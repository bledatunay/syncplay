import os
import subprocess
import time

import re
import glob

from syncplay import constants
from syncplay.players.basePlayer import BasePlayer


try:
    import win32con
    import win32gui
    import win32api
except Exception:  # pragma: no cover
    win32con = None
    win32gui = None
    win32api = None


class PotPlayerPlayer(BasePlayer):
    """PotPlayer integration via Windows messages.

    PotPlayer exposes a small control surface via SendMessage:
      - WM_COMMAND (0x0111) for basic playback actions
      - WM_USER    (0x0400) for state/time queries and setters

    Message IDs follow the widely used community API (e.g. Specter333 AHK lib).
    """

    # Feature flags expected by Syncplay's client UI
    # (Other player adapters expose these as class attributes.)
    alertOSDSupported = False
    chatOSDSupported = False
    speedSupported = False
    customOpenDialog = False
    osdMessageSeparator = "; "

    # Window class names
    _WINDOW_CLASSES = ["PotPlayer64", "PotPlayer"]

    # Messages
    _WM_COMMAND = 0x0111
    _WM_USER = 0x0400

    # WM_COMMAND commands
    _CMD_PLAY = 20001
    _CMD_PAUSE = 20000
    _CMD_STOP = 20002
    _CMD_PLAY_PAUSE = 10014

    # WM_USER commands
    _POT_GET_TOTAL_TIME = 0x5002  # ms
    _POT_GET_PROGRESS_TIME = 0x5003  # ms
    _POT_GET_CURRENT_TIME = 0x5004  # ms
    _POT_SET_CURRENT_TIME = 0x5005  # ms
    _POT_GET_PLAY_STATUS = 0x5006  # -1 stopped, 1 paused, 2 running
    _POT_SET_PLAY_STATUS = 0x5007  # 0 toggle, 1 paused, 2 running

    def __init__(self, client):
        # Twisted reactor is used throughout Syncplay's player adapters to
        # schedule shutdown safely from polling loops.
        from twisted.internet import reactor
        self.reactor = reactor
        self._client = client
        self._hwnd = None
        self._stopping = False
        self._filename = None
        self._filepath = None
        self._duration = None

        # Cache for resolving a window-title filename to an actual filesystem path
        # (used for correct filesize reporting in Syncplay's user list).
        self._resolvedPathCache = {}
        self._lastPathScan = 0

        self._lastFilePoll = 0

        # For play/pause inference in case PotPlayer returns ambiguous play state
        # codes on some builds/skins.
        self._lastPosMs = None
        self._lastPosCheck = None

        # Smoothed position we report to Syncplay.
        # PotPlayer can momentarily snap backwards or return stale values around
        # rapid pause/unpause toggles and shortly after seeks. If we forward those
        # transient values, Syncplay interprets them as user seeks and will
        # propagate "corrections" which can lead to snap-back loops and
        # long-lasting desync.
        self._smoothPosMs = None
        self._smoothPosAt = None

        # PotPlayer time queries can exhibit small backwards "jitter" (e.g. -200ms
        # to -1500ms) during playback. If we forward those small negative deltas to
        # Syncplay, the client may interpret them as a user seek and propagate
        # backwards jumps to the room, causing cascading rewinds.
        # We therefore clamp small backwards movements while *playing*.
        self._maxBackwardJitterMs = 300

        # When *paused*, PotPlayer may still report small backwards adjustments
        # (often due to keyframe alignment or internal refresh). Syncplay treats
        # any backwards jump as a potential seek and will propagate it, which can
        # cause "stuck" or "rewind" loops in mixed-player rooms.
        # We clamp a slightly larger window while paused because keyframe snaps
        # can be a couple of seconds.
        self._maxBackwardJitterPausedMs = 300

        # When Syncplay requests a seek via setPosition(), we expect a legitimate
        # backwards/forwards jump to occur. We mark a short window where we do not
        # clamp backwards movement so real seeks propagate correctly.
        self._allowJumpUntil = 0

        # Some PotPlayer builds/skins occasionally return 0 for the current time
        # while playback is active (likely during internal refresh or when the UI
        # thread is busy). If we forward that 0 to Syncplay, it can be interpreted
        # as a user seek-to-start and cause repeated "rewind to 0" corrections.
        # We therefore require consecutive 0 reads before accepting it as real.
        self._zeroPosReadStreak = 0

        # Pause/play state stabilisation
        self._lastPaused = None
        self._pauseCmdUntil = 0  # during this window, trust status codes over time drift
        self._playingEvidence = 0

        # When we explicitly request pause/unpause, PotPlayer may briefly report
        # inconsistent play state and/or a slightly older position. Syncplay may
        # interpret that as an unintended seek or as an extra unpause, which can
        # trigger "ready" toggles and desync loops. To mimic the stability of
        # other player adapters, we temporarily force the requested paused state.
        self._forcedPausedValue = None
        self._forcedPausedUntil = 0

        # NOTE: We intentionally do NOT attempt to auto-unpause when Syncplay
        # sets readiness. That behaviour is core Syncplay logic and must remain
        # consistent across players. We only stabilise PotPlayer's transient
        # playstate/time bounces so they don't *accidentally* trigger ready logic.

        # When transitioning to paused, PotPlayer can snap the reported time
        # backwards a few seconds (keyframe alignment). Freeze the position
        # briefly on the play->pause edge so the pause point is exact.
        self._freezePosMs = None
        self._freezePosUntil = 0

        # When pausing, PotPlayer may snap back to a nearby keyframe. Capture a
        # pause anchor timestamp and keep reporting it briefly so the paused
        # position matches other players (e.g. MPC).
        self._pauseAnchorMs = None
        self._pauseAnchorUntil = 0

        # Any time we issue a seek (setPosition), allow a short window where raw
        # PotPlayer time can be temporarily inconsistent without being treated as
        # a user-driven seek by Syncplay.
        self._seekCmdUntil = 0

    # ------------------------
    # Window helpers
    # ------------------------
    def _find_hwnd(self):
        if win32gui is None:
            return None
        for cls in self._WINDOW_CLASSES:
            hwnd = win32gui.FindWindow(cls, None)
            if hwnd:
                return hwnd
        return None

    def _ensure_hwnd(self):
        if self._hwnd and win32gui and win32gui.IsWindow(self._hwnd):
            return True
        self._hwnd = self._find_hwnd()
        return bool(self._hwnd)

    def _send(self, msg, wparam=0, lparam=0):
        if not self._ensure_hwnd():
            raise RuntimeError("PotPlayer window not found")
        # win32gui.SendMessage returns LRESULT (int)
        return win32gui.SendMessage(self._hwnd, msg, wparam, lparam)

    def _post(self, msg, wparam=0, lparam=0):
        if not self._ensure_hwnd():
            raise RuntimeError("PotPlayer window not found")
        return win32gui.PostMessage(self._hwnd, msg, wparam, lparam)

    # ------------------------
    # BasePlayer API
    # ------------------------
    def askForStatus(self):
        """Poll PotPlayer for playback status and current position."""
        # If PotPlayer was closed, Syncplay should close too (mirrors VLC behaviour).
        if not self._ensure_hwnd():
            self._requestStop()
            return

        # Update file metadata periodically (required for the Syncplay playlist).
        now = time.time()
        if now - self._lastFilePoll > 1.0:
            self._lastFilePoll = now
            try:
                self._pollFileInfo()
            except Exception:
                # Non-fatal; continue updating play state.
                pass

        try:
            status = int(self._send(self._WM_USER, self._POT_GET_PLAY_STATUS, 0))
            raw_pos_ms = int(self._send(self._WM_USER, self._POT_GET_CURRENT_TIME, 0))
        except Exception:
            self._requestStop()
            return

        now = time.time()

        pos_ms = raw_pos_ms

        # Guard against transient 0 reads which can trigger pathological
        # rewinds via Syncplay's resync logic.
        if pos_ms == 0 and self._lastPosMs is not None and self._lastPosMs > 5000 and status != -1:
            self._zeroPosReadStreak += 1
            if self._zeroPosReadStreak < 2:
                pos_ms = self._lastPosMs
        else:
            self._zeroPosReadStreak = 0

        # Determine paused state.


        # Determine paused state.
        # Prefer PotPlayer's play status codes:
        #   -1 stopped, 1 paused, 2 playing
        # Some PotPlayer builds may briefly report inconsistent values around rapid toggles;
        # we therefore add a small stabilisation window around explicit pause/unpause commands.
        paused = True
        if status == 2:
            paused = False
        elif status == 1:
            paused = True
        else:
            # Unknown: fall back to last known state
            paused = True if self._lastPaused is None else self._lastPaused

        # Freeze position briefly on play->pause transition to avoid keyframe
        # snap-back being interpreted as a seek.
        if self._lastPaused is False and paused is True:
            if self._smoothPosMs is not None:
                self._freezePosMs = int(self._smoothPosMs)
                self._freezePosUntil = now + 1.25
                self._pauseAnchorMs = int(self._smoothPosMs)
                self._pauseAnchorUntil = now + 1.50

        # During a short window after we explicitly requested pause/unpause,
        # force the requested state to prevent transient oscillation.
        if now <= self._forcedPausedUntil and self._forcedPausedValue is not None:
            paused = bool(self._forcedPausedValue)
            # Also clamp a transient backwards position snap during this window
            # (unless a deliberate seek is in progress).
            if (now > self._allowJumpUntil and self._smoothPosMs is not None and pos_ms < self._smoothPosMs):
                pos_ms = self._smoothPosMs

        # If we just paused, ignore a backwards snap for a short time.
        if now <= self._freezePosUntil and self._freezePosMs is not None:
            if pos_ms < self._freezePosMs:
                pos_ms = self._freezePosMs

        # When paused and within the pause-anchor window, keep reporting the
        # anchor timestamp (prevents keyframe snap-back from becoming a room-wide
        # correction seek).
        if paused and now <= self._pauseAnchorUntil and self._pauseAnchorMs is not None:
            if pos_ms < self._pauseAnchorMs:
                pos_ms = self._pauseAnchorMs

        # If we're outside the command stabilisation window, allow a conservative inference:
        # Only treat "paused" as playing if we have sustained evidence over multiple polls.
        if now > self._pauseCmdUntil:
            if paused and self._lastPosMs is not None and self._lastPosCheck is not None:
                dt = max(0.0, now - self._lastPosCheck)
                if dt > 0.25 and (pos_ms - self._lastPosMs) > 350:
                    self._playingEvidence += 1
                else:
                    self._playingEvidence = 0

                if self._playingEvidence >= 3:
                    paused = False
            else:
                self._playingEvidence = 0

        # --- Position smoothing ---
        # Build a monotonic (when playing) position to avoid transient backward snaps.
        # Accept large jumps as real seeks.
        if self._smoothPosMs is None:
            self._smoothPosMs = pos_ms
            self._smoothPosAt = now
        else:
            # Keep reported position stable. PotPlayer can sometimes return brief
            # backwards snaps (keyframe alignment / internal refresh), especially
            # around rapid pause/unpause. Forwarding those raw values causes Syncplay
            # to interpret them as user seeks and can create oscillations.
            jump_ms = abs(pos_ms - self._smoothPosMs)

            # Treat large jumps as real seeks (or when we recently requested a seek).
            if jump_ms > 1500 or now <= self._allowJumpUntil:
                self._smoothPosMs = pos_ms
                self._smoothPosAt = now
            else:
                # Small backwards movement is jitter -> clamp to last known.
                if pos_ms < self._smoothPosMs:
                    pos_ms = self._smoothPosMs
                else:
                    # Small forward drift is fine.
                    self._smoothPosMs = pos_ms
                self._smoothPosAt = now

        pos_ms = self._smoothPosMs

        pos_ms = self._smoothPosMs

        self._lastPaused = paused
        self._lastPosMs = pos_ms
        self._lastPosCheck = now
        position = max(0.0, pos_ms / 1000.0)
        self._client.updatePlayerStatus(paused, position)

    def _requestStop(self):
        if self._stopping:
            return
        self._stopping = True
        try:
            self.reactor.callFromThread(self._client.stop, False,)
        except Exception:
            # Avoid raising from polling thread
            pass

    _TITLE_SUFFIX_RE = re.compile(r"\s*[-–]\s*PotPlayer\s*(?:64)?\s*$", re.IGNORECASE)

    def _getWindowTitleFilename(self):
        if win32gui is None or not self._ensure_hwnd():
            return None
        title = win32gui.GetWindowText(self._hwnd) or ""
        title = title.strip()
        if not title:
            return None
        # Typical title format: "<name> - PotPlayer" (sometimes with different dash types).
        title = re.sub(self._TITLE_SUFFIX_RE, "", title).strip()
        # If user customized the title format, we may still end up with generic titles.
        if not title or "potplayer" in title.lower():
            return None
        return title

    def _pollFileInfo(self):
        # PotPlayer does not expose a stable public API for current file path.
        # For Syncplay purposes we at minimum need a non-empty "path" to avoid
        # the UI showing "(No file being played)".
        filename = self._getWindowTitleFilename()
        if not filename:
            return

        try:
            duration_ms = int(self._send(self._WM_USER, self._POT_GET_TOTAL_TIME, 0))
        except Exception:
            duration_ms = 0
        duration = max(0.0, duration_ms / 1000.0) if duration_ms else 0

        # Try to resolve a real filesystem path so Syncplay can display filesize.
        # PotPlayer's public message API doesn't reliably expose the full path,
        # but the active playlist (.dpl) typically stores absolute paths.
        path = self._resolveFilePathFromTitle(filename) or filename

        if (filename != self._filename) or (path != self._filepath) or (duration != self._duration):
            self._filename = filename
            self._filepath = path
            self._duration = duration
            self._client.updateFile(self._filename, self._duration, self._filepath)

    _DRIVE_PATH_RE = re.compile(r"[A-Za-z]:\\")

    def _resolveFilePathFromTitle(self, titleText):
        """Best-effort mapping from PotPlayer's window title to an absolute path.

        The window title usually contains only the basename, which is insufficient
        to compute filesize. We therefore scan PotPlayer's recent playlist files
        for a matching entry.
        """
        if not titleText:
            return None

        # If title already looks like an absolute path and exists, use it.
        if (self._DRIVE_PATH_RE.search(titleText) or titleText.startswith("\\\\")) and os.path.exists(titleText):
            return titleText

        cached = self._resolvedPathCache.get(titleText)
        if cached and os.path.exists(cached):
            return cached

        now = time.time()
        # Avoid hammering the disk during polling.
        if now - self._lastPathScan < 1.0:
            return None
        self._lastPathScan = now

        resolved = self._scanPotPlayerPlaylistsForFilename(titleText)
        if resolved:
            self._resolvedPathCache[titleText] = resolved
        return resolved

    def _scanPotPlayerPlaylistsForFilename(self, filename):
        """Scan PotPlayer playlist files (.dpl) to find an absolute path."""
        roots = []
        appdata = os.environ.get("APPDATA")
        localappdata = os.environ.get("LOCALAPPDATA")
        if appdata:
            roots.extend([
                os.path.join(appdata, "PotPlayerMini64"),
                os.path.join(appdata, "PotPlayerMini"),
                os.path.join(appdata, "Daum", "PotPlayer"),
                os.path.join(appdata, "PotPlayer"),
            ])
        if localappdata:
            roots.extend([
                os.path.join(localappdata, "PotPlayerMini64"),
                os.path.join(localappdata, "PotPlayerMini"),
                os.path.join(localappdata, "Daum", "PotPlayer"),
                os.path.join(localappdata, "PotPlayer"),
            ])

        dpl_files = []
        for root in roots:
            if root and os.path.isdir(root):
                dpl_files.extend(glob.glob(os.path.join(root, "**", "*.dpl"), recursive=True))
        if not dpl_files:
            return None

        dpl_files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        dpl_files = dpl_files[:10]

        fn = os.path.basename(filename).strip().lower()

        candidate_re = re.compile(
            r"(?i)([A-Za-z]:\\[^\r\n\"]*{}|[A-Za-z]:/[^\r\n\"]*{}|\\\\[^\r\n\"]*{}|\\\\[^\r\n\"]*{})".format(
                re.escape(fn), re.escape(fn), re.escape(fn), re.escape(fn)
            )
        )

        for dpl in dpl_files:
            try:
                with open(dpl, "rb") as f:
                    raw = f.read()

                text = None
                for enc in ("utf-16", "utf-16le", "utf-8"):
                    try:
                        text = raw.decode(enc)
                        break
                    except Exception:
                        continue
                if not text:
                    continue

                if fn not in text.lower():
                    continue

                m = candidate_re.search(text)
                if not m:
                    continue
                path = m.group(1).strip().strip('"')
                if os.path.exists(path):
                    return path
            except Exception:
                continue
        return None

    def displayMessage(self, message, duration=(constants.OSD_DURATION * 1000), OSDType=constants.OSD_NOTIFICATION, mood=constants.MESSAGE_NEUTRAL):
        """PotPlayer does not expose a stable cross-version API for custom OSD text.

        We intentionally no-op rather than simulating keystrokes.
        """
        return

    def drop(self):
        # When Syncplay closes, close PotPlayer too.
        try:
            if self._ensure_hwnd() and win32gui is not None:
                win32gui.PostMessage(self._hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            pass
        self._hwnd = None

    @staticmethod
    def run(client, playerPath, filePath, args):
        # Launch PotPlayer. It accepts the media path as an argument.
        expanded = PotPlayerPlayer.getExpandedPath(playerPath)
        popen_args = [expanded]
        if filePath:
            popen_args.append(filePath)
        if args:
            popen_args.extend(args)
        try:
            subprocess.Popen(popen_args)
        except Exception:
            # Let Syncplay error handling surface this elsewhere.
            raise

        player = PotPlayerPlayer(client)

        # Give the window a moment to appear.
        end = time.time() + 10
        while time.time() < end:
            if player._ensure_hwnd():
                break
            time.sleep(0.1)

        # Populate initial file info if possible (ensures Syncplay shows it in playlist).
        try:
            player._pollFileInfo()
        except Exception:
            pass

        client.initPlayer(player)
        return player

    def setPaused(self, value):
        try:
            now = time.time()
            # Force the requested paused state briefly to avoid transient "bounce"
            # right after sending pause/unpause.
            self._forcedPausedValue = bool(value)
            self._forcedPausedUntil = now + 0.35
            self._pauseCmdUntil = now + 0.35
            self._playingEvidence = 0

            self._send(self._WM_USER, self._POT_SET_PLAY_STATUS, 1 if value else 2)
        except Exception:
            return

    def setFeatures(self, featureList):
        # No-op: PotPlayer feature negotiation is not supported.
        return

    def setPosition(self, value):
        try:
            # Allow an actual jump (including backwards) to be observed for a short
            # period so Syncplay can propagate a real user seek.
            now = time.time()
            self._allowJumpUntil = now + 1.0
            self._seekCmdUntil = now + 1.25
            target_ms = int(float(value) * 1000)
            target_ms = max(0, target_ms)
            self._smoothPosMs = target_ms
            self._smoothPosAt = now
            self._send(self._WM_USER, self._POT_SET_CURRENT_TIME, target_ms)
        except Exception:
            return

    def setSpeed(self, value):
        # Not implemented (PotPlayer has commands, but no stable getter/setter is relied upon by Syncplay).
        return

    def openFile(self, filePath, resetPosition=False):
        # PotPlayer does not expose a simple, reliable "open exact path" window-message API.
        # Syncplay already passes the file at startup via run(); keep this method as a no-op.
        return

    @staticmethod
    def getDefaultPlayerPathsList():
        return constants.POTPLAYER_PATHS

    @staticmethod
    def isValidPlayerPath(path):
        return bool(PotPlayerPlayer.getExpandedPath(path))

    @staticmethod
    def getIconPath(path):
        return constants.POTPLAYER_ICONPATH

    @staticmethod
    def getExpandedPath(path):
        if not path:
            return None
        p = os.path.expandvars(path)
        if os.path.isfile(p):
            if p.lower().endswith("potplayermini.exe") or p.lower().endswith("potplayermini64.exe"):
                return p
        # If path is a folder, try common exe names.
        for exe in ("PotPlayerMini64.exe", "PotPlayerMini.exe", "potplayermini64.exe", "potplayermini.exe"):
            for sep in ("", "\\"):
                cand = p + sep + exe
                if os.path.isfile(cand):
                    return cand
        return None

    @staticmethod
    def openCustomOpenDialog(self):
        # Not implemented.
        return

    @staticmethod
    def getPlayerPathErrors(playerPath, filePath):
        return None
