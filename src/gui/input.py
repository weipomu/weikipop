# src/gui/input.py
import logging
import sys
import threading
import time

from pynput import mouse

from src.config.config import config, IS_LINUX, IS_MACOS

if IS_LINUX:
    from Xlib import display as xlib_display
    from Xlib.error import XError
    from Xlib import XK
elif IS_MACOS:
    import Quartz
    from AppKit import NSEvent
else:
    import ctypes
    import ctypes.wintypes as _wintypes
    import keyboard

    # Constants for WH_MOUSE_LL scroll suppression
    _WH_MOUSE_LL    = 14
    _WM_MOUSEWHEEL  = 0x020A
    _WM_MOUSEHWHEEL = 0x020E
    _LLMHF_INJECTED = 0x00000001

    class _MSLLHOOKSTRUCT(ctypes.Structure):
        _fields_ = [
            ('pt',          _wintypes.POINT),
            ('mouseData',   _wintypes.DWORD),
            ('flags',       _wintypes.DWORD),
            ('time',        _wintypes.DWORD),
            ('dwExtraInfo', ctypes.POINTER(ctypes.c_ulong)),
        ]

    _HOOKPROC = ctypes.WINFUNCTYPE(
        ctypes.c_ssize_t,   # LRESULT (pointer-sized on 64-bit)
        ctypes.c_int,       # nCode
        _wintypes.WPARAM,
        _wintypes.LPARAM,
    )

    class _WindowsScrollHook:
        """Suppression-only WH_MOUSE_LL hook.
        Installed *before* pynput so the chain order is:
          pynput (records delta) → our hook (suppresses app) → [app blocked]
        Windows hooks are LIFO so last-installed = first-called; by blocking in
        start() until our hook is live, pynput is guaranteed to be installed after
        us and therefore called first."""

        def __init__(self, input_loop):
            self._input_loop   = input_loop
            self._hook_handle  = None
            self._proc         = None   # keep reference alive
            self._thread       = None
            self._ready        = threading.Event()

        def start(self):
            self._thread = threading.Thread(
                target=self._run, daemon=True, name="ScrollSuppressHook"
            )
            self._thread.start()
            # Block until the hook is actually installed so pynput (started right
            # after) ends up last-installed = first-called in the hook chain.
            self._ready.wait(timeout=1.0)

        def _run(self):
            u32 = ctypes.windll.user32
            # Declare proper 64-bit-safe types before any call.
            u32.SetWindowsHookExW.restype  = ctypes.c_void_p
            u32.SetWindowsHookExW.argtypes = [
                ctypes.c_int, _HOOKPROC, ctypes.c_void_p, ctypes.c_ulong
            ]
            u32.CallNextHookEx.restype  = ctypes.c_ssize_t
            u32.CallNextHookEx.argtypes = [
                ctypes.c_void_p, ctypes.c_int,
                _wintypes.WPARAM, _wintypes.LPARAM,
            ]
            u32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]

            self._proc        = _HOOKPROC(self._callback)
            self._hook_handle = u32.SetWindowsHookExW(
                _WH_MOUSE_LL, self._proc, None, 0
            )
            self._ready.set()   # unblock start() regardless of success/failure

            if not self._hook_handle:
                logger.warning(
                    "Could not install scroll suppression hook — "
                    "popup scrolling works normally but other apps may also scroll."
                )
                return

            msg = _wintypes.MSG()
            while u32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
                u32.TranslateMessage(ctypes.byref(msg))
                u32.DispatchMessageW(ctypes.byref(msg))
            u32.UnhookWindowsHookEx(self._hook_handle)

        def _callback(self, nCode, wParam, lParam):
            try:
                if nCode >= 0 and wParam in (_WM_MOUSEWHEEL, _WM_MOUSEHWHEEL):
                    data = ctypes.cast(
                        lParam, ctypes.POINTER(_MSLLHOOKSTRUCT)
                    ).contents
                    if (not (data.flags & _LLMHF_INJECTED)
                            and self._input_loop.suppress_scroll):
                        return 1    # suppress: skip CallNextHookEx → app never sees it
            except Exception:
                pass    # never raise inside a low-level hook callback
            return ctypes.windll.user32.CallNextHookEx(
                self._hook_handle, nCode, wParam, lParam
            )


logger = logging.getLogger(__name__)

class LinuxX11KeyboardController:
    def __init__(self, hotkey_str):
        self.hotkey_str = hotkey_str.lower()
        try:
            self.display = xlib_display.Display()
            self._setup_keycodes()
        except (XError, Exception) as e:
            logger.critical("Could not connect to X server. Is DISPLAY environment variable set? Error: %s", e)
            logger.critical("Weikipop cannot run without a graphical session.")
            sys.exit(1)

    def _setup_keycodes(self):
        self.modifier_groups = []
        modifier_map = {
            'shift': ['Shift_L', 'Shift_R'],
            'ctrl': ['Control_L', 'Control_R'],
            'alt': ['Alt_L', 'Alt_R']
        }
        hotkeys = self.hotkey_str.split('+')

        for key in hotkeys:
            target_keysyms = modifier_map.get(key)
            if not target_keysyms:
                logger.critical(f"Unsupported hotkey '{key}' for Linux/X11. Use 'shift', 'ctrl', or 'alt'.")
                sys.exit(1)
            group_keycodes = set()
            for keysym_str in target_keysyms:
                keysym = XK.string_to_keysym(keysym_str)
                if keysym:
                    keycode = self.display.keysym_to_keycode(keysym)
                    if keycode:
                        group_keycodes.add(keycode)

            if not group_keycodes:
                logger.critical(f"Could not find keycodes for hotkey '{key}'.")
                sys.exit(1)

            self.modifier_groups.append(group_keycodes)

    def is_hotkey_pressed(self) -> bool:
        try:
            key_map = self.display.query_keymap()
            for group in self.modifier_groups:
                group_is_pressed = False
                for keycode in group:
                    if (key_map[keycode // 8] >> (keycode % 8)) & 1:
                        group_is_pressed = True
                        break
                if not group_is_pressed:
                    return False
            return True
        except XError:
            return False

    def is_key_pressed(self, key_str: str) -> bool:
        # Not implemented for X11 yet.
        return False


class WindowsKeyboardController:
    def __init__(self, hotkey_str):
        self.hotkey_str = hotkey_str.lower()

    def is_hotkey_pressed(self) -> bool:
        try:
            return keyboard.is_pressed(self.hotkey_str)
        except ImportError:
            logger.critical("FATAL: The 'keyboard' library failed to import a backend. This often means it needs to be run with administrator/sudo privileges.")
            sys.exit(1)
        except Exception:
            return False

    def is_key_pressed(self, key_str: str) -> bool:
        try:
            return keyboard.is_pressed(key_str)
        except Exception:
            return False


class MacOSKeyboardController:
    def __init__(self, hotkey_str):
        self.hotkey_str = hotkey_str.lower()
        self.modifiers = self.hotkey_str.split('+')

        # Map common hotkey strings to macOS key codes
        key_mapping = {
            'shift': [56, 60],  # Left and Right Shift
            'ctrl': [59, 62],   # Left and Right Control
            'alt': [58, 61],    # Left and Right Option/Alt
            'cmd': [55, 54],    # Left and Right Command
        }

        for mod in self.modifiers:
            self.keycodes_to_check = key_mapping.get(mod, [])
            if not self.keycodes_to_check:
                logger.critical(
                    f"Unsupported hotkey '{self.hotkey_str}' for macOS. Use 'shift', 'ctrl', 'alt', or 'cmd'.")
                sys.exit(1)

    def is_hotkey_pressed(self) -> bool:
        try:
            # Get current modifier flags
            flags = NSEvent.modifierFlags()

            # Iterate through all required modifiers in the combo
            for mod in self.modifiers:
                if mod == 'shift':
                    if not (flags & (1 << 17) or flags & (1 << 18)):
                        return False
                elif mod == 'ctrl':
                    if not (flags & (1 << 12)):
                        return False
                elif mod == 'alt':
                    if not (flags & (1 << 19)):
                        return False
                elif mod == 'cmd':
                    if not (flags & (1 << 20)):
                        return False
            return True
        except Exception as e:
            logger.warning(f"Error checking hotkey state: {e}")
            return False

    def is_key_pressed(self, key_str: str) -> bool:
        # Not implemented yet.
        return False

class InputLoop(threading.Thread):
    def __init__(self, shared_state):
        super().__init__(daemon=True, name="InputLoop")
        self.shared_state = shared_state
        self.mouse_controller = mouse.Controller()

        self.hotkey_str = config.hotkey.lower()
        if IS_LINUX:
            self.keyboard_controller = LinuxX11KeyboardController(self.hotkey_str)
        elif IS_MACOS:
            self.keyboard_controller = MacOSKeyboardController(self.hotkey_str)
        else: # IS_WINDOWS
            self.keyboard_controller = WindowsKeyboardController(self.hotkey_str)

        self.hotkey_is_pressed = False  # cached for main thread — avoids extra syscall
        self.started_auto_mode = False

        self.scroll_dy = 0
        self.scroll_lock = threading.Lock()

        # When True, scroll events are suppressed at the OS level (Windows only) so
        # other apps don't receive them while the popup is visible. Set by the popup.
        self.suppress_scroll = False

        # Track mouse button states for mouse shortcuts
        self.mouse_buttons_pressed = set()
        self.mouse_button_lock = threading.Lock()

        # On Windows: install the suppression hook *before* starting pynput's listener.
        # This guarantees pynput ends up last-installed = first-called in the hook chain,
        # so pynput records the scroll delta as normal and our hook suppresses downstream.
        if not IS_LINUX and not IS_MACOS:
            self._scroll_hook = _WindowsScrollHook(self)
            self._scroll_hook.start()   # blocks until hook is live
        else:
            self._scroll_hook = None

        # Start mouse listener for scroll and click events — unchanged from original.
        self.mouse_listener = mouse.Listener(
            on_scroll=self.on_scroll,
            on_click=self.on_click,
        )
        self.mouse_listener.start()

    def on_click(self, x, y, button, pressed):
        with self.mouse_button_lock:
            if pressed:
                self.mouse_buttons_pressed.add(button)
            else:
                self.mouse_buttons_pressed.discard(button)

    def on_scroll(self, x, y, dx, dy):
        with self.scroll_lock:
            self.scroll_dy += dy

    def get_and_reset_scroll_delta(self):
        with self.scroll_lock:
            delta = self.scroll_dy
            self.scroll_dy = 0
        return delta

    def run(self):
        logger.debug("Input thread started.")
        last_mouse_pos = (0, 0)
        hotkey_was_pressed = False

        while self.shared_state.running:
            if not config.is_enabled:
                time.sleep(0.1)
                continue
            try:
                current_mouse_pos = self.mouse_controller.position
                try:
                    hotkey_is_pressed = self.keyboard_controller.is_hotkey_pressed()
                except Exception:
                    hotkey_is_pressed = False

                # trigger screenshots + ocr in manual mode
                if hotkey_is_pressed and not hotkey_was_pressed and not config.auto_scan_mode:
                    logger.info(f"Input: Hotkey '{config.hotkey}' pressed. Triggering screenshot.")
                    self.shared_state.screenshot_trigger_event.set()

                # trigger initial screenshots + ocr in auto mode
                if not self.started_auto_mode and config.auto_scan_mode:
                    self.shared_state.screenshot_trigger_event.set()
                self.started_auto_mode = config.auto_scan_mode

                # trigger screenshots + ocr in auto-on-mouse-move mode
                if config.auto_scan_mode and config.auto_scan_on_mouse_move and current_mouse_pos != last_mouse_pos:
                    self.shared_state.screenshot_trigger_event.set()

                # trigger hit_scans + lookups
                if current_mouse_pos != last_mouse_pos:
                    self.shared_state.hit_scan_queue.trigger()

                if hotkey_was_pressed and not hotkey_is_pressed:
                    logger.info(f"Input: Hotkey '{config.hotkey}' released.")

                last_mouse_pos = current_mouse_pos
                hotkey_was_pressed = hotkey_is_pressed
                self.hotkey_is_pressed = hotkey_is_pressed
            except:
                logger.exception("An unexpected error occurred in the input loop. Continuing...")
            finally:
                time.sleep(0.01)
        logger.debug("Input thread stopped.")

    def is_virtual_hotkey_down(self):
        return self.keyboard_controller.is_hotkey_pressed() or (
                config.auto_scan_mode and config.auto_scan_mode_lookups_without_hotkey)

    def is_key_pressed(self, key_str: str) -> bool:
        key_lower = (key_str or "").lower()

        mouse_button_map = {
            'mouse4': mouse.Button.x1,
            'mouse5': mouse.Button.x2,
            'xbutton1': mouse.Button.x1,
            'xbutton2': mouse.Button.x2,
            'middlemouse': mouse.Button.middle,
            'mouse3': mouse.Button.middle,
        }

        if key_lower in mouse_button_map:
            with self.mouse_button_lock:
                return mouse_button_map[key_lower] in self.mouse_buttons_pressed

        return self.keyboard_controller.is_key_pressed(key_str)

    def reapply_settings(self):
        logger.debug(f"InputLoop: Re-applying settings. New hotkey: '{config.hotkey}'.")
        self.hotkey_str = config.hotkey.lower()
        if IS_LINUX:
            self.keyboard_controller = LinuxX11KeyboardController(self.hotkey_str)
        elif IS_MACOS:
            self.keyboard_controller = MacOSKeyboardController(self.hotkey_str)
        else: # IS_WINDOWS
            self.keyboard_controller = WindowsKeyboardController(self.hotkey_str)

    def get_mouse_pos(self):
        """Reuse the existing mouse controller instead of creating a new one."""
        pos = self.mouse_controller.position
        return (int(pos[0]), int(pos[1]))
