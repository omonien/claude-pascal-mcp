"""Windows application interaction via Win32 API.

Provides click, type, and key input for desktop applications using
PostMessage (for clicks) and SendInput (for keyboard). Coordinates
use screenshot pixels — the same coordinate space as screenshot_app.

Ported from the Tina4Delphi preview_bridge.py battle-tested implementation.
"""

from __future__ import annotations

import ctypes
import ctypes.wintypes
import sys
import time

from pascal_mcp.screenshot import (
    _find_window_by_title,
    _get_window_title,
    _bring_window_to_front,
)


# ---------------------------------------------------------------------------
# Win32 constants
# ---------------------------------------------------------------------------

INPUT_MOUSE = 0
INPUT_KEYBOARD = 1

MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_ABSOLUTE = 0x8000
MOUSEEVENTF_VIRTUALDESK = 0x4000

_ABS_FLAGS = MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK

KEYEVENTF_UNICODE = 0x0004
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_EXTENDEDKEY = 0x0001

WM_LBUTTONDOWN = 0x0201
WM_LBUTTONUP = 0x0202
WM_RBUTTONDOWN = 0x0204
WM_RBUTTONUP = 0x0205
WM_LBUTTONDBLCLK = 0x0203
MK_LBUTTON = 0x0001
MK_RBUTTON = 0x0002

# Virtual key codes for special keys
VK_MAP = {
    "enter": 0x0D, "return": 0x0D, "tab": 0x09, "escape": 0x1B, "esc": 0x1B,
    "backspace": 0x08, "delete": 0x2E, "space": 0x20,
    "up": 0x26, "down": 0x28, "left": 0x25, "right": 0x27,
    "home": 0x24, "end": 0x23, "pageup": 0x21, "pagedown": 0x22,
    "f1": 0x70, "f2": 0x71, "f3": 0x72, "f4": 0x73, "f5": 0x74,
    "f6": 0x75, "f7": 0x76, "f8": 0x77, "f9": 0x78, "f10": 0x79,
    "f11": 0x7A, "f12": 0x7B,
    "ctrl": 0x11, "control": 0x11, "alt": 0x12, "shift": 0x10,
}


# ---------------------------------------------------------------------------
# ctypes structures for SendInput
# ---------------------------------------------------------------------------

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_long),
        ("dy", ctypes.c_long),
        ("mouseData", ctypes.c_ulong),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class _INPUT_UNION(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong), ("ii", _INPUT_UNION)]


def _send_input(*inputs: INPUT) -> int:
    """Send one or more INPUT structs via SendInput."""
    arr = (INPUT * len(inputs))(*inputs)
    return ctypes.windll.user32.SendInput(len(inputs), arr, ctypes.sizeof(INPUT))


# ---------------------------------------------------------------------------
# Window geometry helpers
# ---------------------------------------------------------------------------

def _window_to_screen(hwnd: int, x: int, y: int) -> tuple[int, int]:
    """Convert window-relative (x, y) to screen coordinates."""
    rect = ctypes.wintypes.RECT()
    ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return rect.left + x, rect.top + y


def _find_deepest_child(hwnd: int, screen_x: int, screen_y: int) -> int:
    """Find the deepest child window at a screen position.

    Recursively walks child windows using ChildWindowFromPointEx
    to find the most specific target for mouse messages.
    """
    CWP_SKIPTRANSPARENT = 0x0004

    current = hwnd
    for _ in range(10):  # Max depth to prevent infinite loops
        pt = ctypes.wintypes.POINT(screen_x, screen_y)
        ctypes.windll.user32.ScreenToClient(current, ctypes.byref(pt))

        child = ctypes.windll.user32.ChildWindowFromPointEx(
            current, pt, CWP_SKIPTRANSPARENT,
        )

        if not child or child == current:
            break
        current = child

    return current


# ---------------------------------------------------------------------------
# Click via PostMessage (uses screenshot pixel coordinates)
# ---------------------------------------------------------------------------

def _click_message(
    hwnd: int, x: int, y: int,
    button: str = "left",
    double: bool = False,
) -> dict:
    """Click via PostMessage using screenshot pixel coordinates.

    The screenshot from PrintWindow is in physical pixels relative to
    the window's top-left corner (matching GetWindowRect).

    Automatically finds the deepest child window at the click position
    and sends the click directly to it with properly translated coordinates.
    """
    if sys.platform != "win32":
        return {"ok": False, "error": "Windows only"}

    # Get window geometry
    rect = ctypes.wintypes.RECT()
    ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect))

    # Screenshot pixel -> screen position
    screen_x = rect.left + x
    screen_y = rect.top + y

    # Find the deepest child at this screen position
    child_hwnd = _find_deepest_child(hwnd, screen_x, screen_y)

    # Get child's class name to decide targeting strategy
    class_buf = ctypes.create_unicode_buffer(256)
    ctypes.windll.user32.GetClassNameW(child_hwnd, class_buf, 256)
    child_class = class_buf.value

    # Chrome/WebView2 controls don't process PostMessage clicks properly
    # for focus management — send to the main window instead.
    chrome_classes = {
        "Chrome_WidgetWin_0", "Chrome_WidgetWin_1",
        "Chrome_RenderWidgetHostHWND",
        "Intermediate D3D Window",
    }
    if child_class in chrome_classes or child_hwnd == hwnd:
        target_hwnd = hwnd
    else:
        target_hwnd = child_hwnd

    # Screen -> target's client coords
    pt = ctypes.wintypes.POINT(screen_x, screen_y)
    ctypes.windll.user32.ScreenToClient(target_hwnd, ctypes.byref(pt))
    cx, cy = pt.x, pt.y

    # Pack coordinates into lParam
    lparam = ((cy & 0xFFFF) << 16) | (cx & 0xFFFF)

    if button == "right":
        down_msg = WM_RBUTTONDOWN
        up_msg = WM_RBUTTONUP
        wparam = MK_RBUTTON
    else:
        down_msg = WM_LBUTTONDOWN
        up_msg = WM_LBUTTONUP
        wparam = MK_LBUTTON

    _bring_window_to_front(hwnd)
    time.sleep(0.1)

    clicks = 2 if double else 1
    for i in range(clicks):
        if i == 1 and button == "left":
            down_msg = WM_LBUTTONDBLCLK
        ctypes.windll.user32.PostMessageW(target_hwnd, down_msg, wparam, lparam)
        time.sleep(0.05)
        ctypes.windll.user32.PostMessageW(target_hwnd, up_msg, 0, lparam)
        time.sleep(0.05)

    time.sleep(0.1)
    return {"ok": True, "target_hwnd": str(target_hwnd), "client_pos": [cx, cy]}


# ---------------------------------------------------------------------------
# Keyboard via SendInput
# ---------------------------------------------------------------------------

def _type_text(hwnd: int, text: str) -> bool:
    """Send unicode text to a window character by character."""
    if sys.platform != "win32":
        return False

    _bring_window_to_front(hwnd)

    for char in text:
        code = ord(char)
        inp_down = INPUT()
        inp_down.type = INPUT_KEYBOARD
        inp_down.ii.ki.wScan = code
        inp_down.ii.ki.dwFlags = KEYEVENTF_UNICODE
        inp_up = INPUT()
        inp_up.type = INPUT_KEYBOARD
        inp_up.ii.ki.wScan = code
        inp_up.ii.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP
        _send_input(inp_down, inp_up)
        time.sleep(0.02)

    return True


def _send_key(hwnd: int, key: str) -> bool:
    """Send a special key or key combination (e.g. 'enter', 'ctrl+a').

    Supports: enter, tab, escape, backspace, delete, arrows, F1-F12,
    and modifier combos like ctrl+a, ctrl+shift+s, alt+f4.
    """
    if sys.platform != "win32":
        return False

    _bring_window_to_front(hwnd)

    parts = [p.strip().lower() for p in key.split("+")]
    modifiers = []
    main_key = None

    for p in parts:
        if p in ("ctrl", "control", "alt", "shift"):
            modifiers.append(VK_MAP[p])
        elif p in VK_MAP:
            main_key = VK_MAP[p]
        elif len(p) == 1:
            # Single character — use virtual key code
            main_key = ord(p.upper())
        else:
            return False

    if main_key is None:
        return False

    # Press modifiers
    for vk in modifiers:
        inp = INPUT()
        inp.type = INPUT_KEYBOARD
        inp.ii.ki.wVk = vk
        _send_input(inp)
        time.sleep(0.02)

    # Press and release main key
    inp_down = INPUT()
    inp_down.type = INPUT_KEYBOARD
    inp_down.ii.ki.wVk = main_key
    inp_up = INPUT()
    inp_up.type = INPUT_KEYBOARD
    inp_up.ii.ki.wVk = main_key
    inp_up.ii.ki.dwFlags = KEYEVENTF_KEYUP
    _send_input(inp_down)
    time.sleep(0.03)
    _send_input(inp_up)
    time.sleep(0.02)

    # Release modifiers in reverse
    for vk in reversed(modifiers):
        inp = INPUT()
        inp.type = INPUT_KEYBOARD
        inp.ii.ki.wVk = vk
        inp.ii.ki.dwFlags = KEYEVENTF_KEYUP
        _send_input(inp)
        time.sleep(0.02)

    return True


# ---------------------------------------------------------------------------
# Public API (called from server.py tools)
# ---------------------------------------------------------------------------

def click_window(
    window_title: str,
    x: int,
    y: int,
    button: str = "left",
    double: bool = False,
) -> str:
    """Click at (x, y) in a window using screenshot pixel coordinates.

    Uses PostMessage with automatic child window targeting.

    Args:
        window_title: Full or partial window title (case-insensitive).
        x: X coordinate in screenshot pixels.
        y: Y coordinate in screenshot pixels.
        button: 'left' or 'right'.
        double: If True, send a double-click.

    Returns:
        Status message.

    Raises:
        RuntimeError: If the window is not found.
    """
    hwnd = _find_window_by_title(window_title)
    if hwnd is None:
        raise RuntimeError(f"Window not found: '{window_title}'")

    actual_title = _get_window_title(hwnd)
    result = _click_message(hwnd, x, y, button=button, double=double)

    if not result.get("ok"):
        raise RuntimeError(result.get("error", "Click failed"))

    click_type = "Double-clicked" if double else "Clicked"
    btn = f" ({button})" if button != "left" else ""
    return f"{click_type}{btn} ({x}, {y}) on '{actual_title}'"


def type_in_window(window_title: str, text: str) -> str:
    """Type text into a window using SendInput Unicode events.

    The window's currently focused control receives the text.

    Args:
        window_title: Full or partial window title (case-insensitive).
        text: The text to type.

    Returns:
        Status message.

    Raises:
        RuntimeError: If the window is not found or typing fails.
    """
    hwnd = _find_window_by_title(window_title)
    if hwnd is None:
        raise RuntimeError(f"Window not found: '{window_title}'")

    actual_title = _get_window_title(hwnd)
    success = _type_text(hwnd, text)

    if not success:
        raise RuntimeError("Type failed")

    return f"Typed {len(text)} character(s) into '{actual_title}'"


def send_key_to_window(window_title: str, key: str) -> str:
    """Send a key or key combination to a window.

    Supports special keys (enter, tab, escape, arrows, F1-F12) and
    modifier combinations (ctrl+a, ctrl+shift+s, alt+f4).

    Args:
        window_title: Full or partial window title (case-insensitive).
        key: Key name or combination (e.g., 'enter', 'ctrl+a', 'f5').

    Returns:
        Status message.

    Raises:
        RuntimeError: If the window is not found or key send fails.
    """
    hwnd = _find_window_by_title(window_title)
    if hwnd is None:
        raise RuntimeError(f"Window not found: '{window_title}'")

    actual_title = _get_window_title(hwnd)
    success = _send_key(hwnd, key)

    if not success:
        raise RuntimeError(
            f"Key '{key}' not recognized. Use: enter, tab, escape, backspace, "
            "delete, space, up/down/left/right, f1-f12, or combos like ctrl+a."
        )

    return f"Sent key '{key}' to '{actual_title}'"
