"""ADB (Android Debug Bridge) interaction tools.

Provides device management, screenshots, UI automation, app management,
and file transfer via ADB commands.
"""

from __future__ import annotations

import io
import os
import shutil
import subprocess
from dataclasses import dataclass

from PIL import Image


@dataclass
class AdbDevice:
    """Information about a connected Android device."""
    serial: str
    state: str  # "device", "offline", "unauthorized"
    model: str
    android_version: str
    screen_size: str


# Key alias map for adb_key
KEY_ALIASES = {
    "home": "KEYCODE_HOME",
    "back": "KEYCODE_BACK",
    "enter": "KEYCODE_ENTER",
    "menu": "KEYCODE_MENU",
    "power": "KEYCODE_POWER",
    "volume_up": "KEYCODE_VOLUME_UP",
    "volume_down": "KEYCODE_VOLUME_DOWN",
    "tab": "KEYCODE_TAB",
    "delete": "KEYCODE_DEL",
    "space": "KEYCODE_SPACE",
    "escape": "KEYCODE_ESCAPE",
    "dpad_up": "KEYCODE_DPAD_UP",
    "dpad_down": "KEYCODE_DPAD_DOWN",
    "dpad_left": "KEYCODE_DPAD_LEFT",
    "dpad_right": "KEYCODE_DPAD_RIGHT",
    "dpad_center": "KEYCODE_DPAD_CENTER",
    "app_switch": "KEYCODE_APP_SWITCH",
    "camera": "KEYCODE_CAMERA",
}


def find_adb() -> str | None:
    """Find the adb executable on the system.

    Checks PATH first, then known Android SDK locations.
    Returns the absolute path or None.
    """
    # Check PATH
    adb_path = shutil.which("adb")
    if adb_path:
        return os.path.abspath(adb_path)

    # Check known locations on Windows
    known_locations = []
    local_app_data = os.environ.get("LOCALAPPDATA", "")
    if local_app_data:
        known_locations.append(
            os.path.join(local_app_data, "Android", "Sdk", "platform-tools", "adb.exe")
        )

    android_home = os.environ.get("ANDROID_HOME", "")
    if android_home:
        known_locations.append(
            os.path.join(android_home, "platform-tools", "adb.exe")
        )

    known_locations.extend([
        r"C:\Android\platform-tools\adb.exe",
        r"C:\android-sdk\platform-tools\adb.exe",
        os.path.expanduser(r"~\AppData\Local\Android\Sdk\platform-tools\adb.exe"),
    ])

    for loc in known_locations:
        if os.path.isfile(loc):
            return os.path.abspath(loc)

    return None


def _run_adb(
    args: list[str],
    device: str | None = None,
    timeout: int = 15,
    binary: bool = False,
) -> subprocess.CompletedProcess:
    """Run an adb command and return the result.

    Args:
        args: Command arguments after 'adb' (e.g., ['shell', 'input', 'tap', '100', '200']).
        device: Optional device serial to target.
        timeout: Command timeout in seconds.
        binary: If True, capture stdout as bytes (for screencap).

    Raises:
        RuntimeError: If adb is not found on the system.
    """
    adb_path = find_adb()
    if not adb_path:
        raise RuntimeError(
            "ADB not found. Install Android SDK platform-tools and ensure "
            "adb is on PATH, or set ANDROID_HOME environment variable."
        )

    cmd = [adb_path]
    if device:
        cmd.extend(["-s", device])
    cmd.extend(args)

    return subprocess.run(
        cmd,
        capture_output=True,
        text=not binary,
        timeout=timeout,
    )


def resolve_device(device: str | None = None) -> str:
    """Resolve which device to target.

    If device is provided, validates it exists.
    If None, auto-selects the single connected device.

    Raises:
        RuntimeError: If no devices, multiple devices without selection,
            or specified device not found.
    """
    result = _run_adb(["devices"])
    if result.returncode != 0:
        raise RuntimeError(f"adb devices failed: {result.stderr.strip()}")

    # Parse device list (skip header line)
    devices = []
    for line in result.stdout.strip().splitlines()[1:]:
        parts = line.split()
        if len(parts) >= 2:
            serial, state = parts[0], parts[1]
            if state == "device":
                devices.append(serial)

    if device:
        if device in devices:
            return device
        # Check if it's connected but in another state
        all_serials = []
        for line in result.stdout.strip().splitlines()[1:]:
            parts = line.split()
            if len(parts) >= 2:
                all_serials.append(f"{parts[0]} ({parts[1]})")
        raise RuntimeError(
            f"Device '{device}' not found or not ready. "
            f"Connected devices: {', '.join(all_serials) or 'none'}"
        )

    if len(devices) == 0:
        raise RuntimeError("No ADB devices connected.")

    if len(devices) == 1:
        return devices[0]

    raise RuntimeError(
        f"Multiple devices connected. Specify a device serial.\n"
        f"Connected: {', '.join(devices)}"
    )


# --- Device Management ---


def list_devices() -> list[AdbDevice]:
    """List all connected ADB devices with their details."""
    result = _run_adb(["devices"])
    if result.returncode != 0:
        raise RuntimeError(f"adb devices failed: {result.stderr.strip()}")

    devices = []
    for line in result.stdout.strip().splitlines()[1:]:
        parts = line.split()
        if len(parts) >= 2:
            serial, state = parts[0], parts[1]
            model = ""
            android_version = ""
            screen_size = ""

            if state == "device":
                model = _getprop(serial, "ro.product.model")
                android_version = _getprop(serial, "ro.build.version.release")
                screen_size = _get_screen_size(serial)

            devices.append(AdbDevice(
                serial=serial,
                state=state,
                model=model,
                android_version=android_version,
                screen_size=screen_size,
            ))

    return devices


def get_device_info(device: str | None = None) -> AdbDevice:
    """Get detailed info for a single device."""
    serial = resolve_device(device)
    return AdbDevice(
        serial=serial,
        state="device",
        model=_getprop(serial, "ro.product.model"),
        android_version=_getprop(serial, "ro.build.version.release"),
        screen_size=_get_screen_size(serial),
    )


def _getprop(serial: str, prop: str) -> str:
    """Query a system property from a device."""
    try:
        result = _run_adb(["shell", "getprop", prop], device=serial, timeout=5)
        return result.stdout.strip() if result.returncode == 0 else ""
    except (subprocess.TimeoutExpired, RuntimeError):
        return ""


def _get_screen_size(serial: str) -> str:
    """Get the screen resolution of a device."""
    try:
        result = _run_adb(["shell", "wm", "size"], device=serial, timeout=5)
        if result.returncode == 0:
            # Output: "Physical size: 1080x1920"
            for line in result.stdout.strip().splitlines():
                if "size" in line.lower():
                    parts = line.split(":")
                    if len(parts) >= 2:
                        return parts[-1].strip()
        return ""
    except (subprocess.TimeoutExpired, RuntimeError):
        return ""


# --- Screenshot ---


def capture_device_screen(device: str | None = None) -> tuple[bytes, int, int]:
    """Capture the device screen and return PNG bytes with dimensions.

    Returns:
        Tuple of (png_bytes, width, height).

    Raises:
        RuntimeError: If capture fails.
    """
    serial = resolve_device(device)
    result = _run_adb(
        ["exec-out", "screencap", "-p"],
        device=serial,
        timeout=10,
        binary=True,
    )

    if result.returncode != 0:
        stderr = result.stderr if isinstance(result.stderr, str) else result.stderr.decode("utf-8", errors="replace")
        raise RuntimeError(f"screencap failed: {stderr.strip()}")

    png_data = result.stdout
    if not png_data or len(png_data) < 100:
        raise RuntimeError("screencap returned empty or invalid data")

    # Parse with PIL to get dimensions
    img = Image.open(io.BytesIO(png_data))
    width, height = img.size

    return (png_data, width, height)


# --- UI Automation ---


def tap(x: int, y: int, device: str | None = None) -> str:
    """Tap a point on the device screen."""
    serial = resolve_device(device)
    result = _run_adb(["shell", "input", "tap", str(x), str(y)], device=serial)
    if result.returncode != 0:
        return f"Tap failed: {result.stderr.strip()}"
    return f"Tapped ({x}, {y}) on {serial}"


def swipe(
    x1: int, y1: int, x2: int, y2: int,
    duration_ms: int = 300,
    device: str | None = None,
) -> str:
    """Swipe from one point to another."""
    serial = resolve_device(device)
    result = _run_adb(
        ["shell", "input", "swipe", str(x1), str(y1), str(x2), str(y2), str(duration_ms)],
        device=serial,
    )
    if result.returncode != 0:
        return f"Swipe failed: {result.stderr.strip()}"
    return f"Swiped ({x1},{y1}) -> ({x2},{y2}) over {duration_ms}ms on {serial}"


def type_text(text: str, device: str | None = None) -> str:
    """Type text on the device. Handles escaping for adb shell."""
    serial = resolve_device(device)

    # Escape for adb shell input text
    escaped = _escape_adb_text(text)

    result = _run_adb(["shell", "input", "text", escaped], device=serial)
    if result.returncode != 0:
        return f"Type failed: {result.stderr.strip()}"
    return f"Typed text on {serial}"


def key_event(key: str, device: str | None = None) -> str:
    """Send a key event to the device.

    Accepts key aliases (home, back, enter, etc.) or full
    KEYCODE_* names or numeric codes.
    """
    serial = resolve_device(device)

    # Resolve alias
    keycode = KEY_ALIASES.get(key.lower(), key)

    result = _run_adb(["shell", "input", "keyevent", keycode], device=serial)
    if result.returncode != 0:
        return f"Key event failed: {result.stderr.strip()}"
    return f"Sent key '{key}' ({keycode}) on {serial}"


def _escape_adb_text(text: str) -> str:
    """Escape text for adb shell input text command."""
    # Characters that need escaping for shell
    special = set('&|;<>()$`\\"\'{} !~')
    escaped = []
    for ch in text:
        if ch == " ":
            escaped.append("%s")
        elif ch in special:
            escaped.append(f"\\{ch}")
        else:
            escaped.append(ch)
    return "".join(escaped)


# --- App Management ---


def install_apk(apk_path: str, device: str | None = None) -> str:
    """Install an APK on the device."""
    serial = resolve_device(device)

    if not os.path.isfile(apk_path):
        return f"APK file not found: {apk_path}"

    result = _run_adb(
        ["install", "-r", apk_path],
        device=serial,
        timeout=60,
    )
    if result.returncode != 0:
        return f"Install failed: {result.stderr.strip()}"
    return f"Installed {os.path.basename(apk_path)} on {serial}\n{result.stdout.strip()}"


def list_packages(filter_text: str = "", device: str | None = None) -> list[str]:
    """List installed packages, optionally filtered."""
    serial = resolve_device(device)
    result = _run_adb(["shell", "pm", "list", "packages"], device=serial, timeout=15)
    if result.returncode != 0:
        raise RuntimeError(f"list packages failed: {result.stderr.strip()}")

    packages = []
    for line in result.stdout.strip().splitlines():
        # Lines look like "package:com.example.app"
        pkg = line.replace("package:", "").strip()
        if pkg:
            if not filter_text or filter_text.lower() in pkg.lower():
                packages.append(pkg)

    packages.sort()
    return packages


def launch_app(package: str, activity: str | None = None, device: str | None = None) -> str:
    """Launch an app on the device."""
    serial = resolve_device(device)

    if activity:
        result = _run_adb(
            ["shell", "am", "start", "-n", f"{package}/{activity}"],
            device=serial,
        )
    else:
        # Use monkey to launch the default activity
        result = _run_adb(
            ["shell", "monkey", "-p", package, "-c",
             "android.intent.category.LAUNCHER", "1"],
            device=serial,
        )

    if result.returncode != 0:
        return f"Launch failed: {result.stderr.strip()}"
    return f"Launched {package} on {serial}"


def stop_app(package: str, device: str | None = None) -> str:
    """Force-stop an app on the device."""
    serial = resolve_device(device)
    result = _run_adb(
        ["shell", "am", "force-stop", package],
        device=serial,
    )
    if result.returncode != 0:
        return f"Stop failed: {result.stderr.strip()}"
    return f"Stopped {package} on {serial}"


# --- File Transfer ---


def push_file(local_path: str, remote_path: str, device: str | None = None) -> str:
    """Push a file from the local machine to the device."""
    serial = resolve_device(device)

    if not os.path.exists(local_path):
        return f"Local path not found: {local_path}"

    result = _run_adb(
        ["push", local_path, remote_path],
        device=serial,
        timeout=60,
    )
    if result.returncode != 0:
        return f"Push failed: {result.stderr.strip()}"
    return f"Pushed {local_path} -> {remote_path} on {serial}\n{result.stdout.strip()}"


def pull_file(remote_path: str, local_path: str, device: str | None = None) -> str:
    """Pull a file from the device to the local machine."""
    serial = resolve_device(device)

    result = _run_adb(
        ["pull", remote_path, local_path],
        device=serial,
        timeout=60,
    )
    if result.returncode != 0:
        return f"Pull failed: {result.stderr.strip()}"
    return f"Pulled {remote_path} -> {local_path} from {serial}\n{result.stdout.strip()}"
