"""
com_renderer.py
===============
Refined high-fidelity slide export bridge for Microsoft PowerPoint COM automation.

This module uses the native PowerPoint rendering engine through pywin32 COM automation.
It is intended for Windows hosts where Microsoft PowerPoint is installed.
"""

from __future__ import annotations

import gc
import logging
import os
import subprocess
import sys
import tempfile
import threading
import time
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterator, List, Optional, Sequence, Tuple, Union

from PIL import Image

log = logging.getLogger("pptx_com_renderer")

_POWERPOINT_COM_AVAILABLE: Optional[bool] = None
_AVAILABILITY_LOCK = threading.Lock()

# 0x80010106
_RPC_E_CHANGED_MODE = -2147417850

# PowerPoint / Office automation constants.
_PP_ALERTS_NONE = 1
_MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

try:
    _LANCZOS = Image.Resampling.LANCZOS
except AttributeError:
    _LANCZOS = getattr(Image, "LANCZOS", getattr(Image, "ANTIALIAS", 1))


class PowerPointCOMError(RuntimeError):
    """Base exception for PowerPoint COM rendering failures."""


class PowerPointNotAvailable(PowerPointCOMError):
    """PowerPoint COM automation could not be established."""


class SlideExportError(PowerPointCOMError):
    """A slide could not be exported."""


def _is_windows() -> bool:
    return sys.platform == "win32"


@contextmanager
def _com_scope() -> Iterator[None]:
    """
    Initializes COM for PowerPoint automation on the current thread.

    PowerPoint automation normally requires an STA thread. If the current thread
    has already been initialized with a different apartment model, this raises
    a clear error instead of proceeding in an undefined state.
    """
    if not _is_windows():
        raise PowerPointNotAvailable(
            "PowerPoint COM automation is only supported on Windows."
        )

    import pythoncom

    initialized = False
    try:
        try:
            pythoncom.CoInitialize()
            initialized = True
        except Exception as exc:
            if getattr(exc, "hresult", None) == _RPC_E_CHANGED_MODE:
                raise PowerPointCOMError(
                    "COM is already initialized on this thread with a different apartment model. "
                    "PowerPoint automation requires an STA thread. Run the export in a dedicated "
                    "STA thread or avoid initializing COM as MTA on this thread."
                ) from exc
            raise

        yield
    finally:
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _iter_registry_access() -> Iterator[int]:
    """
    Yields registry access masks useful for locating Office installations.

    This includes both default registry access and WOW64 views where available,
    improving detection on mixed 32-bit/64-bit environments.
    """
    import winreg

    base = winreg.KEY_READ
    seen = {base}
    yield base

    for name in ("KEY_WOW64_64KEY", "KEY_WOW64_32KEY"):
        flag = getattr(winreg, name, 0)
        if not flag:
            continue

        access = base | flag
        if access not in seen:
            seen.add(access)
            yield access


def _reg_default_value(root: int, subkey: str, access: int) -> Optional[str]:
    """Reads the default value of a registry key as an expanded string."""
    import winreg

    try:
        with winreg.OpenKey(root, subkey, 0, access) as key:
            value, _ = winreg.QueryValueEx(key, "")
            if isinstance(value, str) and value.strip():
                return winreg.ExpandEnvironmentStrings(value.strip())
    except Exception:
        return None

    return None


def _exe_from_command_line(command: str) -> Optional[str]:
    """
    Extracts an executable path from a registry command line.

    Handles values such as:
        "C:\\Program Files\\...\\POWERPNT.EXE" /Automation
    """
    command = command.strip()
    if not command:
        return None

    candidate = command

    if candidate.startswith('"'):
        end = candidate.find('"', 1)
        if end != -1:
            candidate = candidate[1:end]
        else:
            candidate = candidate.strip('"')
    else:
        for separator in (" /", " -", "\t/", "\t-"):
            index = candidate.find(separator)
            if index != -1:
                candidate = candidate[:index]
                break

    candidate = os.path.expandvars(candidate.strip().strip('"'))
    if candidate and os.path.isfile(candidate):
        return candidate

    return None


def _powerpoint_clsid() -> Optional[str]:
    """Resolves the registered CLSID for PowerPoint.Application, if available."""
    if not _is_windows():
        return None

    import winreg

    subkey = r"PowerPoint.Application\CLSID"
    for access in _iter_registry_access():
        value = _reg_default_value(winreg.HKEY_CLASSES_ROOT, subkey, access)
        if value:
            return value.strip().strip("{}")

    return None


def find_powerpoint_executable() -> Optional[str]:
    """
    Locates POWERPNT.EXE using registry information and common install paths.
    """
    if not _is_windows():
        return None

    try:
        import winreg  # noqa: F401
    except Exception:
        return None

    app_paths_subkey = (
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE"
    )

    # 1. App Paths registry entries.
    for access in _iter_registry_access():
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            value = _reg_default_value(root, app_paths_subkey, access)
            if not value:
                continue

            exe = _exe_from_command_line(value)
            if exe:
                return exe

            if os.path.isfile(value):
                return value

    # 2. LocalServer32 registration for PowerPoint.Application.
    clsids = ["91493441-5A91-11CF-8700-00AA0060263B"]
    dynamic_clsid = _powerpoint_clsid()
    if dynamic_clsid and dynamic_clsid not in clsids:
        clsids.insert(0, dynamic_clsid)

    for clsid in clsids:
        subkey = rf"CLSID\{{{clsid}}}\LocalServer32"
        for access in _iter_registry_access():
            value = _reg_default_value(winreg.HKEY_CLASSES_ROOT, subkey, access)
            if not value:
                continue

            exe = _exe_from_command_line(value)
            if exe:
                return exe

    # 3. Common installation locations.
    candidates = [
        r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\root\Office15\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office15\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office15\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\Office14\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office14\POWERPNT.EXE",
    ]

    for candidate in candidates:
        expanded = os.path.expandvars(candidate)
        if os.path.isfile(expanded):
            return expanded

    return None


def _powerpoint_com_registered() -> bool:
    """Returns True if PowerPoint.Application appears to be COM-registered."""
    if not _is_windows():
        return False

    try:
        import winreg
    except Exception:
        return False

    subkey = r"PowerPoint.Application\CLSID"
    for access in _iter_registry_access():
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, subkey, 0, access):
                return True
        except Exception:
            pass

    return False


def _quick_check_powerpoint() -> bool:
    """
    Performs a lightweight availability check without launching PowerPoint.

    This checks imports, registry registration, and known executable locations.
    It does not prove that PowerPoint can successfully render a presentation.
    """
    if not _is_windows():
        return False

    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except Exception as exc:
        log.debug("pywin32 import failed: %s", exc)
        return False

    return bool(find_powerpoint_executable() or _powerpoint_com_registered())


def _process_exists(image_name: str) -> bool:
    """
    Best-effort check for a running Windows process by image name.

    This is used conservatively to avoid claiming ownership of a PowerPoint
    process that may already belong to the interactive user.
    """
    if not _is_windows():
        return False

    try:
        output = subprocess.check_output(
            ["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/NH"],
            stderr=subprocess.DEVNULL,
            text=True,
        )
        return image_name.lower() in output.lower()
    except Exception:
        return False


def _get_pid_from_hwnd(hwnd: Any) -> Optional[int]:
    """Resolves a process ID from a window handle."""
    try:
        hwnd_int = int(hwnd)
    except Exception:
        return None

    if not hwnd_int:
        return None

    try:
        import ctypes

        pid = ctypes.c_ulong()
        ctypes.windll.user32.GetWindowThreadProcessId(hwnd_int, ctypes.byref(pid))
        value = int(pid.value)
        return value or None
    except Exception:
        return None


def _terminate_powerpoint_process(
    pid: Optional[int], process: Optional[subprocess.Popen]
) -> None:
    """Terminates a PowerPoint process that this renderer owns."""
    if process is not None:
        try:
            process.kill()
        except Exception:
            pass

    if pid:
        try:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
        except Exception:
            pass


def _format_com_error(exc: Exception) -> str:
    """
    Formats a COM exception with an HRESULT where available.
    """
    hresult = getattr(exc, "hresult", None)
    if hresult is None:
        return str(exc)

    try:
        code = int(hresult)
        return f"{exc} [HRESULT=0x{code & 0xFFFFFFFF:08X}]"
    except Exception:
        return str(exc)


def _short_path(path: str) -> str:
    """
    Attempts to convert a Windows path to its 8.3 short form.

    Falls back to the original path if short names are unavailable or disabled.
    """
    if not _is_windows():
        return path

    try:
        import ctypes

        buffer = ctypes.create_unicode_buffer(32768)
        result = ctypes.windll.kernel32.GetShortPathNameW(path, buffer, len(buffer))

        if 0 < result < len(buffer):
            return buffer.value
    except Exception as exc:
        log.debug("Short path conversion failed for %s: %s", path, exc)

    return path


@dataclass
class _PowerPointSession:
    """Internal representation of a PowerPoint COM session."""

    app: Any
    owns_process: bool
    process: Optional[subprocess.Popen] = None
    pid: Optional[int] = None


def _wrap_powerpoint_app(
    app: Any,
    owns_process: bool,
    process: Optional[subprocess.Popen] = None,
) -> _PowerPointSession:
    """Wraps a PowerPoint Application object with ownership metadata."""
    pid = None

    if owns_process:
        try:
            pid = _get_pid_from_hwnd(getattr(app, "HWND", 0))
        except Exception:
            pid = None

        if pid is None and process is not None:
            pid = getattr(process, "pid", None)

    return _PowerPointSession(
        app=app,
        owns_process=owns_process,
        process=process,
        pid=pid,
    )


def _acquire_powerpoint(
    *,
    reuse_running: bool = False,
    timeout: float = 30.0,
) -> _PowerPointSession:
    """
    Acquires a PowerPoint Application object.

    If reuse_running is False, the function attempts to create or use an instance
    that it can safely own and terminate. If a user instance is already running
    and an isolated instance cannot be obtained, it reuses the running instance
    but does not mark it as owned.

    If reuse_running is True, the function only attempts to attach to an existing
    instance and does not terminate it.
    """
    if not _is_windows():
        raise PowerPointNotAvailable(
            "PowerPoint COM automation is only supported on Windows."
        )

    import win32com.client

    last_error: Optional[Exception] = None

    if reuse_running:
        try:
            app = win32com.client.GetActiveObject("PowerPoint.Application")
            return _wrap_powerpoint_app(app, owns_process=False)
        except Exception as exc:
            last_error = exc
            log.debug("GetActiveObject failed: %s", exc)

        if _process_exists("POWERPNT.EXE"):
            try:
                app = win32com.client.Dispatch("PowerPoint.Application")
                return _wrap_powerpoint_app(app, owns_process=False)
            except Exception as exc:
                last_error = exc
                log.debug("Dispatch to running PowerPoint failed: %s", exc)

        raise PowerPointNotAvailable(
            "No running PowerPoint application could be attached."
        ) from last_error

    already_running = _process_exists("POWERPNT.EXE")

    # Prefer a dedicated instance. If PowerPoint is already running, do not assume
    # ownership unless we are reasonably sure no pre-existing process was present.
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        return _wrap_powerpoint_app(app, owns_process=not already_running)
    except Exception as exc:
        last_error = exc
        log.debug("DispatchEx failed: %s", exc)

    # If an instance is already active, avoid hijacking it.
    try:
        app = win32com.client.GetActiveObject("PowerPoint.Application")
        log.warning(
            "Isolated PowerPoint instance unavailable; reusing running instance "
            "and will not terminate it."
        )
        return _wrap_powerpoint_app(app, owns_process=False)
    except Exception:
        pass

    already_running = _process_exists("POWERPNT.EXE")

    if already_running:
        try:
            app = win32com.client.Dispatch("PowerPoint.Application")
            return _wrap_powerpoint_app(app, owns_process=False)
        except Exception as exc:
            last_error = exc
            log.debug("Dispatch to already-running PowerPoint failed: %s", exc)

        raise PowerPointNotAvailable(
            "A PowerPoint process appears to be running, but COM connection failed."
        ) from last_error

    # No known running process: Dispatch may launch PowerPoint.
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        return _wrap_powerpoint_app(app, owns_process=True)
    except Exception as exc:
        last_error = exc
        log.debug("Dispatch failed: %s", exc)

    # Click-to-Run and some Office configurations may require explicit bootstrap.
    exe = find_powerpoint_executable()
    if not exe:
        raise PowerPointNotAvailable(
            "PowerPoint executable could not be located."
        ) from last_error

    log.debug("Bootstrapping PowerPoint executable: %s", exe)
    process = subprocess.Popen([exe, "/AUTOMATION"])

    deadline = time.monotonic() + max(1.0, float(timeout))
    while time.monotonic() < deadline:
        for connector in (win32com.client.Dispatch, win32com.client.GetActiveObject):
            try:
                app = connector("PowerPoint.Application")
                return _wrap_powerpoint_app(app, owns_process=True, process=process)
            except Exception:
                pass

        time.sleep(0.25)

    _terminate_powerpoint_process(getattr(process, "pid", None), process)
    raise PowerPointNotAvailable(
        "Timed out while waiting for PowerPoint COM application."
    ) from last_error


def _configure_powerpoint_app(app: Any, force_visible: bool = False) -> None:
    """
    Applies conservative automation settings to PowerPoint.

    DisplayAlerts is suppressed to avoid modal hangs. AutomationSecurity is set
    to force-disable where supported.

    If force_visible is True, Application.Visible is set to True. This is a
    diagnostic option because some PowerPoint builds cannot render/export
    reliably while invisible.
    """
    try:
        app.DisplayAlerts = _PP_ALERTS_NONE
    except Exception as exc:
        log.debug("Could not set DisplayAlerts: %s", exc)

    try:
        app.AutomationSecurity = _MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    except Exception as exc:
        log.debug("Could not set AutomationSecurity: %s", exc)

    if force_visible:
        try:
            app.Visible = True
        except Exception as exc:
            log.debug("Could not set Application.Visible=True: %s", exc)


@contextmanager
def _powerpoint_session(
    *,
    reuse_running: bool = False,
    timeout: float = 30.0,
    force_visible: bool = False,
) -> Iterator[_PowerPointSession]:
    """
    Provides a scoped PowerPoint COM session.

    The session is responsible for COM initialization and for terminating the
    PowerPoint process only when the renderer owns that process.
    """
    with _com_scope():
        session = _acquire_powerpoint(reuse_running=reuse_running, timeout=timeout)

        try:
            _configure_powerpoint_app(session.app, force_visible=force_visible)
            yield session
        finally:
            if session.owns_process:
                try:
                    session.app.Quit()

                    if session.process is not None:
                        try:
                            session.process.wait(timeout=5.0)
                        except subprocess.TimeoutExpired:
                            _terminate_powerpoint_process(session.pid, session.process)
                except Exception as exc:
                    log.debug("PowerPoint Quit failed; attempting termination: %s", exc)
                    _terminate_powerpoint_process(session.pid, session.process)

            session.app = None
            gc.collect()


def is_powerpoint_com_available(
    force_recheck: bool = False,
    launch_test: bool = False,
) -> bool:
    """
    Checks whether PowerPoint COM automation appears available.

    By default this is a lightweight check. Set launch_test=True to perform an
    actual PowerPoint connection test. The launch test will not terminate a
    pre-existing user-owned PowerPoint instance.
    """
    global _POWERPOINT_COM_AVAILABLE

    with _AVAILABILITY_LOCK:
        if _POWERPOINT_COM_AVAILABLE is not None and not force_recheck:
            return _POWERPOINT_COM_AVAILABLE

        if not _quick_check_powerpoint():
            _POWERPOINT_COM_AVAILABLE = False
            return False

        if not launch_test:
            _POWERPOINT_COM_AVAILABLE = True
            return True

        try:
            with _powerpoint_session(reuse_running=False, timeout=20.0) as session:
                _ = session.app.Version

            _POWERPOINT_COM_AVAILABLE = True
        except Exception as exc:
            log.debug("PowerPoint COM launch test failed: %s", exc)
            _POWERPOINT_COM_AVAILABLE = False

        return _POWERPOINT_COM_AVAILABLE


def _remove_quiet(path: str) -> None:
    try:
        os.remove(path)
    except Exception:
        pass


def _wait_for_nonempty_file(path: str, timeout: float = 5.0) -> bool:
    """Waits until a file exists and has non-zero size."""
    deadline = time.monotonic() + max(0.25, timeout)

    while time.monotonic() < deadline:
        try:
            if os.path.isfile(path) and os.path.getsize(path) > 0:
                return True
        except OSError:
            pass

        time.sleep(0.05)

    return False


def _load_image(
    path: str,
    image_mode: Optional[str],
    target_size: Tuple[int, int],
) -> Image.Image:
    """
    Loads an exported image into memory.

    If image_mode is provided, the image is converted. If the exported image
    does not match the requested target size, it is resized using Lanczos
    resampling.
    """
    with Image.open(path) as img:
        if image_mode:
            img = img.convert(image_mode)
        else:
            img = img.copy()

        if img.size != target_size:
            img = img.resize(target_size, _LANCZOS)

    return img


def _try_export_slide(
    slide: Any,
    target_png: str,
    width: int,
    height: int,
    *,
    allow_width_only: bool = True,
    allow_default_scale: bool = True,
) -> Tuple[Optional[str], List[str]]:
    """
    Attempts to export a slide using progressively more permissive fallbacks.

    Returns:
        (method, errors)

        method is None if all attempts failed.
        errors contains detailed attempt failures.
    """
    errors: List[str] = []

    attempts: List[Tuple[Optional[int], Optional[int]]] = [(width, height)]

    if allow_width_only:
        attempts.append((width, None))

    if allow_default_scale:
        attempts.append((None, None))

    for attempt_width, attempt_height in attempts:
        label = f"Slide.Export(width={attempt_width}, height={attempt_height})"

        try:
            if attempt_width is not None and attempt_height is not None:
                slide.Export(target_png, "PNG", attempt_width, attempt_height)
            elif attempt_width is not None:
                slide.Export(target_png, "PNG", attempt_width)
            else:
                slide.Export(target_png, "PNG")

            if _wait_for_nonempty_file(target_png, timeout=15.0):
                return label, []

            errors.append(
                f"{label}: PowerPoint returned without an exception, "
                f"but the output file was missing or zero-length."
            )

        except Exception as exc:
            formatted = _format_com_error(exc)
            errors.append(f"{label}: {formatted}")
            log.debug("Slide export attempt failed: %s", formatted)

        _remove_quiet(target_png)

    return None, errors


def _try_shape_range_export(
    slide: Any,
    target_png: str,
    width: int,
    height: int,
) -> Tuple[Optional[str], List[str]]:
    """
    Optional fallback that exports the slide's shape range.

    This may omit backgrounds or other presentation-level rendering context.
    It is intended primarily as a diagnostic fallback.

    Returns:
        (method, errors)
    """
    errors: List[str] = []

    try:
        shape_range = slide.Shapes.Range()
    except Exception as exc:
        errors.append(f"Shapes.Range: {_format_com_error(exc)}")
        return None, errors

    for filter_name in (2, "PNG"):
        label = f"ShapeRange.Export(filter={filter_name!r})"

        try:
            shape_range.Export(target_png, filter_name, width, height)

            if _wait_for_nonempty_file(target_png, timeout=15.0):
                return label, []

            errors.append(
                f"{label}: PowerPoint returned without an exception, "
                f"but the output file was missing or zero-length."
            )

        except Exception as exc:
            errors.append(f"{label}: {_format_com_error(exc)}")

        _remove_quiet(target_png)

    return None, errors


def _try_alternate_filter_export(
    slide: Any,
    target_png: str,
    width: int,
    height: int,
) -> Tuple[Optional[str], List[str]]:
    """
    Attempts to export using alternate raster filters such as JPG or BMP.

    This is useful when the PNG export filter fails because of an Office
    installation issue or a filter-specific rendering problem.

    Returns:
        (alternate_file_path, errors)
    """
    errors: List[str] = []
    base = Path(target_png)

    filters: Tuple[Tuple[str, str], ...] = (
        ("JPG", ".jpg"),
        ("JPEG", ".jpg"),
        ("BMP", ".bmp"),
    )

    seen_paths = set()

    for filter_name, suffix in filters:
        alt_path = base.with_suffix(suffix)
        alt_path_str = str(alt_path)

        if alt_path_str in seen_paths:
            continue

        seen_paths.add(alt_path_str)

        label = f"Slide.Export(filter={filter_name!r}, width={width}, height={height})"

        try:
            slide.Export(alt_path_str, filter_name, width, height)

            if _wait_for_nonempty_file(alt_path_str, timeout=15.0):
                return alt_path_str, []

            errors.append(
                f"{label}: PowerPoint returned without an exception, "
                f"but the output file was missing or zero-length."
            )

        except Exception as exc:
            errors.append(f"{label}: {_format_com_error(exc)}")

        _remove_quiet(alt_path_str)

    return None, errors


def _open_presentation(
    app: Any,
    abs_path: str,
    *,
    open_with_window: Optional[bool] = None,
    force_visible: bool = False,
) -> Any:
    """
    Opens a presentation in PowerPoint.

    open_with_window:
        None  - automatic order. Prefers headless unless force_visible is True.
        True  - prefers WithWindow=True.
        False - attempts WithWindow=False only.
    """
    if open_with_window is True:
        order: Tuple[bool, ...] = (True, False)
    elif open_with_window is False:
        order = (False,)
    else:
        order = (True, False) if force_visible else (False, True)

    last_exc: Optional[Exception] = None

    for with_window in order:
        try:
            return app.Presentations.Open(abs_path, True, False, with_window)
        except Exception as exc:
            last_exc = exc
            log.debug(
                "Presentations.Open failed with WithWindow=%s: %s", with_window, exc
            )

    raise PowerPointCOMError("Unable to open presentation in PowerPoint.") from last_exc


def export_pptx_slides_com(
    pptx_path: Union[str, Path],
    output_dir: Optional[Union[str, Path]] = None,
    width: int = 1280,
    slide_numbers: Optional[Sequence[int]] = None,
    *,
    height: Optional[int] = None,
    image_mode: Optional[str] = "RGB",
    timeout: float = 30.0,
    reuse_running: bool = False,
    strict_slide_numbers: bool = True,
    allow_default_scale_fallback: bool = True,
    allow_shape_range_fallback: bool = False,
    allow_alternate_filter_fallback: bool = False,
    force_visible: bool = False,
    open_with_window: Optional[bool] = None,
    use_short_path: bool = False,
    output_file_prefix: str = "slide_",
) -> List[Image.Image]:
    """
    Exports slides from a PPTX file using PowerPoint COM automation.

    Args:
        pptx_path:
            Path to the .pptx file.

        output_dir:
            Optional directory where exported PNG files are persisted.
            If omitted, temporary files are used and removed after loading.

        width:
            Desired output width in pixels.

        slide_numbers:
            Optional 1-based slide numbers to export. If omitted, all slides
            are exported.

        height:
            Optional explicit output height. If omitted, height is derived
            from the presentation's slide aspect ratio.

        image_mode:
            Pillow image mode for returned images. Defaults to "RGB".
            Use None to preserve the mode produced by Pillow, or "RGBA"
            where alpha preservation is desired.

        timeout:
            Timeout for acquiring PowerPoint.

        reuse_running:
            If True, attach to a running PowerPoint instance if possible.
            The renderer will not terminate a user-owned PowerPoint instance.

        strict_slide_numbers:
            If True, requested slide numbers outside the valid range raise
            ValueError. If False, invalid slide numbers are skipped and logged.

        allow_default_scale_fallback:
            Permits default-scale export if exact-size export fails. The result
            is resized to the requested target dimensions.

        allow_shape_range_fallback:
            Permits shape-range export. Disabled by default because it may not
            include slide backgrounds or full presentation rendering context.

        allow_alternate_filter_fallback:
            Permits export via JPG/BMP if PNG export fails. If output_dir is
            provided, the final file is still written as PNG.

        force_visible:
            Diagnostic option that sets Application.Visible = True. Some
            PowerPoint builds require a visible application to export reliably.

        open_with_window:
            Controls Presentations.Open WithWindow behavior.
            None = automatic, True = prefer window, False = headless only.

        use_short_path:
            Attempts to use Windows 8.3 short paths for source and output
            directories. This can help with Unicode or long-path issues.

        output_file_prefix:
            File name prefix used when output_dir is provided.

    Returns:
        List of Pillow Image.Image instances.

    Raises:
        PowerPointNotAvailable:
            PowerPoint COM automation could not be established.

        SlideExportError:
            A slide could not be exported.

        FileNotFoundError:
            The source presentation does not exist.

        ValueError:
            Invalid arguments were supplied.
    """
    if not _is_windows():
        raise PowerPointNotAvailable(
            "PowerPoint COM export is only supported on Windows."
        )

    try:
        width = int(width)
    except Exception as exc:
        raise ValueError("width must be an integer.") from exc

    if width <= 0:
        raise ValueError("width must be greater than zero.")

    explicit_height = height is not None
    if explicit_height:
        try:
            height = int(height)
        except Exception as exc:
            raise ValueError("height must be an integer when provided.") from exc

        if height <= 0:
            raise ValueError("height must be greater than zero.")

    requested_indices: Optional[List[int]] = None
    if slide_numbers is not None:
        requested_indices = []
        for item in slide_numbers:
            try:
                requested_indices.append(int(item))
            except Exception as exc:
                raise ValueError(
                    f"slide_numbers must contain integers, got {item!r}."
                ) from exc

    source = Path(pptx_path).expanduser()
    if not source.is_file():
        raise FileNotFoundError(f"Presentation file not found: {source}")

    abs_path = str(source.resolve())
    if use_short_path:
        abs_path = _short_path(abs_path)

    out_dir: Optional[Path] = None
    if output_dir:
        out_dir = Path(output_dir).expanduser().resolve()
        out_dir.mkdir(parents=True, exist_ok=True)

        if use_short_path:
            out_dir = Path(_short_path(str(out_dir)))

    safe_prefix = Path(output_file_prefix).name or "slide_"

    images: List[Image.Image] = []
    temp_files: List[str] = []

    try:
        with _powerpoint_session(
            reuse_running=reuse_running,
            timeout=timeout,
            force_visible=force_visible,
        ) as session:
            app = session.app
            pres: Optional[Any] = None

            try:
                pres = _open_presentation(
                    app,
                    abs_path,
                    open_with_window=open_with_window,
                    force_visible=force_visible,
                )

                total_slides = int(pres.Slides.Count)
                slide_w_pt = float(pres.PageSetup.SlideWidth)
                slide_h_pt = float(pres.PageSetup.SlideHeight)

                if explicit_height:
                    target_height = int(height)
                else:
                    if slide_w_pt > 0:
                        computed_height = int(round(width * slide_h_pt / slide_w_pt))
                    else:
                        computed_height = int(round(width * 9.0 / 16.0))
                    target_height = max(1, computed_height)

                target_size = (width, target_height)

                if requested_indices is None:
                    indices = list(range(1, total_slides + 1))
                else:
                    indices = []
                    seen = set()
                    invalid = []

                    for idx in requested_indices:
                        if idx < 1 or idx > total_slides:
                            invalid.append(idx)
                            continue

                        if idx not in seen:
                            seen.add(idx)
                            indices.append(idx)

                    if invalid:
                        if strict_slide_numbers:
                            raise ValueError(
                                f"Slide numbers out of range 1..{total_slides}: {invalid}"
                            )

                        log.warning(
                            "Skipping out-of-range slide numbers %s; presentation has %d slides.",
                            invalid,
                            total_slides,
                        )

                pad = max(3, len(str(total_slides)))

                for slide_index in indices:
                    slide = None
                    target_png: Optional[str] = None

                    try:
                        if out_dir is not None:
                            target_png = str(
                                out_dir / f"{safe_prefix}{slide_index:0{pad}d}.png"
                            )
                        else:
                            fd, target_png = tempfile.mkstemp(
                                suffix=".png",
                                prefix=f"{safe_prefix}{slide_index:0{pad}d}_",
                            )
                            os.close(fd)
                            temp_files.append(target_png)

                        # Remove stale output before asking PowerPoint to write.
                        _remove_quiet(target_png)

                        slide = pres.Slides(slide_index)

                        method, export_errors = _try_export_slide(
                            slide,
                            target_png,
                            width,
                            target_height,
                            allow_width_only=not explicit_height,
                            allow_default_scale=allow_default_scale_fallback,
                        )

                        alternate_source: Optional[str] = None

                        if method is None and allow_shape_range_fallback:
                            shape_method, shape_errors = _try_shape_range_export(
                                slide,
                                target_png,
                                width,
                                target_height,
                            )
                            export_errors.extend(shape_errors)
                            method = shape_method

                        if method is None and allow_alternate_filter_fallback:
                            alt_path, alt_errors = _try_alternate_filter_export(
                                slide,
                                target_png,
                                width,
                                target_height,
                            )
                            export_errors.extend(alt_errors)

                            if alt_path is not None:
                                method = f"AlternateFilter:{Path(alt_path).suffix}"
                                alternate_source = alt_path

                                if alt_path not in temp_files:
                                    temp_files.append(alt_path)

                        source_path = alternate_source or target_png
                        if source_path is None:
                            raise SlideExportError("No export target was created.")

                        if method is None or not _wait_for_nonempty_file(
                            source_path, timeout=5.0
                        ):
                            details = (
                                "; ".join(export_errors)
                                if export_errors
                                else "unknown COM failure"
                            )
                            raise SlideExportError(
                                f"All enabled PowerPoint export methods failed for slide {slide_index}. "
                                f"Details: {details}"
                            )

                        img = _load_image(source_path, image_mode, target_size)

                        # If an alternate filter was used and the caller requested
                        # persisted output, write the final image as PNG.
                        if alternate_source is not None and out_dir is not None:
                            img.save(target_png, format="PNG")

                        images.append(img)

                    except Exception as exc:
                        raise SlideExportError(
                            f"Failed to export slide {slide_index} from {abs_path}: {exc}"
                        ) from exc

                    finally:
                        slide = None

            finally:
                if pres is not None:
                    try:
                        pres.Saved = True
                    except Exception:
                        pass

                    try:
                        pres.Close()
                    except Exception as exc:
                        log.debug("Presentation close failed: %s", exc)

                    pres = None
                    gc.collect()

    except (PowerPointCOMError, ValueError, FileNotFoundError):
        raise
    except Exception as exc:
        raise PowerPointCOMError(
            f"PowerPoint COM export failed for {abs_path}: {exc}"
        ) from exc
    finally:
        for tmp in temp_files:
            _remove_quiet(tmp)

    return images
