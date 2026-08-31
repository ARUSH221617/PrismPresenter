"""
com_renderer.py
===============
High-fidelity slide export bridge utilizing Microsoft PowerPoint COM Automation
(win32com.client / comtypes).

Provides pixel-perfect slide rendering using the native Office PowerPoint engine,
with fallback detection and safe lifecycle management.
"""

from __future__ import annotations

import os
import sys
import time
import subprocess
import tempfile
import logging
from pathlib import Path
from typing import List, Optional, Sequence, Union

from PIL import Image

log = logging.getLogger("pptx_com_renderer")

_POWERPOINT_COM_AVAILABLE: Optional[bool] = None


def find_powerpoint_executable() -> Optional[str]:
    """Locates the POWERPNT.EXE path via Windows Registry or standard Office directories."""
    if sys.platform != "win32":
        return None

    try:
        import winreg

        # 1. App Paths registry
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE") as key:
                    val, _ = winreg.QueryValueEx(key, "")
                    if val and os.path.isfile(val):
                        return val
            except Exception:
                pass

        # 2. LocalServer32 CLSID
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"CLSID\{91493441-5A91-11CF-8700-00AA0060263B}\LocalServer32") as key:
                val, _ = winreg.QueryValueEx(key, "")
                exe = val.split("/")[0].strip(' "')
                if exe and os.path.isfile(exe):
                    return exe
        except Exception:
            pass
    except Exception:
        pass

    # 3. Known Office installation paths
    candidates = [
        r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE",
        r"C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE",
    ]
    for p in candidates:
        if os.path.isfile(p):
            return p

    return None


def get_powerpoint_application():
    """
    Acquires or starts a PowerPoint COM Application instance reliably.
    Handles Click-to-Run Office CO_E_SERVER_EXEC_FAILURE via /AUTOMATION launch fallback.
    """
    import pythoncom
    import win32com.client

    # 1. Try active running instance
    try:
        return win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception:
        pass

    # 2. Try standard Dispatch
    try:
        return win32com.client.Dispatch("PowerPoint.Application")
    except Exception:
        pass

    # 3. Try DispatchEx
    try:
        return win32com.client.DispatchEx("PowerPoint.Application")
    except Exception:
        pass

    # 4. If Click-to-Run Office prevents out-of-process CoCreateInstance directly,
    # bootstrap POWERPNT.EXE /AUTOMATION and connect to ROT (Running Object Table)
    exe_path = find_powerpoint_executable()
    if exe_path:
        log.debug("Bootstrapping PowerPoint process: %s", exe_path)
        proc = subprocess.Popen([exe_path, "/AUTOMATION"])
        try:
            for _ in range(30):
                time.sleep(0.2)
                try:
                    return win32com.client.Dispatch("PowerPoint.Application")
                except Exception:
                    pass
        finally:
            pass

    raise RuntimeError("Could not connect to PowerPoint COM Application.")


def is_powerpoint_com_available(force_recheck: bool = False) -> bool:
    """
    Checks whether Windows PowerPoint COM automation is functional.
    Result is cached after the first check unless force_recheck=True.
    """
    global _POWERPOINT_COM_AVAILABLE
    if _POWERPOINT_COM_AVAILABLE is not None and not force_recheck:
        return _POWERPOINT_COM_AVAILABLE

    if sys.platform != "win32":
        _POWERPOINT_COM_AVAILABLE = False
        return False

    import gc
    try:
        import pythoncom

        pythoncom.CoInitialize()
        app = None
        try:
            app = get_powerpoint_application()
            _ = app.Visible
            _POWERPOINT_COM_AVAILABLE = True
        finally:
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            del app
            gc.collect()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    except Exception as exc:
        log.debug("PowerPoint COM availability check failed: %s", exc)
        _POWERPOINT_COM_AVAILABLE = False

    return bool(_POWERPOINT_COM_AVAILABLE)


def export_pptx_slides_com(
    pptx_path: Union[str, Path],
    output_dir: Optional[Union[str, Path]] = None,
    width: int = 1280,
    slide_numbers: Optional[Sequence[int]] = None,
) -> List[Image.Image]:
    """
    Exports slides from a PPTX file using PowerPoint COM automation.

    Args:
        pptx_path: Absolute or relative path to the .pptx file.
        output_dir: Optional directory to persist exported PNG files.
        width: Desired output width in pixels (aspect ratio is preserved).
        slide_numbers: Optional 1-based list of slide indices to export.

    Returns:
        List of Pillow Image.Image instances in RGB mode.

    Raises:
        RuntimeError: If COM automation fails or PowerPoint is not available.
        FileNotFoundError: If pptx_path does not exist.
    """
    if sys.platform != "win32":
        raise RuntimeError("PowerPoint COM export is only supported on Windows.")

    abs_path = os.path.abspath(str(pptx_path))
    if not os.path.isfile(abs_path):
        raise FileNotFoundError(f"Presentation file not found: {abs_path}")

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()

    app = None
    pres = None
    images: List[Image.Image] = []
    temp_files: List[str] = []

    try:
        app = get_powerpoint_application()
        # Suppress alerts / prompts to prevent modal dialog hangs
        try:
            app.DisplayAlerts = 1  # ppAlertsNone
        except Exception:
            pass

        # Open presentation in headless background mode
        pres = app.Presentations.Open(
            abs_path,
            ReadOnly=True,
            Untitled=False,
            WithWindow=False,
        )

        total_slides = pres.Slides.Count
        slide_w_pt = float(pres.PageSetup.SlideWidth)
        slide_h_pt = float(pres.PageSetup.SlideHeight)
        aspect_ratio = slide_h_pt / slide_w_pt if slide_w_pt > 0 else 9.0 / 16.0
        height = max(1, int(round(width * aspect_ratio)))

        indices_to_export = (
            [idx for idx in slide_numbers if 1 <= idx <= total_slides]
            if slide_numbers is not None
            else list(range(1, total_slides + 1))
        )

        out_path_obj = Path(output_dir) if output_dir else None
        if out_path_obj:
            out_path_obj.mkdir(parents=True, exist_ok=True)

        for s_idx in indices_to_export:
            slide = pres.Slides(s_idx)
            try:
                if out_path_obj:
                    target_png = os.path.abspath(str(out_path_obj / f"slide_{s_idx:03d}.png"))
                else:
                    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tf:
                        target_png = os.path.abspath(tf.name)
                    temp_files.append(target_png)

                # Export slide via COM API with multi-tier fallback
                exported = False

                # Tier 1: Standard slide.Export with requested resolution
                try:
                    slide.Export(target_png, "PNG", width, height)
                    exported = True
                except Exception as slide_exp_err:
                    log.debug("Standard slide.Export failed for slide %d: %s", s_idx, slide_exp_err)

                # Tier 2: slide.Export with default scale
                if not exported:
                    try:
                        slide.Export(target_png, "PNG")
                        exported = True
                    except Exception:
                        pass

                # Tier 3: In a temporary duplicate slide, sanitize shapes (replace textboxes, touch fonts)
                if not exported:
                    s_dup = None
                    try:
                        s_dup = slide.Duplicate()
                        # Convert msoTextBox (17) into regular borderless/transparent shapes
                        textboxes = [sh for sh in s_dup.Shapes if sh.Type == 17]
                        for tb in textboxes:
                            l, t, w, h = tb.Left, tb.Top, tb.Width, tb.Height
                            new_sh = s_dup.Shapes.AddShape(1, l, t, w, h)  # 1 = msoShapeRectangle
                            new_sh.Fill.Visible = 0
                            new_sh.Line.Visible = 0
                            if tb.HasTextFrame and tb.TextFrame.HasText:
                                new_sh.TextFrame.TextRange.Text = tb.TextFrame.TextRange.Text
                                for p_idx in range(1, tb.TextFrame.TextRange.Paragraphs().Count + 1):
                                    p_old = tb.TextFrame.TextRange.Paragraphs(p_idx)
                                    p_new = new_sh.TextFrame.TextRange.Paragraphs(p_idx)
                                    if p_old.Font.Name:
                                        p_new.Font.Name = p_old.Font.Name
                                    p_new.Font.Size = p_old.Font.Size
                                    p_new.ParagraphFormat.Alignment = p_old.ParagraphFormat.Alignment
                            tb.Delete()

                        # Touch font names and reset AutoSize to avoid engine crashes on complex scripts
                        for sh in s_dup.Shapes:
                            try:
                                if sh.HasTextFrame:
                                    tf2 = sh.TextFrame2
                                    if tf2.AutoSize == 1:
                                        tf2.AutoSize = 0
                                    if tf2.HasText:
                                        for p_idx in range(1, tf2.TextRange.Paragraphs.Count + 1):
                                            para = tf2.TextRange.Paragraphs(p_idx)
                                            fname = para.Font.Name
                                            if fname:
                                                para.Font.Name = fname
                            except Exception:
                                pass

                        s_dup.Export(target_png, "PNG", width, height)
                        exported = True
                    except Exception as dup_err:
                        log.debug("Duplicate slide sanitized export failed for slide %d: %s", s_idx, dup_err)
                    finally:
                        if s_dup is not None:
                            try:
                                s_dup.Delete()
                            except Exception:
                                pass

                # Tier 4: Shapes.Range().Export fallback
                if not exported:
                    try:
                        slide.Shapes.Range().Export(target_png, 2, width, height)
                        exported = True
                    except Exception as range_err:
                        log.debug("ShapeRange export failed for slide %d: %s", s_idx, range_err)

                if not exported or not os.path.exists(target_png):
                    raise RuntimeError(f"All PowerPoint COM export methods failed for slide {s_idx}.")

                # Read image into memory and disconnect file handle
                with Image.open(target_png) as img_disk:
                    img_mem = img_disk.convert("RGB").copy()
                    images.append(img_mem)
            finally:
                del slide

    except Exception as e:
        log.warning("PowerPoint COM export error on %s: %s", abs_path, e)
        raise RuntimeError(f"COM export failed for {abs_path}: {e}") from e

    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass

        import gc
        del pres
        del app
        gc.collect()

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

        # Clean up temporary files
        for tmp in temp_files:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except Exception:
                pass

    return images
