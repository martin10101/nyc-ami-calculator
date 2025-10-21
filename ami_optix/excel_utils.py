"""
Excel helper utilities: primarily used to convert xlsx outputs into xlsb files
so the client receives downloads in the same binary format they upload.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Optional


def _ensure_absolute(path: str) -> str:
    """Return an absolute path for consistent COM interactions."""
    if not os.path.isabs(path):
        return os.path.abspath(path)
    return path


def convert_xlsx_to_xlsb(source_path: str, target_path: Optional[str] = None, delete_source: bool = True) -> str:
    """
    Convert a .xlsx workbook into a .xlsb file using the local Excel installation.

    Parameters
    ----------
    source_path
        Path to the source .xlsx file.
    target_path
        Optional destination. Defaults to the same directory with a .xlsb extension.
    delete_source
        When True (default), remove the original .xlsx after a successful conversion.

    Returns
    -------
    str
        Absolute path to the generated .xlsb file.

    Raises
    ------
    RuntimeError
        If Excel conversion fails or pywin32 is not installed.
    """
    source_abs = _ensure_absolute(source_path)
    if target_path:
        target_abs = _ensure_absolute(target_path)
    else:
        root, _ = os.path.splitext(source_abs)
        target_abs = f"{root}.xlsb"

    if not os.path.exists(source_abs):
        raise RuntimeError(f"Cannot convert missing file: {source_abs}")

    errors = []

    if sys.platform.startswith("win"):
        try:
            return _convert_with_win32(source_abs, target_abs, delete_source=delete_source)
        except RuntimeError as exc:  # pragma: no cover - depends on runtime
            errors.append(str(exc))

    soffice = _locate_soffice_binary()
    if soffice:
        try:
            return _convert_with_soffice(soffice, source_abs, target_abs, delete_source=delete_source)
        except RuntimeError as exc:  # pragma: no cover - depends on runtime
            errors.append(str(exc))

    if errors:
        raise RuntimeError(
            "Unable to convert workbook to XLSB. Attempted methods:\n- "
            + "\n- ".join(errors)
        )

    raise RuntimeError(
        "No XLSB conversion utility found. Install LibreOffice (soffice) or run on a Windows host with Excel."
    )


def _convert_with_win32(source_abs: str, target_abs: str, delete_source: bool) -> str:
    try:
        import win32com.client  # type: ignore
        from win32com.client import constants  # type: ignore
    except ImportError as exc:  # pragma: no cover
        raise RuntimeError("pywin32/Excel automation not available on this host.") from exc

    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(source_abs)
        try:
            workbook.SaveAs(target_abs, FileFormat=constants.xlExcel12)
        finally:
            workbook.Close(SaveChanges=False)
    except Exception as exc:  # pragma: no cover
        raise RuntimeError(f"Excel COM conversion failed: {exc}") from exc
    finally:
        if excel is not None:
            excel.Quit()

    if delete_source:
        try:
            os.remove(source_abs)
        except OSError:
            pass

    return target_abs


def _locate_soffice_binary() -> Optional[str]:
    env_path = os.environ.get("LIBREOFFICE_PATH")
    if env_path and os.path.isfile(env_path) and os.access(env_path, os.X_OK):
        return env_path

    system_path = shutil.which("soffice")
    if system_path:
        return system_path

    project_root = Path(__file__).resolve().parents[1]
    candidates = [
        project_root / "libreoffice-appimage" / "usr" / "bin" / "soffice",
        project_root / "squashfs-root" / "usr" / "bin" / "soffice",
    ]
    for candidate in candidates:
        if candidate.is_file():
            candidate_path = str(candidate)
            if os.access(candidate_path, os.X_OK):
                return candidate_path
    return None


def _convert_with_soffice(soffice_path: str, source_abs: str, target_abs: str, delete_source: bool) -> str:
    output_dir = os.path.dirname(target_abs)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    cmd = [
        soffice_path,
        "--headless",
        "--convert-to",
        "xlsb",
        "--outdir",
        output_dir,
        source_abs,
    ]

    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed (exit {result.returncode}): {result.stderr.decode().strip()}"
        )

    converted_path = os.path.join(
        output_dir,
        f"{os.path.splitext(os.path.basename(source_abs))[0]}.xlsb",
    )
    if not os.path.exists(converted_path):
        raise RuntimeError("LibreOffice conversion did not produce an XLSB file.")

    if os.path.abspath(converted_path) != os.path.abspath(target_abs):
        try:
            if os.path.exists(target_abs):
                os.remove(target_abs)
            os.replace(converted_path, target_abs)
        except OSError as exc:
            raise RuntimeError(f"Failed to move LibreOffice output: {exc}") from exc

    if delete_source:
        try:
            os.remove(source_abs)
        except OSError:
            pass

    return target_abs
