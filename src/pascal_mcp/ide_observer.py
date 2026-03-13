"""IDE Observer for RAD Studio / Delphi / Lazarus.

Watches the IDE window via screenshots and reads project source files
from disk. Claude reads compiler errors visually from the screenshot
and uses file access to understand and fix the code.
"""

from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
from pathlib import Path

from pascal_mcp.screenshot import (
    _find_window_by_title,
    _capture_with_printwindow,
    list_windows,
)


# Patterns to match IDE window titles
IDE_TITLE_PATTERNS = [
    r"Embarcadero RAD Studio",
    r"RAD Studio",
    r"Delphi \d",
    r"Lazarus IDE",
    r"Lazarus v",
]


def find_ide_window() -> dict | None:
    """Find a running Delphi/Lazarus IDE window.

    Returns dict with hwnd, title, and parsed project_name, or None.
    """
    windows = list_windows("")
    for w in windows:
        title = w["title"]
        for pattern in IDE_TITLE_PATTERNS:
            if re.search(pattern, title, re.IGNORECASE):
                project_name = _parse_project_name(title)
                return {
                    "hwnd": w["hwnd"],
                    "title": title,
                    "project_name": project_name,
                }
    return None


def _parse_project_name(title: str) -> str | None:
    """Extract project name from IDE window title.

    RAD Studio: "Embarcadero RAD Studio 12.2 - ProjectName"
    Delphi 7:   "Delphi 7 - ProjectName"
    Lazarus:    "Lazarus IDE v3.0 - ProjectName"
    """
    # Try "- ProjectName" pattern
    match = re.search(r"\s-\s+(.+?)(?:\s*\[|$)", title)
    if match:
        return match.group(1).strip()
    return None


def capture_ide_screenshot(hwnd: int):
    """Capture the IDE window and return a PIL Image."""
    return _capture_with_printwindow(hwnd)


def find_project_files(project_dir: str) -> dict:
    """Scan a Delphi/Lazarus project directory for source files.

    Returns a dict with categorized file lists and parsed project info.
    """
    project_path = Path(project_dir)
    if not project_path.is_dir():
        return {"error": f"Directory not found: {project_dir}"}

    result = {
        "project_dir": str(project_path),
        "dpr_files": [],
        "dproj_files": [],
        "pas_files": [],
        "dfm_files": [],
        "fmx_files": [],
        "lfm_files": [],
        "lpr_files": [],
        "lpi_files": [],
        "units_from_dproj": [],
        "search_paths": [],
        "build_config": None,
    }

    for f in project_path.rglob("*"):
        if f.is_file():
            ext = f.suffix.lower()
            rel = str(f.relative_to(project_path))
            if ext == ".dpr":
                result["dpr_files"].append(rel)
            elif ext == ".dproj":
                result["dproj_files"].append(rel)
            elif ext == ".pas":
                result["pas_files"].append(rel)
            elif ext == ".dfm":
                result["dfm_files"].append(rel)
            elif ext == ".fmx":
                result["fmx_files"].append(rel)
            elif ext == ".lfm":
                result["lfm_files"].append(rel)
            elif ext == ".lpr":
                result["lpr_files"].append(rel)
            elif ext == ".lpi":
                result["lpi_files"].append(rel)

    # Parse .dproj for unit list and build config
    for dproj_rel in result["dproj_files"]:
        dproj_path = project_path / dproj_rel
        try:
            info = _parse_dproj(str(dproj_path))
            result["units_from_dproj"] = info.get("units", [])
            result["search_paths"] = info.get("search_paths", [])
            result["build_config"] = info.get("config")
            break  # Use the first .dproj found
        except Exception:
            pass

    return result


def _parse_dproj(dproj_path: str) -> dict:
    """Parse a .dproj (MSBuild) file for unit references and config."""
    tree = ET.parse(dproj_path)
    root = tree.getroot()

    # Handle MSBuild namespace
    ns = ""
    if root.tag.startswith("{"):
        ns = root.tag.split("}")[0] + "}"

    units = []
    search_paths = []
    config = None

    # Find DCCReference elements (source unit includes)
    for ref in root.iter(f"{ns}DCCReference"):
        include = ref.get("Include", "")
        if include:
            units.append(include)

    # Find search paths
    for prop in root.iter(f"{ns}DCC_UnitSearchPath"):
        if prop.text:
            search_paths.extend(prop.text.split(";"))

    # Find active config
    for prop in root.iter(f"{ns}Config"):
        condition = prop.get("Condition", "")
        if "'$(Config)'" in condition and "==" not in condition:
            config = prop.text
            break
    if not config:
        for prop in root.iter(f"{ns}Config"):
            if prop.text:
                config = prop.text
                break

    return {
        "units": units,
        "search_paths": [p.strip() for p in search_paths if p.strip()],
        "config": config,
    }


def read_source_context(
    file_path: str,
    line: int,
    context_lines: int = 10,
) -> str:
    """Read source file and return lines around the specified line number.

    Returns formatted output with line numbers, highlighting the target line.
    """
    path = Path(file_path)
    if not path.is_file():
        return f"File not found: {file_path}"

    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except Exception as e:
        return f"Error reading {file_path}: {e}"

    lines = text.splitlines()
    total = len(lines)

    if line < 1 or line > total:
        return f"Line {line} is out of range (file has {total} lines)"

    start = max(0, line - 1 - context_lines)
    end = min(total, line + context_lines)

    output = [f"--- {path.name} (line {line}) ---"]
    for i in range(start, end):
        line_num = i + 1
        marker = " >> " if line_num == line else "    "
        output.append(f"{marker}{line_num:4d} | {lines[i]}")
    output.append(f"--- end ({total} lines total) ---")

    return "\n".join(output)


def resolve_error_file(
    filename: str,
    project_dir: str,
    search_paths: list[str] | None = None,
) -> str | None:
    """Resolve a filename from a compiler error to an absolute path.

    Searches the project directory and any configured search paths.
    """
    # Try as-is (absolute path)
    if os.path.isfile(filename):
        return filename

    project_path = Path(project_dir)

    # Try relative to project dir
    candidate = project_path / filename
    if candidate.is_file():
        return str(candidate)

    # Try in search paths
    if search_paths:
        for sp in search_paths:
            sp_path = Path(sp) if os.path.isabs(sp) else project_path / sp
            candidate = sp_path / filename
            if candidate.is_file():
                return str(candidate)

    # Recursive search as last resort
    for f in project_path.rglob(filename):
        if f.is_file():
            return str(f)

    return None
