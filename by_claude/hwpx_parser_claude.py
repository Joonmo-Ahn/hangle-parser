"""
Utility functions to extract text content from an HWPX document.

HWPX files are zipped XML packages. We shell out to the system `unzip`
via subprocess (per requirement) and then parse the XML sections to
collect paragraph text.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Iterable, List
import xml.etree.ElementTree as ET

# Namespaces used inside HWPX files.
NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hc": "http://www.hancom.co.kr/hwpml/2011/char",
}


def _require_unzip() -> None:
    """Ensure the system `unzip` command exists."""
    if shutil.which("unzip") is None:
        raise RuntimeError("unzip command is required but not found in PATH.")


def _extract_hwpx(hwpx_path: Path, dest_dir: Path) -> None:
    """Extract the HWPX archive to dest_dir using subprocess + unzip."""
    _require_unzip()
    dest_dir.mkdir(parents=True, exist_ok=True)
    cmd = ["unzip", "-qq", str(hwpx_path), "-d", str(dest_dir)]
    subprocess.run(cmd, check=True)


def _iter_section_files(extracted_root: Path) -> Iterable[Path]:
    contents_dir = extracted_root / "Contents"
    if not contents_dir.exists():
        raise FileNotFoundError(f"Contents directory not found in {extracted_root}")
    for section_file in sorted(contents_dir.glob("section*.xml")):
        if section_file.is_file():
            yield section_file


def _parse_paragraphs(section_file: Path) -> List[str]:
    paragraphs: List[str] = []
    tree = ET.parse(section_file)
    root = tree.getroot()
    for para in root.findall(".//hp:p", namespaces=NS):
        runs = [node.text or "" for node in para.findall(".//hc:t", namespaces=NS)]
        if runs:
            paragraphs.append("".join(runs))
    return paragraphs


def extract_text(hwpx_path: Path) -> List[str]:
    """
    Extract all paragraph text from a .hwpx file.

    Returns a list of paragraph strings in document order.
    """
    hwpx_path = Path(hwpx_path).expanduser().resolve()
    if not hwpx_path.exists():
        raise FileNotFoundError(f"HWPX file not found: {hwpx_path}")

    with tempfile.TemporaryDirectory() as tmpdir:
        extracted_root = Path(tmpdir) / "extracted"
        _extract_hwpx(hwpx_path, extracted_root)

        paragraphs: List[str] = []
        for section_file in _iter_section_files(extracted_root):
            paragraphs.extend(_parse_paragraphs(section_file))

    return paragraphs


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract text from an HWPX file.")
    parser.add_argument("hwpx_path", type=Path, help="Path to the .hwpx file")
    args = parser.parse_args()

    paragraphs = extract_text(args.hwpx_path)
    for para in paragraphs:
        print(para)


if __name__ == "__main__":
    main()
