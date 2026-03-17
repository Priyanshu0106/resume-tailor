"""Rezip modified XML back into a .docx with proper ZIP structure.

Per CLAUDE.md: only word/document.xml is modified; all other files are
copied verbatim from the unpacked temp directory.
"""

import os
import zipfile
from pathlib import Path


def repack_docx(temp_dir: str, output_path: str) -> None:
    """Repack the unpacked .docx directory into a new .docx file.

    - Uses ZIP_DEFLATED compression
    - [Content_Types].xml is written FIRST (OOXML spec requirement)
    - All files from temp_dir are included in the same relative structure
    - Verifies output file size is reasonable
    """
    temp_path = Path(temp_dir)
    all_files = []

    for root, _dirs, files in os.walk(temp_dir):
        for f in files:
            full = os.path.join(root, f)
            arcname = os.path.relpath(full, temp_dir).replace("\\", "/")
            all_files.append((full, arcname))

    # Separate [Content_Types].xml to write first
    content_types = None
    other_files = []
    for full, arcname in all_files:
        if arcname == "[Content_Types].xml":
            content_types = (full, arcname)
        else:
            other_files.append((full, arcname))

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        # [Content_Types].xml MUST be first entry
        if content_types:
            zout.write(content_types[0], content_types[1])

        for full, arcname in other_files:
            zout.write(full, arcname)

    # Basic sanity check on output size
    output_size = os.path.getsize(output_path)
    if output_size < 1000:
        raise RuntimeError(
            f"Output .docx is suspiciously small ({output_size} bytes). "
            "Something went wrong during repacking."
        )


def cleanup_temp(temp_dir: str) -> None:
    """Remove the temporary unpacked directory."""
    import shutil

    if os.path.isdir(temp_dir):
        shutil.rmtree(temp_dir, ignore_errors=True)
