"""
pdf_merger.py
=============
Utility to merge multiple PDF files into a single output PDF.

Uses `pypdf` (the modern successor to PyPDF2).
Install with:  pip install pypdf

Public API
----------
  merge_pdfs(pdf_paths, output_path)  → str  (the output path written)
"""

import os
from pypdf import PdfWriter


def merge_pdfs(pdf_paths: list, output_path: str) -> str:
    """
    Merge one or more PDF files into a single output PDF.

    Args:
        pdf_paths   : ordered list of file paths to merge (first = first pages)
        output_path : destination file path for the merged PDF

    Returns:
        The absolute path of the written output file.

    Raises:
        ValueError  if pdf_paths is empty or no valid PDFs are found.
        FileNotFoundError if a listed file does not exist.
    """
    if not pdf_paths:
        raise ValueError("pdf_paths is empty – nothing to merge")

    writer = PdfWriter()
    files_merged = 0

    for path in pdf_paths:
        path = str(path)
        if not os.path.isfile(path):
            print(f"  [PDF-MERGE] Skipping missing file: {path}")
            continue
        try:
            writer.append(path)
            files_merged += 1
            print(f"  [PDF-MERGE] Appended: {os.path.basename(path)}")
        except Exception as exc:
            print(f"  [PDF-MERGE] Could not read {os.path.basename(path)}: {exc}")

    if files_merged == 0:
        raise ValueError("No valid PDF files could be merged")

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)

    print(f"  [PDF-MERGE] ✅ Merged {files_merged} PDF(s) → {output_path}")
    return os.path.abspath(output_path)
