#!/usr/bin/env python3
"""Convert DSM_NIS_WORKFLOW_ARCHITECTURE.md to PDF using fpdf2."""

import os
from fpdf import FPDF

MD_FILE = "DSM_NIS_WORKFLOW_ARCHITECTURE.md"
PDF_FILE = "DSM_NIS_WORKFLOW_ARCHITECTURE.pdf"


class PDF(FPDF):
    def header(self):
        self.set_font("Courier", "B", 8)
        self.set_text_color(100, 100, 100)
        self.cell(0, 5, "DSM-NIS Booking Automation - Workflow Architecture", align="C")
        self.ln(6)

    def footer(self):
        self.set_y(-10)
        self.set_font("Courier", "I", 7)
        self.set_text_color(128, 128, 128)
        self.cell(0, 5, f"Page {self.page_no()}/{{nb}}", align="C")


def _sanitize(text):
    """Replace Unicode characters that latin-1 can't encode."""
    replacements = {
        '\u2014': '--', '\u2013': '-', '\u2018': "'", '\u2019': "'",
        '\u201c': '"', '\u201d': '"', '\u2026': '...', '\u2192': '->',
        '\u2190': '<-', '\u2022': '*', '\u25cf': '*', '\u2713': '[x]',
        '\u2714': '[x]', '\u2716': '[X]', '\u2717': '[X]', '\u2610': '[ ]',
        '\u2611': '[x]', '\u25b6': '>', '\u23f3': '[wait]', '\u2709': '[mail]',
        '\u2705': '[OK]', '\u274c': '[FAIL]', '\u26a0': '[!]',
        '\u2699': '[*]', '\u231b': '[time]', '\u2328': '[kb]',
        '\u2615': '', '\u2764': '', '\u1f4cb': '', '\u1f4c4': '',
        '\u27a1': '->', '\u2b06': '^', '\u2b07': 'v',
    }
    for char, repl in replacements.items():
        text = text.replace(char, repl)
    # Remove any remaining non-latin1 characters
    return text.encode('latin-1', errors='replace').decode('latin-1')


def main():
    with open(MD_FILE, "r", encoding="utf-8") as f:
        lines = f.readlines()

    pdf = PDF(orientation="P", unit="mm", format="A4")
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=12)
    pdf.set_left_margin(8)
    pdf.set_right_margin(8)
    pdf.add_page()

    in_code_block = False

    for line in lines:
        raw = line.rstrip("\n")

        # Toggle code blocks
        if raw.startswith("```"):
            in_code_block = not in_code_block
            continue

        if in_code_block:
            pdf.set_font("Courier", "", 7)
            pdf.set_text_color(30, 30, 30)
            text = _sanitize(raw[:120] if len(raw) > 120 else raw)
            pdf.cell(0, 3.2, text, new_x="LMARGIN", new_y="NEXT")
            continue

        # Headings
        if raw.startswith("# "):
            pdf.set_font("Courier", "B", 14)
            pdf.set_text_color(0, 0, 0)
            pdf.ln(4)
            pdf.cell(0, 7, _sanitize(raw[2:]), new_x="LMARGIN", new_y="NEXT")
            pdf.ln(2)
        elif raw.startswith("## "):
            # Start a new page for each major section so content stays together
            pdf.add_page()
            pdf.set_font("Courier", "B", 11)
            pdf.set_text_color(20, 20, 80)
            pdf.cell(0, 6, _sanitize(raw[3:]), new_x="LMARGIN", new_y="NEXT")
            pdf.line(8, pdf.get_y(), 202, pdf.get_y())
            pdf.ln(2)
        elif raw.startswith("### "):
            pdf.set_font("Courier", "B", 9)
            pdf.set_text_color(40, 40, 40)
            pdf.ln(2)
            pdf.cell(0, 5, _sanitize(raw[4:]), new_x="LMARGIN", new_y="NEXT")
            pdf.ln(1)
        elif raw.startswith("---"):
            pdf.line(10, pdf.get_y(), 200, pdf.get_y())
            pdf.ln(3)
        elif raw.startswith("|"):
            # Table row
            pdf.set_font("Courier", "", 6)
            pdf.set_text_color(0, 0, 0)
            cells = [c.strip() for c in raw.split("|")[1:-1]]
            if all(set(c) <= set("- :") for c in cells):
                continue  # skip separator row
            col_w = min(194 / max(len(cells), 1), 60)
            for cell in cells:
                cell_text = _sanitize(cell.replace("**", "").replace("`", "")[:40])
                pdf.cell(col_w, 3.5, cell_text, border=1)
            pdf.ln()
        elif raw.strip() == "":
            pdf.ln(2)
        else:
            # Normal text
            pdf.set_font("Courier", "", 8)
            pdf.set_text_color(0, 0, 0)
            text = _sanitize(raw.replace("**", "").replace("`", "").replace("*", ""))
            try:
                pdf.multi_cell(0, 3.5, text)
            except Exception:
                # If multi_cell fails, use cell with truncation
                pdf.cell(0, 3.5, text[:130], new_x="LMARGIN", new_y="NEXT")

    pdf.output(PDF_FILE)
    size = os.path.getsize(PDF_FILE)
    print(f"PDF created: {PDF_FILE} ({size:,} bytes)")


if __name__ == "__main__":
    main()
