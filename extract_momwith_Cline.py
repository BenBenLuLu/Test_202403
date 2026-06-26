"""
MoM (Minutes of Meeting) Raw Extraction Script
===============================================
Phase 1 ONLY: Reads all .docx files from MOM_FOLDER and exports a flat
raw table to  mon_rawtable.xlsx.

Output columns:
  Source_File | Seq | Type | Style | Content | Has_Image

No AI / Ollama required.  The exported Excel is intended for
downstream analysis (e.g. Microsoft Copilot).

Usage (PowerShell on Windows):
  & "C:\\Users\\ben.lu\\OneDrive - shl-group.com\\Documents\\Privacy\\python2\\Ben\\AI agent\\.venv\\Scripts\\python.exe" ^
    "C:\\Users\\ben.lu\\OneDrive - shl-group.com\\Documents\\Privacy\\python2\\Ben\\SHL\\Tool Assessment\\MoM extraction\\extract_momwith_Cline.py"
"""

# =============================================================================
# Last commit : 0995853  (2026-06-26)
# Branch      : cursor/phase1-rawonly-66b9
# Note        : Phase 1 only — no Ollama / AI dependency
# =============================================================================

import sys
import io
import subprocess
import json
import re
from pathlib import Path
from datetime import datetime

# Make stdout safe on cp950 / cp936 terminals (replaces unencodable chars with '?')
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(errors="replace")
else:
    sys.stdout = io.TextIOWrapper(
        sys.stdout.buffer,
        encoding=sys.stdout.encoding or "utf-8",
        errors="replace",
    )

# ── Dependency auto-install ───────────────────────────────────────────────────

def _ensure_package(import_name: str, pip_name: str | None = None) -> None:
    """Install a package via pip if it is not already importable."""
    try:
        __import__(import_name)
    except ImportError:
        pkg = pip_name or import_name
        print(f"[setup] '{pkg}' not found -- installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
        print(f"[setup] '{pkg}' installed.")


_ensure_package("docx",    "python-docx")
_ensure_package("pandas")
_ensure_package("openpyxl")

# ── Imports (after ensuring packages) ────────────────────────────────────────

import docx           # noqa: E402
import pandas as pd   # noqa: E402

# ── Configuration ─────────────────────────────────────────────────────────────

# Path to the folder containing .docx MoM files
MOM_FOLDER = Path(r"c:\Users\ben.lu\OneDrive - shl-group.com\Documents\Privacy\python2\Ben\SHL\Tool Assessment\MoM extraction\MoM")

# Output Excel file — saved inside MOM_FOLDER
OUTPUT_XLSX = MOM_FOLDER / "mon_rawtable.xlsx"

# ── Text normalisation ────────────────────────────────────────────────────────
# Maps Unicode symbols that appear in document content to ASCII equivalents.
# Applied to every extracted text cell so the output Excel is plain-ASCII safe.

_UNICODE_REPLACEMENTS = [
    # Arrows
    ("\u2192", "->"),   # → RIGHT ARROW
    ("\u21D2", "->"),   # => RIGHTWARDS DOUBLE ARROW
    ("\u27A1", "->"),   # -> BLACK RIGHTWARDS ARROW
    ("\u2794", "->"),   # -> HEAVY WIDE-HEADED RIGHTWARDS ARROW
    ("\u279C", "->"),   # -> HEAVY ROUND-TIPPED RIGHTWARDS ARROW
    ("\u2190", "<-"),   # <- LEFT ARROW
    ("\u2194", "<->"),  # <-> LEFT RIGHT ARROW
    # Dashes / bullets
    ("\u2013", "-"),    # en dash
    ("\u2014", "--"),   # em dash
    ("\u2022", "*"),    # bullet
    ("\u25CF", "*"),    # black circle
    # Checkboxes (keep as text label so content is readable)
    ("\u2611", "[x]"),  # checked ballot box
    ("\u2612", "[x]"),  # ballot box with X
    ("\u2610", "[ ]"),  # empty ballot box
    ("\u25A0", "[x]"),  # black square (filled checkbox)
    ("\u25A1", "[ ]"),  # white square (empty checkbox)
    ("\u25AA", "[x]"),  # black small square
    # Misc
    ("\u2026", "..."),  # ellipsis
    ("\u00D7", "x"),    # multiplication sign
    ("\u00B0", "deg"),  # degree sign
]

def _normalize(text: str) -> str:
    """Replace known Unicode symbols with ASCII equivalents."""
    for char, replacement in _UNICODE_REPLACEMENTS:
        if char in text:
            text = text.replace(char, replacement)
    return text


# ── Phase 1: Extract .docx files → Raw DataFrame ─────────────────────────────

def extract_rows_from_docx(filepath: Path) -> list[dict]:
    """
    Walk the XML body of a .docx file in document order.

    Returns one dict per paragraph or table row:
        source_file : str   -- filename only (no path)
        seq         : int   -- 1-based sequence within this file
        type        : str   -- "paragraph" or "table_row"
        style       : str   -- Word paragraph style name (e.g. Heading1, Normal)
                               "table" for table rows
        content     : str   -- extracted text
        has_image   : bool  -- True if the document contains any inline image
    """
    from docx.oxml.ns import qn  # noqa: PLC0415

    doc       = docx.Document(str(filepath))
    has_image = len(doc.inline_shapes) > 0
    rows      = []
    seq       = 1

    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            text = _normalize("".join(run.text for run in child.iter(qn("w:t"))).strip())
            if text:
                pStyle = child.find(".//" + qn("w:pStyle"))
                style  = pStyle.get(qn("w:val")) if pStyle is not None else "Normal"
                rows.append({
                    "source_file": filepath.name,
                    "seq":         seq,
                    "type":        "paragraph",
                    "style":       style,
                    "content":     text,
                    "has_image":   has_image,
                })
                seq += 1

        elif tag == "tbl":
            for row_el in child.iter(qn("w:tr")):
                cells = [
                    _normalize("".join(t.text for t in cell_el.iter(qn("w:t"))).strip())
                    for cell_el in row_el.iter(qn("w:tc"))
                ]
                line = " | ".join(cells)
                if line.strip(" |"):
                    rows.append({
                        "source_file": filepath.name,
                        "seq":         seq,
                        "type":        "table_row",
                        "style":       "table",
                        "content":     line,
                        "has_image":   has_image,
                    })
                    seq += 1

    return rows


def build_raw_dataframe(mom_folder: Path) -> pd.DataFrame:
    """
    Iterate all .docx files in *mom_folder* and return a combined DataFrame:

        Source_File | Seq | Type | Style | Content | Has_Image
    """
    if not mom_folder.exists():
        raise FileNotFoundError(
            f"MoM folder not found:\n  {mom_folder}\n"
            "Check the MOM_FOLDER path in the Configuration section."
        )

    files = sorted(mom_folder.glob("*.docx"))
    if not files:
        raise ValueError(f"No .docx files found in:\n  {mom_folder}")

    all_rows: list[dict] = []
    for fp in files:
        print(f"  Extracting: {fp.name}")
        try:
            rows = extract_rows_from_docx(fp)
            all_rows.extend(rows)
            has_img = rows[0]["has_image"] if rows else False
            print(f"    -> {len(rows)} rows  |  has_image: {has_img}")
        except Exception as exc:  # noqa: BLE001
            print(f"    [WARNING] Could not read '{fp.name}': {exc}")

    df = pd.DataFrame(all_rows)
    df.columns = [c.replace("_", "_") for c in df.columns]   # keep snake_case
    # Rename to display-friendly Title_Case column names
    df = df.rename(columns={
        "source_file": "Source_File",
        "seq":         "Seq",
        "type":        "Type",
        "style":       "Style",
        "content":     "Content",
        "has_image":   "Has_Image",
    })
    return df[["Source_File", "Seq", "Type", "Style", "Content", "Has_Image"]]


# ── Export to Excel ───────────────────────────────────────────────────────────

def save_excel(df: pd.DataFrame, output_path: Path) -> None:
    """
    Write the DataFrame to *output_path* with:
      - Blue header row (bold white text)
      - Frozen header row
      - Auto-fit column widths (max 80 chars)
      - Wrap-text on Content column
    """
    from openpyxl.styles import Font, PatternFill, Alignment  # noqa: PLC0415
    from openpyxl.utils import get_column_letter               # noqa: PLC0415

    with pd.ExcelWriter(str(output_path), engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Raw Table")
        ws = writer.sheets["Raw Table"]

        # Freeze header row
        ws.freeze_panes = "A2"

        # Style header row
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4",
                                  fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        header_align = Alignment(horizontal="center", vertical="center",
                                 wrap_text=True)
        for cell in ws[1]:
            cell.fill      = header_fill
            cell.font      = header_font
            cell.alignment = header_align

        # Column widths + wrap Content
        content_col = list(df.columns).index("Content") + 1
        for ci, col_name in enumerate(df.columns, start=1):
            max_len = max(
                len(str(col_name)),
                df[col_name].astype(str).map(len).max() if len(df) else 0,
            )
            ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 4, 80)
            if ci == content_col:
                for ri in range(2, len(df) + 2):
                    ws.cell(row=ri, column=ci).alignment = Alignment(
                        wrap_text=True, vertical="top"
                    )


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    print("=" * 65)
    print("  MoM Raw Extraction  (Phase 1 only)")
    print("=" * 65)
    print(f"  MoM folder : {MOM_FOLDER}")
    print(f"  Output     : {OUTPUT_XLSX}")
    print()

    print("[Phase 1] Extracting .docx files...")
    df = build_raw_dataframe(MOM_FOLDER)

    n_files = df["Source_File"].nunique()
    n_rows  = len(df)
    styles  = sorted(df["Style"].unique())
    print(f"\n  [OK] {n_files} file(s) -> {n_rows} rows total")
    print(f"  Styles found : {styles}")
    print(f"  Has Image    : {df.groupby('Source_File')['Has_Image'].first().to_dict()}")

    print(f"\n[Saving] {OUTPUT_XLSX} ...")
    save_excel(df, OUTPUT_XLSX)
    print(f"  [OK] Saved -- {n_rows} rows, {len(df.columns)} columns")
    print(f"       Columns: {list(df.columns)}")

    print("\n--- Preview (first 20 rows) ---")
    preview = df.copy()
    preview["Content"] = preview["Content"].str[:80].str.replace("\n", " ", regex=False)
    print(preview.head(20).to_string(index=False))

    print("\n[Done] Open mon_rawtable.xlsx and feed it to Copilot.")


if __name__ == "__main__":
    main()
