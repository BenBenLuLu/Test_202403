"""
MoM (Minutes of Meeting) Data Extraction and Analysis Script
=============================================================
Pipeline:
  Phase 1  - Extract paragraphs + table rows from every .docx into a
             pandas DataFrame (Source_File | Seq | Type | Content).
  Phase 2  - Feed each file's DataFrame rows to a local Ollama LLM for
             structured analysis (root cause, component, action, solution).
  Phase 3  - Consolidate AI results into a 6-column summary DataFrame
             and export to  mom_table.xlsx  (two sheets: Summary + Raw).

Prerequisites:
  1. Install Ollama  ->  https://ollama.com/download  (Windows installer)
  2. Pull the model once:
       ollama pull llama3.1:8b
     (or whichever OLLAMA_MODEL you set below)
  3. Ollama must be running before executing this script.
     It starts automatically after installation, or run:  ollama serve

Usage (PowerShell on Windows):
  & "C:\\Users\\ben.lu\\OneDrive - shl-group.com\\Documents\\Privacy\\python2\\Ben\\AI agent\\.venv\\Scripts\\python.exe" ^
    "C:\\Users\\ben.lu\\OneDrive - shl-group.com\\Documents\\Privacy\\python2\\Ben\\SHL\\Tool Assessment\\MoM extraction\\extract_momwith_Cline.py"
"""

import sys
import subprocess
import json
import re
from pathlib import Path
from datetime import datetime

# -- Dependency auto-install ---------------------------------------------------

def _ensure_package(import_name: str, pip_name: str | None = None) -> None:
    """Install a package via pip if it is not already importable."""
    try:
        __import__(import_name)
    except ImportError:
        pkg = pip_name or import_name
        print(f"[setup] '{pkg}' not found - installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
        print(f"[setup] '{pkg}' installed successfully.")


_ensure_package("docx",    "python-docx")
_ensure_package("ollama")
_ensure_package("pandas")
_ensure_package("openpyxl")

# -- Imports (after ensuring packages) ----------------------------------------

import docx           # noqa: E402
import ollama         # noqa: E402
import pandas as pd   # noqa: E402

# -- Configuration -------------------------------------------------------------

SCRIPT_DIR = Path(__file__).resolve().parent

# -- Path to the MoM folder ----------------------------------------------------
# Change the line below to your actual folder path, e.g.:
#   MOM_FOLDER = Path(r"C:\Users\ben.lu\OneDrive - shl-group.com\...\MoM")
MOM_FOLDER = SCRIPT_DIR / "MoM"

# -- Output files (both saved in SCRIPT_DIR by default) -----------------------
OUTPUT_XLSX        = SCRIPT_DIR / "mom_table.xlsx"           # final summary
PHASE1_XLSX        = SCRIPT_DIR / "mom_phase1_preview.xlsx"  # raw extraction

# -- Run mode ------------------------------------------------------------------
# Set PHASE1_ONLY = True to stop after Phase 1 (extraction) so you can
# verify mom_phase1_preview.xlsx before running the Ollama AI step.
PHASE1_ONLY = False

# -- Ollama settings -----------------------------------------------------------
#
# Recommended free models (pull once with: ollama pull <model>):
#   llama3.1:8b   - best all-round accuracy           (~4.7 GB)
#   qwen2.5:7b    - excellent JSON output + Chinese   (~4.4 GB)
#   mistral:7b    - fast, solid reasoning             (~4.1 GB)
#   llama3.2:3b   - lightest, quick results           (~2.0 GB)
#
OLLAMA_MODEL       = "llama3.1:8b"
OLLAMA_TEMPERATURE = 0.2
OLLAMA_NUM_PREDICT = 1200

# -- Phase 1: Extract .docx -> DataFrame ---------------------------------------

def extract_rows_from_docx(filepath: Path) -> list[dict]:
    """
    Walk the XML body of a .docx file in document order and return a list
    of dicts, one per paragraph or table row:

        {"source_file": str, "seq": int, "type": "paragraph"|"table_row",
         "content": str}
    """
    from docx.oxml.ns import qn  # noqa: PLC0415

    doc  = docx.Document(str(filepath))
    rows = []
    seq  = 1

    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            text = "".join(run.text for run in child.iter(qn("w:t"))).strip()
            if text:
                rows.append({
                    "source_file": filepath.name,
                    "seq":         seq,
                    "type":        "paragraph",
                    "content":     text,
                })
                seq += 1

        elif tag == "tbl":
            for row_el in child.iter(qn("w:tr")):
                cells = [
                    "".join(t.text for t in cell_el.iter(qn("w:t"))).strip()
                    for cell_el in row_el.iter(qn("w:tc"))
                ]
                line = " | ".join(cells)
                if line.strip(" |"):
                    rows.append({
                        "source_file": filepath.name,
                        "seq":         seq,
                        "type":        "table_row",
                        "content":     line,
                    })
                    seq += 1

    return rows


def build_raw_dataframe(mom_folder: Path) -> pd.DataFrame:
    """
    Load all .docx files in *mom_folder* and return a combined DataFrame:

        Source_File | Seq | Type | Content
    """
    if not mom_folder.exists():
        raise FileNotFoundError(
            f"MoM folder not found: {mom_folder}\n"
            "Please set MOM_FOLDER to the correct path."
        )

    files = sorted(mom_folder.glob("*.docx"))
    if not files:
        raise ValueError(f"No .docx files found in: {mom_folder}")

    all_rows: list[dict] = []
    for fp in files:
        print(f"  [Phase 1] Extracting: {fp.name}")
        try:
            all_rows.extend(extract_rows_from_docx(fp))
        except Exception as exc:  # noqa: BLE001
            print(f"  [Phase 1] WARNING - could not read '{fp.name}': {exc}")

    df = pd.DataFrame(all_rows, columns=["source_file", "seq", "type", "content"])
    df.columns = ["Source_File", "Seq", "Type", "Content"]
    return df


def df_to_text(file_df: pd.DataFrame) -> str:
    """
    Convert a single file's DataFrame rows into a numbered plain-text block
    that the AI can read easily.
    """
    lines = []
    for _, row in file_df.iterrows():
        lines.append(f"[{row['Seq']}] ({row['Type']})  {row['Content']}")
    return "\n".join(lines)


# -- Phase 2: AI analysis via Ollama ------------------------------------------

ANALYSIS_SYSTEM_PROMPT = """
You are an expert manufacturing quality engineer and meeting analyst.
Your task is to analyse a structured table of a Minutes of Meeting (MoM)
document related to mold tooling issues and return a JSON object with
EXACTLY the following keys:

{
  "subject":        "<Meeting Subject or a concise title>",
  "according_to":   "<Reported by / source name(s) found in the document>",
  "problems":       "<The abnormal issue or symptom observed>",
  "root_cause":     "<The analysed root cause of the issue>",
  "mold_component": "<The mold component name>",
  "part_number":    "<Part number or drawing number of the component, or N/A>",
  "action_type":    "<MUST be one of: Replacement | Repair/Welding | Other>",
  "solution":       "<Full solution implemented>",
  "action":         "<Follow-up actions and responsible person(s)>"
}

Rules:
- All values must be concise plain text (no nested JSON, no Markdown).
- If a field cannot be determined from the text, use the string N/A.
- Return ONLY the raw JSON object - no explanation, no code fences.
""".strip()

ANALYSIS_USER_TEMPLATE = """
Below is the extracted content of the MoM document "{filename}",
presented as a numbered table (each row shows its sequence number,
type [paragraph or table_row], and content).

Analyse the content and respond with the JSON object as instructed.

---
{table_text}
---
""".strip()


def _extract_json(raw: str) -> dict:
    """Robustly extract a JSON object from the model's response string."""
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass
    fence = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", raw, re.DOTALL)
    if fence:
        try:
            return json.loads(fence.group(1))
        except json.JSONDecodeError:
            pass
    brace = re.search(r"\{.*\}", raw, re.DOTALL)
    if brace:
        try:
            return json.loads(brace.group(0))
        except json.JSONDecodeError:
            pass
    return {}


def analyse_file_df(filename: str, file_df: pd.DataFrame) -> dict[str, str]:
    """
    Convert a file's DataFrame to text, send to Ollama, return parsed dict.
    """
    table_text = df_to_text(file_df)
    user_msg   = ANALYSIS_USER_TEMPLATE.format(
        filename=filename, table_text=table_text
    )

    try:
        response = ollama.chat(
            model=OLLAMA_MODEL,
            messages=[
                {"role": "system", "content": ANALYSIS_SYSTEM_PROMPT},
                {"role": "user",   "content": user_msg},
            ],
            format="json",
            options={
                "temperature": OLLAMA_TEMPERATURE,
                "num_predict": OLLAMA_NUM_PREDICT,
            },
        )
        data = _extract_json(response["message"]["content"])

    except ollama.ResponseError as exc:
        print(f"\n  [Phase 2] Ollama model error: {exc}")
        print(f"  Run:  ollama pull {OLLAMA_MODEL}\n")
        data = {}
    except Exception as exc:  # noqa: BLE001
        print(f"  [Phase 2] ERROR analysing '{filename}': {exc}")
        data = {}

    defaults = {
        "subject": filename, "according_to": "N/A", "problems": "N/A",
        "root_cause": "N/A", "mold_component": "N/A", "part_number": "N/A",
        "action_type": "N/A", "solution": "N/A", "action": "N/A",
    }
    for k, v in defaults.items():
        data.setdefault(k, v)

    # Merge component details into the Solution cell
    comp   = data.get("mold_component", "N/A")
    part   = data.get("part_number",    "N/A")
    atype  = data.get("action_type",    "N/A")
    sol    = data.get("solution",       "N/A")
    if comp != "N/A" or part != "N/A":
        data["solution"] = (
            f"{sol} [Component: {comp}, Part No.: {part}, Action: {atype}]"
        )

    data["source_file"] = filename
    return data


# -- Phase 3: Build summary DataFrame and export to Excel ---------------------

SUMMARY_COLUMNS = {
    "subject":      "Subject",
    "according_to": "According to",
    "problems":     "Problems",
    "root_cause":   "Root Cause",
    "solution":     "Solution",
    "action":       "Action",
}


def build_summary_dataframe(records: list[dict]) -> pd.DataFrame:
    """Return the 6-column summary DataFrame from a list of analysis dicts."""
    rows = [
        {display: rec.get(key, "N/A") for key, display in SUMMARY_COLUMNS.items()}
        for rec in records
    ]
    return pd.DataFrame(rows, columns=list(SUMMARY_COLUMNS.values()))


def save_to_excel(summary_df: pd.DataFrame, raw_df: pd.DataFrame,
                  output_path: Path) -> None:
    """
    Write two sheets to *output_path*:
      * Summary  - the 6-column AI analysis table
      * Raw      - every extracted paragraph / table row
    """
    with pd.ExcelWriter(str(output_path), engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        raw_df.to_excel(writer,     sheet_name="Raw",     index=False)

        # Auto-fit column widths on both sheets
        for sheet_name, df in [("Summary", summary_df), ("Raw", raw_df)]:
            ws = writer.sheets[sheet_name]
            for col_idx, col_name in enumerate(df.columns, start=1):
                max_len = max(
                    len(str(col_name)),
                    df[col_name].astype(str).map(len).max() if len(df) else 0,
                )
                ws.column_dimensions[
                    ws.cell(row=1, column=col_idx).column_letter
                ].width = min(max_len + 4, 80)


# -- Ollama pre-flight check ---------------------------------------------------

def _check_ollama_running() -> None:
    try:
        ollama.list()
    except Exception:  # noqa: BLE001
        print(
            "\n[ERROR] Cannot connect to Ollama.\n"
            "  * Download: https://ollama.com/download\n"
            "  * Then run in a terminal: ollama serve\n"
            f"  * Pull the model:        ollama pull {OLLAMA_MODEL}\n"
        )
        sys.exit(1)


# -- Main orchestration --------------------------------------------------------

def main() -> None:
    print("=" * 70)
    print("  MoM Extraction & Analysis Pipeline  (Ollama / local LLM)")
    print("=" * 70)
    print(f"  Mode            : {'PHASE 1 ONLY (no AI)' if PHASE1_ONLY else 'Full pipeline (Phase 1 + 2 + 3)'}")
    print(f"  MoM folder      : {MOM_FOLDER}")
    print(f"  Phase 1 preview : {PHASE1_XLSX}")
    print(f"  Final output    : {OUTPUT_XLSX}")
    print()

    # -- Phase 1: Extract -> Raw DataFrame -------------------------------------
    print("[Phase 1] Extracting .docx content into DataFrame ...")
    raw_df = build_raw_dataframe(MOM_FOLDER)
    n_files    = raw_df["Source_File"].nunique()
    total_rows = len(raw_df)
    print(f"  [OK] {n_files} file(s) -> {total_rows} rows extracted.\n")
    print(raw_df.to_string(max_rows=15, max_colwidth=70))
    print()

    # Always save the Phase 1 raw extraction so it can be reviewed
    with pd.ExcelWriter(str(PHASE1_XLSX), engine="openpyxl") as writer:
        raw_df.to_excel(writer, sheet_name="Phase1 Raw", index=False)
        ws = writer.sheets["Phase1 Raw"]
        ws.freeze_panes = "A2"
        for col_idx, col_name in enumerate(raw_df.columns, start=1):
            max_len = max(
                len(str(col_name)),
                raw_df[col_name].astype(str).map(len).max() if total_rows else 0,
            )
            ws.column_dimensions[
                ws.cell(row=1, column=col_idx).column_letter
            ].width = min(max_len + 4, 80)
    print(f"  [OK] Phase 1 preview saved -> {PHASE1_XLSX}")

    if PHASE1_ONLY:
        print("\n[Phase 1 only mode] Stopping here.")
        print("  Open the preview file above, verify the extraction,")
        print("  then set PHASE1_ONLY = False and re-run to continue.\n")
        return

    # -- Phase 2: AI analysis per file ----------------------------------------
    _check_ollama_running()
    print("\n[Phase 2] Analysing each file with Ollama ...")
    records: list[dict] = []

    for filename, file_df in raw_df.groupby("Source_File", sort=False):
        print(f"  -> {filename}  ({len(file_df)} rows fed to AI)")
        record = analyse_file_df(str(filename), file_df)
        records.append(record)
        rc = record["root_cause"]
        print(f"     Subject   : {record['subject']}")
        print(f"     Root Cause: {rc[:90]}{'...' if len(rc) > 90 else ''}")

    print(f"\n  [OK] Analysed {len(records)} file(s).\n")

    # -- Phase 3: Summary DataFrame -> Excel -----------------------------------
    print("[Phase 3] Building summary DataFrame and saving to Excel ...")
    summary_df = build_summary_dataframe(records)

    print("\n" + "=" * 70)
    print("  SUMMARY TABLE")
    print("=" * 70)
    print(summary_df.to_string(index=False, max_colwidth=60))
    print()

    save_to_excel(summary_df, raw_df, OUTPUT_XLSX)
    print(f"  [OK] Saved -> {OUTPUT_XLSX}")
    print(f"     Sheet 'Summary' : {len(summary_df)} row(s), 6 columns")
    print(f"     Sheet 'Raw'     : {len(raw_df)} row(s), 4 columns")
    print("\n[Done]")


if __name__ == "__main__":
    main()
