"""
MoM (Minutes of Meeting) Data Extraction and Analysis Script
=============================================================
Extracts text from .docx files, analyzes them using an LLM (OpenAI),
and consolidates results into a structured Markdown summary table.

Pipeline:
  Phase 1  – python-docx extraction of paragraphs + tables
  Phase 2  – AI analysis (root cause, mold component, action type, solution)
  Phase 3  – Consolidated Markdown table output (console + .md file)

Usage (PowerShell on Windows):
  & "C:\\Users\\ben.lu\\OneDrive - shl-group.com\\Documents\\Privacy\\python2\\Ben\\AI agent\\.venv\\Scripts\\python.exe" ^
    "C:\\Users\\ben.lu\\OneDrive - shl-group.com\\Documents\\Privacy\\python2\\Ben\\SHL\\Tool Assessment\\MoM extraction\\extract_momwith_Cline.py"
"""

import os
import sys
import subprocess
import json
import re
from pathlib import Path
from datetime import datetime

# ── Dependency auto-install ───────────────────────────────────────────────────

def _ensure_package(import_name: str, pip_name: str | None = None) -> None:
    """Install a package via pip if it is not already importable."""
    try:
        __import__(import_name)
    except ImportError:
        pkg = pip_name or import_name
        print(f"[setup] Package '{pkg}' not found – installing…")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
        print(f"[setup] '{pkg}' installed successfully.")


_ensure_package("docx", "python-docx")
_ensure_package("openai")

# ── Imports (after ensuring packages) ────────────────────────────────────────

import docx  # noqa: E402  (imported after auto-install guard)
from openai import OpenAI  # noqa: E402

# ── Configuration ─────────────────────────────────────────────────────────────

# Path to the folder containing .docx MoM files.
# By default, resolved relative to this script's own directory.
SCRIPT_DIR = Path(__file__).resolve().parent
MOM_FOLDER = SCRIPT_DIR / "MoM"

# OpenAI settings.
# Set the OPENAI_API_KEY environment variable before running, or paste the key
# directly into OPENAI_API_KEY_FALLBACK below (not recommended for shared code).
OPENAI_API_KEY_FALLBACK = ""          # ← optional hard-coded fallback key
OPENAI_MODEL = "gpt-4o"              # change to "gpt-4-turbo" / "gpt-3.5-turbo" etc.
OPENAI_MAX_TOKENS = 1200
OPENAI_TEMPERATURE = 0.2

# Output file for the final Markdown table (saved next to this script).
OUTPUT_MD_FILE = SCRIPT_DIR / f"MoM_Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"

# ── Phase 1: Extract text from .docx files ────────────────────────────────────

def extract_text_from_docx(filepath: Path) -> str:
    """
    Return a single flat string containing all paragraph text and all table
    text (formatted as Markdown rows) from the given .docx file, preserving
    document order.
    """
    doc = docx.Document(str(filepath))
    chunks: list[str] = []

    # We iterate the document body's XML children to respect original order
    # of paragraphs and tables.
    from docx.oxml.ns import qn  # noqa: PLC0415

    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            # Paragraph
            para_text = "".join(run.text for run in child.iter(qn("w:t")))
            if para_text.strip():
                chunks.append(para_text.strip())

        elif tag == "tbl":
            # Table – render each row as pipe-separated cells
            for row_el in child.iter(qn("w:tr")):
                cells = []
                for cell_el in row_el.iter(qn("w:tc")):
                    cell_text = "".join(
                        t.text for t in cell_el.iter(qn("w:t"))
                    )
                    cells.append(cell_text.strip())
                row_line = " | ".join(cells)
                if row_line.strip(" |"):
                    chunks.append(row_line)

    return "\n".join(chunks)


def load_all_moms(mom_folder: Path) -> dict[str, str]:
    """
    Iterate all .docx files in *mom_folder* and return a dict mapping
    filename → extracted plain text.
    """
    if not mom_folder.exists():
        raise FileNotFoundError(
            f"MoM folder not found: {mom_folder}\n"
            "Please ensure the 'MoM' folder exists next to this script and "
            "contains .docx files."
        )

    files = sorted(mom_folder.glob("*.docx"))
    if not files:
        raise ValueError(f"No .docx files found in: {mom_folder}")

    results: dict[str, str] = {}
    for fp in files:
        print(f"  [Phase 1] Extracting: {fp.name}")
        try:
            results[fp.name] = extract_text_from_docx(fp)
        except Exception as exc:  # noqa: BLE001
            print(f"  [Phase 1] WARNING – could not read '{fp.name}': {exc}")

    return results


# ── Phase 2: AI analysis ──────────────────────────────────────────────────────

ANALYSIS_SYSTEM_PROMPT = """
You are an expert manufacturing quality engineer and meeting analyst.
Your task is to analyse a cleaned, plain-text extract of a Minutes of Meeting
(MoM) document related to mold tooling issues and return a structured JSON
object with EXACTLY the following keys:

{
  "subject":        "<Meeting Subject or a concise title>",
  "according_to":   "<Reported by / source name(s) found in the document>",
  "problems":       "<The abnormal issue or symptom observed>",
  "root_cause":     "<The analysed root cause of the issue>",
  "mold_component": "<The mold component name>",
  "part_number":    "<Part number or drawing number of the component, or 'N/A'>",
  "action_type":    "<MUST be one of: 'Replacement' | 'Repair/Welding' | 'Other'>",
  "solution":       "<Full solution implemented, integrating component, part number and action type>",
  "action":         "<Follow-up actions and responsible person(s)>"
}

Rules:
- All values must be concise plain text (no nested JSON, no Markdown inside).
- If a field cannot be determined from the text, use the string 'N/A'.
- Return ONLY the raw JSON object – no explanation, no code fences.
""".strip()

ANALYSIS_USER_TEMPLATE = """
Below is the full extracted text of the MoM document named "{filename}".
Analyse it and respond with the JSON object as instructed.

---
{content}
---
""".strip()


def analyse_mom_with_ai(
    client: OpenAI,
    filename: str,
    content: str,
) -> dict[str, str]:
    """
    Send the extracted MoM text to the OpenAI API and return parsed JSON.
    Falls back to error placeholders if the API call or parsing fails.
    """
    user_msg = ANALYSIS_USER_TEMPLATE.format(filename=filename, content=content)

    try:
        response = client.chat.completions.create(
            model=OPENAI_MODEL,
            messages=[
                {"role": "system", "content": ANALYSIS_SYSTEM_PROMPT},
                {"role": "user", "content": user_msg},
            ],
            max_tokens=OPENAI_MAX_TOKENS,
            temperature=OPENAI_TEMPERATURE,
            response_format={"type": "json_object"},
        )
        raw = response.choices[0].message.content or "{}"
        data = json.loads(raw)
    except Exception as exc:  # noqa: BLE001
        print(f"  [Phase 2] ERROR analysing '{filename}': {exc}")
        data = {}

    # Ensure all expected keys exist
    defaults = {
        "subject": filename,
        "according_to": "N/A",
        "problems": "N/A",
        "root_cause": "N/A",
        "mold_component": "N/A",
        "part_number": "N/A",
        "action_type": "N/A",
        "solution": "N/A",
        "action": "N/A",
    }
    for key, default in defaults.items():
        data.setdefault(key, default)

    # Build the combined Solution cell
    component = data.get("mold_component", "N/A")
    part_no = data.get("part_number", "N/A")
    action_type = data.get("action_type", "N/A")
    base_solution = data.get("solution", "N/A")

    if component != "N/A" or part_no != "N/A":
        data["solution"] = (
            f"{base_solution} "
            f"[Component: {component}, Part No.: {part_no}, Action: {action_type}]"
        )

    data["filename"] = filename
    return data


# ── Phase 3: Markdown table consolidation ─────────────────────────────────────

_MD_SEP = re.compile(r"\|")  # used to escape pipe characters in cells


def _escape_md(text: str) -> str:
    """Escape pipe characters so they don't break the Markdown table."""
    return text.replace("|", "\\|")


def build_markdown_table(records: list[dict[str, str]]) -> str:
    """
    Produce a Markdown table string from a list of analysis dicts.

    Columns: Subject | According to | Problems | Root Cause | Solution | Action
    """
    header = (
        "| Subject | According to | Problems | Root Cause | Solution | Action |"
    )
    divider = "|---------|--------------|----------|------------|----------|--------|"

    rows = []
    for rec in records:
        row = "| {subject} | {according_to} | {problems} | {root_cause} | {solution} | {action} |".format(
            subject=_escape_md(rec.get("subject", "N/A")),
            according_to=_escape_md(rec.get("according_to", "N/A")),
            problems=_escape_md(rec.get("problems", "N/A")),
            root_cause=_escape_md(rec.get("root_cause", "N/A")),
            solution=_escape_md(rec.get("solution", "N/A")),
            action=_escape_md(rec.get("action", "N/A")),
        )
        rows.append(row)

    lines = [header, divider] + rows
    return "\n".join(lines)


# ── Main orchestration ─────────────────────────────────────────────────────────

def main() -> None:
    print("=" * 70)
    print("  MoM Extraction & Analysis Pipeline")
    print("=" * 70)

    # ── Phase 1 ──────────────────────────────────────────────────────────────
    print("\n[Phase 1] Extracting text from .docx files …")
    mom_texts = load_all_moms(MOM_FOLDER)
    print(f"  ✓ Extracted {len(mom_texts)} file(s).")

    # ── Phase 2 ──────────────────────────────────────────────────────────────
    print("\n[Phase 2] Analysing each MoM with AI …")

    api_key = os.environ.get("OPENAI_API_KEY") or OPENAI_API_KEY_FALLBACK
    if not api_key:
        print(
            "\n[ERROR] OpenAI API key not found.\n"
            "  Set the OPENAI_API_KEY environment variable, or paste your key\n"
            "  into the OPENAI_API_KEY_FALLBACK variable at the top of this script.\n"
        )
        sys.exit(1)

    client = OpenAI(api_key=api_key)

    records: list[dict[str, str]] = []
    for filename, content in mom_texts.items():
        print(f"  [Phase 2] Analysing: {filename}")
        record = analyse_mom_with_ai(client, filename, content)
        records.append(record)
        print(f"    → Subject  : {record['subject']}")
        print(f"    → Root cause: {record['root_cause'][:80]}…" if len(record['root_cause']) > 80 else f"    → Root cause: {record['root_cause']}")

    print(f"  ✓ Analysed {len(records)} file(s).")

    # ── Phase 3 ──────────────────────────────────────────────────────────────
    print("\n[Phase 3] Building consolidated Markdown table …")
    md_table = build_markdown_table(records)

    # Save to file
    OUTPUT_MD_FILE.write_text(md_table, encoding="utf-8")
    print(f"  ✓ Saved to: {OUTPUT_MD_FILE}")

    # Print to console
    print("\n" + "=" * 70)
    print("  MASTER SUMMARY TABLE")
    print("=" * 70 + "\n")
    print(md_table)
    print()

    # Also dump raw JSON records for reference
    json_out = SCRIPT_DIR / f"MoM_RawAnalysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    json_out.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  Raw JSON analysis saved to: {json_out}")
    print("\n[Done]")


if __name__ == "__main__":
    main()
