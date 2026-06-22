"""
MoM (Minutes of Meeting) Data Extraction and Analysis Script
=============================================================
Extracts text from .docx files, analyses them with a LOCAL Ollama LLM
(free, open-source, no API key required), and consolidates the results
into a structured Markdown summary table.

Pipeline:
  Phase 1  – python-docx extraction of paragraphs + tables
  Phase 2  – AI analysis via Ollama (root cause, mold component,
              action type, solution)
  Phase 3  – Consolidated Markdown table output (console + .md file)

Prerequisites:
  1. Install Ollama  →  https://ollama.com/download  (Windows installer)
  2. Pull the model once:
       ollama pull llama3.1:8b
     (or whichever OLLAMA_MODEL you set below)
  3. Ollama must be running in the background before you execute this script.
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
_ensure_package("ollama")

# ── Imports (after ensuring packages) ────────────────────────────────────────

import docx    # noqa: E402
import ollama  # noqa: E402

# ── Configuration ─────────────────────────────────────────────────────────────

# Path to the folder containing .docx MoM files.
# Resolved relative to this script's own directory by default.
SCRIPT_DIR = Path(__file__).resolve().parent
MOM_FOLDER  = SCRIPT_DIR / "MoM"

# ── Ollama settings ───────────────────────────────────────────────────────────
#
# Recommended free models (pull once with: ollama pull <model>):
#
#   llama3.1:8b   – Meta Llama 3.1 8B   – best all-round accuracy  (~4.7 GB)
#   qwen2.5:7b    – Alibaba Qwen 2.5 7B – excellent JSON + Chinese  (~4.4 GB)
#   mistral:7b    – Mistral 7B v0.3      – fast, solid reasoning    (~4.1 GB)
#   llama3.2:3b   – Meta Llama 3.2 3B   – lightest, quick results  (~2.0 GB)
#
OLLAMA_MODEL       = "llama3.1:8b"   # ← change to whichever model you have pulled
OLLAMA_TEMPERATURE = 0.2             # lower = more deterministic
OLLAMA_NUM_PREDICT = 1200            # max tokens to generate per document

# Output file for the final Markdown table (saved next to this script).
OUTPUT_MD_FILE = SCRIPT_DIR / f"MoM_Summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"

# ── Phase 1: Extract text from .docx files ────────────────────────────────────

def extract_text_from_docx(filepath: Path) -> str:
    """
    Return a single flat string of all paragraph text and all table text
    (formatted as Markdown pipe-separated rows) from the given .docx file,
    preserving the original document order by walking the raw XML body.
    """
    doc = docx.Document(str(filepath))
    chunks: list[str] = []

    from docx.oxml.ns import qn  # noqa: PLC0415

    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            para_text = "".join(run.text for run in child.iter(qn("w:t")))
            if para_text.strip():
                chunks.append(para_text.strip())

        elif tag == "tbl":
            for row_el in child.iter(qn("w:tr")):
                cells = []
                for cell_el in row_el.iter(qn("w:tc")):
                    cell_text = "".join(t.text for t in cell_el.iter(qn("w:t")))
                    cells.append(cell_text.strip())
                row_line = " | ".join(cells)
                if row_line.strip(" |"):
                    chunks.append(row_line)

    return "\n".join(chunks)


def load_all_moms(mom_folder: Path) -> dict[str, str]:
    """
    Walk *mom_folder* for .docx files and return {filename: extracted_text}.
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


# ── Phase 2: AI analysis via Ollama ──────────────────────────────────────────

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
  "part_number":    "<Part number or drawing number of the component, or N/A>",
  "action_type":    "<MUST be one of: Replacement | Repair/Welding | Other>",
  "solution":       "<Full solution implemented, integrating component, part number and action type>",
  "action":         "<Follow-up actions and responsible person(s)>"
}

Rules:
- All values must be concise plain text (no nested JSON, no Markdown inside).
- If a field cannot be determined from the text, use the string N/A.
- Return ONLY the raw JSON object – no explanation, no code fences, no extra text.
""".strip()

ANALYSIS_USER_TEMPLATE = """
Below is the full extracted text of the MoM document named "{filename}".
Analyse it and respond with the JSON object as instructed.

---
{content}
---
""".strip()


def _extract_json(raw: str) -> dict:
    """
    Robustly pull a JSON object out of the model's raw response string.
    Handles cases where the model wraps the JSON in markdown code fences
    or adds surrounding explanation text.
    """
    # Try direct parse first
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # Strip markdown code fences if present
    fence_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", raw, re.DOTALL)
    if fence_match:
        try:
            return json.loads(fence_match.group(1))
        except json.JSONDecodeError:
            pass

    # Extract the outermost { ... } block
    brace_match = re.search(r"\{.*\}", raw, re.DOTALL)
    if brace_match:
        try:
            return json.loads(brace_match.group(0))
        except json.JSONDecodeError:
            pass

    return {}


def analyse_mom_with_ollama(filename: str, content: str) -> dict[str, str]:
    """
    Send the extracted MoM text to the local Ollama model and return
    a parsed dict.  Falls back to N/A placeholders on any error.
    """
    user_msg = ANALYSIS_USER_TEMPLATE.format(filename=filename, content=content)

    try:
        response = ollama.chat(
            model=OLLAMA_MODEL,
            messages=[
                {"role": "system", "content": ANALYSIS_SYSTEM_PROMPT},
                {"role": "user",   "content": user_msg},
            ],
            format="json",          # instructs Ollama to constrain output to JSON
            options={
                "temperature": OLLAMA_TEMPERATURE,
                "num_predict": OLLAMA_NUM_PREDICT,
            },
        )
        raw  = response["message"]["content"]
        data = _extract_json(raw)

    except ollama.ResponseError as exc:
        # Model not found locally
        print(f"\n  [Phase 2] Ollama model error: {exc}")
        print(
            f"  Please pull the model first by running in a terminal:\n"
            f"    ollama pull {OLLAMA_MODEL}\n"
        )
        data = {}
    except Exception as exc:  # noqa: BLE001
        print(f"  [Phase 2] ERROR analysing '{filename}': {exc}")
        data = {}

    # Fill in any missing keys with defaults
    defaults = {
        "subject":        filename,
        "according_to":   "N/A",
        "problems":       "N/A",
        "root_cause":     "N/A",
        "mold_component": "N/A",
        "part_number":    "N/A",
        "action_type":    "N/A",
        "solution":       "N/A",
        "action":         "N/A",
    }
    for key, default in defaults.items():
        data.setdefault(key, default)

    # Build the enriched Solution cell (component + part no. + action type)
    component    = data.get("mold_component", "N/A")
    part_no      = data.get("part_number",    "N/A")
    action_type  = data.get("action_type",    "N/A")
    base_solution = data.get("solution",      "N/A")

    if component != "N/A" or part_no != "N/A":
        data["solution"] = (
            f"{base_solution} "
            f"[Component: {component}, Part No.: {part_no}, Action: {action_type}]"
        )

    data["filename"] = filename
    return data


# ── Phase 3: Markdown table consolidation ─────────────────────────────────────

def _escape_md(text: str) -> str:
    """Escape pipe characters so they don't break the Markdown table."""
    return str(text).replace("|", "\\|")


def build_markdown_table(records: list[dict[str, str]]) -> str:
    """
    Produce a 6-column Markdown table from the list of analysis dicts.
    Columns: Subject | According to | Problems | Root Cause | Solution | Action
    """
    header  = "| Subject | According to | Problems | Root Cause | Solution | Action |"
    divider = "|---------|--------------|----------|------------|----------|--------|"

    rows = []
    for rec in records:
        row = (
            "| {subject} | {according_to} | {problems} | "
            "{root_cause} | {solution} | {action} |"
        ).format(
            subject      = _escape_md(rec.get("subject",      "N/A")),
            according_to = _escape_md(rec.get("according_to", "N/A")),
            problems     = _escape_md(rec.get("problems",     "N/A")),
            root_cause   = _escape_md(rec.get("root_cause",   "N/A")),
            solution     = _escape_md(rec.get("solution",     "N/A")),
            action       = _escape_md(rec.get("action",       "N/A")),
        )
        rows.append(row)

    return "\n".join([header, divider] + rows)


# ── Main orchestration ─────────────────────────────────────────────────────────

def _check_ollama_running() -> None:
    """Warn the user early if Ollama is not reachable."""
    try:
        ollama.list()   # lightweight probe – lists locally available models
    except Exception:  # noqa: BLE001
        print(
            "\n[ERROR] Cannot connect to Ollama.\n"
            "  Make sure Ollama is installed and running:\n"
            "    • Download: https://ollama.com/download\n"
            "    • Then run in a separate terminal: ollama serve\n"
            "  After Ollama is running, pull the model if you haven't yet:\n"
            f"    ollama pull {OLLAMA_MODEL}\n"
        )
        sys.exit(1)


def main() -> None:
    print("=" * 70)
    print("  MoM Extraction & Analysis Pipeline  (Ollama / local LLM)")
    print("=" * 70)
    print(f"  Model : {OLLAMA_MODEL}")
    print(f"  MoM folder : {MOM_FOLDER}")
    print()

    # Pre-flight check
    _check_ollama_running()

    # ── Phase 1 ──────────────────────────────────────────────────────────────
    print("[Phase 1] Extracting text from .docx files …")
    mom_texts = load_all_moms(MOM_FOLDER)
    print(f"  ✓ Extracted {len(mom_texts)} file(s).\n")

    # ── Phase 2 ──────────────────────────────────────────────────────────────
    print("[Phase 2] Analysing each MoM with local Ollama LLM …")
    records: list[dict[str, str]] = []

    for filename, content in mom_texts.items():
        print(f"  → {filename}")
        record = analyse_mom_with_ollama(filename, content)
        records.append(record)
        rc = record["root_cause"]
        print(f"     Subject   : {record['subject']}")
        print(f"     Root Cause: {rc[:90]}{'…' if len(rc) > 90 else ''}")

    print(f"\n  ✓ Analysed {len(records)} file(s).\n")

    # ── Phase 3 ──────────────────────────────────────────────────────────────
    print("[Phase 3] Building consolidated Markdown table …")
    md_table = build_markdown_table(records)

    OUTPUT_MD_FILE.write_text(md_table, encoding="utf-8")
    print(f"  ✓ Markdown table saved to: {OUTPUT_MD_FILE}")

    print("\n" + "=" * 70)
    print("  MASTER SUMMARY TABLE")
    print("=" * 70 + "\n")
    print(md_table)
    print()

    json_out = SCRIPT_DIR / f"MoM_RawAnalysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    json_out.write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  Raw JSON saved to: {json_out}")
    print("\n[Done]")


if __name__ == "__main__":
    main()
