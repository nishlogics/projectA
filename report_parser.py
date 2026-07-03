"""
report_parser.py
================
Auto-extract everything the scale-up engine needs directly from a PR&D report PDF,
so the user only has to upload the report (plus the one plant decision: scale-up ratio).

Extracts, per stage:
  * numbered process steps (+ notes)      -> flowchart
  * the "List of Raw materials" table      -> lab quantities, unit, MW, mole ratio, remarks
  * molecular weights from the scheme      -> input MW / product MW
  * theoretical / reported yield           -> yield fraction

Works on text-based PDFs (pdfplumber tables + text). Falls back to line regex when
the table grid isn't detected. Multi-stage reports are split on the stage headings
(CGTR1 / CGTR2 / CGTR3 draft process, or "Stage n").
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

import pdfplumber

from pfd_pipeline import extract_steps_from_lines, Step
from pfd_scaleup import RawMaterial, StageCalc


# ---------------------------------------------------------------------------
@dataclass
class ParsedStage:
    name: str
    raw_text: str
    steps: list[Step] = field(default_factory=list)
    materials: list[RawMaterial] = field(default_factory=list)
    input_mw: Optional[float] = None
    product_mw: Optional[float] = None
    yield_frac: Optional[float] = None
    lab_input_g: Optional[float] = None


_STAGE_SPLIT = re.compile(
    r"(?im)^\s*(CGTR\s*\d|CTGR\s*\d|Stage[\s\-]*\d|.*Draft\s+process)\b")
_UNIT_SOLID = {"g", "gm", "gms", "kg", "grams"}
_UNIT_LIQ = {"ml", "mL", "l", "L", "litre", "liters"}


_TABLE_ROW_RE = re.compile(
    r"^\s*\d{1,3}\s+.+?\s+\d+(\.\d+)?\s+(g|gm|gms|ml|mL|kg|L)\b", re.IGNORECASE)


def _looks_like_table_row(line: str) -> bool:
    """True if a line is really a raw-material table row (e.g. '1 L-Tyrosine 100 g 181.19 ...')
    rather than a genuine numbered process step (e.g. '1. Charge water into the RBF')."""
    if _TABLE_ROW_RE.match(line):
        return True
    nums = re.findall(r"\d+\.\d+", line)
    return len(nums) >= 2 and not re.search(
        r"(?i)\b(charge|add|stir|cool|heat|adjust|maintain|filter|wash|dry|separate|"
        r"submit|raise|distill|unload|suck|release|stop|allow)\b", line)


def _to_float(x) -> Optional[float]:
    if x is None:
        return None
    s = re.sub(r"[^\d.\-]", "", str(x))
    if not s or s in {".", "-"}:
        return None
    try:
        return float(s)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
def parse_report(file) -> list[ParsedStage]:
    """Main entry: PDF file-like -> list of ParsedStage (one per synthesis stage)."""
    with pdfplumber.open(file) as pdf:
        pages_text = [p.extract_text() or "" for p in pdf.pages]
        pages_tables = [p.extract_tables() or [] for p in pdf.pages]

    full_text = "\n".join(pages_text)
    stages = _split_stages(full_text)

    # attach tables to whichever stage's text region they fall in
    all_tables = [t for tbls in pages_tables for t in tbls]
    for st in stages:
        st.steps = extract_steps_from_lines(_procedure_lines(st.raw_text))
        st.input_mw, st.product_mw = _extract_mws(st.raw_text)
        st.yield_frac = _extract_yield(st.raw_text)
        st.lab_input_g = _extract_lab_input(st.raw_text)

    # material tables: match a table to a stage by header proximity
    _assign_tables(stages, all_tables, pages_text, pages_tables)
    return stages


# ---------------------------------------------------------------------------
def _procedure_lines(text: str) -> list[str]:
    """Return only the lines belonging to the process procedure, dropping the
    raw-materials table and scheme so they don't leak into the step list.

    Starts at the first 'Process description' / 'Reaction procedure' heading (or
    the first genuine numbered charge/stir/cool step) and ends before the yield.
    """
    lines = text.splitlines()
    start = 0
    for i, ln in enumerate(lines):
        if re.search(r"(?i)\b(process description|reaction procedure)\b", ln):
            start = i + 1
            break
    else:
        # no explicit heading: begin at the first action-verb numbered step
        for i, ln in enumerate(lines):
            if re.match(r"(?i)^\s*\d{1,3}[.\)]\s*(charge|stir|cool|add|heat|"
                        r"maintain|adjust|filter|distill|dry|slowly)", ln):
                start = i
                break

    end = len(lines)
    for i in range(start, len(lines)):
        if re.search(r"(?i)\b(theoretical yield|reported yield)\b", lines[i]):
            end = i
            break
    return [ln for ln in lines[start:end] if not _looks_like_table_row(ln)]


def _split_stages(full_text: str) -> list[ParsedStage]:
    lines = full_text.splitlines()
    # find explicit stage sections like "CGTR2 Draft process" / "Synthetic scheme of CGTR3"
    idxs = []
    for i, ln in enumerate(lines):
        m = re.search(r"(?i)\b(CGTR|CTGR)\s*([123])\b.*(draft process|process description)", ln)
        if m:
            idxs.append((i, f"Stage{m.group(2)}"))
    if not idxs:
        # single stage
        return [ParsedStage(name="Stage1", raw_text=full_text)]

    stages = []
    for k, (start, name) in enumerate(idxs):
        end = idxs[k + 1][0] if k + 1 < len(idxs) else len(lines)
        stages.append(ParsedStage(name=name, raw_text="\n".join(lines[start:end])))
    # dedupe by name, keeping the richest block
    best: dict[str, ParsedStage] = {}
    for s in stages:
        if s.name not in best or len(s.raw_text) > len(best[s.name].raw_text):
            best[s.name] = s
    return list(best.values())


def _extract_mws(text: str) -> tuple[Optional[float], Optional[float]]:
    mws = [float(m) for m in re.findall(r"(?i)M\.?\s*Wt:?\s*([\d.]+)", text)]
    if not mws:
        return None, None
    return mws[0], mws[-1]  # first = SM/input, last = product


def _extract_yield(text: str) -> Optional[float]:
    m = re.search(r"(?i)reported\s+yield:?\s*([\d.]+)\s*%", text)
    if m:
        return float(m.group(1)) / 100.0
    # theoretical: "142.20 g for 100 g of input" -> not a fraction; skip
    return None


def _extract_lab_input(text: str) -> Optional[float]:
    m = re.search(r"(?i)for\s+([\d.]+)\s*g\s+of\s+input", text)
    return float(m.group(1)) if m else None


# ---------------------------------------------------------------------------
def _assign_tables(stages, all_tables, pages_text, pages_tables):
    """Assign each extracted table to the nearest stage by matching stage keyword
    presence on the same page; parse rows into RawMaterial objects."""
    # Build page->stagename hint
    for pi, tbls in enumerate(pages_tables):
        if not tbls:
            continue
        ptxt = pages_text[pi]
        m = re.search(r"(?i)\b(CGTR|CTGR)\s*([123])\b", ptxt)
        stage_name = f"Stage{m.group(2)}" if m else (stages[0].name if stages else "Stage1")
        target = next((s for s in stages if s.name == stage_name), stages[0] if stages else None)
        if target is None:
            continue
        for tbl in tbls:
            mats = _parse_material_table(tbl)
            if mats:
                target.materials.extend(mats)


def _parse_material_table(tbl: list[list]) -> list[RawMaterial]:
    """Turn a raw pdfplumber table into RawMaterial rows.
    Expected columns (order tolerant): S.No | Raw Material | Qty | Unit | M.Wt | Mol | Mole Ratio | Remarks
    """
    if not tbl or len(tbl) < 2:
        return []

    # locate header row
    header = None
    for r in tbl[:3]:
        joined = " ".join((c or "").lower() for c in r)
        if "raw material" in joined or ("qty" in joined and ("unit" in joined or "uom" in joined)):
            header = [(c or "").strip().lower() for c in r]
            break
    if header is None:
        return []

    def col(*names):
        for i, h in enumerate(header):
            if any(n in h for n in names):
                return i
        return None

    ci_name = col("raw material", "material")
    ci_qty = col("qty")
    ci_unit = col("unit", "uom")
    ci_mw = col("m.wt", "m. wt", "mol. wt", "mw")
    ci_mol = col("mol.", "mol ")
    ci_ratio = col("mole ratio", "ratio")
    ci_rem = col("remark")
    ci_sr = col("s.no", "s. no", "sr")
    ci_dens = col("density", "dens", "sp. gr", "specific gravity")

    if ci_name is None:
        return []

    mats = []
    hdr_idx = tbl.index(next(r for r in tbl if [(c or "").strip().lower() for c in r] == header))
    for r in tbl[hdr_idx + 1:]:
        if not r or all((c or "").strip() == "" for c in r):
            continue
        name = (r[ci_name] or "").strip() if ci_name < len(r) else ""
        if not name or name.lower() in {"raw material", "material"}:
            continue
        qty = _to_float(r[ci_qty]) if ci_qty is not None and ci_qty < len(r) else None
        unit = (r[ci_unit] or "").strip().lower() if ci_unit is not None and ci_unit < len(r) else ""
        mw = _to_float(r[ci_mw]) if ci_mw is not None and ci_mw < len(r) else None
        dens = _to_float(r[ci_dens]) if ci_dens is not None and ci_dens < len(r) else None
        ratio = (r[ci_ratio] or "").strip() if ci_ratio is not None and ci_ratio < len(r) else ""
        rem = (r[ci_rem] or "").strip() if ci_rem is not None and ci_rem < len(r) else ""
        sr = int(_to_float(r[ci_sr]) or (len(mats) + 1)) if ci_sr is not None and ci_sr < len(r) else len(mats) + 1

        lab_g = qty if unit in _UNIT_SOLID else None
        lab_ml = qty if unit in _UNIT_LIQ else None
        # if unit missing, guess: has MW & 'g' vibes -> solid; else liquid
        if lab_g is None and lab_ml is None and qty is not None:
            lab_g = qty if mw else None
            lab_ml = qty if not mw else None

        remarks = (ratio + (" | " if ratio and rem else "") + rem).strip()
        mats.append(RawMaterial(sr_no=sr, name=name, lab_qty_g=lab_g,
                                lab_vol_ml=lab_ml, mw=mw, density=dens, remarks=remarks))
    return mats


# ---------------------------------------------------------------------------
def to_stage_calc(ps: ParsedStage, scale_up_ratio: float,
                  yield_override: Optional[float] = None,
                  project: str = "") -> StageCalc:
    """Convert a ParsedStage + manual ratio into a ready-to-compute StageCalc."""
    yld = yield_override if yield_override is not None else (ps.yield_frac or 0.8)
    lab_in = ps.lab_input_g or (ps.materials[0].lab_qty_g if ps.materials else 100.0) or 100.0
    return StageCalc(
        name=ps.name, project=project,
        lab_input_g=lab_in,
        input_mw=ps.input_mw or (ps.materials[0].mw if ps.materials else 1.0) or 1.0,
        product_mw=ps.product_mw or 1.0,
        scale_up_ratio=scale_up_ratio, yield_frac=yld,
        materials=ps.materials, steps=ps.steps,
    )
