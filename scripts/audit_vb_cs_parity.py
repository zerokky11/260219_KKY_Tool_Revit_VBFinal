#!/usr/bin/env python3
import re
from pathlib import Path
from dataclasses import dataclass

ROOT = Path(__file__).resolve().parents[1] / "KKY_Tool_Revit_2019-2023"
AREAS = [ROOT / "Services", ROOT / "UI" / "Hub", ROOT / "Infrastructure", ROOT / "Exports", ROOT / "My Project"]

VB_METHOD_RE = re.compile(r"\b(?:Public|Private|Friend|Protected)?\s*(?:Shared\s+)?(?:Function|Sub)\s+([A-Za-z_][A-Za-z0-9_]*)", re.I)
CS_METHOD_RE = re.compile(r"\b(?:public|private|internal|protected)\s+(?:static\s+)?[A-Za-z0-9_<>,\[\]\.? ]+\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(", re.I)

NOISE = {
    "end", "public", "private", "friend", "if", "for", "new", "case", "dim", "handles", "then", "else", "loop"
}

@dataclass
class Row:
    vb: Path
    cs: Path
    vb_lines: int
    cs_lines: int
    vb_methods: set
    cs_methods: set

    @property
    def missing(self):
        return sorted(self.vb_methods - self.cs_methods)

    @property
    def missing_count(self):
        return len(self.missing)

    @property
    def line_ratio(self):
        return 0.0 if self.vb_lines == 0 else self.cs_lines / self.vb_lines


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig", errors="ignore")


def extract_vb_methods(text: str) -> set:
    return {m.lower() for m in VB_METHOD_RE.findall(text) if m and m.lower() not in NOISE}


def extract_cs_methods(text: str) -> set:
    return {m.lower() for m in CS_METHOD_RE.findall(text) if m and m.lower() not in NOISE}


def risk_label(row: Row) -> str:
    if row.missing_count >= 20 or row.line_ratio < 0.35:
        return "HIGH"
    if row.missing_count >= 8 or row.line_ratio < 0.6:
        return "MEDIUM"
    return "LOW"


def gather_rows():
    rows = []
    for area in AREAS:
        if not area.exists():
            continue
        for vb in sorted(area.glob("*.vb")):
            cs = vb.with_suffix(".cs")
            if not cs.exists():
                continue
            vb_txt = read_text(vb)
            cs_txt = read_text(cs)
            rows.append(
                Row(
                    vb=vb,
                    cs=cs,
                    vb_lines=len(vb_txt.splitlines()),
                    cs_lines=len(cs_txt.splitlines()),
                    vb_methods=extract_vb_methods(vb_txt),
                    cs_methods=extract_cs_methods(cs_txt),
                )
            )
    return rows


def main():
    rows = gather_rows()
    rows.sort(key=lambda r: (-r.missing_count, r.vb.as_posix()))

    high = [r for r in rows if risk_label(r) == "HIGH"]
    medium = [r for r in rows if risk_label(r) == "MEDIUM"]

    out = []
    out.append("# VB->C# Parity Audit\n")
    out.append(f"- Total pairs: {len(rows)}")
    out.append(f"- HIGH risk: {len(high)}")
    out.append(f"- MEDIUM risk: {len(medium)}")
    out.append("")
    out.append("## Summary Table")
    out.append("|Risk|VB File|CS File|VB Lines|CS Lines|Line Ratio|VB Methods|CS Methods|Missing Methods|")
    out.append("|---|---|---:|---:|---:|---:|---:|---:|---:|")

    for r in rows:
        rel_vb = r.vb.relative_to(ROOT).as_posix()
        rel_cs = r.cs.relative_to(ROOT).as_posix()
        out.append(
            f"|{risk_label(r)}|`{rel_vb}`|`{rel_cs}`|{r.vb_lines}|{r.cs_lines}|{r.line_ratio:.2f}|{len(r.vb_methods)}|{len(r.cs_methods)}|{r.missing_count}|"
        )

    out.append("\n## Missing Method Samples (Top 12 per file)")
    for r in rows:
        if r.missing_count == 0:
            continue
        rel_vb = r.vb.relative_to(ROOT).as_posix()
        sample = ", ".join(r.missing[:12])
        out.append(f"- `{rel_vb}`: {sample}")

    audit_path = Path(__file__).resolve().parents[1] / "PARITY_AUDIT.md"
    audit_path.write_text("\n".join(out) + "\n", encoding="utf-8")
    print(audit_path)


if __name__ == "__main__":
    main()
