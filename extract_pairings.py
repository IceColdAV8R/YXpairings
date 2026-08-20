#!/usr/bin/env python3
"""Extract Republic Airways pairing-packet PDFs into PairingsViewer JSON.

Reads two-column UA / DL / AA bid packets (and pdftotext -layout .txt files)
and writes the same structure as pairings.json:

  { metadata: { generated, schedMonth, schedYear, startDate, endDate },
    pairings: [ { code, base, reportTime, days, totals, operatingDates, length } ] }

Examples:
  python3 extract_pairings.py ua.pdf dl.pdf aa.pdf -o pairings.json
  python3 extract_pairings.py ~/Downloads/September\\ 2026*Pairings.pdf
  python3 extract_pairings.py packet.txt --dump-text unwrapped.txt

Requires: pymupdf (PyMuPDF) for PDFs.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

try:
    import pymupdf
except ImportError:  # pragma: no cover
    pymupdf = None


HEADER_RE = re.compile(
    r"(?P<sMon>\w+)\s+(?P<sYr>20\d{2})\s+"
    r"(?:(?:AA|UA|DL)\s+Pilot|Pilot\s+(?:AA|UA|DL))\s+Pairings\s+"
    r"(?P<mBgn>\w{3})\s*(?P<dBgn>\d{1,2})\s*-\s*"
    r"(?P<mEnd>\w{3})\s*(?P<dEnd>\d{1,2})",
    re.I,
)
SECTION_RE = re.compile(r"\s([A-Z]\d{4})(?:.*\n)*?(?=\s?={72})")
CODE_RE = re.compile(r"^[A-Z]\d{4}$")
LEG_RE = re.compile(
    r"(?P<DAY>\w{2}|\d)\s*(?P<DH>DH)?\s*(?P<FLTN>\d{4})\s*"
    r"(?P<DPS>[A-Z]{3})-(?P<ARS>[A-Z]{3})\s*"
    r"(?P<DEPL>\d{4})\s*(?P<ARRL>\d{4})\s*"
    r"(?P<BLKT>\d{1,3})\s*(?P<GRNT>\d{1,3})?\s*"
    r"(?P<EQP>\w[A-Z\d]{1,2})"
    r"(?:\s+(?P<TBLK>\d{1,4})\s+(?P<TCRD>\d{1,4})\s*(?P<TPAY>\d{1,4})\s+"
    r"(?P<DUTY>\d{1,4})\s*(?P<LAYO>\d{3,4})?)?",
    re.I,
)
DEND_RE = re.compile(
    r"D-END:\s*(?P<DEND>\d{4}L)\s*"
    r"(?:REPT:\s*(?P<REPT>\d{4}L)\s*)?"
    r"FDP:\s*(?P<FDP>\d{4})\s*FDPLim:\s*(?P<FDPlim>\d{4})",
    re.I,
)
TOTALS_RE = re.compile(
    r"TOTALS\sBLK\s+(?P<TrBLK>\d{1,4})\sDHD\s+(?P<TrDHD>\d{1,4})\s"
    r"(?:TRIP\sRIG|PRG\sCRED):\s+(?P<TrRIG>\d{1,4})\sCDT\s*(?P<TrCDT>\d{1,4})\s"
    r"T.A.F.B.\s+(?P<TAFB>\d{1,5})\s+LDGS:\s+(?P<LDGS>\d{1,2})"
)
BASE_RE = re.compile(r"Base\s*:\s*(?P<BASE>\w{3})", re.I)
BRPT_RE = re.compile(r"BASE\sREPT:\s*(?P<BRPT>\d{4}L)", re.I)
CAL_HEADER_RE = re.compile(r"Mo\s+Tu\s+We\s+Th\s+Fr\s+Sa\s+Su")

MONTHS = {
    "JAN": 0, "FEB": 1, "MAR": 2, "APR": 3, "MAY": 4, "JUN": 5,
    "JUL": 6, "AUG": 7, "SEP": 8, "OCT": 9, "NOV": 10, "DEC": 11,
    "JANUARY": 0, "FEBRUARY": 1, "MARCH": 2, "APRIL": 3, "MAY": 4, "JUNE": 5,
    "JULY": 6, "AUGUST": 7, "SEPTEMBER": 8, "OCTOBER": 9, "NOVEMBER": 10, "DECEMBER": 11,
}
MMM = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
       "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
WEEKDAYS = ["MO", "TU", "WE", "TH", "FR", "SA", "SU"]
TXT_COLUMN_WIDTH = 76
LEG_FIELD_ORDER = [
    "DAY", "DH", "FLTN", "DPS", "ARS", "DEPL", "ARRL",
    "BLKT", "GRNT", "EQP", "TBLK", "TCRD", "TPAY", "DUTY", "LAYO",
]
DEND_FIELD_ORDER = ["DEND", "REPT", "FDP", "FDPlim"]
TOTALS_FIELD_ORDER = ["TrBLK", "TrDHD", "TrRIG", "TrCDT", "TAFB", "LDGS"]


def named_groups(regex: re.Pattern, text: str) -> dict[str, str] | None:
    match = regex.search(text)
    if not match:
        return None
    return {k: v for k, v in match.groupdict().items() if v}


def ordered(data: dict[str, str], fields: list[str]) -> dict[str, str]:
    return {k: data[k] for k in fields if k in data}


def collapse_ws(text: str) -> str:
    return " ".join(text.split())


def unwrap_two_column_text(text: str, width: int = TXT_COLUMN_WIDTH) -> str:
    """Split pdftotext -layout lines into left column, then overflow column."""
    main: list[str] = []
    overflow: list[str] = []
    for line in text.splitlines():
        line = line.replace("\x0c", "")
        if len(line) <= width:
            main.append(line)
        else:
            main.append(line[:width])
            overflow.append(line[width:])
    if overflow:
        return "\n".join(main) + "\n\n" + "\n".join(overflow)
    return "\n".join(main)


def extract_text_from_pdf(path: Path) -> str:
    """Read a pairing packet PDF, left column then right column per page."""
    if pymupdf is None:
        raise SystemExit(
            "PDF input requires pymupdf. Install with: pip install pymupdf"
        )
    doc = pymupdf.open(path)
    chunks: list[str] = []
    try:
        for page in doc:
            mid = page.rect.width / 2.0
            left = page.get_text(
                "text", clip=pymupdf.Rect(0, 0, mid, page.rect.height)
            )
            right = page.get_text(
                "text",
                clip=pymupdf.Rect(mid, 0, page.rect.width, page.rect.height),
            )
            chunks.append(left)
            if right.strip():
                chunks.append(right)
    finally:
        doc.close()
    return "\n\n".join(chunks)


def load_source(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return extract_text_from_pdf(path)
    if suffix in {".txt", ".text"}:
        return unwrap_two_column_text(path.read_text(encoding="utf-8", errors="replace"))
    raise SystemExit(f"Unsupported file type: {path}")


def parse_header(text: str) -> dict[str, str] | None:
    match = HEADER_RE.search(text)
    if not match:
        return None
    g = match.groupdict()
    return {
        "schedMonth": g["sMon"].strip(),
        "schedYear": g["sYr"],
        "startDate": (g["mBgn"].upper() + f"{int(g['dBgn']):02d}")[:5],
        "endDate": (g["mEnd"].upper() + f"{int(g['dEnd']):02d}")[:5],
    }


def split_sections(text: str) -> list[dict[str, str]]:
    sections: list[dict[str, str]] = []
    for match in SECTION_RE.finditer(text):
        sections.append({"name": match.group(1), "content": match.group(0)})
    sections.sort(key=lambda s: s["name"])
    return sections


def insert_standover_days(pairing: dict) -> None:
    days = pairing["days"]
    if len(days) <= 1:
        return
    if all(re.fullmatch(r"\d+", d["dayKey"]) for d in days):
        days.sort(key=lambda d: int(d["dayKey"]))
        i = 0
        while i < len(days) - 1:
            expected = int(days[i]["dayKey"]) + 1
            if int(days[i + 1]["dayKey"]) != expected:
                days.insert(
                    i + 1,
                    {
                        "dayKey": str(expected),
                        "legs": [],
                        "dayEnd": None,
                        "isStandover": True,
                    },
                )
            else:
                i += 1
        return

    i = 0
    while i < len(days) - 1:
        current = days[i]["dayKey"].upper()
        nxt = days[i + 1]["dayKey"].upper()
        if current not in WEEKDAYS:
            i += 1
            continue
        expected = WEEKDAYS[(WEEKDAYS.index(current) + 1) % 7]
        if nxt != expected:
            days.insert(
                i + 1,
                {
                    "dayKey": expected,
                    "legs": [],
                    "dayEnd": None,
                    "isStandover": True,
                },
            )
        else:
            i += 1


def parse_pairing(text: str) -> dict:
    pairing: dict = {
        "base": None,
        "reportTime": None,
        "days": [],
        "totals": None,
    }
    lines = text.splitlines()

    for line in lines:
        base = named_groups(BASE_RE, line)
        if base:
            pairing["base"] = base["BASE"]
        rept = named_groups(BRPT_RE, line)
        if rept:
            pairing["reportTime"] = rept["BRPT"]

    current_day = None
    i = 0
    while i < len(lines):
        trimmed = lines[i].strip()
        if not trimmed:
            i += 1
            continue

        leg = named_groups(LEG_RE, trimmed)
        if leg:
            day_key = leg["DAY"]
            day_entry = next((d for d in pairing["days"] if d["dayKey"] == day_key), None)
            if day_entry is None:
                day_entry = {"dayKey": day_key, "legs": [], "dayEnd": None}
                pairing["days"].append(day_entry)
            day_entry["legs"].append(ordered(leg, LEG_FIELD_ORDER))
            current_day = day_key
            i += 1
            continue

        dend = named_groups(DEND_RE, trimmed)
        if dend and current_day:
            day_entry = next((d for d in pairing["days"] if d["dayKey"] == current_day), None)
            if day_entry is not None:
                day_entry["dayEnd"] = ordered(dend, DEND_FIELD_ORDER)
                hotel_i = i + 1
                nxt = lines[hotel_i].strip() if hotel_i < len(lines) else ""
                if nxt and re.search(r"STANDOVER", nxt, re.I):
                    hotel_i += 1
                    nxt = lines[hotel_i].strip() if hotel_i < len(lines) else ""
                if (
                    nxt
                    and not LEG_RE.search(nxt)
                    and not DEND_RE.search(nxt)
                    and not re.search(r"TOTALS", nxt, re.I)
                    and not re.fullmatch(r"-{10,}", nxt)
                    and not nxt.startswith("=")
                ):
                    day_entry["hotel"] = collapse_ws(nxt)
            i += 1
            continue

        i += 1

    insert_standover_days(pairing)

    for line in lines:
        totals = named_groups(TOTALS_RE, line)
        if totals:
            pairing["totals"] = ordered(totals, TOTALS_FIELD_ORDER)
            break

    return pairing


def extract_operating_dates(
    text: str, start_date: str | None, end_date: str | None, year: str | None
) -> list[str]:
    if not text or not start_date or not end_date or not year:
        return []

    lines = text.splitlines()
    cal_start = -1
    for line in lines:
        if CAL_HEADER_RE.search(line):
            cal_start = line.find("Mo")
            break
    if cal_start == -1:
        return []

    tokens: list[str] = []
    in_cal = False
    for line in lines:
        part = line[cal_start:].strip() if len(line) > cal_start else ""
        if in_cal and part == "":
            break
        if part:
            in_cal = True
            found = [t for t in part.split() if t == "--" or t.isdigit()]
            if found:
                tokens.extend(found)
    if not tokens:
        return []

    start_month = MONTHS.get(start_date[:3].upper(), 3)
    end_month = MONTHS.get(end_date[:3].upper(), 3)
    current = date(int(year), start_month + 1, int(start_date[3:]))
    last = date(int(year), end_month + 1, int(end_date[3:]))

    dates: list[str] = []
    idx = 0
    while current <= last and idx < len(tokens):
        if tokens[idx] != "--":
            dates.append(MMM[current.month - 1] + f"{current.day:02d}")
        current += timedelta(days=1)
        idx += 1
    return dates


def extract_pairings(texts: list[str]) -> dict:
    header = None
    for text in texts:
        header = parse_header(text)
        if header:
            break

    combined = ("\n\n" + "=" * 80 + "\n\n").join(texts)
    if header is None:
        header = parse_header(combined)

    start_date = header["startDate"] if header else None
    end_date = header["endDate"] if header else None
    year = header["schedYear"] if header else None

    pairings = []
    for section in split_sections(combined):
        if not CODE_RE.match(section["name"]):
            continue
        pairing = parse_pairing(section["content"])
        pairing["code"] = section["name"]
        pairing["operatingDates"] = extract_operating_dates(
            section["content"], start_date, end_date, year
        )
        pairing["length"] = len(pairing["days"])
        # Keep field order stable for diffs against PairingsViewer JSON.
        pairings.append(
            {
                "base": pairing["base"],
                "reportTime": pairing["reportTime"],
                "days": pairing["days"],
                "totals": pairing["totals"],
                "code": pairing["code"],
                "operatingDates": pairing["operatingDates"],
                "length": pairing["length"],
            }
        )

    now = datetime.now(timezone.utc)
    metadata = {
        "generated": now.strftime("%Y-%m-%dT%H:%M:%S.") + f"{now.microsecond // 1000:03d}Z",
        "schedMonth": header["schedMonth"] if header else None,
        "schedYear": header["schedYear"] if header else None,
        "startDate": start_date,
        "endDate": end_date,
    }
    return {"metadata": metadata, "pairings": pairings}


def collect_inputs(raw_paths: list[str]) -> list[Path]:
    out: list[Path] = []
    for raw in raw_paths:
        path = Path(raw).expanduser()
        if not path.exists():
            raise SystemExit(f"Not found: {path}")
        if path.is_dir():
            out.extend(sorted(path.glob("*.pdf")))
            out.extend(sorted(path.glob("*.txt")))
        else:
            out.append(path)
    if not out:
        raise SystemExit("No PDF or .txt pairing files given")
    return out


def summarize(data: dict, sources: list[Path], stream) -> None:
    pairings = data["pairings"]
    meta = data["metadata"]
    missing_legs = sum(1 for p in pairings if not p["days"])
    missing_dates = sum(1 for p in pairings if not p["operatingDates"])
    missing_totals = sum(1 for p in pairings if not p["totals"])
    bases = sorted({p["base"] for p in pairings if p["base"]})
    stream.write(
        f"Sources: {len(sources)} file(s)\n"
        f"Header: {meta['schedMonth']} {meta['schedYear']} "
        f"{meta['startDate']}-{meta['endDate']}\n"
        f"Pairings: {len(pairings)}\n"
        f"Bases: {', '.join(bases)}\n"
    )
    if missing_legs or missing_dates or missing_totals:
        stream.write(
            f"Warnings: no-legs={missing_legs} no-dates={missing_dates} "
            f"no-totals={missing_totals}\n"
        )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Extract UA/DL/AA pairing packet PDFs into pairings.json"
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        help="PDF or pdftotext .txt files, or a directory of them",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="pairings.json",
        help="Output JSON path (default: pairings.json)",
    )
    parser.add_argument(
        "--dump-text",
        metavar="FILE",
        help="Write the unwrapped packet text (debug)",
    )
    parser.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        help="Only print errors",
    )
    args = parser.parse_args(argv)

    sources = collect_inputs(args.inputs)
    texts = [load_source(path) for path in sources]

    if args.dump_text:
        Path(args.dump_text).write_text("\n\n".join(texts), encoding="utf-8")

    data = extract_pairings(texts)
    out_path = Path(args.output)
    out_path.write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")

    if not args.quiet:
        summarize(data, sources, sys.stderr)
        sys.stderr.write(f"Wrote {out_path} ({out_path.stat().st_size} bytes)\n")
        if not data["pairings"]:
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
