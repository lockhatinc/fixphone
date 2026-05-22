"""fixphone — find busted South African telephone numbers in CSV/XLSX files.

Run:  uvx fixphone INPUT.csv         # CSV (zero deps)
      uvx fixphone INPUT.xlsx        # XLSX (pulls openpyxl)
      pipx run fixphone INPUT.xlsx   # equivalent
"""
from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Iterator

PHONE_HEADER_RE = re.compile(r"(?i)(?:^|[\s_\-/])(phone|tel(?:ephone)?|mobile|cell(?:ular)?|fax|whatsapp|msisdn|contact[\s_\-]*(?:no|num|number))(?:$|[\s_\-/])|^(phone|tel(?:ephone)?|mobile|cell(?:ular)?|fax|whatsapp|msisdn)$")

# Shape gate: must look phone-ish before we even try to validate.
# Allows digits + phone separators; optional leading +.
SHAPE_RE = re.compile(r"^\s*\+?[\d\s\-().\/]{6,25}\s*$")

# Phone-prefix gate: a bare digit blob (e.g. "8001015009087" or "20240115") is
# NOT a phone candidate. To count, the value must have either:
#   - a leading + (international form), or
#   - a leading 0 or 00 (SA national or international), or
#   - any of the typical phone separators (space, dash, dot, paren, slash).
PHONE_PREFIX_RE = re.compile(r"^\s*(\+|0|[\d]+[\s\-().\/])")
SEPARATOR_RE = re.compile(r"[\s\-().\/]")

# Approximate set of country-code lengths we recognize for "+X..." prefix extraction.
# We don't need a full table — anything not +27 is rejected as foreign anyway;
# we just want to report the right country code in the reason.
_CC_PREFIXES_1 = {"1", "7"}
_CC_PREFIXES_2 = {"20", "27", "30", "31", "32", "33", "34", "36", "39", "40", "41", "43", "44", "45", "46", "47", "48", "49",
                  "51", "52", "53", "54", "55", "56", "57", "58", "60", "61", "62", "63", "64", "65", "66", "81", "82", "84",
                  "86", "90", "91", "92", "93", "94", "95", "98"}
def _looks_like_phone_attempt(s: str) -> bool:
    """A value is a phone-entry attempt if it has phone-like shape: leading +,
    leading 0, parenthesized leading 0, OR a digit-only blob whose shape
    matches a recognizable SA pattern (starts with 27 / 0027 / SA mobile lead
    digit and falls in 9-12 digit range). Pure non-phone strings like
    '#ERROR!', short numerics like '12345.0', or generic digit blobs starting
    with 9 or non-SA leads are not phone-entry attempts."""
    s = s.strip()
    if not s:
        return False
    # Excel scientific notation: only treat as a phone attempt if decoding to
    # an integer yields something phone-shaped (10-12 digits starting with 27,
    # or 11+ digits — phones are typically 9-12 digits total, anything shorter
    # is a tiny number). Excel mangles long phone-like integers into
    # '2.782123456E9' form. A real fractional number like '3.14159' has very
    # few digits and decodes to '3', so it won't be misclassified.
    sci = re.fullmatch(r"-?\d+\.\d+[eE][+-]?\d+", s) or re.fullmatch(r"-?\d+[eE][+-]?\d+", s)
    if sci:
        try:
            n = abs(int(float(s)))
        except (ValueError, OverflowError):
            return False
        d = str(n)
        if d.startswith("27") and 10 <= len(d) <= 12:
            return True
        if len(d) >= 10 and d[0] in "678":
            return True
        return False
    # Pure numeric literals (decimals, currency-like) are data-noise.
    if re.fullmatch(r"-?\d+\.\d+", s):
        return False
    if re.fullmatch(r"-?\d+\.0+", s):  # Excel-style "100.0"
        return False
    if re.fullmatch(r"[#].+[!]?", s):   # '#ERROR!', '#N/A'
        return False
    digits = re.sub(r"\D", "", s)
    lead = s.lstrip("(").lstrip()
    # Phone-shape lead beats currency/unit detection: '082 123 4567 ext 123'
    # is a phone attempt, but '£20' or '50kg' are not.
    if lead.startswith("+") or lead.startswith("0"):
        return True
    if re.match(r"^[£$€¥]", s) or re.search(r"\d\s*[a-zA-Z]{1,3}$", s):
        return False
    # Pure-digit attempts (Excel may have stripped formatting/leading 0/+).
    if digits.startswith("0027") and 12 <= len(digits) <= 14:
        return True
    if digits.startswith("27") and 10 <= len(digits) <= 12:
        return True
    if len(digits) in (9, 10) and digits[:1] in "678":
        return True
    return False


def _extract_cc(digits: str) -> str:
    if digits[:1] in _CC_PREFIXES_1:
        return digits[:1]
    if digits[:2] in _CC_PREFIXES_2:
        return digits[:2]
    if digits[:3]:
        return digits[:3]
    return digits

# SA geographic area codes (leading digit after the 0 or +27).
# Mobile: 6, 7, 8.  Geo/landline: 1-5.  (0 and 9 are not valid SA subscriber leads.)
SA_VALID_LEAD = set("12345678")


@dataclass(frozen=True)
class Result:
    ok: bool
    reason: str = ""
    normalized: str = ""


def validate_sa(raw: str) -> Result:
    """Validate a single value as a South African phone number.

    Returns ok=False with a reason when the value is *clearly a phone candidate*
    that is busted. Returns ok=True for valid SA numbers. Returns ok=True with
    reason="not-a-phone" for values that don't look like phones at all — the
    caller decides whether that's a false positive based on column context.
    """
    if raw is None:
        return Result(True, "empty")
    s = str(raw).strip()
    if not s:
        return Result(True, "empty")

    if not SHAPE_RE.match(s):
        return Result(True, "not-a-phone")

    has_plus = s.lstrip().startswith("+")
    has_sep = bool(SEPARATOR_RE.search(s))
    digits = re.sub(r"\D", "", s)

    if not digits:
        return Result(True, "not-a-phone")

    # Phone-prefix gate (operates on digits, not raw chars, so "(082) ..." works):
    # A phone candidate must be marked by EITHER a leading + OR digits that begin
    # with a recognizable phone prefix (0 = national, 0027 = SA international).
    # Bare "27..." without a + is NOT a phone — that's an account number, not an
    # international number (the + is mandatory for international form).
    if not has_plus:
        if not digits.startswith("0"):
            # Sole exception: 9-digit subscriber with separators like "82 123 4567".
            if len(digits) == 9 and digits[0] in SA_VALID_LEAD and has_sep:
                return Result(False, "missing-leading-0-or-+27")
            return Result(True, "not-a-phone")

    # Foreign country code that isn't +27.
    if has_plus and not digits.startswith("27"):
        return Result(False, f"foreign-country-code:+{_extract_cc(digits)}")

    # Normalize to national 0XXXXXXXXX (10 digits).
    if has_plus and digits.startswith("27"):
        national = "0" + digits[2:]
    elif digits.startswith("0027"):
        national = "0" + digits[4:]
    elif digits.startswith("0"):
        national = digits
    else:
        return Result(False, f"unrecognized-prefix:{digits[:3]}")

    # Structural sanity on the raw input (catches malformed-but-right-length cases
    # like trailing dots or wrong digit groupings).
    stripped = s.strip()
    if stripped and not stripped[-1].isdigit():
        return Result(False, "trailing-non-digit")
    # Inspect first run of digits in the raw string — for a national-format number
    # the first run is the area code prefix; if it isn't 3 digits or the full 10,
    # the groupings are wrong (e.g. "08 2123 4567").
    if national[0] == "0" and not has_plus and has_sep:
        first_run = re.match(r"\D*(\d+)", stripped)
        if first_run:
            n_first = len(first_run.group(1))
            if n_first not in (3, 10):  # 082 ... or unseparated 0821234567
                return Result(False, "malformed-grouping")

    if len(national) != 10:
        return Result(False, f"wrong-length:{len(national)}-digits")

    if national[0] != "0":
        return Result(False, "missing-leading-0")

    lead = national[1]
    if lead not in SA_VALID_LEAD:
        return Result(False, f"invalid-area-lead:{lead}")

    return Result(True, "valid", normalized=national)


# ---------- readers ----------

def _read_csv(path: Path) -> tuple[list[str], Iterator[list[str]]]:
    f = open(path, "r", encoding="utf-8-sig", newline="")
    reader = csv.reader(f)
    try:
        header = next(reader)
    except StopIteration:
        header = []
    return header, ((row) for row in reader)


def _read_xlsx(path: Path) -> tuple[list[str], Iterator[list[str]]]:
    """Read an .xlsx with the stdlib only — no openpyxl.

    .xlsx is a ZIP of XML. We pull the shared-strings table and the first
    worksheet, walk rows in document order, and expand each cell to its
    string value. Good enough for finding phone numbers — we don't care
    about formatting, formulas, styles, or types beyond "what would the
    user see in the cell".
    """
    import xml.etree.ElementTree as ET
    import zipfile

    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

    def col_letters_to_index(s: str) -> int:
        n = 0
        for ch in s:
            if not ("A" <= ch <= "Z"):
                break
            n = n * 26 + (ord(ch) - ord("A") + 1)
        return n - 1

    try:
        zf = zipfile.ZipFile(str(path))
    except zipfile.BadZipFile:
        sys.exit(f"error: {path} is not a valid .xlsx (not a zip archive)")

    with zf:
        names = set(zf.namelist())

        # Shared strings table (optional).
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            with zf.open("xl/sharedStrings.xml") as f:
                for _, si in ET.iterparse(f, events=("end",)):
                    if si.tag == NS + "si":
                        # Concatenate all <t> descendants (handles rich-text runs).
                        text_parts = [t.text or "" for t in si.iter(NS + "t")]
                        shared.append("".join(text_parts))
                        si.clear()

        # Pick the first worksheet by walking the workbook part.
        sheet_path = "xl/worksheets/sheet1.xml"
        if "xl/workbook.xml" in names:
            with zf.open("xl/workbook.xml") as f:
                wb_root = ET.parse(f).getroot()
            sheets = wb_root.find(NS + "sheets")
            if sheets is not None and len(sheets) > 0:
                first = sheets[0]
                # Resolve r:id → target via workbook rels.
                rid = first.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
                if rid and "xl/_rels/workbook.xml.rels" in names:
                    with zf.open("xl/_rels/workbook.xml.rels") as rf:
                        rels = ET.parse(rf).getroot()
                    for rel in rels:
                        if rel.get("Id") == rid:
                            target = rel.get("Target", "")
                            sheet_path = "xl/" + target.lstrip("/").removeprefix("xl/").lstrip("/")
                            break

        if sheet_path not in names:
            # Fallback: first thing under xl/worksheets/.
            cands = sorted(n for n in names if n.startswith("xl/worksheets/") and n.endswith(".xml"))
            if not cands:
                sys.exit(f"error: no worksheets found in {path}")
            sheet_path = cands[0]

        def cell_text(c) -> str:
            t = c.get("t")
            if t == "inlineStr":
                is_el = c.find(NS + "is")
                if is_el is None:
                    return ""
                return "".join(tt.text or "" for tt in is_el.iter(NS + "t"))
            v = c.find(NS + "v")
            if v is None or v.text is None:
                return ""
            val = v.text
            if t == "s":
                try:
                    return shared[int(val)]
                except (ValueError, IndexError):
                    return ""
            if t == "b":
                return "TRUE" if val == "1" else "FALSE"
            return val

        def iter_rows_from_sheet() -> Iterator[list[str]]:
            with zf.open(sheet_path) as f:
                for _, row in ET.iterparse(f, events=("end",)):
                    if row.tag != NS + "row":
                        continue
                    out: list[str] = []
                    for c in row.findall(NS + "c"):
                        ref = c.get("r", "")
                        idx = col_letters_to_index(ref) if ref else len(out)
                        while len(out) < idx:
                            out.append("")
                        out.append(cell_text(c))
                    row.clear()
                    yield out

        # We need to fully materialize so we can close the zip after the generator
        # finishes — iterparse holds the stream open. For our use case (phone
        # validation on staff spreadsheets) row count is tiny; this is fine.
        all_rows = list(iter_rows_from_sheet())

    if not all_rows:
        return [], iter(())
    header = all_rows[0]
    return header, iter(all_rows[1:])


def read_table(path: Path) -> tuple[list[str], Iterator[list[str]]]:
    ext = path.suffix.lower()
    if ext == ".csv":
        return _read_csv(path)
    if ext in (".xlsx", ".xlsm"):
        return _read_xlsx(path)
    sys.exit(f"error: unsupported file type {ext!r}. supported: .csv .xlsx")


# ---------- scanning ----------

@dataclass
class Busted:
    row: int          # 1-based; row 1 = header, data starts at row 2
    column: str
    value: str
    reason: str


def pick_phone_columns(header: list[str], explicit: list[str] | None, all_columns: bool) -> list[int]:
    if all_columns:
        return list(range(len(header)))
    if explicit:
        names = {h.strip().lower(): i for i, h in enumerate(header)}
        out = []
        for c in explicit:
            i = names.get(c.strip().lower())
            if i is None:
                sys.exit(f"error: column {c!r} not found. headers: {header}")
            out.append(i)
        return out
    return [i for i, h in enumerate(header) if PHONE_HEADER_RE.search(h or "")]


def scan(path: Path, columns: list[str] | None, all_columns: bool) -> tuple[list[Busted], int, list[str]]:
    header, rows = read_table(path)
    idxs = pick_phone_columns(header, columns, all_columns)
    if not idxs:
        # No phone columns detected and user didn't force --all-columns:
        # refuse to scan the whole sheet — that's how false positives happen.
        sys.exit(
            "error: no phone-like columns found by header name.\n"
            "  headers seen: " + ", ".join(repr(h) for h in header) + "\n"
            "  pass --column NAME to target one, or --all-columns to scan everything (will use shape-gate)."
        )

    # Strict columns are the ones the user clearly *named* as phones (explicit
    # --column or header matched PHONE_HEADER_RE). In those, anything with
    # phone-attempt shape that isn't a valid SA number is busted. Values that
    # clearly aren't phone-entry attempts (#ERROR! cells, short numerics,
    # bare digit blobs without SA shape) are not flagged — they're noise that
    # belongs in a phone column by mistake, not broken phone entries.
    strict = set(idxs) if (columns or not all_columns) else set()

    busted: list[Busted] = []
    scanned = 0
    for r, row in enumerate(rows, start=2):
        for i in idxs:
            if i >= len(row):
                continue
            val = row[i]
            scanned += 1
            res = validate_sa(val)
            colname = header[i] if i < len(header) else f"col{i+1}"
            if res.ok and res.reason == "valid":
                continue
            if res.ok and res.reason == "empty":
                continue
            if res.ok and res.reason == "not-a-phone":
                if i not in strict:
                    continue
                if not _looks_like_phone_attempt(str(val)):
                    continue
                busted.append(Busted(row=r, column=colname, value=str(val), reason="not-a-valid-phone"))
                continue
            # validate_sa returned busted, but in strict mode we still verify
            # the value was actually a phone-entry attempt — otherwise we'd
            # flag decimal noise like '0.5' that trips length/grouping checks.
            if i in strict and not _looks_like_phone_attempt(str(val)):
                continue
            busted.append(Busted(row=r, column=colname, value=str(val), reason=res.reason))
    return busted, scanned, [header[i] if i < len(header) else f"col{i+1}" for i in idxs]


# ---------- CLI ----------

def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        prog="fixphone",
        description="Find busted South African telephone numbers in a CSV or XLSX file.",
    )
    p.add_argument("input", help="path to .csv or .xlsx file")
    p.add_argument("--column", "-c", action="append",
                   help="column name to scan (repeatable). default: auto-detect by header.")
    p.add_argument("--all-columns", action="store_true",
                   help="scan every column (uses shape gate to avoid false positives).")
    p.add_argument("--output", "-o", help="write CSV report to this path.")
    p.add_argument("--quiet", "-q", action="store_true", help="suppress per-row stdout output.")
    args = p.parse_args(argv)

    path = Path(args.input)
    if not path.exists():
        print(f"error: {path} does not exist", file=sys.stderr)
        return 2

    busted, scanned, scanned_cols = scan(path, args.column, args.all_columns)

    if not args.quiet:
        print(f"scanned {scanned} cell(s) across columns: {', '.join(scanned_cols)}")
        if not busted:
            print("no busted SA phone numbers found.")
        else:
            print(f"found {len(busted)} busted number(s):")
            for b in busted:
                print(f"  row {b.row}  [{b.column}]  {b.value!r}  -> {b.reason}")

    if args.output:
        with open(args.output, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            w.writerow(["row", "column", "value", "reason"])
            for b in busted:
                w.writerow([b.row, b.column, b.value, b.reason])
        if not args.quiet:
            print(f"report written to {args.output}")

    return 1 if busted else 0


def _is_frozen() -> bool:
    """True when running from a PyInstaller-built exe."""
    return getattr(sys, "frozen", False)


def _pause_if_interactive() -> None:
    """When launched by double-click / drag-and-drop / Send-To, the console
    window disappears the moment the process exits. Pause so the user can
    read the output. Skipped when redirected (CI, piping, --quiet flag of a
    shell user)."""
    if not sys.stdin or not sys.stdin.isatty():
        return
    try:
        input("\npress Enter to close...")
    except (EOFError, KeyboardInterrupt):
        pass


def _frozen_entry() -> int:
    """Entry used by the PyInstaller exe. Adds drag-and-drop ergonomics:
    if no file was passed, prompt for one; on completion, pause."""
    argv = sys.argv[1:]
    if not argv:
        try:
            entered = input("drag a .csv or .xlsx file onto this window, or type a path, then Enter: ").strip().strip('"')
        except (EOFError, KeyboardInterrupt):
            entered = ""
        if not entered:
            print("no file given. exiting.")
            _pause_if_interactive()
            return 2
        argv = [entered]
    try:
        rc = main(argv)
    except SystemExit as e:
        rc = int(e.code) if isinstance(e.code, int) else (0 if e.code is None else 1)
    except Exception as e:  # noqa: BLE001
        print(f"error: {e}", file=sys.stderr)
        rc = 1
    _pause_if_interactive()
    return rc


if __name__ == "__main__":
    if _is_frozen() or len(sys.argv) <= 1:
        raise SystemExit(_frozen_entry())
    raise SystemExit(main())
