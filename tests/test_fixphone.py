"""Accuracy guard for fixphone.

The point of this suite is not coverage theatre — it is the false-positive guard:
non-phone data (IDs, dates, postal codes, names, amounts) must NOT be flagged as
busted phone numbers. Real busted numbers must be flagged with a reason.
"""
import csv
from pathlib import Path

import pytest

from fixphone import validate_sa, scan, pick_phone_columns, PHONE_HEADER_RE


# --- validate_sa --------------------------------------------------------------

VALID = [
    "+27 82 123 4567",
    "+27821234567",
    "082 123 4567",
    "0821234567",
    "082-123-4567",
    "(082) 123 4567",
    "+27 11 555 0100",
    "011 555 0100",
    "0027821234567",
    "021 555 1234",   # Cape Town landline
    "031 555 1234",   # Durban landline
]

BUSTED = [
    ("082 123 456", "wrong-length:9-digits"),       # missing a digit
    ("08212345678", "wrong-length:11-digits"),      # extra digit
    ("0921234567", "invalid-area-lead:9"),          # 09x isn't allowed
    ("0021234567", "invalid-area-lead:0"),
    ("+44 20 7946 0958", "foreign-country-code:+44"),
    ("+1 415 555 0100", "foreign-country-code:+1"),
    ("82 123 4567", "missing-leading-0-or-+27"),    # 9-digit subscriber with separators
]

# Bare digit blobs with no separators and no phone prefix must NOT be treated
# as phones at all — that's how IDs/account numbers stay un-flagged.
NOT_A_PHONE_DIGIT_BLOBS = [
    "821234567",        # bare 9-digit
    "123456789",
    "8001015009087",    # SA ID
    "20240115",         # ISO date without separators
]

# These must NOT be flagged when scanned (validate_sa returns ok=True with reason="not-a-phone").
NON_PHONE = [
    "",
    "   ",
    "John Smith",
    "123 Main Road",
    "8001",                   # postal code
    "8001234567890",          # SA ID (13 digits) — yes it'll trip the shape gate? check below
    "R 1,234.56",
    "ORD-2024-001",
    "2024-01-15",
    "info@example.com",
    "Yes",
    "N/A",
]


@pytest.mark.parametrize("v", VALID)
def test_valid_numbers(v):
    r = validate_sa(v)
    assert r.ok, f"{v!r} should be valid, got reason={r.reason}"
    assert r.normalized.startswith("0") and len(r.normalized) == 10


@pytest.mark.parametrize("v,expected_reason", BUSTED)
def test_busted_numbers(v, expected_reason):
    r = validate_sa(v)
    assert not r.ok, f"{v!r} should be busted"
    assert r.reason == expected_reason, f"{v!r}: expected {expected_reason}, got {r.reason}"


@pytest.mark.parametrize("v", NOT_A_PHONE_DIGIT_BLOBS)
def test_digit_blobs_without_phone_prefix_not_flagged(v):
    r = validate_sa(v)
    assert r.ok and r.reason == "not-a-phone", f"{v!r}: ok={r.ok} reason={r.reason}"


@pytest.mark.parametrize("v", NON_PHONE)
def test_non_phone_not_flagged(v):
    """Critical false-positive guard: non-phone strings must not be reported as busted."""
    r = validate_sa(v)
    # Either flagged as not-a-phone (ok=True) — never as a busted phone number.
    assert r.ok, f"non-phone value {v!r} was flagged as busted phone: {r.reason}"


def test_id_number_is_not_flagged_as_busted_phone():
    # 13-digit SA ID shouldn't slip through the shape gate as a phone.
    r = validate_sa("8001015009087")
    assert r.ok, f"SA ID was flagged: {r.reason}"


# --- header detection ---------------------------------------------------------

@pytest.mark.parametrize("h", [
    "Phone", "phone", "Phone Number", "Mobile", "Cell", "cell_no",
    "Telephone", "tel", "Fax", "WhatsApp", "Contact No", "contact_number", "MSISDN",
])
def test_phone_header_matches(h):
    assert PHONE_HEADER_RE.search(h)


@pytest.mark.parametrize("h", [
    "Name", "Address", "Email", "Postal Code", "ID Number", "Amount", "Order ID",
    "Province", "Date of Birth",
])
def test_non_phone_header_does_not_match(h):
    assert not PHONE_HEADER_RE.search(h), f"{h!r} should not be picked as a phone column"


# --- end-to-end CSV scan ------------------------------------------------------

def _write_csv(tmp_path: Path, rows: list[list[str]]) -> Path:
    p = tmp_path / "sample.csv"
    with open(p, "w", encoding="utf-8", newline="") as f:
        csv.writer(f).writerows(rows)
    return p


def test_scan_csv_only_flags_phone_columns(tmp_path: Path):
    p = _write_csv(tmp_path, [
        ["Name", "ID Number", "Postal Code", "Mobile", "Amount"],
        ["Alice", "8001015009087", "8001", "082 123 4567", "1234.56"],   # all clean (ID/postal must NOT flag)
        ["Bob",   "8001015009087", "8001", "082 123 456",  "9999.99"],   # bust on Mobile only
        ["Cara",  "9999999999999", "0001", "+44 20 7946 0958", "0"],     # bust on Mobile only
    ])
    busted, scanned, cols = scan(p, columns=None, all_columns=False)
    assert cols == ["Mobile"], f"only Mobile should be scanned, got {cols}"
    reasons = [(b.row, b.column, b.reason) for b in busted]
    assert reasons == [
        (3, "Mobile", "wrong-length:9-digits"),
        (4, "Mobile", "foreign-country-code:+44"),
    ], reasons


def test_scan_csv_all_columns_does_not_flag_non_phone_values(tmp_path: Path):
    """--all-columns must still avoid flagging clearly-non-phone data."""
    p = _write_csv(tmp_path, [
        ["Name", "ID", "Note"],
        ["Alice", "8001015009087", "called back 2024-01-15"],
        ["Bob",   "Yes",            "R 1,200.00 paid"],
    ])
    busted, scanned, _ = scan(p, columns=None, all_columns=True)
    assert busted == [], f"non-phone all-columns scan produced false positives: {busted}"


def _write_minimal_xlsx(path: Path, rows: list[list[str]]) -> None:
    """Hand-roll a minimal .xlsx so we can test the stdlib reader without
    depending on openpyxl (which is exactly what we're avoiding)."""
    import io
    import zipfile

    NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'

    # Build shared strings.
    strings: list[str] = []
    index: dict[str, int] = {}
    for r in rows:
        for v in r:
            if v not in index:
                index[v] = len(strings)
                strings.append(v)

    def xml_escape(s: str) -> str:
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    sst_items = "".join(f"<si><t xml:space=\"preserve\">{xml_escape(s)}</t></si>" for s in strings)
    sst = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst {NS} count="{len(strings)}" uniqueCount="{len(strings)}">{sst_items}</sst>'

    def col_letter(i: int) -> str:
        s = ""
        i += 1
        while i:
            i, rem = divmod(i - 1, 26)
            s = chr(ord("A") + rem) + s
        return s

    sheet_rows = []
    for ri, row in enumerate(rows, start=1):
        cells = "".join(
            f'<c r="{col_letter(ci)}{ri}" t="s"><v>{index[v]}</v></c>'
            for ci, v in enumerate(row)
        )
        sheet_rows.append(f"<row r=\"{ri}\">{cells}</row>")
    sheet = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet {NS}><sheetData>{"".join(sheet_rows)}</sheetData></worksheet>'

    content_types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>'
    root_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    workbook = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook {NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
    workbook_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>'

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/sharedStrings.xml", sst)
        z.writestr("xl/worksheets/sheet1.xml", sheet)


def test_strict_phone_column_catches_all_busted_variants(tmp_path: Path):
    """Regression for the user-validation spreadsheet: a column explicitly
    named for phone numbers must flag every malformed variant, not just
    length errors. All 20 rows below should be flagged."""
    p = _write_csv(tmp_path, [
        ["Broken SA Phone Numbers"],
        ["082123456"],            # short
        ["08212345678"],          # long
        ["+270821234567"],        # +27 then leading 0 then full number = 11 digits
        ["27821234567"],          # no +, just 27 prefix (not international form)
        ["082-123-456A"],         # stray letter
        ["082 123 45 6"],         # wrong groupings + short
        ["+28 82 123 4567"],      # foreign cc
        ["00278212345678"],       # too long
        ["821234567"],            # bare 9-digit blob
        ["082!1234567"],          # bad separator
        ["+27 (0) 82 123 45678"], # spurious (0) makes it too long
        ["08 2123 4567"],         # malformed groupings (10 digits)
        ["+27-82-123-456"],       # short under +27
        ["082_123_4567"],         # underscore not a phone sep
        ["+27821234567a"],        # trailing alpha
        ["082.123.4567."],        # trailing dot
        ["082123"],               # way too short
        ["+27 82 123 4567 890"],  # too long
        ["082 123 4567 ext 123"], # extension text
        ["(082) 123-456"],        # short
    ])
    busted, scanned, cols = scan(p, columns=None, all_columns=False)
    assert cols == ["Broken SA Phone Numbers"]
    assert scanned == 20
    assert len(busted) == 20, f"expected 20 busted, got {len(busted)}: " + \
        ", ".join(f"row{b.row}={b.value!r}({b.reason})" for b in busted)


def test_scan_xlsx_stdlib_reader(tmp_path: Path):
    p = tmp_path / "sample.xlsx"
    _write_minimal_xlsx(p, [
        ["Name", "ID Number", "Mobile"],
        ["Alice", "8001015009087", "082 123 4567"],
        ["Bob",   "8001015009087", "082 123 456"],     # busted
        ["Cara",  "9999999999999", "+44 20 7946 0958"], # busted (foreign)
    ])
    busted, _, cols = scan(p, columns=None, all_columns=False)
    assert cols == ["Mobile"]
    reasons = [(b.row, b.column, b.reason) for b in busted]
    assert reasons == [
        (3, "Mobile", "wrong-length:9-digits"),
        (4, "Mobile", "foreign-country-code:+44"),
    ], reasons


def test_scan_csv_explicit_column(tmp_path: Path):
    p = _write_csv(tmp_path, [
        ["Name", "Office", "Home"],
        ["Alice", "021 555 1234", "082 123 456"],
    ])
    busted, _, cols = scan(p, columns=["Home"], all_columns=False)
    assert cols == ["Home"]
    assert len(busted) == 1 and busted[0].column == "Home"
