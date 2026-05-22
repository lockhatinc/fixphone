# fixphone

Find busted South African telephone numbers in CSV and Excel files. One command, no setup.

## Run it (no install)

```bash
uvx fixphone contacts.xlsx
```

or

```bash
pipx run fixphone contacts.xlsx
```

That's it. Works on Windows, macOS, and Linux. The first run pulls the tool and its lone dependency (`openpyxl`, only used for `.xlsx`); subsequent runs are instant.

Don't have `uv` or `pipx`?

- `uv`: `winget install astral-sh.uv` (Windows) · `brew install uv` (macOS) · `curl -LsSf https://astral.sh/uv/install.sh | sh` (Linux)
- `pipx`: `python -m pip install --user pipx`

## What it does

Scans columns whose header looks like a phone column (`phone`, `tel`, `mobile`, `cell`, `fax`, `whatsapp`, `contact_no`, …) and reports values that aren't a valid SA number. It deliberately does **not** scan non-phone columns by default — that's how tools accidentally flag IDs, postal codes, and order numbers.

### What counts as valid

- `+27 82 123 4567`, `082 123 4567`, `0821234567` — mobile (`06/07/08`)
- `+27 11 555 0100`, `011 555 0100` — landline (`01–05`)
- Spaces, dashes, dots, parens, slashes are all fine.

### What gets flagged

- Wrong length, wrong leading digit, foreign country code (`+44 …`)
- Missing leading `0` or `+27`
- Letters or other garbage in a phone column

## Options

```
fixphone INPUT.{csv,xlsx} [--column NAME ...] [--all-columns] [--output report.csv] [--quiet]
```

- `--column / -c` — target a specific column (repeatable).
- `--all-columns` — scan everything; a shape filter keeps non-phone cells from being flagged.
- `--output / -o` — write the busted rows to a CSV report.
- Exit code is non-zero when busted numbers are found (handy for CI).

## Examples

```bash
fixphone clients.csv
fixphone clients.xlsx -c "Cell number" -c "Office"
fixphone export.xlsx --all-columns -o busted.csv
```

## Develop

```bash
git clone https://github.com/lockhatinc/fixphone
cd fixphone
python -m pip install -e ".[dev]"
pytest
```

## License

MIT.
