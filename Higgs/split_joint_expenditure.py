r"""
split_joint_expenditure.py

Splits expenditure rows that carry two client UUIDs into two single-client rows.

Column `client_uuid` (expected to be the first column) may hold one or two UUIDs.
Three shapes are handled:

  1. Single UUID            -> row passes through, amount untouched.
  2. Two UUIDs in one cell  -> "xxx,yyy" (quoted) or "xxx | yyy". Cell is split.
  3. Two UUIDs unquoted     -> "xxx,yyy" written without quotes, so the CSV reader
                               saw them as two separate fields and the row is one
                               column too wide. The leading UUID-shaped fields are
                               detected and stitched back before splitting.

For every joint row the amount is divided in two. The two halves are computed so
they sum back to the original exactly: the first row gets round(total / 2, 2) and
the second gets the remainder, avoiding penny drift.

Every other column is copied verbatim onto both output rows. Only client_uuid and
amount change. If x_client_ref / x_client_name are also joint and need splitting,
that is a separate change.

Usage (PowerShell):
    py .\split_joint_expenditure.py .\input.csv
    py .\split_joint_expenditure.py .\input.csv .\output.csv

With no output path the result is written alongside the input as
<inputname>_split.csv. The source is never overwritten.
"""

import csv
import os
import re
import sys
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

UUID_COLUMN = "client_uuid"
AMOUNT_COLUMN = "amount"

# Separators tried when two UUIDs share a single cell. " | " before "," so the
# pipe wins if both somehow appear.
CELL_SEPARATORS = [" | ", ","]

# Canonical UUID shape: 8-4-4-4-12 hex. Used to detect and stitch ragged rows.
UUID_RE = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-"
    r"[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$"
)


def looks_like_uuid(value):
    return bool(UUID_RE.match((value or "").strip()))


def split_cell_uuids(raw):
    """Split a single cell that may contain one or two UUIDs. Returns a list."""
    value = (raw or "").strip()
    if not value:
        return [""]
    for sep in CELL_SEPARATORS:
        if sep in value:
            parts = [p.strip() for p in value.split(sep)]
            parts = [p for p in parts if p]
            if len(parts) >= 2:
                return parts
    return [value]


def parse_amount(raw):
    """Parse a currency-ish string into a Decimal. Returns None if not numeric."""
    if raw is None:
        return None
    text = re.sub(r"[^\d.\-]", "", str(raw).strip())
    if text in ("", "-", ".", "-."):
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def format_amount(value):
    return str(value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def halve(total):
    """Split a Decimal into two halves summing exactly to the original."""
    first = (total / 2).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    second = total - first
    return first, second


def resolve_row(fields, header, uuid_index):
    """
    Return (uuids, aligned_row) or (None, reason) if the row cannot be resolved.

    aligned_row is a list of len(header) values with the UUID cell holding the
    joined UUID string, so callers can index it by header position.
    """
    n = len(header)

    # Well-formed width: trust the reader's field boundaries.
    if len(fields) == n:
        uuids = split_cell_uuids(fields[uuid_index])
        return uuids, list(fields)

    # Too wide: an unquoted comma inside the UUID cell split it into extra fields.
    # Only safely recoverable when the UUID column is first.
    if len(fields) > n and uuid_index == 0:
        k = 0
        while k < len(fields) and looks_like_uuid(fields[k]):
            k += 1
        # Need at least two leading UUIDs, and the remaining fields must line up
        # with the rest of the header exactly.
        if k >= 2 and (len(fields) - k) == (n - 1):
            uuids = [f.strip() for f in fields[:k]]
            aligned = [",".join(uuids)] + list(fields[k:])
            return uuids, aligned
        return None, "leading UUID fields do not align with the header"

    if len(fields) < n:
        return None, "row has fewer columns than the header"

    return None, "row is wider than the header and UUID column is not first"


def main():
    if len(sys.argv) < 2:
        print("Usage: py split_joint_expenditure.py <input.csv> [output.csv]")
        sys.exit(1)

    in_path = sys.argv[1]
    if not os.path.isfile(in_path):
        print(f"Input file not found: {in_path}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        out_path = sys.argv[2]
    else:
        base, ext = os.path.splitext(in_path)
        out_path = f"{base}_split{ext or '.csv'}"

    if os.path.abspath(out_path) == os.path.abspath(in_path):
        print("Refusing to overwrite the source file. Choose a different output path.")
        sys.exit(1)

    rows_in = 0
    rows_out = 0
    joint_rows = 0
    reconstructed = 0
    blank_rows = 0
    amount_warnings = 0
    unresolved = 0

    with open(in_path, newline="", encoding="utf-8-sig") as f_in:
        reader = csv.reader(f_in)
        try:
            header = next(reader)
        except StopIteration:
            print("Input CSV is empty.")
            sys.exit(1)

        if UUID_COLUMN not in header:
            print(f"Column '{UUID_COLUMN}' not found. Header is: {header}")
            sys.exit(1)
        if AMOUNT_COLUMN not in header:
            print(f"Column '{AMOUNT_COLUMN}' not found. Header is: {header}")
            sys.exit(1)

        uuid_index = header.index(UUID_COLUMN)
        amount_index = header.index(AMOUNT_COLUMN)

        with open(out_path, "w", newline="", encoding="utf-8") as f_out:
            writer = csv.writer(f_out)
            writer.writerow(header)

            for line_no, fields in enumerate(reader, start=2):
                if not fields or all((c or "").strip() == "" for c in fields):
                    blank_rows += 1
                    continue

                rows_in += 1
                was_ragged = len(fields) != len(header)

                uuids, aligned = resolve_row(fields, header, uuid_index)
                if uuids is None:
                    unresolved += 1
                    print(f"Line {line_no}: could not resolve row ({aligned}). Skipped.")
                    continue

                if was_ragged:
                    reconstructed += 1

                # Single UUID (or empty): write the aligned row through unchanged
                # except for normalising the UUID cell to the single value.
                if len(uuids) < 2:
                    out = list(aligned)
                    out[uuid_index] = uuids[0]
                    writer.writerow(out)
                    rows_out += 1
                    continue

                joint_rows += 1
                total = parse_amount(aligned[amount_index])

                if total is None:
                    amount_warnings += 1
                    print(
                        f"Line {line_no}: joint row with non-numeric amount "
                        f"'{aligned[amount_index]}'. UUIDs split, amount left as-is."
                    )
                    for u in uuids[:2]:
                        out = list(aligned)
                        out[uuid_index] = u
                        writer.writerow(out)
                        rows_out += 1
                    continue

                halves = halve(total)
                for u, amt in zip(uuids[:2], halves):
                    out = list(aligned)
                    out[uuid_index] = u
                    out[amount_index] = format_amount(amt)
                    writer.writerow(out)
                    rows_out += 1

    print("Done.")
    print(f"  Input rows:            {rows_in}")
    print(f"  Joint rows split:      {joint_rows}")
    print(f"  Ragged rows rebuilt:   {reconstructed}")
    if blank_rows:
        print(f"  Blank rows skipped:    {blank_rows}")
    if amount_warnings:
        print(f"  Amount warnings:       {amount_warnings} (see lines above)")
    if unresolved:
        print(f"  Unresolved rows:       {unresolved} (skipped, see lines above)")
    print(f"  Output rows:           {rows_out}")
    print(f"  Output written to:     {out_path}")


if __name__ == "__main__":
    main()
