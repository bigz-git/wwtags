import argparse
import csv
from openpyxl import load_workbook


PLACEHOLDER_MAP = {
    "HMI_TAG": "HMI_TAG",
    "PLC_TAG": "PLC_TAG",
    "COMMENT001": "COMMENT",
    "ACCESS_NAME": "ACCESS NAME",
    "ALARM_GROUP": "ALARM GROUP",
}

REQUIRED_COLUMNS = {
    "HMI_TAG",
    "PLC_TAG",
    "DEVICE_TYPE",
}


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate Wonderware tag import CSV from Excel templates"
    )
    parser.add_argument(
        "workbook", help="Excel workbook containing DEVICE_LIST and templates")
    parser.add_argument(
        "--output",
        default="ww_tag_import.csv",
        help="Output CSV file (default: ww_tag_import.csv)",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Fail immediately on missing templates or required fields",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate and count tags without writing output file",
    )
    return parser.parse_args()


def read_device_list(ws, strict=False):
    headers = [cell.value for cell in ws[1]]

    missing = REQUIRED_COLUMNS - set(headers)
    if missing:
        raise RuntimeError(f"Missing required columns: {', '.join(missing)}")

    devices = []

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        record = dict(zip(headers, row))

        for field in REQUIRED_COLUMNS:
            if not record.get(field):
                if strict:
                    raise RuntimeError(
                        f"Row {row_idx}: missing required value '{field}'"
                    )
                else:
                    record = None
                    break

        if record:
            devices.append(record)

    return devices


def is_control_row(row):  # helper function to count only rows with tags
    return isinstance(row[0], str) and row[0].startswith(":")


def expand_template(ws, device):
    expanded_rows = []

    for row in ws.iter_rows(values_only=True):
        row = list(row)

        new_row = []
        for cell in row:
            if isinstance(cell, str):
                for token, column in PLACEHOLDER_MAP.items():
                    value = device.get(column, "") or ""
                    cell = cell.replace(token, str(value))
            new_row.append(cell)

        expanded_rows.append(new_row)

    return expanded_rows


def main():
    args = parse_args()

    wb = load_workbook(args.workbook, data_only=True)

    if "DEVICE_LIST" not in wb.sheetnames:
        raise RuntimeError("DEVICE_LIST sheet not found")

    device_ws = wb["DEVICE_LIST"]
    devices = read_device_list(device_ws, strict=args.strict)

    output_rows = []

    tag_count = 0

    for device in devices:
        sheet_name = device["DEVICE_TYPE"]

        if sheet_name not in wb.sheetnames:
            msg = f"Template sheet '{sheet_name}' not found"
            if args.strict:
                raise RuntimeError(msg)
            else:
                print(f"WARNING: {msg} — skipping")
                continue

        template_ws = wb[sheet_name]
        rows = expand_template(template_ws, device)
        output_rows.extend(rows)

        for row in rows:
            if row and not is_control_row(row):
                tag_count += 1

    if args.dry_run:
        print("Dry run enabled — no output file written")
    else:
        with open(args.output, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([":mode=replace"])
            for row in output_rows:
                writer.writerow(["" if v is None else v for v in row])

    if args.dry_run:
        print(f"[DRY RUN] {tag_count} tags would be generated")
    else:
        print(f"Generated {tag_count} tags → {args.output}")


if __name__ == "__main__":
    main()
