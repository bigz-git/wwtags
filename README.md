# wwtags

A command-line tool that generates Wonderware tag import CSV files from Excel templates.

Instead of manually entering hundreds of tags, you define your devices once in an Excel workbook and let `wwtags` generate the import file automatically.

---

## Installation

Navigate to the folder containing the `.whl` file and run:

```bash
pip install wwtags-0.3.1-py3-none-any.whl
```

To upgrade from a previous version:

```bash
pip install --upgrade wwtags-0.3.1-py3-none-any.whl
```

Verify the installation:

```bash
wwtags --version
```

---

## How It Works

`wwtags` reads an Excel workbook that contains two types of sheets:

1. **`DEVICE_LIST`** — one row per device, listing its tag name, PLC address, device type, etc.
2. **Template sheets** — one sheet per device type (e.g. `VFD`, `Pump`, `S7 Template`), defining the tag structure for that type.

For each device in `DEVICE_LIST`, the tool finds its matching template sheet, substitutes the device-specific values into the template rows, and writes the result to a CSV file ready to import into Wonderware.

---

## Setting Up the Excel Workbook

### The `DEVICE_LIST` Sheet

The first sheet must be named `DEVICE_LIST`. Each row represents one device.

| Column | Required | Description |
|--------|----------|-------------|
| `HMI_TAG` | Yes | The tag name used in the HMI (e.g. `PUMP_01`) |
| `ACCESS_NAME` | Yes | The access name configured in Wonderware (e.g. `PLC1`) |
| `DEVICE_TYPE` | Yes | Name of the template sheet to use for this device |
| `PLC_TAG` | No | PLC-side address or DB name (e.g. `DB61`) |
| `COMMENT` | No | Description of the device |
| `ALARM_GROUP` | No | Alarm group name |
| `OFFSET` | No | Numeric base offset for Step 7 DB addressing |

**Example rows:**

| HMI_TAG | PLC_TAG | COMMENT | ACCESS_NAME | ALARM_GROUP | DEVICE_TYPE | OFFSET |
|---------|---------|---------|-------------|-------------|-------------|--------|
| PUMP_01 | DB61 | Feed Pump 1 | PLC1 | Pumps | S7 Template | 20 |
| VFD_02 | DB62 | Conveyor Drive | PLC1 | Drives | VFD | 40 |

> Blank rows in the sheet are automatically skipped.

---

### Template Sheets

Each template sheet defines the tag structure for one device type. The sheet name must exactly match the `DEVICE_TYPE` value used in `DEVICE_LIST`.

Template sheets use Wonderware's CSV format with **control rows** and **data rows**:

- **Control rows** start with `:` (e.g. `:IODisc`, `:IOInt`, `:IOReal`) and define the tag type and column headers for the rows that follow.
- **Data rows** contain the actual tag values, using **placeholders** that get replaced with each device's data.

#### Placeholders

Use these tokens in your template data rows and they will be replaced automatically:

| Placeholder | Replaced With |
|-------------|---------------|
| `HMI_TAG` | Value from the `HMI_TAG` column |
| `PLC_TAG` | Value from the `PLC_TAG` column |
| `COMMENT001` | Value from the `COMMENT` column |
| `ACCESS_NAME` | Value from the `ACCESS_NAME` column |
| `ALARM_GROUP` | Value from the `ALARM_GROUP` column |
| `{OFFSET}` | The raw value from the `OFFSET` column |
| `{OFFSET+N}` | `OFFSET` value plus N (e.g. `{OFFSET+4}` adds 4 to the offset) |

The `{OFFSET+N}` placeholder is useful when a device maps to multiple consecutive DB addresses. For example, if a device starts at offset 20, `{OFFSET+0}` gives `20`, `{OFFSET+2}` gives `22`, `{OFFSET+4}` gives `24`, and so on.

---

## Usage

### Basic — Generate the import file

```bash
wwtags my_workbook.xlsx
```

This creates `ww_tag_import.csv` in the current directory.

### Specify a custom output filename

```bash
wwtags my_workbook.xlsx --output site_a_tags.csv
```

### Inspect the workbook before generating

List which template sheets are available:

```bash
wwtags my_workbook.xlsx --list-templates
```

List the columns present in `DEVICE_LIST` (and whether they are required, optional, or unrecognised):

```bash
wwtags my_workbook.xlsx --list-columns
```

### Validate without writing output

Run a dry run to check for errors and see how many tags would be generated, without writing any file:

```bash
wwtags my_workbook.xlsx --dry-run
```

### Process only a subset of devices

Use `--filter` to process only rows where a column matches a specific value:

```bash
wwtags my_workbook.xlsx --filter DEVICE_TYPE=VFD
wwtags my_workbook.xlsx --filter ALARM_GROUP=Pumps
```

### Strict mode

By default the tool continues processing even if some rows have warnings (e.g. a missing optional field). In strict mode it stops immediately on the first error:

```bash
wwtags my_workbook.xlsx --strict
```

---

## Recommended Workflow

1. **Inspect** — confirm templates and columns look correct:
   ```bash
   wwtags my_workbook.xlsx --list-templates
   wwtags my_workbook.xlsx --list-columns
   ```

2. **Validate** — check for errors before generating:
   ```bash
   wwtags my_workbook.xlsx --dry-run
   ```

3. **Generate** — create the import file:
   ```bash
   wwtags my_workbook.xlsx --output ww_import.csv
   ```

4. **Import** into Wonderware using the generated CSV.

---

## Output Format

The generated CSV follows Wonderware's standard import format:

- First line is always `:mode=replace`
- Control rows (`:IODisc`, `:IOInt`, `:IOReal`, etc.) define the tag type and columns
- Data rows contain the expanded tag values

Example:

```
:mode=replace
:IODisc,Group,Comment,AccessName,...
PUMP_01_RUN,Pumps,Feed Pump 1 Running,PLC1,...
PUMP_01_FAULT,Pumps,Feed Pump 1 Fault,PLC1,...
:IOInt,Group,Comment,AccessName,...
PUMP_01_SPEED,Pumps,Feed Pump 1 Speed,PLC1,...
```

---

## Troubleshooting

| Problem | Check |
|---------|-------|
| `No template found for DEVICE_TYPE 'X'` | The sheet name in the workbook must exactly match the `DEVICE_TYPE` value (case-sensitive) |
| Missing required field errors | Ensure every row in `DEVICE_LIST` has `HMI_TAG`, `ACCESS_NAME`, and `DEVICE_TYPE` filled in |
| `OFFSET` not substituting correctly | The `OFFSET` column value must be a number; use `{OFFSET+0}` not `{OFFSET}` for consistency |
| Unexpected blank output | Run with `--dry-run` to see how many tags are being found, and `--list-templates` to verify sheet names |
