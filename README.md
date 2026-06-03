# wwtags

A tool that generates Wonderware tag import CSV files from an Excel template. Available as both a graphical application (`wwtags-gui`) and a command-line tool (`wwtags`). Use to generate UDT Wonderware tags. Also includes a Studio5000 UDT importer that builds a ready-to-use template sheet directly from a `.L5X` data type export.


---

## Installation

### Step 1 — Install Python

Python 3.9 or later is required. To check if Python is already installed, open **Command Prompt** (press `Win + R`, type `cmd`, press Enter) and run:

```
python --version
```

If a version number is printed (e.g. `Python 3.12.0`), Python is installed and you can skip to Step 2.

If you see an error or the version is below 3.9, download and install Python from [python.org/downloads](https://www.python.org/downloads/). During installation, check the box that says **"Add Python to PATH"** before clicking Install.

---

### Step 2 — Download the installer file

Go to the [Releases](../../releases) page of this repository on GitHub and download the latest `.whl` file (e.g. `wwtags-0.3.2-py3-none-any.whl`). It will save to your Downloads folder.

---

### Step 3 — Install the package

Open **Command Prompt** (or **PowerShell**) and navigate to your Downloads folder:

```
cd %USERPROFILE%\Downloads
```

Then install the package (replace the filename with the version you downloaded):

```
pip install wwtags-0.3.2-py3-none-any.whl
```

You should see output ending with `Successfully installed wwtags-...`.

---

### Step 4 — Verify the installation

```
wwtags --version
```

This should print the version number. Both the GUI and CLI are now available:

- `wwtags-gui` — launches the graphical application
- `wwtags` — the command-line tool

---

### Upgrading to a newer version

Download the new `.whl` from the Releases page, then run the same command:

```
pip install wwtags-0.3.5-py3-none-any.whl
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
| `HMI_TAG` | Yes | The unique part of the tag name used in the HMI (e.g. `PUMP_01`, `PUMP_02`, etc.) |
| `ACCESS_NAME` | Yes | The access name configured in Wonderware (e.g. `PLC1`) |
| `DEVICE_TYPE` | Yes | Name of the template sheet to use for this device |
| `PLC_TAG` | No | The unique part of the tag name used in the PLC (Rockwell) or the DB name (Siemens) (e.g. `PUMP_01` or `DB61`) |
| `COMMENT` | No | Description of the device (this will replace the text `COMMENT001` in the 'Comment' column of the Device Template sheet) |
| `ALARM_GROUP` | No | Alarm group name |
| `OFFSET` | No | Numeric base offset for Step 7 DB addressing |

**Example rows:**

| HMI_TAG | PLC_TAG | COMMENT | ACCESS_NAME | ALARM_GROUP | DEVICE_TYPE | OFFSET |
|---------|---------|---------|-------------|-------------|-------------|--------|
| PUMP_01 | DB61 | Feed Pump 1 | S7_PLC1 | Pumps | S7 Template | 20 |
| PUMP_02 | DB61 | Feed Pump 2 | S7_PLC1 | Pumps | S7 Template | 40 |
| PUMP_03 | DB61 | Feed Pump 3 | S7_PLC1 | Pumps | S7 Template | 60 |
| VFD_01 | VFD1 | Dryer Zone 1 Fan | AB_PLC1 | Drives | VFD | na |
| VFD_02 | VFD2 | Dryer Zone 2 Fan | AB_PLC1 | Drives | VFD | na |

> Blank rows in the sheet are automatically skipped.

---

### Template Sheets

Each template sheet defines the tag structure for one device type. The sheet name must exactly match the `DEVICE_TYPE` value used in `DEVICE_LIST`.

Template sheets use Wonderware's CSV format with **control rows** and **data rows**:

- **Control rows** start with `:` (e.g. `:IODisc`, `:IOInt`, `:IOReal`) and define the tag type and column headers for the rows that follow. 
- **Data rows** contain the actual tag values, using **placeholders** that get replaced with each device's data.

> If you are unsure how to format a Control row, follow these steps:
> - Create a tag of the desired data type in WindowMaker
> - Save and close WindowMaker
> - Export the tag database (DBDump)
> - Open the tag database and find the created tag
> - The first control row above the newly created tag will correspond to that tag's datatype

#### Placeholders

Use these tokens in your template data rows and they will be replaced automatically:

| Placeholder | Replaced With |
|-------------|---------------|
| `HMI_TAG` | Value from the `HMI_TAG` column |
| `PLC_TAG` | Value from the `PLC_TAG` column |
| `COMMENT001` | Value from the `COMMENT` column |
| `ACCESS_NAME` | Value from the `ACCESS_NAME` column |
| `ALARM_GROUP` | Value from the `ALARM_GROUP` column |
| `{OFFSET+N}` (Step 7 Addressing) | `OFFSET` value plus N (e.g. `{OFFSET+0}` gives the raw value, `{OFFSET+4}` adds 4) |

The `{OFFSET+N}` placeholder is useful when a device maps to multiple consecutive DB addresses. For example, if a device starts at offset 20, `{OFFSET+0}` gives `20`, `{OFFSET+2}` gives `22`, `{OFFSET+4}` gives `24`, and so on.

---

## GUI Usage

Launch the graphical application:

```bash
wwtags-gui
```

The window exposes all the same options as the CLI:

| Control | Description |
|---------|-------------|
| **Workbook** | Path to the Excel workbook. Use **Browse…** to open a file picker. |
| **Output** | Path for the generated CSV (default: `ww_tag_import.csv`). Use **Browse…** to choose a save location. |
| **Dry run** | Checkbox — equivalent to `--dry-run`. Validates and counts tags without writing any file. |
| **Filter** | Text field — equivalent to `--filter`. Enter `COL=VAL` to process only matching rows (e.g. `DEVICE_TYPE=VFD`). Leave blank to process all rows. |
| **Generate** | Runs the tag export with the current settings. Disabled until a workbook is selected. |
| **List Templates** | Lists available template sheets in the workbook. |
| **List Columns** | Lists the columns present in `DEVICE_LIST` and whether they are required, optional, or unrecognised. |
| **UDT (.L5X)** | Path to a Studio5000 User Defined Data Type export file. Use **Browse…** to open a file picker. |
| **Import UDT** | Parses the selected `.L5X` file and adds a new template sheet to the workbook named after the UDT. Disabled until both a workbook and a UDT file are selected. |

All output — including validation errors and warnings — appears in the **Log** panel at the bottom of the window.

---

## CLI Usage

### Basic — Generate the import file

```bash
wwtags my_workbook.xlsx
```

This creates `ww_tag_import.csv` in the current directory (assuming that my_workbook.xlsx exists and is formatted properly).

### Specify a custom output filename

```bash
wwtags my_workbook.xlsx --output winder_tag_import.csv
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

### Import a Studio5000 UDT as a template sheet

Export a User Defined Data Type from Studio5000 as an `.L5X` file, then add it as a template sheet to your workbook:

```bash
wwtags my_workbook.xlsx --import-udt my_udt.L5X
```

The sheet is named after the UDT (e.g. `my_udt`) and is added in-place to the workbook. Each member becomes a tag row in the appropriate Wonderware section (`:IODisc`, `:IOReal`, or `:IOInt`). `SINT` packed-bit containers and `TIMER` members are excluded automatically.

The placeholders `HMI_TAG`, `PLC_TAG`, `COMMENT001`, `ACCESS_NAME`, and `ALARM_GROUP` are written into every row so the sheet works with the standard tag generation workflow once it is added to `DEVICE_LIST`.

The command exits with an error if a sheet with the UDT name already exists in the workbook.

---

### Strict mode

By default the tool continues processing even if some rows have warnings (e.g. a missing optional field). In strict mode it stops immediately on the first error:

```bash
wwtags my_workbook.xlsx --strict
```

---

## Recommended Workflow

> All steps below can be performed in the GUI (`wwtags-gui`) or via the CLI commands shown.

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
| `OFFSET` not substituting correctly | The `OFFSET` column value must be a number; use `{OFFSET+0}`, `{OFFSET+2}`, etc. |
| Unexpected blank output | Run with `--dry-run` to see how many tags are being found, and `--list-templates` to verify sheet names |
