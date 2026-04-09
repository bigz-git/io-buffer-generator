# IO Buffer Generator

Python tool for generating Rockwell Studio 5000 `.l5x` import files from an Excel project workbook. Produces IO buffer routines and module status programs for Point IO, Flex IO, and ControlLogix IO racks.

Available as both a CLI (`io_buffer_tool.py`) and a GUI (`gui.py`).

---

## Setup

> **TODO:** Installation instructions will be updated once the tool is packaged as a wheel.
>
> For now, install dependencies manually:
> ```
> pip install -r requirements.txt
> ```
> Requires Python 3.10+.

---

## Workflow Overview

```
1. init          — Create a new project workbook
2. add-rack      — Add a rack (one sheet per rack)
3. (open Excel)  — Fill in Module Type, PLC Routine Name, Tag Names, Descriptions
4. fill-tags     — (optional) Auto-fill blank tag names from module type + routine name
5. fill-desc     — (optional) Fill blank descriptions with "spare"
6. validate      — Check for errors before generating
7. generate      — Write .l5x file(s)
8. generate-cad  — (optional, WIP) Write CAD description .xlsx
```

---

## Workbook Structure

Each project is stored in a single `.xlsx` workbook:

| Sheet | Contents |
|---|---|
| Cover Sheet | Project metadata (Software Version, Controller Name, IO Network Card Name, Project Number, Project Description) and rack summary table |
| CLI Tool Help | Built-in command reference |
| *(one sheet per rack)* | Module data — see columns below |

**Rack sheet columns:**

| Col | Field | Notes |
|---|---|---|
| A | Module Type | Dropdown: Input, Output, Safety Input, Safety Output, Analog Input, Analog Output, Thermocouple/RTD, Other |
| B | Module Slot Number | Auto-filled by `add-rack`; merged per module |
| C | PLC Routine Name | Enter manually in Excel (or leave placeholder and use `fill-tags`) |
| D | I/O Bit | 0-indexed channel number; auto-filled by `add-rack` |
| E | I/O Buffer Tag Name | Enter manually or use `fill-tags` to auto-generate |
| F | I/O Buffer Tag Description | Enter manually or use `fill-descriptions` to fill blanks with "spare" |

**Module types:**

| Type | Tag format | Notes |
|---|---|---|
| Input / Output | `ROUTINE.bit` (DINT) | Standard digital IO |
| Safety Input / Safety Output | `ROUTINE.bit` (DINT, Safety class) | Point IO only |
| Analog Input / Output | `ROUTINE[index]` (INT array) | |
| Thermocouple/RTD | `ROUTINE[index]` (INT array) | |
| Other | *(no buffer tag)* | JSR enable bit and GSV fault detect only |

**IO families supported:** 1734 (Point IO), 1794 (Flex IO), 1756 (ControlLogix IO).

---

## CLI Usage

```
python io_buffer_tool.py <command> [--workbook <path>] [--output <dir>]
```

If `--workbook` is not specified, the tool automatically uses the first `.xlsx` file found in the current directory.

If `--output` is not specified, generated files are written to the same directory as the workbook.

### Commands

| Command | Description |
|---|---|
| `init` | Create a new project workbook with a Cover Sheet |
| `add-rack` | Add a rack sheet (prompts for rack name, IO family, number of modules, and channels per module) |
| `rename-rack` | Rename an existing rack sheet and update the Cover Sheet |
| `remove-rack` | Remove a rack sheet and its Cover Sheet entry |
| `add-module` | Append modules to an existing rack |
| `fill-tags` | Auto-fill blank tag names in column E |
| `fill-descriptions` | Fill blank descriptions in column F with "spare" |
| `validate` | Check the workbook for errors and warnings without generating any files |
| `list` | Print a summary of all racks and modules |
| `generate` | Generate `.l5x` file(s) from the workbook |
| `generate-cad` | *(WIP)* Generate a CAD description `.xlsx` from the workbook |

**Examples:**
```
python io_buffer_tool.py init
python io_buffer_tool.py add-rack --workbook MyProject.xlsx
python io_buffer_tool.py validate
python io_buffer_tool.py generate --output ./output
```

### `fill-tags` — tag name convention

Tag names are generated from the module type and the PLC routine name. For the drawing number to be embedded in the tag, the routine name must begin with `R` followed by a 4-digit drawing number (e.g. `R4103_IB8`). If the routine name does not follow this format, `fill-tags` will still run but the tag will use `XXXX` as a placeholder (e.g. `DI_XXXX.0`).

Existing tag values are never overwritten. Rows where the module type is not set are skipped with a warning.

---

## GUI Usage

```
python gui.py
```

The GUI provides the same functionality as the CLI in a Tkinter window. On launch it auto-detects any `.xlsx` in the current directory.

**Buttons available:**

- **Setup:** Add Rack, Add Module, Rename Rack, Remove Rack
- **Populate:** Fill Tags, Fill Descriptions
- **Tools:** List, Validate, Generate

> Note: `generate-cad` is not currently exposed in the GUI.

---

## Output Files

`generate` produces up to two `.l5x` files in the output directory:

| File | Contents |
|---|---|
| `IO_Files_Rev_<mmddyy_hhmm>.l5x` | Standard IO buffer routines + module status program (all non-safety modules) |
| `Safety_IO_Files_Rev_<mmddyy_hhmm>.l5x` | Safety IO buffer routines + module status program (only written if Safety Input or Safety Output modules are present) |

These files are importable into Rockwell Studio 5000.

`generate-cad` produces:

| File | Contents |
|---|---|
| `CAD_Descriptions_<ddmmyy_hhmm>.xlsx` | Tag descriptions formatted for CAD import, one sheet per rack |

---

## CAD Generation (Work in Progress)

`generate-cad` is functional but incomplete:

- Only tested with **Point IO** modules.
- Module output format is determined by the suffix after the last `_` in the routine name (e.g. `R4103_IB8` → suffix `IB8`). If a routine name has no `_`, the tool will prompt for the module type identifier at runtime.
- Not all Point IO module types are mapped yet. Unknown module types fall back to Format A (consecutive descriptions) and a note is written in column D of the output row.
- Safety IO modules are not yet supported in CAD output.
