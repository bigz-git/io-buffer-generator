# Point IO Buffer Generator — Python Rewrite Plan

## Overview

Python CLI tool that replicates a VBA/Excel-based tool for generating `.l5x` files
importable into Rockwell Studio 5000 PLC programming software.

---

## Workflow

```
python tool.py init          # Create new project workbook + cover sheet
python tool.py add-rack      # Interactive prompts → adds formatted sheet to workbook
python tool.py generate      # Read workbook → output .l5x file(s)
```

### 1. `init`
- Prompt for: Software Version, Controller Name, IO Network Card Name
- Create a new `.xlsx` workbook with a formatted Cover Sheet
- Cover Sheet cells:
  - B2: Software Version
  - B3: Controller Name
  - B4: IO Network Card Name
  - A6:B onwards: rack name / IO count summary (auto-populated by `add-rack`)

### 2. `add-rack`
- Prompt for:
  - Rack name (becomes the sheet name)
  - Number of IO modules in the rack
  - For each module: number of bits/channels (always ask — no defaults)
- Python writes a new sheet to the workbook with:
  - Headers: Module Type | Module Slot Number | PLC Routine Name | I/O Bit | I/O Buffer Tag Name | I/O Buffer Tag Description | Drawing File Name
  - One row per bit, grouped by module
  - Module Slot Number pre-filled (1-indexed)
  - I/O Bit pre-filled (0-indexed within module)
  - Module Type column: dropdown validation (Input, Output, Safety Input, Safety Output, Analog Input, Analog Output, Thermocouple/RTD)
  - PLC Routine Name: placeholder text "ENTER ROUTINE NAME HERE"
  - Drawing File Name: placeholder text "ENTER DRAWING NAME HERE"
  - Columns A, B, C, G merged per module block
  - "End" sentinel written in column B after last row
- Updates Cover Sheet rack/IO-count summary row

### 3. `add-module` (add modules to existing rack)
- Prompt for rack name (or select from list), number of new modules, bits per module
- Appends to existing rack sheet (removing "End" sentinel first)

### 4. `generate`
- Read Cover Sheet for: Software Version, Controller Name, IO Network Card Name
- Loop through all rack sheets (sheet index >= 3, i.e. after Cover Sheet and any fixed sheets)
- For each rack, parse modules by locating rows where column B (slot number) is non-empty
- Build and write `.l5x` output file(s)

---

## L5X Output

Two possible output files per `generate` run:

### `IO_Files_Revision_<mmddyy_hhmm>.l5x` (Standard)
Structure:
```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<RSLogix5000Content SchemaRevision="1.0" SoftwareRevision="<ver>"
  TargetName="IO_Buffer_Files" TargetType="Program" TargetClass="Standard" ...>
  <Controller Use="Context" Name="<controller>">
    <DataTypes Use="Context">
      <!-- QP_PLC_TAGS_v02, QP_MODULE_TAGS_v01 UDT definitions -->
    </DataTypes>
    <Tags Use="Context">
      <!-- PLC tag (QP_PLC_TAGS_v02) -->
      <!-- Per-rack QP_MODULE_TAGS_v01 tag -->
      <!-- Per-module: _S_FAULT BOOL tag, buffer tag (DINT or INT array) -->
    </Tags>
    <Programs Use="Context">
      <Program ... Name="IO_Buffer_Files" MainRoutineName="Subroutine_Calls" Class="Standard">
        <Tags>
          <!-- JSR_ENABLE_<routine> BOOL local tags, one per standard module -->
        </Tags>
        <Routines>
          <!-- Subroutine_Calls routine: XIC(JSR_ENABLE_x)JSR(x,0) per module -->
          <!-- Per-module buffer routine (XIC/OTE for digital, MOV for analog) -->
        </Routines>
      </Program>
      <Program ... Name="IO_Module_Status" MainRoutineName="Subroutine_Calls" Class="Standard">
        <Tags>
          <!-- JSR_ENABLE_<rack> BOOL local tags, one per rack -->
        </Tags>
        <Routines>
          <!-- Subroutine_Calls routine -->
          <!-- Per-rack module status routine (GSV fault detection, slot status bits) -->
        </Routines>
      </Program>
    </Programs>
  </Controller>
</RSLogix5000Content>
```

### `Safety_Buffer_Revision_<mmddyy_hhmm>.l5x` (Safety — only if safety modules present)
- Same structure but `TargetClass="Safety"`, `TargetName="IO_Safety_Buffer_Files"`
- Only includes Safety Input / Safety Output modules
- Safety controller tags use `Safety` class instead of `Standard`

---

## Ladder Logic Generation by Module Type

| Module Type       | Data Rung                                                                  |
|-------------------|---------------------------------------------------------------------------|
| Input             | `XIC(<rack>:<slot>:I.<bit>)OTE(<tag>)`                                    |
| Output            | `XIC(<tag>)OTE(<rack>:<slot>:O.<bit>)`                                    |
| Analog Input      | `MOV(<rack>:<slot>:I.Ch<bit>Data,<tag>)`                                  |
| Analog Output     | `MOV(<tag>,<rack>:<slot>:O.Ch<bit>Data)`                                  |
| Thermocouple/RTD  | `MOV(<rack>:<slot>:I.Ch<bit>Data,<tag>)`                                  |
| Safety Input      | `XIC(<rack>:<slot>:I.Pt<bit:02d>Data)OTE(<tag>)` + status NOP rung        |
| Safety Output     | `XIC(<tag>)OTE(<rack>:<slot>:O.Pt<bit:02d>Data)`                          |

Tag format:
- Digital (Input/Output/Safety): `DINT` — tag referenced as `TagName.bitN`
- Analog / TC/RTD: `INT` array — tag referenced as `TagName[N]`
- Safety tags use `Safety` class; standard tags use `Standard` class

---

## Module Fault Detection (IO_Module_Status routines)

Each rack gets one routine. First rung is MCR with fault/reset logic for the AENT module.
Subsequent rungs are per-module:
- Slot numbers 1–31: reference `<rack>:I.SlotStatusBits0_31.<slot>`
- Slot numbers 32–62: reference `<rack>:I.SlotStatusBits32_63.<slot - 32>`
- Analog / TC/RTD / Safety: use `GSV(Module,<module_base>,FaultCode,<module_base>._S_FaultCode)`

---

## UDTs (Deferred)

Two UDTs are currently hardcoded in the VBA tool:
- `QP_PLC_TAGS_v02` — generic PLC tag UDT
- `QP_MODULE_TAGS_v01` — per-rack/module tag UDT

Plan: make UDT names configurable (either via CLI prompts at `init` time, or via
dedicated cells on the Cover Sheet). XML body for each UDT either hardcoded as
defaults or loaded from an external file. **To be designed.**

---

## CAD Description Export (Deferred)

The VBA tool has a `D_Generate_CAD.bas` module that reads a `CAD_Descriptions`
sheet and exports a sorted `.XLS` file. **Out of scope for initial implementation.**

---

## Technology Stack

- **Python 3.x**
- **openpyxl** — read/write `.xlsx` workbooks (sheet management, dropdowns, merges)
- **Standard library** — `xml.etree.ElementTree` or string templating for L5X output
- **click** or **argparse** — CLI framework (TBD, interactive prompts preferred)

---

## Key Data Structures (proposed)

```python
@dataclass
class Bit:
    index: int         # 0-indexed bit/channel number within module
    tag: str           # buffer tag name (e.g. "CONV_01_IN.0")
    description: str   # tag description
    drawing: str       # drawing file name (shared per module, stored on first bit row)

@dataclass
class Module:
    slot: int          # slot number as entered by user
    type: str          # "Input", "Output", "Safety Input", etc.
    routine: str       # PLC routine name
    bits: list[Bit]

@dataclass
class Rack:
    name: str
    modules: list[Module]

@dataclass
class Project:
    software_version: str
    controller_name: str
    io_network_card: str
    racks: list[Rack]
```

---

## File Naming

- Project workbook: user-defined name, e.g. `MyProject.xlsx`
- Standard L5X: `IO_Files_Revision_<mmddyy_hhmm>.l5x`
- Safety L5X: `Safety_Buffer_Revision_<mmddyy_hhmm>.l5x`
- Output written to same directory as the workbook

---

## Open Questions

- [ ] UDT storage strategy (hardcoded defaults vs. Cover Sheet cells vs. external file)
- [ ] CAD description export (future scope)
- [ ] Whether to support a `clear` / `reset` command (VBA had a button that wiped all rack sheets)
- [ ] CLI framework: `click` vs `argparse`
- [ ] How to specify which `.xlsx` file to operate on (flag, env var, or always in current directory?)
