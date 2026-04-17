# Aurora P6 Excel Converter

A Python tool that converts project schedule data from an Excel template into
valid Primavera P6 18.8 XML for import. Fill in the provided
`P6_Import_Template.xlsx` with your WBS, activity, and relationship data, run
the converter, and get two ready-to-import XML files.

---

## What This Tool Does

The converter reads `P6_Import_Template.xlsx` and produces two Primavera P6
XML files that together import a complete project schedule — including WBS
structure, activities, and relationships — into P6 Professional 18.8. Because
P6 18.8 cannot resolve relationship foreign-key constraints within a new-project
import transaction (activities are not yet committed to the database when
relationships are processed), the output is split into two files that must be
imported in sequence.

---

## Requirements

- **Python 3.9+**
- **openpyxl** — `pip install openpyxl`
- **p6_reference.xml** — a P6 XML export from your own environment (see setup below)

---

## Setup: Get Your p6_reference.xml

Before running the converter, export any existing project from **your** P6:

1. Open P6 Professional
2. **File → Export → Primavera P6 XML**
3. Select any project and export
4. Rename the exported file to **`p6_reference.xml`**
5. Place it in the same folder as `aurora_p6_converter.py`

The converter reads Calendar, OBS, Currency, UDFType, Role, and RoleRate
elements from this file. P6 requires these in every import XML or the import
crashes during CleanupActivities. Because these elements contain ObjectIds
specific to each P6 database, every user must supply their own reference file.

> `p6_reference.xml` is listed in `.gitignore` — it is not committed to the repository.

---

## Files

| File | Purpose |
|---|---|
| `aurora_p6_converter.py` | Main converter — reads template, writes pass1 + pass2 XML |
| `create_template.py` | Generates `P6_Import_Template.xlsx` with sample rows from your p6_reference.xml |
| `P6_Import_Template.xlsx` | Fill this with your project data |

---

## How to Use

### Step 1 — Generate a blank template (first time only)

```
py create_template.py
```

This reads `p6_reference.xml` and writes `P6_Import_Template.xlsx` pre-filled
with sample rows from that reference project.

### Step 2 — Fill the template

Open `P6_Import_Template.xlsx` and populate the following sheets with your data:

| Sheet | What to fill |
|---|---|
| `_Config` | **CalendarObjectId**, **OBSObjectId**, **CurrencyObjectId** from your P6 environment |
| `Project` | One row: project Id, Name, dates, EPS parent, WBS root |
| `WBS` | One row per WBS node |
| `Activity` | One row per activity (Type, dates, status, percent complete) |
| `Relationship` | One row per dependency (predecessor, successor, type, lag) |
| `Resource` | Optional — resource definitions |
| `ActivityCodeType` / `ActivityCode` | Optional — activity code definitions |

Row 1 = column names, Row 2 = data types (do not edit), Row 3+ = your data.

### Step 3 — Run the converter

```
py aurora_p6_converter.py P6_Import_Template.xlsx
```

Produces two files in the same folder as the template:

| File | Import as |
|---|---|
| `P6_Import_pass1.xml` | **Create New Project** |
| `P6_Import_pass2.xml` | **Update Existing Project** |

---

## Import Instructions (P6 Professional 18.8)

> **Important:** Both passes must be imported in order. Do not skip pass 1.

### Pass 1 — Create the project

1. In P6: **File → Import → Primavera P6 XML**
2. Select `P6_Import_pass1.xml`
3. Import action: **Create New Project**
4. Complete the import wizard

*Imports: Project, WBS, Activities, Resources, ActivityCodes. No relationships.*

### Pass 2 — Add relationships

1. In P6: **File → Import → Primavera P6 XML**
2. Select `P6_Import_pass2.xml`
3. Import action: **Update Existing Project**
4. Complete the import wizard

*Imports: Activities (matched by ObjectId, updated in place) + Relationships.*
*P6 can now resolve all FK constraints because activities are already committed.*

---

## Limitations

This tool uses P6's standard XML import interface. Compared to a direct P6 API
integration, the following limitations apply:

- **Two-pass import required.** P6 18.8 has a transaction-level bug: relationship
  FK constraints fail when activities and relationships are imported in the same
  transaction. Two separate imports are the workaround.
- **No real-time sync.** Batch-only workflow — there is no live connection between
  the spreadsheet and the P6 database.
- **ObjectIds are generated, not stable.** P6 assigns its own database IDs on
  import. Subsequent re-imports may create duplicates rather than updating existing
  records unless ObjectIds are carefully managed.
- **WBS Summary activities cannot be relationship endpoints.** P6 does not allow
  WBS Summary activity types as predecessors or successors. The converter
  automatically drops such relationships.
- **No resource assignments.** Resource definitions are supported but
  ResourceAssignment elements (linking resources to specific activities) are not.
- **Limited field coverage.** Only the fields needed for a standard task-based
  schedule are included. Baselines, earned value overrides, financial periods,
  and other advanced P6 features are not supported.
- **p6_reference.xml is environment-specific.** Calendar/OBS/Role ObjectIds
  differ between P6 databases. Each user must supply a reference export from
  their own environment.

---

## AURORA P6 Integration

For a full API-based integration with live sync, validation, and automated
import — see the AURORA P6 module:

> **[AURORA P6 — link coming soon]**

---

## License

MIT
