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
- **EC00630.xml** — reference P6 export file that must be placed in the same
  directory as the converter. It provides the Calendar, OBS, Currency, Role, and
  RoleRate elements that P6 CleanupActivities requires in every import XML.
  Without it the import crashes with a NullReferenceException.

---

## Files

| File | Purpose |
|---|---|
| `aurora_p6_converter.py` | Main converter — reads template, writes pass1 + pass2 XML |
| `create_template.py` | Generates a blank `P6_Import_Template.xlsx` with sample EC00630 data |
| `P6_Import_Template.xlsx` | Fill this with your project data |
| `EC00630.xml` | Reference P6 export (required at runtime, not included in repo) |

---

## How to Use

### Step 1 — Fill the template

Open `P6_Import_Template.xlsx` and populate the following sheets:

| Sheet | What to fill |
|---|---|
| `_Config` | CalendarObjectId, OBSObjectId, CurrencyObjectId from your P6 environment |
| `Project` | One row: project Id, Name, dates, EPS parent, WBS root |
| `WBS` | One row per WBS node |
| `Activity` | One row per activity (Type, dates, status, percent complete) |
| `Relationship` | One row per dependency (predecessor, successor, type, lag) |
| `Resource` | Optional — one row per resource |
| `ActivityCodeType` / `ActivityCode` | Optional — activity code definitions |

Row 1 = column names, Row 2 = data types (do not edit), Row 3+ = your data.

To regenerate a blank template with sample data from EC00630.xml:
```
py create_template.py
```

### Step 2 — Run the converter

```
py aurora_p6_converter.py P6_Import_Template.xlsx
```

This writes two files to the same directory as the template:
- `P6_Import_pass1.xml`
- `P6_Import_pass2.xml`

---

## Import Instructions (P6 Professional 18.8)

> **Important:** Both passes must be imported in order. Do not skip pass 1.

### Pass 1 — Create the project

1. In P6, go to **File → Import**
2. Select **Primavera P6 XML**
3. Choose `P6_Import_pass1.xml`
4. Import action: **Create New Project**
5. Complete the import wizard

*Pass 1 imports: Project, WBS, Activities, Resources, ActivityCodes.*
*No relationships are included.*

### Pass 2 — Add relationships

1. In P6, go to **File → Import**
2. Select **Primavera P6 XML**
3. Choose `P6_Import_pass2.xml`
4. Import action: **Update Existing Project**
5. Complete the import wizard

*Pass 2 imports: Activities (matched by ObjectId, updated in place) and*
*Relationships. P6 can now resolve all FK constraints because activities are*
*already committed to the database from pass 1.*

---

## Limitations

This tool uses P6's standard XML import interface. Compared to a direct P6 API
integration, the following limitations apply:

- **Two-pass import required.** P6 18.8 has a transaction-level bug where
  relationship FK constraints fail for new-project imports. The workaround
  (pass 1 then pass 2) is manual and requires two separate import operations.
- **No real-time sync.** The workflow is batch-only: export from Excel, import
  into P6. There is no live connection between the spreadsheet and the P6 database.
- **ObjectIds are generated, not stable.** P6 assigns its own database IDs on
  import. Subsequent re-imports may create duplicates rather than updating
  existing records unless ObjectIds are carefully managed.
- **WBS Summary activities cannot be relationship endpoints.** P6 does not
  allow WBS Summary activity types as predecessors or successors. Such
  relationships are automatically dropped by the converter.
- **No resource assignments.** The current template supports resource
  definitions but not ResourceAssignment elements (linking resources to specific
  activities).
- **Limited field coverage.** Only the fields needed for a standard task-based
  schedule are supported. Baselines, earned value overrides, financial periods,
  and other advanced P6 features are not included.
- **EC00630.xml dependency.** The converter requires a specific reference export
  file from your P6 environment. A different P6 database will have different
  Calendar/OBS/Role ObjectIds and will need a new reference export.

---

## AURORA P6 Integration

For a full API-based integration with live sync, validation, and automated
import — see the AURORA P6 module:

> **[AURORA P6 — link coming soon]**

---

## License

MIT
