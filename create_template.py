"""
create_template.py
Creates a blank P6_Import_Template.xlsx.
Requires p6_reference.xml in the same folder (your own P6 XML export).

  Row 1 = field names
  Row 2 = field types
  Row 3+ = your project data

_Config sheet holds user-configurable P6 object IDs (CalendarObjectId,
OBSObjectId, CurrencyObjectId) - fill these from your P6 environment.

Workflow:
  1. Export any project from P6: File -> Export -> Primavera P6 XML
     Rename the file to p6_reference.xml, place alongside this script.
  2. Run: py create_template.py
     Produces P6_Import_Template.xlsx with sample rows from p6_reference.xml.
  3. Fill P6_Import_Template.xlsx with your project data.
  4. Run: py aurora_p6_converter.py P6_Import_Template.xlsx
     Produces:
       P6_Import_pass1.xml  -- import first  (Create New Project)
       P6_Import_pass2.xml  -- import second (Update Existing Project)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import xml.etree.ElementTree as ET

NS  = "http://xmlns.oracle.com/Primavera/P6Professional/V18.8/API/BusinessObjects"
XSI = "http://www.w3.org/2001/XMLSchema-instance"

# ── Colour scheme ─────────────────────────────────────────────────────────────
HDR_FILL  = PatternFill("solid", fgColor="1F4E79")   # dark blue
TYPE_FILL = PatternFill("solid", fgColor="2E75B6")   # mid blue
HDR_FONT  = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
TYPE_FONT = Font(italic=True, color="FFFFFF", name="Calibri", size=9)
DATA_FONT = Font(name="Calibri", size=10)

def style_row(ws, row_idx, fill, font):
    for cell in ws[row_idx]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)

def auto_width(ws):
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 40)

def write_sheet(wb, name, headers, types, rows):
    ws = wb.create_sheet(name)
    ws.append(headers)
    ws.append(types)
    for row in rows:
        ws.append(row)
    style_row(ws, 1, HDR_FILL,  HDR_FONT)
    style_row(ws, 2, TYPE_FILL, TYPE_FONT)
    for r in range(3, ws.max_row + 1):
        for cell in ws[r]:
            cell.font = DATA_FONT
    ws.freeze_panes = "A3"
    auto_width(ws)
    return ws

# ── Load p6_reference.xml ────────────────────────────────────────────────────
import os, sys
if not os.path.exists("p6_reference.xml"):
    print("ERROR: p6_reference.xml not found.")
    print("  Export any project from P6: File -> Export -> Primavera P6 XML")
    print("  Rename it to p6_reference.xml and place it in this folder.")
    sys.exit(1)
tree = ET.parse("p6_reference.xml")
root = tree.getroot()
proj = root.find(f"{{{NS}}}Project")

def fv(el, tag):
    """Get field value from element."""
    ch = el.find(f"{{{NS}}}{tag}")
    if ch is None:
        return None
    if ch.get(f"{{{XSI}}}nil", "false") == "true":
        return None
    return ch.text

# ── Build UDFType lookup: ObjectId -> Title (for UDFValue sample rows) ────────
udf_oid_to_title = {}
udf_oid_to_datatype = {}
for u in root.findall(f"{{{NS}}}UDFType"):
    oid   = fv(u, "ObjectId")
    title = fv(u, "Title")
    dtype = fv(u, "DataType")
    if oid and title:
        udf_oid_to_title[oid]    = title
        udf_oid_to_datatype[oid] = dtype or ""

# ── PROJECT sheet ─────────────────────────────────────────────────────────────
proj_headers = [
    "ObjectId","Id","Name","Status","OBSObjectId","ParentEPSObjectId",
    "WBSObjectId","PlannedStartDate","ScheduledFinishDate","DataDate",
    "ActivityIdPrefix","ActivityIdSuffix","ActivityIdIncrement",
    "ActivityDefaultCalendarObjectId","ActivityDefaultDurationType",
    "ActivityDefaultActivityType","ActivityDefaultPercentCompleteType",
    "OriginalBudget","LevelingPriority","GUID","AddedBy",
    "CriticalActivityPathType","CriticalActivityFloatLimit",
    "EarnedValueComputeType","EarnedValueETCComputeType",
    "WBSCodeSeparator","SummarizeToWBSLevel","SummaryLevel"
]
proj_types = [
    "ObjectId","String","String","Enum","ObjectId","ObjectId",
    "ObjectId","Date","Date","Date",
    "String","String","Integer",
    "ObjectId","Enum",
    "Enum","Enum",
    "Cost","Integer","String","String",
    "Enum","Duration",
    "Enum","Enum",
    "String","Boolean","Enum"
]
proj_data = [[
    fv(proj,"ObjectId"), fv(proj,"Id"), fv(proj,"Name"), fv(proj,"Status"),
    fv(proj,"OBSObjectId"), fv(proj,"ParentEPSObjectId"),
    fv(proj,"WBSObjectId"),
    fv(proj,"PlannedStartDate"), fv(proj,"ScheduledFinishDate"), fv(proj,"DataDate"),
    fv(proj,"ActivityIdPrefix"), fv(proj,"ActivityIdSuffix"), fv(proj,"ActivityIdIncrement"),
    fv(proj,"ActivityDefaultCalendarObjectId"), fv(proj,"ActivityDefaultDurationType"),
    fv(proj,"ActivityDefaultActivityType"), fv(proj,"ActivityDefaultPercentCompleteType"),
    fv(proj,"OriginalBudget"), fv(proj,"LevelingPriority"),
    fv(proj,"GUID"), fv(proj,"AddedBy"),
    fv(proj,"CriticalActivityPathType"), fv(proj,"CriticalActivityFloatLimit"),
    fv(proj,"EarnedValueComputeType"), fv(proj,"EarnedValueETCComputeType"),
    fv(proj,"WBSCodeSeparator"), fv(proj,"SummarizeToWBSLevel"), fv(proj,"SummaryLevel")
]]

# ── WBS sheet ─────────────────────────────────────────────────────────────────
wbs_headers = [
    "ObjectId","Code","Name","ProjectObjectId","ParentObjectId",
    "OBSObjectId","Status","SequenceNumber","OriginalBudget",
    "EarnedValueComputeType","EarnedValueETCComputeType",
    "EarnedValueETCUserValue","EarnedValueUserPercent",
    "IndependentETCLaborUnits","IndependentETCTotalCost","GUID",
    "UDF_TypeObjectId","UDF_IndicatorValue"
]
wbs_types = [
    "ObjectId","String","String","ObjectId","ObjectId",
    "ObjectId","Enum","Integer","Cost",
    "Enum","Enum",
    "Float","Float",
    "Unit","Cost","String",
    "ObjectId","Enum"
]
wbs_rows = []
for w in proj.findall(f"{{{NS}}}WBS"):
    udf_el = w.find(f"{{{NS}}}UDF")
    udf_type = fv(udf_el, "TypeObjectId") if udf_el is not None else None
    udf_ind  = fv(udf_el, "IndicatorValue") if udf_el is not None else None
    wbs_rows.append([
        fv(w,"ObjectId"), fv(w,"Code"), fv(w,"Name"),
        fv(w,"ProjectObjectId"), fv(w,"ParentObjectId"),
        fv(w,"OBSObjectId"), fv(w,"Status"), fv(w,"SequenceNumber"),
        fv(w,"OriginalBudget"),
        fv(w,"EarnedValueComputeType"), fv(w,"EarnedValueETCComputeType"),
        fv(w,"EarnedValueETCUserValue"), fv(w,"EarnedValueUserPercent"),
        fv(w,"IndependentETCLaborUnits"), fv(w,"IndependentETCTotalCost"),
        fv(w,"GUID"),
        udf_type, udf_ind
    ])

# ── ACTIVITY sheet (5 samples) ────────────────────────────────────────────────
act_headers = [
    "ObjectId","Id","Name","ProjectObjectId","WBSObjectId",
    "Type","Status","CalendarObjectId",
    "PlannedStartDate","PlannedFinishDate","PlannedDuration",
    "ActualStartDate","ActualFinishDate","ActualDuration",
    "RemainingDuration","RemainingEarlyStartDate","RemainingEarlyFinishDate",
    "RemainingLateStartDate","RemainingLateFinishDate",
    "StartDate","FinishDate",
    "PercentComplete","PercentCompleteType","PhysicalPercentComplete",
    "DurationPercentComplete","UnitsPercentComplete",
    "PlannedLaborCost","PlannedLaborUnits","PlannedNonLaborCost","PlannedNonLaborUnits",
    "ActualLaborCost","ActualLaborUnits","ActualNonLaborCost","ActualNonLaborUnits",
    "RemainingLaborCost","RemainingLaborUnits","RemainingNonLaborCost","RemainingNonLaborUnits",
    "AtCompletionDuration","AtCompletionLaborCost","AtCompletionLaborUnits",
    "AtCompletionNonLaborCost","AtCompletionNonLaborUnits","AtCompletionExpenseCost",
    "DurationType","LevelingPriority","AutoComputeActuals",
    "PrimaryResourceObjectId","EstimatedWeight",
    "PrimaryConstraintType","PrimaryConstraintDate",
    "GUID",
    "Code1_TypeObjectId","Code1_ValueObjectId",
    "Code2_TypeObjectId","Code2_ValueObjectId",
    "Code3_TypeObjectId","Code3_ValueObjectId",
    "Code4_TypeObjectId","Code4_ValueObjectId",
]
act_types = [
    "ObjectId","String","String","ObjectId","ObjectId",
    "Enum","Enum","ObjectId",
    "Date","Date","Duration",
    "Date","Date","Duration",
    "Duration","Date","Date",
    "Date","Date",
    "Date","Date",
    "Float","Enum","Float",
    "Float","Float",
    "Cost","Unit","Cost","Unit",
    "Cost","Unit","Cost","Unit",
    "Cost","Unit","Cost","Unit",
    "Duration","Cost","Unit",
    "Cost","Unit","Cost",
    "Enum","Enum","Boolean",
    "ObjectId","Float",
    "Enum","Date",
    "String",
    "ObjectId","ObjectId",
    "ObjectId","ObjectId",
    "ObjectId","ObjectId",
    "ObjectId","ObjectId",
]

act_rows = []
acts = proj.findall(f"{{{NS}}}Activity")[:5]
for a in acts:
    codes = a.findall(f"{{{NS}}}Code")
    def gc(i, sub):
        if i < len(codes):
            ch = codes[i].find(f"{{{NS}}}{sub}")
            return ch.text if ch is not None else None
        return None
    act_rows.append([
        fv(a,"ObjectId"), fv(a,"Id"), fv(a,"Name"),
        fv(a,"ProjectObjectId"), fv(a,"WBSObjectId"),
        fv(a,"Type"), fv(a,"Status"), fv(a,"CalendarObjectId"),
        fv(a,"PlannedStartDate"), fv(a,"PlannedFinishDate"), fv(a,"PlannedDuration"),
        fv(a,"ActualStartDate"), fv(a,"ActualFinishDate"), fv(a,"ActualDuration"),
        fv(a,"RemainingDuration"),
        fv(a,"RemainingEarlyStartDate"), fv(a,"RemainingEarlyFinishDate"),
        fv(a,"RemainingLateStartDate"),  fv(a,"RemainingLateFinishDate"),
        fv(a,"StartDate"), fv(a,"FinishDate"),
        fv(a,"PercentComplete"), fv(a,"PercentCompleteType"), fv(a,"PhysicalPercentComplete"),
        fv(a,"DurationPercentComplete"), fv(a,"UnitsPercentComplete"),
        fv(a,"PlannedLaborCost"), fv(a,"PlannedLaborUnits"),
        fv(a,"PlannedNonLaborCost"), fv(a,"PlannedNonLaborUnits"),
        fv(a,"ActualLaborCost"), fv(a,"ActualLaborUnits"),
        fv(a,"ActualNonLaborCost"), fv(a,"ActualNonLaborUnits"),
        fv(a,"RemainingLaborCost"), fv(a,"RemainingLaborUnits"),
        fv(a,"RemainingNonLaborCost"), fv(a,"RemainingNonLaborUnits"),
        fv(a,"AtCompletionDuration"), fv(a,"AtCompletionLaborCost"), fv(a,"AtCompletionLaborUnits"),
        fv(a,"AtCompletionNonLaborCost"), fv(a,"AtCompletionNonLaborUnits"),
        fv(a,"AtCompletionExpenseCost"),
        fv(a,"DurationType"), fv(a,"LevelingPriority"), fv(a,"AutoComputeActuals"),
        fv(a,"PrimaryResourceObjectId"), fv(a,"EstimatedWeight"),
        fv(a,"PrimaryConstraintType"), fv(a,"PrimaryConstraintDate"),
        fv(a,"GUID"),
        gc(0,"TypeObjectId"), gc(0,"ValueObjectId"),
        gc(1,"TypeObjectId"), gc(1,"ValueObjectId"),
        gc(2,"TypeObjectId"), gc(2,"ValueObjectId"),
        gc(3,"TypeObjectId"), gc(3,"ValueObjectId"),
    ])

# ── RELATIONSHIP sheet ────────────────────────────────────────────────────────
rel_headers = [
    "ObjectId",
    "PredecessorActivityObjectId","PredecessorActivityId","PredecessorProjectObjectId",
    "SuccessorActivityObjectId","SuccessorActivityId","SuccessorProjectObjectId",
    "Type","Lag",
]
rel_types = [
    "ObjectId",
    "ObjectId","String","ObjectId",
    "ObjectId","String","ObjectId",
    "Enum","Duration",
]

sample_acts = proj.findall(f"{{{NS}}}Activity")[:5]
sample_act_oids = set(fv(a,"ObjectId") for a in sample_acts)
oid_to_id = {fv(a,"ObjectId"): fv(a,"Id") for a in sample_acts}

rel_rows = []
for r in proj.findall(f"{{{NS}}}Relationship"):
    pred_oid = fv(r,"PredecessorActivityObjectId")
    succ_oid = fv(r,"SuccessorActivityObjectId")
    if pred_oid in sample_act_oids and succ_oid in sample_act_oids:
        rel_rows.append([
            fv(r,"ObjectId"),
            pred_oid, oid_to_id.get(pred_oid,""), fv(r,"PredecessorProjectObjectId"),
            succ_oid, oid_to_id.get(succ_oid,""), fv(r,"SuccessorProjectObjectId"),
            fv(r,"Type"), fv(r,"Lag"),
        ])

# ── RESOURCEASSIGNMENT sheet ──────────────────────────────────────────────────
ra_headers = [
    "ProjectId","ActivityId","ResourceId","ResourceType","RateType",
    "PlannedUnits","PlannedCost","PlannedUnitsPerTime",
    "PlannedStartDate","PlannedFinishDate",
    "ActualUnits","ActualCost",
    "RemainingUnits","RemainingCost","RemainingDuration",
    "IsPrimaryResource","Proficiency",
    "ObjectId","ActivityObjectId","ResourceObjectId","ProjectObjectId","WBSObjectId",
]
ra_types = [
    "Lookup","Lookup","Lookup","Enum","Enum",
    "Unit","Cost","Float",
    "Date","Date",
    "Unit","Cost",
    "Unit","Cost","Duration",
    "Boolean","Enum",
    "ObjectId","ObjectId","ObjectId","ObjectId","ObjectId",
]

# Only include sample RAs whose ActivityObjectId is in the 5-activity sample set,
# so the sample data is self-consistent and importable.
ra_sample_rows = []
for ra in proj.findall(f"{{{NS}}}ResourceAssignment"):
    act_oid = fv(ra, "ActivityObjectId")
    if act_oid not in sample_act_oids:
        continue
    ra_sample_rows.append([
        fv(ra,"ProjectId"),
        fv(ra,"ActivityId"),        fv(ra,"ResourceId"),
        fv(ra,"ResourceType"),      fv(ra,"RateType"),
        fv(ra,"PlannedUnits"),      fv(ra,"PlannedCost"),
        fv(ra,"PlannedUnitsPerTime"),
        fv(ra,"PlannedStartDate"),  fv(ra,"PlannedFinishDate"),
        fv(ra,"ActualUnits"),       fv(ra,"ActualCost"),
        fv(ra,"RemainingUnits"),    fv(ra,"RemainingCost"),
        fv(ra,"RemainingDuration"),
        fv(ra,"IsPrimaryResource"), fv(ra,"Proficiency"),
        fv(ra,"ObjectId"),          act_oid,
        fv(ra,"ResourceObjectId"),  fv(ra,"ProjectObjectId"),
        fv(ra,"WBSObjectId"),
    ])
    if len(ra_sample_rows) >= 3:
        break

# ── UDFTYPE sheet ─────────────────────────────────────────────────────────────
udftype_headers = ["ObjectId","DataType","SubjectArea","Title","IsSecureCode"]
udftype_types   = ["ObjectId","Enum","Enum","String","Boolean"]
udftype_rows = []
for u in root.findall(f"{{{NS}}}UDFType"):
    udftype_rows.append([
        fv(u,"ObjectId"), fv(u,"DataType"), fv(u,"SubjectArea"),
        fv(u,"Title"), fv(u,"IsSecureCode")
    ])

# ── UDFVALUE sheet ────────────────────────────────────────────────────────────
udfv_headers = [
    "ProjectId","ObjectType","ObjectId_Ref","UDFTypeTitle",
    "TextValue","NumberValue","DateValue","IndicatorValue",
]
udfv_types = [
    "Lookup","Enum","Lookup","Lookup",
    "String","Cost","Date","Enum",
]
udfv_rows = []
proj_id_ref = fv(proj, "Id")

# Sample from Activity UDFs only (up to 3 activities, first UDF each).
# WBS UDFs are already covered by the UDF_TypeObjectId / UDF_IndicatorValue
# columns in the WBS sheet — putting them here too would create duplicates.
for a in proj.findall(f"{{{NS}}}Activity")[:3]:
    udf_el = a.find(f"{{{NS}}}UDF")
    if udf_el is not None:
        type_oid = fv(udf_el, "TypeObjectId")
        title = udf_oid_to_title.get(type_oid, "") if type_oid else ""
        if title:
            udfv_rows.append([
                proj_id_ref, "Activity", fv(a, "Id"), title,
                fv(udf_el, "TextValue"),
                fv(udf_el, "NumberValue"),
                fv(udf_el, "DateValue"),
                fv(udf_el, "IndicatorValue"),
            ])

# ── ACTIVITYCODETYPE sheet ────────────────────────────────────────────────────
actype_headers = [
    "ObjectId","Name","Scope","Length","IsSecureCode","SequenceNumber","RefProjectObjectIds"
]
actype_types = [
    "ObjectId","String","Enum","Integer","Boolean","Integer","String"
]
actype_rows = []
for act in root.findall(f"{{{NS}}}ActivityCodeType"):
    actype_rows.append([
        fv(act,"ObjectId"), fv(act,"Name"), fv(act,"Scope"),
        fv(act,"Length"), fv(act,"IsSecureCode"), fv(act,"SequenceNumber"),
        fv(act,"RefProjectObjectIds")
    ])

# ── ACTIVITYCODE sheet ────────────────────────────────────────────────────────
ac_headers = [
    "ObjectId","CodeTypeObjectId","CodeValue","Description","Color","SequenceNumber"
]
ac_types = [
    "ObjectId","ObjectId","String","String","String","Integer"
]
ac_rows = []
for ac in root.findall(f"{{{NS}}}ActivityCode"):
    ac_rows.append([
        fv(ac,"ObjectId"), fv(ac,"CodeTypeObjectId"), fv(ac,"CodeValue"),
        fv(ac,"Description"), fv(ac,"Color"), fv(ac,"SequenceNumber")
    ])

# ── RESOURCE sheet ────────────────────────────────────────────────────────────
res_headers = [
    "ObjectId","Id","Name","Code","ResourceType","CalendarObjectId",
    "CurrencyObjectId","DefaultUnitsPerTime","OvertimeFactor",
    "IsActive","IsOverTimeAllowed","AutoComputeActuals",
    "CalculateCostFromUnits","SequenceNumber","GUID",
    "ParentObjectId","PrimaryRoleObjectId","ResourceNotes","UnitOfMeasureObjectId"
]
res_types = [
    "ObjectId","String","String","String","Enum","ObjectId",
    "ObjectId","Unit","Float",
    "Boolean","Boolean","Boolean",
    "Boolean","Integer","String",
    "ObjectId","ObjectId","String","ObjectId"
]
res_rows = []
for r in root.findall(f"{{{NS}}}Resource"):
    res_rows.append([
        fv(r,"ObjectId"), fv(r,"Id"), fv(r,"Name"), fv(r,"Code"),
        fv(r,"ResourceType"), fv(r,"CalendarObjectId"),
        fv(r,"CurrencyObjectId"), fv(r,"DefaultUnitsPerTime"), fv(r,"OvertimeFactor"),
        fv(r,"IsActive"), fv(r,"IsOverTimeAllowed"), fv(r,"AutoComputeActuals"),
        fv(r,"CalculateCostFromUnits"), fv(r,"SequenceNumber"), fv(r,"GUID"),
        fv(r,"ParentObjectId"), fv(r,"PrimaryRoleObjectId"),
        fv(r,"ResourceNotes"), fv(r,"UnitOfMeasureObjectId")
    ])

# ── EXPENSE sheet (empty placeholder) ────────────────────────────────────────
exp_headers = [
    "ObjectId","ActivityObjectId","ProjectObjectId","Description",
    "CostAccountObjectId","PlannedCost","ActualCost","RemainingCost",
    "AtCompletionCost","AutoComputeActuals","AccrualType","VendorId"
]
exp_types = [
    "ObjectId","ObjectId","ObjectId","String",
    "ObjectId","Cost","Cost","Cost",
    "Cost","Boolean","Enum","String"
]

# ── _CONFIG sheet ─────────────────────────────────────────────────────────────
cfg_headers = ["Field", "Value", "Description"]
cfg_types   = ["String", "ObjectId", "String"]
cfg_rows = [
    ["CalendarObjectId",  fv(proj, "ActivityDefaultCalendarObjectId"),
     "ObjectId of the global Calendar used by activities (default: 939 = Corporate Full Time)"],
    ["OBSObjectId",       fv(proj, "OBSObjectId"),
     "ObjectId of the OBS node assigned to this project and WBS (default: 745 = E&C)"],
    ["CurrencyObjectId",  "1",
     "ObjectId of the Currency (default: 1 = USD Dollar)"],
]

# ── Write workbook ─────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
wb.remove(wb.active)  # remove default sheet

write_sheet(wb, "_Config",            cfg_headers,      cfg_types,      cfg_rows)
write_sheet(wb, "Project",            proj_headers,     proj_types,     proj_data)
write_sheet(wb, "WBS",                wbs_headers,      wbs_types,      wbs_rows)
write_sheet(wb, "Activity",           act_headers,      act_types,      act_rows)
write_sheet(wb, "Relationship",       rel_headers,      rel_types,      rel_rows)
write_sheet(wb, "ResourceAssignment", ra_headers,       ra_types,       ra_sample_rows)
write_sheet(wb, "UDFType",            udftype_headers,  udftype_types,  udftype_rows)
write_sheet(wb, "UDFValue",           udfv_headers,     udfv_types,     udfv_rows)
write_sheet(wb, "ActivityCodeType",   actype_headers,   actype_types,   actype_rows)
write_sheet(wb, "ActivityCode",       ac_headers,       ac_types,       ac_rows)
write_sheet(wb, "Expense",            exp_headers,      exp_types,      [])
write_sheet(wb, "Resource",           res_headers,      res_types,      res_rows)

out = "P6_Import_Template.xlsx"
wb.save(out)
print(f"Saved {out}")
print(f"  _Config rows         : {len(cfg_rows)}")
print(f"  Project rows         : {len(proj_data)}")
print(f"  WBS rows             : {len(wbs_rows)}")
print(f"  Activity rows        : {len(act_rows)}  (5 samples from p6_reference.xml)")
print(f"  Relationship         : {len(rel_rows)}")
print(f"  ResourceAssignment   : {len(ra_sample_rows)}")
print(f"  UDFType              : {len(udftype_rows)}")
print(f"  UDFValue             : {len(udfv_rows)}")
print(f"  ActivityCodeType     : {len(actype_rows)}")
print(f"  ActivityCode         : {len(ac_rows)}")
print(f"  Resource rows        : {len(res_rows)}")
