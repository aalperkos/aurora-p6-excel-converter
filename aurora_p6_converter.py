"""
aurora_p6_converter.py
Reads P6_Import_Template.xlsx and produces TWO P6 18.8 XML files for two-pass import.

P6 18.8 cannot resolve relationship FK constraints within a new-project import
transaction (activities are not yet committed when relationships are processed).
The two-pass workaround:

  Pass 1 - P6_Import_pass1.xml   Import action: Create New Project
    Contains: Calendar/OBS/Currency/Role/RoleRate (required global refs)
              Resource, ActivityCodeType, ActivityCode (top-level)
              Project -> WBS, Activity  (NO Relationships, NO ResourceAssignments)

  Pass 2 - P6_Import_pass2.xml   Import action: Update Existing Project
    Contains: Calendar/OBS/Currency/Role/RoleRate
              Project -> Activity (same ObjectIds as pass 1), Relationship
              (NO WBS, NO ResourceAssignments)

  Pass 3 - P6_Import_pass3.xml   Import action: Update Existing Project
    Contains: Calendar/OBS/Currency/Role/RoleRate
              Project -> Activity (same ObjectIds), ResourceAssignment
              (NO WBS, NO Relationships)

  P6 18.8 FK constraint behaviour:
    fk_taskpred_task / fk_taskpred_task_2  (TASKPRED) - relationships need
      activities already committed -> pass 2 after pass 1
    fk_taskactv_task (TASKACTV) - resource assignments need activities
      already committed -> separate pass 3

Sheet layout:
  Row 1: field names
  Row 2: field types (Date/String/Enum/Cost/Unit/Duration/Boolean/Lookup/ObjectId)
  Row 3+: data rows

Requires p6_reference.xml:
  Before running, export ANY project from your P6 as XML
  (File -> Export -> Primavera P6 XML), rename it to p6_reference.xml
  and place it in the same folder as this script.
  The converter reads Calendar, OBS, Currency, UDFType, Role, and RoleRate
  from that file - elements P6 requires in every import XML.

Usage:
  py aurora_p6_converter.py [template.xlsx]
  Output files P6_Import_pass1.xml and P6_Import_pass2.xml are written
  to the same directory as the template.
"""

import sys
import os
import copy
import uuid
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime
import openpyxl

# ── Namespace constants ──────────────────────────────────────────────────────
NS_BO  = "http://xmlns.oracle.com/Primavera/P6Professional/V18.8/API/BusinessObjects"
NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"

# ── Default P6 ObjectId fallbacks (overridden by _Config sheet) ──────────────
DEFAULT_CALENDAR_OID = ""   # set in _Config: CalendarObjectId
DEFAULT_OBS_OID      = ""   # set in _Config: OBSObjectId
DEFAULT_ROOT_WBS_OID = "17583"
SCHEMA_LOC = (
    "http://xmlns.oracle.com/Primavera/P6Professional/V18.8/API/BusinessObjects "
    "http://xmlns.oracle.com/Primavera/P6Professional/V18.8/API/p6apibo.xsd"
)

# ── Fields that are always xsi:nil (Activity) ────────────────────────────────
ACTIVITY_NIL_ALWAYS = {
    "ExpectedFinishDate",
    "ExternalEarlyStartDate", "ExternalLateFinishDate",
    "ResumeDate",
    "SecondaryConstraintDate", "SecondaryConstraintType", "SuspendDate",
}
# Optionally nil: emit value when present, xsi:nil when blank
ACTIVITY_NIL_OPTIONAL = {
    "ActualFinishDate", "ActualStartDate",   # nil for Not Started; value for In Progress
    "PrimaryConstraintDate", "PrimaryConstraintType",
    "PrimaryResourceObjectId", "WBSObjectId",
}

WBS_NIL_ALWAYS   = {"AnticipatedFinishDate", "AnticipatedStartDate", "WBSCategoryObjectId"}
WBS_NIL_OPTIONAL = {"ParentObjectId"}

RA_NIL_ALWAYS   = {"ActualCurve", "ActualFinishDate", "ActualStartDate",
                    "PlannedCurve", "RemainingCurve", "ResourceCurveObjectId"}
RA_NIL_OPTIONAL = {"CostAccountObjectId", "RoleObjectId"}


# ── Helpers ──────────────────────────────────────────────────────────────────
def new_guid():
    return "{" + str(uuid.uuid4()).upper() + "}"

def fmt_date(val):
    """Normalise a date value to ISO 8601 datetime string used by P6."""
    if val is None:
        return None
    if isinstance(val, (datetime,)):
        return val.strftime("%Y-%m-%dT%H:%M:%S")
    s = str(val).strip()
    if not s:
        return None
    for pat in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
                "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, pat).strftime("%Y-%m-%dT%H:%M:%S")
        except ValueError:
            pass
    return s  # pass through as-is

def make_tag(parent, tag, text=None, nil=False):
    """Add a child element, with optional xsi:nil."""
    el = ET.SubElement(parent, tag)
    if nil:
        el.set(f"{{{NS_XSI}}}nil", "true")
    elif text is not None:
        el.text = str(text)
    return el

def nil_or_text(parent, tag, value, nil_always_set, nil_optional_set):
    """Emit element with correct nil/value handling."""
    if tag in nil_always_set:
        make_tag(parent, tag, nil=True)
    elif tag in nil_optional_set:
        if value is None or str(value).strip() == "":
            make_tag(parent, tag, nil=True)
        else:
            make_tag(parent, tag, text=value)
    else:
        make_tag(parent, tag, text=value if value is not None else "")

def read_config(wb):
    """Read _Config sheet as a {Field: Value} dict."""
    if "_Config" not in wb.sheetnames:
        return {}
    ws = wb["_Config"]
    rows = list(ws.iter_rows(values_only=True))
    cfg = {}
    for row in rows[2:]:
        if row and row[0] is not None and row[1] is not None:
            cfg[str(row[0]).strip()] = str(row[1]).strip()
    return cfg


def load_reference_elements(ref_path, *tag_names):
    """
    Load named top-level elements verbatim from a P6 reference XML file.
    Returns {tag_name: [list of ET.Element copies]}.

    Aborts with a clear error if the file is missing.
    """
    if not os.path.exists(ref_path):
        print()
        print("ERROR: p6_reference.xml not found.")
        print()
        print("  Before running the converter, export any project from your P6:")
        print("    File -> Export -> Primavera P6 XML")
        print(f"  Rename the exported file to:  p6_reference.xml")
        print(f"  Place it in the same folder:  {os.path.dirname(ref_path)}")
        print()
        print("  The converter reads Calendar, OBS, Currency, UDFType, Role, and")
        print("  RoleRate from that file. P6 requires these in every import XML.")
        sys.exit(1)
    try:
        ref_tree = ET.parse(ref_path)
        ref_root = ref_tree.getroot()
        result = {t: [] for t in tag_names}
        for tname in tag_names:
            result[tname] = [copy.deepcopy(el)
                             for el in ref_root.findall(f"{{{NS_BO}}}{tname}")]
        counts = {t: len(v) for t, v in result.items()}
        print(f"  Loaded from {os.path.basename(ref_path)}: {counts}")
        return result
    except Exception as e:
        print(f"ERROR: Could not parse {ref_path}: {e}")
        sys.exit(1)


def read_sheet(wb, name):
    """Return list of dicts from a sheet (row1=headers, row2=types, row3+=data)."""
    if name not in wb.sheetnames:
        return []
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 3:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = []
    for row in rows[2:]:
        if all(v is None for v in row):
            continue
        rec = {}
        for i, h in enumerate(headers):
            if h:
                v = row[i] if i < len(row) else None
                rec[h] = v
        data.append(rec)
    return data

def cell(rec, *keys):
    """Return first non-None value from record for any of the given keys."""
    for k in keys:
        v = rec.get(k)
        if v is not None and str(v).strip() != "":
            return str(v).strip()
    return None


# ── UDFValue injection helper ─────────────────────────────────────────────────
def _inject_udf(parent_el, udf_v, udf_type_lookup):
    """
    Append a <UDF> child to parent_el from a UDFValue sheet row.
    udf_type_lookup: {Title: ObjectId}
    """
    title = cell(udf_v, "UDFTypeTitle")
    if not title:
        return
    type_oid = (udf_type_lookup or {}).get(title)
    if not type_oid:
        print(f"  [WARN] UDFValue: UDFTypeTitle '{title}' not found in UDFType data — skipped")
        return
    udf_el = ET.SubElement(parent_el, "UDF")
    make_tag(udf_el, "TypeObjectId", text=type_oid)
    text_v = cell(udf_v, "TextValue")
    num_v  = cell(udf_v, "NumberValue")
    date_v = cell(udf_v, "DateValue")
    ind_v  = cell(udf_v, "IndicatorValue")
    if text_v:
        make_tag(udf_el, "TextValue",      text=text_v)
    elif num_v:
        make_tag(udf_el, "NumberValue",    text=num_v)
    elif date_v:
        make_tag(udf_el, "DateValue",      text=fmt_date(date_v))
    elif ind_v:
        make_tag(udf_el, "IndicatorValue", text=ind_v)


# ── ResourceAssignment builder ────────────────────────────────────────────────
def build_resource_assignment_element(ra_data, act_oid_to_id, proj_oid):
    """Build a <ResourceAssignment> element in P6 field order."""
    ra_el = ET.Element("ResourceAssignment")

    make_tag(ra_el, "ActualCost",            text=cell(ra_data, "ActualCost") or "0")
    make_tag(ra_el, "ActualCurve",           nil=True)
    make_tag(ra_el, "ActualFinishDate",      nil=True)
    make_tag(ra_el, "ActualOvertimeCost",    text="0")
    make_tag(ra_el, "ActualOvertimeUnits",   text="0.000000")
    make_tag(ra_el, "ActualRegularCost",     text="0")
    make_tag(ra_el, "ActualRegularUnits",    text="0.000000")
    make_tag(ra_el, "ActualStartDate",       nil=True)
    make_tag(ra_el, "ActualThisPeriodCost",  text="0")
    make_tag(ra_el, "ActualThisPeriodUnits", text="0.000000")
    make_tag(ra_el, "ActualUnits",           text=cell(ra_data, "ActualUnits") or "0.000000")

    act_oid = cell(ra_data, "ActivityObjectId")
    act_id  = cell(ra_data, "ActivityId") or (act_oid_to_id.get(act_oid, "") if act_oid else "")
    make_tag(ra_el, "ActivityId",            text=act_id)
    make_tag(ra_el, "ActivityObjectId",      text=act_oid or "")
    make_tag(ra_el, "AutoComputeActuals",    text="1")

    cost_acct = cell(ra_data, "CostAccountObjectId")
    if cost_acct:
        make_tag(ra_el, "CostAccountObjectId", text=cost_acct)
    else:
        make_tag(ra_el, "CostAccountObjectId", nil=True)

    make_tag(ra_el, "GUID",                  text=new_guid())
    make_tag(ra_el, "IsPrimaryResource",     text=cell(ra_data, "IsPrimaryResource") or "0")
    make_tag(ra_el, "ObjectId",              text=cell(ra_data, "ObjectId") or "")
    make_tag(ra_el, "OverBudgetCost",        text="0")
    make_tag(ra_el, "OverBudgetUnits",       text="0.000000")
    make_tag(ra_el, "OvertimeFactor",        text="0")
    make_tag(ra_el, "PlannedCost",           text=cell(ra_data, "PlannedCost") or "0")
    make_tag(ra_el, "PlannedCurve",          nil=True)
    make_tag(ra_el, "PlannedUnits",          text=cell(ra_data, "PlannedUnits") or "0.000000")
    make_tag(ra_el, "ProjectObjectId",       text=cell(ra_data, "ProjectObjectId") or proj_oid or "")
    make_tag(ra_el, "RateSource",            text="Resource")
    make_tag(ra_el, "RateType",              text=cell(ra_data, "RateType") or "Price / Unit")
    make_tag(ra_el, "RemainingCost",         text=cell(ra_data, "RemainingCost") or "0")
    make_tag(ra_el, "RemainingCurve",        nil=True)
    make_tag(ra_el, "RemainingUnits",        text=cell(ra_data, "RemainingUnits") or "0.000000")
    make_tag(ra_el, "ResourceCurveObjectId", nil=True)

    res_id  = cell(ra_data, "ResourceId")
    res_oid = cell(ra_data, "ResourceObjectId")
    make_tag(ra_el, "ResourceId",            text=res_id or "")
    make_tag(ra_el, "ResourceObjectId",      text=res_oid or "")
    make_tag(ra_el, "RoleObjectId",          nil=True)

    return ra_el


# ── XML builders ─────────────────────────────────────────────────────────────
def build_project_element(proj_data, wbs_list, activity_list,
                          relationship_list, ra_list,
                          udf_value_list=None, udf_type_lookup=None):
    """Build the <Project> element with all nested children."""
    proj = ET.Element("Project")

    p = proj_data

    # ── Scalar project fields in EC00630 order ──
    def ps(tag, key=None, val=None, nil=False, date=False, dft=None):
        v = val if val is not None else cell(p, key or tag)
        if v is None:
            v = dft
        if nil:
            make_tag(proj, tag, nil=True)
        else:
            make_tag(proj, tag, text=(fmt_date(v) if date else v) if v is not None else "")

    ps("ActivityDefaultActivityType",    dft="Task Dependent")
    ps("ActivityDefaultCalendarObjectId")
    make_tag(proj, "ActivityDefaultCostAccountObjectId", nil=True)
    ps("ActivityDefaultDurationType",    dft="Fixed Duration and Units/Time")
    ps("ActivityDefaultPercentCompleteType", dft="Duration")
    ps("ActivityDefaultPricePerUnit",    dft="100.00000000")
    ps("ActivityIdBasedOnSelectedActivity", dft="1")
    ps("ActivityIdIncrement",            dft="10")
    ps("ActivityIdPrefix",               dft="EC")
    ps("ActivityIdSuffix",               dft="1000")
    ps("ActivityPercentCompleteBasedOnActivitySteps", dft="1")
    ps("AddActualToRemaining",           dft="0")
    ps("AddedBy",                        dft="admin")
    ps("AllowNegativeActualUnitsFlag",   dft="0")
    ps("AnnualDiscountRate",             dft="5.000000")
    make_tag(proj, "AnticipatedFinishDate", nil=True)
    make_tag(proj, "AnticipatedStartDate",  nil=True)
    ps("AssignmentDefaultDrivingFlag",   dft="0")
    ps("AssignmentDefaultRateType",      dft="Price / Unit")
    ps("CheckOutStatus",                 dft="0")
    ps("CostQuantityRecalculateFlag",    dft="0")
    ps("CriticalActivityFloatLimit",     dft="0.00")
    ps("CriticalActivityPathType",       dft="Critical Float")
    make_tag(proj, "CurrentBaselineProjectObjectId", nil=True)
    ps("DataDate",                       date=True, dft="2014-01-01T00:00:00")
    ps("DateAdded",                      date=True, dft=datetime.now().strftime("%Y-%m-%dT%H:%M:%S"))
    ps("DefaultPriceTimeUnits",          dft="Hour")
    ps("DiscountApplicationPeriod",      dft="Month")
    ps("EarnedValueComputeType",         dft="Activity Percent Complete")
    ps("EarnedValueETCComputeType",      dft="PF = 1 / CPI")
    ps("EarnedValueETCUserValue",        dft="0")
    ps("EarnedValueUserPercent",         dft="0.0")
    ps("EnableSummarization",            dft="1")
    ps("FiscalYearStartMonth",           dft="1")
    guid = cell(p, "GUID") or new_guid()
    make_tag(proj, "GUID", text=guid)
    ps("Id")
    ps("IndependentETCLaborUnits",       dft="0.000000")
    ps("IndependentETCTotalCost",        dft="0.000000")
    make_tag(proj, "LastFinancialPeriodObjectId", nil=True)
    ps("LevelingPriority",               dft="10")
    ps("LinkActualToActualThisPeriod",   dft="1")
    ps("LinkPercentCompleteWithActual",  dft="0")
    ps("LinkPlannedAndAtCompletionFlag", dft="1")
    make_tag(proj, "MustFinishByDate", nil=True)
    ps("Name")
    ps("OBSObjectId",                    dft=DEFAULT_OBS_OID)
    ps("ObjectId")
    ps("OriginalBudget",                 dft="0.000000")
    ps("ParentEPSObjectId")
    ps("PlannedStartDate",               date=True)
    ps("PrimaryResourcesCanMarkActivitiesAsCompleted", dft="1")
    make_tag(proj, "ProjectForecastStartDate", nil=True)
    ps("ResetPlannedToRemainingFlag",    dft="0")
    ps("ResourceCanBeAssignedToSameActivityMoreThanOnce", dft="1")
    ps("ResourcesCanAssignThemselvesToActivities",        dft="1")
    ps("ScheduledFinishDate",            date=True)
    ps("Status",                         dft="Active")
    ps("StrategicPriority",              dft="100")
    ps("SummarizeToWBSLevel",            dft="0")
    ps("SummaryLevel",                   dft="Assignment Level")
    ps("UseProjectBaselineForEarnedValue", dft="1")
    ps("WBSCodeSeparator",               dft=".")
    ps("WBSObjectId")
    make_tag(proj, "WebSiteRootDirectory", nil=True)
    make_tag(proj, "WebSiteURL",           nil=True)

    # ── WBS ──────────────────────────────────────────────────────────────────
    for w in wbs_list:
        wbs_el = ET.SubElement(proj, "WBS")
        make_tag(wbs_el, "AnticipatedFinishDate",  nil=True)
        make_tag(wbs_el, "AnticipatedStartDate",   nil=True)
        make_tag(wbs_el, "Code",                   text=cell(w, "Code") or "1")
        make_tag(wbs_el, "EarnedValueComputeType", text=cell(w, "EarnedValueComputeType") or "Activity Percent Complete")
        make_tag(wbs_el, "EarnedValueETCComputeType", text=cell(w, "EarnedValueETCComputeType") or "PF = 1 / CPI")
        make_tag(wbs_el, "EarnedValueETCUserValue", text=cell(w, "EarnedValueETCUserValue") or "0")
        make_tag(wbs_el, "EarnedValueUserPercent",  text=cell(w, "EarnedValueUserPercent") or "0.0")
        make_tag(wbs_el, "GUID",  text=cell(w, "GUID") or new_guid())
        make_tag(wbs_el, "IndependentETCLaborUnits",  text=cell(w, "IndependentETCLaborUnits") or "0.000000")
        make_tag(wbs_el, "IndependentETCTotalCost",   text=cell(w, "IndependentETCTotalCost")  or "0.000000")
        make_tag(wbs_el, "Name",           text=cell(w, "Name") or "")
        make_tag(wbs_el, "OBSObjectId",    text=cell(w, "OBSObjectId") or DEFAULT_OBS_OID)
        make_tag(wbs_el, "ObjectId",       text=cell(w, "ObjectId") or "")
        make_tag(wbs_el, "OriginalBudget", text=cell(w, "OriginalBudget") or "0.000000")
        parent_oid = cell(w, "ParentObjectId")
        if parent_oid:
            make_tag(wbs_el, "ParentObjectId", text=parent_oid)
        else:
            make_tag(wbs_el, "ParentObjectId", nil=True)
        make_tag(wbs_el, "ProjectObjectId", text=cell(w, "ProjectObjectId") or cell(p, "ObjectId") or "")
        make_tag(wbs_el, "SequenceNumber",  text=cell(w, "SequenceNumber") or "0")
        make_tag(wbs_el, "Status",          text=cell(w, "Status") or "Active")
        make_tag(wbs_el, "WBSCategoryObjectId", nil=True)

        # UDF from WBS sheet columns (UDF_TypeObjectId / UDF_IndicatorValue)
        udf_type_oid = cell(w, "UDF_TypeObjectId")
        udf_ind_val  = cell(w, "UDF_IndicatorValue")
        if udf_type_oid:
            udf_el = ET.SubElement(wbs_el, "UDF")
            make_tag(udf_el, "TypeObjectId",   text=udf_type_oid)
            make_tag(udf_el, "IndicatorValue", text=udf_ind_val or "Green")

        # UDF from UDFValue sheet — matched by WBS Code
        wbs_code = cell(w, "Code")
        for uv in (udf_value_list or []):
            if cell(uv, "ObjectType") == "WBS" and cell(uv, "ObjectId_Ref") == wbs_code:
                _inject_udf(wbs_el, uv, udf_type_lookup)

    # ── Build activity lookup maps (used by both RA and Relationship sections) ──
    act_oid_to_id = {}
    act_oid_set   = set()
    for a in activity_list:
        oid = cell(a, "ObjectId")
        aid = cell(a, "Id")
        if oid:
            act_oid_set.add(oid)
            if aid:
                act_oid_to_id[oid] = aid

    # ── Build set of WBS ObjectIds for validation ─────────────────────────────
    wbs_oid_set = set()
    first_top_wbs_oid = None
    for w in wbs_list:
        oid = cell(w, "ObjectId")
        if oid:
            wbs_oid_set.add(oid)
            if first_top_wbs_oid is None and not cell(w, "ParentObjectId"):
                first_top_wbs_oid = oid
    if first_top_wbs_oid is None and wbs_oid_set:
        first_top_wbs_oid = sorted(wbs_oid_set)[0]

    # ── Activities ────────────────────────────────────────────────────────────
    for a in activity_list:
        act_el = ET.SubElement(proj, "Activity")

        def af(tag, key=None, date=False, default=None):
            v = cell(a, key or tag) or default
            if tag in ACTIVITY_NIL_ALWAYS:
                make_tag(act_el, tag, nil=True)
            elif tag in ACTIVITY_NIL_OPTIONAL:
                if v:
                    make_tag(act_el, tag, text=fmt_date(v) if date else v)
                else:
                    make_tag(act_el, tag, nil=True)
            else:
                make_tag(act_el, tag, text=(fmt_date(v) if date else v) if v is not None else "")

        af("ActualDuration",              default="0")
        af("ActualFinishDate",            date=True)
        af("ActualLaborCost",             default="0")
        af("ActualLaborUnits",            default="0.000000")
        af("ActualNonLaborCost",          default="0")
        af("ActualNonLaborUnits",         default="0.000000")
        af("ActualStartDate",             date=True)
        af("ActualThisPeriodLaborCost",   default="0")
        af("ActualThisPeriodLaborUnits",  default="0.000000")
        af("ActualThisPeriodNonLaborCost",  default="0")
        af("ActualThisPeriodNonLaborUnits", default="0.000000")
        af("AtCompletionDuration",        default="0")
        af("AtCompletionExpenseCost",     default="0")
        af("AtCompletionLaborCost",       default="0")
        af("AtCompletionLaborUnits",      default="0.000000")
        af("AtCompletionNonLaborCost",    default="0")
        af("AtCompletionNonLaborUnits",   default="0.000000")
        af("AutoComputeActuals",          default="1")
        af("CalendarObjectId",            default=DEFAULT_CALENDAR_OID)
        af("DurationPercentComplete",     default="0")
        af("DurationType",                default="Fixed Duration and Units/Time")
        af("EstimatedWeight",             default="1.00")
        af("ExpectedFinishDate",          date=True)
        af("ExternalEarlyStartDate",      date=True)
        af("ExternalLateFinishDate",      date=True)
        af("Feedback",                    default="")
        af("FinishDate",                  date=True)
        guid_a = cell(a, "GUID") or new_guid()
        make_tag(act_el, "GUID", text=guid_a)
        af("Id")
        af("IsNewFeedback",               default="0")
        af("LevelingPriority",            default="Normal")
        af("Name")
        af("NonLaborUnitsPercentComplete", default="0")
        af("NotesToResources",            default="")
        af("ObjectId")
        af("PercentComplete",             default="0")
        af("PercentCompleteType",         default="Physical")
        af("PhysicalPercentComplete",     default="0")
        af("PlannedDuration",             default="0.000000")
        af("PlannedFinishDate",           date=True)
        af("PlannedLaborCost",            default="0")
        af("PlannedLaborUnits",           default="0.000000")
        af("PlannedNonLaborCost",         default="0")
        af("PlannedNonLaborUnits",        default="0.000000")
        af("PlannedStartDate",            date=True)
        # PrimaryConstraintDate/Type: emit value when filled, xsi:nil when blank
        af("PrimaryConstraintDate",       date=True)
        af("PrimaryConstraintType")
        af("PrimaryResourceObjectId")
        af("ProjectObjectId",             default=cell(p, "ObjectId"))
        af("RemainingDuration",           default="0.000000")
        af("RemainingEarlyFinishDate",    date=True)
        af("RemainingEarlyStartDate",     date=True)
        af("RemainingLaborCost",          default="0")
        af("RemainingLaborUnits",         default="0.000000")
        af("RemainingLateFinishDate",     date=True)
        af("RemainingLateStartDate",      date=True)
        af("RemainingNonLaborCost",       default="0")
        af("RemainingNonLaborUnits",      default="0.000000")
        af("ResumeDate",                  date=True)
        af("SecondaryConstraintDate",     date=True)
        af("SecondaryConstraintType")
        af("StartDate",                   date=True)
        af("Status",                      default="Not Started")
        af("SuspendDate",                 date=True)
        af("Type",                        default="Task Dependent")
        af("UnitsPercentComplete",        default="0")

        # WBSObjectId — use nil when no value (valid for project-level activities)
        wbs_oid = cell(a, "WBSObjectId")
        if wbs_oid and wbs_oid in wbs_oid_set:
            make_tag(act_el, "WBSObjectId", text=wbs_oid)
        elif wbs_oid and wbs_oid not in wbs_oid_set:
            print(f"  [WARN] Activity {cell(a,'Id')}: WBSObjectId={wbs_oid} not in WBS list, "
                  f"remapped to {first_top_wbs_oid}")
            if first_top_wbs_oid:
                make_tag(act_el, "WBSObjectId", text=first_top_wbs_oid)
            else:
                make_tag(act_el, "WBSObjectId", nil=True)
        else:
            make_tag(act_el, "WBSObjectId", nil=True)

        # Activity codes (Code1_TypeObjectId / Code1_ValueObjectId …)
        for i in range(1, 6):
            type_oid = cell(a, f"Code{i}_TypeObjectId")
            val_oid  = cell(a, f"Code{i}_ValueObjectId")
            if type_oid and val_oid:
                code_el = ET.SubElement(act_el, "Code")
                make_tag(code_el, "TypeObjectId",  text=type_oid)
                make_tag(code_el, "ValueObjectId", text=val_oid)

        # UDF from UDFValue sheet — matched by Activity Id
        act_id_str = cell(a, "Id")
        for uv in (udf_value_list or []):
            if cell(uv, "ObjectType") == "Activity" and cell(uv, "ObjectId_Ref") == act_id_str:
                _inject_udf(act_el, uv, udf_type_lookup)

    # ── ResourceAssignments ───────────────────────────────────────────────────
    proj_oid = cell(p, "ObjectId") or ""
    for ra in (ra_list or []):
        ra_el = build_resource_assignment_element(ra, act_oid_to_id, proj_oid)
        proj.append(ra_el)

    # ── Project-level UDFs ────────────────────────────────────────────────────
    for uv in (udf_value_list or []):
        if cell(uv, "ObjectType") == "Project":
            _inject_udf(proj, uv, udf_type_lookup)

    # ── Relationships ─────────────────────────────────────────────────────────
    proj_id = cell(p, "Id") or ""

    skipped_rels = 0
    for r in relationship_list:
        pred_oid = cell(r, "PredecessorActivityObjectId")
        succ_oid = cell(r, "SuccessorActivityObjectId")
        if pred_oid not in act_oid_set or succ_oid not in act_oid_set:
            skipped_rels += 1
            print(f"  [WARN] Dropped relationship {cell(r,'ObjectId')}: "
                  f"pred={pred_oid}({'OK' if pred_oid in act_oid_set else 'MISSING'}) "
                  f"succ={succ_oid}({'OK' if succ_oid in act_oid_set else 'MISSING'})")
            continue
        pred_id = cell(r, "PredecessorActivityId") or act_oid_to_id.get(pred_oid, "")
        succ_id = cell(r, "SuccessorActivityId") or act_oid_to_id.get(succ_oid, "")

        rel_el = ET.SubElement(proj, "Relationship")
        make_tag(rel_el, "Lag",                         text=cell(r, "Lag") or "0.000000")
        make_tag(rel_el, "ObjectId",                    text=cell(r, "ObjectId") or "")
        make_tag(rel_el, "PredecessorActivityId",       text=pred_id)
        make_tag(rel_el, "PredecessorActivityObjectId", text=pred_oid)
        make_tag(rel_el, "PredecessorProjectObjectId",  text=cell(r, "PredecessorProjectObjectId") or proj_oid)
        make_tag(rel_el, "SuccessorActivityId",         text=succ_id)
        make_tag(rel_el, "SuccessorActivityObjectId",   text=succ_oid)
        make_tag(rel_el, "SuccessorProjectObjectId",    text=cell(r, "SuccessorProjectObjectId") or proj_oid)
        make_tag(rel_el, "Type",                        text=cell(r, "Type") or "Finish to Start")

    return proj


def build_resource_element(r_data):
    """Build a top-level <Resource> element matching EC00630 field order."""
    res_el = ET.Element("Resource")
    make_tag(res_el, "AutoComputeActuals",     text=cell(r_data, "AutoComputeActuals") or "1")
    make_tag(res_el, "CalculateCostFromUnits", text=cell(r_data, "CalculateCostFromUnits") or "1")
    make_tag(res_el, "CalendarObjectId",       text=cell(r_data, "CalendarObjectId") or "")
    code_v = cell(r_data, "Code")
    if code_v:
        make_tag(res_el, "Code", text=code_v)
    make_tag(res_el, "CurrencyObjectId",       text=cell(r_data, "CurrencyObjectId") or "1")
    make_tag(res_el, "DefaultUnitsPerTime",    text=cell(r_data, "DefaultUnitsPerTime") or "1.000000")
    make_tag(res_el, "EmailAddress",           nil=True)
    make_tag(res_el, "EmployeeId",             nil=True)
    make_tag(res_el, "GUID",                   text=cell(r_data, "GUID") or new_guid())
    make_tag(res_el, "Id",                     text=cell(r_data, "Id") or "")
    make_tag(res_el, "IsActive",               text=cell(r_data, "IsActive") or "1")
    make_tag(res_el, "IsOverTimeAllowed",      text=cell(r_data, "IsOverTimeAllowed") or "0")
    make_tag(res_el, "Name",                   text=cell(r_data, "Name") or "")
    make_tag(res_el, "ObjectId",               text=cell(r_data, "ObjectId") or "")
    make_tag(res_el, "OfficePhone",            nil=True)
    make_tag(res_el, "OtherPhone",             nil=True)
    make_tag(res_el, "OvertimeFactor",         text=cell(r_data, "OvertimeFactor") or "0")

    def r_opt(tag, key=None):
        v = cell(r_data, key or tag)
        if v:
            make_tag(res_el, tag, text=v)
        else:
            make_tag(res_el, tag, nil=True)

    r_opt("ParentObjectId")
    r_opt("PrimaryRoleObjectId")
    r_opt("ResourceNotes")
    make_tag(res_el, "ResourceType",           text=cell(r_data, "ResourceType") or "Labor")
    make_tag(res_el, "SequenceNumber",         text=cell(r_data, "SequenceNumber") or "0")
    make_tag(res_el, "ShiftObjectId",          nil=True)
    make_tag(res_el, "Title",                  nil=True)
    r_opt("UnitOfMeasureObjectId")
    make_tag(res_el, "UserObjectId",           nil=True)
    return res_el


def build_activity_code_element(ac_data):
    """Build a top-level <ActivityCode> element matching EC00630 field order."""
    ac_el = ET.Element("ActivityCode")
    make_tag(ac_el, "CodeTypeObjectId", text=cell(ac_data, "CodeTypeObjectId") or "")
    make_tag(ac_el, "CodeValue",        text=cell(ac_data, "CodeValue") or "")
    make_tag(ac_el, "Color",            text=cell(ac_data, "Color") or "#0000FF")
    make_tag(ac_el, "Description",      text=cell(ac_data, "Description") or "")
    make_tag(ac_el, "ObjectId",         text=cell(ac_data, "ObjectId") or "")
    make_tag(ac_el, "ParentObjectId",   nil=True)
    make_tag(ac_el, "ProjectObjectId",  nil=True)
    make_tag(ac_el, "SequenceNumber",   text=cell(ac_data, "SequenceNumber") or "0")
    return ac_el


def build_activity_code_type_element(act_data):
    """Build a top-level <ActivityCodeType> element."""
    el = ET.Element("ActivityCodeType")
    make_tag(el, "EPSObjectId",         nil=True)
    make_tag(el, "IsSecureCode",        text=cell(act_data, "IsSecureCode") or "0")
    make_tag(el, "Length",              text=cell(act_data, "Length") or "20")
    make_tag(el, "Name",                text=cell(act_data, "Name") or "")
    make_tag(el, "ObjectId",            text=cell(act_data, "ObjectId") or "")
    make_tag(el, "ProjectObjectId",     nil=True)
    make_tag(el, "RefProjectObjectIds", text=cell(act_data, "RefProjectObjectIds") or "")
    make_tag(el, "Scope",               text=cell(act_data, "Scope") or "Global")
    make_tag(el, "SequenceNumber",      text=cell(act_data, "SequenceNumber") or "0")
    return el


def build_udf_type_element(udf_data):
    """Build a top-level <UDFType> element."""
    el = ET.Element("UDFType")
    make_tag(el, "DataType",      text=cell(udf_data, "DataType") or "Indicator")
    make_tag(el, "IsSecureCode",  text=cell(udf_data, "IsSecureCode") or "0")
    make_tag(el, "ObjectId",      text=cell(udf_data, "ObjectId") or "")
    make_tag(el, "SubjectArea",   text=cell(udf_data, "SubjectArea") or "")
    make_tag(el, "Title",         text=cell(udf_data, "Title") or "")
    return el


# ── Pretty-print helper ───────────────────────────────────────────────────────
def prettify(element):
    rough = ET.tostring(element, encoding="unicode")
    reparsed = minidom.parseString(rough.encode("utf-8"))
    lines = reparsed.toprettyxml(indent="  ", encoding=None).splitlines()
    return "\n".join(l for l in lines if l.strip() and not l.startswith("<?xml"))


# ── XML write helper ──────────────────────────────────────────────────────────
def write_p6_xml(root_el, path):
    """Serialise root_el to a P6-compatible UTF-8 XML file and print stats."""
    ET.register_namespace("",    NS_BO)
    ET.register_namespace("xsi", NS_XSI)
    xml_str = '<?xml version="1.0" encoding="utf-8"?>\n' + prettify(root_el) + "\n"
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_str)

    tree  = ET.parse(path)
    r2    = tree.getroot()
    p_el  = r2.find(f"{{{NS_BO}}}Project")
    def cnt(tag): return len(r2.findall(f"{{{NS_BO}}}{tag}"))
    def pcnt(tag): return len(p_el.findall(f"{{{NS_BO}}}{tag}")) if p_el is not None else 0
    print(f"  Written: {path}")
    print(f"    Global refs — Currency:{cnt('Currency')} UDFType:{cnt('UDFType')} "
          f"OBS:{cnt('OBS')} Calendar:{cnt('Calendar')} "
          f"Role:{cnt('Role')} RoleRate:{cnt('RoleRate')}")
    print(f"    Top-level  — Resource:{cnt('Resource')} "
          f"ActivityCodeType:{cnt('ActivityCodeType')} ActivityCode:{cnt('ActivityCode')}")
    print(f"    In Project — WBS:{pcnt('WBS')} Activity:{pcnt('Activity')} "
          f"ResourceAssignment:{pcnt('ResourceAssignment')} "
          f"Relationship:{pcnt('Relationship')}")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    template_path = sys.argv[1] if len(sys.argv) > 1 else "P6_Import_Template.xlsx"
    output_arg    = sys.argv[2] if len(sys.argv) > 2 else "P6_Import_test.xml"

    out_dir    = os.path.dirname(os.path.abspath(template_path))
    pass1_path = os.path.join(out_dir, "P6_Import_pass1.xml")
    pass2_path = os.path.join(out_dir, "P6_Import_pass2.xml")
    pass3_path = os.path.join(out_dir, "P6_Import_pass3.xml")

    print(f"Reading template: {template_path}")
    wb = openpyxl.load_workbook(template_path, data_only=True)

    cfg = read_config(wb)
    if cfg:
        print(f"  _Config: {cfg}")

    proj_rows   = read_sheet(wb, "Project")
    wbs_rows    = read_sheet(wb, "WBS")
    act_rows    = read_sheet(wb, "Activity")
    rel_rows    = read_sheet(wb, "Relationship")
    udf_rows    = read_sheet(wb, "UDFType")
    act_rows_ac = read_sheet(wb, "ActivityCodeType")
    ac_rows     = read_sheet(wb, "ActivityCode")
    res_rows    = read_sheet(wb, "Resource")
    ra_rows     = read_sheet(wb, "ResourceAssignment")
    udf_val_rows = read_sheet(wb, "UDFValue")

    if not proj_rows:
        print("ERROR: No Project data found in template.")
        sys.exit(1)

    global DEFAULT_CALENDAR_OID, DEFAULT_OBS_OID
    if cfg.get("CalendarObjectId"):
        DEFAULT_CALENDAR_OID = cfg["CalendarObjectId"]
    if cfg.get("OBSObjectId"):
        DEFAULT_OBS_OID = cfg["OBSObjectId"]

    # Load required global reference elements from p6_reference.xml
    ref_xml = os.path.join(os.path.dirname(os.path.abspath(template_path)), "p6_reference.xml")
    ref_els = load_reference_elements(
        ref_xml, "Currency", "UDFType", "OBS", "Calendar", "Role", "RoleRate"
    )

    # ── Build UDFType title -> ObjectId lookup for UDFValue resolution ─────────
    # Prefer the UDFType sheet; fall back to p6_reference.xml elements.
    udf_type_lookup = {}
    if udf_rows:
        for u in udf_rows:
            title = cell(u, "Title")
            oid   = cell(u, "ObjectId")
            if title and oid:
                udf_type_lookup[title] = oid
    else:
        for el in ref_els.get("UDFType", []):
            title_el = el.find(f"{{{NS_BO}}}Title")
            oid_el   = el.find(f"{{{NS_BO}}}ObjectId")
            if title_el is not None and oid_el is not None and title_el.text and oid_el.text:
                udf_type_lookup[title_el.text] = oid_el.text
    if udf_type_lookup:
        print(f"  UDFType lookup: {len(udf_type_lookup)} types")
    if udf_val_rows:
        print(f"  UDFValue rows : {len(udf_val_rows)}")
    if ra_rows:
        print(f"  ResourceAssignment rows: {len(ra_rows)}")

    ET.register_namespace("",    NS_BO)
    ET.register_namespace("xsi", NS_XSI)

    def add_ns(el):
        if not el.tag.startswith("{"):
            el.tag = f"{{{NS_BO}}}{el.tag}"
        for child in el:
            add_ns(child)

    def make_ref_root():
        """Root element pre-loaded with the required global reference elements."""
        root = ET.Element(
            f"{{{NS_BO}}}APIBusinessObjects",
            attrib={f"{{{NS_XSI}}}schemaLocation": SCHEMA_LOC}
        )
        for el in ref_els.get("Currency", []):
            root.append(el)
        udf_source = udf_rows if udf_rows else []
        if udf_source:
            for u in udf_source:
                u_el = build_udf_type_element(u)
                u_el.tag = f"{{{NS_BO}}}UDFType"
                root.append(u_el)
        else:
            for el in ref_els.get("UDFType", []):
                root.append(el)
        for el in ref_els.get("OBS", []):
            root.append(el)
        for el in ref_els.get("Calendar", []):
            root.append(el)
        for el in ref_els.get("Role", []):
            root.append(el)
        for el in ref_els.get("RoleRate", []):
            root.append(el)
        return root

    # Build the full project element
    proj_full = build_project_element(
        proj_rows[0], wbs_rows, act_rows, rel_rows, ra_rows,
        udf_value_list=udf_val_rows,
        udf_type_lookup=udf_type_lookup,
    )
    proj_full.tag = f"{{{NS_BO}}}Project"
    add_ns(proj_full)

    # ── Pass 1: Create New Project ────────────────────────────────────────────
    # Global refs + Resource + ActivityCodeType + ActivityCode
    # Project: scalars + WBS + Activity  (NO Relationships, NO ResourceAssignments)
    root1 = make_ref_root()
    for r in res_rows:
        res_el = build_resource_element(r)
        res_el.tag = f"{{{NS_BO}}}Resource"
        root1.append(res_el)
    for act in act_rows_ac:
        act_el = build_activity_code_type_element(act)
        act_el.tag = f"{{{NS_BO}}}ActivityCodeType"
        root1.append(act_el)
    for ac in ac_rows:
        ac_el = build_activity_code_element(ac)
        ac_el.tag = f"{{{NS_BO}}}ActivityCode"
        root1.append(ac_el)

    proj1 = copy.deepcopy(proj_full)
    for rel in proj1.findall(f"{{{NS_BO}}}Relationship"):
        proj1.remove(rel)
    for ra in proj1.findall(f"{{{NS_BO}}}ResourceAssignment"):
        proj1.remove(ra)
    root1.append(proj1)

    # ── Pass 2: Update Existing Project — Relationships ───────────────────────
    # Global refs only
    # Project: scalars + Activity (same ObjectIds) + Relationship
    # (NO WBS, NO ResourceAssignments — activities now committed after pass 1)
    root2 = make_ref_root()
    proj2 = copy.deepcopy(proj_full)
    for wbs in proj2.findall(f"{{{NS_BO}}}WBS"):
        proj2.remove(wbs)
    for ra in proj2.findall(f"{{{NS_BO}}}ResourceAssignment"):
        proj2.remove(ra)
    root2.append(proj2)

    # ── Pass 3: Update Existing Project — ResourceAssignments ────────────────
    # Global refs only
    # Project: scalars + Activity (same ObjectIds) + ResourceAssignment
    # (NO WBS, NO Relationships — fk_taskactv_task needs committed activities)
    root3 = make_ref_root()
    proj3 = copy.deepcopy(proj_full)
    for wbs in proj3.findall(f"{{{NS_BO}}}WBS"):
        proj3.remove(wbs)
    for rel in proj3.findall(f"{{{NS_BO}}}Relationship"):
        proj3.remove(rel)
    root3.append(proj3)

    # ── Write all three files ─────────────────────────────────────────────────
    print()
    print("Pass 1 - Create New Project:")
    write_p6_xml(root1, pass1_path)
    print()
    print("Pass 2 - Update Existing Project (Relationships):")
    write_p6_xml(root2, pass2_path)
    print()
    print("Pass 3 - Update Existing Project (ResourceAssignments):")
    write_p6_xml(root3, pass3_path)
    print()
    print("Import instructions:")
    print(f"  Step 1: Import {pass1_path}")
    print( "          P6 action: File -> Import -> Primavera P6 XML -> Create New Project")
    print(f"  Step 2: Import {pass2_path}")
    print( "          P6 action: File -> Import -> Primavera P6 XML -> Update Existing Project")
    print(f"  Step 3: Import {pass3_path}")
    print( "          P6 action: File -> Import -> Primavera P6 XML -> Update Existing Project")


if __name__ == "__main__":
    main()
