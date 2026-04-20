"""
excel_to_opc_converter.py
Reads P6_Import_Template.xlsx and produces a single OPC XML file for
one-pass import into Oracle Primavera Cloud.

OPC differences from P6:
  - Activity Type "WBS Summary" is not supported; those activities are
    filtered out (skipped with warning) and represented only as WBS elements
  - UDFType with DataType "Indicator" or "Formula" are skipped + warned
  - UDFValue rows for skipped UDFTypes are also suppressed
  - Single output file: {name}_OPC.xml  (no 3-pass needed)
  - Project Id truncated to 20 characters (OPC limit)
  - Calendar copied from p6_reference.xml
  - Post-import reminder: reschedule project and recalculate costs in OPC

Sheet layout:
  Row 1: field names
  Row 2: field types
  Row 3+: data rows

Requires p6_reference.xml in the same folder as this script.

Usage:
  py excel_to_opc_converter.py [template.xlsx] [output_dir]
"""

import sys
import os
import copy
import uuid
import re
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime
import openpyxl

NS_BO  = "http://xmlns.oracle.com/Primavera/P6Professional/V18.8/API/BusinessObjects"
NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
SCHEMA_LOC = (
    "http://xmlns.oracle.com/Primavera/P6Professional/V18.8/API/BusinessObjects "
    "http://xmlns.oracle.com/Primavera/P6Professional/V18.8/API/p6apibo.xsd"
)

DEFAULT_CALENDAR_OID = ""
DEFAULT_OBS_OID      = ""
DEFAULT_ROOT_WBS_OID = "17583"

ACTIVITY_NIL_ALWAYS = {
    "ExpectedFinishDate",
    "ExternalEarlyStartDate", "ExternalLateFinishDate",
    "ResumeDate",
    "SecondaryConstraintDate", "SecondaryConstraintType", "SuspendDate",
}
ACTIVITY_NIL_OPTIONAL = {
    "ActualFinishDate", "ActualStartDate",
    "PrimaryConstraintDate", "PrimaryConstraintType",
    "PrimaryResourceObjectId", "WBSObjectId",
}

# Activity types not supported in OPC — activities with these types are skipped
OPC_UNSUPPORTED_ACTIVITY_TYPES = {"WBS Summary"}

# UDFType DataType values not supported in OPC
OPC_SKIP_UDF_DATATYPES = {"Indicator", "Formula"}


def new_guid():
    return "{" + str(uuid.uuid4()).upper() + "}"

def fmt_date(val):
    if val is None:
        return None
    if isinstance(val, datetime):
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
    return s

def make_tag(parent, tag, text=None, nil=False):
    el = ET.SubElement(parent, tag)
    if nil:
        el.set(f"{{{NS_XSI}}}nil", "true")
    elif text is not None:
        el.text = str(text)
    return el

def read_config(wb):
    if "_Config" not in wb.sheetnames:
        return {}
    ws   = wb["_Config"]
    rows = list(ws.iter_rows(values_only=True))
    cfg  = {}
    for row in rows[2:]:
        if row and row[0] is not None and row[1] is not None:
            cfg[str(row[0]).strip()] = str(row[1]).strip()
    return cfg

def load_reference_elements(ref_path, *tag_names):
    if not os.path.exists(ref_path):
        print()
        print("ERROR: p6_reference.xml not found.")
        print()
        print("  Export any project from P6/OPC: File -> Export -> Primavera P6 XML")
        print(f"  Rename it to p6_reference.xml and place it in: {os.path.dirname(ref_path)}")
        print()
        sys.exit(1)
    try:
        ref_root = ET.parse(ref_path).getroot()
        result   = {}
        for t in tag_names:
            result[t] = [copy.deepcopy(el)
                         for el in ref_root.findall(f"{{{NS_BO}}}{t}")]
        print(f"  Loaded from {os.path.basename(ref_path)}: "
              f"{ {t: len(v) for t, v in result.items()} }")
        return result
    except Exception as e:
        print(f"ERROR: Could not parse {ref_path}: {e}")
        sys.exit(1)

def read_sheet(wb, name):
    if name not in wb.sheetnames:
        return []
    ws   = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 3:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    data    = []
    for row in rows[2:]:
        if all(v is None for v in row):
            continue
        rec = {}
        for i, h in enumerate(headers):
            if h:
                rec[h] = row[i] if i < len(row) else None
        data.append(rec)
    return data

def cell(rec, *keys):
    for k in keys:
        v = rec.get(k)
        if v is not None and str(v).strip() != "":
            return str(v).strip()
    return None


def filter_udf_for_opc(udf_rows, ref_els, warnings):
    """
    Remove UDFType entries with DataType Indicator or Formula from both the
    template rows and the ref_els dict.  Returns a set of skipped titles so
    that UDFValue rows referencing them can also be suppressed.
    """
    skipped_titles = set()

    # Filter template UDFType rows
    filtered_rows = []
    for u in udf_rows:
        dt    = cell(u, "DataType") or ""
        title = cell(u, "Title") or "?"
        if dt in OPC_SKIP_UDF_DATATYPES:
            warnings.append(
                f"[WARN] UDFType skipped (not supported in OPC): {title} ({dt})"
            )
            skipped_titles.add(title)
        else:
            filtered_rows.append(u)

    # Filter ref_els UDFType entries (used when template has no UDFType sheet)
    filtered_ref = []
    for el in ref_els.get("UDFType", []):
        dt_el    = el.find(f"{{{NS_BO}}}DataType")
        title_el = el.find(f"{{{NS_BO}}}Title")
        dt    = (dt_el.text    or "").strip() if dt_el    is not None else ""
        title = (title_el.text or "?").strip() if title_el is not None else "?"
        if dt in OPC_SKIP_UDF_DATATYPES:
            if title not in skipped_titles:
                warnings.append(
                    f"[WARN] UDFType skipped (not supported in OPC): {title} ({dt})"
                )
                skipped_titles.add(title)
        else:
            filtered_ref.append(el)
    ref_els["UDFType"] = filtered_ref

    return filtered_rows, skipped_titles


def filter_activities_for_opc(act_rows, warnings):
    """
    Remove Activity rows with unsupported Type values (e.g. 'WBS Summary').
    Returns filtered list.
    """
    filtered = []
    for a in act_rows:
        act_type = cell(a, "Type") or ""
        if act_type in OPC_UNSUPPORTED_ACTIVITY_TYPES:
            act_id = cell(a, "Id") or cell(a, "ObjectId") or "?"
            warnings.append(
                f"[WARN] Activity {act_id}: Type='{act_type}' not supported in OPC "
                f"— filtered out (emit as WBS element only)"
            )
        else:
            filtered.append(a)
    return filtered


def _inject_udf(parent_el, udf_v, udf_type_lookup, skipped_titles):
    """Append a <UDF> child, skipping entries for OPC-unsupported types."""
    title = cell(udf_v, "UDFTypeTitle")
    if not title:
        return
    if title in skipped_titles:
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


def build_resource_assignment_element(ra_data, act_oid_to_id, proj_oid):
    ra_el = ET.Element("ResourceAssignment")

    act_oid = cell(ra_data, "ActivityObjectId")
    res_oid = cell(ra_data, "ResourceObjectId")

    make_tag(ra_el, "ActivityObjectId",      text=act_oid or "")
    make_tag(ra_el, "ActualCost",            text=cell(ra_data, "ActualCost") or "0.000000")
    make_tag(ra_el, "ActualCurve",           nil=True)
    make_tag(ra_el, "ActualFinishDate",      nil=True)
    make_tag(ra_el, "ActualOvertimeCost",    text="0.000000")
    make_tag(ra_el, "ActualOvertimeUnits",   text="0.000000")
    make_tag(ra_el, "ActualRegularCost",     text="0.000000")
    make_tag(ra_el, "ActualRegularUnits",    text="0.000000")
    make_tag(ra_el, "ActualStartDate",       nil=True)
    make_tag(ra_el, "ActualThisPeriodCost",  text="0.000000")
    make_tag(ra_el, "ActualThisPeriodUnits", text="0.000000")
    make_tag(ra_el, "ActualUnits",           text=cell(ra_data, "ActualUnits") or "0.000000")
    planned_cost  = cell(ra_data, "PlannedCost")  or "0.000000"
    planned_units = cell(ra_data, "PlannedUnits") or "0.000000"
    make_tag(ra_el, "AtCompletionCost",      text=cell(ra_data, "AtCompletionCost")  or planned_cost)
    make_tag(ra_el, "AtCompletionUnits",     text=cell(ra_data, "AtCompletionUnits") or planned_units)
    cost_acct = cell(ra_data, "CostAccountObjectId")
    if cost_acct:
        make_tag(ra_el, "CostAccountObjectId", text=cost_acct)
    else:
        make_tag(ra_el, "CostAccountObjectId", nil=True)
    make_tag(ra_el, "DrivingActivityDatesFlag", text="1")
    planned_finish = cell(ra_data, "PlannedFinishDate")
    planned_start  = cell(ra_data, "PlannedStartDate")
    if planned_finish:
        make_tag(ra_el, "FinishDate",        text=fmt_date(planned_finish))
    else:
        make_tag(ra_el, "FinishDate",        nil=True)
    make_tag(ra_el, "GUID",                  text=new_guid())
    make_tag(ra_el, "IsCostUnitsLinked",     text="1")
    make_tag(ra_el, "IsPrimaryResource",     text=cell(ra_data, "IsPrimaryResource") or "0")
    make_tag(ra_el, "ObjectId",              text=cell(ra_data, "ObjectId") or "")
    make_tag(ra_el, "OvertimeFactor",        text="0")
    make_tag(ra_el, "PlannedCost",           text=planned_cost)
    make_tag(ra_el, "PlannedCurve",          nil=True)
    if planned_finish:
        make_tag(ra_el, "PlannedFinishDate", text=fmt_date(planned_finish))
    else:
        make_tag(ra_el, "PlannedFinishDate", nil=True)
    make_tag(ra_el, "PlannedLag",            text="0.000000")
    if planned_start:
        make_tag(ra_el, "PlannedStartDate",  text=fmt_date(planned_start))
    else:
        make_tag(ra_el, "PlannedStartDate",  nil=True)
    make_tag(ra_el, "PlannedUnits",          text=planned_units)
    make_tag(ra_el, "PlannedUnitsPerTime",   text=cell(ra_data, "PlannedUnitsPerTime") or "0.000000")
    make_tag(ra_el, "Proficiency",           text=cell(ra_data, "Proficiency") or "3 - Skilled")
    make_tag(ra_el, "ProjectObjectId",       text=cell(ra_data, "ProjectObjectId") or proj_oid or "")
    make_tag(ra_el, "RateSource",            text="Resource")
    make_tag(ra_el, "RateType",              text=cell(ra_data, "RateType") or "Price / Unit")
    rem_cost  = cell(ra_data, "RemainingCost")  or planned_cost
    rem_units = cell(ra_data, "RemainingUnits") or planned_units
    make_tag(ra_el, "RemainingCost",         text=rem_cost)
    make_tag(ra_el, "RemainingCurve",        nil=True)
    make_tag(ra_el, "RemainingDuration",     text=cell(ra_data, "RemainingDuration") or "0.000000")
    if planned_finish:
        make_tag(ra_el, "RemainingFinishDate", text=fmt_date(planned_finish))
    else:
        make_tag(ra_el, "RemainingFinishDate", nil=True)
    make_tag(ra_el, "RemainingLag",          text="0.000000")
    if planned_start:
        make_tag(ra_el, "RemainingStartDate", text=fmt_date(planned_start))
    else:
        make_tag(ra_el, "RemainingStartDate", nil=True)
    make_tag(ra_el, "RemainingUnits",        text=rem_units)
    make_tag(ra_el, "RemainingUnitsPerTime", text=cell(ra_data, "RemainingUnitsPerTime") or "0.000000")
    make_tag(ra_el, "ResourceCurveObjectId", nil=True)
    make_tag(ra_el, "ResourceObjectId",      text=res_oid or "")
    make_tag(ra_el, "ResourceType",          text=cell(ra_data, "ResourceType") or "Labor")
    make_tag(ra_el, "RoleObjectId",          nil=True)
    if planned_start:
        make_tag(ra_el, "StartDate",         text=fmt_date(planned_start))
    else:
        make_tag(ra_el, "StartDate",         nil=True)
    make_tag(ra_el, "UnitsPercentComplete",  text="0")
    wbs_oid = cell(ra_data, "WBSObjectId")
    if wbs_oid:
        make_tag(ra_el, "WBSObjectId",       text=wbs_oid)
    else:
        make_tag(ra_el, "WBSObjectId",       nil=True)
    return ra_el


def build_project_element(proj_data, wbs_list, activity_list,
                           relationship_list, ra_list,
                           udf_value_list=None, udf_type_lookup=None,
                           skipped_udf_titles=None):
    """Build the <Project> element with all nested children."""
    proj     = ET.Element("Project")
    p        = proj_data
    skipped_udf_titles = skipped_udf_titles or set()

    def ps(tag, key=None, val=None, nil=False, date=False, dft=None):
        v = val if val is not None else cell(p, key or tag)
        if v is None:
            v = dft
        if nil:
            make_tag(proj, tag, nil=True)
        else:
            make_tag(proj, tag, text=(fmt_date(v) if date else v) if v is not None else "")

    # OPC: truncate Id to 20 chars
    raw_id = cell(p, "Id") or ""
    opc_id = raw_id[:20]
    if len(raw_id) > 20:
        print(f"  [WARN] Project Id '{raw_id}' truncated to 20 chars: '{opc_id}'")

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
    make_tag(proj, "GUID",               text=cell(p, "GUID") or new_guid())
    make_tag(proj, "Id",                 text=opc_id)  # truncated
    ps("IndependentETCLaborUnits",       dft="0.000000")
    ps("IndependentETCTotalCost",        dft="0.000000")
    make_tag(proj, "LastFinancialPeriodObjectId", nil=True)
    ps("LevelingPriority",               dft="10")
    ps("LinkActualToActualThisPeriod",   dft="1")
    ps("LinkPercentCompleteWithActual",  dft="0")
    ps("LinkPlannedAndAtCompletionFlag", dft="1")
    make_tag(proj, "MustFinishByDate",   nil=True)
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
        make_tag(wbs_el, "AnticipatedFinishDate",     nil=True)
        make_tag(wbs_el, "AnticipatedStartDate",      nil=True)
        make_tag(wbs_el, "Code",                      text=cell(w, "Code") or "1")
        make_tag(wbs_el, "EarnedValueComputeType",    text=cell(w, "EarnedValueComputeType") or "Activity Percent Complete")
        make_tag(wbs_el, "EarnedValueETCComputeType", text=cell(w, "EarnedValueETCComputeType") or "PF = 1 / CPI")
        make_tag(wbs_el, "EarnedValueETCUserValue",   text=cell(w, "EarnedValueETCUserValue") or "0")
        make_tag(wbs_el, "EarnedValueUserPercent",    text=cell(w, "EarnedValueUserPercent") or "0.0")
        make_tag(wbs_el, "GUID",                      text=cell(w, "GUID") or new_guid())
        make_tag(wbs_el, "IndependentETCLaborUnits",  text=cell(w, "IndependentETCLaborUnits") or "0.000000")
        make_tag(wbs_el, "IndependentETCTotalCost",   text=cell(w, "IndependentETCTotalCost")  or "0.000000")
        make_tag(wbs_el, "Name",                      text=cell(w, "Name") or "")
        make_tag(wbs_el, "OBSObjectId",               text=cell(w, "OBSObjectId") or DEFAULT_OBS_OID)
        make_tag(wbs_el, "ObjectId",                  text=cell(w, "ObjectId") or "")
        make_tag(wbs_el, "OriginalBudget",             text=cell(w, "OriginalBudget") or "0.000000")
        parent_oid = cell(w, "ParentObjectId")
        if parent_oid:
            make_tag(wbs_el, "ParentObjectId",        text=parent_oid)
        else:
            make_tag(wbs_el, "ParentObjectId",        nil=True)
        make_tag(wbs_el, "ProjectObjectId",           text=cell(w, "ProjectObjectId") or cell(p, "ObjectId") or "")
        make_tag(wbs_el, "SequenceNumber",            text=cell(w, "SequenceNumber") or "0")
        make_tag(wbs_el, "Status",                    text=cell(w, "Status") or "Active")
        make_tag(wbs_el, "WBSCategoryObjectId",       nil=True)

        # Inline UDF columns from WBS sheet — only if type not skipped
        udf_type_oid = cell(w, "UDF_TypeObjectId")
        udf_ind_val  = cell(w, "UDF_IndicatorValue")
        if udf_type_oid:
            udf_el = ET.SubElement(wbs_el, "UDF")
            make_tag(udf_el, "TypeObjectId",   text=udf_type_oid)
            make_tag(udf_el, "IndicatorValue", text=udf_ind_val or "Green")

        wbs_code = cell(w, "Code")
        for uv in (udf_value_list or []):
            if cell(uv, "ObjectType") == "WBS" and cell(uv, "ObjectId_Ref") == wbs_code:
                _inject_udf(wbs_el, uv, udf_type_lookup, skipped_udf_titles)

    # ── Activity lookup maps ──────────────────────────────────────────────────
    act_oid_to_id = {}
    act_oid_set   = set()
    for a in activity_list:
        oid = cell(a, "ObjectId")
        aid = cell(a, "Id")
        if oid:
            act_oid_set.add(oid)
            if aid:
                act_oid_to_id[oid] = aid

    wbs_oid_set        = set()
    first_top_wbs_oid  = None
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
                make_tag(act_el, tag,
                         text=(fmt_date(v) if date else v) if v is not None else "")

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
        make_tag(act_el, "GUID",          text=cell(a, "GUID") or new_guid())
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

        for i in range(1, 6):
            type_oid = cell(a, f"Code{i}_TypeObjectId")
            val_oid  = cell(a, f"Code{i}_ValueObjectId")
            if type_oid and val_oid:
                code_el = ET.SubElement(act_el, "Code")
                make_tag(code_el, "TypeObjectId",  text=type_oid)
                make_tag(code_el, "ValueObjectId", text=val_oid)

        act_id_str = cell(a, "Id")
        for uv in (udf_value_list or []):
            if cell(uv, "ObjectType") == "Activity" and cell(uv, "ObjectId_Ref") == act_id_str:
                _inject_udf(act_el, uv, udf_type_lookup, skipped_udf_titles)

    # ── ResourceAssignments ───────────────────────────────────────────────────
    proj_oid = cell(p, "ObjectId") or ""
    for ra in (ra_list or []):
        ra_el = build_resource_assignment_element(ra, act_oid_to_id, proj_oid)
        proj.append(ra_el)

    # ── Project-level UDFs ────────────────────────────────────────────────────
    for uv in (udf_value_list or []):
        if cell(uv, "ObjectType") == "Project":
            _inject_udf(proj, uv, udf_type_lookup, skipped_udf_titles)

    # ── Relationships ─────────────────────────────────────────────────────────
    proj_oid_str = cell(p, "ObjectId") or ""
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
        succ_id = cell(r, "SuccessorActivityId")   or act_oid_to_id.get(succ_oid, "")
        rel_el  = ET.SubElement(proj, "Relationship")
        make_tag(rel_el, "Lag",                         text=cell(r, "Lag") or "0.000000")
        make_tag(rel_el, "ObjectId",                    text=cell(r, "ObjectId") or "")
        make_tag(rel_el, "PredecessorActivityId",       text=pred_id)
        make_tag(rel_el, "PredecessorActivityObjectId", text=pred_oid)
        make_tag(rel_el, "PredecessorProjectObjectId",  text=cell(r, "PredecessorProjectObjectId") or proj_oid_str)
        make_tag(rel_el, "SuccessorActivityId",         text=succ_id)
        make_tag(rel_el, "SuccessorActivityObjectId",   text=succ_oid)
        make_tag(rel_el, "SuccessorProjectObjectId",    text=cell(r, "SuccessorProjectObjectId") or proj_oid_str)
        make_tag(rel_el, "Type",                        text=cell(r, "Type") or "Finish to Start")

    return proj


def build_resource_element(r_data):
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

    def r_opt(tag):
        v = cell(r_data, tag)
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
    el = ET.Element("UDFType")
    make_tag(el, "DataType",     text=cell(udf_data, "DataType") or "Text")
    make_tag(el, "IsSecureCode", text=cell(udf_data, "IsSecureCode") or "0")
    make_tag(el, "ObjectId",     text=cell(udf_data, "ObjectId") or "")
    make_tag(el, "SubjectArea",  text=cell(udf_data, "SubjectArea") or "")
    make_tag(el, "Title",        text=cell(udf_data, "Title") or "")
    return el


def prettify(element):
    rough    = ET.tostring(element, encoding="unicode")
    reparsed = minidom.parseString(rough.encode("utf-8"))
    lines    = reparsed.toprettyxml(indent="  ", encoding=None).splitlines()
    return "\n".join(l for l in lines if l.strip() and not l.startswith("<?xml"))


def write_opc_xml(root_el, path):
    ET.register_namespace("",    NS_BO)
    ET.register_namespace("xsi", NS_XSI)
    xml_str = '<?xml version="1.0" encoding="utf-8"?>\n' + prettify(root_el) + "\n"
    with open(path, "w", encoding="utf-8") as f:
        f.write(xml_str)

    tree = ET.parse(path)
    r2   = tree.getroot()
    p_el = r2.find(f"{{{NS_BO}}}Project")
    def cnt(tag):  return len(r2.findall(f"{{{NS_BO}}}{tag}"))
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


def main():
    template_path = sys.argv[1] if len(sys.argv) > 1 else "P6_Import_Template.xlsx"
    out_dir       = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(template_path))

    base     = os.path.splitext(os.path.basename(template_path))[0]
    opc_path = os.path.join(out_dir, f"{base}_OPC.xml")

    print(f"Reading template: {template_path}")
    wb = openpyxl.load_workbook(template_path, data_only=True)

    cfg = read_config(wb)
    if cfg:
        print(f"  _Config: {cfg}")

    proj_rows    = read_sheet(wb, "Project")
    wbs_rows     = read_sheet(wb, "WBS")
    act_rows     = read_sheet(wb, "Activity")
    rel_rows     = read_sheet(wb, "Relationship")
    udf_rows     = read_sheet(wb, "UDFType")
    act_type_rows = read_sheet(wb, "ActivityCodeType")
    ac_rows      = read_sheet(wb, "ActivityCode")
    res_rows     = read_sheet(wb, "Resource")
    ra_rows      = read_sheet(wb, "ResourceAssignment")
    udf_val_rows = read_sheet(wb, "UDFValue")

    if not proj_rows:
        print("ERROR: No Project data found in template.")
        sys.exit(1)

    global DEFAULT_CALENDAR_OID, DEFAULT_OBS_OID
    if cfg.get("CalendarObjectId"):
        DEFAULT_CALENDAR_OID = cfg["CalendarObjectId"]
    if cfg.get("OBSObjectId"):
        DEFAULT_OBS_OID = cfg["OBSObjectId"]

    ref_xml = os.path.join(os.path.dirname(os.path.abspath(__file__)), "p6_reference.xml")
    ref_els = load_reference_elements(
        ref_xml, "Currency", "UDFType", "OBS", "Calendar", "Role", "RoleRate"
    )

    # ── OPC: filter out Indicator/Formula UDFTypes ────────────────────────────
    warnings = []
    udf_rows_filtered, skipped_udf_titles = filter_udf_for_opc(udf_rows, ref_els, warnings)

    # ── OPC: filter out WBS Summary activities ────────────────────────────────
    act_rows_filtered = filter_activities_for_opc(act_rows, warnings)

    if warnings:
        print()
        for w in warnings:
            print(" ", w)
        print()

    # ── UDFType title → ObjectId lookup ──────────────────────────────────────
    udf_type_lookup = {}
    if udf_rows_filtered:
        for u in udf_rows_filtered:
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

    def make_root():
        root = ET.Element(
            f"{{{NS_BO}}}APIBusinessObjects",
            attrib={f"{{{NS_XSI}}}schemaLocation": SCHEMA_LOC}
        )
        for el in ref_els.get("Currency", []):
            root.append(el)
        if udf_rows_filtered:
            for u in udf_rows_filtered:
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

    # ── Build full project element (all data in one pass) ─────────────────────
    proj_el = build_project_element(
        proj_rows[0], wbs_rows, act_rows_filtered, rel_rows, ra_rows,
        udf_value_list=udf_val_rows,
        udf_type_lookup=udf_type_lookup,
        skipped_udf_titles=skipped_udf_titles,
    )
    proj_el.tag = f"{{{NS_BO}}}Project"
    add_ns(proj_el)

    root = make_root()
    for r in res_rows:
        res_el = build_resource_element(r)
        res_el.tag = f"{{{NS_BO}}}Resource"
        root.append(res_el)
    for act in act_type_rows:
        act_el = build_activity_code_type_element(act)
        act_el.tag = f"{{{NS_BO}}}ActivityCodeType"
        root.append(act_el)
    for ac in ac_rows:
        ac_el = build_activity_code_element(ac)
        ac_el.tag = f"{{{NS_BO}}}ActivityCode"
        root.append(ac_el)
    root.append(proj_el)

    print("\nWriting OPC XML ...")
    write_opc_xml(root, opc_path)

    wbs_cnt  = len(wbs_rows)
    act_cnt  = len(act_rows_filtered)
    skp_act  = len(act_rows) - act_cnt
    rel_cnt  = len(rel_rows)
    ra_cnt   = len(ra_rows)
    skp_udf  = len(skipped_udf_titles)

    print()
    print(f"Summary: {wbs_cnt} WBS  |  {act_cnt} Activities"
          + (f"  ({skp_act} WBS Summary filtered)" if skp_act else "")
          + f"  |  {rel_cnt} Relationships  |  {ra_cnt} ResourceAssignments")
    if skp_udf:
        print(f"         {skp_udf} UDFType(s) skipped (Indicator/Formula not supported in OPC)")
    if warnings:
        print(f"         {len(warnings)} warnings (see above)")

    print(f"\nOutput: {opc_path}")
    print("\n[WARN] After OPC import: reschedule project and recalculate costs manually in OPC")

    return opc_path


if __name__ == "__main__":
    main()
