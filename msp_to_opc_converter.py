"""
msp_to_opc_converter.py
Converts MS Project XML (.xml) to Oracle Primavera Cloud (OPC) XML.

Single-pass import — one output file: {name}_OPC.xml

Differences from P6 converter:
  - No WBS Summary activity type; summary tasks become WBS elements only
  - UDFType elements with DataType "Indicator" or "Formula" are skipped
  - Single output file (no 3-pass needed)
  - ProjectId truncated to 20 characters (OPC limit)
  - Calendar copied from p6_reference.xml
  - Post-import reminder: reschedule project and recalculate costs in OPC

Requires p6_reference.xml in the same folder as this script.
Export any project from P6/OPC: File → Export → Primavera P6 XML → rename to p6_reference.xml

Usage:
  py msp_to_opc_converter.py <input_msp.xml> [output_dir]
"""

import sys, os, copy, uuid, re
import xml.etree.ElementTree as ET
from xml.dom import minidom
from datetime import datetime

MSP_NS = "http://schemas.microsoft.com/project"
NS_BO  = "http://xmlns.oracle.com/Primavera/P6Professional/V18.8/API/BusinessObjects"
NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
SCHEMA_LOC = (
    "http://xmlns.oracle.com/Primavera/P6Professional/V18.8/API/BusinessObjects "
    "http://xmlns.oracle.com/Primavera/P6Professional/V18.8/API/p6apibo.xsd"
)

REL_TYPE_MAP = {
    "0": "Finish to Finish",
    "1": "Finish to Start",
    "2": "Start to Finish",
    "3": "Start to Start",
}
CONSTRAINT_MAP = {
    "2": "Mandatory Start",
    "3": "Mandatory Finish",
    "4": "Start On or After",
    "5": "Start On or Before",
    "6": "Finish On or After",
    "7": "Finish On or Before",
}
RESOURCE_TYPE_MAP = {
    "0": "Labor",
    "1": "Material",
    "2": "Nonlabor",
}

# UDFType DataType values unsupported in OPC
OPC_SKIP_UDF_TYPES = {"Indicator", "Formula"}

ACTIVITY_NIL_ALWAYS = {
    "ExpectedFinishDate", "ExternalEarlyStartDate", "ExternalLateFinishDate",
    "ResumeDate", "SecondaryConstraintDate", "SecondaryConstraintType", "SuspendDate",
}
ACTIVITY_NIL_OPTIONAL = {
    "ActualFinishDate", "ActualStartDate",
    "PrimaryConstraintDate", "PrimaryConstraintType",
    "PrimaryResourceObjectId", "WBSObjectId",
}

WBS_OID_OFFSET = 300000


def new_guid():
    return "{" + str(uuid.uuid4()).upper() + "}"

def fmt_date(val):
    if val is None:
        return None
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

def _m(tag):
    return f"{{{MSP_NS}}}{tag}"

def _mtext(el, tag, default=""):
    child = el.find(_m(tag))
    return (child.text or "").strip() if child is not None else default

def parse_iso_dur_hours(s):
    m = re.match(r'PT(?:(\d+(?:\.\d+)?)H)?(?:(\d+(?:\.\d+)?)M)?(?:(\d+(?:\.\d+)?)S)?', s or "")
    if not m:
        return 0.0
    return float(m.group(1) or 0) + float(m.group(2) or 0)/60 + float(m.group(3) or 0)/3600

def parse_lag_hours(link_lag_str):
    try:
        return float(link_lag_str) / 600.0
    except (TypeError, ValueError):
        return 0.0


def load_reference_elements(ref_path, *tag_names):
    if not os.path.exists(ref_path):
        print("\nERROR: p6_reference.xml not found.")
        print("  Export any project from P6/OPC: File > Export > Primavera P6 XML")
        print(f"  Rename it to p6_reference.xml and place it here: {os.path.dirname(ref_path)}\n")
        sys.exit(1)
    try:
        ref_root = ET.parse(ref_path).getroot()
        result = {}
        for t in tag_names:
            result[t] = [copy.deepcopy(e) for e in ref_root.findall(f"{{{NS_BO}}}{t}")]
        print(f"  Loaded from p6_reference.xml: { {t: len(v) for t,v in result.items()} }")
        return result
    except Exception as e:
        print(f"ERROR: Cannot parse p6_reference.xml: {e}")
        sys.exit(1)


def filter_udf_types(ref_els, warnings):
    """Remove UDFType entries with unsupported DataType values (Indicator, Formula)."""
    filtered = []
    skipped  = []
    for el in ref_els.get("UDFType", []):
        dt_el = el.find(f"{{{NS_BO}}}DataType")
        dt    = (dt_el.text or "").strip() if dt_el is not None else ""
        if dt in OPC_SKIP_UDF_TYPES:
            subjt_el = el.find(f"{{{NS_BO}}}Title")
            title    = (subjt_el.text or "").strip() if subjt_el is not None else "?"
            skipped.append(f"{title} ({dt})")
        else:
            filtered.append(el)
    if skipped:
        for s in skipped:
            warnings.append(f"[WARN] UDFType skipped (not supported in OPC): {s}")
    ref_els["UDFType"] = filtered


def load_reference_project_fields(ref_path):
    XSI = "http://www.w3.org/2001/XMLSchema-instance"
    try:
        ref_root = ET.parse(ref_path).getroot()
        proj_el  = ref_root.find(f"{{{NS_BO}}}Project")
        if proj_el is None:
            return {}
        def _get(tag):
            el = proj_el.find(f"{{{NS_BO}}}{tag}")
            if el is None:
                return None
            if el.get(f"{{{XSI}}}nil"):
                return None
            return (el.text or "").strip() or None
        return {
            "ParentEPSObjectId": _get("ParentEPSObjectId"),
            "ref_WBSObjectId":   _get("WBSObjectId"),
        }
    except Exception:
        return {}

def _ref_first_oid(ref_els, tag):
    for el in ref_els.get(tag, []):
        oid = el.find(f"{{{NS_BO}}}ObjectId")
        if oid is not None and oid.text:
            return oid.text
    return ""


class MSPProject:
    def __init__(self, path):
        print(f"Parsing: {path}")
        root = ET.parse(path).getroot()

        self.name      = _mtext(root, "Name", "Untitled")
        self.title     = _mtext(root, "Title") or self.name
        self.start     = _mtext(root, "StartDate")
        self.finish    = _mtext(root, "FinishDate")
        self.data_date = _mtext(root, "CurrentDate") or _mtext(root, "StatusDate")

        self._p6_fid = "188743731"
        for ea in root.findall(f".//{_m('ExtendedAttribute')}"):
            if _mtext(ea, "Alias") == "P6_ActivityId" and _mtext(ea, "FieldName") == "Text1":
                self._p6_fid = _mtext(ea, "FieldID")
                break

        self.tasks      = {}
        self.task_order = []
        tasks_el = root.find(_m("Tasks"))
        if tasks_el is not None:
            for t_el in tasks_el.findall(_m("Task")):
                t = self._parse_task(t_el)
                if t:
                    self.tasks[t["uid"]] = t
                    self.task_order.append(t["uid"])

        self._build_hierarchy()

        self.resources = {}
        res_el = root.find(_m("Resources"))
        if res_el is not None:
            for r_el in res_el.findall(_m("Resource")):
                r = self._parse_resource(r_el)
                if r:
                    self.resources[r["uid"]] = r

        self.assignments = {}
        asgn_el = root.find(_m("Assignments"))
        if asgn_el is not None:
            for a_el in asgn_el.findall(_m("Assignment")):
                a = self._parse_assignment(a_el)
                if a:
                    self.assignments[a["uid"]] = a

        summary_cnt  = sum(1 for t in self.tasks.values() if t["summary"] == "1")
        activity_cnt = sum(1 for t in self.tasks.values()
                          if t["summary"] == "0" and t["active"] == "1")
        print(f"  Tasks: {len(self.tasks)} total  |  WBS/Summary: {summary_cnt}  "
              f"|  Activities: {activity_cnt}")
        print(f"  Resources: {len(self.resources)}  |  Assignments: {len(self.assignments)}")

    def _parse_task(self, el):
        if _mtext(el, "IsNull") == "1":
            return None
        uid = _mtext(el, "UID")
        if not uid:
            return None

        p6_id = None
        for ea in el.findall(_m("ExtendedAttribute")):
            if _mtext(ea, "FieldID") == self._p6_fid:
                p6_id = _mtext(ea, "Value") or None
                break

        pred_links = []
        for pl in el.findall(_m("PredecessorLink")):
            if _mtext(pl, "CrossProject") == "1":
                continue
            pu = _mtext(pl, "PredecessorUID")
            if pu:
                pred_links.append({
                    "pred_uid": pu,
                    "type":     _mtext(pl, "Type", "1"),
                    "link_lag": _mtext(pl, "LinkLag", "0"),
                })

        return {
            "uid":           uid,
            "name":          _mtext(el, "Name"),
            "active":        _mtext(el, "Active", "1"),
            "summary":       _mtext(el, "Summary", "0"),
            "manual":        _mtext(el, "Manual", "0"),
            "milestone":     _mtext(el, "Milestone", "0"),
            "outline_level": int(_mtext(el, "OutlineLevel", "0") or 0),
            "outline_num":   _mtext(el, "OutlineNumber"),
            "start":         _mtext(el, "Start"),
            "finish":        _mtext(el, "Finish"),
            "dur_hours":     parse_iso_dur_hours(_mtext(el, "Duration")),
            "pct":           _mtext(el, "PercentComplete", "0"),
            "actual_start":  _mtext(el, "ActualStart"),
            "actual_finish": _mtext(el, "ActualFinish"),
            "constraint":    _mtext(el, "ConstraintType", "0"),
            "constraint_dt": _mtext(el, "ConstraintDate"),
            "p6_id":         p6_id,
            "pred_links":    pred_links,
        }

    def _parse_resource(self, el):
        if _mtext(el, "IsNull") == "1":
            return None
        uid = _mtext(el, "UID")
        if not uid or uid == "0":
            return None
        return {
            "uid":      uid,
            "name":     _mtext(el, "Name"),
            "type":     _mtext(el, "Type", "0"),
            "initials": _mtext(el, "Initials"),
        }

    def _parse_assignment(self, el):
        uid     = _mtext(el, "UID")
        res_uid = _mtext(el, "ResourceUID", "-1")
        try:
            if int(res_uid) < 0:
                return None
        except ValueError:
            return None
        return {
            "uid":        uid,
            "task_uid":   _mtext(el, "TaskUID"),
            "res_uid":    res_uid,
            "units":      _mtext(el, "Units", "1"),
            "work_h":     parse_iso_dur_hours(_mtext(el, "Work")),
            "act_work_h": parse_iso_dur_hours(_mtext(el, "ActualWork")),
            "rem_work_h": parse_iso_dur_hours(_mtext(el, "RemainingWork")),
            "start":      _mtext(el, "Start"),
            "finish":     _mtext(el, "Finish"),
        }

    def _build_hierarchy(self):
        self.parent_uid = {}
        self.children   = {uid: [] for uid in self.task_order}
        stack = []
        for uid in self.task_order:
            lvl = self.tasks[uid]["outline_level"]
            while stack and stack[-1][0] >= lvl:
                stack.pop()
            self.parent_uid[uid] = stack[-1][1] if stack else None
            if stack:
                self.children[stack[-1][1]].append(uid)
            stack.append((lvl, uid))

    def first_activity(self, uid, activity_set):
        if uid in activity_set:
            return uid
        for c in self.children.get(uid, []):
            r = self.first_activity(c, activity_set)
            if r:
                return r
        return None

    def last_activity(self, uid, activity_set):
        if uid in activity_set:
            return uid
        for c in reversed(self.children.get(uid, [])):
            r = self.last_activity(c, activity_set)
            if r:
                return r
        return None


def extract_data(msp, ref_els, ref_proj_fields=None):
    default_cal_oid = _ref_first_oid(ref_els, "Calendar")
    default_obs_oid = _ref_first_oid(ref_els, "OBS")
    proj_oid = "100000"
    warnings = []

    ref_proj       = ref_proj_fields or {}
    parent_eps_oid = ref_proj.get("ParentEPSObjectId") or ""
    if not parent_eps_oid:
        warnings.append("[WARN] ParentEPSObjectId not found in p6_reference.xml.")

    root_uid      = None
    summary_uids  = set()
    activity_uids = set()
    for uid in msp.task_order:
        t = msp.tasks[uid]
        if t["outline_level"] == 0:
            root_uid = uid
        elif t["summary"] == "1":
            summary_uids.add(uid)
        elif t["active"] == "1":
            activity_uids.add(uid)

    act_uid_to_aid = {}
    for uid in activity_uids:
        t = msp.tasks[uid]
        act_uid_to_aid[uid] = t["p6_id"] if t["p6_id"] else f"T{uid.zfill(4)}"

    def wbs_oid(uid_str):
        return str(int(uid_str) + WBS_OID_OFFSET)

    root_wbs_oid = wbs_oid(root_uid) if root_uid is not None else str(WBS_OID_OFFSET)

    rel_counter = [0]
    def next_rel_oid():
        rel_counter[0] += 1
        return str(200000 + rel_counter[0])

    rel_list = []
    for uid in msp.task_order:
        t = msp.tasks[uid]
        succ_uid = uid

        for pl in t["pred_links"]:
            pred_uid_raw = pl["pred_uid"]
            p6_type  = REL_TYPE_MAP.get(pl["type"], "Finish to Start")
            lag_h    = parse_lag_hours(pl["link_lag"])

            pred_uid = pred_uid_raw
            if pred_uid_raw in summary_uids:
                last = msp.last_activity(pred_uid_raw, activity_uids)
                if last:
                    warnings.append(
                        f"[WARN] Summary UID={pred_uid_raw} is predecessor of UID={succ_uid} "
                        f"→ transferred to last child activity UID={last}"
                    )
                    pred_uid = last
                else:
                    warnings.append(
                        f"[WARN] Summary UID={pred_uid_raw} has no activities; "
                        f"relationship to UID={succ_uid} skipped"
                    )
                    continue

            act_succ = succ_uid
            if succ_uid in summary_uids:
                first = msp.first_activity(succ_uid, activity_uids)
                if first:
                    warnings.append(
                        f"[WARN] Summary UID={succ_uid} is successor of UID={pred_uid_raw} "
                        f"→ transferred to first child activity UID={first}"
                    )
                    act_succ = first
                else:
                    warnings.append(
                        f"[WARN] Summary UID={succ_uid} has no activities; "
                        f"relationship from UID={pred_uid_raw} skipped"
                    )
                    continue

            if pred_uid not in activity_uids:
                warnings.append(f"[WARN] Predecessor UID={pred_uid} not an activity — skipped")
                continue
            if act_succ not in activity_uids:
                warnings.append(f"[WARN] Successor UID={act_succ} not an activity — skipped")
                continue

            rel_list.append({
                "oid":      next_rel_oid(),
                "pred_uid": pred_uid,
                "succ_uid": act_succ,
                "type":     p6_type,
                "lag":      f"{lag_h:.6f}",
            })

    wbs_list = []

    def add_wbs(uid, parent_wbs_oid):
        t    = msp.tasks[uid]
        code = t["outline_num"] or str(uid)
        wbs_list.append({
            "oid":        wbs_oid(uid),
            "parent_oid": parent_wbs_oid,
            "code":       code,
            "name":       t["name"] or "",
        })
        for c in msp.children.get(uid, []):
            if c in summary_uids:
                add_wbs(c, wbs_oid(uid))

    start_children = msp.children.get(root_uid, []) if root_uid is not None else msp.children.get(None, [])
    for c in start_children:
        if c in summary_uids:
            add_wbs(c, None)

    def find_parent_wbs(uid):
        p = msp.parent_uid.get(uid)
        while p is not None:
            if p in summary_uids:
                return wbs_oid(p)
            if p == root_uid:
                return None
            p = msp.parent_uid.get(p)
        return None

    activity_list = []
    for uid in msp.task_order:
        if uid not in activity_uids:
            continue
        t  = msp.tasks[uid]
        pc = int(t["pct"] or "0")

        if pc == 100 or t["actual_finish"]:
            status = "Completed"
        elif pc > 0 or t["actual_start"]:
            status = "In Progress"
        else:
            status = "Not Started"

        dur      = 0.0 if t["milestone"] == "1" else t["dur_hours"]
        act_type = "Start Milestone" if t["milestone"] == "1" else "Task Dependent"

        con_type = None
        con_date = None
        if t["manual"] == "1":
            con_type = "Mandatory Start"
            con_date = fmt_date(t["start"])
        elif t["constraint"] in CONSTRAINT_MAP:
            con_type = CONSTRAINT_MAP[t["constraint"]]
            con_date = fmt_date(t["constraint_dt"])

        activity_list.append({
            "uid":        uid,
            "oid":        uid,
            "id":         act_uid_to_aid[uid],
            "name":       t["name"] or "",
            "start":      fmt_date(t["start"]),
            "finish":     fmt_date(t["finish"]),
            "dur":        f"{dur:.6f}",
            "status":     status,
            "type":       act_type,
            "wbs_oid":    find_parent_wbs(uid),
            "con_type":   con_type,
            "con_date":   con_date,
            "pct":        str(pc),
            "act_start":  fmt_date(t["actual_start"]) if t["actual_start"] else None,
            "act_finish": fmt_date(t["actual_finish"]) if t["actual_finish"] else None,
        })

    resource_list = []
    res_uid_to_oid = {}
    _used_res_ids = {}
    for uid, r in msp.resources.items():
        oid = str(int(uid) + 40000)
        res_uid_to_oid[uid] = oid
        base_id = (r["initials"] or f"R{uid.zfill(4)}").strip()
        if base_id in _used_res_ids:
            _used_res_ids[base_id] += 1
            res_id = f"{base_id}{_used_res_ids[base_id]:02d}"
        else:
            _used_res_ids[base_id] = 0
            res_id = base_id
        resource_list.append({
            "oid":  oid,
            "id":   res_id,
            "name": r["name"] or "",
            "type": RESOURCE_TYPE_MAP.get(r["type"], "Labor"),
        })

    ra_list = []
    for uid, a in msp.assignments.items():
        if a["task_uid"] not in activity_uids:
            continue
        if a["res_uid"] not in res_uid_to_oid:
            continue
        ra_list.append({
            "oid":    str(int(uid) + 50000),
            "act_oid": a["task_uid"],
            "res_oid": res_uid_to_oid[a["res_uid"]],
            "units":  f"{float(a['units'] or 1):.6f}",
            "plan_u": f"{a['work_h']:.6f}",
            "act_u":  f"{a['act_work_h']:.6f}",
            "rem_u":  f"{a['rem_work_h']:.6f}",
            "start":  fmt_date(a["start"]),
            "finish": fmt_date(a["finish"]),
        })

    # OPC: ProjectId max 20 characters
    proj_name  = msp.title or msp.name or "Converted Project"
    proj_id    = re.sub(r'[^A-Za-z0-9_-]', '', proj_name)[:20] or "MSPROJ"
    plan_start = fmt_date(msp.start) or "2000-01-01T08:00:00"
    data_date  = fmt_date(msp.data_date) or plan_start

    return {
        "proj_oid":       proj_oid,
        "proj_id":        proj_id,
        "proj_name":      proj_name,
        "plan_start":     plan_start,
        "data_date":      data_date,
        "root_wbs_oid":   root_wbs_oid,
        "parent_eps_oid": parent_eps_oid,
        "default_cal":    default_cal_oid,
        "default_obs":    default_obs_oid,
        "wbs_list":       wbs_list,
        "activity_list":  activity_list,
        "rel_list":       rel_list,
        "resource_list":  resource_list,
        "ra_list":        ra_list,
        "act_uid_to_aid": act_uid_to_aid,
        "warnings":       warnings,
    }


def build_resource_el(r):
    el = ET.Element("Resource")
    make_tag(el, "AutoComputeActuals",     text="1")
    make_tag(el, "CalculateCostFromUnits", text="1")
    make_tag(el, "CalendarObjectId",       text="")
    make_tag(el, "CurrencyObjectId",       text="1")
    make_tag(el, "DefaultUnitsPerTime",    text="1.000000")
    make_tag(el, "EmailAddress",           nil=True)
    make_tag(el, "EmployeeId",             nil=True)
    make_tag(el, "GUID",                   text=new_guid())
    make_tag(el, "Id",                     text=r["id"])
    make_tag(el, "IsActive",               text="1")
    make_tag(el, "IsOverTimeAllowed",      text="0")
    make_tag(el, "Name",                   text=r["name"])
    make_tag(el, "ObjectId",               text=r["oid"])
    make_tag(el, "OfficePhone",            nil=True)
    make_tag(el, "OtherPhone",             nil=True)
    make_tag(el, "OvertimeFactor",         text="0")
    make_tag(el, "ParentObjectId",         nil=True)
    make_tag(el, "PrimaryRoleObjectId",    nil=True)
    make_tag(el, "ResourceNotes",          nil=True)
    make_tag(el, "ResourceType",           text=r["type"])
    make_tag(el, "SequenceNumber",         text="0")
    make_tag(el, "ShiftObjectId",          nil=True)
    make_tag(el, "Title",                  nil=True)
    make_tag(el, "UnitOfMeasureObjectId",  nil=True)
    make_tag(el, "UserObjectId",           nil=True)
    return el


def build_project_el(data):
    proj     = ET.Element("Project")
    proj_oid = data["proj_oid"]

    def ps(tag, val="", nil=False):
        if nil:
            make_tag(proj, tag, nil=True)
        else:
            make_tag(proj, tag, text=val)

    ps("ActivityDefaultActivityType",    "Task Dependent")
    ps("ActivityDefaultCalendarObjectId", data["default_cal"])
    ps("ActivityDefaultCostAccountObjectId", nil=True)
    ps("ActivityDefaultDurationType",    "Fixed Duration and Units/Time")
    ps("ActivityDefaultPercentCompleteType", "Duration")
    ps("ActivityDefaultPricePerUnit",    "100.00000000")
    ps("ActivityIdBasedOnSelectedActivity", "1")
    ps("ActivityIdIncrement",            "10")
    ps("ActivityIdPrefix",               "T")
    ps("ActivityIdSuffix",               "1000")
    ps("ActivityPercentCompleteBasedOnActivitySteps", "1")
    ps("AddActualToRemaining",           "0")
    ps("AddedBy",                        "admin")
    ps("AllowNegativeActualUnitsFlag",   "0")
    ps("AnnualDiscountRate",             "5.000000")
    ps("AnticipatedFinishDate",          nil=True)
    ps("AnticipatedStartDate",           nil=True)
    ps("AssignmentDefaultDrivingFlag",   "0")
    ps("AssignmentDefaultRateType",      "Price / Unit")
    ps("CheckOutStatus",                 "0")
    ps("CostQuantityRecalculateFlag",    "0")
    ps("CriticalActivityFloatLimit",     "0.00")
    ps("CriticalActivityPathType",       "Critical Float")
    ps("CurrentBaselineProjectObjectId", nil=True)
    ps("DataDate",                       data["data_date"])
    ps("DateAdded",                      datetime.now().strftime("%Y-%m-%dT%H:%M:%S"))
    ps("DefaultPriceTimeUnits",          "Hour")
    ps("DiscountApplicationPeriod",      "Month")
    ps("EarnedValueComputeType",         "Activity Percent Complete")
    ps("EarnedValueETCComputeType",      "PF = 1 / CPI")
    ps("EarnedValueETCUserValue",        "0")
    ps("EarnedValueUserPercent",         "0.0")
    ps("EnableSummarization",            "1")
    ps("FiscalYearStartMonth",           "1")
    make_tag(proj, "GUID",               text=new_guid())
    ps("Id",                             data["proj_id"])
    ps("IndependentETCLaborUnits",       "0.000000")
    ps("IndependentETCTotalCost",        "0.000000")
    ps("LastFinancialPeriodObjectId",    nil=True)
    ps("LevelingPriority",               "10")
    ps("LinkActualToActualThisPeriod",   "1")
    ps("LinkPercentCompleteWithActual",  "0")
    ps("LinkPlannedAndAtCompletionFlag", "1")
    ps("MustFinishByDate",               nil=True)
    ps("Name",                           data["proj_name"])
    ps("OBSObjectId",                    data["default_obs"])
    ps("ObjectId",                       proj_oid)
    ps("OriginalBudget",                 "0.000000")
    if data.get("parent_eps_oid"):
        ps("ParentEPSObjectId",          data["parent_eps_oid"])
    else:
        ps("ParentEPSObjectId",          nil=True)
    ps("PlannedStartDate",               data["plan_start"])
    ps("PrimaryResourcesCanMarkActivitiesAsCompleted", "1")
    ps("ProjectForecastStartDate",       nil=True)
    ps("ResetPlannedToRemainingFlag",    "0")
    ps("ResourceCanBeAssignedToSameActivityMoreThanOnce", "1")
    ps("ResourcesCanAssignThemselvesToActivities",        "1")
    ps("ScheduledFinishDate",            nil=True)
    ps("Status",                         "Active")
    ps("StrategicPriority",              "100")
    ps("SummarizeToWBSLevel",            "0")
    ps("SummaryLevel",                   "Assignment Level")
    ps("UseProjectBaselineForEarnedValue", "1")
    ps("WBSCodeSeparator",               ".")
    ps("WBSObjectId",                    data["root_wbs_oid"])
    ps("WebSiteRootDirectory",           nil=True)
    ps("WebSiteURL",                     nil=True)

    for w in data["wbs_list"]:
        wbs_el = ET.SubElement(proj, "WBS")
        make_tag(wbs_el, "AnticipatedFinishDate",     nil=True)
        make_tag(wbs_el, "AnticipatedStartDate",      nil=True)
        make_tag(wbs_el, "Code",                      text=str(w["code"]))
        make_tag(wbs_el, "EarnedValueComputeType",    text="Activity Percent Complete")
        make_tag(wbs_el, "EarnedValueETCComputeType", text="PF = 1 / CPI")
        make_tag(wbs_el, "EarnedValueETCUserValue",   text="0")
        make_tag(wbs_el, "EarnedValueUserPercent",    text="0.0")
        make_tag(wbs_el, "GUID",                      text=new_guid())
        make_tag(wbs_el, "IndependentETCLaborUnits",  text="0.000000")
        make_tag(wbs_el, "IndependentETCTotalCost",   text="0.000000")
        make_tag(wbs_el, "Name",                      text=w["name"])
        make_tag(wbs_el, "OBSObjectId",               text=data["default_obs"] or "")
        make_tag(wbs_el, "ObjectId",                  text=str(w["oid"]))
        make_tag(wbs_el, "OriginalBudget",             text="0.000000")
        if w["parent_oid"] is not None:
            make_tag(wbs_el, "ParentObjectId",        text=str(w["parent_oid"]))
        else:
            make_tag(wbs_el, "ParentObjectId",        nil=True)
        make_tag(wbs_el, "ProjectObjectId",           text=proj_oid)
        make_tag(wbs_el, "SequenceNumber",            text="0")
        make_tag(wbs_el, "Status",                    text="Active")
        make_tag(wbs_el, "WBSCategoryObjectId",       nil=True)

    for a in data["activity_list"]:
        act_el = ET.SubElement(proj, "Activity")

        def e(tag, val=None, nil_a=False, nil_o=False):
            if nil_a or tag in ACTIVITY_NIL_ALWAYS:
                make_tag(act_el, tag, nil=True)
            elif nil_o or tag in ACTIVITY_NIL_OPTIONAL:
                if val:
                    make_tag(act_el, tag, text=val)
                else:
                    make_tag(act_el, tag, nil=True)
            else:
                make_tag(act_el, tag, text=val if val is not None else "")

        e("ActualDuration",              "0")
        e("ActualFinishDate",            a["act_finish"])
        e("ActualLaborCost",             "0")
        e("ActualLaborUnits",            "0.000000")
        e("ActualNonLaborCost",          "0")
        e("ActualNonLaborUnits",         "0.000000")
        e("ActualStartDate",             a["act_start"])
        e("ActualThisPeriodLaborCost",   "0")
        e("ActualThisPeriodLaborUnits",  "0.000000")
        e("ActualThisPeriodNonLaborCost",  "0")
        e("ActualThisPeriodNonLaborUnits", "0.000000")
        e("AtCompletionDuration",        a["dur"])
        e("AtCompletionExpenseCost",     "0")
        e("AtCompletionLaborCost",       "0")
        e("AtCompletionLaborUnits",      "0.000000")
        e("AtCompletionNonLaborCost",    "0")
        e("AtCompletionNonLaborUnits",   "0.000000")
        e("AutoComputeActuals",          "1")
        e("CalendarObjectId",            data["default_cal"])
        e("DurationPercentComplete",     a["pct"])
        e("DurationType",                "Fixed Duration and Units/Time")
        e("EstimatedWeight",             "1.00")
        e("ExpectedFinishDate",          None)
        e("ExternalEarlyStartDate",      None)
        e("ExternalLateFinishDate",      None)
        e("Feedback",                    "")
        e("FinishDate",                  a["finish"])
        make_tag(act_el, "GUID",         text=new_guid())
        e("Id",                          a["id"])
        e("IsNewFeedback",               "0")
        e("LevelingPriority",            "Normal")
        e("Name",                        a["name"])
        e("NonLaborUnitsPercentComplete", "0")
        e("NotesToResources",            "")
        e("ObjectId",                    a["oid"])
        e("PercentComplete",             a["pct"])
        e("PercentCompleteType",         "Physical")
        e("PhysicalPercentComplete",     a["pct"])
        e("PlannedDuration",             a["dur"])
        e("PlannedFinishDate",           a["finish"])
        e("PlannedLaborCost",            "0")
        e("PlannedLaborUnits",           "0.000000")
        e("PlannedNonLaborCost",         "0")
        e("PlannedNonLaborUnits",        "0.000000")
        e("PlannedStartDate",            a["start"])
        e("PrimaryConstraintDate",       a["con_date"])
        e("PrimaryConstraintType",       a["con_type"])
        e("PrimaryResourceObjectId",     None)
        e("ProjectObjectId",             proj_oid)
        e("RemainingDuration",           a["dur"])
        e("RemainingEarlyFinishDate",    a["finish"])
        e("RemainingEarlyStartDate",     a["start"])
        e("RemainingLaborCost",          "0")
        e("RemainingLaborUnits",         "0.000000")
        e("RemainingLateFinishDate",     a["finish"])
        e("RemainingLateStartDate",      a["start"])
        e("RemainingNonLaborCost",       "0")
        e("RemainingNonLaborUnits",      "0.000000")
        e("ResumeDate",                  None)
        e("SecondaryConstraintDate",     None)
        e("SecondaryConstraintType",     None)
        e("StartDate",                   a["start"])
        e("Status",                      a["status"])
        e("SuspendDate",                 None)
        e("Type",                        a["type"])
        e("UnitsPercentComplete",        "0")
        e("WBSObjectId",                 str(a["wbs_oid"]) if a["wbs_oid"] else None)

    for ra in data["ra_list"]:
        ra_el = ET.SubElement(proj, "ResourceAssignment")
        make_tag(ra_el, "ActivityObjectId",        text=ra["act_oid"])
        make_tag(ra_el, "ActualCost",              text="0.000000")
        make_tag(ra_el, "ActualCurve",             nil=True)
        make_tag(ra_el, "ActualFinishDate",        nil=True)
        make_tag(ra_el, "ActualOvertimeCost",      text="0.000000")
        make_tag(ra_el, "ActualOvertimeUnits",     text="0.000000")
        make_tag(ra_el, "ActualRegularCost",       text="0.000000")
        make_tag(ra_el, "ActualRegularUnits",      text="0.000000")
        make_tag(ra_el, "ActualStartDate",         nil=True)
        make_tag(ra_el, "ActualThisPeriodCost",    text="0.000000")
        make_tag(ra_el, "ActualThisPeriodUnits",   text="0.000000")
        make_tag(ra_el, "ActualUnits",             text=ra["act_u"])
        make_tag(ra_el, "AtCompletionCost",        text="0.000000")
        make_tag(ra_el, "AtCompletionUnits",       text=ra["plan_u"])
        make_tag(ra_el, "CostAccountObjectId",     nil=True)
        make_tag(ra_el, "DrivingActivityDatesFlag", text="1")
        if ra["finish"]:
            make_tag(ra_el, "FinishDate",          text=ra["finish"])
        else:
            make_tag(ra_el, "FinishDate",          nil=True)
        make_tag(ra_el, "GUID",                    text=new_guid())
        make_tag(ra_el, "IsCostUnitsLinked",       text="1")
        make_tag(ra_el, "IsPrimaryResource",       text="0")
        make_tag(ra_el, "ObjectId",                text=ra["oid"])
        make_tag(ra_el, "OvertimeFactor",          text="0")
        make_tag(ra_el, "PlannedCost",             text="0.000000")
        make_tag(ra_el, "PlannedCurve",            nil=True)
        if ra["finish"]:
            make_tag(ra_el, "PlannedFinishDate",   text=ra["finish"])
        else:
            make_tag(ra_el, "PlannedFinishDate",   nil=True)
        make_tag(ra_el, "PlannedLag",              text="0.000000")
        if ra["start"]:
            make_tag(ra_el, "PlannedStartDate",    text=ra["start"])
        else:
            make_tag(ra_el, "PlannedStartDate",    nil=True)
        make_tag(ra_el, "PlannedUnits",            text=ra["plan_u"])
        make_tag(ra_el, "PlannedUnitsPerTime",     text=ra["units"])
        make_tag(ra_el, "Proficiency",             text="3 - Skilled")
        make_tag(ra_el, "ProjectObjectId",         text=data["proj_oid"])
        make_tag(ra_el, "RateSource",              text="Resource")
        make_tag(ra_el, "RateType",                text="Price / Unit")
        make_tag(ra_el, "RemainingCost",           text="0.000000")
        make_tag(ra_el, "RemainingCurve",          nil=True)
        make_tag(ra_el, "RemainingDuration",       text="0.000000")
        if ra["finish"]:
            make_tag(ra_el, "RemainingFinishDate", text=ra["finish"])
        else:
            make_tag(ra_el, "RemainingFinishDate", nil=True)
        make_tag(ra_el, "RemainingLag",            text="0.000000")
        if ra["start"]:
            make_tag(ra_el, "RemainingStartDate",  text=ra["start"])
        else:
            make_tag(ra_el, "RemainingStartDate",  nil=True)
        make_tag(ra_el, "RemainingUnits",          text=ra["rem_u"])
        make_tag(ra_el, "RemainingUnitsPerTime",   text=ra["units"])
        make_tag(ra_el, "ResourceCurveObjectId",   nil=True)
        make_tag(ra_el, "ResourceObjectId",        text=ra["res_oid"])
        make_tag(ra_el, "ResourceType",            text="Labor")
        make_tag(ra_el, "RoleObjectId",            nil=True)
        if ra["start"]:
            make_tag(ra_el, "StartDate",           text=ra["start"])
        else:
            make_tag(ra_el, "StartDate",           nil=True)
        make_tag(ra_el, "UnitsPercentComplete",    text="0")
        make_tag(ra_el, "WBSObjectId",             nil=True)

    aid = data["act_uid_to_aid"]
    for r in data["rel_list"]:
        rel_el = ET.SubElement(proj, "Relationship")
        make_tag(rel_el, "Lag",                         text=r["lag"])
        make_tag(rel_el, "ObjectId",                    text=r["oid"])
        make_tag(rel_el, "PredecessorActivityId",       text=aid.get(r["pred_uid"], r["pred_uid"]))
        make_tag(rel_el, "PredecessorActivityObjectId", text=r["pred_uid"])
        make_tag(rel_el, "PredecessorProjectObjectId",  text=data["proj_oid"])
        make_tag(rel_el, "SuccessorActivityId",         text=aid.get(r["succ_uid"], r["succ_uid"]))
        make_tag(rel_el, "SuccessorActivityObjectId",   text=r["succ_uid"])
        make_tag(rel_el, "SuccessorProjectObjectId",    text=data["proj_oid"])
        make_tag(rel_el, "Type",                        text=r["type"])

    return proj


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

    r2   = ET.parse(path).getroot()
    p_el = r2.find(f"{{{NS_BO}}}Project")
    def cnt(tag):  return len(r2.findall(f"{{{NS_BO}}}{tag}"))
    def pcnt(tag): return len(p_el.findall(f"{{{NS_BO}}}{tag}")) if p_el is not None else 0
    print(f"  Written: {path}")
    print(f"    Global — Calendar:{cnt('Calendar')} OBS:{cnt('OBS')} "
          f"Currency:{cnt('Currency')} Role:{cnt('Role')}")
    print(f"    Top-level — Resource:{cnt('Resource')}")
    print(f"    In Project — WBS:{pcnt('WBS')} Activity:{pcnt('Activity')} "
          f"Relationship:{pcnt('Relationship')} ResourceAssignment:{pcnt('ResourceAssignment')}")


def main():
    if len(sys.argv) < 2:
        print("Usage: py msp_to_opc_converter.py <input_msp.xml> [output_dir]")
        sys.exit(1)

    msp_path   = sys.argv[1]
    out_dir    = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(msp_path))
    script_dir = os.path.dirname(os.path.abspath(__file__))
    ref_path   = os.path.join(script_dir, "p6_reference.xml")

    base     = os.path.splitext(os.path.basename(msp_path))[0]
    opc_path = os.path.join(out_dir, f"{base}_OPC.xml")

    msp = MSPProject(msp_path)

    print("Loading p6_reference.xml ...")
    ref_els = load_reference_elements(ref_path,
                                      "Currency", "UDFType", "OBS",
                                      "Calendar", "Role", "RoleRate")
    ref_proj_flds = load_reference_project_fields(ref_path)
    if ref_proj_flds.get("ParentEPSObjectId"):
        print(f"  ParentEPSObjectId: {ref_proj_flds['ParentEPSObjectId']}")
    else:
        print("  WARNING: ParentEPSObjectId not found in p6_reference.xml")

    print("Converting ...")
    warnings_list = []
    filter_udf_types(ref_els, warnings_list)

    data = extract_data(msp, ref_els, ref_proj_fields=ref_proj_flds)
    data["warnings"] = warnings_list + data["warnings"]

    if data["warnings"]:
        print()
        for w in data["warnings"]:
            print(" ", w)
        print()

    ET.register_namespace("",    NS_BO)
    ET.register_namespace("xsi", NS_XSI)

    def add_ns(el):
        if not el.tag.startswith("{"):
            el.tag = f"{{{NS_BO}}}{el.tag}"
        for child in el:
            add_ns(child)

    proj_el = build_project_el(data)
    proj_el.tag = f"{{{NS_BO}}}Project"
    add_ns(proj_el)

    root = ET.Element(
        f"{{{NS_BO}}}APIBusinessObjects",
        attrib={f"{{{NS_XSI}}}schemaLocation": SCHEMA_LOC}
    )
    for tag in ("Currency", "UDFType", "OBS", "Calendar", "Role", "RoleRate"):
        for el in ref_els.get(tag, []):
            root.append(el)
    for r in data["resource_list"]:
        res_el = build_resource_el(r)
        res_el.tag = f"{{{NS_BO}}}Resource"
        add_ns(res_el)
        root.append(res_el)
    root.append(proj_el)

    print("\nWriting OPC XML ...")
    write_opc_xml(root, opc_path)

    print(f"\nSummary: {len(data['wbs_list'])} WBS  |  {len(data['activity_list'])} Activities  "
          f"|  {len(data['rel_list'])} Relationships  |  {len(data['ra_list'])} ResourceAssignments")
    if data["warnings"]:
        print(f"         {len(data['warnings'])} warnings (see above)")

    print(f"\nOutput: {opc_path}")
    print("\n[WARN] After OPC import: reschedule project and recalculate costs manually in OPC")

    return opc_path


if __name__ == "__main__":
    main()
