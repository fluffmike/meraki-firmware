#!/usr/bin/env python3
"""
Meraki inventory + firmware scheduling tool.

Modes
-----
  report    Read-only. Collects every device in every org/network the API key can
            see, writes a dated Devices / Changes / Summary sheet set into the
            inventory workbook (one workbook, new sheets per run).
  plan      Read-only. Builds firmware_plan.xlsx: one row per upgrade target
            (config template or standalone network) and product type, with the
            versions Meraki says are available for it. You mark rows with "Y",
            pick a target version and a local date/time.
  schedule  Reads the plan and schedules the selected upgrades. DRY-RUN by
            default; nothing changes unless --apply is given (and confirmed).

Settings come from .env next to this script (see .env.example).
"""

import argparse
import csv
import logging
import os
import re
import sys
import tempfile
from collections import Counter, defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

try:
    import meraki
    from dotenv import load_dotenv
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError as exc:
    sys.exit(f"Missing dependency ({exc}). Run: pip install -r requirements.txt")

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")


def _path(env_name, default):
    p = Path(os.getenv(env_name) or default).expanduser()
    return p if p.is_absolute() else BASE_DIR / p


API_KEY = os.getenv("MERAKI_API_KEY", "").strip()
OUTPUT_FILE = _path("OUTPUT_FILE", "output/meraki_inventory.xlsx")
PLAN_FILE = _path("PLAN_FILE", "output/firmware_plan.xlsx")
EXCLUSIONS_FILE = _path("EXCLUSIONS_FILE", "exclusions.txt")
LOG_DIR = _path("LOG_DIR", "logs")
KEEP_RUNS = int(os.getenv("KEEP_RUNS", "30"))
CSV_EXPORT = os.getenv("CSV_EXPORT", "true").strip().lower() in ("1", "true", "yes", "y")
MIN_LEAD_MINUTES = int(os.getenv("MIN_LEAD_MINUTES", "15"))
# "template": networks bound to a config template are scheduled once, on the template.
# "network":  every network is scheduled individually, bound or not.
TEMPLATE_FIRMWARE_MODE = os.getenv("TEMPLATE_FIRMWARE_MODE", "template").strip().lower()

SHEET_RE = re.compile(r"^(Devices|Changes|Summary)_(\d{4}-\d{2}-\d{2}_\d{4}(?:\d{2})?)$")
SCHEDULED_SHEET = "Scheduled_Upgrades"
YES_VALUES = {"y", "yes", "x", "1", "true"}

log = logging.getLogger("meraki_tool")


# --------------------------------------------------------------------------- #
# Setup helpers
# --------------------------------------------------------------------------- #
def setup_logging():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s %(levelname)-7s %(message)s")
    fh = logging.FileHandler(LOG_DIR / f"meraki_tool_{datetime.now():%Y%m%d}.log")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler()
    ch.setFormatter(logging.Formatter("%(levelname)-7s %(message)s"))
    log.addHandler(fh)
    log.addHandler(ch)


def get_local_tz():
    """LOCAL_TZ from .env, else the system's IANA zone, else a fixed offset."""
    name = os.getenv("LOCAL_TZ", "").strip()
    if name:
        try:
            return ZoneInfo(name)
        except ZoneInfoNotFoundError:
            sys.exit(f"LOCAL_TZ '{name}' is not a valid IANA time zone (e.g. America/Chicago).")
    real = os.path.realpath("/etc/localtime")
    if "zoneinfo/" in real:
        try:
            return ZoneInfo(real.split("zoneinfo/", 1)[1])
        except ZoneInfoNotFoundError:
            pass
    log.warning("Could not detect an IANA time zone; set LOCAL_TZ in .env for DST-correct times.")
    return datetime.now().astimezone().tzinfo


LOCAL_TZ = None  # set in main()


def get_dashboard():
    if not API_KEY:
        sys.exit("MERAKI_API_KEY is not set. Copy .env.example to .env and add your key.")
    return meraki.DashboardAPI(
        API_KEY,
        suppress_logging=True,
        wait_on_rate_limit=True,
        maximum_retries=5,
        retry_4xx_error=False,
    )


def api_error_text(exc):
    if isinstance(exc, meraki.APIError):
        return f"{exc.status} {exc.message}"
    return str(exc)


# --------------------------------------------------------------------------- #
# Small formatting helpers
# --------------------------------------------------------------------------- #
def parse_iso(value):
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None


def fmt_local(value):
    dt = value if isinstance(value, datetime) else parse_iso(value)
    if not dt:
        return ""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(LOCAL_TZ).strftime("%Y-%m-%d %H:%M %Z")


def version_label(v):
    if not v:
        return ""
    return v.get("shortName") or v.get("firmware") or str(v.get("id", ""))


def joined(values):
    return ", ".join(str(v) for v in values if v not in (None, ""))


def fw_product_for(device, products):
    """Map a device to the key used in the firmwareUpgrades 'products' object."""
    pt = device.get("productType") or ""
    model = (device.get("model") or "").upper()
    candidates = []
    if pt == "switch" and model.startswith("C9"):
        candidates.append("switchCatalyst")
    if pt == "wireless" and model.startswith("C9800"):
        candidates.append("wirelessController")
    candidates.append(pt)
    for c in candidates:
        if c in products:
            return c
    return candidates[0]


def latest_stable(available):
    stable = [v for v in available or [] if (v.get("releaseType") or "").lower() == "stable"]
    stable.sort(key=lambda v: v.get("releaseDate") or "")
    return stable[-1] if stable else None


# --------------------------------------------------------------------------- #
# Exclusions
# --------------------------------------------------------------------------- #
def load_exclusions():
    entries = set()
    if EXCLUSIONS_FILE.exists():
        for line in EXCLUSIONS_FILE.read_text().splitlines():
            line = line.split("#", 1)[0].strip()
            if line:
                entries.add(line.lower())
    return entries


def exclusion_reason(exclusions, **keys):
    """keys: label -> value or list of values. Returns 'label=value' of first match."""
    for label, values in keys.items():
        if not isinstance(values, (list, tuple, set)):
            values = [values]
        for v in values:
            if v and str(v).lower() in exclusions:
                return f"{label}={v}"
    return ""


# --------------------------------------------------------------------------- #
# Data collection
# --------------------------------------------------------------------------- #
class Collector:
    """Pulls everything once per run and caches firmware lookups."""

    def __init__(self, dashboard):
        self.d = dashboard
        self.errors = []
        self._fw_cache = {}

    def firmware(self, network_or_template_id):
        if network_or_template_id not in self._fw_cache:
            try:
                self._fw_cache[network_or_template_id] = self.d.networks.getNetworkFirmwareUpgrades(
                    network_or_template_id)
            except Exception as exc:  # noqa: BLE001 - keep going, record the error
                self._fw_cache[network_or_template_id] = exc
        result = self._fw_cache[network_or_template_id]
        return result if not isinstance(result, Exception) else None

    def firmware_error(self, network_or_template_id):
        result = self._fw_cache.get(network_or_template_id)
        return api_error_text(result) if isinstance(result, Exception) else ""

    def _call(self, org, what, fn, *args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except Exception as exc:  # noqa: BLE001
            msg = f"{org['name']} ({org['id']}): {what} failed: {api_error_text(exc)}"
            log.warning(msg)
            self.errors.append(msg)
            return None

    def collect(self, with_firmware=True):
        orgs = self.d.organizations.getOrganizations(total_pages="all")
        log.info("Found %d organization(s)", len(orgs))
        data = []
        for org in sorted(orgs, key=lambda o: o.get("name", "")):
            oid = org["id"]
            log.info("Collecting org '%s' (%s)", org.get("name"), oid)
            networks = self._call(org, "networks", self.d.organizations.getOrganizationNetworks,
                                  oid, total_pages="all")
            if networks is None:  # API access disabled for this org, etc.
                continue
            templates = self._call(org, "config templates",
                                   self.d.organizations.getOrganizationConfigTemplates, oid) or []
            devices = self._call(org, "devices", self.d.organizations.getOrganizationDevices,
                                 oid, total_pages="all") or []
            statuses = self._call(org, "device statuses",
                                  self.d.organizations.getOrganizationDevicesStatuses,
                                  oid, total_pages="all") or []
            unassigned = self._call(org, "unassigned inventory",
                                    self.d.organizations.getOrganizationInventoryDevices,
                                    oid, total_pages="all", usedState="unused") or []
            upgrades = []
            if with_firmware:
                upgrades = self._call(org, "firmware upgrades by device",
                                      self.d.organizations.getOrganizationFirmwareUpgradesByDevice,
                                      oid, total_pages="all", perPage=1000) or []
                for nid in {d.get("networkId") for d in devices if d.get("networkId")}:
                    if self.firmware(nid) is None:
                        self.errors.append(f"{org['name']}: firmware for network {nid} failed: "
                                           f"{self.firmware_error(nid)}")
            data.append({
                "org": org,
                "networks": {n["id"]: n for n in networks},
                "templates": {t["id"]: t for t in templates},
                "devices": devices,
                "statuses": {s.get("serial"): s for s in statuses},
                "unassigned": unassigned,
                "upgrades": group_upgrades(upgrades),
            })
            log.info("  %d network(s), %d template(s), %d device(s), %d unassigned",
                     len(networks), len(templates), len(devices), len(unassigned))
        return data


def group_upgrades(entries):
    """serial -> {'pending': upgrade or None, 'last': upgrade or None}"""
    done_states = {"completed", "canceled", "cancelled", "failed", "skipped"}
    out = defaultdict(lambda: {"pending": None, "last": None})
    for e in entries:
        serial = e.get("serial")
        items = e.get("upgrades") or ([e["upgrade"]] if e.get("upgrade") else [])
        for up in items:
            status = (up.get("status") or "").lower()
            slot = "last" if status in done_states else "pending"
            cur = out[serial][slot]
            if cur is None or (up.get("time") or "") > (cur.get("time") or ""):
                out[serial][slot] = up
    return out


# --------------------------------------------------------------------------- #
# REPORT
# --------------------------------------------------------------------------- #
DEVICE_HEADERS = [
    "Org Name", "Org ID", "Network Name", "Network ID", "Network Tags", "Network Time Zone",
    "Bound To Template", "Template Name", "Template ID", "Switch Profile ID",
    "Device Name", "Serial", "MAC", "Model", "Product Type", "Status", "Last Reported",
    "Public IP", "LAN IP", "WAN1 IP", "Gateway", "Tags", "Address", "Notes",
    "Device Firmware", "Network Firmware Version", "Latest Stable Available", "Upgrade Available",
    "Scheduled Upgrade Time", "Scheduled To Version", "Scheduled Upgrade Status",
    "Scheduled Source", "Last Upgrade Time", "Last Upgrade To Version", "Last Upgrade Status",
    "Excluded From Upgrades", "Dashboard URL",
]

CHANGE_FIELDS = [
    "Org Name", "Network Name", "Template Name", "Device Name", "Model", "Status",
    "Device Firmware", "Network Firmware Version", "Scheduled Upgrade Time",
    "Scheduled To Version", "LAN IP", "Public IP", "Tags",
]


def build_device_rows(collector, data, exclusions):
    rows = []
    for o in data:
        org = o["org"]
        for dev in o["devices"]:
            net = o["networks"].get(dev.get("networkId"), {})
            tpl_id = net.get("configTemplateId") if net.get("isBoundToConfigTemplate") else None
            tpl = o["templates"].get(tpl_id, {}) if tpl_id else {}
            st = o["statuses"].get(dev.get("serial"), {})
            fw = collector.firmware(net["id"]) if net.get("id") else None
            products = (fw or {}).get("products") or {}
            prod = products.get(fw_product_for(dev, products)) or {}
            current = prod.get("currentVersion") or {}
            newest = latest_stable(prod.get("availableVersions"))
            upgrade_avail = ""
            if newest and current:
                upgrade_avail = "Yes" if (newest.get("releaseDate") or "") > (current.get("releaseDate") or "") else "No"

            ups = o["upgrades"].get(dev.get("serial"), {"pending": None, "last": None})
            pending, last = ups["pending"], ups["last"]
            nxt = prod.get("nextUpgrade") or {}
            if pending:
                sched_time, sched_ver = pending.get("time"), version_label(pending.get("toVersion"))
                sched_status, sched_src = pending.get("status") or "", "device"
            elif nxt.get("time"):
                sched_time, sched_ver = nxt.get("time"), version_label(nxt.get("toVersion"))
                sched_status, sched_src = "scheduled", "template" if tpl_id else "network"
            else:
                sched_time = sched_ver = sched_status = sched_src = ""
            if not last and prod.get("lastUpgrade"):
                lu = prod["lastUpgrade"]
                last = {"time": lu.get("time"), "toVersion": lu.get("toVersion"), "status": ""}

            excluded = exclusion_reason(
                exclusions, org=[org["id"], org.get("name")],
                network=[net.get("id"), net.get("name")],
                template=[tpl_id, tpl.get("name")],
                device=[dev.get("serial"), dev.get("name")])

            rows.append({
                "Org Name": org.get("name"), "Org ID": org["id"],
                "Network Name": net.get("name", ""), "Network ID": dev.get("networkId", ""),
                "Network Tags": joined(net.get("tags") or []),
                "Network Time Zone": net.get("timeZone", ""),
                "Bound To Template": "Yes" if tpl_id else "No",
                "Template Name": tpl.get("name", tpl_id or ""), "Template ID": tpl_id or "",
                "Switch Profile ID": dev.get("switchProfileId") or "",
                "Device Name": dev.get("name") or "", "Serial": dev.get("serial"),
                "MAC": dev.get("mac"), "Model": dev.get("model"),
                "Product Type": dev.get("productType"), "Status": st.get("status", ""),
                "Last Reported": fmt_local(st.get("lastReportedAt")),
                "Public IP": st.get("publicIp") or "", "LAN IP": st.get("lanIp") or dev.get("lanIp") or "",
                "WAN1 IP": st.get("wan1Ip") or "", "Gateway": st.get("gateway") or "",
                "Tags": joined(dev.get("tags") or []), "Address": dev.get("address") or "",
                "Notes": dev.get("notes") or "",
                "Device Firmware": dev.get("firmware") or "",
                "Network Firmware Version": version_label(current),
                "Latest Stable Available": version_label(newest), "Upgrade Available": upgrade_avail,
                "Scheduled Upgrade Time": fmt_local(sched_time), "Scheduled To Version": sched_ver,
                "Scheduled Upgrade Status": sched_status, "Scheduled Source": sched_src,
                "Last Upgrade Time": fmt_local((last or {}).get("time")),
                "Last Upgrade To Version": version_label((last or {}).get("toVersion")),
                "Last Upgrade Status": (last or {}).get("status") or "",
                "Excluded From Upgrades": excluded, "Dashboard URL": dev.get("url") or "",
            })
        for inv in o["unassigned"]:
            rows.append({h: "" for h in DEVICE_HEADERS} | {
                "Org Name": org.get("name"), "Org ID": org["id"],
                "Network Name": "(unassigned)", "Bound To Template": "No",
                "Device Name": inv.get("name") or "", "Serial": inv.get("serial"),
                "MAC": inv.get("mac"), "Model": inv.get("model"),
                "Product Type": inv.get("productType"), "Status": "unassigned",
                "Notes": f"Claimed {fmt_local(inv.get('claimedAt'))}" if inv.get("claimedAt") else "",
            })
    rows.sort(key=lambda r: (str(r["Org Name"]), str(r["Network Name"]), str(r["Device Name"]), str(r["Serial"])))
    return rows


def build_summary_rows(device_rows, data, errors):
    rows = [("Totals", "Organizations", len(data)),
             ("Totals", "Networks", sum(len(o["networks"]) for o in data)),
             ("Totals", "Config templates", sum(len(o["templates"]) for o in data)),
             ("Totals", "Devices (incl. unassigned)", len(device_rows)),
             ("Totals", "Devices bound to a template",
              sum(1 for r in device_rows if r["Bound To Template"] == "Yes")),
             ("Totals", "Upgrades scheduled", sum(1 for r in device_rows if r["Scheduled Upgrade Time"])),
             ("Totals", "Newer stable firmware available",
              sum(1 for r in device_rows if r["Upgrade Available"] == "Yes")),
             ("Totals", "Excluded from upgrades", sum(1 for r in device_rows if r["Excluded From Upgrades"]))]
    for section, key in [("By Status", "Status"), ("By Org", "Org Name"), ("By Product Type", "Product Type"),
                         ("By Model", "Model"), ("By Network Firmware", "Network Firmware Version"),
                         ("By Template", "Template Name")]:
        counts = Counter((r[key] or "(none)") for r in device_rows)
        rows += [(section, k, v) for k, v in sorted(counts.items(), key=lambda kv: (-kv[1], str(kv[0])))]
    rows += [("Errors", e, "") for e in errors]
    return rows


def previous_devices(wb):
    stamps = sorted((m.group(2) for m in map(SHEET_RE.match, wb.sheetnames) if m and m.group(1) == "Devices"),
                    reverse=True)
    if not stamps:
        return None, {}
    ws = wb[f"Devices_{stamps[0]}"]
    it = ws.iter_rows(values_only=True)
    headers = list(next(it, []))
    prev = {}
    for row in it:
        rec = dict(zip(headers, row))
        if rec.get("Serial"):
            prev[rec["Serial"]] = rec
    return stamps[0], prev


def build_change_rows(prev_stamp, prev, device_rows):
    if prev_stamp is None:
        return [("Info", "", "", "", "", "First run - nothing to compare against")]
    norm = lambda v: "" if v is None else str(v)  # noqa: E731
    cur = {r["Serial"]: r for r in device_rows}
    rows = []
    for serial in sorted(cur.keys() - prev.keys()):
        r = cur[serial]
        rows.append(("Added", serial, r["Device Name"], r["Network Name"], "", ""))
    for serial in sorted(prev.keys() - cur.keys()):
        r = prev[serial]
        rows.append(("Removed", serial, norm(r.get("Device Name")), norm(r.get("Network Name")), "", ""))
    for serial in sorted(cur.keys() & prev.keys()):
        for f in CHANGE_FIELDS:
            old, new = norm(prev[serial].get(f)), norm(cur[serial].get(f))
            if old != new:
                rows.append(("Changed", serial, cur[serial]["Device Name"], f, old, new))
    if not rows:
        rows.append(("Info", "", "", "", "", f"No changes since {prev_stamp}"))
    return rows


def style_sheet(ws, header_count, text_cols=()):
    fill = PatternFill("solid", fgColor="1F4E78")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill
        c.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(header_count)}{max(ws.max_row, 1)}"
    for idx in range(1, header_count + 1):
        letter = get_column_letter(idx)
        width = max((len(str(c.value)) for c in ws[letter] if c.value is not None), default=8)
        ws.column_dimensions[letter].width = min(max(width + 2, 10), 60)
    for letter in text_cols:
        for c in ws[letter][1:]:
            c.number_format = "@"


def write_sheet(wb, title, headers, rows, index=None):
    ws = wb.create_sheet(title, index)
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers] if isinstance(r, dict) else list(r))
    style_sheet(ws, len(headers))
    return ws


def open_workbook(path):
    if path.exists():
        return load_workbook(path)
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def save_workbook(wb, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(fd)
    wb.save(tmp)
    try:
        os.replace(tmp, path)
        return path
    except PermissionError:  # usually: file open in Excel (WSL /mnt/c)
        alt = path.with_name(f"{path.stem}_{datetime.now():%Y%m%d_%H%M%S}{path.suffix}")
        os.replace(tmp, alt)
        log.warning("%s is locked (open in Excel?). Saved to %s instead.", path, alt)
        return alt


def prune_runs(wb):
    stamps = sorted({m.group(2) for m in map(SHEET_RE.match, wb.sheetnames) if m}, reverse=True)
    for stamp in stamps[KEEP_RUNS:]:
        for prefix in ("Devices", "Changes", "Summary"):
            name = f"{prefix}_{stamp}"
            if name in wb.sheetnames:
                wb.remove(wb[name])
                log.info("Retention: removed sheet %s", name)


def run_stamp(wb):
    now = datetime.now(LOCAL_TZ)
    stamp = now.strftime("%Y-%m-%d_%H%M")
    if f"Devices_{stamp}" in wb.sheetnames:
        stamp = now.strftime("%Y-%m-%d_%H%M%S")
    return stamp


def cmd_report(args):
    collector = Collector(get_dashboard())
    data = collector.collect()
    exclusions = load_exclusions()
    device_rows = build_device_rows(collector, data, exclusions)

    wb = open_workbook(OUTPUT_FILE)
    prev_stamp, prev = previous_devices(wb)
    stamp = run_stamp(wb)

    write_sheet(wb, f"Summary_{stamp}", ["Section", "Item", "Count"],
                build_summary_rows(device_rows, data, collector.errors), index=0)
    write_sheet(wb, f"Changes_{stamp}",
                ["Change", "Serial", "Device Name", "Field / Network", "Old Value", "New Value"],
                build_change_rows(prev_stamp, prev, device_rows), index=0)
    write_sheet(wb, f"Devices_{stamp}", DEVICE_HEADERS, device_rows, index=0)
    if SCHEDULED_SHEET in wb.sheetnames:  # keep the persistent log at the end
        wb.move_sheet(SCHEDULED_SHEET, offset=len(wb.sheetnames) - 1 - wb.sheetnames.index(SCHEDULED_SHEET))
    prune_runs(wb)
    wb.active = 0
    saved = save_workbook(wb, OUTPUT_FILE)
    log.info("Wrote %d device rows to %s (sheet Devices_%s)", len(device_rows), saved, stamp)

    if CSV_EXPORT:
        csv_path = saved.parent / f"devices_{stamp}.csv"
        with open(csv_path, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=DEVICE_HEADERS)
            w.writeheader()
            w.writerows(device_rows)
        log.info("CSV copy: %s", csv_path)
    if collector.errors:
        log.warning("%d error(s) during collection - see Summary sheet 'Errors' section.", len(collector.errors))


# --------------------------------------------------------------------------- #
# PLAN
# --------------------------------------------------------------------------- #
PLAN_HEADERS = [
    "Select (Y)", "Target Type", "Org Name", "Org ID", "Target Name", "Target ID", "Product",
    "Device Count", "Models", "Bound Networks", "Current Version", "Currently Scheduled",
    "Latest Stable", "Available Versions", "Target Version", "Upgrade Time Local (YYYY-MM-DD HH:MM)",
    "Excluded", "Notes",
]
PLAN_DEVICE_HEADERS = ["Org Name", "Target Type", "Target Name", "Target ID", "Product",
                       "Network Name", "Serial", "Device Name", "Model", "Device Firmware"]
VERSION_HEADERS = ["Org Name", "Target Type", "Target Name", "Target ID", "Product", "Version ID",
                   "Short Name", "Firmware", "Release Type", "Release Date"]


def plan_targets(collector, o):
    """Yield (type, id, name, fw_response, [networks], note)."""
    devices_by_net = defaultdict(list)
    for d in o["devices"]:
        if d.get("networkId"):
            devices_by_net[d["networkId"]].append(d)
    bound = defaultdict(list)
    standalone = []
    for nid, net in o["networks"].items():
        if nid not in devices_by_net:
            continue
        tpl_id = net.get("configTemplateId") if net.get("isBoundToConfigTemplate") else None
        if tpl_id and TEMPLATE_FIRMWARE_MODE == "template":
            bound[tpl_id].append(net)
        else:
            standalone.append((net, "Bound to template; scheduled per network" if tpl_id else ""))

    for tpl_id, nets in bound.items():
        tpl = o["templates"].get(tpl_id, {"name": tpl_id})
        fw = collector.firmware(tpl_id)
        if fw is None:
            note = f"Template-level firmware API failed ({collector.firmware_error(tpl_id)}); listed per network"
            log.warning("Template '%s': %s", tpl.get("name"), note)
            standalone += [(n, note) for n in nets]
            continue
        yield "template", tpl_id, tpl.get("name"), fw, nets, ""
    for net, note in standalone:
        yield "network", net["id"], net.get("name"), collector.firmware(net["id"]), [net], note


def cmd_plan(args):
    collector = Collector(get_dashboard())
    data = collector.collect()
    exclusions = load_exclusions()
    plan_rows, dev_rows, ver_rows = [], [], []

    for o in data:
        org = o["org"]
        devices_by_net = defaultdict(list)
        for d in o["devices"]:
            devices_by_net[d.get("networkId")].append(d)
        for ttype, tid, tname, fw, nets, note in plan_targets(collector, o):
            if fw is None:
                plan_rows.append({"Target Type": ttype, "Org Name": org["name"], "Org ID": org["id"],
                                  "Target Name": tname, "Target ID": tid,
                                  "Notes": f"Firmware info unavailable: {collector.firmware_error(tid)}"})
                continue
            products = fw.get("products") or {}
            by_product = defaultdict(list)
            for net in nets:
                for d in devices_by_net[net["id"]]:
                    by_product[fw_product_for(d, products)].append((net, d))
            for product, members in sorted(by_product.items()):
                info = products.get(product)
                if not info:
                    continue
                available = info.get("availableVersions") or []
                current = info.get("currentVersion") or {}
                nxt = info.get("nextUpgrade") or {}
                devs = [d for _, d in members]
                excluded = exclusion_reason(
                    exclusions, org=[org["id"], org.get("name")],
                    target=[tid, tname],
                    network=[x for n in nets for x in (n["id"], n.get("name"))],
                    device=[x for d in devs for x in (d.get("serial"), d.get("name"))])
                plan_rows.append({
                    "Select (Y)": "", "Target Type": ttype, "Org Name": org["name"], "Org ID": org["id"],
                    "Target Name": tname, "Target ID": tid, "Product": product,
                    "Device Count": len(devs), "Models": joined(sorted({d.get("model") for d in devs})),
                    "Bound Networks": joined(n.get("name") for n in nets) if ttype == "template" else "",
                    "Current Version": version_label(current),
                    "Currently Scheduled": (f"{fmt_local(nxt.get('time'))} -> {version_label(nxt.get('toVersion'))}"
                                            if nxt.get("time") else ""),
                    "Latest Stable": version_label(latest_stable(available)),
                    "Available Versions": joined(f"{version_label(v)} ({v.get('releaseType')})" for v in
                                                 sorted(available, key=lambda v: v.get("releaseDate") or "")),
                    "Target Version": "", "Upgrade Time Local (YYYY-MM-DD HH:MM)": "",
                    "Excluded": excluded, "Notes": note,
                    "_shortnames": [version_label(v) for v in available],
                })
                for net, d in members:
                    dev_rows.append({"Org Name": org["name"], "Target Type": ttype, "Target Name": tname,
                                     "Target ID": tid, "Product": product, "Network Name": net.get("name"),
                                     "Serial": d.get("serial"), "Device Name": d.get("name") or "",
                                     "Model": d.get("model"), "Device Firmware": d.get("firmware") or ""})
                for v in available:
                    ver_rows.append({"Org Name": org["name"], "Target Type": ttype, "Target Name": tname,
                                     "Target ID": tid, "Product": product, "Version ID": str(v.get("id")),
                                     "Short Name": v.get("shortName"), "Firmware": v.get("firmware"),
                                     "Release Type": v.get("releaseType"),
                                     "Release Date": fmt_local(v.get("releaseDate"))})

    wb = Workbook()
    ws = wb.active
    ws.title = "Plan"
    ws.append(PLAN_HEADERS)
    for r in plan_rows:
        ws.append([r.get(h, "") for h in PLAN_HEADERS])
    style_sheet(ws, len(PLAN_HEADERS), text_cols=("O", "P"))
    ws.column_dimensions["N"].width = 60

    select_dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(select_dv)
    highlight = PatternFill("solid", fgColor="FFF2CC")
    grey = PatternFill("solid", fgColor="D9D9D9")
    for i, r in enumerate(plan_rows, start=2):
        if r.get("Excluded") or not r.get("Product"):
            for c in ws[i]:
                c.fill = grey
            continue
        select_dv.add(f"A{i}")
        for col in ("A", "O", "P"):
            ws[f"{col}{i}"].fill = highlight
        names = [n for n in r.get("_shortnames", []) if n and "," not in n and '"' not in n]
        formula = '"' + ",".join(names) + '"'
        if names and len(formula) <= 255:  # Excel's inline list limit
            dv = DataValidation(type="list", formula1=formula, allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"O{i}")

    write_sheet(wb, "Devices", PLAN_DEVICE_HEADERS, dev_rows)
    write_sheet(wb, "Available_Versions", VERSION_HEADERS, ver_rows)
    tz_name = getattr(LOCAL_TZ, "key", str(LOCAL_TZ))
    instructions = [
        ["How to use this plan"],
        ["1. Each Plan row is one upgrade target (template or standalone network) + product type."],
        ["   Meraki schedules firmware per network/template per product - every device listed under"],
        ["   that row in the Devices sheet is upgraded together."],
        ["2. Put Y in 'Select (Y)' for the rows you want to upgrade."],
        ["3. Target Version: pick from the dropdown, or type a Short Name / Version ID from Available_Versions."],
        ["   Only versions Meraki lists as available for that target are accepted (compatibility)."],
        [f"4. Upgrade Time Local: your local time ({tz_name}), format YYYY-MM-DD HH:MM, e.g. 2026-11-02 02:00."],
        ["5. Grey rows are excluded (exclusions.txt) or have no firmware info and will be refused."],
        ["6. Save, then run:  python3 meraki_tool.py schedule            (dry run, shows what would happen)"],
        ["                    python3 meraki_tool.py schedule --apply    (actually schedules, asks to confirm)"],
        [f"Plan generated {fmt_local(datetime.now(timezone.utc))}. Re-run 'plan' if it is old."],
    ]
    wi = wb.create_sheet("Instructions", 0)
    for line in instructions:
        wi.append(line)
    wi["A1"].font = Font(bold=True, size=14)
    wi.column_dimensions["A"].width = 110
    wb.active = 1

    saved = save_workbook(wb, PLAN_FILE)
    log.info("Plan written: %s (%d target rows, %d devices)", saved, len(plan_rows), len(dev_rows))
    if collector.errors:
        log.warning("%d collection error(s); see log.", len(collector.errors))


# --------------------------------------------------------------------------- #
# SCHEDULE
# --------------------------------------------------------------------------- #
def parse_local_time(value):
    if isinstance(value, datetime):
        dt = value
    else:
        s = str(value or "").strip()
        dt = None
        for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M", "%Y-%m-%dT%H:%M:%S",
                    "%m/%d/%Y %H:%M", "%d-%m-%Y %H:%M"):
            try:
                dt = datetime.strptime(s, fmt)
                break
            except ValueError:
                continue
        if dt is None:
            raise ValueError(f"cannot parse time '{s}' (use YYYY-MM-DD HH:MM)")
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=LOCAL_TZ)
    return dt


def resolve_version(wanted, available):
    w = str(wanted).strip()
    if w.endswith(".0") and w[:-2].isdigit():  # Excel turned an ID into a float
        w = w[:-2]
    for v in available:
        if w == str(v.get("id")) or w.lower() in {(v.get("shortName") or "").lower(),
                                                   (v.get("firmware") or "").lower()}:
            return v
    return None


def load_plan():
    if not PLAN_FILE.exists():
        sys.exit(f"Plan file {PLAN_FILE} not found. Run: python3 meraki_tool.py plan")
    wb = load_workbook(PLAN_FILE, data_only=True)
    ws = wb["Plan"]
    it = ws.iter_rows(values_only=True)
    headers = list(next(it))
    rows = [dict(zip(headers, r)) for r in it]
    serials = defaultdict(list)
    if "Devices" in wb.sheetnames:
        dit = wb["Devices"].iter_rows(values_only=True)
        dh = list(next(dit))
        for r in dit:
            d = dict(zip(dh, r))
            serials[(d["Target ID"], d["Product"])].append((d.get("Serial"), d.get("Device Name")))
    return rows, serials


def append_scheduled_log(results):
    wb = open_workbook(OUTPUT_FILE)
    headers = ["Logged At", "Org Name", "Target Type", "Target Name", "Target ID", "Product",
               "From Version", "To Version", "Upgrade Time Local", "Upgrade Time UTC", "Devices",
               "Result", "Message"]
    if SCHEDULED_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(SCHEDULED_SHEET)
        ws.append(headers)
    ws = wb[SCHEDULED_SHEET]
    for r in results:
        ws.append([r.get(h, "") for h in headers])
    style_sheet(ws, len(headers))
    save_workbook(wb, OUTPUT_FILE)


def cmd_schedule(args):
    rows, plan_serials = load_plan()
    selected = [r for r in rows if str(r.get("Select (Y)") or "").strip().lower() in YES_VALUES]
    if not selected:
        sys.exit("No rows marked Y in the plan. Nothing to do.")

    dashboard = get_dashboard()
    exclusions = load_exclusions()
    now = datetime.now(LOCAL_TZ)
    tz_name = getattr(LOCAL_TZ, "key", str(LOCAL_TZ))
    log.info("Local time zone: %s (now %s)", tz_name, now.strftime("%Y-%m-%d %H:%M %Z"))

    jobs, problems = [], []
    for r in selected:
        label = f"{r.get('Target Type')} '{r.get('Target Name')}' / {r.get('Product')}"
        tid, product = r.get("Target ID"), r.get("Product")
        try:
            if not tid or not product:
                raise ValueError("row has no target/product")
            reason = exclusion_reason(
                exclusions, org=[r.get("Org ID"), r.get("Org Name")], target=[tid, r.get("Target Name")],
                device=[x for pair in plan_serials.get((tid, product), []) for x in pair])
            if reason:
                raise ValueError(f"excluded ({reason})")
            if not r.get("Target Version"):
                raise ValueError("Target Version is empty")
            when = parse_local_time(r.get("Upgrade Time Local (YYYY-MM-DD HH:MM)"))
            if when < now + timedelta(minutes=MIN_LEAD_MINUTES):
                raise ValueError(f"time {when:%Y-%m-%d %H:%M} is in the past or less than "
                                 f"{MIN_LEAD_MINUTES} min from now")
            fw = dashboard.networks.getNetworkFirmwareUpgrades(tid)  # live re-check
            info = (fw.get("products") or {}).get(product)
            if not info:
                raise ValueError(f"product '{product}' not present on target anymore")
            available = info.get("availableVersions") or []
            version = resolve_version(r["Target Version"], available)
            if not version:
                raise ValueError(f"version '{r['Target Version']}' is not available for this target "
                                 f"(available: {joined(version_label(v) for v in available)})")
            current = info.get("currentVersion") or {}
            if str(version.get("id")) == str(current.get("id")):
                raise ValueError(f"already running {version_label(current)}")
            downgrade = (version.get("releaseDate") or "") < (current.get("releaseDate") or "")
            if downgrade and not args.allow_downgrade:
                raise ValueError(f"{version_label(version)} is older than current {version_label(current)} "
                                 f"(use --allow-downgrade)")
            jobs.append({"row": r, "label": label, "when": when, "version": version, "current": current,
                         "release": (version.get("releaseType") or "").lower(),
                         "devices": len(plan_serials.get((tid, product), []))})
        except Exception as exc:  # noqa: BLE001
            problems.append((label, api_error_text(exc)))

    print(f"\nTime zone: {tz_name}")
    print(f"\n{'#':>2}  {'Target':<55} {'From':<16} {'To':<16} {'Local time':<22} Devices")
    print("-" * 125)
    for i, j in enumerate(jobs, 1):
        flag = "" if j["release"] == "stable" else f"  [{j['release'].upper() or 'NON-STABLE'}]"
        print(f"{i:>2}  {j['label'][:55]:<55} {version_label(j['current'])[:16]:<16} "
              f"{version_label(j['version'])[:16]:<16} {j['when']:%Y-%m-%d %H:%M %Z}  {j['devices']}{flag}")
    for label, msg in problems:
        print(f" !  {label}: SKIPPED - {msg}")
        log.warning("Skipped %s: %s", label, msg)
    print()

    if not jobs:
        sys.exit("No valid rows to schedule.")
    if not args.apply:
        print("DRY RUN - nothing was changed. Re-run with --apply to schedule the rows above.")
        return
    if not args.yes:
        answer = input(f"Type YES to schedule {len(jobs)} upgrade(s): ").strip()
        if answer != "YES":
            sys.exit("Aborted - nothing was changed.")

    results = []
    for j in jobs:
        r = j["row"]
        utc = j["when"].astimezone(timezone.utc)
        rec = {"Logged At": fmt_local(datetime.now(timezone.utc)), "Org Name": r.get("Org Name"),
               "Target Type": r.get("Target Type"), "Target Name": r.get("Target Name"),
               "Target ID": r.get("Target ID"), "Product": r.get("Product"),
               "From Version": version_label(j["current"]), "To Version": version_label(j["version"]),
               "Upgrade Time Local": j["when"].strftime("%Y-%m-%d %H:%M %Z"),
               "Upgrade Time UTC": utc.strftime("%Y-%m-%dT%H:%M:%SZ"), "Devices": j["devices"]}
        body = {r["Product"]: {"nextUpgrade": {"time": rec["Upgrade Time UTC"],
                                               "toVersion": {"id": str(j["version"]["id"])}}}}
        try:
            resp = dashboard.networks.updateNetworkFirmwareUpgrades(r["Target ID"], products=body)
            nxt = ((resp.get("products") or {}).get(r["Product"]) or {}).get("nextUpgrade") or {}
            ok = str((nxt.get("toVersion") or {}).get("id")) == str(j["version"]["id"])
            rec["Result"] = "Scheduled" if ok else "Submitted (unverified)"
            rec["Message"] = f"Dashboard shows {fmt_local(nxt.get('time'))}" if nxt.get("time") else ""
            log.info("%s: %s -> %s at %s", rec["Result"], j["label"], rec["To Version"], rec["Upgrade Time Local"])
        except Exception as exc:  # noqa: BLE001
            rec["Result"], rec["Message"] = "FAILED", api_error_text(exc)
            log.error("FAILED %s: %s", j["label"], rec["Message"])
        results.append(rec)

    append_scheduled_log(results)
    ok = sum(1 for r in results if r["Result"] != "FAILED")
    log.info("Done: %d scheduled, %d failed. Logged to sheet '%s' in %s",
             ok, len(results) - ok, SCHEDULED_SHEET, OUTPUT_FILE)
    log.info("Run 'report' to see the schedule on the device sheet.")


# --------------------------------------------------------------------------- #
def main():
    global LOCAL_TZ
    parser = argparse.ArgumentParser(description="Meraki inventory report and firmware scheduler")
    sub = parser.add_subparsers(dest="mode", required=True)
    sub.add_parser("report", help="read-only: write dated device sheets to the inventory workbook")
    sub.add_parser("plan", help="read-only: build the firmware plan workbook to fill in")
    sp = sub.add_parser("schedule", help="schedule upgrades from the plan (dry run unless --apply)")
    sp.add_argument("--apply", action="store_true", help="actually schedule (otherwise dry run)")
    sp.add_argument("--yes", action="store_true", help="skip the YES confirmation prompt")
    sp.add_argument("--allow-downgrade", action="store_true", help="permit scheduling an older version")
    args = parser.parse_args()

    setup_logging()
    LOCAL_TZ = get_local_tz()
    {"report": cmd_report, "plan": cmd_plan, "schedule": cmd_schedule}[args.mode](args)


if __name__ == "__main__":
    main()
