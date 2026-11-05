# Meraki Inventory & Firmware Scheduler

One script, three modes:

| Mode | Changes anything? | What it does |
|---|---|---|
| `report` | No | Gets every device from every org and network your key can see, and writes **dated sheets** into one workbook. |
| `plan` | No | Builds `firmware_plan.xlsx` for you to choose what to upgrade, to which version, and when. |
| `schedule` | Only with `--apply` | Reads the plan and schedules the selected upgrades. Without `--apply` it's a dry run. |

## Setup (Kali on WSL or Ubuntu server)

```bash
cd meraki_tool
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
chmod 600 .env
nano .env          # set MERAKI_API_KEY and LOCAL_TZ
```

> **Set `LOCAL_TZ`** (for example `America/Chicago` or `Europe/Amsterdam`). On WSL the system time zone is often UTC.
> Every time shown, and every time you type in the plan, uses this zone.
> The scheduler converts it to UTC for Meraki and handles daylight saving time.

## 1. Report

```bash
python3 meraki_tool.py report
```

Each run adds three sheets to `output/meraki_inventory.xlsx`. The newest run comes first:

- **Devices_YYYY-MM-DD_HHMM**: one row per device:
  - org, network, network tags and time zone
  - **template bound, template name and ID**, switch profile
  - name, serial, MAC, model, type
  - status, last reported, IPs, tags, address, notes
  - device firmware, network firmware, latest stable version, upgrade available
  - **scheduled upgrade time, target version and status** (and whether it comes from the device, network or template)
  - last upgrade, whether it's excluded, dashboard URL
  
  Devices that are claimed but not assigned to a network are listed as `(unassigned)`.
- **Changes_...**: what changed since the previous run: devices added or removed, and changed fields (name, network, template, status, firmware, schedule, IPs, tags).
- **Summary_...**: totals, plus counts by status, org, product type, model, firmware and template, and any **errors**, for example an org with API access disabled.
- **Scheduled_Upgrades** (a single sheet kept at the end): a history of every upgrade the tool has scheduled.

Only the last `KEEP_RUNS` runs are kept (default 30). A CSV copy of each Devices sheet is also written to `output/` (turn this off with `CSV_EXPORT=false`).

## 2. Plan → 3. Schedule

```bash
python3 meraki_tool.py plan                   # creates output/firmware_plan.xlsx
# open it, fill in the yellow cells, save
python3 meraki_tool.py schedule               # DRY RUN: shows what would happen
python3 meraki_tool.py schedule --apply       # schedules; you type YES to confirm
```

In the **Plan** sheet:
1. Put `Y` in **Select** for each row you want to upgrade.
2. Pick a **Target Version** from the dropdown, or type a short name or ID from the `Available_Versions` sheet.
3. Enter the **Upgrade Time Local**, for example `2026-11-15 02:00`.

Before scheduling anything, `schedule` checks each row **live against Meraki**:
- The version must be one that Meraki lists as available for that target. This is how it handles compatibility.
- The time must be at least `MIN_LEAD_MINUTES` in the future.
- The target must not be in `exclusions.txt`.
- The target must not already be on that version.
- The version must not be older than the current one, unless you pass `--allow-downgrade`.

Beta and release-candidate versions are flagged in the preview.

### Important: what gets upgraded together
Meraki schedules firmware **per network (or template) per product type**, not per device.
Each plan row is one such unit. The **Devices** sheet in the plan shows exactly which devices are included in each row.

- **Template-bound networks**: by default, scheduled once on the template (`TEMPLATE_FIRMWARE_MODE=template`).
  To get one row per bound network instead, set it to `network`.
  If Meraki refuses the template-level call, the plan automatically lists those networks individually and adds a note.
- **Standalone networks**: one row per network per product type.
- **Exclusions**: if a serial in `exclusions.txt` is part of a row, the whole row is blocked. Upgrading the rest of that row would also upgrade the excluded device.

## Running on a schedule (Ubuntu server)

```cron
# Inventory report every day at 06:10
10 6 * * * cd /opt/meraki_tool && .venv/bin/python meraki_tool.py report >> logs/cron.log 2>&1
```

## Files
- `logs/meraki_tool_YYYYMMDD.log`: audit log of every run
- `exclusions.txt`: never-upgrade list (org, network, template, serial or name)
- `.env`: settings and API key (never commit it; it's in `.gitignore`)

## Notes
- Close the workbook in Excel before running. If it's locked (common on WSL with `/mnt/c/...`), the tool saves a timestamped copy instead.
- Rate limits, retries and pagination are handled by the official `meraki` SDK.
- Use an API key from an account with **read-only** org access for `report` and `plan`. `schedule --apply` needs **full** access.
