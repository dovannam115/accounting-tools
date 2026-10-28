# Paths and update settings shared by server.py (the window) and worker.py (one run of a tool).
import os, sys

APP     = os.path.dirname(os.path.abspath(__file__))
ROOT    = os.path.dirname(APP)
ENGINES = os.path.join(APP, "engines")
DUMP    = os.path.join(APP, "dump")           # token, log, uploads
OUTPUT  = os.path.join(ROOT, "output")        # one folder per run: output\<Tool>\<date time>\
BACKUPS = os.path.join(ROOT, "backups")       # old settings files, each time one is replaced

if APP not in sys.path: sys.path.insert(0, APP)

# folder name of each tool under output\
TOOLS = {"debt": "Debt_Report", "pias": "PIAS_Fix", "offset": "Offset_5111", "update": "_update"}

# Settings files the user edits on this PC. Update replaces one only while it is still the shipped copy
# (its hash is the one recorded at the last install / update), so a file edited here is never overwritten.
SETTINGS = {
    "config":  os.path.join(ENGINES, "debt", "Config", "config.xlsx"),
    "units":   os.path.join(ENGINES, "pias", "Unit_Codes.xlsx"),
    "rates":   os.path.join(ENGINES, "offset", "Offset_Rates.xlsx"),
}

# Where "Update" looks for new versions. app\update.json wins when it names a repo.
REPO = {"repo": "dovannam115/accounting-tools", "branch": "main", "token": ""}


def update_cfg():
    import json
    out = dict(REPO)
    try:
        j = json.load(open(os.path.join(APP, "update.json"), encoding="utf-8"))
        if j.get("repo"): out.update({k: v for k, v in j.items() if v or k == "token"})
    except Exception:
        pass
    return out


def http_get(url, token="", timeout=25):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": "AccountingTools", "Cache-Control": "no-cache"})
    if token: req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=timeout) as r: return r.read()
