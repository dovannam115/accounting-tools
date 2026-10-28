# One run of one tool, started by server.py as a child process:
#   python -u worker.py <tool> <job dir>            a run: reads args.json, writes result.json
#   python -u worker.py --inspect <kind> <file>     a quick look at uploaded files: prints one JSON line
#   python -u worker.py --setting <name> <file>     replace a settings file (checked first): prints one JSON line
# Lines it prints go to the run panel in the window. Markers the server reads:
#   [[P]] done|total|label     progress bar
#   [[ERROR]] text             the message the user sees when the run fails
import os, sys, json, traceback, datetime, decimal, zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import common as C

try: sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
except Exception: pass

DIR = ""


def say(*a):
    print(" ".join(str(x) for x in a), flush=True)


def prog(done, total, label=""):
    print("[[P]] %d|%d|%s" % (done, total, str(label).replace("\n", " ")[:120]), flush=True)


class UserError(Exception):
    """A failure the user can fix (wrong file, missing type...): shown without a traceback."""


def jsonable(v):
    if v is None or isinstance(v, (bool, int, str)): return v
    if isinstance(v, float): return None if v != v else v
    if isinstance(v, decimal.Decimal): return float(v)
    if isinstance(v, (datetime.datetime, datetime.date)): return v.strftime("%d/%m/%Y")
    if isinstance(v, (bytes, bytearray)): return None
    return str(v)


def table(title, cols, rows, limit=2000):
    recs = [r if isinstance(r, dict) else dict(zip(cols, r)) for r in rows]
    return {"title": title, "cols": cols, "total": len(recs), "links": [],
            "rows": [[jsonable(r.get(c)) for c in cols] for r in recs[:limit]]}


def stat(label, value, tone=""):
    return {"label": label, "value": value, "tone": tone}


def day(s):
    """'2026-10-22' or '22/10/2026' -> date; '' -> None."""
    if not s: return None
    for f in ("%Y-%m-%d", "%d/%m/%Y"):
        try: return datetime.datetime.strptime(s, f).date()
        except ValueError: pass
    raise UserError("Bad date: %s" % s)


def take(paths):
    """[(name, bytes)] of the staged inputs. The runner saves its own copy under Input\\ of this run,
    so the staged copy is removed to keep one copy of each file."""
    out = []
    for p in paths:
        out.append((os.path.basename(p), open(p, "rb").read()))
        if os.path.dirname(os.path.abspath(p)) == os.path.abspath(DIR):
            try: os.remove(p)
            except OSError: pass
    return out


def zip_to(path, files):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for p in files: z.write(p, os.path.basename(p))
    return path


def excel_installed():
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Excel.Application\CLSID"))
        return True
    except Exception:
        return False


def runners():
    """Import the engines once, with every run folder = this run (no cleanup of other runs)."""
    import runners as R
    R.RUNS = DIR or os.path.join(C.DUMP, "inspect")
    R.BACKUPS = C.BACKUPS
    if DIR:
        R.new_run = lambda kind: DIR
    return R


def stamp():
    return datetime.datetime.now().strftime("%Y%m%d")


# ------------------------------------------------------------------ Debt report 1311
DEBT_LABEL = {"raw": "BCPBHTT export", "bk": "BK_HDVAT (invoice amounts)", "list": "List hóa đơn",
              "report": "BC công nợ already built", None: "Not recognised, ignored"}


def run_debt(a):
    R = runners()
    scope = a.get("scope") or "full"
    files = take(a["files"])
    has_raw = any(R.debt_kind(n) == "raw" for n, _ in files)
    verify = bool(a.get("verify")) and has_raw
    warns = []
    if verify and not excel_installed():
        verify = False
        warns.append("Excel is not on this PC, so the report was not re-checked in Excel.")
    say("Reading %d file(s) ..." % len(files))
    try:
        res = R.debt_run(files, verify=verify, clean=bool(a.get("clean")), kh=bool(a.get("kh")) and scope != "0051",
                         pkd=bool(a.get("pkd")), log=lambda m, tag=None: say(m),
                         progress=lambda p, m: prog(min(int(float(p)), 100), 100, m), scope=scope)
    except R.RunError as e:
        raise UserError(str(e))
    outs = [p for p in res["outputs"] if os.path.isfile(p)]
    xlsx = [p for p in outs if p.lower().endswith(".xlsx")]
    if len(outs) > 1:
        tag = {"0051": "0051", "13111": "13111"}.get(scope, "1311")
        files_, more = [zip_to(os.path.join(DIR, "Debt_Report_%s_%s.zip" % (tag, stamp())), outs)], outs
    else:
        files_, more = outs, []
    stats = [stat("Files", len(xlsx))]
    if res["verified"] is True:
        stats.append(stat("Excel re-check", "Passed", "good"))
    elif res["verified"] is False:
        stats.append(stat("Excel re-check", "Differences", "bad"))
        warns.append("The Excel re-check found differences. Read the run log before sending the report.")
    return {"stats": stats, "files": files_, "more": more, "warns": warns, "details": res["log_text"]}


# ------------------------------------------------------------------ PIAS fix for other PVI units
def run_pias(a):
    R = runners()
    files = take(a["files"])
    say("Fixing %d file(s) ..." % len(files))
    try:
        res = R.pias_run(files, sheet=(a.get("sheet") or "").strip() or None, dates=(a.get("dates") or "").strip() or None)
    except R.RunError as e:
        raise UserError(str(e))
    outs = [p for p in res["outputs"] if os.path.isfile(p)]
    xlsx = [p for p in outs if p.lower().endswith(".xlsx")]
    txt = [p for p in outs if p.lower().endswith(".txt")]
    if res["error"] and not xlsx:
        say(res["console"])
        raise UserError(res["error"])
    if len(xlsx) > 1:
        files_, more = [zip_to(os.path.join(DIR, "PIAS_Fixed_%s.zip" % stamp()), outs)], xlsx + txt
    else:
        files_, more = xlsx, txt
    fixed = sum(int(f["fixed"]) for f in res["files"])
    warns = [res["error"]] if res["error"] else []
    if not xlsx: warns.append("No row needed fixing, so no file was written.")
    return {"stats": [stat("Rows fixed", fixed, "good" if fixed else ""),
                      stat("To check by hand", res["check"] + res["look"], "warn" if res["check"] + res["look"] else ""),
                      stat("List items not found in PIAS", res["left"], "warn" if res["left"] else "")],
            "files": files_, "more": more, "warns": warns, "console": res["console"],
            "details": res["log_text"]}


# ------------------------------------------------------------------ Offset with the 5111 ledger
def run_offset(a):
    R = runners()
    led, det = take([a["ledger"]])[0], take([a["detail"]])[0]
    say("Reading the ledger and the detail file ...")
    try:
        res = R.offset_run(led, det, chot=day(a.get("chot")), ky=(a.get("ky") or "").strip() or None,
                           progress=lambda n: prog(0, 0, "Reading the ledger, %s rows" % format(n, ",").replace(",", ".")))
    except R.RunError as e:
        raise UserError(str(e))
    for ln in res["log_text"].splitlines(): say(ln)
    xlsx = [p for p in res["outputs"] if p.lower().endswith(".xlsx")]
    rows = [{"Invoice": r["hd"], "Detail": r["phi_det"], "Ledger": r["phi_led"], "Difference": r["lech"]}
            for r in res["bad_invoices"]]
    notes = []
    if res["n_unassigned"]:
        notes.append("%s key(s) without an invoice number, %s of premium: sheet Hóa đơn chưa gán."
                     % (format(res["n_unassigned"], ",").replace(",", "."),
                        format(round(res["phi_unassigned"]), ",").replace(",", ".")))
    n = lambda v: format(v, ",").replace(",", ".")
    return {"stats": [stat("Invoices off", "%s/%s" % (n(res["n_invoice_bad"]), n(res["n_invoice"])),
                           "warn" if res["n_invoice_bad"] else "good"),
                      stat("Policy keys off", "%s/%s" % (n(res["n_key_bad"]), n(res["n_key"])),
                           "warn" if res["n_key_bad"] else "good"),
                      stat("Premium not in the ledger", round(res["phi_missing"]), "warn" if res["phi_missing"] else "")],
            "files": xlsx, "notes": notes,
            "tables": [table("Invoices off", ["Invoice", "Detail", "Ledger", "Difference"], rows)] if rows else [],
            "details": res["log_text"]}


RUN = {"debt": run_debt, "pias": run_pias, "offset": run_offset}


# ------------------------------------------------------------------ quick looks for the form
def inspect(kind, a):
    global DIR
    DIR = ""
    R = runners()
    try:
        if kind == "debt":
            out = []
            for p in a["files"]:
                k = R.debt_kind(p)
                acc = R.raw_account(open(p, "rb").read()) if k == "raw" else None
                out.append({"name": os.path.basename(p), "kind": k, "type": DEBT_LABEL[k], "account": acc or ""})
            return {"files": out, "scopes": {k: list(v) for k, v in R.SCOPE_ACCOUNTS.items()}}
        if kind == "pias":
            out = []
            for p in a["files"]:
                k = R.pias_kind(open(p, "rb").read())
                out.append({"name": os.path.basename(p), "kind": k,
                            "type": "PIAS receipt file" if k == "pias" else "PVI khác list"})
            return {"files": out}
        return {"error": "unknown"}
    except R.RunError as e:
        return {"error": str(e)}


def setting(name, path):
    R = runners()
    fn = {"config": R.replace_debt_config, "units": R.replace_unit_codes, "rates": R.replace_offset_rates}[name]
    try:
        fn(open(path, "rb").read())
    except R.RunError as e:
        return {"error": str(e)}
    return {"ok": True}


def main():
    global DIR
    if sys.argv[1] in ("--inspect", "--setting"):
        try:
            if sys.argv[1] == "--inspect":
                out = inspect(sys.argv[2], json.load(open(sys.argv[3], encoding="utf-8")))
            else:
                out = setting(sys.argv[2], sys.argv[3])
        except Exception as e:
            out = {"error": "%s: %s" % (type(e).__name__, e)}
        print("[[JSON]] " + json.dumps(out, ensure_ascii=False, default=jsonable), flush=True)
        return
    tool, DIR = sys.argv[1], sys.argv[2]
    a = json.load(open(os.path.join(DIR, "args.json"), encoding="utf-8"))
    res = {"ok": False}
    try:
        res = RUN[tool](a) or {}
        res["ok"] = True
        rc = 0
    except UserError as e:
        print("[[ERROR]] " + str(e), flush=True); res["error"] = str(e); rc = 2
    except Exception as e:
        tb = traceback.format_exc()
        say(tb)
        msg = "Unexpected error: %s" % e
        print("[[ERROR]] " + msg.replace("\n", " "), flush=True)
        res["error"] = msg; res["trace"] = tb; rc = 2
    for k in ("files", "more"):
        res[k] = [{"name": os.path.basename(p), "path": p} for p in res.get(k, []) if p and os.path.exists(p)]
    json.dump(res, open(os.path.join(DIR, "result.json"), "w", encoding="utf-8"), ensure_ascii=False, default=jsonable)
    sys.exit(rc)


if __name__ == "__main__":
    main()
