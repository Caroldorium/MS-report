#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10,<3.14"
# dependencies = ["numpy", "openpyxl", "pyteomics", "psims", "lxml", "rdkit", "python-pptx", "pywebview"]
# ///
"""MS Report dashboard: a window over make_ms_report.py (same folder).

    uv run dashboard.py                 # opens the window
    uv run dashboard.py --browser       # same page in the default web browser instead
    uv run dashboard.py --folder PATH   # start in this data folder

The page is served on 127.0.0.1 only and every request must carry a random token.
Reports are made by running make_ms_report.py as a separate process, so the dashboard
and Terminal give the same result.
"""
import argparse
import contextlib
import csv
import io
import json
import mimetypes
import os
import secrets
import shutil
import signal
import subprocess
import sys
import threading
import time
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

HERE = Path(__file__).resolve().parent
ENGINE = HERE / "make_ms_report.py"
sys.path.insert(0, str(HERE))
import make_ms_report as E  # noqa: E402

UI_FILE = HERE / "dashboard.html"
MANUAL_CANDIDATES = [HERE / "MS Report manual.html", HERE / "app_source" / "MS Report manual.html"]
ICON_CANDIDATES = [HERE / "MS Report.icns", HERE / "app_source" / "MS Report.icns"]
SETTINGS = E.MASTER_DIR / "dashboard.json"
LOCK = E.MASTER_DIR / ".dashboard-running.json"
HPLC_EXTS = {".txt", ".csv", ".tsv", ".asc", ".arw", ".dat"}
NOT_HPLC = {"samples.csv", "samples_old.csv", "building_blocks.csv", "summary.csv"}
DEFAULTS = dict(ppm=5.0, max_charge=6, min_rel=1.0, min_scans=8)


# ----------------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------------
def load_settings():
    try:
        return json.loads(SETTINGS.read_text())
    except Exception:
        return {}


def save_settings(s):
    try:
        SETTINGS.parent.mkdir(parents=True, exist_ok=True)
        SETTINGS.write_text(json.dumps(s, indent=1))
    except OSError:
        pass


def remember_folder(folder):
    s = load_settings()
    rec = [r for r in s.get("recent", []) if r.get("path") != str(folder)]
    rec.insert(0, {"path": str(folder), "opened": datetime.now().isoformat(timespec="minutes")})
    s["recent"] = rec[:12]
    save_settings(s)


def as_folder(path):
    if not path:
        raise E.UserError("no folder chosen")
    f = Path(path).expanduser().resolve()
    if not f.is_dir():
        raise E.UserError(f"folder not found: {f}")
    return f


def mtime(p):
    try:
        return p.stat().st_mtime
    except OSError:
        return None


def blocks_for(folder):
    path = E.find_blocks_file(folder)
    return path, E.load_blocks(path)


def capture(fn, *a, **k):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        fn(*a, **k)
    return buf.getvalue()


# ----------------------------------------------------------------------------
# folder, samples
# ----------------------------------------------------------------------------
def folder_files(folder):
    files = [p for p in folder.iterdir() if p.is_file() and not p.name.startswith((".", "~$"))]
    files.sort(key=lambda p: E.natural_key(p.name))
    raws = [p.name for p in files if p.suffix.lower() in (".raw", ".mzml")]
    blanks = [n for n in raws if E.BLANK_RE.search(Path(n).stem)]
    hplc = [p.name for p in files if p.suffix.lower() in HPLC_EXTS and p.name not in NOT_HPLC]
    sources = [p.name for p in files if p.suffix.lower() in (".xlsx", ".xlsm")
               or (p.suffix.lower() == ".csv" and p.name not in NOT_HPLC)]
    return dict(raw_files=raws, blank_guess=blanks, hplc_files=hplc, sources=sources)


def read_rows(folder):
    sp = folder / "samples.csv"
    if not sp.exists():
        return None, []
    raw = sp.read_bytes()
    text = None
    for enc in ("utf-8-sig", "mac_roman"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    header = next(csv.reader(io.StringIO(text or "")), [])
    rows = E.read_samples(sp)
    extra = [c for c in header if c and c not in E.SAMPLE_COLUMNS and c != "also_look_for"]
    out = []
    for r in rows:
        d = {c: r.get(c, "") for c in E.SAMPLE_COLUMNS}
        d.update({c: r.get(c, "") for c in extra})
        out.append(d)
    return extra, out


def api_open_folder(a):
    folder = as_folder(a.get("folder"))
    remember_folder(folder)
    extra, rows = read_rows(folder)
    bpath = E.find_blocks_file(folder)
    out = folder / E.OUT_DIR_NAME
    return dict(folder=str(folder), name=folder.name, parent=str(folder.parent),
                has_samples=(folder / "samples.csv").exists(), rows=rows,
                extra_columns=extra or [], samples_mtime=mtime(folder / "samples.csv"),
                blocks_path=str(bpath), has_results=(out / "summary.csv").exists(),
                **folder_files(folder))


def api_setup(a):
    folder = as_folder(a.get("folder"))
    blocks = E.load_blocks(E.find_blocks_file(folder))
    log = capture(E.cmd_setup, folder, blocks, force=bool(a.get("force")))
    res = api_open_folder(a)
    res["log"] = log
    return res


def api_save(a):
    folder = as_folder(a.get("folder"))
    sp = folder / "samples.csv"
    known = a.get("samples_mtime")
    now = mtime(sp)
    if sp.exists() and known is not None and now is not None and abs(now - known) > 1e-3 and not a.get("force"):
        return dict(conflict=True, error="samples.csv was changed outside the dashboard (Excel?) since it was loaded.")
    extra = [c for c in a.get("extra_columns", []) if c not in E.SAMPLE_COLUMNS]
    cols = E.SAMPLE_COLUMNS + extra
    rows = [{c: str(r.get(c, "") or "").strip() for c in cols} for r in a.get("rows", [])]
    if sp.exists():
        shutil.copy(sp, folder / ".samples_backup.csv")
    tmp = folder / ".samples.csv.tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    os.replace(tmp, sp)
    return dict(samples_mtime=mtime(sp))


def validate_rows(folder, rows, blocks, max_z=4):
    ids = {}
    for r in rows:
        if (r.get("role") or "sample") != "blank":
            ids[r.get("id", "")] = ids.get(r.get("id", ""), 0) + 1
    out = []
    for r in rows:
        errs, info = {}, {}
        role = r.get("role") or "sample"
        raw = r.get("raw_file", "")
        if role not in ("sample", "blank"):
            errs["role"] = "sample or blank"
        if raw and not (folder / raw).exists():
            errs["raw_file"] = "file not in the folder"
        if role == "blank":
            out.append(dict(errors=errs, info=info, runnable=False, blank=True))
            continue
        sid = r.get("id", "")
        if not sid:
            errs["id"] = "needs an id"
        elif ids.get(sid, 0) > 1:
            errs["id"] = "id used twice"
        bf = r.get("blank_file", "")
        if bf not in ("", "auto") and not (folder / bf).exists():
            errs["blank_file"] = "file not in the folder"
        if r.get("hplc_file") and not (folder / r["hplc_file"]).exists():
            errs["hplc_file"] = "file not in the folder"
        try:
            E.parse_observed(r.get("observed_mz", ""))
        except E.UserError as e:
            errs["observed_mz"] = str(e).replace("observed_mz: ", "")
        try:
            E.parse_rt(r.get("spectrum_rt", ""))
        except E.UserError as e:
            errs["spectrum_rt"] = str(e).replace("spectrum_rt: ", "")
        pep = None
        seq = r.get("sequence", "")
        if not seq.strip():
            errs["sequence"] = "no sequence"
        elif r.get("topology", "") not in E.TOPOLOGIES:
            errs["topology"] = "choose linear or head-to-tail"
            try:   # still show the mass, guessing the topology
                pep = E.parse_sequence(seq, "linear" if "-" in seq else "head-to-tail", blocks)
                info["guess"] = True
            except E.UserError:
                pep = None
        else:
            try:
                pep = E.parse_sequence(seq, r["topology"], blocks)
            except E.UserError as e:
                msg = str(e)
                errs["topology" if "head-to-tail peptide has no ends" in msg else "sequence"] = msg
        if pep is not None:
            info.update(residues=len(pep.units), formula=E.hill(pep.formula), mass=round(pep.mass, 4),
                        one_letter=E.one_letter(pep))
            ions = E.make_species(pep, "none", blocks, (0, 0), max_z)[0].ions
            info["ions"] = [[i.label, round(i.mz, 4)] for i in ions if i.z >= 2] or \
                           [[i.label, round(i.mz, 4)] for i in ions]
            try:
                info["deltas"] = len(E.expand_deltas(r.get("find_deltas", ""), pep, blocks))
            except E.UserError as e:
                errs["find_deltas"] = str(e).split(". Use")[0].replace("find_deltas: ", "")
        runnable = bool(raw) and not errs
        out.append(dict(errors=errs, info=info, runnable=runnable, skipped=not raw, blank=False))
    return out


def api_validate(a):
    folder = as_folder(a.get("folder"))
    _, blocks = blocks_for(folder)
    return dict(results=validate_rows(folder, a.get("rows", []), blocks))


def api_deltas(a):
    folder = as_folder(a.get("folder"))
    path, blocks = blocks_for(folder)
    groups = []
    for g in E.delta_groups(blocks):
        items = [b for b in blocks.values() if b.group == g]
        groups.append(dict(name=g, items=[dict(code=b.code, name=b.name, mass=round(b.mass, 4),
                                                 residues=b.residues) for b in items]))
    loose = [dict(code=b.code, name=b.name, mass=round(b.mass, 4), residues=b.residues)
             for b in blocks.values() if b.kind in ("shift", "adduct") and not b.group]
    builtin = [dict(name="linear", text="uncyclised precursor (+H2O), head-to-tail only"),
               dict(name="deletions", text="each residue missing once"),
               dict(name="insertions", text="each residue present twice"),
               dict(name="dimer", text="2M, odd charge states")]
    return dict(auto=list(E.AUTO_GROUPS), builtin=builtin, groups=groups, loose=loose)


# ----------------------------------------------------------------------------
# runs
# ----------------------------------------------------------------------------
class Job:
    def __init__(self):
        self.lock = threading.Lock()
        self.proc = None
        self.lines = []
        self.ids = []
        self.current = None
        self.done_ids = []
        self.started = None
        self.ended = None
        self.code = None
        self.pptx = None
        self.folder = None
        self.cancelled = False

    def running(self):
        return self.proc is not None and self.proc.poll() is None


JOB = Job()


def api_run(a):
    folder = as_folder(a.get("folder"))
    with JOB.lock:
        if JOB.running():
            raise E.UserError("a report is already being made")
        ids = [str(i) for i in a.get("ids") or []]
        s = {k: a.get(k, v) for k, v in DEFAULTS.items()}
        cmd = [sys.executable, "-u", str(ENGINE), "run", "--folder", str(folder),
               "--ppm", str(float(s["ppm"])), "--max-charge", str(int(s["max_charge"])),
               "--min-rel", str(float(s["min_rel"])), "--min-scans", str(int(s["min_scans"]))]
        if a.get("only") and ids:
            cmd += ["--only", ",".join(ids)]
        env = dict(os.environ, PYTHONUNBUFFERED="1", MS_REPORT_HOME=str(E.MASTER_DIR))
        JOB.__init__()
        JOB.ids, JOB.folder, JOB.started = ids, str(folder), time.time()
        JOB.proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                                    bufsize=1, env=env, start_new_session=True)
        threading.Thread(target=_reader, args=(JOB.proc,), daemon=True).start()
    return dict(started=True)


def _reader(proc):
    import re
    for line in proc.stdout:
        line = line.rstrip("\n")
        with JOB.lock:
            JOB.lines.append(line)
            m = re.match(r"^\[(.+?)\] ", line)
            if m:
                if JOB.current and JOB.current not in JOB.done_ids:
                    JOB.done_ids.append(JOB.current)
                JOB.current = m.group(1)
            m = re.match(r"^Done: (.+\.pptx)$", line)
            if m:
                JOB.pptx = m.group(1)
    proc.wait()
    with JOB.lock:
        if JOB.current and proc.returncode == 0 and JOB.current not in JOB.done_ids:
            JOB.done_ids.append(JOB.current)
        JOB.code, JOB.ended = proc.returncode, time.time()


def api_run_status(a):
    since = int(a.get("since") or 0)
    with JOB.lock:
        return dict(running=JOB.running(), lines=JOB.lines[since:], n_lines=len(JOB.lines),
                    current=JOB.current, done=len(JOB.done_ids), ids=JOB.ids, code=JOB.code,
                    pptx=JOB.pptx, folder=JOB.folder, cancelled=JOB.cancelled,
                    elapsed=(JOB.ended or time.time()) - JOB.started if JOB.started else 0)


def stop_job():
    with JOB.lock:
        if JOB.running():
            JOB.cancelled = True
            try:
                os.killpg(JOB.proc.pid, signal.SIGTERM)
            except OSError:
                JOB.proc.terminate()


def api_cancel(a):
    stop_job()
    return {}


# ----------------------------------------------------------------------------
# results
# ----------------------------------------------------------------------------
def api_results(a):
    folder = as_folder(a.get("folder"))
    out = folder / E.OUT_DIR_NAME
    sp = out / "summary.csv"
    if not sp.exists():
        return dict(rows=[], pptx=None)
    with open(sp, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    decks = sorted(out.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
    import re
    for r in rows:
        sid = re.sub(r"[^\w.-]+", "_", r.get("id", ""))
        svg = out / "data" / f"{sid}.svg"
        r["_svg"] = str(svg) if svg.exists() else ""
        mp = out / "data" / f"{sid}_main_peaks.csv"
        r["_main_peaks"] = []
        if mp.exists():
            with open(mp, newline="") as f:
                r["_main_peaks"] = list(csv.DictReader(f))
        r["_mz_table"] = []
        mt = out / "data" / f"{sid}_mz_table.csv"
        if mt.exists():
            with open(mt, newline="") as f:
                r["_mz_table"] = list(csv.DictReader(f))
    return dict(rows=rows, pptx=str(decks[0]) if decks else None,
                made=datetime.fromtimestamp(sp.stat().st_mtime).strftime("%d %b %Y, %H:%M"),
                out=str(out), summary=str(sp))


# ----------------------------------------------------------------------------
# building blocks
# ----------------------------------------------------------------------------
BLOCK_COLUMNS = ["code", "code3", "name", "kind", "formula", "smiles", "group", "residues", "notes"]


def api_blocks(a):
    folder = as_folder(a.get("folder"))
    path = E.find_blocks_file(folder)
    with open(path, newline="", encoding="utf-8-sig") as f:
        rd = csv.DictReader(f)
        cols = list(rd.fieldnames or [])
        rows = [{c: (r.get(c) or "") for c in cols} for r in rd]
    for c in BLOCK_COLUMNS:
        if c not in cols:
            cols.append(c)
    in_folder = path.parent.resolve() == folder
    return dict(path=str(path), columns=cols, rows=rows, mtime=mtime(path), in_folder=in_folder)


def block_info(r):
    kind = (r.get("kind") or "").strip()
    errs, info = {}, {}
    if not (r.get("code") or "").strip():
        errs["code"] = "needs a code"
    if kind not in ("residue", "n_cap", "c_cap", "tag", "adduct", "shift"):
        errs["kind"] = "choose a kind"
    try:
        fc = E.parse_formula(r.get("formula") or "")
        info["mass"] = round(E.formula_mass(fc), 4)
        info["hill"] = E.hill(fc)
    except E.UserError as e:
        errs["formula"] = str(e)
        fc = None
    smi = (r.get("smiles") or "").strip()
    if smi and fc is not None and "kind" not in errs and kind in ("residue", "n_cap", "c_cap", "tag"):
        try:
            msg = E.check_block_smiles(kind, smi, fc)
        except Exception as e:   # RDKit sanitisation errors
            msg = f"smiles: {e}"
        if msg:
            errs["smiles"] = msg
    elif not smi and kind in ("residue", "n_cap", "c_cap", "tag"):
        errs["smiles"] = "needs a smiles"
    g = (r.get("group") or "").strip()
    if g and kind not in ("shift", "adduct"):
        errs["group"] = "only shift/adduct rows have a group"
    elif g and not all(ch.islower() or ch.isdigit() or ch == "_" for ch in g):
        errs["group"] = "lower case, e.g. my_group"
    elif g in E.BUILTIN_DELTA_GROUPS + ("auto", "all", "none"):
        errs["group"] = "built-in name; choose another"
    return dict(errors=errs, info=info)


def api_blocks_check(a):
    rows = a.get("rows", [])
    res = [block_info(r) for r in rows]
    seen = {}
    for i, r in enumerate(rows):
        c = (r.get("code") or "").strip()
        if c:
            if c in seen:
                res[i]["errors"]["code"] = "code used twice"
            seen[c] = i
    return dict(results=res)


def api_blocks_save(a):
    folder = as_folder(a.get("folder"))
    path = E.find_blocks_file(folder)
    if a.get("mtime") is not None and abs((mtime(path) or 0) - a["mtime"]) > 1e-3 and not a.get("force"):
        return dict(conflict=True, error=f"{path.name} was changed outside the dashboard since it was loaded.")
    cols = list(dict.fromkeys(list(a.get("columns") or []) + BLOCK_COLUMNS))
    rows = [{c: str(r.get(c, "") or "").strip() for c in cols} for r in a.get("rows", [])
            if any(str(v or "").strip() for v in r.values())]
    tmp = path.with_name(".building_blocks.tmp.csv")
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    try:
        E.load_blocks(tmp)
    except E.UserError as e:
        tmp.unlink(missing_ok=True)
        return dict(error=str(e).replace(str(tmp), path.name))
    shutil.copy(path, path.with_name(".building_blocks_backup.csv"))
    os.replace(tmp, path)
    return dict(mtime=mtime(path), path=str(path))


# ----------------------------------------------------------------------------
# Finder, Excel, folder picker
# ----------------------------------------------------------------------------
WINDOW = None
CURRENT = {"folder": None}


def allowed_path(p):
    p = Path(p).expanduser().resolve()
    roots = [E.MASTER_DIR.resolve()] + ([Path(CURRENT["folder"])] if CURRENT["folder"] else [])
    return any(p == r or r in p.parents for r in roots)


def api_open(a):
    p = Path(a.get("path", "")).expanduser().resolve()
    if not allowed_path(p) or not p.exists():
        raise E.UserError("that file is not in the current folder")
    if a.get("reveal"):
        subprocess.run(["open", "-R", str(p)])
    elif a.get("excel"):
        if subprocess.run(["open", "-b", "com.microsoft.Excel", str(p)], capture_output=True).returncode:
            subprocess.run(["open", str(p)])
    else:
        subprocess.run(["open", str(p)])
    return {}


def api_choose_folder(a):
    start = a.get("start") or str(Path.home() / "Downloads")
    if WINDOW is not None:
        import webview
        kind = getattr(getattr(webview, "FileDialog", None), "FOLDER", None) or webview.FOLDER_DIALOG
        got = WINDOW.create_file_dialog(kind, directory=start)
        path = got[0] if got else None
    else:
        r = subprocess.run(["osascript", "-e",
                            f'POSIX path of (choose folder with prompt "Folder with the .raw files" '
                            f'default location (POSIX file "{start}"))'], capture_output=True, text=True)
        path = r.stdout.strip() or None
    return dict(folder=path)


def api_state(a):
    s = load_settings()
    recent = [r for r in s.get("recent", []) if Path(r["path"]).is_dir()]
    return dict(recent=recent, home=str(E.MASTER_DIR), defaults=DEFAULTS, settings=s.get("run", {}),
                engine=str(ENGINE), window=WINDOW is not None)


def api_remember_settings(a):
    s = load_settings()
    s["run"] = {k: a[k] for k in DEFAULTS if k in a}
    save_settings(s)
    return {}


def api_focus(a):
    if WINDOW is not None:
        try:
            WINDOW.restore()
            WINDOW.show()
            from AppKit import NSApplication
            NSApplication.sharedApplication().activateIgnoringOtherApps_(True)
        except Exception:
            pass
    return {}


API = {n[4:]: f for n, f in globals().items() if n.startswith("api_")}


# ----------------------------------------------------------------------------
# server
# ----------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    token = ""
    port = 0

    def log_message(self, *a):
        pass

    def _host_ok(self):
        return self.headers.get("Host", "") in (f"127.0.0.1:{self.port}", f"localhost:{self.port}")

    def _send(self, code, body, ctype="application/json"):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, default=str).encode()
        elif isinstance(body, str):
            body = body.encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        if not self._host_ok() or q.get("t", [""])[0] != self.token:
            return self._send(403, "forbidden", "text/plain")
        if u.path == "/":
            html = UI_FILE.read_text().replace("__TOKEN__", self.token)
            return self._send(200, html, "text/html; charset=utf-8")
        if u.path == "/manual":
            for p in MANUAL_CANDIDATES:
                if p.exists():
                    return self._send(200, p.read_bytes(), "text/html; charset=utf-8")
            return self._send(404, "manual not found", "text/plain")
        if u.path == "/file":
            p = Path(q.get("path", [""])[0]).expanduser().resolve()
            if p.suffix.lower() in (".svg", ".png") and allowed_path(p) and p.exists():
                return self._send(200, p.read_bytes(), mimetypes.guess_type(p.name)[0] or "application/octet-stream")
            return self._send(404, "not found", "text/plain")
        return self._send(404, "not found", "text/plain")

    def do_POST(self):
        u = urlparse(self.path)
        if not self._host_ok() or self.headers.get("X-Token") != self.token:
            return self._send(403, {"ok": False, "error": "forbidden"})
        name = u.path.removeprefix("/api/")
        fn = API.get(name)
        if fn is None:
            return self._send(404, {"ok": False, "error": f"unknown call {name}"})
        try:
            n = int(self.headers.get("Content-Length") or 0)
            a = json.loads(self.rfile.read(n) or b"{}")
            if a.get("folder"):
                CURRENT["folder"] = str(Path(a["folder"]).expanduser().resolve())
            res = fn(a) or {}
            return self._send(200, {"ok": "error" not in res, **res})
        except E.UserError as e:
            return self._send(200, {"ok": False, "error": str(e)})
        except Exception as e:   # show it in the window rather than losing it
            import traceback
            traceback.print_exc()
            return self._send(200, {"ok": False, "error": f"{type(e).__name__}: {e}"})


def already_running():
    """Another dashboard is open: bring it to the front and say so."""
    try:
        info = json.loads(LOCK.read_text())
        os.kill(int(info["pid"]), 0)
    except Exception:
        return False
    import urllib.request
    try:
        req = urllib.request.Request(f"http://127.0.0.1:{info['port']}/api/focus", data=b"{}",
                                     headers={"X-Token": info["token"], "Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=2).read()
        return True
    except Exception:
        return False


def main():
    global WINDOW
    ap = argparse.ArgumentParser(description="MS Report dashboard")
    ap.add_argument("--browser", action="store_true", help="open in the web browser instead of a window")
    ap.add_argument("--port", type=int, default=0)
    ap.add_argument("--folder")
    ap.add_argument("--no-open", action="store_true", help="only serve (for testing)")
    args = ap.parse_args()
    if not args.no_open and already_running():
        print("MS Report is already open.")
        return
    token = secrets.token_urlsafe(24)
    Handler.token = token
    srv = ThreadingHTTPServer(("127.0.0.1", args.port), Handler)
    Handler.port = srv.server_address[1]
    url = f"http://127.0.0.1:{Handler.port}/?t={token}"
    if args.folder:
        url += "&folder=" + __import__("urllib.parse").parse.quote(str(Path(args.folder).expanduser().resolve()))
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    try:
        E.MASTER_DIR.mkdir(parents=True, exist_ok=True)
        LOCK.write_text(json.dumps(dict(pid=os.getpid(), port=Handler.port, token=token)))
    except OSError:
        pass
    print(f"MS Report dashboard: {url}", flush=True)
    try:
        if args.no_open:
            threading.Event().wait()
        elif args.browser:
            subprocess.run(["open", url])
            threading.Event().wait()
        else:
            try:
                from Foundation import NSBundle   # Dock and menu bar say "MS Report", not "Python"
                NSBundle.mainBundle().infoDictionary()["CFBundleName"] = "MS Report"
            except Exception:
                pass
            import webview
            WINDOW = webview.create_window("MS Report", url, width=1380, height=880, min_size=(1000, 640),
                                           text_select=True)

            def on_start():
                icon = next((p for p in ICON_CANDIDATES if p.exists()), None)
                if icon:
                    try:
                        from AppKit import NSApplication, NSImage
                        from PyObjCTools import AppHelper
                        img = NSImage.alloc().initWithContentsOfFile_(str(icon))
                        AppHelper.callAfter(NSApplication.sharedApplication().setApplicationIconImage_, img)
                    except Exception:
                        pass
            webview.start(on_start)
    except KeyboardInterrupt:
        pass
    finally:
        stop_job()
        srv.shutdown()
        try:
            if json.loads(LOCK.read_text()).get("pid") == os.getpid():
                LOCK.unlink()
        except Exception:
            pass


if __name__ == "__main__":
    main()
