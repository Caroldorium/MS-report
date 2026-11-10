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
import re
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
        try:
            E.parse_purity_rt(r.get("purity_rt", ""))
        except E.UserError as e:
            errs["purity_rt"] = str(e).replace("purity_rt: ", "")
        pep = None
        seq = r.get("sequence", "")
        if not seq.strip():
            errs["sequence"] = "no sequence"
        elif r.get("topology", "") not in E.TOPOLOGIES:
            errs["topology"] = "choose linear or head-to-tail"
            try:   # still show the mass, guessing the topology
                pep = E.parse_sequence(seq, E.guess_topology(seq), blocks)
                info["guess"] = True
            except E.UserError:
                pep = None
        else:
            try:
                pep = E.parse_sequence(seq, r["topology"], blocks, r.get("bridges", ""))
            except E.UserError as e:
                msg = str(e)
                if msg.startswith("bridges: "):
                    errs["bridges"] = msg[len("bridges: "):]
                    try:   # still show the mass, without the bridge
                        pep = E.parse_sequence(seq, r["topology"], blocks)
                        info["guess"] = True
                    except E.UserError:
                        pep = None
                else:
                    errs["topology" if "head-to-tail peptide has no ends" in msg
                         or "written as cyclo" in msg else "sequence"] = msg
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


# ----------------------------------------------------------------------------
# pasted sequences, mass calculator
# ----------------------------------------------------------------------------
def blocks_any(a):
    """Building blocks of the open folder, or the shared ones when no folder is open."""
    folder = Path(a["folder"]).expanduser().resolve() if a.get("folder") else E.MASTER_DIR
    E.MASTER_DIR.mkdir(parents=True, exist_ok=True)
    with contextlib.redirect_stdout(io.StringIO()):
        path = E.find_blocks_file(folder if folder.is_dir() else E.MASTER_DIR)
    return E.load_blocks(path)


def has_ends(seq, blocks):
    caps = [b.code for b in blocks.values() if b.kind in ("n_cap", "c_cap")]
    s = re.sub(r"\s+", "", seq)
    return any(s.startswith(c) for c in caps if c.endswith("-")) or any(s.endswith(c) for c in caps if c.startswith("-"))


NUM_RE = re.compile(r"[\d.,]+")
CELL_RE = re.compile(r"[A-Za-z][A-Za-z0-9]{1,5}(\([^)]+\)|\[[^\]]+\])?")


def split_line(line, blocks):
    """One pasted line -> (name, sequence, numbers, from_cells).
    Cells come tab-separated from Excel or Word, else separated by spaces, commas or semicolons.
    A batch sheet has one residue per cell (Ser | Ala | Arg ...): that run of cells is the sequence.
    Otherwise the sequence is the last cell that reads as one."""
    tabbed = "\t" in line
    raw = [c.strip() for c in (line.split("\t") if tabbed else re.split(r"[,;]\s*|\s+", line))]
    cells = [c for c in raw if c]
    if not cells:
        return None
    numbers = []
    for c in cells:
        if NUM_RE.fullmatch(c) and "." in c:
            try:
                numbers.append(float(c.replace(",", "")))
            except ValueError:
                pass
    if tabbed:      # one residue per cell
        three = {b.code3 for b in blocks.values() if b.kind == "residue" and b.code3}
        best, k = None, 0
        while k < len(raw):
            j = k
            while j < len(raw) and CELL_RE.fullmatch(raw[j]):
                j += 1
            known = sum(1 for c in raw[k:j] if re.split(r"[(\[]", c)[0] in three)
            if j - k >= 2 and known >= 2 and (best is None or j - k > best[1] - best[0]):
                best = (k, j)
            k = max(j, k + 1)
        if best:
            before = [c for c in raw[:best[0]] if c]
            return (before[-1] if before else ""), "-".join(raw[best[0]:best[1]]), numbers, True
    hit = None
    for k, c in enumerate(cells):
        if NUM_RE.fullmatch(c):
            continue
        try:
            E.parse_sequence(c, E.guess_topology(c), blocks)
            hit = k
        except E.UserError:
            pass
    if hit is None:
        texty = [k for k, c in enumerate(cells) if not NUM_RE.fullmatch(c)]
        # nothing reads as a sequence: take the cell that looks most like one, so its error is shown
        hit = max(texty, key=lambda k: (sum(ch.isupper() for ch in cells[k]) / len(cells[k]), k)) if texty \
            else len(cells) - 1
    others = cells[:hit] + (cells[hit + 1:] if tabbed else [])
    if tabbed:
        named = [c for c in others if not NUM_RE.fullmatch(c)]
        name = named[0] if named else (others[0] if others else "")
    else:
        name = " ".join(others)
    return name, cells[hit], numbers, False


def fit_sheet_mass(d, seq, top, blocks, bridges, numbers):
    """A mass written on the sheet next to the sequence settles what the sequence cannot say:
    the ends. Tries the peptide as written, then as C-terminal amide, free acid and head-to-tail."""
    nums = [n for n in numbers if n > 150]
    if not nums:
        return
    def fits(pep):
        for n in nums:
            for what, m in (("M", pep.mass), ("[M+H]+", pep.mass + E.PROTON)):
                if abs(n - m) <= 0.02:
                    return n, what
        return None
    got = fits(d["pep"])
    if got:
        d.update(sheet=got[0], sheet_note=f"= {got[1]}")
        return
    bare = not has_ends(seq, blocks) and not E.CYCLO_RE.match(seq)
    one = E.read_sequence(seq, blocks)[0]
    forms = [("linear", one + "-NH2", "as C-terminal amide"), ("linear", one, "as free acid"),
             ("head-to-tail", one, "as head-to-tail")] if bare else []
    for t2, s2, text in forms:
        try:
            pep = E.parse_sequence(s2, t2, blocks, bridges)
        except E.UserError:
            continue
        got = fits(pep)
        if got:
            d.update(pep=pep, sequence=s2, topology=t2, sheet=got[0], sheet_note=f"= {got[1]} {text}", sheet_set=True)
            return
    near = min(nums, key=lambda n: abs(n - d["pep"].mass))
    d.update(sheet=near, sheet_note=f"differs by {near - d['pep'].mass:+.3f}", sheet_bad=True)


def parse_lines(a, blocks):
    """Lines of pasted text -> [dict(name, sequence, topology, pep | error)]. Lines that hold no
    sequence at all (a header row copied along) are left out and counted in 'ignored'."""
    default = a.get("topology") if a.get("topology") in E.TOPOLOGIES else "linear"
    bridges = (a.get("bridges") or "").strip()
    out, ignored = [], 0
    for line in (a.get("text") or "").replace("\r", "").split("\n"):
        got = split_line(line, blocks)
        if got is None:
            continue
        name, seq, numbers, from_cells = got
        top = "head-to-tail" if E.CYCLO_RE.match(seq) else "linear" if has_ends(seq, blocks) else default
        d = dict(name=name, sequence=seq, topology=top)
        try:
            d["pep"] = E.parse_sequence(seq, top, blocks, bridges)
        except E.UserError as e:
            if not re.fullmatch(r"[A-Za-z0-9()\[\]+\-]+", seq) or (not re.search(r"[A-Z]", seq)):
                ignored += 1        # 'Sequences (1 AA per cell)', '#', notes ...
                continue
            d["error"] = str(e)
        if "pep" in d:
            if from_cells:
                d["sequence"] = E.read_sequence(seq, blocks)[0]
            fit_sheet_mass(d, d["sequence"], top, blocks, bridges, numbers)
        out.append(d)
    a["_ignored"] = ignored
    return out


def api_parse_sequences(a):
    blocks = blocks_any(a)
    rows = []
    for d in parse_lines(a, blocks)[:2000]:
        pep = d.pop("pep", None)
        if pep is not None:
            d.update(mass=round(pep.mass, 4), formula=E.hill(pep.formula), residues=len(pep.units))
        rows.append(d)
    return dict(rows=rows, ignored=a.get("_ignored", 0))


def api_mass(a):
    blocks = blocks_any(a)
    max_z = max(1, min(12, int(a.get("max_charge") or 4)))
    from rdkit.Chem import Descriptors
    rows = []
    for d in parse_lines(a, blocks)[:500]:
        pep = d.pop("pep", None)
        if pep is not None:
            ions = E.make_species(pep, "none", blocks, (0, 0), max_z)[0].ions
            d.update(mass=round(pep.mass, 4), formula=E.hill(pep.formula), residues=len(pep.units),
                     one_letter=E.one_letter(pep), three_letter=E.three_letter(pep, blocks),
                     ions=[[i.label, round(i.mz, 4)] for i in ions])
            try:
                d["average"] = round(Descriptors.MolWt(E.build_molecule(pep, blocks)), 2)
            except Exception:
                d["average"] = None
        rows.append(d)
    return dict(rows=rows)


def api_structure(a):
    import tempfile
    blocks = blocks_any(a)
    top = a.get("topology") if a.get("topology") in E.TOPOLOGIES else "linear"
    pep = E.parse_sequence(a.get("sequence", ""), top, blocks, a.get("bridges", ""))
    with tempfile.TemporaryDirectory() as t:
        t = Path(t)
        E.draw_structure(E.build_molecule(pep, blocks), t / "s.png", t / "s.svg", t / "s.mol", cyclic=pep.ringed)
        return dict(svg=(t / "s.svg").read_text(), mol=(t / "s.mol").read_text())


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
               dict(name="unbridged", text="bridge not formed: reduced disulfide (+2H), open lactam (+H2O)"),
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
    pdfs = sorted(out.glob("MS_report_notebook*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
    return dict(rows=rows, pptx=str(decks[0]) if decks else None, pdf=str(pdfs[0]) if pdfs else None,
                made=datetime.fromtimestamp(sp.stat().st_mtime).strftime("%d %b %Y, %H:%M"),
                out=str(out), summary=str(sp))


# ----------------------------------------------------------------------------
# interactive plots: chromatogram of a sample, spectrum at any retention time
# ----------------------------------------------------------------------------
RUNS = {}        # (path, mtime) -> Run; the last few, so clicking along a chromatogram is quick


def used_row(folder, sid):
    """The row of this sample as it was when the report was made."""
    for name in (Path(E.OUT_DIR_NAME) / "samples_used.csv", Path("samples.csv")):
        p = folder / name
        if p.exists():
            for r in E.read_samples(p):
                if r.get("id") == sid and r.get("role", "sample") != "blank":
                    return r
    raise E.UserError(f"sample {sid} is not in samples.csv")


def cached_run(folder, raw_file):
    cache = folder / E.OUT_DIR_NAME / "cache"
    npz = cache / (Path(raw_file).stem + ".npz")
    if not npz.exists():
        raise E.UserError("the converted MS data of this sample is not in MS_report/cache; make the report again")
    key = (str(npz), npz.stat().st_mtime)
    if key not in RUNS:
        while len(RUNS) >= 3:
            RUNS.pop(next(iter(RUNS)))
        with contextlib.redirect_stdout(io.StringIO()):
            RUNS[key] = E.load_run(folder / raw_file, cache)
    return RUNS[key]


def thin(x, y, n=2500):
    """Fewer points for drawing, keeping every peak top."""
    if len(x) <= n:
        return [round(float(v), 4) for v in x], [float(f"{v:.4g}") for v in y]
    k = int(-(-len(x) // n))
    idx = [s + int(y[s:s + k].argmax()) for s in range(0, len(x), k)]
    return [round(float(x[i]), 4) for i in idx], [float(f"{y[i]:.4g}") for i in idx]


def api_trace(a):
    folder = as_folder(a.get("folder"))
    sid = str(a.get("id", ""))
    r = used_row(folder, sid)
    run = cached_run(folder, r["raw_file"])
    data = folder / E.OUT_DIR_NAME / "data"
    safe = re.sub(r"[^\w.-]+", "_", sid)
    xic = None
    tp = data / f"{safe}_tic.csv"
    if tp.exists():
        with open(tp, newline="") as f:
            rows = list(csv.DictReader(f))
        if rows and "xic_expected_ions" in rows[0]:
            import numpy as np
            xic = thin(np.array([float(x["rt_min"]) for x in rows]), np.array([float(x["xic_expected_ions"]) for x in rows]))
    blank = None
    bp = data / f"{safe}_blank_tic.csv"
    if bp.exists():
        import numpy as np
        with open(bp, newline="") as f:
            rows = list(csv.DictReader(f))
        blank = thin(np.array([float(x["rt_min"]) for x in rows]), np.array([float(x["tic"]) for x in rows]))
    return dict(tic=thin(run.rt, run.tic), xic=xic, blank=blank, raw=r["raw_file"],
                rt_range=[float(run.rt[0]), float(run.rt[-1])])


def api_spectrum(a):
    import numpy as np
    folder = as_folder(a.get("folder"))
    r = used_row(folder, str(a.get("id", "")))
    run = cached_run(folder, r["raw_file"])
    k = int(np.argmin(np.abs(run.rt - float(a.get("rt", 0)))))
    lo, hi = max(0, k - 3), min(len(run.rt) - 1, k + 3)
    smz, sint = E.averaged_spectrum(run, lo, hi)
    if len(smz) > 5000:
        keep = np.sort(np.argsort(sint)[-5000:])
        smz, sint = smz[keep], sint[keep]
    ions = []
    try:
        blocks = blocks_any(a)
        pep = E.parse_sequence(r["sequence"], r["topology"], blocks, r.get("bridges", ""))
        for sp in E.make_species(pep, r.get("find_deltas", ""), blocks, (run.scan_lo, run.scan_hi), 6):
            for ion in sp.ions:
                j = E.match_in_spectrum(smz, sint, ion, 5.0) if len(smz) else None
                if j is not None:
                    name = ion.label if sp.name == "main" or sp.carrier else f"{sp.name} {ion.label}"
                    ions.append(dict(mz=round(float(smz[j]), 4), calc=round(ion.mz, 4), label=name, main=sp.name == "main"))
    except E.UserError:
        pass
    return dict(mz=[round(float(v), 4) for v in smz], inten=[float(f"{v:.4g}") for v in sint], ions=ions,
                rt=float(run.rt[k]), rt_from=float(run.rt[lo]), rt_to=float(run.rt[hi]),
                mz_range=[run.scan_lo, min(run.scan_hi, 1e6)])


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
        if a.get("open"):
            show_folder(a["open"])
    return {}


def show_folder(path):
    """Open a dropped folder (or the folder of a dropped file) in the window."""
    p = Path(path).expanduser()
    folder = p if p.is_dir() else p.parent
    if WINDOW is not None and folder.is_dir():
        WINDOW.evaluate_js(f"dropFolder({json.dumps(str(folder))})")


def bind_drop():
    """Dropping a folder on the window opens it (pywebview gives the full path of dropped items)."""
    try:
        from webview.dom import DOMEventHandler

        def on_drop(e):
            files = (e.get("dataTransfer") or {}).get("files") or []
            paths = [f.get("pywebviewFullPath") for f in files if f.get("pywebviewFullPath")]
            if paths:
                show_folder(paths[0])
        doc = WINDOW.dom.document
        doc.events.dragover += DOMEventHandler(lambda e: None, True, True)
        doc.events.drop += DOMEventHandler(on_drop, True, True)
    except Exception as e:      # an older pywebview: Choose folder still works
        print(f"drag and drop not available: {e}", flush=True)


# ----------------------------------------------------------------------------
# is there a newer version?
# ----------------------------------------------------------------------------
UPDATE = {"latest": None}


def check_update():
    """Reads the VERSION line of the published script. Nothing about this computer is sent.
    MS_REPORT_NO_UPDATE_CHECK=1 turns it off."""
    if os.environ.get("MS_REPORT_NO_UPDATE_CHECK"):
        return
    try:
        import urllib.request
        req = urllib.request.Request(f"https://raw.githubusercontent.com/{E.REPO}/main/make_ms_report.py",
                                     headers={"User-Agent": "MS-Report"})
        text = urllib.request.urlopen(req, timeout=6).read(400000).decode("utf-8", "replace")
        m = re.search(r'^VERSION = "([0-9.]+)"', text, re.M)
        if m:
            UPDATE["latest"] = m.group(1)
    except Exception:
        pass


def version_tuple(v):
    return tuple(int(x) for x in re.findall(r"\d+", v or "")[:4])


def api_update(a):
    latest = UPDATE["latest"]
    return dict(current=E.VERSION, latest=latest, url=f"https://github.com/{E.REPO}",
                newer=bool(latest) and version_tuple(latest) > version_tuple(E.VERSION))


def api_open_update(a):
    subprocess.run(["open", f"https://github.com/{E.REPO}"])
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


def already_running(folder=None):
    """Another dashboard is open: bring it to the front (and open this folder in it) and say so."""
    try:
        info = json.loads(LOCK.read_text())
        os.kill(int(info["pid"]), 0)
    except Exception:
        return False
    import urllib.request
    try:
        body = json.dumps({"open": str(Path(folder).expanduser().resolve())} if folder else {}).encode()
        req = urllib.request.Request(f"http://127.0.0.1:{info['port']}/api/focus", data=body,
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
    if not args.no_open and already_running(args.folder):
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
    threading.Thread(target=check_update, daemon=True).start()
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
            WINDOW.events.loaded += bind_drop
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
