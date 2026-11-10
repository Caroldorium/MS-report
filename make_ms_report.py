#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10,<3.14"
# dependencies = ["numpy", "openpyxl", "pyteomics", "psims", "lxml", "rdkit", "python-pptx", "matplotlib"]
# ///
"""make_ms_report.py - one-command LC-MS report for synthesised peptides.

Copy this file into a folder of Thermo .raw files and run it there.
Run `python3 make_ms_report.py help` for the full guide.
"""

GUIDE = r"""
make_ms_report.py - LC-MS report for synthesised peptides
==========================================================

HOW TO USE
  1. Copy make_ms_report.py into the folder that holds the .raw files
     (and the synthesis .xlsx, or a .csv of sequences).
  2. In Terminal:
        cd "<that folder>"
        python3 make_ms_report.py
     First time: no samples.csv yet, so the script WRITES samples.csv from the
     .xlsx/.csv it finds and matches the .raw files to it, then stops.
  3. Open samples.csv, check every row, fill the blanks (topology is required),
     save, and run the same command again. The report is written to MS_report/.

COMMANDS
  python3 make_ms_report.py            setup if there is no samples.csv, otherwise run
  python3 make_ms_report.py help       this guide
  python3 make_ms_report.py setup      (re)write samples.csv from the .xlsx/.csv here
                                       (an existing one is kept as samples_old.csv)
  python3 make_ms_report.py check      check samples.csv and print calculated m/z; no MS data read
  python3 make_ms_report.py run        make the report
  python3 make_ms_report.py mass SEQ   quick calculator, e.g.
        python3 make_ms_report.py mass RMNEILDFIDPRFISNDP --topology head-to-tail
        python3 make_ms_report.py mass "Ac-QHREDGSGSGK[Biotin]-NH2"
  python3 make_ms_report.py blocks     list the building blocks the script knows
  python3 make_ms_report.py deltas     list what find_deltas can search for

OPTIONS (for run)
  --only 1,5,13       only these sample ids
  --ppm 5             m/z tolerance in ppm (default 5; this Orbitrap is ~2 ppm)
  --max-charge 6      highest charge state to look for (default 6)
  --min-rel 1         an ion counts only if it is at least this % of the base peak
                      in its averaged spectrum (default 1)
  --min-scans 8       the ions must be seen in at least this many consecutive scans
                      (default 8, about 1 s on this method)

HOW "FOUND" IS DECIDED
  An ion counts when its mono peak, first 13C peak (at the ratio the formula predicts)
  and, for larger peptides, the second 13C peak are all within --ppm, for --min-scans
  scans. summary.csv then says
    yes   2 or more charge states, or one charge state at 5% of the base peak or more
    weak  one charge state under 5%. Unrelated sequences reach this level by chance,
          so treat weak as not proven.
    no    nothing that passes
  --folder PATH       work on another folder instead of the one this script is in

OUTPUT  (MS_report/ inside the data folder)
  MS_report_notebook.pdf   the same report as a PDF for the lab notebook: an overview page,
                      then one page per sample (not editable; the .pptx is the working copy)
  MS_report.pptx      an overview slide (every sample: found or not, retention time, ppm error,
                      purity), then one slide per sample, everything editable:
                        A structure picture, name and sequence (text)
                        B HPLC trace (native chart); without an hplc_file, a zoom on the
                          product's isotope cluster, each peak labelled with its m/z
                        C calculated vs observed m/z table with retention times (native table)
                        D MS spectrum averaged over the peptide's elution peak, full scan
                          range, Xcalibur style: expected ions labelled in red, other
                          major peaks in green (native chart)
                        E TIC, Xcalibur style, retention times on the peaks, blank in grey;
                          red line where the product elutes, grey line (TIC maximum) when it
                          is not found; spectrum_rt in samples.csv moves it (native chart)
                      Right-click a chart > Edit Data to change it; all axes/colours editable.
                      Labels are data labels: click one to move, edit or delete it.
  summary.csv         one line per sample: found or not, retention time, ppm error, purity
  data/<id>_*.csv     the numbers behind every panel (paste into Prism if you like);
                      _tic.csv also has the XIC of the expected ions, _blank_tic.csv the blank
  data/<id>_main_peaks.csv  the biggest LC peaks: neutral mass, difference from the
                      expected mass, and possible explanations (protecting groups, adducts,
                      side reactions from the Novabiochem delta-mass sheet, a missing or extra
                      residue, two of these together, a dimer, or another peptide of the run)
  data/<id>.mol       structure for ChemDraw (open, then Clean Up Structure)
  data/<id>.svg       structure as vector art (PowerPoint: Insert > Pictures, then Convert to Shape)
  cache/              converted MS data; delete it to force re-reading the .raw files

samples.csv COLUMNS
  id            sample id (any text); used in file names
  name          slide title
  sequence      one letter per residue; non-canonical in parentheses (Aib);
                side-chain tag in square brackets after its residue K[Biotin];
                ends of a linear peptide with dashes: Ac-QHREDGSGSGK[TAMRA]-NH2
                (a linear peptide with no ends written is H-...-OH)
                or three-letter codes with dashes: H-Ala-Gly-Lys(Biotin)-NH2,
                cyclo(-Ala-Gly-Pro-); D residues as D-Phe need their own row
  topology      linear | head-to-tail   (required; the backbone)
  bridges       optional side-chain bridges, by residue position, separated by ", "
                  C11-C18 (or 11-18)   disulfide between two Cys (-2H)
                  K3-D7                lactam between Lys and Asp/Glu (-H2O)
                  4-8                  hydrocarbon staple between two olefin residues such as
                                       (S5) or (R8), closed by metathesis (-C2H4)
                  disulfide            short for the only two Cys of the peptide
                The kind follows from the two residues. Works with either topology.
                e.g. thanatin: GSKKPVPIIYCNRRTGKCQRM, linear, C11-C18
  raw_file     the .raw (or .mzML) file for this sample
  blank_file    the blank (solvent) run for this sample, any .raw in the folder. Its TIC is
                drawn in grey behind panel E, it is subtracted from the spectrum in panel D,
                and main peaks also in the blank are marked. Empty = no blank.
                auto = the first row with role blank (what an older sheet without this column gets)
  hplc_file     optional: a 2-column text/csv export (time, absorbance). Panel B then shows
                the trace and the purity: area of the product peak as a percentage of all
                integrated peaks (baseline in grey, product peak in red; the peak list is in
                data/<id>_hplc_peaks.csv). Area % at the recorded wavelength, nothing more.
  purity_rt     optional: retention time (min) of the product peak in the HPLC trace.
                Empty = the peak with the largest area.
  spectrum_rt   optional: retention time (min) for the red line in panel E and the spectrum in
                panel D. Empty = at the expected product if found, else at the TIC maximum (grey).
  role          sample | blank     (blank rows get no slide; pick them in blank_file)
  observed_mz   optional: m/z values you read yourself (e.g. in FreeStyle), separated by ;
                Each goes to the nearest calculated ion (within 20 ppm) and is shown in
                table C instead of the value the script found. Setup fills this from a
                "Found m/z" / "Observed" column of the .xlsx or .csv if there is one.
  find_deltas   what else to search for besides the expected peptide, separated by ", "
                  auto     everything below that fits the sequence (default)
                  all      every delta, whatever the sequence (more chance hits)
                  none     only the expected peptide
                  linear, unbridged, deletions, insertions, dimer, adducts, side_reactions,
                  protecting_groups, or your own group from building_blocks.csv
                  +Na, +Pbf, ...  one delta by its code
                (python3 make_ms_report.py deltas lists them all)
  notes         free text; copied into the slide notes

BUILDING BLOCKS (building_blocks.csv)
  Looked for in: this folder, then the master Script folder
  (%MASTER%), else a built-in copy is written here.
  To add a monomer, add a row: code, code3, name, kind, formula, smiles.
    kind     residue | n_cap | c_cap | tag | adduct | shift
    formula  what the unit adds to the peptide (residue = amino acid minus H2O;
             tag = group added minus the H it replaces)
    smiles   with attachment points: [*:1] N side, [*:2] C side, [*:3] side chain
             (tags and caps use [*:1] / [*:2]). The script refuses a row whose
             smiles and formula disagree, so drawing and mass always match.
  D-amino acids: add rows, e.g. code dF with the D smiles, and write (dF).

REQUIREMENTS
  Python with rdkit pyteomics psims python-pptx numpy openpyxl. If they are missing, the
  script switches by itself to a conda env called 'msreport' (or to MS_REPORT_PYTHON).
  ThermoRawFileParser in ~/software/ThermoRawFileParser (only needed for .raw files;
  THERMORAWFILEPARSER=/path/to/it points somewhere else).
  --ppm: 5 suits Orbitrap and FT-ICR data, about 20 a Q-TOF. Spectra must be centroided.
"""

import os
import sys
from pathlib import Path

VERSION = "1.1.0"
REPO = "Caroldorium/MS-report"       # where new versions and the issue tracker live


def _master_dir():
    """Where the shared building_blocks.csv and the app's settings live: MS_REPORT_HOME if set,
    else this script's folder when it holds a building_blocks.csv, else Documents/MS Report
    (the app's folder) when it exists, else this script's folder."""
    if os.environ.get("MS_REPORT_HOME"):
        return Path(os.environ["MS_REPORT_HOME"]).expanduser()
    here = Path(__file__).resolve().parent
    app_home = Path.home() / "Documents" / "MS Report"
    if (here / "building_blocks.csv").exists() or not app_home.is_dir():
        return here
    return app_home


def _env_python():
    """A conda environment called 'msreport', wherever conda is installed (or MS_REPORT_PYTHON)."""
    if os.environ.get("MS_REPORT_PYTHON"):
        return Path(os.environ["MS_REPORT_PYTHON"]).expanduser()
    roots = [Path.home() / n for n in ("miniconda3", "anaconda3", "miniforge3", "mambaforge", "micromamba")]
    roots += [Path("/opt/anaconda3"), Path("/opt/miniconda3"), Path("/opt/homebrew/Caskroom/miniconda/base"),
              Path("/opt/homebrew/Caskroom/miniforge/base")]
    for r in roots:
        p = r / "envs" / "msreport" / "bin" / "python"
        if p.exists():
            return p
    return Path("/nonexistent")


MASTER_DIR = _master_dir()
ENV_PYTHON = _env_python()
CONVERTER_CANDIDATES = [
    Path.home() / "software" / "ThermoRawFileParser" / "ThermoRawFileParser",
]
OUT_DIR_NAME = "MS_report"
FONT = "Arial"                                                   # every text on the slides
STRUCTURE_FONT = "/System/Library/Fonts/Supplemental/Arial.ttf"  # atom labels in structures

if len(sys.argv) > 1 and sys.argv[1] in ("help", "-h", "--help"):
    print(GUIDE.replace("%MASTER%", str(MASTER_DIR)))
    sys.exit(0)

try:
    import numpy as np
    import openpyxl
    from pyteomics import mzml  # noqa: F401
    import psims  # noqa: F401  (pyteomics needs it for mzML)
    from rdkit import Chem, RDLogger
    from pptx import Presentation  # noqa: F401
except ImportError as exc:
    if ENV_PYTHON.exists() and Path(sys.executable).resolve() != ENV_PYTHON.resolve():
        os.execv(str(ENV_PYTHON), [str(ENV_PYTHON), *sys.argv])
    sys.exit(
        f"Missing package ({exc.name}). Create the environment once with:\n"
        "  conda create -n msreport -c conda-forge python=3.12 rdkit pyteomics psims "
        "python-pptx numpy openpyxl\n"
    )

import argparse
import csv
import io
import re
import shutil
import subprocess
import warnings
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime

warnings.filterwarnings("ignore")
RDLogger.DisableLog("rdApp.*")

# ----------------------------------------------------------------------------
# Constants
# ----------------------------------------------------------------------------
PROTON = 1.007276466812
ELECTRON = 0.000548579909
C13_SHIFT = 1.0033548378
MONO = {
    "H": 1.00782503207, "C": 12.0, "N": 14.0030740048, "O": 15.99491461956,
    "S": 31.97207100, "P": 30.97376163, "Na": 22.9897692809, "K": 38.96370668,
    "Cl": 34.96885268, "F": 18.99840322, "Br": 78.9183371, "I": 126.904473,
    "Se": 79.9165213, "Si": 27.9769265325, "B": 11.0093054, "Li": 7.01600455,
}

OPTIONAL_COLUMNS = ["observed_mz", "find_deltas", "blank_file", "spectrum_rt", "bridges", "purity_rt"]
SAMPLE_COLUMNS = ["id", "name", "sequence", "topology", "bridges", "raw_file", "blank_file", "spectrum_rt", "hplc_file",
                  "purity_rt",
                  "role", "find_deltas", "observed_mz", "notes"]
TOPOLOGIES = ("linear", "head-to-tail")

DEFAULT_BLOCKS = """code,code3,name,kind,formula,smiles,group,residues,notes
G,Gly,Glycine,residue,C2H3NO,[*:1]NCC(=O)[*:2],,,
A,Ala,Alanine,residue,C3H5NO,[*:1]N[C@@H](C)C(=O)[*:2],,,
S,Ser,Serine,residue,C3H5NO2,[*:1]N[C@@H](CO)C(=O)[*:2],,,
P,Pro,Proline,residue,C5H7NO,[*:1]N1CCC[C@H]1C(=O)[*:2],,,
V,Val,Valine,residue,C5H9NO,[*:1]N[C@@H](C(C)C)C(=O)[*:2],,,
T,Thr,Threonine,residue,C4H7NO2,[*:1]N[C@@H]([C@@H](C)O)C(=O)[*:2],,,
C,Cys,Cysteine,residue,C3H5NOS,[*:1]N[C@@H](CS[*:3])C(=O)[*:2],,,[*:3] = side-chain attachment for tags
L,Leu,Leucine,residue,C6H11NO,[*:1]N[C@@H](CC(C)C)C(=O)[*:2],,,
I,Ile,Isoleucine,residue,C6H11NO,[*:1]N[C@@H]([C@@H](C)CC)C(=O)[*:2],,,
N,Asn,Asparagine,residue,C4H6N2O2,[*:1]N[C@@H](CC(N)=O)C(=O)[*:2],,,
D,Asp,Aspartic acid,residue,C4H5NO3,[*:1]N[C@@H](CC(=O)O)C(=O)[*:2],,,
Q,Gln,Glutamine,residue,C5H8N2O2,[*:1]N[C@@H](CCC(N)=O)C(=O)[*:2],,,
K,Lys,Lysine,residue,C6H12N2O,[*:1]N[C@@H](CCCCN[*:3])C(=O)[*:2],,,[*:3] = side-chain attachment for tags
E,Glu,Glutamic acid,residue,C5H7NO3,[*:1]N[C@@H](CCC(=O)O)C(=O)[*:2],,,
M,Met,Methionine,residue,C5H9NOS,[*:1]N[C@@H](CCSC)C(=O)[*:2],,,
H,His,Histidine,residue,C6H7N3O,[*:1]N[C@@H](Cc1c[nH]cn1)C(=O)[*:2],,,
F,Phe,Phenylalanine,residue,C9H9NO,[*:1]N[C@@H](Cc1ccccc1)C(=O)[*:2],,,
R,Arg,Arginine,residue,C6H12N4O,[*:1]N[C@@H](CCCNC(=N)N)C(=O)[*:2],,,
Y,Tyr,Tyrosine,residue,C9H9NO2,[*:1]N[C@@H](Cc1ccc(O)cc1)C(=O)[*:2],,,
W,Trp,Tryptophan,residue,C11H10N2O,[*:1]N[C@@H](Cc1c[nH]c2ccccc12)C(=O)[*:2],,,
Aib,Aib,2-Aminoisobutyric acid,residue,C4H7NO,[*:1]NC(C)(C)C(=O)[*:2],,,
Mox,Mox,Methoxinine (O-methylhomoserine),residue,C5H9NO2,[*:1]N[C@@H](CCOC)C(=O)[*:2],,,
Nle,Nle,Norleucine,residue,C6H11NO,[*:1]N[C@@H](CCCC)C(=O)[*:2],,,
Nva,Nva,Norvaline,residue,C5H9NO,[*:1]N[C@@H](CCC)C(=O)[*:2],,,
Abu,Abu,2-Aminobutyric acid,residue,C4H7NO,[*:1]N[C@@H](CC)C(=O)[*:2],,,
Orn,Orn,Ornithine,residue,C5H10N2O,[*:1]N[C@@H](CCCN[*:3])C(=O)[*:2],,,[*:3] = side-chain attachment for tags
Dab,Dab,"2,4-Diaminobutyric acid",residue,C4H8N2O,[*:1]N[C@@H](CCN[*:3])C(=O)[*:2],,,[*:3] = side-chain attachment for tags
Dap,Dap,"2,3-Diaminopropionic acid",residue,C3H6N2O,[*:1]N[C@@H](CN[*:3])C(=O)[*:2],,,[*:3] = side-chain attachment for tags
Cit,Cit,Citrulline,residue,C6H11N3O2,[*:1]N[C@@H](CCCNC(N)=O)C(=O)[*:2],,,
Hyp,Hyp,4-Hydroxyproline,residue,C5H7NO2,[*:1]N1C[C@H](O)C[C@H]1C(=O)[*:2],,,
Sar,Sar,Sarcosine (N-methylglycine),residue,C3H5NO,[*:1]N(C)CC(=O)[*:2],,,
S5,S5,(S)-2-(4-Pentenyl)alanine (stapling),residue,C8H13NO,[*:1]N[C@@](C)(CCCC=C)C(=O)[*:2],,,olefin for a hydrocarbon staple
R5,R5,(R)-2-(4-Pentenyl)alanine (stapling),residue,C8H13NO,[*:1]N[C@](C)(CCCC=C)C(=O)[*:2],,,olefin for a hydrocarbon staple
S8,S8,(S)-2-(7-Octenyl)alanine (stapling),residue,C11H19NO,[*:1]N[C@@](C)(CCCCCCC=C)C(=O)[*:2],,,olefin for a hydrocarbon staple
R8,R8,(R)-2-(7-Octenyl)alanine (stapling),residue,C11H19NO,[*:1]N[C@](C)(CCCCCCC=C)C(=O)[*:2],,,olefin for a hydrocarbon staple
H-,,Free amine (N-terminus),n_cap,H,[H][*:1],,,default N-terminus of a linear peptide
Ac-,,Acetyl (N-terminus),n_cap,C2H3O,CC(=O)[*:1],,,
-OH,,Free acid (C-terminus),c_cap,HO,O[*:2],,,default C-terminus of a linear peptide
-NH2,,Amide (C-terminus),c_cap,H2N,N[*:2],,,
Biotin,,Biotin (amide),tag,C10H14N2O2S,[*:1]C(=O)CCCC[C@@H]1SC[C@@H]2NC(=O)N[C@H]12,,,acylates the side-chain amine
TAMRA,,5-Carboxytetramethylrhodamine (amide),tag,C25H20N2O4,[*:1]C(=O)c1ccc(-c2c3ccc(N(C)C)cc3oc3cc(=[N+](C)C)ccc23)c(C(=O)[O-])c1,,,5-isomer; zwitterion so M is neutral
+Na,,Sodium adduct,adduct,Na,,adducts,,[M+zH] becomes [M+(z-1)H+Na]
+K,,Potassium adduct,adduct,K,,adducts,,[M+zH] becomes [M+(z-1)H+K]
+CO2,,CO2 adduct (e.g. Trp carbamate),shift,CO2,,side_reactions,W,neutral mass shift
-H2O,,Water loss,shift,-H2O,,side_reactions,,neutral mass shift
+O,,Oxidation (e.g. Met),shift,O,,side_reactions,MWC,neutral mass shift
+NH3,,Ammonium adduct ([M+NH4]+ vs [M+H]+),shift,NH3,,adducts,,Novabiochem delta mass sheet
-NH3,,Pyroglutamate from N-terminal Gln; aspartimide from Asn,shift,-NH3,,side_reactions,QN,Novabiochem delta mass sheet
+deamid,,Deamidation of Asn/Gln; Cit from Arg,shift,O-HN,,side_reactions,NQR,Novabiochem delta mass sheet
+O2,,Oxidation of Met to sulfone,shift,O2,,side_reactions,M,Novabiochem delta mass sheet
+O3,,Cysteic acid from Cys,shift,O3,,side_reactions,C,Novabiochem delta mass sheet
+CH2,,Methyl ester,shift,CH2,,side_reactions,DE,Novabiochem delta mass sheet
+C2H4,,Ethyl,shift,C2H4,,side_reactions,DE,Novabiochem delta mass sheet
+CO,,Formyl,shift,CO,,side_reactions,KW,Novabiochem delta mass sheet
+Ac,,Acetyl (e.g. capping),shift,C2H2O,,side_reactions,,Novabiochem delta mass sheet
+Tfa,,Trifluoroacetyl,shift,C2F3O-H,,side_reactions,KSTY,Novabiochem delta mass sheet
+TMG,,Tetramethylguanidinium cap from HBTU/HATU,shift,C5H10N2,,side_reactions,,Novabiochem delta mass sheet
+pip,,Asp piperidide,shift,C5H9N-O,,side_reactions,D,Novabiochem delta mass sheet
+pipAla,,3-(1-Piperidinyl)alanine from C-terminal Cys,shift,C5H9N-S,,side_reactions,C,Novabiochem delta mass sheet
+tBu,,"tBu left on, or tBu alkylation",shift,C4H8,,protecting_groups,STDEYCMW,Novabiochem delta mass sheet
+Boc,,Boc left on,shift,C5H8O2,,protecting_groups,KWH,Novabiochem delta mass sheet
+Trt,,Trt left on,shift,C19H14,,protecting_groups,NQHC,Novabiochem delta mass sheet
+Pbf,,Pbf left on (Arg),shift,C13H16O3S,,protecting_groups,R,Novabiochem delta mass sheet
+Pmc,,Pmc left on (Arg),shift,C14H18O3S,,protecting_groups,R,Novabiochem delta mass sheet
+Mtt,,Mtt left on,shift,C20H16,,protecting_groups,K,Novabiochem delta mass sheet
+Mmt,,Mmt left on,shift,C20H16O,,protecting_groups,CK,Novabiochem delta mass sheet
+Fmoc,,Fmoc left on,shift,C15H10O2,,protecting_groups,,Novabiochem delta mass sheet
+Acm,,Acm left on (Cys),shift,C3H5NO,,protecting_groups,C,Novabiochem delta mass sheet
+StBu,,StBu left on (Cys),shift,C4H8S,,protecting_groups,C,Novabiochem delta mass sheet
+Bzl,,Benzyl,shift,C7H6,,protecting_groups,CWY,Novabiochem delta mass sheet
+Alloc,,Alloc left on,shift,C4H4O2,,protecting_groups,K,Novabiochem delta mass sheet
+ivDde,,ivDde left on,shift,C13H18O2,,protecting_groups,K,Novabiochem delta mass sheet
+Dmb,,Dmb left on,shift,C9H10O2,,protecting_groups,G,Novabiochem delta mass sheet
+Hmb,,Hmb left on,shift,C8H8O2,,protecting_groups,G,Novabiochem delta mass sheet
+SO3,,"Sulfonation from Pbf/Pmc (Arg, Ser, Thr)",shift,SO3,,side_reactions,RSTY,Novabiochem delta mass sheet
+EDT,,EDT disulfide adduct,shift,C2H4S2,,side_reactions,C,Novabiochem delta mass sheet
+HMP,,HMP linker adduct (Wang resin),shift,C7H6O,,side_reactions,,Novabiochem delta mass sheet
"""


class UserError(Exception):
    """A problem the user fixes in a file; printed without a traceback."""


# ----------------------------------------------------------------------------
# Formulas
# ----------------------------------------------------------------------------
def parse_formula(text):
    """'C2H3NO', '-H2O' (a loss) or 'C5H9N-O' (gain C5H9N, lose O)."""
    text = text.strip()
    if "-" in text[1:]:
        gain, loss = text.split("-", 1) if not text.startswith("-") else ("", text[1:])
        c = parse_formula(gain) if gain else Counter()
        c.subtract(parse_formula(loss))
        return c
    sign = 1
    if text.startswith("-"):
        sign, text = -1, text[1:]
    elif text.startswith("+"):
        text = text[1:]
    if not re.fullmatch(r"([A-Z][a-z]?\d*)+", text):
        raise UserError(f"Cannot read formula '{text}'")
    c = Counter()
    for el, n in re.findall(r"([A-Z][a-z]?)(\d*)", text):
        if el not in MONO:
            raise UserError(f"Element '{el}' has no mass in the script (formula {text})")
        c[el] += sign * int(n or 1)
    return c


def formula_mass(c):
    return sum(MONO[e] * n for e, n in c.items())


def hill(c):
    c = {e: n for e, n in c.items() if n}
    if c and all(n < 0 for n in c.values()):      # a pure loss, e.g. -H2O
        return "-" + hill({e: -n for e, n in c.items()})
    if any(n < 0 for n in c.values()):            # mixed change, e.g. C5H9N-O
        return hill({e: n for e, n in c.items() if n > 0}) + "-" + hill({e: -n for e, n in c.items() if n < 0})
    if "C" in c:   # Hill order: C, H, then the rest alphabetically
        order = ["C"] + (["H"] if "H" in c else []) + sorted(e for e in c if e not in ("C", "H"))
    else:
        order = sorted(c)
    return "".join(f"{e}{c[e] if c[e] != 1 else ''}" for e in order)


def rdkit_formula(mol):
    from rdkit.Chem.rdMolDescriptors import CalcMolFormula
    f = CalcMolFormula(mol)
    f = re.sub(r"[+-]\d*$", "", f)
    return parse_formula(f)


# ----------------------------------------------------------------------------
# Building blocks
# ----------------------------------------------------------------------------
@dataclass
class Block:
    code: str
    code3: str
    name: str
    kind: str
    formula: Counter
    smiles: str
    mass: float
    group: str = ""      # deltas only: adducts, side_reactions, protecting_groups, or your own
    residues: str = ""   # deltas only: one-letter codes it can happen on (blank = any peptide)


def find_blocks_file(folder):
    for p in (folder / "building_blocks.csv", MASTER_DIR / "building_blocks.csv"):
        if p.exists():
            return p
    p = folder / "building_blocks.csv"
    p.write_text(DEFAULT_BLOCKS)
    print(f"No building_blocks.csv found; wrote the built-in one to {p}")
    return p


def load_blocks(path):
    blocks, problems = {}, []
    with open(path, newline="", encoding="utf-8-sig") as f:
        for i, r in enumerate(csv.DictReader(f), start=2):
            code = (r.get("code") or "").strip()
            if not code:
                continue
            kind = (r.get("kind") or "").strip()
            if kind not in ("residue", "n_cap", "c_cap", "tag", "adduct", "shift"):
                problems.append(f"line {i} ({code}): unknown kind '{kind}'")
                continue
            try:
                fc = parse_formula(r["formula"])
            except UserError as e:
                problems.append(f"line {i} ({code}): {e}")
                continue
            smi = (r.get("smiles") or "").strip()
            if smi:
                msg = check_block_smiles(kind, smi, fc)
                if msg:
                    problems.append(f"line {i} ({code}): {msg}")
                    continue
            elif kind in ("residue", "n_cap", "c_cap", "tag"):
                problems.append(f"line {i} ({code}): needs a smiles")
                continue
            if code in blocks:
                problems.append(f"line {i}: code '{code}' is used twice")
            group = (r.get("group") or "").strip()
            if group and kind not in ("shift", "adduct"):
                problems.append(f"line {i} ({code}): only shift/adduct rows can have a group")
            if group and not re.fullmatch(r"[a-z0-9_]+", group):
                problems.append(f"line {i} ({code}): group '{group}' should be lower case, e.g. my_group")
            if group in BUILTIN_DELTA_GROUPS + ("auto", "all", "none"):
                problems.append(f"line {i} ({code}): '{group}' is a built-in name; choose another group name")
            blocks[code] = Block(code, (r.get("code3") or "").strip(), (r.get("name") or "").strip(),
                                 kind, fc, smi, formula_mass(fc), group,
                                 re.sub(r"[^A-Za-z]", "", r.get("residues") or ""))
    if problems:
        raise UserError(f"Problems in {path}:\n  " + "\n  ".join(problems))
    return blocks


def check_block_smiles(kind, smi, fc):
    """The formula column must equal what the smiles adds to a peptide."""
    mol = Chem.MolFromSmiles(smi)
    if mol is None:
        return "smiles cannot be read"
    maps = sorted(a.GetAtomMapNum() for a in mol.GetAtoms() if a.GetAtomicNum() == 0)
    need = {"residue": [1, 2], "n_cap": [1], "c_cap": [2], "tag": [1]}[kind]
    if [m for m in maps if m in (1, 2)] != need:
        return f"smiles needs attachment point(s) {', '.join(f'[*:{m}]' for m in need)}"
    rw = Chem.RWMol(mol)
    for a in rw.GetAtoms():
        if a.GetAtomicNum() == 0:
            a.SetAtomicNum(1)
            a.SetAtomMapNum(0)
            a.SetIsotope(0)
    m2 = rw.GetMol()
    Chem.SanitizeMol(m2)
    got = rdkit_formula(Chem.RemoveHs(m2))
    got["H"] -= {"residue": 2, "n_cap": 1, "c_cap": 1, "tag": 2}[kind]
    got = Counter({e: n for e, n in got.items() if n})
    want = Counter({e: n for e, n in fc.items() if n})
    if got != want:
        return f"smiles gives {hill(got)} but formula says {hill(want)}"
    return None


def residue_lookup(blocks):
    look = {}
    for b in blocks.values():
        if b.kind == "residue":
            look[b.code] = b
            if b.code3:
                look.setdefault(b.code3, b)
                look.setdefault(b.code3.lower(), b)
    return look


# ----------------------------------------------------------------------------
# Sequences
# ----------------------------------------------------------------------------
@dataclass
class Unit:
    code: str
    tag: str = None


@dataclass
class Peptide:
    units: list
    topology: str
    n_cap: str = None
    c_cap: str = None
    formula: Counter = field(default_factory=Counter)
    bridges: list = field(default_factory=list)   # [(i, j, "disulfide" | "lactam")], 0-based residue indices

    @property
    def mass(self):
        return formula_mass(self.formula)

    @property
    def ringed(self):
        return self.topology == "head-to-tail" or bool(self.bridges)


SIDE_ACID = "[CX3](=O)[OX2H1]"
SIDE_OLEFIN = "[CX3;H1;!R]=[CX3;H2]"                    # terminal vinyl, e.g. on S5 / R8
BRIDGE_LOSS = {"disulfide": "H2", "lactam": "H2O", "staple": "C2H4"}       # what closing the bridge removes
BRIDGE_OPEN = {"disulfide": "reduced", "lactam": "open lactam", "staple": "unstapled"}


def side_groups(block):
    """What a residue's side chain can bridge with: thiol / amine (at its [*:3]) and acid (a free COOH)."""
    mol = Chem.MolFromSmiles(block.smiles)
    out = set()
    for a in mol.GetAtoms():
        if a.GetAtomicNum() == 0 and a.GetAtomMapNum() == 3:
            sym = a.GetNeighbors()[0].GetSymbol()
            if sym in ("S", "Se"):
                out.add("thiol")
            elif sym == "N":
                out.add("amine")
    if mol.HasSubstructMatch(Chem.MolFromSmarts(SIDE_ACID)):
        out.add("acid")
    if mol.HasSubstructMatch(Chem.MolFromSmarts(SIDE_OLEFIN)):
        out.add("olefin")
    return out


def parse_bridges(text, units, blocks):
    """'C11-C18' / '11-18, 2-9' / 'disulfide' -> [(i, j, kind)]; the kind follows from the two residues."""
    toks = split_list(text)
    if not toks:
        return []
    groups = [side_groups(blocks[u.code]) for u in units]

    def name(i):
        c = units[i].code
        return (c if len(c) == 1 else f"({c})") + str(i + 1)
    pairs = []
    for tok in toks:
        if tok.lower() in ("disulfide", "disulphide", "ss", "s-s"):
            thiols = [i for i, g in enumerate(groups) if "thiol" in g and not units[i].tag]
            if len(thiols) != 2:
                raise UserError(f"bridges: 'disulfide' alone needs exactly two free Cys (this peptide has "
                                f"{len(thiols)}); write the positions, e.g. C2-C9, C5-C14")
            pairs.append((thiols[0], thiols[1], tok))
            continue
        m = re.fullmatch(r"(\(?[A-Za-z0-9]*?\)?)(\d+)\s*[-–:]\s*(\(?[A-Za-z0-9]*?\)?)(\d+)", tok)
        if not m:
            raise UserError(f"bridges: cannot read '{tok}'. Write two residue positions, e.g. C11-C18 or 11-18")
        ends = []
        for letters, num in ((m.group(1), m.group(2)), (m.group(3), m.group(4))):
            i = int(num) - 1
            if not 0 <= i < len(units):
                raise UserError(f"bridges: position {num} in '{tok}' is outside the peptide ({len(units)} residues)")
            b = blocks[units[i].code]
            want = letters.strip("()")
            if want and want not in (b.code, b.code3, (b.code3 or "").lower()):
                raise UserError(f"bridges: '{tok}' says {want}{num}, but residue {num} is {b.code3 or b.code}")
            ends.append(i)
        if ends[0] == ends[1]:
            raise UserError(f"bridges: '{tok}' joins a residue to itself")
        pairs.append((min(ends), max(ends), tok))
    out, used = [], {}
    for i, j, tok in pairs:
        for k in (i, j):
            if k in used:
                raise UserError(f"bridges: residue {name(k)} is in two bridges ({used[k]} and {tok})")
            if units[k].tag:
                raise UserError(f"bridges: residue {name(k)} carries the tag [{units[k].tag}], so it cannot bridge")
            used[k] = tok
        gi, gj = groups[i], groups[j]
        if "thiol" in gi and "thiol" in gj:
            kind = "disulfide"
        elif ("amine" in gi and "acid" in gj) or ("acid" in gi and "amine" in gj):
            kind = "lactam"
        elif "olefin" in gi and "olefin" in gj:
            kind = "staple"
        else:
            raise UserError(f"bridges: no bridge known between {name(i)} and {name(j)} (disulfide = two Cys; "
                            f"lactam = Lys with Asp or Glu; staple = two olefin residues such as S5, R8)")
        out.append((i, j, kind))
    return sorted(out)


def bridge_text(p):
    """'disulfide C11-C18, lactam K3-D7' for titles and tables."""
    def name(i):
        c = p.units[i].code
        return (c if len(c) == 1 else f"({c})") + str(i + 1)
    return ", ".join(f"{kind} {name(i)}-{name(j)}" for i, j, kind in p.bridges)


def open_forms(pep):
    """Precursors with a ring not closed: [(name, formula to add)]."""
    out = []
    if pep.topology == "head-to-tail":
        out.append(("linear (+H2O)", parse_formula("H2O")))
    for kind in ("disulfide", "lactam", "staple"):
        n = sum(1 for b in pep.bridges if b[2] == kind)
        for k in range(1, n + 1):
            add = add_formula(Counter(), parse_formula(BRIDGE_LOSS[kind]), k)
            unit = {"lactam": "H2O", "staple": "C2H4"}.get(kind)
            what = f"+{2 * k}H" if kind == "disulfide" else (f"+{unit}" if k == 1 else f"+{k}{unit}")
            count = "" if n == 1 else (f"{k} of {n} " if k < n else "all ")
            out.append((f"{count}{BRIDGE_OPEN[kind]} ({what})", add))
    return out


CYCLO_RE = re.compile(r"^(?:cyclo|cyc|c)[(\[]-?(.*?)-?[)\]]$", re.I)


def split_dashes(text):
    """Split at dashes that are not inside brackets: 'Lys(5-FAM)-Gly' -> ['Lys(5-FAM)', 'Gly']."""
    out, depth, cur = [], 0, ""
    for ch in text:
        if ch in "([":
            depth += 1
        elif ch in ")]":
            depth -= 1
        if ch == "-" and depth <= 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    return out + [cur]


def three_to_one(seq, blocks):
    """'H-Ala-Gly-Lys(Biotin)-NH2' -> 'H-AGK[Biotin]-NH2'. None when the text is not three-letter.
    Three-letter means: dash-separated, and at least one part is a three-letter code as written in
    building_blocks.csv (Ala or ala; ALA stays one-letter Ala-Leu-Ala)."""
    look = residue_lookup(blocks)
    three = {b.code3 for b in blocks.values() if b.kind == "residue" and b.code3}
    three |= {c.lower() for c in three}
    toks = split_dashes(seq)
    base = lambda t: re.split(r"[(\[]", t, maxsplit=1)[0]
    if len(toks) < 2 or not any(base(t) in three for t in toks):
        return None
    n_caps = {b.code[:-1]: b.code for b in blocks.values() if b.kind == "n_cap"}
    c_caps = {b.code[1:]: b.code for b in blocks.values() if b.kind == "c_cap"}
    n_cap = c_cap = ""
    if toks[0] in n_caps and not (toks[0] in look and len(toks) > 1 and base(toks[1]) not in look
                                  and toks[1] not in ("D", "L")):
        n_cap, toks = n_caps[toks[0]], toks[1:]
    if toks and toks[-1] in c_caps:
        c_cap, toks = c_caps[toks[-1]], toks[:-1]
    body, i = "", 0
    while i < len(toks):
        t = toks[i]
        i += 1
        if t == "":
            raise UserError("two dashes in a row in the sequence")
        if t in ("D", "L") and i < len(toks):          # D-Phe is written with a dash of its own
            nxt, name = toks[i], base(toks[i])
            if t == "L":
                continue
            found = next((k for k in (f"D-{name}", f"D{name}", f"d{name}",
                                      "d" + look[name].code if name in look else None) if k and k in look), None)
            if found is None:
                raise UserError(f"D-{name} is not in building_blocks.csv (add a residue row, e.g. code "
                                f"d{look[name].code if name in look else name} with the D smiles)")
            t = found + nxt[len(name):]
            i += 1
        name = base(t)
        rest = t[len(name):]
        if name not in look:
            if i == 1 and not n_cap:
                raise UserError(f"unknown N-terminal group or residue '{name}' (N-terminal groups: "
                                f"{', '.join(n_caps.values())}; add a row to building_blocks.csv)")
            if i == len(toks) and not c_cap:
                raise UserError(f"unknown C-terminal group or residue '{name}' (C-terminal groups: "
                                f"{', '.join(c_caps.values())}; add a row to building_blocks.csv)")
            raise UserError(f"unknown residue '{name}' (add a row to building_blocks.csv)")
        code = look[name].code
        body += code if len(code) == 1 else f"({code})"
        if rest:
            if not re.fullmatch(r"\(.+\)|\[.+\]", rest):
                raise UserError(f"cannot read '{t}'; a side-chain tag is written Lys(Biotin)")
            body += f"[{rest[1:-1]}]"
    return n_cap + body + c_cap


def read_sequence(seq, blocks):
    """Whatever the user wrote -> (one-letter text, True when written as cyclo(...))."""
    seq = re.sub(r"\s+", "", seq or "").replace("\u2013", "-").replace("\u2212", "-")
    m = CYCLO_RE.match(seq)
    if m:
        seq = m.group(1)
    one = three_to_one(seq, blocks)
    return (seq if one is None else one), bool(m)


def guess_topology(seq, blocks=None):
    """For a row whose topology is not set yet."""
    s = re.sub(r"\s+", "", seq or "")
    if CYCLO_RE.match(s):
        return "head-to-tail"
    return "linear" if "-" in s else "head-to-tail"


def parse_sequence(seq, topology, blocks, bridges=""):
    seq, cyclo = read_sequence(seq, blocks)
    if not seq:
        raise UserError("empty sequence")
    if cyclo and topology == "linear":
        raise UserError("the sequence is written as cyclo(...), so the topology is head-to-tail")
    if topology not in TOPOLOGIES:
        raise UserError(f"topology must be one of {', '.join(TOPOLOGIES)} (got '{topology}')")
    look = residue_lookup(blocks)
    n_caps = sorted((b.code for b in blocks.values() if b.kind == "n_cap"), key=len, reverse=True)
    c_caps = sorted((b.code for b in blocks.values() if b.kind == "c_cap"), key=len, reverse=True)
    n_cap = c_cap = None
    for c in n_caps:
        if seq.startswith(c):
            n_cap, seq = c, seq[len(c):]
            break
    for c in c_caps:
        if seq.endswith(c):
            c_cap, seq = c, seq[: -len(c)]
            break
    if topology == "head-to-tail" and (n_cap or c_cap):
        raise UserError("a head-to-tail peptide has no ends; remove the caps from the sequence")
    if topology == "linear":
        n_cap = n_cap or "H-"
        c_cap = c_cap or "-OH"
        for c in (n_cap, c_cap):
            if c not in blocks:
                raise UserError(f"end group '{c}' is not in building_blocks.csv")
    units, i = [], 0
    while i < len(seq):
        ch = seq[i]
        if ch == "(":
            j = seq.find(")", i)
            if j < 0:
                raise UserError(f"unclosed '(' in {seq}")
            code, i = seq[i + 1:j], j + 1
        elif ch.isalpha():
            code, i = ch, i + 1
        elif ch == "-":
            raise UserError(f"unexpected '-' inside the sequence (ends must be one of "
                            f"{', '.join(n_caps)} / {', '.join(c_caps)})")
        else:
            raise UserError(f"unexpected character '{ch}' in sequence")
        if code not in look:
            hint = " (lower case letters are not defined; add a row to building_blocks.csv)" \
                if len(code) == 1 and code.islower() else ""
            raise UserError(f"unknown residue '{code}'{hint}")
        tag = None
        if i < len(seq) and seq[i] == "[":
            j = seq.find("]", i)
            if j < 0:
                raise UserError(f"unclosed '[' in {seq}")
            tag, i = seq[i + 1:j], j + 1
            if tag not in blocks or blocks[tag].kind != "tag":
                raise UserError(f"unknown tag '{tag}' (add it to building_blocks.csv as kind 'tag')")
            if 3 not in [a.GetAtomMapNum() for a in Chem.MolFromSmiles(look[code].smiles).GetAtoms()
                         if a.GetAtomicNum() == 0]:
                raise UserError(f"residue '{code}' has no side-chain attachment [*:3] for tag '{tag}'")
        units.append(Unit(look[code].code, tag))
    if len(units) < 2:
        raise UserError("a peptide needs at least 2 residues")
    fc = Counter()
    for u in units:
        fc.update(blocks[u.code].formula)
        if u.tag:
            fc.update(blocks[u.tag].formula)
    if topology == "linear":
        fc.update(blocks[n_cap].formula)
        fc.update(blocks[c_cap].formula)
    links = parse_bridges(bridges, units, blocks)
    for _, _, kind in links:
        fc = add_formula(fc, parse_formula(BRIDGE_LOSS[kind]), -1)
    return Peptide(units, topology, n_cap, c_cap, Counter({e: n for e, n in fc.items() if n}), links)


def one_letter(p):
    body = "".join((u.code if len(u.code) == 1 else f"({u.code})") + (f"[{u.tag}]" if u.tag else "")
                   for u in p.units)
    tail = f" ({bridge_text(p)})" if p.bridges else ""
    if p.topology == "head-to-tail":
        return f"cyclo({body}){tail}"
    return f"{p.n_cap}{body}{p.c_cap}{tail}"


def three_letter(p, blocks):
    parts = [(blocks[u.code].code3 or u.code) + (f"({u.tag})" if u.tag else "") for u in p.units]
    tail = f" ({bridge_text(p)})" if p.bridges else ""
    if p.topology == "head-to-tail":
        return "cyclo(-" + "-".join(parts) + "-)" + tail
    return p.n_cap + "-".join(parts) + p.c_cap + tail


# ----------------------------------------------------------------------------
# Species and ions
# ----------------------------------------------------------------------------
@dataclass
class Ion:
    label: str
    z: int
    mz: float
    iso: float = 0.0   # expected (first 13C peak / mono) intensity ratio


# fraction of the first heavy isotope per atom, relative to the light one
ISO1 = {"C": 0.010816, "H": 0.000115, "N": 0.003663, "O": 0.000381, "S": 0.007893, "Si": 0.0508}


def iso_ratio(formula):
    return sum(ISO1.get(e, 0.0) * n for e, n in formula.items() if n > 0)


# natural isotope abundances by nominal mass shift (0, +1, +2, ...)
ISO_ABUND = {"C": [0.9893, 0.0107], "H": [0.999885, 0.000115], "N": [0.99636, 0.00364],
             "O": [0.99757, 0.00038, 0.00205], "S": [0.9499, 0.0075, 0.0425, 0.0, 0.0001],
             "Cl": [0.7576, 0.0, 0.2424], "Br": [0.5069, 0.0, 0.4931], "K": [0.932581, 0.000117, 0.067302],
             "Si": [0.92223, 0.04685, 0.03092]}


def iso_pattern(formula, n=6):
    """Relative intensities of M, M+1, ... M+n-1 (largest = 1)."""
    dist = np.zeros(n)
    dist[0] = 1.0
    for e, cnt in formula.items():
        ab = ISO_ABUND.get(e)
        if not ab or cnt <= 0:
            continue
        base = np.zeros(n)
        base[:min(n, len(ab))] = ab[:n]
        result = np.zeros(n)
        result[0] = 1.0
        c = int(cnt)
        while c:
            if c & 1:
                result = np.convolve(result, base)[:n]
            base = np.convolve(base, base)[:n]
            c >>= 1
        dist = np.convolve(dist, result)[:n]
    return dist / dist.max()


@dataclass
class Species:
    name: str          # "main", "linear", "+Na", ...
    formula: Counter
    carrier: str = None  # adduct element (Na/K) or None for protons
    ions: list = field(default_factory=list)       # ions inside the scan range
    all_ions: list = field(default_factory=list)   # every charge state up to max_z

    @property
    def mass(self):
        return formula_mass(self.formula)


BUILTIN_DELTA_GROUPS = ("linear", "unbridged", "deletions", "insertions", "dimer")
AUTO_GROUPS = ("linear", "unbridged", "adducts", "side_reactions", "protecting_groups", "deletions", "insertions", "dimer")


def split_list(text):
    """'auto, +Na' or 'linear;+Na' -> ['auto', '+Na']"""
    return [t.strip() for t in re.split(r"[,;]", text or "") if t.strip()]


def add_formula(a, b, k=1):
    c = Counter(a)
    for e, n in b.items():
        c[e] += k * n
    return Counter({e: n for e, n in c.items() if n})


def delta_groups(blocks):
    return sorted({b.group for b in blocks.values() if b.group})


def expand_deltas(text, pep, blocks):
    """find_deltas text -> [(name, formula, carrier)] for everything to search besides the peptide itself."""
    groups = set(delta_groups(blocks))
    present = {u.code for u in pep.units}
    out = []

    def from_block(b):
        if b.kind == "adduct":
            return (b.code, pep.formula, hill(b.formula))
        return (b.code, add_formula(pep.formula, b.formula), None)

    def add_group(g, use_residues=True):
        if g in ("linear", "unbridged"):
            for name, add in open_forms(pep):
                if name.startswith("linear") == (g == "linear"):
                    out.append((name, add_formula(pep.formula, add), None))
        elif g in ("deletions", "insertions"):
            for code in dict.fromkeys(u.code for u in pep.units):
                if g == "deletions" and len(pep.units) <= 2:
                    continue
                name = blocks[code].code3 or code
                out.append((f"minus {name}" if g == "deletions" else f"extra {name}",
                            add_formula(pep.formula, blocks[code].formula, -1 if g == "deletions" else 1), None))
        elif g == "dimer":
            out.append(("dimer (2M)", add_formula(pep.formula, pep.formula), None))
        else:
            for b in blocks.values():
                if b.group == g and (not use_residues or not b.residues or any(c in present for c in b.residues)):
                    out.append(from_block(b))

    for tok in split_list(text) or ["auto"]:
        if tok == "none":
            continue
        if tok == "auto":
            for g in AUTO_GROUPS:
                if g in BUILTIN_DELTA_GROUPS or g in groups:
                    add_group(g)
        elif tok == "all":       # every delta in the file, whatever the sequence
            for g in BUILTIN_DELTA_GROUPS + tuple(sorted(groups)):
                add_group(g, use_residues=False)
            for b in blocks.values():
                if b.kind in ("shift", "adduct") and not b.group:
                    out.append(from_block(b))
        elif tok in BUILTIN_DELTA_GROUPS or tok in groups:
            add_group(tok)
        elif tok in blocks and blocks[tok].kind in ("shift", "adduct"):
            out.append(from_block(blocks[tok]))
        else:
            raise UserError(f"find_deltas: '{tok}' is not known. Use auto, all, none, a group "
                            f"({', '.join(BUILTIN_DELTA_GROUPS + tuple(sorted(groups)))}) or a code such as +Na "
                            f"(python3 make_ms_report.py deltas lists them)")
    # the same mass twice (e.g. extra Ala = +Acm): search once, keep both names
    merged = {}
    for name, f, carrier in out:
        key = (round(formula_mass(f), 4), carrier)
        if key == (round(pep.mass, 4), None):
            continue
        if key in merged:
            if name not in merged[key][0].split(" / "):
                merged[key][0] += f" / {name}"
        else:
            merged[key] = [name, f, carrier]
    return [tuple(v) for v in merged.values()]


def make_species(pep, deltas, blocks, scan_range, max_z):
    out = [Species("main", pep.formula)]
    for name, f, carrier in expand_deltas(deltas, pep, blocks):
        out.append(Species(name, f, carrier=carrier))
    lo, hi = scan_range
    for sp in out:
        M = sp.mass
        for z in range(1, max_z + 1):
            if sp.carrier:
                mz = (M + (z - 1) * PROTON + MONO[sp.carrier] - ELECTRON) / z
                hs = "" if z == 1 else ("H" if z == 2 else f"{z - 1}H")
                label = f"M+{hs}+{sp.carrier}" if hs else f"M+{sp.carrier}"
            else:
                mz = (M + z * PROTON) / z
                label = "M+H" if z == 1 else f"M+{z}H"
            sp.ions.append(Ion(label, z, mz, iso_ratio(sp.formula)))
        if sp.name.startswith("dimer"):     # even charges of 2M sit on top of the monomer ions
            sp.ions = [i for i in sp.ions if i.z % 2 == 1]
            for i in sp.ions:
                i.label = i.label.replace("M+", "2M+")
        sp.all_ions = list(sp.ions)
        sp.ions = [i for i in sp.ions if lo - 1 <= i.mz <= hi + 1] if lo else sp.ions
    return out


# ----------------------------------------------------------------------------
# Structure drawing
# ----------------------------------------------------------------------------
def build_molecule(pep, blocks):
    from functools import reduce
    n = len(pep.units)
    frags, tag_flags = [], []

    def relabel(smi, mapping, is_tag=False):
        m = Chem.RWMol(Chem.MolFromSmiles(smi))
        for a in m.GetAtoms():
            if a.GetAtomicNum() == 0:
                lab = mapping.get(a.GetAtomMapNum())
                if lab is None:
                    a.SetAtomicNum(1)
                    a.SetAtomMapNum(0)
                else:
                    a.SetAtomMapNum(lab)
            elif is_tag:
                a.SetIntProp("is_tag", 1)
        return m.GetMol()

    cyc = pep.topology == "head-to-tail"
    link, acid_end, olefin_end = {}, set(), set()   # residue index -> label shared by the two ends of its bridge
    for k, (i, j, kind) in enumerate(pep.bridges):
        link[i] = link[j] = 2000 + k
        if kind == "staple":
            olefin_end.update((i, j))
        if kind == "lactam":
            acid_end.add(i if "acid" in side_groups(blocks[pep.units[i].code]) and
                         "amine" in side_groups(blocks[pep.units[j].code]) else j)

    def acid_as_attachment(smi):
        """The side-chain COOH loses its OH: C(=O)O -> C(=O)[*:3]."""
        m = Chem.RWMol(Chem.MolFromSmiles(smi))
        for a in m.GetAtoms():                      # an unused [*:3] elsewhere becomes H
            if a.GetAtomicNum() == 0 and a.GetAtomMapNum() == 3:
                a.SetAtomMapNum(9)
        o = m.GetAtomWithIdx(m.GetSubstructMatch(Chem.MolFromSmarts(SIDE_ACID))[2])
        o.SetAtomicNum(0)
        o.SetAtomMapNum(3)
        return Chem.MolToSmiles(m)

    def olefin_as_attachment(smi):
        """Metathesis: the terminal =CH2 leaves (as ethylene with its partner): C=C -> C=[*:3]."""
        m = Chem.RWMol(Chem.MolFromSmiles(smi))
        for a in m.GetAtoms():
            if a.GetAtomicNum() == 0 and a.GetAtomMapNum() == 3:
                a.SetAtomMapNum(9)
        c = m.GetAtomWithIdx(m.GetSubstructMatch(Chem.MolFromSmarts(SIDE_OLEFIN))[1])
        c.SetAtomicNum(0)
        c.SetAtomMapNum(3)
        c.SetNoImplicit(True)
        c.SetNumExplicitHs(0)
        return Chem.MolToSmiles(m)

    for i, u in enumerate(pep.units):
        mp = {1: (100 + i - 1) if i > 0 else (100 + n - 1 if cyc else 1),
              2: (100 + i) if (i < n - 1 or cyc) else 2,
              3: (1000 + i) if u.tag else link.get(i)}
        smi = acid_as_attachment(blocks[u.code].smiles) if i in acid_end else \
            olefin_as_attachment(blocks[u.code].smiles) if i in olefin_end else blocks[u.code].smiles
        frags.append(relabel(smi, mp))
        if u.tag:
            frags.append(relabel(blocks[u.tag].smiles, {1: 1000 + i}, is_tag=True))
    if not cyc:
        frags.append(relabel(blocks[pep.n_cap].smiles, {1: 1}))
        frags.append(relabel(blocks[pep.c_cap].smiles, {2: 2}))
    mol = Chem.molzip(reduce(Chem.CombineMols, frags))
    mol = Chem.RemoveHs(mol)
    Chem.SanitizeMol(mol)
    if Counter({e: n for e, n in rdkit_formula(mol).items() if n}) != pep.formula:
        raise UserError(f"built structure {hill(rdkit_formula(mol))} does not match formula {hill(pep.formula)}")
    return mol


def crossed_bonds(mol):
    pos = mol.GetConformer().GetPositions()[:, :2]
    bonds = [(b.GetBeginAtomIdx(), b.GetEndAtomIdx()) for b in mol.GetBonds()]

    def ccw(a, b, c):
        return (c[1] - a[1]) * (b[0] - a[0]) > (b[1] - a[1]) * (c[0] - a[0])
    n = 0
    for i, (a, b) in enumerate(bonds):
        for c, d in bonds[i + 1:]:
            if len({a, b, c, d}) == 4 and ccw(pos[a], pos[c], pos[d]) != ccw(pos[b], pos[c], pos[d]) \
                    and ccw(pos[a], pos[b], pos[c]) != ccw(pos[a], pos[b], pos[d]):
                n += 1
    return n


def layout_2d(mol, cyclic):
    """Try RDKit's ring layout and CoordGen; keep the one with fewer crossed bonds."""
    from rdkit.Chem import rdDepictor
    options = []
    for coordgen in ((False, True) if cyclic else (True, False)):
        m = Chem.Mol(mol)
        rdDepictor.SetPreferCoordGen(coordgen)
        rdDepictor.Compute2DCoords(m, useRingTemplates=not coordgen)
        options.append((crossed_bonds(m), len(options), m))
    rdDepictor.SetPreferCoordGen(False)
    return min(options, key=lambda t: (t[0], t[1]))[2]


def draw_structure(mol, png_path, svg_path, mol_path, cyclic=False):
    from rdkit.Chem.Draw import rdMolDraw2D
    mol = layout_2d(mol, cyclic)
    Chem.MolToMolFile(mol, str(mol_path))
    tag_atoms = [a.GetIdx() for a in mol.GetAtoms() if a.HasProp("is_tag")]
    tag_bonds = [b.GetIdx() for b in mol.GetBonds()
                 if b.GetBeginAtomIdx() in tag_atoms and b.GetEndAtomIdx() in tag_atoms]
    colors = {i: (0.95, 0.55, 0.5) for i in tag_atoms}
    bcolors = {i: (0.95, 0.55, 0.5) for i in tag_bonds}
    pos = mol.GetConformer().GetPositions()
    span = pos.max(axis=0) - pos.min(axis=0)
    aspect = float(min(2.6, max(1.0, (span[0] + 1.5) / (span[1] + 1.5))))
    h = 1100
    for drawer, path, binary in ((rdMolDraw2D.MolDraw2DCairo(int(h * aspect), h), png_path, True),
                                 (rdMolDraw2D.MolDraw2DSVG(int(h * aspect / 2), h // 2), svg_path, False)):
        o = drawer.drawOptions()
        o.addStereoAnnotation = False
        o.bondLineWidth = 2 if binary else 1
        o.minFontSize = 22 if binary else 11
        if Path(STRUCTURE_FONT).exists():
            o.fontFile = STRUCTURE_FONT
        o.padding = 0.02
        o.clearBackground = True
        drawer.DrawMolecule(mol, highlightAtoms=tag_atoms, highlightAtomColors=colors,
                            highlightBonds=tag_bonds, highlightBondColors=bcolors)
        drawer.FinishDrawing()
        data = drawer.GetDrawingText()
        Path(path).write_bytes(data if binary else data.encode())
    return aspect


# ----------------------------------------------------------------------------
# MS data
# ----------------------------------------------------------------------------
def find_converter():
    env = os.environ.get("THERMORAWFILEPARSER")
    for p in ([Path(env)] if env else []) + CONVERTER_CANDIDATES:
        if p.exists():
            return p
    w = shutil.which("ThermoRawFileParser")
    return Path(w) if w else None


@dataclass
class Run:
    rt: np.ndarray
    tic: np.ndarray
    offsets: np.ndarray
    mz: np.ndarray
    inten: np.ndarray
    scan_lo: float
    scan_hi: float
    filter: str = ""

    def scan(self, k):
        a, b = self.offsets[k], self.offsets[k + 1]
        return self.mz[a:b], self.inten[a:b]


def load_run(path, cache_dir):
    path = Path(path)
    cache_dir.mkdir(parents=True, exist_ok=True)
    npz = cache_dir / (path.stem + ".npz")
    if npz.exists() and npz.stat().st_mtime >= path.stat().st_mtime:
        d = np.load(npz)
        if "filt" in d.files:
            return Run(d["rt"], d["tic"], d["offsets"], d["mz"], d["inten"], float(d["lo"]), float(d["hi"]),
                       str(d["filt"]))
    if path.suffix.lower() == ".raw":
        conv = find_converter()
        if conv is None:
            raise UserError("ThermoRawFileParser not found (expected ~/software/ThermoRawFileParser/"
                            "ThermoRawFileParser, or set THERMORAWFILEPARSER=/path/to/it)")
        mzml_path = cache_dir / (path.stem + ".mzML")
        print(f"    converting {path.name} ...")
        r = subprocess.run([str(conv), "-i", str(path.resolve()), "-b", str(mzml_path), "-f", "1", "-l", "4"],
                           capture_output=True, text=True)
        if r.returncode != 0 or not mzml_path.exists():
            raise UserError(f"conversion of {path.name} failed:\n{r.stdout}\n{r.stderr}")
    else:
        mzml_path = path
    print(f"    reading {mzml_path.name} ...")
    rts, tics, mzs, ints, lo, hi, filt = [], [], [], [], [], [], ""
    for sp in mzml.MzML(str(mzml_path)):
        if sp.get("ms level") != 1:
            continue
        sc = sp["scanList"]["scan"][0]
        rts.append(float(sc["scan start time"]))
        filt = filt or str(sc.get("filter string", ""))
        m, i = sp["m/z array"].astype(float), sp["intensity array"].astype(float)
        o = np.argsort(m)
        mzs.append(m[o])
        ints.append(i[o])
        tics.append(float(sp.get("total ion current", i.sum())))
        try:
            w = sc["scanWindowList"]["scanWindow"][0]
            lo.append(float(w["scan window lower limit"]))
            hi.append(float(w["scan window upper limit"]))
        except (KeyError, IndexError):
            pass
    if not rts:
        raise UserError(f"{path.name} has no MS1 spectra")
    offsets = np.concatenate([[0], np.cumsum([len(m) for m in mzs])])
    run = Run(np.array(rts), np.array(tics), offsets, np.concatenate(mzs), np.concatenate(ints),
              min(lo) if lo else 0.0, max(hi) if hi else 1e9, filt)
    np.savez_compressed(npz, rt=run.rt, tic=run.tic, offsets=run.offsets, mz=run.mz, inten=run.inten,
                        lo=run.scan_lo, hi=run.scan_hi, filt=filt)
    if path.suffix.lower() == ".raw":
        mzml_path.unlink(missing_ok=True)
    return run


def window_max(m, i, target, ppm):
    lo, hi = np.searchsorted(m, [target * (1 - ppm * 1e-6), target * (1 + ppm * 1e-6)])
    return float(i[lo:hi].max()) if hi > lo else 0.0


def iso_ok(mono, heavy, ion):
    """Observed first-isotope ratio within a factor of 2 of the formula's."""
    return mono > 0 and heavy > 0 and 0.5 * ion.iso <= heavy / mono <= 2.0 * ion.iso


def xic(run, ions, ppm):
    """Per-scan sum over ions of (mono + first 13C peak), counted only when the isotope ratio fits."""
    out = np.zeros(len(run.rt))
    for k in range(len(run.rt)):
        m, i = run.scan(k)
        s = 0.0
        for ion in ions:
            a = window_max(m, i, ion.mz, ppm)
            if a > 0:
                b = window_max(m, i, ion.mz + C13_SHIFT / ion.z, ppm)
                if iso_ok(a, b, ion):
                    s += a + b
        out[k] = s
    return out


def xic_many(run, species, ppm):
    """xic() for many species at once: one pass over the scans, all target m/z searched together."""
    pairs = [(si, ion) for si, sp in enumerate(species) for ion in sp.ions]
    out = np.zeros((len(species), len(run.rt)))
    if not pairs:
        return out
    n = len(pairs)
    mono = np.array([ion.mz for _, ion in pairs])
    heavy = mono + C13_SHIFT / np.array([ion.z for _, ion in pairs])
    targets = np.concatenate([mono, heavy])
    lo_t, hi_t = targets * (1 - ppm * 1e-6), targets * (1 + ppm * 1e-6)
    iso = np.array([ion.iso for _, ion in pairs])
    sidx = np.array([si for si, _ in pairs])
    for k in range(len(run.rt)):
        m, inten = run.scan(k)
        if len(m) == 0:
            continue
        lo, hi = np.searchsorted(m, lo_t), np.searchsorted(m, hi_t)
        idx = np.nonzero(hi > lo)[0]
        if len(idx) == 0:
            continue
        v = np.zeros(len(targets))
        v[idx] = inten[lo[idx]]
        for t in idx[(hi[idx] - lo[idx]) > 1]:
            v[t] = inten[lo[t]:hi[t]].max()
        a, b = v[:n], v[n:]
        ok = (a > 0) & (b > 0) & (b >= 0.5 * iso * a) & (b <= 2.0 * iso * a)
        if ok.any():
            np.add.at(out[:, k], sidx[ok], (a + b)[ok])
    return out


def smooth(y, n=5):
    if len(y) < n:
        return y
    return np.convolve(y, np.ones(n) / n, mode="same")


def supported_apex(y, min_scans, max_gap=2):
    """Highest smoothed-XIC apex whose signal lasts at least min_scans scans (gaps of max_gap allowed)."""
    ys, hit = smooth(y), y > 0
    tried = []
    for k in np.argsort(ys)[::-1]:
        if ys[k] <= 0 or len(tried) >= 25:
            break
        if any(abs(k - t) < 2 * min_scans for t in tried):
            continue
        tried.append(k)
        a = b = k
        gap = 0
        while a > 0 and gap <= max_gap:
            a -= 1
            gap = 0 if hit[a] else gap + 1
        gap = 0
        while b < len(y) - 1 and gap <= max_gap:
            b += 1
            gap = 0 if hit[b] else gap + 1
        if hit[a:b + 1].sum() >= min_scans:
            return int(k)
    return None


def peak_window(rt, y, k=None, frac=0.5, max_half_width=0.3):
    ys = smooth(y)
    if k is None:
        k = int(np.argmax(ys))
    if ys[k] <= 0:
        return None
    a = b = k
    while a > 0 and ys[a - 1] >= frac * ys[k] and rt[k] - rt[a - 1] <= max_half_width:
        a -= 1
    while b < len(ys) - 1 and ys[b + 1] >= frac * ys[k] and rt[b + 1] - rt[k] <= max_half_width:
        b += 1
    return a, b, k


TOL = {"merge": 5.0}     # peaks closer than this (ppm) are one peak; cmd_run widens it with --ppm


def averaged_spectrum(run, a, b, merge_ppm=None):
    merge_ppm = merge_ppm or TOL["merge"]
    s, e = run.offsets[a], run.offsets[b + 1]
    m, i = run.mz[s:e], run.inten[s:e]
    if len(m) == 0:
        return np.array([]), np.array([])
    o = np.argsort(m)
    m, i = m[o], i[o]
    gid = np.concatenate([[0], np.cumsum(np.diff(m) / m[:-1] * 1e6 > merge_ppm)])
    tot = np.bincount(gid, i)
    wmz = np.bincount(gid, i * m) / np.where(tot > 0, tot, 1)
    return wmz, tot / (b - a + 1)


def match_in_spectrum(smz, sint, ion, ppm):
    """Observed mono m/z if the mono and first isotope peaks are both present."""
    def best(t):
        lo, hi = np.searchsorted(smz, [t * (1 - ppm * 1e-6), t * (1 + ppm * 1e-6)])
        if hi <= lo:
            return None
        j = lo + int(np.argmax(sint[lo:hi]))
        return j
    j = best(ion.mz)
    if j is None:
        return None
    h = best(ion.mz + C13_SHIFT / ion.z)
    if h is None or not iso_ok(sint[j], sint[h], ion):
        return None
    if ion.iso >= 0.4 and best(ion.mz + 2 * C13_SHIFT / ion.z) is None:   # M+2 expected to be visible
        return None
    return j



# ----------------------------------------------------------------------------
# What are the big peaks?  (mass of each main LC peak, explained by deltas)
# ----------------------------------------------------------------------------
def cluster_mass(smz, sint, j, tol=0.004):
    """Charge (from isotope spacing) and monoisotopic neutral mass of the cluster containing peak j."""
    def peak(x):
        lo, hi = np.searchsorted(smz, [x - tol, x + tol])
        return lo + int(np.argmax(sint[lo:hi])) if hi > lo else None
    z = None
    for zz in range(8, 0, -1):
        up, down = peak(smz[j] + C13_SHIFT / zz), peak(smz[j] - C13_SHIFT / zz)
        if (up is not None and sint[up] > 0.05 * sint[j]) or (down is not None and sint[down] > 0.3 * sint[j]):
            z = zz
            break
    if z is None:
        return None
    cur = j
    for _ in range(8):   # walk down to the monoisotopic peak
        down = peak(smz[cur] - C13_SHIFT / z)
        if down is None or sint[down] < 0.1 * sint[j]:
            break
        cur = down
    return (float(smz[cur]) - PROTON) * z, z, float(smz[cur]), float(sint[cur])


def main_components(run, n=5, min_rel=0.05, min_mass=400.0):
    """The biggest LC peaks (by TIC) and the mass of the strongest isotope cluster in each."""
    t = smooth(run.tic, 7)
    tmax = t.max() if t.max() > 0 else 1
    picks = []
    for k in np.argsort(t)[::-1]:
        if t[k] < min_rel * tmax or len(picks) >= 4 * n:
            break
        if run.rt[k] < 1.0 or run.rt[k] > run.rt[-1] - 0.2:
            continue
        if any(abs(run.rt[k] - run.rt[q]) < 0.15 for q in picks):
            continue
        picks.append(int(k))
    comps = []
    for k in picks:
        smz, sint = averaged_spectrum(run, max(0, k - 3), min(len(run.rt) - 1, k + 3))
        for j in np.argsort(sint)[::-1][:6]:
            got = cluster_mass(smz, sint, int(j))
            if got is None or got[0] < min_mass:
                continue
            M, z, mono, inten = got
            if not any(abs(M - c["M"]) / M < 20e-6 and abs(run.rt[k] - c["rt"]) < 0.5 for c in comps):
                comps.append(dict(rt=float(run.rt[k]), rel=float(t[k] / tmax * 100), mz=mono, z=z, M=M, inten=inten))
            break
        if len(comps) >= n:
            break
    return comps


def explain_mass(M, pep, blocks, others, tol_ppm=10.0, max_out=3):
    """Possible explanations of a measured neutral mass M for this peptide (fewest steps first)."""
    tol = max(M * tol_ppm * 1e-6, 0.003)
    refs = [("expected", pep.mass)]
    for oname, add in open_forms(pep):
        refs.append((oname.split(" (")[0], pep.mass + formula_mass(add)))
    terms =[(b.code, b.mass) for b in blocks.values() if b.kind == "shift"]
    terms += [(f"{b.code} (as [M+{hill(b.formula)}]+)", b.mass - MONO["H"])
              for b in blocks.values() if b.kind == "adduct"]
    seen = set()
    for u in pep.units:
        if u.code in seen:
            continue
        seen.add(u.code)
        name = blocks[u.code].code3 or u.code
        terms.append((f"minus {name}", -blocks[u.code].mass))
        terms.append((f"extra {name}", blocks[u.code].mass))
    out = []
    for rname, rmass in refs:
        d = M - rmass
        if abs(d) <= tol:
            out.append((0, abs(d), f"{rname} ({d / M * 1e6:+.1f} ppm)"))
        if abs(M - 2 * rmass) <= tol:
            out.append((1, abs(M - 2 * rmass), f"{rname} dimer"))
        for i, (n1, m1) in enumerate(terms):
            if abs(d - m1) <= tol:
                out.append((1, abs(d - m1), f"{rname} {n1}"))
            for n2, m2 in terms[i:]:
                if abs(m1 + m2) < 0.01:     # e.g. minus Pro and extra Pro
                    continue
                if abs(d - m1 - m2) <= tol:
                    out.append((2, abs(d - m1 - m2), f"{rname} {n1} {n2}"))
    for oid, oname, omass in others:
        if abs(M - omass) <= tol:
            out.append((0, abs(M - omass), f"= peptide {oid} {oname}"))
    out.sort(key=lambda t: (t[0], t[1]))
    return [t[2] for t in out[:max_out]]


def blank_level(blank_run, mz, ppm=None):
    """Highest intensity of this m/z in any blank scan."""
    ppm = ppm or TOL["merge"]
    if blank_run is None:
        return 0.0
    best = 0.0
    for k in range(len(blank_run.rt)):
        m, i = blank_run.scan(k)
        best = max(best, window_max(m, i, mz, ppm))
    return best

# ----------------------------------------------------------------------------
# Samples
# ----------------------------------------------------------------------------
def read_samples(path):
    raw = Path(path).read_bytes()
    for enc in ("utf-8-sig", "mac_roman"):      # Excel on a Mac may save "CSV" as Mac Roman
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    rows = list(csv.DictReader(io.StringIO(text, newline="")))
    missing = [c for c in SAMPLE_COLUMNS if rows and c not in rows[0] and c not in OPTIONAL_COLUMNS]
    if missing:
        raise UserError(f"samples.csv is missing column(s): {', '.join(missing)}")
    out = []
    for r in rows:
        if not any((v or "").strip() for v in r.values() if isinstance(v, str)):
            continue
        d = {k: (v or "").strip() for k, v in r.items() if k and isinstance(v, (str, type(None)))}
        if "blank_file" not in d:                               # older sheets: first blank row
            d["blank_file"] = "auto"
        if "find_deltas" not in d and "also_look_for" in d:     # older sheets
            d["find_deltas"] = d["also_look_for"]
        for c in OPTIONAL_COLUMNS:
            d.setdefault(c, "")
        out.append(d)
    return out


OBSERVED_HEADER = re.compile(r"found|observ|obs\b|measured|exp(erimental)?\.?\s*m/?z", re.I)


def parse_rt(text):
    t = (text or "").strip()
    if not t:
        return None
    try:
        v = float(t)
    except ValueError:
        raise UserError(f"spectrum_rt: '{t}' is not a time in minutes")
    if v < 0:
        raise UserError("spectrum_rt: the time cannot be negative")
    return v


def parse_observed(text):
    vals = []
    for t in re.split(r"[;,\s]+", (text or "").strip()):
        if t:
            try:
                vals.append(float(t))
            except ValueError:
                raise UserError(f"observed_mz: '{t}' is not a number (separate values with ;)")
    return vals


def cell_observed(v):
    if isinstance(v, (int, float)) and v > 0:
        return f"{float(v):.4f}"
    if isinstance(v, str) and re.search(r"\d", v):
        nums = re.findall(r"\d+(?:\.\d+)?", v)
        return ";".join(nums)
    return ""


def assign_observed(values, species, tol_ppm=20.0):
    """Each entered m/z goes to the nearest calculated ion (any species, any charge)."""
    assigned, unassigned = {}, []
    for v in values:
        best = None
        for sp in species:
            for ion in sp.all_ions:
                e = (v - ion.mz) / ion.mz * 1e6
                if best is None or abs(e) < abs(best[0]):
                    best = (e, sp, ion)
        if best and abs(best[0]) <= tol_ppm:
            assigned[(best[1].name, best[2].label)] = v
        else:
            unassigned.append((v, best))
    return assigned, unassigned


def natural_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


BLANK_RE = re.compile(r"h2o|blank|water|solvent|wash", re.I)


def data_files(folder, exts):
    return sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in exts
                   and not p.name.startswith(("~$", "."))], key=lambda p: natural_key(p.name))


def sequences_from_xlsx(path, blocks):
    look = residue_lookup(blocks)
    tags = {b.code.lower(): b.code for b in blocks.values() if b.kind == "tag"}
    out = []
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)

    def cell_unit(v):
        if not isinstance(v, str):
            return None
        mt = re.fullmatch(r"\s*([A-Za-z0-9]+)\s*(?:\(\s*([A-Za-z0-9]+)\s*\))?\s*", v)
        if not mt or mt.group(1) not in look or len(mt.group(1)) < 3:
            return None
        res = look[mt.group(1)].code
        tag = tags.get((mt.group(2) or "").lower()) if mt.group(2) else None
        if mt.group(2) and not tag:
            return None
        return (res if len(res) == 1 else f"({res})") + (f"[{tag}]" if tag else "")

    for ws in wb.worksheets:
        obs_col = None
        for row in ws.iter_rows(max_row=5, values_only=True):
            for j, v in enumerate(row):
                if obs_col is None and isinstance(v, str) and OBSERVED_HEADER.search(v):
                    obs_col = j
        for row in ws.iter_rows(values_only=True):
            row = list(row)
            start = None
            for j, v in enumerate(row):
                if cell_unit(v) and j + 1 < len(row) and cell_unit(row[j + 1]):
                    start = j
                    break
            if start is None:
                continue
            seq = []
            for v in row[start:]:
                u = cell_unit(v)
                if not u:
                    break
                seq.append(u)
            sid = row[start - 1] if start > 0 else None
            sid = str(int(sid)) if isinstance(sid, float) and sid.is_integer() else (str(sid).strip() if sid is not None else "")
            if not sid:      # e.g. an id formula whose cached value is missing: number by order
                sid = str(len(out) + 1)
            obs = cell_observed(row[obs_col]) if obs_col is not None and obs_col < len(row) else ""
            out.append((sid, "".join(seq), f"{path.name} / {ws.title}", obs))
        if out:
            break
    wb.close()
    return out


def sequences_from_csv(path, blocks):
    out, obs_col = [], None
    with open(path, newline="", encoding="utf-8-sig") as f:
        for n, row in enumerate(csv.reader(f), start=1):
            cells = [c.strip() for c in row]
            if n == 1:
                obs_col = next((j for j, c in enumerate(cells) if OBSERVED_HEADER.search(c)), None)
            for j, c in enumerate(cells):
                if j == obs_col:
                    continue
                if len(c) < 3 or not re.fullmatch(r"[A-Za-z()\[\]0-9+\-]+", c):
                    continue
                try:
                    if not re.search(r"[A-Z]{2}", c) and three_to_one(CYCLO_RE.sub(r"\1", c), blocks) is None:
                        continue
                    parse_sequence(c, guess_topology(c), blocks)
                except UserError:
                    continue
                others = [x for k, x in enumerate(cells) if k not in (j, obs_col) and x]
                obs = cell_observed(cells[obs_col]) if obs_col is not None and obs_col < len(cells) else ""
                out.append((others[0] if others else str(n), c, path.name, obs))
                break
    return out


def cmd_setup(folder, blocks, force=False):
    sp = folder / "samples.csv"
    if sp.exists():
        if not force:
            raise UserError("samples.csv already exists; use 'setup' to rewrite it")
        old = folder / "samples_old.csv"
        shutil.copy(sp, old)
        print(f"Kept the previous samples.csv as {old.name}")
    raws = data_files(folder, {".raw", ".mzml"})
    xlsx = data_files(folder, {".xlsx", ".xlsm"})
    csvs = [p for p in data_files(folder, {".csv"})
            if p.name not in ("samples.csv", "samples_old.csv", "building_blocks.csv", "summary.csv")]
    found = []
    for p in xlsx:
        got = sequences_from_xlsx(p, blocks)
        print(f"  {p.name}: {len(got)} sequence(s) (3-letter cells)")
        found += got
    for p in csvs:
        got = sequences_from_csv(p, blocks)
        print(f"  {p.name}: {len(got)} sequence(s) (one-letter cells)")
        found += got
    if not raws:
        print("  WARNING: no .raw or .mzML files in this folder. Is the script in the data folder?")

    by_number = {}
    for p in raws:
        mt = re.search(r"(\d+)$", p.stem)
        if mt and not BLANK_RE.search(p.stem):
            by_number.setdefault(str(int(mt.group(1))), []).append(p)
    blank_raws = [p for p in raws if BLANK_RE.search(p.stem)]
    default_blank = blank_raws[0].name if blank_raws else ""
    used, rows = set(), []
    for sid, seq, src, obs in found:
        raw = ""
        num = re.search(r"(\d+)$", sid or "")
        cands = by_number.get(str(int(num.group(1))), []) if num else []
        cands = [c for c in cands if c not in used]
        if len(cands) == 1:
            raw = cands[0].name
            used.add(cands[0])
        rows.append(dict(id=sid, name="", sequence=seq, topology="", raw_file=raw, blank_file=default_blank, hplc_file="",
                         role="sample", find_deltas="auto", observed_mz=obs, notes=f"from {src}"))
    # sequences and raw files left over in equal numbers: pair them in order, flagged
    left_rows = [r for r in rows if not r["raw_file"]]
    left_raws = [p for p in raws if p not in used and not BLANK_RE.search(p.stem)]
    if left_rows and len(left_rows) == len(left_raws):
        for r, p in zip(left_rows, left_raws):
            r["raw_file"] = p.name
            r["notes"] += "; raw file paired by order - CHECK"
            used.add(p)
    for r in rows:
        if r["sequence"] and not r["raw_file"]:
            r["notes"] += "; no raw file found - row is skipped until raw_file is filled"
        sid = r["id"]
        if sid and not sid.isdigit():
            r["name"] = sid
        elif r["raw_file"]:
            r["name"] = Path(r["raw_file"]).stem.replace("_", " ")
        else:
            r["name"] = f"Peptide {sid}"
    for p in raws:
        if p in used:
            continue
        blank = bool(BLANK_RE.search(p.stem))
        rows.append(dict(id="blank" if blank else p.stem, name=p.stem.replace("_", " "), sequence="",
                         topology="", raw_file=p.name, blank_file="" if blank else default_blank,
                         hplc_file="", role="blank" if blank else "sample",
                         find_deltas="" if blank else "auto", observed_mz="",
                         notes="" if blank else "no sequence matched this file"))
    seen = Counter(r["id"] for r in rows)
    for r in rows:
        if seen[r["id"]] > 1 or not r["id"]:
            r["notes"] = (r["notes"] + "; id must be unique, please edit").strip("; ")
    with open(sp, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=SAMPLE_COLUMNS)
        w.writeheader()
        w.writerows(rows)
    n_seq = sum(1 for r in rows if r["sequence"])
    n_raw = sum(1 for r in rows if r["raw_file"] and r["sequence"])
    print(f"\nWrote {sp}\n  {n_seq} sequence(s), {n_raw} matched to a raw file, "
          f"{sum(1 for r in rows if r['role'] == 'blank')} blank(s).")
    print("\nNEXT: open samples.csv and")
    print("  - fill 'topology' for every sample: linear or head-to-tail")
    print("  - check sequences, ends (Ac-...-NH2) and tags (K[Biotin])")
    print("  - check raw_file matches; add hplc_file if you have one")
    print(f"  - blank_file = the blank run for each sample ({default_blank or 'none found; empty = no blank'})")
    print("  - find_deltas says what else to look for; 'auto' covers the usual suspects")
    print("  - optional: observed_mz = m/z values you read yourself, separated by ;")
    n_obs = sum(1 for r in rows if r.get("observed_mz"))
    if n_obs:
        print(f"    ({n_obs} row(s) already have observed m/z copied from the sequence file)")
    print("then run:  python3 make_ms_report.py")


def prepare_samples(folder, blocks, only=None):
    rows = read_samples(folder / "samples.csv")
    problems, samples, blanks = [], [], []
    ids = Counter(r["id"] for r in rows)
    for n, r in enumerate(rows, start=2):
        where = f"row {n} (id {r['id'] or '?'})"
        if r["role"] == "blank":
            if r["raw_file"]:
                if not (folder / r["raw_file"]).exists():
                    problems.append(f"{where}: raw_file '{r['raw_file']}' not found")
                else:
                    blanks.append(r)
            continue
        if r["role"] not in ("sample", ""):
            problems.append(f"{where}: role must be sample or blank")
            continue
        if only and r["id"] not in only:
            continue
        if not r["raw_file"]:
            print(f"  skipping {where}: no raw_file")
            continue
        if not r["id"] or ids[r["id"]] > 1:
            problems.append(f"{where}: id must be filled and unique")
        if not r["sequence"]:
            problems.append(f"{where}: no sequence")
            continue
        try:
            pep = parse_sequence(r["sequence"], r["topology"], blocks, r.get("bridges", ""))
        except UserError as e:
            problems.append(f"{where}: {e}")
            continue
        if not (folder / r["raw_file"]).exists():
            problems.append(f"{where}: raw_file '{r['raw_file']}' not found")
        bf = r.get("blank_file", "")
        if bf not in ("", "auto") and not (folder / bf).exists():
            problems.append(f"{where}: blank_file '{bf}' not found")
        if r["hplc_file"] and not (folder / r["hplc_file"]).exists():
            problems.append(f"{where}: hplc_file '{r['hplc_file']}' not found")
        try:
            expand_deltas(r["find_deltas"], pep, blocks)
        except UserError as e:
            problems.append(f"{where}: {e}")
        try:
            parse_observed(r["observed_mz"])
            parse_rt(r.get("spectrum_rt", ""))
            parse_purity_rt(r.get("purity_rt", ""))
        except UserError as e:
            problems.append(f"{where}: {e}")
        samples.append((r, pep))
    if problems:
        raise UserError("Please fix samples.csv:\n  " + "\n  ".join(problems))
    if only and not samples:
        raise UserError(f"no samples with id in {sorted(only)}")
    return samples, blanks


def cmd_check(folder, blocks, max_z):
    samples, blanks = prepare_samples(folder, blocks)
    print(f"samples.csv OK: {len(samples)} sample(s), {len(blanks)} blank(s)\n")
    for r, pep in samples:
        print(f"{r['id']:>6}  {one_letter(pep)}")
        print(f"        {hill(pep.formula)}   M (mono) = {pep.mass:.4f}")
        species = make_species(pep, r["find_deltas"], blocks, (0, 0), max_z)
        ions = "  ".join(f"{i.label} {i.mz:.4f}" for i in species[0].ions[:4])
        print(f"        expected       {ions}")
        names = [sp.name for sp in species[1:]]
        print(f"        find_deltas '{r['find_deltas'] or 'auto'}': {len(names)} other species"
              + (f" ({', '.join(names[:6])}{', ...' if len(names) > 6 else ''})" if names else ""))
        obs = parse_observed(r["observed_mz"])
        if obs:
            assigned, unassigned = assign_observed(obs, species)
            calc = {(sp.name, i.label): i.mz for sp in species for i in sp.all_ions}
            for (name, label), v in assigned.items():
                c = calc[(name, label)]
                print(f"        entered {v:.4f} = {name} {label} ({(v - c) / c * 1e6:+.1f} ppm)")
            for v, best in unassigned:
                print(f"        entered {v:.4f}: no ion within 20 ppm (nearest {best[1].name} {best[2].label} "
                      f"{best[2].mz:.4f}, {best[0]:+.0f} ppm)")


def cmd_mass(seq, topology, blocks, max_z, bridges=""):
    if not topology:
        topology = "linear" if bridges and not CYCLO_RE.match(seq.strip()) else guess_topology(seq)
        print(f"(topology not given; using {topology})")
    pep = parse_sequence(seq, topology, blocks, bridges)
    print(one_letter(pep))
    print(three_letter(pep, blocks))
    print(f"Formula {hill(pep.formula)}   M (mono) = {pep.mass:.4f}")
    for sp in make_species(pep, "", blocks, (0, 0), max_z):
        for i in sp.ions:
            print(f"  {i.label:8} {i.mz:.4f}")


def cmd_deltas(blocks, path):
    print(f"find_deltas options (from {path})\n")
    print("  auto        = " + ", ".join(AUTO_GROUPS) + "  (default; rows with 'residues' only when the peptide has one)")
    print("  all         = every delta below, whatever the sequence (more chance hits)")
    print("  none        = only the expected peptide")
    print("  linear      = uncyclised precursor of a head-to-tail peptide (+H2O)")
    print("  unbridged   = a bridge not formed: reduced disulfide (+2H), open lactam (+H2O), unstapled (+C2H4)")
    print("  deletions   = each residue of the sequence missing once")
    print("  insertions  = each residue of the sequence present twice")
    print("  dimer       = 2M (odd charge states, the even ones overlap the monomer)")
    for g in delta_groups(blocks):
        items = [b for b in blocks.values() if b.group == g]
        print(f"\n  {g}:")
        for b in items:
            res_ = f"  [{b.residues}]" if b.residues else ""
            print(f"    {b.code:9} {hill(b.formula):12} {b.mass:+10.4f}  {b.name}{res_}")
    loose = [b for b in blocks.values() if b.kind in ("shift", "adduct") and not b.group]
    if loose:
        print("\n  no group (use the code itself):")
        for b in loose:
            print(f"    {b.code:9} {hill(b.formula):12} {b.mass:+10.4f}  {b.name}")
    print("\nWrite them in samples.csv separated by comma space, e.g.:  auto   or   linear, adducts, +Pbf")


def cmd_blocks(blocks, path):
    print(f"Building blocks from {path}\n")
    for kind in ("residue", "n_cap", "c_cap", "tag", "adduct", "shift"):
        items = [b for b in blocks.values() if b.kind == kind]
        if items:
            print(f"{kind}:")
            for b in items:
                print(f"  {b.code:8} {b.code3:5} {hill(b.formula):14} {b.mass:11.5f}  {b.name}")


# ----------------------------------------------------------------------------
# Report
# ----------------------------------------------------------------------------
def downsample(x, y, n=2000):
    if len(x) <= n:
        return x, y
    k = int(np.ceil(len(x) / n))
    idx = [s + int(np.argmax(y[s:s + k])) for s in range(0, len(x), k)]
    return x[idx], y[idx]


def nl_text(v):
    """Xcalibur-style normalisation level, e.g. 7.55E8."""
    if v <= 0:
        return "0"
    e = int(np.floor(np.log10(v)))
    return f"{v / 10 ** e:.2f}E{e}"


def write_csv(path, header, rows):
    with open(path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def isotope_of_main(smz, sint, ion, height, main, pattern, ppm):
    """True when this ion sits on an isotope peak of the expected product and is no taller than that
    peak should be. A reduced disulfide (+2H, 2.0157) is 9 mDa from the product's own M+2 (2.0067):
    without this check the oxidised peptide would always 'contain' its reduced form."""
    for m_ion in main.all_ions:
        if m_ion.z != ion.z:
            continue
        for n in range(1, len(pattern)):
            target = m_ion.mz + n * C13_SHIFT / ion.z
            if abs(ion.mz - target) / target * 1e6 > 2 * ppm:
                continue
            lo, hi = np.searchsorted(smz, [m_ion.mz * (1 - ppm * 1e-6), m_ion.mz * (1 + ppm * 1e-6)])
            if hi > lo and pattern[0] > 0:
                return height <= 2.0 * float(sint[lo:hi].max()) * pattern[n] / pattern[0]
    return False


def analyse(r, pep, blocks, run, ppm, max_z, min_rel=1.0, min_scans=8):
    species = make_species(pep, r["find_deltas"], blocks, (run.scan_lo, run.scan_hi), max_z)
    res = {"species": [], "window": None, "found": False}
    all_xic = xic_many(run, species, ppm)
    main_pattern = iso_pattern(pep.formula)
    for sp, y in zip(species, all_xic):
        k = supported_apex(y, min_scans)
        w = peak_window(run.rt, y, k) if k is not None else None
        hits, status = {}, "no"
        if w:
            smz, sint = averaged_spectrum(run, w[0], w[1])
            floor = min_rel / 100 * (sint.max() if len(sint) else 0)
            for ion in sp.ions:
                j = match_in_spectrum(smz, sint, ion, ppm)
                if j is not None and sint[j] >= floor and not (
                        sp is not species[0] and not sp.carrier
                        and isotope_of_main(smz, sint, ion, float(sint[j]), species[0], main_pattern, ppm)):
                    hits[ion.label] = (float(smz[j]), float(sint[j]))
            if hits:
                charges = {i.z for i in sp.ions if i.label in hits}
                best_rel = max(h[1] for h in hits.values()) / sint.max()
                status = "yes" if len(charges) >= 2 or best_rel >= 0.05 else "weak"
        res["species"].append(dict(sp=sp, xic=y, window=w if hits else None, hits=hits, status=status))
    main = res["species"][0]
    res["status"] = main["status"]
    if main["window"]:
        res["window"], res["found"] = main["window"], True
        res["note"] = ("Weak evidence: one charge state, under 5% of the base peak."
                       if main["status"] == "weak" else "")
        res["marker"] = "main"
    else:
        t = smooth(run.tic)
        inside = (run.rt >= 1.0) & (run.rt <= run.rt[-1] - 0.2)   # skip the injection front and the end
        k = int(np.argmax(np.where(inside, t, -1))) if inside.any() else int(np.argmax(t))
        a, b = max(0, k - 5), min(len(run.rt) - 1, k + 5)
        res["window"], res["marker"] = (a, b, k), None
        res["note"] = "Expected ions not found; spectrum averaged at the TIC maximum."
    user_rt = parse_rt(r.get("spectrum_rt", ""))
    if user_rt is not None:   # set in samples.csv: line in E and spectrum in D at this time
        k = int(np.argmin(np.abs(run.rt - user_rt)))
        a, b = max(0, k - 5), min(len(run.rt) - 1, k + 5)
        res["window"], res["marker"] = (a, b, k), "user"
        found_txt = "" if res["found"] else " Expected product not found."
        res["note"] = f"Spectrum at {run.rt[k]:.2f} min (spectrum_rt in samples.csv).{found_txt}"
    a, b, _ = res["window"]
    res["smz"], res["sint"] = averaged_spectrum(run, a, b)
    # extra species seen in the main window, for labels in panel D
    res["labels"] = []
    floor = min_rel / 100 * (res["sint"].max() if len(res["sint"]) else 0)
    for s in res["species"]:
        if s["sp"].name != "main" and s["status"] != "yes":
            continue     # only confident hits get a label on the spectrum
        for ion in s["sp"].ions:
            j = match_in_spectrum(res["smz"], res["sint"], ion, ppm)
            if j is not None and res["sint"][j] >= floor:
                tag = ion.label if (s["sp"].name == "main" or s["sp"].carrier) else f"{s['sp'].name} {ion.label}"
                res["labels"].append((float(res["smz"][j]), float(res["sint"][j]), tag))
    return res


def read_hplc(path):
    xs, ys = [], []
    with open(path, encoding="utf-8-sig", errors="replace") as f:
        for line in f:
            parts = [p for p in re.split(r"[,\t; ]+", line.strip()) if p]
            nums = []
            for p in parts:
                try:
                    nums.append(float(p))
                except ValueError:
                    pass
            if len(nums) >= 2:
                xs.append(nums[0])
                ys.append(nums[1])
    if len(xs) < 5:
        raise UserError(f"could not read time/absorbance columns from {path.name}")
    xs, ys = np.array(xs), np.array(ys)
    o = np.argsort(xs, kind="stable")
    return xs[o], ys[o]


def parse_purity_rt(text):
    text = (text or "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        raise UserError(f"purity_rt: '{text}' is not a retention time in minutes")


# ----------------------------------------------------------------------------
# HPLC purity (area %)
# ----------------------------------------------------------------------------
def hplc_baseline(y, lam=None, p=0.002, rounds=12):
    """Asymmetric least-squares baseline (Eilers): a smooth curve that hugs the trace from below,
    so a drifting gradient baseline is not counted as peak area."""
    n = len(y)
    lam = lam if lam is not None else 1e-5 * n ** 4 / 600.0     # same stiffness per run, whatever the sampling
    d = np.diff(np.eye(n), 2, axis=0)
    pen = lam * d.T @ d
    w = np.ones(n)
    z = y
    for _ in range(rounds):
        z = np.linalg.solve(np.diag(w) + pen, w * y)
        w = np.where(y > z, p, 1 - p)
    return z


def trapezoid(y, x):
    return float(np.sum((y[1:] + y[:-1]) * np.diff(x)) / 2)


def hplc_peaks(t, y, skip_before=None):
    """Integrate an HPLC trace. Returns (peaks, baseline on t); each peak is a dict with
    rt, start, end, area, height, pct (of the summed peak area).

    Baseline: asymmetric least squares on the trace resampled to 600 points. A peak is a stretch
    above 3x the noise (and above 0.5 % of the tallest peak); stretches are cut at valleys between
    two tops (perpendicular drop). Peaks before skip_before min (default 1 min for runs over 5 min:
    the injection front) are left out, and so is anything under 0.1 % of the total area."""
    t, y = np.asarray(t, float), np.asarray(y, float)
    o = np.argsort(t)
    t, y = t[o], y[o]
    grid = np.linspace(t[0], t[-1], 600)
    base = np.interp(t, grid, hplc_baseline(np.interp(grid, t, y)))
    sig = y - base
    k = max(1, int(round(len(t) / 2000)))                 # light smoothing against detector noise
    sm = np.convolve(sig, np.ones(2 * k + 1) / (2 * k + 1), mode="same")
    resid = sig - sm if k > 1 else np.diff(sig, prepend=sig[0]) / np.sqrt(2)
    noise = 1.4826 * float(np.median(np.abs(resid - np.median(resid)))) or 1e-12
    thr = max(3 * noise, 0.005 * float(sm.max()))
    if skip_before is None:
        skip_before = 1.0 if t[-1] - t[0] > 5 else 0.0
    above = sm > thr
    peaks, i, n = [], 0, len(t)
    while i < n:
        if not above[i]:
            i += 1
            continue
        j = i
        while j + 1 < n and above[j + 1]:
            j += 1
        # tops inside this stretch: local maxima that stand out from the valley next to them
        seg = sm[i:j + 1]
        tops = [a for a in range(1, len(seg) - 1) if seg[a] >= seg[a - 1] and seg[a] > seg[a + 1]] or [int(np.argmax(seg))]
        kept = []
        for a in tops:
            if not kept:
                kept.append(a)
                continue
            valley = float(seg[kept[-1]:a + 1].min())
            if min(seg[kept[-1]], seg[a]) - valley > max(3 * noise, 0.02 * min(seg[kept[-1]], seg[a])):
                kept.append(a)
            elif seg[a] > seg[kept[-1]]:
                kept[-1] = a
        cuts = [0] + [kept[q] + int(np.argmin(seg[kept[q]:kept[q + 1] + 1])) for q in range(len(kept) - 1)] + [len(seg) - 1]
        for q, a in enumerate(kept):
            lo, hi = i + cuts[q], i + cuts[q + 1]
            if hi > lo:
                peaks.append(dict(rt=float(t[i + a]), start=float(t[lo]), end=float(t[hi]), lo=lo, hi=hi,
                                  area=trapezoid(np.clip(sig[lo:hi + 1], 0, None), t[lo:hi + 1]),
                                  height=float(sm[i + a])))
        i = j + 1
    peaks = [pk for pk in peaks if pk["rt"] >= t[0] + skip_before]
    total = sum(pk["area"] for pk in peaks)
    peaks = [pk for pk in peaks if total and pk["area"] >= 0.001 * total]
    total = sum(pk["area"] for pk in peaks)
    for pk in peaks:
        pk["pct"] = 100 * pk["area"] / total if total else 0.0
    return peaks, base


def hplc_purity(t, y, product_rt=None):
    """(peaks, baseline, index of the product peak or None). The product peak is the one at
    product_rt (within its own width, else the nearest within 0.3 min); without product_rt, the
    one with the largest area."""
    peaks, base = hplc_peaks(t, y)
    if not peaks:
        return peaks, base, None
    if product_rt is None:
        return peaks, base, int(np.argmax([pk["area"] for pk in peaks]))
    inside = [k for k, pk in enumerate(peaks) if pk["start"] <= product_rt <= pk["end"]]
    if inside:
        return peaks, base, min(inside, key=lambda k: abs(peaks[k]["rt"] - product_rt))
    k = int(np.argmin([abs(pk["rt"] - product_rt) for pk in peaks]))
    return peaks, base, (k if abs(peaks[k]["rt"] - product_rt) <= 0.3 else None)


# Slide layout in inches (x, y, w, h); slide is 13.33 x 12 in. Nothing here overlaps.
L = {
    "title": (0.3, 0.12, 12.7, 0.45),
    "seq": (0.45, 0.66, 12.4, 0.42),
    "A": (0.0, 1.2, 0.4, 0.3), "B": (7.55, 1.3, 5.6, 4.7), "C": (0.0, 3.95, 0.4, 0.3),
    "D": (0.0, 6.12, 0.4, 0.3), "E": (0.0, 9.22, 0.4, 0.26),
    "A_pic": (0.45, 1.3, 6.6, 2.28), "A_formula": (0.45, 3.62, 6.6, 0.2),
    "C_table": (0.45, 4.0, 6.6, 1.78),       # x, y, width, max height; caption goes right below
    "D_head": (0.75, 6.14, 6.6, 0.28), "D_note": (7.45, 6.14, 5.7, 0.28),
    "D_chart": (0.3, 6.45, 12.85, 2.75),
    "E_head_l": (0.75, 9.32, 4.0, 0.15), "E_head_r": (8.15, 9.32, 5.0, 0.15),
    "E_chart": (0.3, 9.5, 12.85, 2.38),
}
L["B_letter"] = (7.15, 1.2, 0.4, 0.3)
SPEC_MARGINS = (0.62, 0.08, 0.2, 0.48)
ZOOM_MARGINS = (0.62, 0.08, 0.2, 0.48)   # left, top, right, bottom of the plot inside its chart (in)
TIC_MARGINS = (0.62, 0.08, 0.2, 0.48)
SPEC_YMAX, TIC_YMAX = 132, 118          # headroom above 100 % so peak labels stay inside the plot box
Y_FMT = '[<=100]0;""'                   # tick labels only up to 100, as in Xcalibur
LABEL_PT = 8
ARIAL_EM = {**{d: 0.556 for d in "0123456789"}, ".": 0.278, " ": 0.278, "+": 0.584, "-": 0.333,
            "(": 0.333, ")": 0.333, "M": 0.833, "H": 0.722, "N": 0.722, "K": 0.667, "a": 0.556, "O": 0.778}


def label_size(text, pt=LABEL_PT):
    """Width and height (in) of a data label in Arial, with a small safety margin."""
    rows = text.split("\n")
    w = max(sum(ARIAL_EM.get(c, 0.65) for c in row) for row in rows) * pt / 72
    return w + 0.08, len(rows) * pt * 1.2 / 72 + 0.04


class PlotGeom:
    """Maps data coordinates to slide inches for a chart whose plot area is fixed by margins."""
    def __init__(self, x, y, w, h, margins, xlim, ylim):
        ml, mt, mr, mb = margins
        self.top, self.left, self.right = y, x, x + w
        self.L, self.T, self.W, self.H = x + ml, y + mt, w - ml - mr, h - mt - mb
        self.xlim, self.ylim = xlim, ylim

    def X(self, v):
        return self.L + (np.asarray(v, float) - self.xlim[0]) / (self.xlim[1] - self.xlim[0]) * self.W

    def Y(self, v):
        return self.T + (1 - (np.asarray(v, float) - self.ylim[0]) / (self.ylim[1] - self.ylim[0])) * self.H


def stick_top(sticks, left, right, own_x):
    """Highest point (smallest y, in) of the spectrum sticks inside [left, right], ignoring the labelled one."""
    px, py = sticks
    m = (px >= left) & (px <= right) & (np.abs(px - own_x) > 1e-6)
    return float(py[m].min()) if m.any() else np.inf


def line_top(trace, left, right):
    px, py = trace
    m = (px >= left) & (px <= right)
    return float(py[m].min()) if m.any() else np.inf


def place_labels(cands, geo, obstacle_top, max_n=30, gap=0.04, step=0.04):
    """Greedy, collision-free label placement.
    cands: (point index, x, y, text, colour, must) in priority order. A 'must' label is lifted
    until it is clear; others are dropped if they would touch a placed label or the trace.
    Returns (placed [(idx, text, colour, lift)], dropped [cand])."""
    rects, placed, dropped = [], [], []
    for cand in cands:
        idx, xv, yv, text, color, must = cand
        if not must and len(placed) >= max_n:
            break
        w, h = label_size(text)
        cx, base = float(geo.X(xv)), float(geo.Y(yv)) - gap
        left, right = cx - w / 2, cx + w / 2
        if left < geo.left + 0.05 or right > geo.right - 0.05:
            dropped.append(cand)
            continue
        floor_ = obstacle_top(left, right, float(xv))
        lift, ok = 0.0, False
        while True:
            bottom = base - lift
            top = bottom - h
            if top < geo.T + 0.02:          # stay inside the grey plot box
                break
            if bottom <= floor_ and all(not (left < r2 and right > l2 and top < b2 and bottom > t2)
                                        for l2, t2, r2, b2 in rects):
                ok = True
                break
            if not must:
                break
            lift += step
        if ok:
            rects.append((left, top, right, bottom))
            placed.append((idx, text, color, lift))
        else:
            dropped.append(cand)
    return placed, dropped


class Deck:
    def __init__(self):
        from pptx.util import Emu
        self.prs = Presentation()
        self.prs.slide_width = Emu(12192000)    # 13.33 in
        self.prs.slide_height = Emu(10972800)   # 12 in, as the lab template

    # --- helpers
    def text(self, slide, x, y, w, h, text, size=12, bold=False, color=None, align=None):
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.02)
        tf.margin_top = tf.margin_bottom = Inches(0.0)
        for n, line in enumerate(text.split("\n")):
            p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
            run = p.add_run()
            run.text = line
            run.font.size, run.font.bold, run.font.name = Pt(size), bold, FONT
            if color:
                run.font.color.rgb = RGBColor.from_string(color)
            if align:
                p.alignment = align
        return tb

    def xy_chart(self, slide, x, y, w, h, series, xtitle, ytitle, xlim=None, ylim=(0, 100),
                 x_major=None, x_fmt="0", plot_fill="E6E6E6", margins=None, y_fmt="0"):
        """Xcalibur-style XY chart: grey plot area, black trace, 0-100 relative abundance.
        series: (name, xs, ys, colour, line width pt)."""
        from pptx.chart.data import XyChartData
        from pptx.enum.chart import XL_CHART_TYPE, XL_TICK_MARK
        from pptx.oxml.ns import qn
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from lxml import etree
        cd = XyChartData()
        for name, xs, ys, *_ in series:
            sr = cd.add_series(name)
            for a, b in zip(xs, ys):
                sr.add_data_point(float(a), float(b))
        gf = slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
                                    Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = gf.chart
        ch.font.size, ch.font.name = Pt(9), FONT
        ch.has_legend = False
        ch.has_title = False
        for sr, (_, _, _, color, width) in zip(ch.plots[0].series, series):
            sr.smooth = False
            sr.format.line.width = Pt(width)
            sr.format.line.color.rgb = RGBColor.from_string(color)
            sr.marker.style = None
        for ax, t, lim, major, fmt in ((ch.category_axis, xtitle, xlim, x_major, x_fmt),
                                       (ch.value_axis, ytitle, ylim, 10, y_fmt)):
            ax.has_major_gridlines = False
            ax.has_title = True
            ax.axis_title.text_frame.text = t
            for p in ax.axis_title.text_frame.paragraphs:
                for rn in p.runs:
                    rn.font.size, rn.font.bold, rn.font.name = Pt(9), False, FONT
            ax.tick_labels.font.size, ax.tick_labels.font.name = Pt(8), FONT
            ax.tick_labels.number_format, ax.tick_labels.number_format_is_linked = fmt, False
            ax.major_tick_mark = XL_TICK_MARK.OUTSIDE
            ax.minor_tick_mark = XL_TICK_MARK.OUTSIDE
            ax.format.line.color.rgb = RGBColor(0x40, 0x40, 0x40)
            if lim:
                ax.minimum_scale, ax.maximum_scale = lim
            if major:
                ax.major_unit, ax.minor_unit = major, major / 5
        if plot_fill:   # grey plot area, as in Xcalibur/FreeStyle
            pa = ch._chartSpace.chart.plotArea
            sp = pa.find(qn("c:spPr"))
            if sp is None:
                sp = etree.SubElement(pa, qn("c:spPr"))
            fill = etree.SubElement(sp, qn("a:solidFill"))
            etree.SubElement(fill, qn("a:srgbClr")).set("val", plot_fill)
        if margins:     # fixed plot-area position, so label positions can be computed exactly
            ml_, mt_, mr_, mb_ = margins
            pa = ch._chartSpace.chart.plotArea
            lay = pa.find(qn("c:layout"))
            if lay is None:
                lay = etree.Element(qn("c:layout"))
                pa.insert(0, lay)
            for c in list(lay):
                lay.remove(c)
            man = etree.SubElement(lay, qn("c:manualLayout"))
            for tag, val in (("layoutTarget", "inner"), ("xMode", "edge"), ("yMode", "edge"),
                             ("x", ml_ / w), ("y", mt_ / h), ("w", (w - ml_ - mr_) / w), ("h", (h - mt_ - mb_) / h)):
                etree.SubElement(man, qn("c:" + tag)).set("val", val if isinstance(val, str) else f"{val:.5f}")
        return ch

    @staticmethod
    def point_labels(series, labels, size=None, chart_h=None):
        """labels: (point index, text, colour hex, lift in inches) placed above the point."""
        from pptx.enum.chart import XL_LABEL_POSITION
        from pptx.oxml.ns import qn
        from pptx.util import Pt
        from pptx.dml.color import RGBColor
        from lxml import etree
        size = size or LABEL_PT
        for idx, text, color, lift in labels:
            dl = series.points[idx].data_label
            dl.has_text_frame = True
            dl.text_frame.text = text
            dl.position = XL_LABEL_POSITION.ABOVE
            if lift and chart_h:
                lay = etree.Element(qn("c:layout"))
                man = etree.SubElement(lay, qn("c:manualLayout"))
                etree.SubElement(man, qn("c:x")).set("val", "0")
                etree.SubElement(man, qn("c:y")).set("val", f"{-lift / chart_h:.5f}")
                dl._dLbl.insert(1, lay)   # right after c:idx
            for p in dl.text_frame.paragraphs:
                for rn in p.runs:
                    rn.font.size, rn.font.name = Pt(size), FONT
                    rn.font.color.rgb = RGBColor.from_string(color)

    def isotope_zoom(self, s, pep, res, run, data_dir, sid, ppm):
        """Panel B when there is no HPLC file: the intended product's isotope cluster, Xcalibur style,
        each peak labelled with its measured m/z (the calculated pattern goes to the CSV only)."""
        main = res["species"][0]
        sp = main["sp"]
        product = "cyclic" if pep.topology == "head-to-tail" else "expected"
        if main["window"]:     # the product's own elution peak, whatever panel D shows
            a, b, _ = main["window"]
            smz, sint = averaged_spectrum(run, a, b)
        else:                  # not found: the spectrum behind the line in E
            a, b, _ = res["window"]
            smz, sint = res["smz"], res["sint"]
        ions = sp.ions or sp.all_ions
        if main["hits"]:
            ion = max((i for i in ions if i.label in main["hits"]), key=lambda i: main["hits"][i.label][1])
        else:
            scored = [(window_max(smz, sint, i.mz, ppm), i) for i in ions]
            best = max(scored, key=lambda t: t[0]) if scored else (0.0, None)
            ion = best[1] if best[0] > 0 else next((i for i in ions if i.z == 2), ions[0])
        z = ion.z
        f_ion = add_formula(sp.formula, Counter({"H": z}))
        pat = iso_pattern(f_ion, 6)
        calc_x = np.array([ion.mz + j * C13_SHIFT / z for j in range(len(pat))])
        lo, hi = ion.mz - 1.2 / z, ion.mz + 5.6 / z
        sel = (smz >= lo) & (smz <= hi)
        wmz, wint = smz[sel], sint[sel]
        wmax = float(wint.max()) if len(wint) else 0.0
        mono_obs = window_max(wmz, wint, ion.mz, ppm) if len(wmz) else 0.0
        near = None   # closest measured peak when the monoisotopic one is outside the tolerance
        if wmax > 0 and mono_obs > 0:
            calc_y = pat / pat[0] * mono_obs / wmax * 100
            scale_txt = "scaled to the measured mono peak"
        else:
            calc_y = pat * 100
            scale_txt = f"no peak within {ppm:g} ppm"
            if len(wmz):
                j = int(np.argmin(np.abs(wmz - ion.mz)))
                e = (wmz[j] - ion.mz) / ion.mz * 1e6
                if abs(e) <= 50 and wint[j] >= 0.05 * wmax:
                    near = (float(wmz[j]), e)
        calc_y = np.minimum(calc_y, 125)
        if mono_obs <= 0 and near is None and len(sint) and sint.max() > 0:
            wmax = float(sint.max())    # nothing here: scale to the whole spectrum so noise stays flat
        rel = wint / wmax * 100 if wmax > 0 else wint
        write_csv(data_dir / f"{sid}_isotope_zoom.csv", ["mz", "measured_intensity", "calculated_rel"],
                  [(f"{m_:.5f}", f"{i_:.4g}", "") for m_, i_ in zip(wmz, wint)] +
                  [(f"{m_:.5f}", "", f"{c_:.2f}") for m_, c_ in zip(calc_x, calc_y)])
        sx, sy = [], []
        keep = rel >= 0.5
        for m_, i_ in zip(wmz[keep], rel[keep]):
            sx += [m_, m_, m_]
            sy += [0, i_, 0]
        if not sx:
            sx, sy = [lo, hi], [0, 0]
        span = hi - lo
        step = next(st for st in (0.1, 0.2, 0.25, 0.5, 1.0, 2.0) if span / st <= 7)
        xlim = (float(np.floor(lo / step) * step), float(np.ceil(hi / step) * step))
        x_fmt = "0" if step >= 1 else ("0.00" if step == 0.25 else "0.0")
        bx, by, bw, bh = L["B"]
        head_h = 0.36
        chart_rect = (bx, by + head_h, bw, bh - head_h)
        geo = PlotGeom(*chart_rect, ZOOM_MARGINS, xlim, (0, SPEC_YMAX))
        ch = self.xy_chart(s, *chart_rect, [("Measured", sx, sy, "000000", 1.0)],
                           "m/z", "Relative Abundance", xlim=xlim, ylim=(0, SPEC_YMAX), x_major=step, x_fmt=x_fmt,
                           margins=ZOOM_MARGINS, y_fmt=Y_FMT)
        # labels: the monoisotopic peak in red, then the other measured peaks (isotopes) with their m/z
        px = np.array(sx[1::3]) if len(sx) > 2 else np.array([])
        py = np.array(sy[1::3]) if len(sy) > 2 else np.array([])
        cands = []
        if len(px):
            jm = int(np.argmin(np.abs(px - ion.mz)))
            if abs(px[jm] - ion.mz) / ion.mz * 1e6 <= ppm:
                cands.append((3 * jm + 1, px[jm], py[jm], f"{ion.label}\n{px[jm]:.4f}", "C00000", True))
            elif near and abs(px[jm] - near[0]) < 1e-6:
                cands.append((3 * jm + 1, px[jm], py[jm], f"{px[jm]:.4f}\n{near[1]:+.1f} ppm", "C55A11", True))
            cands += [(3 * int(j) + 1, px[j], py[j], f"{px[j]:.4f}", "008000", bool(py[j] >= 5))
                      for j in np.argsort(py)[::-1] if py[j] >= 2 and not (cands and 3 * int(j) + 1 == cands[0][0])]
        sticks = (geo.X(px), geo.Y(py)) if len(px) else (np.array([]), np.array([]))
        placed, _ = place_labels(cands, geo, lambda l_, r_, own: stick_top(sticks, l_, r_, own), max_n=12)
        self.point_labels(ch.plots[0].series[0], placed, chart_h=chart_rect[3])
        found_txt = "" if main["window"] else " (not found)"
        self.text(s, bx + 0.3, by, bw - 0.3, head_h - 0.02,
                  f"Zoom: {product} {ion.label}{found_txt}, calc. m/z {ion.mz:.4f}   RT: {run.rt[a]:.2f}-{run.rt[b]:.2f}"
                  f"   NL: {nl_text(wmax) if wmax > 0 else '0'}"
                  f"{'' if (mono_obs > 0 or near) else ' (base peak of the whole spectrum)'}\n"
                  "Isotope peaks labelled with their measured m/z" +
                  ("" if mono_obs > 0 else
                   f". Orange: {near[1]:+.1f} ppm from the calculated m/z, not the product" if near else
                   f"; nothing within {ppm:g} ppm of the calculated m/z"),
                  size=8, color="000000")

    # --- one slide
    def add_overview(self, summary, title):
        """Table of every sample on the first slide(s): found or not, retention time, ppm, purity."""
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN
        if not summary:
            return
        with_purity = any(x.get("purity_pct") for x in summary)
        header = ["ID", "Name", "Sequence", "M (mono)", "Found", "RT (min)", "Error (ppm)"] + \
                 (["Purity (%)"] if with_purity else []) + ["Also found"]
        widths = [0.7, 2.3, 3.6, 1.0, 0.7, 0.8, 0.9] + ([0.85] if with_purity else []) + [0.0]
        widths[-1] = 12.7 - sum(widths)
        word = {"yes": "yes", "weak": "weak", "no": "no"}
        colour = {"yes": "1A7F37", "weak": "8A5A00", "no": "B42318"}
        per, row_h = 34, 0.3
        chunks = [summary[i:i + per] for i in range(0, len(summary), per)]
        for n, chunk in enumerate(chunks):
            s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
            found = sum(1 for x in summary if x["found"] == "yes")
            weak = sum(1 for x in summary if x["found"] == "weak")
            self.text(s, *L["title"], title + (f" ({n + 1}/{len(chunks)})" if len(chunks) > 1 else ""),
                      size=24, bold=True, color="002A5C")
            self.text(s, 0.45, 0.66, 12.4, 0.3,
                      f"{len(summary)} samples: expected mass found in {found}"
                      + (f", weak evidence in {weak}" if weak else "") + ". One slide per sample follows.",
                      size=11, color="1F2933")
            shape = s.shapes.add_table(len(chunk) + 1, len(header), Inches(0.3), Inches(1.15), Inches(12.7),
                                       Inches(row_h * (len(chunk) + 1)))
            shape._element.graphic.graphicData.tbl.tblPr.find(
                "{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId").text = \
                "{5940675A-B579-460E-94D1-54222C63F5DA}"
            tbl = shape.table
            for ci, wd in enumerate(widths):
                tbl.columns[ci].width = Inches(wd)
            body = []
            for x in chunk:
                also = x.get("deltas_found", "")
                also = re.sub(r" \([^)]*min[^)]*\)", "", also)
                row = [x["id"], x["name"], x["sequence"], x["mono_mass"], word.get(x["found"], x["found"]),
                       x["rt_min"], x["mean_ppm"]] + ([x.get("purity_pct", "")] if with_purity else []) + [also]
                body.append(row)
            for ri, vals in enumerate([header] + body):
                for ci, v in enumerate(vals):
                    c = tbl.cell(ri, ci)
                    c.margin_top = c.margin_bottom = Inches(0.02)
                    limit = int(widths[ci] * 17)
                    v = str(v)
                    c.text = v if len(v) <= limit else v[:limit - 1].rstrip() + "\u2026"
                    for p in c.text_frame.paragraphs:
                        p.alignment = PP_ALIGN.LEFT if ci in (1, 2, len(header) - 1) else PP_ALIGN.CENTER
                        for rn in p.runs:
                            rn.font.size, rn.font.name, rn.font.bold = Pt(9), FONT, ri == 0
                            if ri and ci == 4:
                                rn.font.bold = True
                                rn.font.color.rgb = RGBColor.from_string(colour.get(vals[4], "1F2933"))
                tbl.rows[ri].height = Inches(row_h)
        lst = self.prs.slides._sldIdLst            # the overview goes first
        for n in range(len(chunks)):
            el = lst[len(lst) - len(chunks) + n]
            lst.remove(el)
            lst.insert(n, el)

    def add_sample(self, r, pep, blocks, res, run, blank_run, data_dir, folder, ppm, blank_name=""):
        from pptx.util import Inches, Pt
        from pptx.dml.color import RGBColor
        from pptx.enum.chart import XL_LABEL_POSITION
        from pptx.enum.dml import MSO_LINE_DASH_STYLE
        from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
        sid = re.sub(r"[^\w.-]+", "_", r["id"])
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.text(s, *L["title"], r["name"] or f"Sample {r['id']}", size=24, bold=True, color="002A5C")
        n_res = len(pep.units)
        seq1 = f"{one_letter(pep)}   ({n_res} residues)"
        seq3 = three_letter(pep, blocks)
        seq_pt = 11 if max(len(seq1), len(seq3)) <= 120 else (9 if max(len(seq1), len(seq3)) <= 150 else 7)
        self.text(s, *L["seq"], f"{seq1}\n{seq3}", size=seq_pt, color="1F2933")
        for lab in "ABCDE":
            self.text(s, *(L["B_letter"] if lab == "B" else L[lab]), f"{lab}.", size=16, bold=True)

        # A. structure
        png, svg, mol = data_dir / f"{sid}_structure.png", data_dir / f"{sid}.svg", data_dir / f"{sid}.mol"
        try:
            m = build_molecule(pep, blocks)
            aspect = draw_structure(m, png, svg, mol, cyclic=pep.ringed)
            bx, by, box_w, box_h = L["A_pic"]
            w = min(box_w, box_h * aspect)
            s.shapes.add_picture(str(png), Inches(bx + (box_w - w) / 2), Inches(by), Inches(w), Inches(w / aspect))
        except Exception as e:  # drawing must never stop the report
            self.text(s, L["A_pic"][0], 2.3, 6.7, 0.4, f"(structure could not be drawn: {e})", size=10, color="B00020")
        self.text(s, *L["A_formula"],
                  f"{hill(pep.formula)}    Exact mass (mono.): {pep.mass:.4f}", size=10, color="5E6873",
                  align=PP_ALIGN.CENTER)

        # B. HPLC
        if r["hplc_file"]:
            hx, hy = read_hplc(folder / r["hplc_file"])
            write_csv(data_dir / f"{sid}_hplc.csv", ["time_min", "absorbance"], zip(hx, hy))
            series, caption = [], "Purity: not calculated"
            try:      # purity = area of the product peak / area of all integrated peaks
                peaks, base, k = hplc_purity(hx, hy, parse_purity_rt(r.get("purity_rt", "")))
                write_csv(data_dir / f"{sid}_hplc_peaks.csv", ["rt_min", "start_min", "end_min", "area", "area_pct", "product"],
                          [(f"{pk['rt']:.3f}", f"{pk['start']:.3f}", f"{pk['end']:.3f}", f"{pk['area']:.5g}",
                            f"{pk['pct']:.2f}", "yes" if n == k else "") for n, pk in enumerate(peaks)])
                bx = np.linspace(float(hx.min()), float(hx.max()), 200)
                series.append(("baseline", bx, np.interp(bx, hx, base), "9AA0A6", 0.5))
                res["hplc_base"] = base
                if k is not None:
                    pk = peaks[k]
                    series.append(("product peak", *downsample(hx[pk["lo"]:pk["hi"] + 1], hy[pk["lo"]:pk["hi"] + 1], 400),
                                   "C00000", 1.25))
                    res["purity"] = dict(pct=pk["pct"], rt=pk["rt"], n=len(peaks))
                    res["hplc_seg"] = (hx[pk["lo"]:pk["hi"] + 1], hy[pk["lo"]:pk["hi"] + 1])
                    caption = (f"Purity {pk['pct']:.1f} % (area of the peak at {pk['rt']:.2f} min; "
                               f"{len(peaks)} peak{'s' if len(peaks) != 1 else ''} integrated)")
                else:
                    caption = "Purity: no peak at purity_rt"
            except Exception as e:      # an odd trace must never stop the report
                caption = f"Purity: not calculated ({e})"
            dx, dy = downsample(hx, hy)
            self.xy_chart(s, *L["B"], [("HPLC", dx, dy, "000000", 0.75)] + series, "Time (min)",
                          "Absorbance", xlim=(float(hx.min()), float(hx.max())), ylim=None, x_fmt="0.0")
            self.text(s, L["B"][0] + L["B"][2] - 4.6, L["B"][1] - 0.22, 4.6, 0.18, caption,
                      size=9, color="5E6873", align=PP_ALIGN.RIGHT)
            res["hplc"] = (hx, hy, res.get("hplc_base"), res.get("hplc_seg"), caption)
        else:
            self.isotope_zoom(s, pep, res, run, data_dir, sid, ppm)

        # C. table: entered observed m/z (samples.csv) first, otherwise the value found in the MS data
        main = res["species"][0]
        entered, unassigned = assign_observed(parse_observed(r.get("observed_mz", "")),
                                              [sx["sp"] for sx in res["species"]])
        rows = [("M", "", f"{pep.mass:.4f}", "", "", "")]

        def ion_row(sx, ion, label_species):
            v = entered.get((sx["sp"].name, ion.label))
            if v is None and ion.label in sx["hits"]:
                v = sx["hits"][ion.label][0]
            obs = f"{v:.4f}" if v is not None else ("n.d." if ion in sx["sp"].ions else "out of range")
            err = f"{(v - ion.mz) / ion.mz * 1e6:.1f}" if v is not None else ""
            rt_ = f"{run.rt[sx['window'][2]]:.2f}" if sx["window"] and v is not None else ""
            return (ion.label, label_species, f"{ion.mz:.4f}", obs, err, rt_)

        top_z = max([i.z for i in main["sp"].ions if i.label in main["hits"]] + [3])
        for ion in main["sp"].all_ions:
            if (ion in main["sp"].ions and ion.z <= top_z) or ("main", ion.label) in entered:
                rows.append(ion_row(main, ion, ""))
        n_main = len(rows)
        weak_rows = []
        ranked = sorted(res["species"][1:], key=lambda sx: -max([h[1] for h in sx["hits"].values()] or [0]))
        for sx in ranked:
            for ion in sx["sp"].all_ions:
                if ion.label in sx["hits"] or (sx["sp"].name, ion.label) in entered:
                    row = ion_row(sx, ion, sx["sp"].name)
                    (rows if sx["status"] == "yes" or (sx["sp"].name, ion.label) in entered else weak_rows).append(row)
        for v, best in unassigned:
            rows.append(("?", "no ion within 20 ppm", f"{best[2].mz:.4f}" if best else "", f"{v:.4f}",
                         f"{best[0]:.0f}" if best else "", ""))
        has_species = any(rw[1] for rw in rows)
        header = ["Ion"] + (["Species"] if has_species else []) + ["Calculated (mono.)", "Observed", "Error (ppm)",
                                                                  "RT (min)"]
        body = [[rw[0]] + ([rw[1] or "cyclic" if pep.topology == "head-to-tail" else rw[1] or "expected"]
                           if has_species else []) + list(rw[2:]) for rw in rows]
        weak_body = [[rw[0]] + ([rw[1]] if has_species else []) + list(rw[2:]) for rw in weak_rows]
        write_csv(data_dir / f"{sid}_mz_table.csv", header + ["evidence"],
                  [b_ + ["yes" if i >= n_main else ""] for i, b_ in enumerate(body)] + [w_ + ["weak"] for w_ in weak_body])
        tx, ty, tw, avail = L["C_table"]
        max_rows = int(avail / 0.15) - 1
        hidden = 0
        if len(body) > max_rows:
            hidden = len(body) - max_rows
            body = body[:max_rows]
        row_h = min(0.24, avail / (len(body) + 1))
        size = 10 if row_h >= 0.24 else (9 if row_h >= 0.2 else (8 if row_h >= 0.17 else 7))
        widths = [0.7, 1.9, 1.1, 1.1, 0.9, 0.9] if has_species else [1.0, 1.5, 1.5, 1.3, 1.3]
        if has_species:   # long species names must fit on one line, or the table grows into the caption
            col = 1
            fits = lambda pt: all(label_size(rw[col], pt)[0] + 0.12 <= widths[col] for rw in body)
            while size > 7 and not fits(size):
                size -= 1
            for rw in body:
                while label_size(rw[col], size)[0] + 0.12 > widths[col] and len(rw[col]) > 4:
                    rw[col] = rw[col][:-2].rstrip() + "\u2026"
        tbl_shape = s.shapes.add_table(len(body) + 1, len(header), Inches(tx), Inches(ty),
                                       Inches(tw), Inches(min(avail, row_h * (len(body) + 1))))
        tbl_shape._element.graphic.graphicData.tbl.tblPr.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId").text = \
            "{5940675A-B579-460E-94D1-54222C63F5DA}"   # "No Style, Table Grid"
        tbl = tbl_shape.table
        for ci, wd in enumerate(widths):
            tbl.columns[ci].width = Inches(wd)
        for ri, vals in enumerate([header] + body):
            for ci, v in enumerate(vals):
                c = tbl.cell(ri, ci)
                c.margin_top = c.margin_bottom = Inches(0.02)
                c.text = v
                for p in c.text_frame.paragraphs:
                    p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
                    for rn in p.runs:
                        rn.font.size, rn.font.name, rn.font.bold = Pt(size), FONT, ri == 0
            tbl.rows[ri].height = Inches(row_h)
        caption = "n.d. = not detected in the MS data."
        if hidden or weak_rows:
            caption += f" More in data/{sid}_mz_table.csv" + (f" ({hidden} more hits, {len(weak_rows)} weak)." if hidden
                                                              else f" ({len(weak_rows)} weak).")
        if entered:
            caption = "Observed: values entered in samples.csv where given, otherwise found in the MS data. " + caption
        self.text(s, tx, ty + row_h * (len(body) + 1) + 0.03, tw, 0.16, caption, size=8, color="5E6873")

        stem = Path(r["raw_file"]).stem
        a, b, k = res["window"]

        # D. spectrum, full scan range, Xcalibur style
        smz, sint = res["smz"], res["sint"]
        bsub = ""
        if blank_run is not None and len(smz):
            # blank spectrum over the same retention-time window, subtracted peak by peak (within ppm)
            bk = np.where((blank_run.rt >= run.rt[a] - 1e-6) & (blank_run.rt <= run.rt[b] + 1e-6))[0]
            if len(bk) == 0:
                bk = np.array([int(np.argmin(np.abs(blank_run.rt - run.rt[k])))])
            bmz, bint = averaged_spectrum(blank_run, int(bk[0]), int(bk[-1]))
            bsub_int = np.zeros_like(sint)
            if len(bmz):
                pos = np.clip(np.searchsorted(bmz, smz), 1, len(bmz) - 1)
                near = np.where(np.abs(bmz[pos - 1] - smz) < np.abs(bmz[pos] - smz), pos - 1, pos)
                ok = np.abs(bmz[near] - smz) / smz * 1e6 <= ppm
                bsub_int[ok] = bint[near[ok]]
            write_csv(data_dir / f"{sid}_spectrum.csv", ["mz", "intensity", "blank_intensity", "blank_subtracted"],
                      zip(smz, sint, bsub_int, np.clip(sint - bsub_int, 0, None)))
            sint = np.clip(sint - bsub_int, 0, None)
            bsub = f"   minus blank ({Path(blank_name).stem})"
        else:
            write_csv(data_dir / f"{sid}_spectrum.csv", ["mz", "intensity"], zip(smz, sint))
        x0, x1 = run.scan_lo, run.scan_hi
        if not (x1 - x0 < 1e6):
            x0, x1 = float(smz.min()), float(smz.max())
        sel = (smz >= x0) & (smz <= x1)
        pmz, pint = smz[sel], sint[sel]
        nl_s = pint.max() if len(pint) and pint.max() > 0 else 1
        keep = pint >= 0.003 * nl_s
        if keep.sum() > 1500:
            keep = pint >= np.sort(pint)[-1500]
        expected = {round(x, 4): tag for x, _, tag in res["labels"]}
        keep |= np.isin(np.round(pmz, 4), list(expected))
        pmz, pint = pmz[keep], pint[keep] / nl_s * 100
        sx, sy = [], []
        for m_, i_ in zip(pmz, pint):
            sx += [m_, m_, m_]
            sy += [0, i_, 0]
        if not sx:
            sx, sy = [x0, x1], [0, 0]
        step = 100 if x1 - x0 > 600 else 50
        xlim = (float(np.floor(x0 / step) * step), float(np.ceil(x1 / step) * step))
        geo = PlotGeom(*L["D_chart"], SPEC_MARGINS, xlim, (0, SPEC_YMAX))
        ch = self.xy_chart(s, *L["D_chart"], [("Spectrum", sx, sy, "000000", 0.75)], "m/z",
                           "Relative Abundance", xlim=xlim, ylim=(0, SPEC_YMAX), x_major=step, margins=SPEC_MARGINS,
                           y_fmt=Y_FMT)
        # expected ions first (red, lifted if needed to stay clear), then the tallest other peaks (green)
        cands = [(3 * j + 1, m_, pint[j], f"{expected[round(m_, 4)]}\n{m_:.4f}", "C00000", True)
                 for j, m_ in enumerate(pmz) if round(m_, 4) in expected]
        cands += [(3 * int(j) + 1, pmz[j], pint[j], f"{pmz[j]:.4f}", "008000", False)
                  for j in np.argsort(pint)[::-1][:200] if pint[j] >= 1.0 and round(pmz[j], 4) not in expected]
        sticks = (geo.X(pmz), geo.Y(pint))
        placed, dropped = place_labels(cands, geo, lambda l_, r_, own: stick_top(sticks, l_, r_, own), max_n=30)
        self.point_labels(ch.plots[0].series[0], placed, chart_h=L["D_chart"][3])
        missing_expected = [c[3].split("\n")[0] for c in dropped if c[4] == "C00000"]
        filt = run.filter or f"Full ms [{x0:.4f}-{x1:.4f}]"
        self.text(s, *L["D_head"],
                  f"{stem}  RT: {run.rt[a]:.2f}-{run.rt[b]:.2f}  AV: {b - a + 1}  NL: {nl_text(nl_s)}{bsub}\nT: {filt}",
                  size=8, color="000000")
        note = res["note"]
        if missing_expected:
            note = (note + " " if note else "") + "Not labelled (no room): " + ", ".join(missing_expected)
        if note:
            self.text(s, *L["D_note"], note, size=8, color="C00000", align=PP_ALIGN.RIGHT)

        # E. TIC, Xcalibur style, with a red marker at the spectrum's retention time
        rt, tic = run.rt, run.tic
        nl = tic.max() if tic.max() > 0 else 1
        xs, ys = downsample(rt, tic / nl * 100, 1500)
        write_csv(data_dir / f"{sid}_tic.csv", ["rt_min", "tic", "xic_expected_ions"], zip(rt, tic, main["xic"]))
        if blank_run is not None:
            write_csv(data_dir / f"{sid}_blank_tic.csv", ["rt_min", "tic"], zip(blank_run.rt, blank_run.tic))
        xmax = float(np.ceil(rt.max() * 2) / 2)
        geo = PlotGeom(*L["E_chart"], TIC_MARGINS, (0, xmax), (0, TIC_YMAX))
        mk = res.get("marker")
        if mk == "main":
            mk_text = ("cyclic" if pep.topology == "head-to-tail" else "expected") + \
                (" (weak)" if res["status"] == "weak" else "")
            mk_color = "E00000"
        elif mk == "user":
            mk_text, mk_color = "set by user", "E00000"
        else:
            mk_text, mk_color = None, "808080"
        product = "cyclic" if pep.topology == "head-to-tail" else "expected"
        mk_text = f"{mk_text} {rt[k]:.2f}" if mk_text else f"TIC max {rt[k]:.2f} ({product} not found)"
        e_series = [("TIC", xs, ys, "000000", 0.75),
                    ("Spectrum RT", [rt[k], rt[k]], [0, 100], mk_color, 1.0)]
        blank_clipped = False
        if blank_run is not None:   # same scale as the sample (divided by the sample's NL)
            bx, by = downsample(blank_run.rt, blank_run.tic / nl * 100, 1500)
            blank_clipped = bool(np.max(by) > 100) if len(by) else False
            by = np.minimum(by, 100)
            e_series.insert(0, ("Blank", bx, by, "A6A6A6", 0.75))   # first = drawn behind the TIC
        ch = self.xy_chart(s, *L["E_chart"], e_series,
                           "Time (min)", "Relative Abundance", xlim=(0, xmax), ylim=(0, TIC_YMAX), y_fmt=Y_FMT,
                           x_major=0.5, x_fmt="0.0", margins=TIC_MARGINS)
        peaks = [i for i in range(len(ys)) if ys[i] >= 1.0 and ys[i] == ys[max(0, i - 3):i + 4].max()]
        cands = [(i, xs[i], ys[i], f"{xs[i]:.2f}", "008000", False) for i in sorted(peaks, key=lambda i: -ys[i])]
        trace = (geo.X(xs), geo.Y(ys))
        cands.insert(0, (-1, rt[k], 100.0, mk_text, mk_color, True))   # label on top of the line, placed first
        placed, _ = place_labels(cands, geo, lambda l_, r_, own: line_top(trace, l_, r_), max_n=41)
        off = 1 if blank_run is not None else 0
        self.point_labels(ch.plots[0].series[off], [p_ for p_ in placed if p_[0] >= 0], chart_h=L["E_chart"][3])
        self.point_labels(ch.plots[0].series[off + 1], [(1, *p_[1:]) for p_ in placed if p_[0] < 0],
                          chart_h=L["E_chart"][3])
        self.text(s, *L["E_head_l"], f"RT: {rt.min():.2f}-{rt.max():.2f}", size=8, color="000000")
        blank_txt = (f"   grey: blank {Path(blank_name).stem}" + (" (cut at 100)" if blank_clipped else "")) \
            if blank_run is not None else ""
        self.text(s, *L["E_head_r"], f"NL: {nl_text(nl)}   TIC MS   {stem}{blank_txt}", size=8, color="000000",
                  align=PP_ALIGN.RIGHT)

        notes = [f"Sample {r['id']}: {r['raw_file']}", f"Tolerance {ppm} ppm; each ion also needs its first 13C peak at the expected ratio."]
        for sx_ in res["species"]:
            w = sx_["window"]
            if w:
                found = ", ".join(sx_["hits"])
                notes.append(f"{sx_['sp'].name}: found at {run.rt[w[2]]:.2f} min ({found})")
            else:
                notes.append(f"{sx_['sp'].name}: not found")
        if blank_run is not None:
            notes.append(f"Blank: {blank_name}. Grey trace in E on the sample's scale; panel D shows the spectrum "
                         f"minus the blank's spectrum over the same retention-time window "
                         f"(both are in data/{sid}_spectrum.csv). Detection (found / weak) uses the "
                         f"spectrum without subtraction.")
        else:
            notes.append("No blank for this sample.")
        notes.append(f"TIC and the XIC of the expected ions are in data/{sid}_tic.csv"
                     + (" and the blank TIC in _blank_tic.csv." if blank_run is not None else "."))
        if res.get("decoy_hits") is not None:
            notes.append(f"Delta hits: {sum(1 for sx in res['species'][1:] if sx['status'] == 'yes')}; the same search "
                         f"with another peptide's sequence gave {res['decoy_hits']} (chance level).")
        if res.get("components"):
            notes.append("Biggest LC peaks (neutral monoisotopic mass; possible explanations, not proof):")
            for c in res["components"]:
                notes.append(f"  {c['rt']:.2f} min, {c['rel']:.0f}% TIC: M={c['M']:.4f} (z={c['z']}, "
                             f"{c['M'] - pep.mass:+.4f} vs expected) - {', '.join(c['explain']) or 'unexplained'}"
                             f"{' - also in blank' if c['blank'] else ''}")
        if r["notes"]:
            notes.append(r["notes"])
        s.notes_slide.notes_text_frame.text = "\n".join(notes)


# ----------------------------------------------------------------------------
# PDF copy of the report (for lab notebooks)
# ----------------------------------------------------------------------------
def write_pdf(path, title, summary, pages):
    """One overview page, then one page per sample: structure, m/z table, HPLC or the reason there
    is none, spectrum and TIC. Drawn with matplotlib; returns False when it is not installed."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_pdf import PdfPages
    except ImportError:
        return False
    plt.rcParams.update({"font.family": ["Arial", "Helvetica", "DejaVu Sans"], "font.size": 8,
                         "axes.spines.top": False, "axes.spines.right": False, "pdf.fonttype": 42})
    ink, grey, red = "#16191d", "#5f6b76", "#c00000"
    colour = {"yes": "#1a7f37", "weak": "#8a5a00", "no": "#b42318"}
    with PdfPages(path) as pdf:
        per = 38
        with_purity = any(x.get("purity_pct") for x in summary)
        cols = ["ID", "Name", "Sequence", "M (mono)", "Found", "RT (min)", "ppm"] + (["Purity %"] if with_purity else [])
        xs = [0.06, 0.12, 0.34, 0.66, 0.76, 0.83, 0.90] + ([0.95] if with_purity else [])
        for n in range(0, max(1, len(summary)), per):
            fig = plt.figure(figsize=(8.27, 11.69))
            fig.text(0.06, 0.95, title, fontsize=15, weight="bold", color="#002a5c")
            found = sum(1 for x in summary if x["found"] == "yes")
            fig.text(0.06, 0.93, f"{len(summary)} samples: expected mass found in {found}. "
                                 f"Made {datetime.now():%Y-%m-%d %H:%M} with MS Report {VERSION}.", color=grey)
            y = 0.90
            for cx, c in zip(xs, cols):
                fig.text(cx, y, c, weight="bold", color=grey)
            for x in summary[n:n + per]:
                y -= 0.0215
                vals = [x["id"], x["name"][:30], x["sequence"][:44] + ("\u2026" if len(x["sequence"]) > 44 else ""),
                        x["mono_mass"], x["found"], x["rt_min"], x["mean_ppm"]] + ([x.get("purity_pct", "")] if with_purity else [])
                for k, (cx, v) in enumerate(zip(xs, vals)):
                    fig.text(cx, y, str(v), color=colour.get(x["found"], ink) if k == 4 else ink,
                             weight="bold" if k == 4 else "normal", family="monospace" if k == 2 else None,
                             fontsize=7 if k == 2 else 8)
            pdf.savefig(fig)
            plt.close(fig)
        for pg in pages:
            fig = plt.figure(figsize=(8.27, 11.69))
            fig.text(0.06, 0.955, pg["title"], fontsize=14, weight="bold", color="#002a5c")
            fig.text(0.06, 0.938, pg["seq"][:150], fontsize=7 if len(pg["seq"]) > 90 else 8, family="monospace")
            fig.text(0.06, 0.924, pg["formula"], color=grey)
            fig.text(0.94, 0.955, {"yes": "found", "weak": "weak evidence", "no": "not found"}[pg["found"]],
                     ha="right", fontsize=11, weight="bold", color=colour[pg["found"]])
            ax = fig.add_axes([0.06, 0.725, 0.42, 0.19])          # structure
            ax.axis("off")
            if pg.get("png") and Path(pg["png"]).exists():
                ax.imshow(plt.imread(pg["png"]))
            ax = fig.add_axes([0.54, 0.725, 0.40, 0.185])         # HPLC, when there is one
            if pg.get("hplc"):
                hx, hy, base, seg, caption = pg["hplc"]
                ax.plot(hx, hy, color=ink, lw=0.6)
                if base is not None:
                    ax.plot(hx, base, color="#9aa0a6", lw=0.5)
                if seg is not None:
                    ax.plot(seg[0], seg[1], color=red, lw=1.0)
                ax.set_xlabel("Time (min)")
                ax.set_ylabel("Absorbance")
                ax.set_title(caption, fontsize=7.5, color=grey, loc="right")
            else:
                ax.axis("off")
            y = 0.675
            fig.text(0.06, y, "Ion", weight="bold", color=grey)
            fig.text(0.20, y, "Calculated m/z", weight="bold", color=grey)
            fig.text(0.36, y, "Observed m/z", weight="bold", color=grey)
            fig.text(0.52, y, "Error (ppm)", weight="bold", color=grey)
            fig.text(0.66, y, "RT (min)", weight="bold", color=grey)
            for row in pg["table"][:8]:
                y -= 0.0165
                for cx, v in zip((0.06, 0.20, 0.36, 0.52, 0.66), row):
                    fig.text(cx, y, v)
            ax = fig.add_axes([0.08, 0.30, 0.86, 0.22])           # spectrum
            smz, sint = pg["smz"], pg["sint"]
            if len(smz):
                top = float(sint.max())
                keep = np.argsort(sint)[-4000:]
                ax.vlines(smz[keep], 0, sint[keep] / top * 100, color=ink, lw=0.5)
                for mz, inten, tag in pg["labels"]:
                    ax.vlines([mz], 0, [inten / top * 100], color=red, lw=0.8)
                    ax.text(mz, inten / top * 100 + 2, f"{tag}\n{mz:.4f}", color=red, ha="center", va="bottom", fontsize=6.5)
                ax.set_xlim(pg["mz_range"])
            ax.set_ylim(0, 125)
            ax.set_xlabel("m/z")
            ax.set_ylabel("Relative abundance")
            ax.set_title(pg["spec_title"], fontsize=7.5, color=grey, loc="left")
            ax = fig.add_axes([0.08, 0.05, 0.86, 0.19])           # TIC
            rt, tic = pg["rt"], pg["tic"]
            top = float(tic.max()) or 1.0
            if pg.get("blank") is not None:
                ax.plot(pg["blank"][0], pg["blank"][1] / top * 100, color="#9aa0a6", lw=0.6)
            ax.plot(rt, tic / top * 100, color=ink, lw=0.6)
            if pg.get("line_rt") is not None:
                ax.axvline(pg["line_rt"], color=red if pg["found"] != "no" else "#9aa0a6", lw=0.8)
                ax.text(pg["line_rt"], 106, f" {pg['line_rt']:.2f}", ha="left", fontsize=7,
                        color=red if pg["found"] != "no" else grey)
            ax.set_xlim(float(rt[0]), float(rt[-1]))
            ax.set_ylim(0, 115)
            ax.set_xlabel("Time (min)")
            ax.set_ylabel("Relative abundance")
            ax.set_title("TIC" + (", blank in grey" if pg.get("blank") is not None else ""), fontsize=7.5, color=grey, loc="left")
            if pg.get("note"):
                fig.text(0.06, 0.012, pg["note"][:170], fontsize=6.5, color=grey)
            pdf.savefig(fig)
            plt.close(fig)
    return True


def cmd_run(folder, blocks, blocks_path, ppm, max_z, only, min_rel, min_scans):
    TOL["merge"] = max(5.0, ppm)
    samples, blanks = prepare_samples(folder, blocks, only)
    out = folder / OUT_DIR_NAME
    data_dir, cache = out / "data", out / "cache"
    data_dir.mkdir(parents=True, exist_ok=True)
    default_blank = blanks[0]["raw_file"] if blanks else ""
    blank_runs = {}

    def blank_for(r):
        name = r.get("blank_file", "auto")
        name = default_blank if name == "auto" else name
        if not name:
            return "", None
        if name not in blank_runs:
            blank_runs[name] = load_run(folder / name, cache)
        return name, blank_runs[name]
    deck = Deck()
    summary, pages = [], []
    others = []   # every peptide of this run, to spot carry-over or swapped vials
    for r_, p_ in samples:
        others.append((r_["id"], "", p_.mass))
        for oname, add in open_forms(p_):
            others.append((r_["id"], oname.split(" (")[0], p_.mass + formula_mass(add)))
    for r, pep in samples:
        print(f"[{r['id']}] {r['name']}")
        run = load_run(folder / r["raw_file"], cache)
        blank_name, blank_run = blank_for(r)
        if blank_name:
            print(f"    blank: {blank_name}")
        res = analyse(r, pep, blocks, run, ppm, max_z, min_rel, min_scans)
        # chance level: the same search with another peptide's sequence on this run
        decoy_hits = None
        if len(samples) > 1:
            j = (samples.index((r, pep)) + 1) % len(samples)
            res_d = analyse(r, samples[j][1], blocks, run, ppm, max_z, min_rel, min_scans)
            decoy_hits = sum(1 for sx in res_d["species"][1:] if sx["status"] == "yes")
        res["decoy_hits"] = decoy_hits
        comps = main_components(run)
        for c in comps:
            c["explain"] = explain_mass(c["M"], pep, blocks, [o for o in others if o[0] != r["id"]], tol_ppm=ppm)
            c["blank"] = blank_level(blank_run, c["mz"]) >= 0.2 * c["inten"]
        res["components"] = comps
        sid_ = re.sub(r"[^\w.-]+", "_", r["id"])
        write_csv(data_dir / f"{sid_}_main_peaks.csv",
                  ["rt_min", "tic_rel_pct", "mono_mz", "z", "neutral_mass", "delta_vs_expected", "also_in_blank",
                   "possible_explanations"],
                  [(f"{c['rt']:.2f}", f"{c['rel']:.0f}", f"{c['mz']:.4f}", c["z"], f"{c['M']:.4f}",
                    f"{c['M'] - pep.mass:+.4f}", "yes" if c["blank"] else "",
                    " | ".join(c["explain"])) for c in comps])
        deck.add_sample(r, pep, blocks, res, run, blank_run, data_dir, folder, ppm, blank_name)
        main = res["species"][0]
        errs = [(h[0] - i.mz) / i.mz * 1e6 for i in main["sp"].ions for lbl, h in main["hits"].items() if lbl == i.label]
        line = dict(id=r["id"], name=r["name"], sequence=one_letter(pep), formula=hill(pep.formula),
                    mono_mass=f"{pep.mass:.4f}", found=res["status"],
                    rt_min=f"{run.rt[main['window'][2]]:.2f}" if main["window"] else "",
                    ions_found=" ".join(main["hits"]),
                    mean_ppm=f"{np.mean(errs):.1f}" if errs else "",
                    intensity=f"{max(h[1] for h in main['hits'].values()):.2e}" if main["hits"] else "",
                    blank=blank_name)
        pur = res.get("purity")
        line["purity_pct"] = f"{pur['pct']:.1f}" if pur else ""
        line["purity_rt"] = f"{pur['rt']:.2f}" if pur else ""
        a_, b_, k_ = res["window"]
        table = [("M", f"{pep.mass:.4f}", "", "", "")]
        for ion in main["sp"].ions:
            h = main["hits"].get(ion.label)
            table.append((f"[{ion.label}]", f"{ion.mz:.4f}", f"{h[0]:.4f}" if h else "n.d.",
                          f"{(h[0] - ion.mz) / ion.mz * 1e6:+.1f}" if h else "", line["rt_min"] if h else ""))
        pages.append(dict(
            title=r["name"] or f"Sample {r['id']}", seq=one_letter(pep), found=res["status"],
            formula=f"{hill(pep.formula)}    Exact mass (mono.): {pep.mass:.4f}    {len(pep.units)} residues",
            png=data_dir / f"{sid_}_structure.png", hplc=res.get("hplc"), table=table,
            smz=res["smz"], sint=res["sint"], labels=res["labels"], mz_range=(run.scan_lo, run.scan_hi),
            spec_title=f"{Path(r['raw_file']).stem}, spectrum averaged {run.rt[a_]:.2f}-{run.rt[b_]:.2f} min",
            rt=run.rt, tic=run.tic, blank=(blank_run.rt, blank_run.tic) if blank_run is not None else None,
            line_rt=float(run.rt[k_]), note=res.get("note", "")))
        def hit_text(sx):
            return f"{sx['sp'].name} ({run.rt[sx['window'][2]]:.2f} min, {max(h[1] for h in sx['hits'].values()):.1e})"
        line["deltas_searched"] = len(res["species"]) - 1
        line["deltas_found"] = "; ".join(hit_text(sx) for sx in res["species"][1:] if sx["status"] == "yes")
        line["deltas_weak"] = "; ".join(hit_text(sx) for sx in res["species"][1:] if sx["status"] == "weak")
        line["delta_hits"] = sum(1 for sx in res["species"][1:] if sx["status"] == "yes")
        line["delta_hits_with_wrong_sequence"] = "" if decoy_hits is None else decoy_hits
        entered, unassigned = assign_observed(parse_observed(r.get("observed_mz", "")),
                                              [sx["sp"] for sx in res["species"]])
        calc = {(sx["sp"].name, i.label): i.mz for sx in res["species"] for i in sx["sp"].all_ions}
        line["observed_entered"] = "; ".join(
            [f"{v:.4f}={'' if n == 'main' else n + ' '}{lbl} ({(v - calc[(n, lbl)]) / calc[(n, lbl)] * 1e6:+.1f} ppm)"
             for (n, lbl), v in entered.items()] + [f"{v:.4f}=no match" for v, _ in unassigned])
        line["main_peaks"] = "; ".join(
            f"{c['rt']:.2f} min M={c['M']:.3f} ({(', '.join(c['explain']) or 'unexplained')}"
            f"{', also in blank' if c['blank'] else ''})" for c in comps[:3])
        tops = [(max(h[1] for h in sx["hits"].values()), sx["sp"].name) for sx in res["species"] if sx["hits"]]
        line["strongest"] = max(tops)[1] if tops else ""
        if line["strongest"] == "main":
            line["strongest"] = "expected"
        summary.append(line)
        st = {"yes": "FOUND", "weak": "weak"}
        print(f"    expected: {st[res['status']] + ' ' + line['rt_min'] + ' min, ' + line['ions_found'] if res['found'] else 'not found'}"
              + f"; {len(res['species']) - 1} deltas searched"
              + (f", found: {', '.join(sx['sp'].name for sx in res['species'][1:] if sx['status'] == 'yes')}"
                 if any(sx["status"] == "yes" for sx in res["species"][1:]) else ""))
        if comps:
            c = comps[0]
            print(f"    biggest peak: {c['rt']:.2f} min, M={c['M']:.3f} (z={c['z']}, {c['M'] - pep.mass:+.3f} vs expected): "
                  f"{', '.join(c['explain']) or 'unexplained'}{' [also in blank]' if c['blank'] else ''}")
    deck.add_overview(summary, f"{folder.name}: overview")
    pptx_path = out / "MS_report.pptx"
    if (out / "~$MS_report.pptx").exists():   # open in PowerPoint: never write under it
        pptx_path = out / f"MS_report_{datetime.now():%Y%m%d-%H%M%S}.pptx"
        print(f"MS_report.pptx is open in PowerPoint; saving as {pptx_path.name}")
    try:
        deck.prs.save(pptx_path)
    except PermissionError:
        pptx_path = out / f"MS_report_{datetime.now():%Y%m%d-%H%M%S}.pptx"
        deck.prs.save(pptx_path)
    cols = list(dict.fromkeys(k for line in summary for k in line))
    with open(out / "summary.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(summary)
    pdf_path = out / "MS_report_notebook.pdf"     # not MS_report.pdf: that is what PowerPoint calls its own export
    try:
        made = write_pdf(pdf_path, f"{folder.name}: overview", summary, pages)
    except PermissionError:
        pdf_path = out / f"MS_report_notebook_{datetime.now():%Y%m%d-%H%M%S}.pdf"
        made = write_pdf(pdf_path, f"{folder.name}: overview", summary, pages)
    except Exception as e:          # the PDF is a copy; never lose the report over it
        made = False
        print(f"PDF not written: {e}")
    shutil.copy(folder / "samples.csv", out / "samples_used.csv")
    shutil.copy(blocks_path, out / "building_blocks_used.csv")
    print(f"\nDone: {pptx_path}\n      {out / 'summary.csv'}")
    print(f"      {pdf_path}" if made else "      (no PDF: matplotlib is not installed)")
    print(f"      {sum(1 for x in summary if x['found'] == 'yes')}/{len(summary)} samples: expected mass found"
          f" ({sum(1 for x in summary if x['found'] == 'weak')} more weak)")
    if len(summary) > 1:
        own = sum(x["delta_hits"] for x in summary)
        wrong = sum(x["delta_hits_with_wrong_sequence"] or 0 for x in summary)
        print(f"      delta hits: {own} with the right sequences, {wrong} when each run is searched with another "
              f"peptide's sequence (chance level). Treat delta hits as leads, not proof.")


# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument("command", nargs="?", default="auto",
                    choices=["auto", "setup", "check", "run", "mass", "blocks", "deltas"])
    ap.add_argument("sequence", nargs="?")
    ap.add_argument("--folder")
    ap.add_argument("--only")
    ap.add_argument("--ppm", type=float, default=5.0)
    ap.add_argument("--max-charge", type=int, default=6)
    ap.add_argument("--min-rel", type=float, default=1.0)
    ap.add_argument("--min-scans", type=int, default=8)
    ap.add_argument("--topology", choices=TOPOLOGIES)
    ap.add_argument("--bridges")
    try:
        args = ap.parse_args()
    except SystemExit:
        print("\nRun  python3 make_ms_report.py help  for the guide.")
        raise
    folder = Path(args.folder).expanduser().resolve() if args.folder else Path(__file__).resolve().parent
    try:
        blocks_path = find_blocks_file(folder)
        blocks = load_blocks(blocks_path)
        cmd = args.command
        if cmd == "auto":
            cmd = "run" if (folder / "samples.csv").exists() else "setup"
        if cmd == "setup":
            print(f"Setting up samples.csv in {folder}")
            cmd_setup(folder, blocks, force=args.command == "setup")
        elif cmd == "check":
            cmd_check(folder, blocks, args.max_charge)
        elif cmd == "run":
            only = {x.strip() for x in args.only.split(",")} if args.only else None
            cmd_run(folder, blocks, blocks_path, args.ppm, args.max_charge, only, args.min_rel, args.min_scans)
        elif cmd == "mass":
            if not args.sequence:
                raise UserError('give a sequence, e.g.  python3 make_ms_report.py mass "Ac-KGG-NH2"')
            cmd_mass(args.sequence, args.topology, blocks, args.max_charge, args.bridges or "")
        elif cmd == "blocks":
            cmd_blocks(blocks, blocks_path)
        elif cmd == "deltas":
            cmd_deltas(blocks, blocks_path)
    except UserError as e:
        print(f"\n{e}\n\n(python3 make_ms_report.py help  shows the guide)")
        sys.exit(1)


if __name__ == "__main__":
    main()
