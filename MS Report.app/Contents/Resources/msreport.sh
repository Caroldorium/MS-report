#!/bin/bash
# Helper behind "MS Report.app" (the app shows dialogs; this does the work).
#   msreport.sh ready                 -> prints "ready" when the tools are installed
#   msreport.sh installed_before      -> "yes" if an older version was installed (update, not first install)
#   msreport.sh dashboard             -> opens the MS Report window (dashboard.py) and returns
#   msreport.sh install               -> one-time install (uv + Python packages + ThermoRawFileParser)
#   msreport.sh install_bg            -> the same in the background, for the progress window
#   msreport.sh install_status        -> "step|what it is doing", "done" or "failed|why"
#   msreport.sh state  <folder>       -> new | samples | noraw
#   msreport.sh setup|check|run <folder>
#   msreport.sh open_csv <folder>     -> open samples.csv in Excel (or the default app)
set -u
set -o pipefail
RES="$(cd "$(dirname "$0")" && pwd)"
APP_HOME="$HOME/Documents/MS Report"           # building_blocks.csv lives here; add monomers here
UV="$HOME/.local/bin/uv"
TRFP_DIR="$HOME/software/ThermoRawFileParser"
TRFP="$TRFP_DIR/ThermoRawFileParser"
TRFP_URL="https://github.com/CompOmics/ThermoRawFileParser/releases/download/v.2.0.0-dev"
MARKER="$APP_HOME/.installed-v3"      # v2: dashboard (pywebview) packages; v3: matplotlib for the PDF
INSTALL_LOG="$APP_HOME/install.log"
INSTALL_PID="$APP_HOME/.install.pid"
export MS_REPORT_HOME="$APP_HOME"
export PATH="$HOME/.local/bin:/usr/bin:/bin:/usr/sbin:/sbin"

sync_blocks() {   # keep Documents/MS Report/building_blocks.csv current unless the user has edited it
    local shipped="$RES/building_blocks.csv" user="$APP_HOME/building_blocks.csv" stamp="$APP_HOME/.blocks-shipped"
    local new; new=$(shasum "$shipped" | cut -d' ' -f1)
    mkdir -p "$APP_HOME"
    if [ ! -f "$user" ]; then
        cp "$shipped" "$user"; echo "$new" > "$stamp"
    elif [ -f "$stamp" ] && [ "$(shasum "$user" | cut -d' ' -f1)" = "$(cat "$stamp")" ] && [ "$(cat "$stamp")" != "$new" ]; then
        cp "$shipped" "$user"; echo "$new" > "$stamp"           # untouched copy: update it
    elif [ ! -f "$stamp" ] && ! head -1 "$user" | grep -q group; then
        cp "$user" "$user.old"; cp "$shipped" "$user"; echo "$new" > "$stamp"   # first-version copy: update, keep old
    fi
    cp "$RES/MS Report manual.html" "$APP_HOME/" 2>/dev/null || true
}

script() {   # run make_ms_report.py; on failure send its output to stderr so the app can show it
    local out code
    out=$("$UV" run --quiet --python 3.12 "$RES/make_ms_report.py" "$@" 2>&1)
    code=$?
    if [ $code -ne 0 ]; then echo "$out" | tail -25 >&2; exit $code; fi
    echo "$out"
}

case "${1:-}" in
ready)
    if [ -x "$UV" ] && [ -x "$TRFP" ] && [ -f "$MARKER" ]; then echo ready; else echo missing; fi
    ;;
installed_before)
    if ls "$APP_HOME"/.installed-v* >/dev/null 2>&1; then echo yes; else echo no; fi
    ;;
install)
    set -e
    mkdir -p "$APP_HOME"
    echo "STEP 1|Getting the Python installer (uv)"
    if [ ! -x "$UV" ]; then
        echo "Installing uv ..."
        curl -LsSf https://astral.sh/uv/install.sh | env UV_NO_MODIFY_PATH=1 sh
    fi
    echo "STEP 2|Getting the reader for .raw files (ThermoRawFileParser)"
    if [ ! -x "$TRFP" ]; then
        echo "Installing ThermoRawFileParser ..."
        if [ "$(uname -m)" = "arm64" ]; then zip=ThermoRawFileParser-v.2.0.0-dev-osx-arm64.zip
        else zip=ThermoRawFileParser-v.2.0.0-dev-osx.zip; fi
        mkdir -p "$TRFP_DIR"
        curl -fsSL -o "$TRFP_DIR/trfp.zip" "$TRFP_URL/$zip"
        (cd "$TRFP_DIR" && unzip -q -o trfp.zip && rm trfp.zip)
        if [ ! -f "$TRFP" ]; then     # some builds unpack into a sub-folder: move it up
            found=$(find "$TRFP_DIR" -name ThermoRawFileParser -type f | head -1)
            [ -n "$found" ] && mv "$(dirname "$found")"/* "$TRFP_DIR"/
        fi
        chmod +x "$TRFP"
        xattr -dr com.apple.quarantine "$TRFP_DIR" 2>/dev/null || true
        # macOS kills unsigned downloaded binaries; sign them locally (ad hoc)
        codesign --force -s - "$TRFP" >/dev/null 2>&1
        for f in "$TRFP_DIR"/*.dylib; do codesign --force -s - "$f" >/dev/null 2>&1 || true; done
        "$TRFP" --version >/dev/null
    fi
    sync_blocks
    echo "STEP 3|Downloading the chemistry and plotting packages (the long part)"
    "$UV" run --quiet --python 3.12 "$RES/make_ms_report.py" blocks >/dev/null
    echo "STEP 4|Downloading the window packages"
    "$UV" run --quiet --python 3.12 "$RES/dashboard.py" --help >/dev/null     # window packages
    touch "$MARKER"
    echo "done"
    ;;
install_bg)
    mkdir -p "$APP_HOME"
    : > "$INSTALL_LOG"
    nohup /bin/bash "$0" install >"$INSTALL_LOG" 2>&1 </dev/null &
    echo $! > "$INSTALL_PID"
    ;;
install_status)
    if [ "$(tail -1 "$INSTALL_LOG" 2>/dev/null)" = "done" ]; then echo done
    elif [ -f "$INSTALL_PID" ] && kill -0 "$(cat "$INSTALL_PID")" 2>/dev/null; then
        step=$(grep '^STEP ' "$INSTALL_LOG" 2>/dev/null | tail -1)
        if [ -n "$step" ]; then echo "${step#STEP }"; else echo "0|Starting"; fi
    else
        echo "failed|$(grep -v '^STEP ' "$INSTALL_LOG" 2>/dev/null | tail -4 | tr '\n' ' ')"
    fi
    ;;
state)
    d="$2"
    shopt -s nullglob nocaseglob
    data=("$d"/*.raw "$d"/*.mzml)
    if [ -f "$d/samples.csv" ]; then echo samples
    elif [ ${#data[@]} -gt 0 ]; then echo new
    else echo noraw; fi
    ;;
setup)  sync_blocks; script setup --folder "$2" ;;
check)  sync_blocks; script check --folder "$2" ;;
run)
    sync_blocks
    out=$(script run --folder "$2") || exit $?
    echo "$out" | grep -vE "converting|reading|^\[|^    expected|^    biggest|skipping" | tail -6
    latest=$(ls -t "$2/MS_report/"*.pptx 2>/dev/null | head -1)
    echo "PPTX=$latest"
    ;;
dashboard)
    sync_blocks
    log="$APP_HOME/dashboard.log"
    if [ -n "${2:-}" ]; then      # a folder (or a file in it) dropped on the app icon
        target="$2"; [ -d "$target" ] || target="$(dirname "$target")"
        nohup "$UV" run --quiet --python 3.12 "$RES/dashboard.py" --folder "$target" >"$log" 2>&1 </dev/null &
    else
        nohup "$UV" run --quiet --python 3.12 "$RES/dashboard.py" >"$log" 2>&1 </dev/null &
    fi
    pid=$!
    for _ in $(seq 1 80); do     # wait until the window is up (or an open one was brought forward)
        if grep -qE "MS Report dashboard:|already open" "$log" 2>/dev/null; then exit 0; fi
        if ! kill -0 "$pid" 2>/dev/null; then tail -15 "$log" >&2; exit 1; fi
        sleep 0.5
    done
    ;;
open_csv)
    open -b com.microsoft.Excel "$2/samples.csv" 2>/dev/null || open "$2/samples.csv"
    ;;
*)
    echo "usage: msreport.sh ready|installed_before|install|install_bg|install_status|dashboard|state|setup|check|run|open_csv [folder]" >&2; exit 2 ;;
esac
