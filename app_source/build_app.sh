#!/bin/bash
# Rebuild "MS Report.app" and "MS Report for Mac.zip" from the files in Script/.
# Run after changing make_ms_report.py, building_blocks.csv or anything in app_source/:
#   bash ~/Downloads/MSdata/Script/app_source/build_app.sh
set -euo pipefail
SRC="$(cd "$(dirname "$0")" && pwd)"
OUT="$(dirname "$SRC")"
APP="$OUT/MS Report.app"
rm -rf "$APP"
osacompile -o "$APP" "$SRC/main.applescript"
R="$APP/Contents/Resources"
cp "$SRC/msreport.sh" "$SRC/MS Report manual.html" "$SRC/MS Report.icns" \
   "$OUT/make_ms_report.py" "$OUT/dashboard.py" "$OUT/dashboard.html" "$OUT/building_blocks.csv" "$R/"
ICON="$(/usr/libexec/PlistBuddy -c "Print :CFBundleIconFile" "$APP/Contents/Info.plist")"   # applet, or droplet once the script has "on open"
rm -f "$R/applet.icns" "$R/droplet.icns"
cp "$SRC/MS Report.icns" "$R/$ICON.icns"            # the app's own icon instead of the script icon
rm -f "$R/Assets.car"                                # its icon would win over applet.icns
/usr/libexec/PlistBuddy -c "Delete :CFBundleIconName" "$APP/Contents/Info.plist" 2>/dev/null || true
chmod +x "$R/msreport.sh"
codesign --force --deep -s - "$APP"
rm -f "$OUT/MS Report for Mac.zip"
STAGE="$(mktemp -d)"
ditto "$APP" "$STAGE/MS Report for Mac/MS Report.app"
cp "$SRC/MS Report manual.html" "$STAGE/MS Report for Mac/"
(cd "$STAGE" && ditto -c -k --keepParent "MS Report for Mac" "$OUT/MS Report for Mac.zip")
rm -rf "$STAGE"
echo "Built: $APP"
echo "       $OUT/MS Report for Mac.zip"
