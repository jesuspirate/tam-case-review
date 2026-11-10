#!/bin/bash
# Builds "TAM Case Review.app" and a drag-to-install .dmg in app/dist/.
# Run from anywhere:  bash app/build_mac.sh
set -euo pipefail
cd "$(dirname "$0")"
NAME="TAM Case Review"
VERSION="$(sed -n 's/^VERSION = "\(.*\)"/\1/p' tam_app.py)"

if [ ! -x .venv/bin/python ]; then
  python3 -m venv .venv
fi
.venv/bin/pip install -q --upgrade pip
.venv/bin/pip install -q -r requirements-build.txt

# Organization profile (vendor name, case types, Jira URL): bundled when ../profile.json exists.
PROFILE_ARGS=()
if [ -f ../profile.json ]; then
  PROFILE_ARGS=(--add-data "../profile.json:.")
  echo "Bundling profile.json"
else
  echo "No profile.json: building with neutral defaults (see profiles/example.json)"
fi

rm -rf build "dist/$NAME.app" "dist/$NAME.dmg"
.venv/bin/python -m PyInstaller --noconfirm --windowed --clean \
  --name "$NAME" \
  --icon assets/icon.png \
  --osx-bundle-identifier com.tamtoolkit.casereview \
  --paths .. \
  --add-data "ui.html:." \
  ${PROFILE_ARGS[@]+"${PROFILE_ARGS[@]}"} \
  tam_app.py

# Version number and a friendlier Finder description in the app's Info.plist.
PLIST="dist/$NAME.app/Contents/Info.plist"
/usr/libexec/PlistBuddy -c "Set :CFBundleShortVersionString $VERSION" "$PLIST"
/usr/libexec/PlistBuddy -c "Add :NSHumanReadableCopyright string 'Internal tool for TAM customer case reviews'" "$PLIST" 2>/dev/null || true
codesign --force --deep --sign - "dist/$NAME.app"     # ad-hoc signature (no Apple Developer ID)

# Disk image: the app plus an Applications shortcut to drag it onto.
STAGE="$(mktemp -d)"
cp -R "dist/$NAME.app" "$STAGE/"
ln -s /Applications "$STAGE/Applications"
hdiutil create -quiet -volname "$NAME" -srcfolder "$STAGE" -ov -format UDZO "dist/$NAME.dmg"
rm -rf "$STAGE"

echo
echo "Built: app/dist/$NAME.app"
echo "       app/dist/$NAME.dmg   <- send this one to coworkers"
