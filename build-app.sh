#!/bin/bash
# Build AnyView.app bundle from the SPM executable
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

CONFIG="${1:-release}"
if [ "$CONFIG" = "release" ]; then
    swift build -c release
    EXECUTABLE=".build/release/AnyView"
    BUILD_DIR=".build/release"
else
    swift build
    EXECUTABLE=".build/debug/AnyView"
    BUILD_DIR=".build/debug"
fi

APP_DIR="$SCRIPT_DIR/.build/AnyView.app"
CONTENTS="$APP_DIR/Contents"
MACOS="$CONTENTS/MacOS"

# Clean previous bundle
rm -rf "$APP_DIR"

# Create bundle structure
mkdir -p "$MACOS"

# Copy executable
cp "$EXECUTABLE" "$MACOS/AnyView"

# Copy Info.plist
cp "$SCRIPT_DIR/Sources/AnyViewApp/Info.plist" "$CONTENTS/Info.plist"

# Copy icon
RESOURCES="$CONTENTS/Resources"
mkdir -p "$RESOURCES"
cp "$SCRIPT_DIR/Resources/AppIcon.icns" "$RESOURCES/AppIcon.icns"

# Copy SPM resource bundle(s)
for spm_bundle in "$SCRIPT_DIR/$BUILD_DIR"/*.bundle; do
    [ -d "$spm_bundle" ] && cp -R "$spm_bundle" "$RESOURCES/"
done

# Bundle docmod CLI if available (for .docx rendering)
# Search order: DOCMOD_PATH env, ~/.local/bin, ~/.docmod/bin, PATH
DOCMOD_BIN=""
if [ -n "${DOCMOD_PATH:-}" ] && [ -x "$DOCMOD_PATH" ]; then
    DOCMOD_BIN="$DOCMOD_PATH"
elif [ -x "$HOME/.local/bin/docmod" ]; then
    DOCMOD_BIN="$HOME/.local/bin/docmod"
elif [ -x "$HOME/.docmod/bin/docmod" ]; then
    DOCMOD_BIN="$HOME/.docmod/bin/docmod"
elif command -v docmod &>/dev/null; then
    DOCMOD_BIN="$(command -v docmod)"
fi

if [ -n "$DOCMOD_BIN" ]; then
    cp "$DOCMOD_BIN" "$MACOS/docmod"
    echo "Bundled docmod CLI from $DOCMOD_BIN"
else
    echo "Note: docmod CLI not found — .docx preview will not work without it"
fi

# Ad-hoc code sign
codesign --force --deep --sign - "$APP_DIR" 2>/dev/null && echo "Code signed (ad-hoc)" || true

echo "Built: $APP_DIR"
