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

# Bundle tectonic (LaTeX compiler) if available
TECTONIC_BIN=""
if [ -n "${TECTONIC_PATH:-}" ] && [ -x "$TECTONIC_PATH" ]; then
    TECTONIC_BIN="$TECTONIC_PATH"
elif [ -x "/opt/homebrew/bin/tectonic" ]; then
    TECTONIC_BIN="/opt/homebrew/bin/tectonic"
elif [ -x "/usr/local/bin/tectonic" ]; then
    TECTONIC_BIN="/usr/local/bin/tectonic"
elif command -v tectonic &>/dev/null; then
    TECTONIC_BIN="$(command -v tectonic)"
fi

if [ -n "$TECTONIC_BIN" ]; then
    cp "$TECTONIC_BIN" "$MACOS/tectonic"
    echo "Bundled tectonic from $TECTONIC_BIN"
else
    echo "Note: tectonic not found — .tex compilation will not work without it"
fi

# Bundle ffmpeg if available (for non-native video format transcoding)
FFMPEG_BIN=""
if [ -n "${FFMPEG_PATH:-}" ] && [ -x "$FFMPEG_PATH" ]; then
    FFMPEG_BIN="$FFMPEG_PATH"
elif [ -x "/opt/homebrew/bin/ffmpeg" ]; then
    FFMPEG_BIN="/opt/homebrew/bin/ffmpeg"
elif [ -x "/usr/local/bin/ffmpeg" ]; then
    FFMPEG_BIN="/usr/local/bin/ffmpeg"
elif command -v ffmpeg &>/dev/null; then
    FFMPEG_BIN="$(command -v ffmpeg)"
fi

if [ -n "$FFMPEG_BIN" ]; then
    cp "$FFMPEG_BIN" "$MACOS/ffmpeg"
    echo "Bundled ffmpeg from $FFMPEG_BIN"
else
    echo "Note: ffmpeg not found — mkv/avi/flv/rmvb transcoding will not work without it (brew install ffmpeg)"
fi

# Bundle av CLI script
cp "$SCRIPT_DIR/scripts/av" "$RESOURCES/av"
chmod +x "$RESOURCES/av"

# Install av to ~/.local/bin/av
LOCAL_BIN="$HOME/.local/bin"
mkdir -p "$LOCAL_BIN"
cp "$RESOURCES/av" "$LOCAL_BIN/av"
chmod +x "$LOCAL_BIN/av"
echo "Installed: $LOCAL_BIN/av"
echo "  (make sure $LOCAL_BIN is in your PATH)"

# Ad-hoc code sign
codesign --force --deep --sign - "$APP_DIR" 2>/dev/null && echo "Code signed (ad-hoc)" || true

echo "Built: $APP_DIR"
