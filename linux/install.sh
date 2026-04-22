#!/usr/bin/env bash
# Install AnyView to ~/.local (user-local, no sudo required).
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

PREFIX="${PREFIX:-$HOME/.local}"
BIN_DIR="$PREFIX/bin"
APP_DIR="$PREFIX/share/applications"
ICON_DIR="$PREFIX/share/icons/hicolor/scalable/apps"
MIME_DIR="$PREFIX/share/mime/packages"

echo "→ Building release binary…"
cargo build --release

echo "→ Installing to $PREFIX"
install -d "$BIN_DIR" "$APP_DIR" "$ICON_DIR" "$MIME_DIR"

install -m 0755 target/release/anyview "$BIN_DIR/anyview"
install -m 0644 data/anyview.desktop "$APP_DIR/anyview.desktop"

if [ -f data/anyview.svg ]; then
    install -m 0644 data/anyview.svg "$ICON_DIR/anyview.svg"
fi

if [ -f data/anyview.xml ]; then
    install -m 0644 data/anyview.xml "$MIME_DIR/anyview.xml"
    if command -v update-mime-database &>/dev/null; then
        update-mime-database "$PREFIX/share/mime" || true
    fi
fi

if command -v update-desktop-database &>/dev/null; then
    update-desktop-database "$APP_DIR" || true
fi

if command -v gtk-update-icon-cache &>/dev/null; then
    gtk-update-icon-cache -f -t "$PREFIX/share/icons/hicolor" 2>/dev/null || true
fi

echo "✓ Installed: $BIN_DIR/anyview"
echo "  Desktop:   $APP_DIR/anyview.desktop"
echo ""
echo "Make sure $BIN_DIR is in your PATH, then run: anyview <file>"
