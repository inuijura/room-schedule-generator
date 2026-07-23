#!/usr/bin/env bash

set -Eeuo pipefail

APP_NAME="room-schedule-generator"
VERSION="1.0"
GEM_NAME="${APP_NAME}-${VERSION}.gem"

REPOSITORY="inuijura/room-schedule-generator"
DOWNLOAD_URL="https://github.com/${REPOSITORY}/releases/download/v${VERSION}/${GEM_NAME}"

TEMP_DIR="$(mktemp -d)"
GEM_FILE="${TEMP_DIR}/${GEM_NAME}"

cleanup() {
    rm -rf "$TEMP_DIR"
}

trap cleanup EXIT

echo "Downloading $APP_NAME..."
curl -fL "$DOWNLOAD_URL" -o "$GEM_FILE"

echo "Installing $APP_NAME..."
gem install --user-install "$GEM_FILE"

GEM_BIN="$(ruby -r rubygems -e 'puts Gem.user_dir')/bin"

echo "Gem executable path:"
echo "$GEM_BIN"

case ":$PATH:" in
    *":$GEM_BIN:"*)
        echo "PATH already contains $GEM_BIN"
        exit 0
        ;;
esac

CURRENT_SHELL="$(basename "${SHELL:-}")"

case "$CURRENT_SHELL" in
    bash)
        RC_FILE="$HOME/.bashrc"
        ;;
    zsh)
        RC_FILE="$HOME/.zshrc"
        ;;
    fish)
        RC_FILE="$HOME/.config/fish/config.fish"
        ;;
    *)
        echo "Unsupported shell: ${CURRENT_SHELL:-unknown}"
        echo
        echo "Please add the following directory to PATH:"
        echo "$GEM_BIN"
        exit 0
        ;;
esac

if [ ! -f "$RC_FILE" ]; then
    echo "$RC_FILE does not exist."
    echo
    echo "Please add the following directory to PATH manually:"
    echo "$GEM_BIN"
    exit 0
fi

if grep -Fq "# BEGIN $APP_NAME" "$RC_FILE"; then
    echo "PATH setting already exists in $RC_FILE."
    exit 0
fi

{
    echo
    echo "# BEGIN $APP_NAME"

    case "$CURRENT_SHELL" in
        fish)
            echo "fish_add_path \"$GEM_BIN\""
            ;;
        *)
            echo "export PATH=\"$GEM_BIN:\$PATH\""
            ;;
    esac

    echo "# END $APP_NAME"
} >> "$RC_FILE"

echo
echo "Installation completed."
echo
echo "Restart the terminal or execute:"
echo "source \"$RC_FILE\""