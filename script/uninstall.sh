#!/usr/bin/env bash

set -e

APP_NAME="room-schedule-generator"


echo "Uninstalling $APP_NAME..."


# gem uninstall
if gem list -i "$APP_NAME" > /dev/null 2>&1; then
    gem uninstall "$APP_NAME"
else
    echo "$APP_NAME is not installed."
fi

# ログインシェルの種類を取得
CURRENT_SHELL=$(basename "$SHELL")

# .xxxrc ファイルのパスを決定
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
        echo "Unsupported shell: $CURRENT_SHELL"
        exit 0
        ;;
esac


# .xxxrc がなければ終了
if [ ! -f "$RC_FILE" ]; then
    exit 0
fi


# install.sh が追加した範囲のパスの記述だけ削除
if grep -q "# BEGIN $APP_NAME" "$RC_FILE"; then

    sed -i.bak \
        "/# BEGIN $APP_NAME/,/# END $APP_NAME/d" \
        "$RC_FILE"

    rm -f "${RC_FILE}.bak"

    echo "Removed PATH setting from:"
    echo "$RC_FILE"

else
    echo "PATH setting not found."
fi


echo ""
echo "Uninstallation completed."