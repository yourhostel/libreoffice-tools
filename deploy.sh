#!/bin/bash
set -e

ODS="$HOME/Scripts/libreoffice-tools/temp/saperka_dev.ods"
DST="$HOME/.config/libreoffice/4/user/basic/Standard"
# DST="$HOME/Scripts/libreoffice-tools/basic/Standard"

TMP="$(mktemp -d)"

echo "[*] Чищу старі файли"
rm -f "$DST"/*.xba
rm -f "$DST/script.xlb"

echo "[*] Розпаковую $ODS"
unzip -q "$ODS" -d "$TMP"

echo "[*] Копіюю модулі"
for f in "$TMP/Basic/Standard/"*.xml; do
    base="$(basename "$f" .xml)"
    if [[ "$base" == "script-lb" ]]; then
        cp "$f" "$DST/script.xlb"
        echo "[+] Додав script.xlb"
    else
        cp "$f" "$DST/$base.xba"
        echo "[+] Додав модуль $base.xba"
    fi
done

echo "[*] Прибираю тимчасові файли"
rm -rf "$TMP"

echo "[*] Готово! Перезапусти LibreOffice."





