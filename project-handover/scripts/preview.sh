#!/usr/bin/env bash
# Render a .pptx to PDF + per-slide PNGs for visual QA.
# Usage: preview.sh <deck.pptx> [out_dir]   (default out_dir: <deck_dir>/preview)
#
# The deck uses Microsoft JhengHei / Consolas (what recipients' Office has). LibreOffice
# here usually can't see those (or, inside a sandbox, any system font) and renders CJK
# as tofu boxes. So for the preview ONLY we:
#   1. render a temp copy with the fonts swapped to Heiti TC / Menlo, and
#   2. run soffice with a private profile whose user/fonts dir links the macOS font files.
# The .pptx you deliver is untouched.
set -euo pipefail
deck="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
out="${2:-$(dirname "$deck")/preview}"; mkdir -p "$out"; out="$(cd "$out" && pwd)"
name="$(basename "${deck%.*}")"
tmp="$(mktemp -d)"; trap 'rm -rf "$tmp"' EXIT

profile="${HANDOVER_LO_PROFILE:-${TMPDIR:-/tmp}/project-handover-lo-profile}"
mkdir -p "$profile/user/fonts"
for f in "/System/Library/Fonts/STHeiti Medium.ttc" "/System/Library/Fonts/STHeiti Light.ttc" \
         "/System/Library/Fonts/Menlo.ttc"; do
  [ -f "$f" ] && [ ! -e "$profile/user/fonts/$(basename "$f")" ] && ln -s "$f" "$profile/user/fonts/" || true
done
cjk="Heiti TC"; mono="Menlo"
[ -f "/System/Library/Fonts/STHeiti Medium.ttc" ] || { cjk="Noto Sans CJK TC"; mono="DejaVu Sans Mono"; }

python3 - "$deck" "$tmp/$name.pptx" "$cjk" "$mono" <<'PY'
import sys, zipfile
src, dst, cjk, mono = sys.argv[1:5]
with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
    for item in zi.infolist():
        data = zi.read(item.filename)
        if item.filename.endswith(".xml"):
            data = (data.decode("utf-8").replace('typeface="Microsoft JhengHei"', f'typeface="{cjk}"')
                    .replace('typeface="Consolas"', f'typeface="{mono}"').encode("utf-8"))
        zo.writestr(item, data)
PY
soffice "-env:UserInstallation=file://$profile" --headless --convert-to pdf --outdir "$out" \
  "$tmp/$name.pptx" >/dev/null 2>&1
rm -f "$out"/slide-*.png
pdftoppm -png -r 60 "$out/$name.pdf" "$out/slide"
echo "preview font: $cjk / $mono"
echo "pdf: $out/$name.pdf"
ls "$out"/slide-*.png
