#!/bin/bash
# PPTXをPDFへ変換する統合スクリプト．
#
# 実行環境がWSLであり，かつWindows側にPowerPointが実際にインストールされている場合は，
# `convert_to_pdf.ps1`（PowerPoint COMオートメーション）を用いて変換する．
# これは，PowerPoint本体のレンダリングを用いることで，LibreOfficeで発生する数式・記号
# オブジェクトの消失を避けるためである．
#
# WSLでない場合，PowerPointが利用できない場合，またはPowerPoint経由の変換が失敗した場合は，
# LibreOfficeへフォールバックする．LibreOfficeでの変換結果は，数式・記号オブジェクトが
# 正しく表示されない場合があるため，生成されたPDFを必ず目視確認すること．
#
# 使い方:
#   convert_to_pdf.sh <入力pptxのパス> [出力pdfのパス]

set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
INPUT="${1:-}"
OUTPUT="${2:-}"

if [ -z "$INPUT" ] || [ ! -f "$INPUT" ]; then
    echo "[convert_to_pdf] 入力ファイルが指定されていない，または存在しない: $INPUT" >&2
    exit 1
fi

is_wsl() {
    grep -qi microsoft /proc/version 2>/dev/null
}

powershell_usable() {
    # WSL⇔Windows間のプロセス実行（interop）自体が機能しない環境では，
    # powershell.exeを呼び出すこと自体がエラーとなる場合があるため，
    # 単純なコマンドの実行可否も含めて確認する．
    command -v powershell.exe >/dev/null 2>&1 || return 1
    timeout 15 powershell.exe -NoProfile -Command "exit 0" >/dev/null 2>&1
}

powerpoint_available() {
    timeout 20 powershell.exe -NoProfile -Command \
        "try { \$p = New-Object -ComObject PowerPoint.Application; \$p.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject(\$p) | Out-Null; exit 0 } catch { exit 1 }" \
        >/dev/null 2>&1
}

convert_with_powerpoint() {
    local win_input win_script
    win_input="$(wslpath -w "$INPUT")"
    win_script="$(wslpath -w "$SCRIPT_DIR/convert_to_pdf.ps1")"

    if [ -n "$OUTPUT" ]; then
        local win_output
        win_output="$(wslpath -w "$OUTPUT")"
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$win_script" -InputFile "$win_input" -OutputFile "$win_output"
    else
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$win_script" -InputFile "$win_input"
    fi
}

convert_with_libreoffice() {
    local outdir default_pdf
    outdir="$(dirname "${OUTPUT:-$INPUT}")"
    libreoffice --headless --convert-to pdf --outdir "$outdir" "$INPUT" >&2
    default_pdf="$outdir/$(basename "${INPUT%.*}").pdf"
    if [ -n "$OUTPUT" ] && [ "$default_pdf" != "$OUTPUT" ] && [ -f "$default_pdf" ]; then
        mv "$default_pdf" "$OUTPUT"
    fi
}

if is_wsl && powershell_usable && powerpoint_available; then
    echo "[convert_to_pdf] WSL上でWindowsのPowerPointが利用可能なため，PowerPoint経由で変換する．" >&2
    if convert_with_powerpoint; then
        exit 0
    fi
    echo "[convert_to_pdf] PowerPoint経由の変換に失敗したため，LibreOfficeにフォールバックする．" >&2
else
    echo "[convert_to_pdf] WindowsのPowerPointが利用できないため，LibreOfficeで変換する．数式・記号オブジェクトが正しく表示されない場合があるため，生成されたPDFを必ず目視確認すること．" >&2
fi

convert_with_libreoffice
