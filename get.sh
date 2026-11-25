#!/usr/bin/env bash
# Підготовка робочої теки sdi-tools у новому середовищі. Автор: В'ячеслав Плескач.
# Запуск із теки клону: bash get.sh
#   1) завантажує скрипт ЄДРСР з публічного репозиторію edrsr-tool як edrsr_vybirka.py (єдина копія скрипта);
#   2) запускає його перевірку доступу (check);
#   3) якщо поруч склоновано postanovi_VP-VS-monitoring — робить посилання на реєстр постанов ВП ВС;
#   4) ставить українську мову для OCR сканів (tesseract-ocr-ukr через apt; github з tessdata заблоковано).
set -e
cd "$(dirname "$0")"
curl -sSfL https://raw.githubusercontent.com/viacheslavpleskach-cpu/edrsr-tool/main/edrsr.py -o edrsr_vybirka.py || {
  echo "! raw.githubusercontent.com недоступний — беру edrsr.py клонуванням edrsr-tool"
  rm -rf .edrsr-tool
  git clone --depth 1 -q https://github.com/viacheslavpleskach-cpu/edrsr-tool .edrsr-tool
  cp .edrsr-tool/edrsr.py edrsr_vybirka.py
  rm -rf .edrsr-tool
}
echo "edrsr_vybirka.py: $(grep -m1 -o 'Версія [0-9.]*\|version [0-9.]*' edrsr_vybirka.py || echo завантажено)"
python3 edrsr_vybirka.py check || echo "! check ЄДРСР не пройшов — див. повідомлення вище"
for d in ../postanovi_vp-vs-monitoring ../postanovi_VP-VS-monitoring; do
  if [ -f "$d/indeks/Реєстр_постанов_ВП_ВС.csv" ]; then
    ln -sf "$d/indeks/Реєстр_постанов_ВП_ВС.csv" Реєстр_постанов_ВП_ВС.csv
    echo "Реєстр ВП ВС: посилання на $d/indeks"
    break
  fi
done
python3 -c "import duckdb, openpyxl, docx, py7zr" 2>/dev/null || pip install -q --break-system-packages duckdb openpyxl python-docx py7zr
if command -v tesseract >/dev/null && ! tesseract --list-langs 2>/dev/null | grep -qx ukr; then
  (apt-get install -y -qq tesseract-ocr-ukr >/dev/null 2>&1 || (apt-get update -qq >/dev/null 2>&1 && apt-get install -y -qq tesseract-ocr-ukr >/dev/null 2>&1)) \
    && echo "OCR: українську мову встановлено" || echo "! OCR: tesseract-ocr-ukr не встановлено (apt недоступний) — скани читати візуально"
fi
echo "Готово."
