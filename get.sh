#!/usr/bin/env bash
# Підготовка робочої теки sdi-tools у новому середовищі. Автор: В'ячеслав Плескач.
# Запуск із теки клону: bash get.sh
#   1) завантажує скрипт ЄДРСР з публічного репозиторію edrsr-tool як edrsr_vybirka.py (єдина копія скрипта);
#   2) запускає його перевірку доступу (check);
#   3) якщо поруч склоновано postanovi_VP-VS-monitoring — робить посилання на реєстр постанов ВП ВС.
set -e
cd "$(dirname "$0")"
curl -sSfL https://raw.githubusercontent.com/viacheslavpleskach-cpu/edrsr-tool/main/edrsr.py -o edrsr_vybirka.py
echo "edrsr_vybirka.py: $(grep -m1 -o 'Версія [0-9.]*\|version [0-9.]*' edrsr_vybirka.py || echo завантажено)"
python3 edrsr_vybirka.py check || echo "! check ЄДРСР не пройшов — див. повідомлення вище"
for d in ../postanovi_vp-vs-monitoring ../postanovi_VP-VS-monitoring; do
  if [ -f "$d/indeks/Реєстр_постанов_ВП_ВС.csv" ]; then
    ln -sf "$d/indeks/Реєстр_постанов_ВП_ВС.csv" Реєстр_постанов_ВП_ВС.csv
    echo "Реєстр ВП ВС: посилання на $d/indeks"
    break
  fi
done
python3 -c "import duckdb, openpyxl, docx" 2>/dev/null || pip install -q --break-system-packages duckdb openpyxl python-docx
echo "Готово."
