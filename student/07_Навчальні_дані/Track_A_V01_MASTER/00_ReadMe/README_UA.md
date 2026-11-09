# Track A / Variant 01 / STARTER

Це еталонний синтетичний навчальний набір для курсу «Аналітика даних в оборонній галузі».

## Наскрізний процес
Контракт -> специфікація -> версія плану -> поставка -> приймання -> рахунок -> оплата.

## Основний баланс первинних джерел
- Input rows: 24 719
- Accepted / Rejected: визначаєте самостійно (контроль: Input = Accepted + Rejected)

У цей баланс НЕ включені вторинні контрольні джерела warehouse_movements.txt, legacy_deliveries_fixed_width.txt та історичний файл для лабораторної Parquet.

## Формати
XLSX, CSV, TXT, fixed-width TXT, LOG, digital PDF, semi-structured PDF, scanned PDF, gzip CSV for DuckDB/Parquet lab.

## DuckDB / Parquet
У поточному середовищі генерації бібліотека DuckDB недоступна, тому пакет містить готові staging CSV та скрипт 05_DuckDB/build_training_A_V01.sql. На машині з DuckDB виконайте build_training_A_V01.bat або CLI-команду з коментаря у SQL-файлі. Скрипт створить training_A_V01.duckdb і historical_payments.parquet.

## Staging
Папка 04_STAGING видається окремо після перевірки HW04 (вона містить очищені таблиці).

## 13-й файл
01_Source/Track_A/ADD_AFTER_DEMO/deliveries_2027_01.csv копіюється в Deliveries лише за командою викладача (HW03, захист).

## Важливо
Усі назви, організації, постачальники, документи та суми синтетичні. Набір не містить реальних оборонних або персональних даних.
