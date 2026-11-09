# Track A / Variant 01 / STARTER

Це еталонний синтетичний навчальний набір для курсу «Аналітика даних в оборонній галузі».

## Наскрізний процес
Контракт -> специфікація -> версія плану -> поставка -> приймання -> рахунок -> оплата.

## Основний баланс первинних джерел
- Input rows: 24 719
- Accepted after base QA: 24 550
- Rejected: 169
- Перевірка: 24 719 = 24 550 + 169

У цей баланс НЕ включені вторинні контрольні джерела warehouse_movements.txt, legacy_deliveries_fixed_width.txt та історичний файл для лабораторної Parquet.

## Golden aggregates
- Original Contract Amount: 148 500 000.00
- Current Contract Amount: 152 350 000.00
- Current Plan: 139 800 000.00
- Delivered: 112 640 000.00
- Accepted: 108 910 000.00
- Invoiced: 110 240 000.00
- Paid: 101 570 000.00

## Формати
XLSX, CSV, TXT, fixed-width TXT, LOG, digital PDF, semi-structured PDF, scanned PDF, gzip CSV for DuckDB/Parquet lab.

## DuckDB / Parquet
У поточному середовищі генерації бібліотека DuckDB недоступна, тому пакет містить готові staging CSV та скрипт 05_DuckDB/build_training_A_V01.sql. На машині з DuckDB виконайте build_training_A_V01.bat або CLI-команду з коментаря у SQL-файлі. Скрипт створить training_A_V01.duckdb і historical_payments.parquet.

## Для викладача
99_Instructor_only/V01_Instructor/ містить canonical truth, SQLite fallback, defect manifest, business exceptions та expected results. Цю папку не треба видавати слухачам.

## Важливо
Усі назви, організації, постачальники, документи та суми синтетичні. Набір не містить реальних оборонних або персональних даних.
