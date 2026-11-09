@echo off
cd /d "%~dp0"
duckdb training_A_V01.duckdb < build_training_A_V01.sql
echo DuckDB created: training_A_V01.duckdb
pause
