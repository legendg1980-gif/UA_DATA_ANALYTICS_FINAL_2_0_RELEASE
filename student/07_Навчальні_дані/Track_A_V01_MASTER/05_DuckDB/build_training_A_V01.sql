-- Run from 05_DuckDB folder: duckdb training_A_V01.duckdb < build_training_A_V01.sql

CREATE OR REPLACE TABLE stg_contracts AS SELECT * FROM read_csv_auto('../04_STAGING/stg_contracts.csv', header=true, sample_size=-1);
CREATE OR REPLACE TABLE stg_contract_lines AS SELECT * FROM read_csv_auto('../04_STAGING/stg_contract_lines.csv', header=true, sample_size=-1);
CREATE OR REPLACE TABLE stg_delivery_plan AS SELECT * FROM read_csv_auto('../04_STAGING/stg_delivery_plan.csv', header=true, sample_size=-1);
CREATE OR REPLACE TABLE stg_deliveries AS SELECT * FROM read_csv_auto('../04_STAGING/stg_deliveries.csv', header=true, sample_size=-1);
CREATE OR REPLACE TABLE stg_acceptance AS SELECT * FROM read_csv_auto('../04_STAGING/stg_acceptance.csv', header=true, sample_size=-1);
CREATE OR REPLACE TABLE stg_invoices AS SELECT * FROM read_csv_auto('../04_STAGING/stg_invoices.csv', header=true, sample_size=-1);
CREATE OR REPLACE TABLE stg_payments AS SELECT * FROM read_csv_auto('../04_STAGING/stg_payments.csv', header=true, sample_size=-1);
CREATE OR REPLACE TABLE stg_events AS SELECT * FROM read_csv_auto('../04_STAGING/stg_events.csv', header=true, sample_size=-1);
CREATE OR REPLACE VIEW vw_current_plan AS SELECT * EXCLUDE(rn) FROM (SELECT *, ROW_NUMBER() OVER(PARTITION BY ContractLineID ORDER BY VersionNo DESC) rn FROM stg_delivery_plan) WHERE rn=1;
CREATE OR REPLACE VIEW vw_payment_by_invoice AS SELECT InvoiceID, COUNT(*) PaymentCount, SUM(PaymentAmount) PaidAmount FROM stg_payments GROUP BY InvoiceID;
COPY (SELECT * FROM read_csv_auto('../06_Parquet/historical_payments.csv.gz', header=true)) TO '../06_Parquet/historical_payments.parquet' (FORMAT PARQUET);
