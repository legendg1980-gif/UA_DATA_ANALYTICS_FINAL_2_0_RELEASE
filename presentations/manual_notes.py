# Auto-extracted from instructor/02_Instructor_Manual/Instructor_Manual_FINAL_2_0.docx (sections C, E, I, J, K)
M = {
 "1": {
  "C": [
   "07_Навчальні_дані/Track_A_V01_MASTER/02_Reference/contracts_sample.xlsx",
   "07_Навчальні_дані/Track_A_V01_MASTER/02_Reference/contract_lines_sample.xlsx",
   "07_Навчальні_дані/Track_A_V01_MASTER/02_Reference/payments_sample.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Common/source_priority.xlsx"
  ],
  "E": "Contract → Plan → Delivery → Acceptance → Invoice → Payment; під кожним блоком дописати grain. Праворуч — логістичний аналог Need → Receipt → Issue → Balance.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Окреме ДЗ цього заняття входить у наступну наскрізну роботу.",
  "K": [
   "Microsoft Learn — Power BI star schema guidance; Microsoft Support — Числа, відформатовані як текст: як знайти й виправити",
   "Kimball Group — Dimensional Modeling Techniques; NIST SP 800-188 — De-Identifying Government Datasets"
  ]
 },
 "2": {
  "C": [
   "07_Навчальні_дані/Exercises/EX02_V01_dirty_register_800.xlsx",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Contracts/contracts_2026_H1.xlsx",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Contracts/contracts_2026_H2.xlsx"
  ],
  "E": "RAW rows/total → Excel Table → formulas → ErrorLog → Clean rows/total; Delta має бути пояснена.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Після цього заняття видається HW01 з папки 03_Домашні_завдання. Умова вже об’єднує відповідні матеріали MASTER, тому старі 14 HW не видаються як окремий core-потік.",
  "K": [
   "Microsoft Support — Excel tables",
   "Microsoft Support — XLOOKUP"
  ]
 },
 "3": {
  "C": [
   "07_Навчальні_дані/Track_A_V01_MASTER/09_QA/DQ_RuleCatalog.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Contracts/contracts_2026_H1.xlsx",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Contracts/contracts_2026_H2.xlsx",
   "99_Instructor_only/V01_Instructor/Manifests/defect_manifest_A_V01.csv"
  ],
  "E": "6 вимірів DQ: Completeness / Validity / Uniqueness / Consistency / Timeliness / Balance; праворуч класи Critical / Material / Cosmetic / Business Exception.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Після цього заняття видається HW02 з папки 03_Домашні_завдання. Умова вже об’єднує відповідні матеріали MASTER, тому старі 14 HW не видаються як окремий core-потік.",
  "K": [
   "Microsoft Learn — Data profiling tools",
   "Great Expectations — Data quality"
  ]
 },
 "4": {
  "C": [
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Deliveries/deliveries_2026_01.csv … deliveries_2026_12.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Contracts/contracts_2026_H1.xlsx",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Contracts/contracts_2026_H2.xlsx",
   "07_Навчальні_дані/Track_B_SUPPLEMENT/issues_2026_01.csv … issues_2026_12.csv"
  ],
  "E": "Folder → File list → Sample transform → Combine → Explicit types → SourceFile → Rejected files → Control totals. Окремо намалювати Schema drift.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Після цього заняття видається HW03 з папки 03_Домашні_завдання. Умова вже об’єднує відповідні матеріали MASTER, тому старі 14 HW не видаються як окремий core-потік.",
  "K": [
   "Microsoft Learn — Folder connector",
   "Microsoft Learn — Combine files overview"
  ]
 },
 "5": {
  "C": [
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Suppliers/suppliers_master.xlsx",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Suppliers/supplier_updates.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Common/units_of_measure.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Common/category_mapping.xlsx",
   "07_Навчальні_дані/Track_B_SUPPLEMENT/uom_mapping.csv"
  ],
  "E": "Source value → Trim/Clean/Locale/Key normalization → Mapping table → Canonical value; окремо null ≠ 0.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Окреме ДЗ цього заняття входить у наступну наскрізну роботу.",
  "K": [
   "Microsoft Learn — Power Query data types",
   "Microsoft Learn — Replace values"
  ]
 },
 "6": {
  "C": [
   "07_Навчальні_дані/Track_A_V01_MASTER/02_Reference/contracts_sample.xlsx",
   "07_Навчальні_дані/Track_A_V01_MASTER/02_Reference/payments_sample.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Payments/payments_2026_12.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/PDF/acceptance_act_00041.pdf"
  ],
  "E": "Append vs Merge; Left/Inner/Left Anti/Right Anti; поруч 1 contract × N lines × M payments = multiplication risk.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Після цього заняття видається HW04 з папки 03_Домашні_завдання. Умова вже об’єднує відповідні матеріали MASTER, тому старі 14 HW не видаються як окремий core-потік.",
  "K": [
   "Microsoft Learn — Merge queries overview",
   "Microsoft Learn — Append queries"
  ]
 },
 "7": {
  "C": [
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Plan/delivery_plan_v1.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Plan/delivery_plan_v2.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Plan/delivery_plan_v3.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_delivery_plan.csv",
   "07_Навчальні_дані/Track_B_SUPPLEMENT/need_plan_v1.csv",
   "07_Навчальні_дані/Track_B_SUPPLEMENT/need_plan_v2.csv"
  ],
  "E": "Timeline V1 → V2 → V3 із вертикальною As Of Date; нижче три дельти: Quantity / Amount / Days.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Окреме ДЗ цього заняття входить у наступну наскрізну роботу.",
  "K": [
   "Microsoft Learn — Date tables in Power BI",
   "Microsoft Learn — Star schema guidance"
  ]
 },
 "8": {
  "C": [
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_contracts.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_contract_lines.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_deliveries.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_payments.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Common/calendar.xlsx"
  ],
  "E": "Неправильне fact-to-fact павутиння перекреслити; побудувати star: DimDate/Contract/Supplier/Item → Facts. Підписати grain кожної таблиці.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Після цього заняття видається HW05 з папки 03_Домашні_завдання. Умова вже об’єднує відповідні матеріали MASTER, тому старі 14 HW не видаються як окремий core-потік.",
  "K": [
   "Microsoft Learn — Star schema guidance",
   "Microsoft Learn — Active/inactive relationships"
  ]
 },
 "9": {
  "C": [
   "07_Навчальні_дані/Track_A_V01_MASTER/05_DuckDB/build_training_A_V01.sql",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_contracts.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_payments.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_deliveries.csv"
  ],
  "E": "SELECT → FROM/JOIN → WHERE → GROUP BY → HAVING → ORDER BY; поруч SQL control = Power Query control.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Окреме ДЗ цього заняття входить у наступну наскрізну роботу.",
  "K": [
   "DuckDB — SQL introduction",
   "DuckDB — GROUP BY"
  ]
 },
 "10": {
  "C": [
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_contracts.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_deliveries.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_payments.csv",
   "06_REFERENCE_REPORT/Control_Results_V01.xlsx"
  ],
  "E": "Measure = Formula + Filter Context; три slicers змінюють один KPI; поряд KPI passport: definition/source/grain/date/control.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Після цього заняття видається HW06 з папки 03_Домашні_завдання. Умова вже об’єднує відповідні матеріали MASTER, тому старі 14 HW не видаються як окремий core-потік.",
  "K": [
   "Microsoft Learn — DAX overview",
   "Microsoft Learn — CALCULATE"
  ]
 },
 "11": {
  "C": [
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_invoices.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_payments.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Common/calendar.xlsx",
   "07_Навчальні_дані/Track_B_SUPPLEMENT/balance_snapshot_2026_12_31.xlsx",
   "07_Навчальні_дані/Track_B_SUPPLEMENT/norms.csv"
  ],
  "E": "Past → Reporting As Of Date → Future; aging buckets і три сценарії Optimistic/Base/Pessimistic; Days of Supply = Balance / Daily rate.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Окреме ДЗ цього заняття входить у наступну наскрізну роботу.",
  "K": [
   "Microsoft Learn — Time intelligence functions",
   "Storytelling with Data"
  ]
 },
 "12": {
  "C": [
   "06_REFERENCE_REPORT/Final_Report_Specification_2_0.docx",
   "06_REFERENCE_REPORT/Control_Results_V01.xlsx",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_deliveries.csv",
   "07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING/stg_payments.csv"
  ],
  "E": "Executive → Analytical/Reconciliation → Profile/Exceptions; від KPI стрілка Drill-through → Detail/Source. QA controls вбудовані у Analysis/Profile.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Після цього заняття видається HW07 з папки 03_Домашні_завдання. Умова вже об’єднує відповідні матеріали MASTER, тому старі 14 HW не видаються як окремий core-потік.",
  "K": [
   "Microsoft Learn — Drillthrough",
   "Microsoft Learn — Visual interactions"
  ]
 },
 "13": {
  "C": [
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Contracts/contracts_2026_H1.xlsx",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Invoices/invoices_Q1.xlsx",
   "07_Навчальні_дані/Track_A_V01_MASTER/01_Source/Track_A/Payments/payments_2026_12.csv",
   "07_Навчальні_дані/Track_B_SUPPLEMENT/need_plan_v2.csv",
   "07_Навчальні_дані/Track_B_SUPPLEMENT/balance_snapshot_2026_12_31.xlsx"
  ],
  "E": "Дві профільні гілки: Contract → Accepted → Invoice → Paid → Outstanding; Need → Balance → Days of Supply. Унизу Exception Register: rule/source/threshold/as-of/owner.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Окреме ДЗ цього заняття входить у наступну наскрізну роботу.",
  "K": [
   "Microsoft Learn — Conditional formatting",
   "Storytelling with Data — resources"
  ]
 },
 "14": {
  "C": [
   "08_Шаблони/Student_Templates.xlsx",
   "08_Шаблони/Analytical_Note_Template.docx",
   "08_Шаблони/Safe_Sharing_Checklist.docx",
   "06_REFERENCE_REPORT/Control_Results_V01.xlsx"
  ],
  "E": "Source → Transform → Model → Measure → Visual → Note; біля кожної стрілки rows / unique keys / total / test case. Праворуч Fact → Cause → Risk → Proposal → Limitations.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Після цього заняття видається HW08 з папки 03_Домашні_завдання. Умова вже об’єднує відповідні матеріали MASTER, тому старі 14 HW не видаються як окремий core-потік.",
  "K": [
   "Microsoft Learn — Query diagnostics",
   "Microsoft Learn — Power BI data protection planning"
  ]
 },
 "15": {
  "C": [
   "04_Курсовий_проєкт/Курсовий_проєкт_FINAL_2_0.docx",
   "05_Оцінювання/Рубрика_оцінювання_FINAL_2_0.docx",
   "Студентський PBIX/SQL/PQ/Note"
  ],
  "E": "Refresh → Model → 3 KPI → Exception → Drill-through → Analytical note → Limits; поряд рубрика 20/15/20/15/10/10/10.",
  "I": "Не скорочуйте контрольний принцип. Скоротіть кількість rows/джерел до мінімального прикладу, поверніться до останнього PASS checkpoint, після чого дайте recovery-файл/інструкцію на платформі. Індивідуальні локальні проблеми не забирають ефір великої групи.",
  "J": "Це фінальний захист курсового проєкту. Окремого core-іспиту немає.",
  "K": [
   "Microsoft Learn — Power BI adoption roadmap",
   "Kimball Group — Techniques"
  ]
 }
}
M = {int(k): v for k, v in M.items()}
