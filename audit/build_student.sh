set -e
SRC=$1; OUT=$2; rm -rf $OUT; mkdir -p $OUT; S=$OUT/UA_DATA_ANALYTICS_FINAL_2_1_STUDENT; cp -r $SRC $S; cd $S
rm -rf 99_Instructor_only 02_Instructor_Manual 06_REFERENCE_REPORT/Control_Results_V01.xlsx "09_Додаткові_матеріали/MASTER_v9" "09_Додаткові_матеріали/MASTER_14_HW" "09_Додаткові_матеріали/MASTER_REFERENCE" 07_Навчальні_дані/Track_A_V01_MASTER/Generator 07_Навчальні_дані/Track_A_V01_MASTER/00_ReadMe/Validation_Report.txt PUBLISHING_QA_FINAL_2_0.txt SHA256_MANIFEST.txt README_FINAL_2_0.txt CHANGELOG_2_1.txt
mkdir -p $OUT/RELEASE_2_after_HW04/Track_A_V01_MASTER; mv 07_Навчальні_дані/Track_A_V01_MASTER/04_STAGING $OUT/RELEASE_2_after_HW04/Track_A_V01_MASTER/
python3 -I -W ignore - "$S" <<'PY'
import sys,re,openpyxl
S=sys.argv[1]
p=S+'/07_Навчальні_дані/Track_A_V01_MASTER/00_ReadMe/README_UA.md'; t=open(p,encoding='utf-8').read()
t=re.sub(r'## Golden aggregates.*?(?=## Формати)','',t,flags=re.S)
t=re.sub(r'- Accepted after base QA.*?\n- Перевірка.*?\n','- Accepted / Rejected: визначаєте самостійно (контроль: Input = Accepted + Rejected)\n',t,flags=re.S)
t=re.sub(r'## Для викладача.*?(?=## Важливо)','## Staging\nПапка 04_STAGING видається окремо після перевірки HW04 (вона містить очищені таблиці).\n\n## 13-й файл\n01_Source/Track_A/ADD_AFTER_DEMO/deliveries_2027_01.csv копіюється в Deliveries лише за командою викладача (HW03, захист).\n\n',t,flags=re.S)
open(p,'w',encoding='utf-8').write(t)
f=S+'/07_Навчальні_дані/Track_B_SUPPLEMENT/balance_snapshot_2026_12_31.xlsx'; wb=openpyxl.load_workbook(f); ws=wb.active
for name in ['CaseID','Status','CalculatedClosing','DaysOfSupply']:
    h=[c.value for c in ws[1]]
    if name in h: ws.delete_cols(h.index(name)+1)
wb.save(f)
PY
cat > README_STUDENT.txt <<'R'
UA DATA ANALYTICS FINAL 2.1 — СТУДЕНТСЬКИЙ ПАКЕТ (Release 1)

00_Навігатор/           — програма вимог і навігатор
01_Програма/            — 15 занять
03_Домашні_завдання/    — HW01–HW08 (core)
04_Курсовий_проєкт/     — накопичувальний проєкт, захист на занятті 15
05_Оцінювання/          — рубрика
06_REFERENCE_REPORT/    — специфікація фінального звіту
07_Навчальні_дані/      — Track A V01 (сирі джерела) і Track B (облік/рух майна)
08_Шаблони/             — шаблони ДЗ і чек-листи
09_Додаткові_матеріали/ — optional-модулі

Release 2 (04_STAGING — очищені таблиці для SQL/DAX) видається після перевірки HW04.
Файли ADD_AFTER_DEMO (Track A і Track B) додаються в папку джерел лише за командою викладача.
Інструкторські матеріали, ключі та еталонні таблиці в цей пакет не входять.
Посилання в документах на Instructor Manual чи Control_Results стосуються викладацького пакета.
Усі дані синтетичні. Перевірка цілісності: SHA256_MANIFEST.txt.
R
cd $OUT
for d in UA_DATA_ANALYTICS_FINAL_2_1_STUDENT RELEASE_2_after_HW04; do (cd $d && find . -type f ! -name SHA256_MANIFEST.txt -print0 | sort -z | xargs -0 sha256sum | sed 's#  \./#  #' > SHA256_MANIFEST.txt && sha256sum -c --quiet SHA256_MANIFEST.txt); done
zip -qr -X UA_DATA_ANALYTICS_FINAL_2_1_STUDENT_R1.zip UA_DATA_ANALYTICS_FINAL_2_1_STUDENT; (cd RELEASE_2_after_HW04 && zip -qr -X ../UA_DATA_ANALYTICS_FINAL_2_1_STUDENT_R2_STAGING.zip .)
