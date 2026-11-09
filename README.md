# UA DATA ANALYTICS IN THE DEFENSE BUSINESS — FINAL 2.0

Курс: 15 занять, 8 наскрізних ДЗ, курсовий проєкт і захист на 15-му занятті. Усі дані синтетичні.

| Папка | Що всередині | Кому |
| --- | --- | --- |
| `student/` | Release 1: програма, ДЗ HW01–HW08, курсовий, рубрика, шаблони, сирі дані Track A/Track B | студентам |
| `student_release2_after_HW04/` | Release 2: `04_STAGING` (очищені таблиці для SQL/DAX) | студентам після перевірки HW04 |
| `instructor/` | Instructor Manual, ключі, еталонні таблиці, маніфести дефектів, Control_Results, MASTER v9, генератор, оригінальні README/QA/маніфест | **лише викладачу** |
| `audit/` | Аудит пакета: розв'язання HW01–HW08 (`Student_Solutions_HW01_HW08.xlsx`) і скрипт `solve.py` | викладачу |

Відомі проблеми та пріоритети виправлень див. у звіті аудиту (Claude Docs) і на аркуші `00_Comparison` у `audit/`.

> Репозиторій має бути **приватним**: `instructor/` містить відповіді.

Запуск аудиту: `python3 audit/solve.py <шлях до повного пакета> out.xlsx` (потрібні pandas, openpyxl). Скрипт очікує оригінальну структуру пакета FINAL 2.0.
