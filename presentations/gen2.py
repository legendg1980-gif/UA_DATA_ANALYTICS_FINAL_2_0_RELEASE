# -*- coding: utf-8 -*-
"""v2 builder: rebuilds a v1 deck (S01..S16 slides) in place into 29 slides.

Usage:
  python gen2.py build N   -> batchUpdate requests for deck N (duplicate v1 slides, write text, delete v1)
  python gen2.py notes N notes_ids.json -> requests that write speaker notes
     (notes_ids.json: {"V01_pg": "<speakerNotesObjectId>", ...} read from the deck after build)
The rebuild keeps the same Drive file, so links in r_d stay valid.
"""
import json, sys
sys.path.insert(0, '.')
from content import L
from content_v2 import X
from manual_notes import M

# layout source slide in a v1 deck -> roles that hold text
SRC = {'cover': 'S01', 'copy': 'S02', 'plan': 'S03', 'result': 'S04', 'div': 'S05', 'bul': 'S06',
       'proc': 'S07', 'code': 'S09', 'three': 'S11', 'four': 'S12', 'cols': 'S13', 'quiz': 'S14'}
ROLES = {'cover': ['t', 'n', 'c'], 'copy': [], 'plan': ['b'], 'result': ['h', 'b'], 'div': ['h', 's'],
         'bul': ['t', 'b'], 'proc': ['t'] + [f'L{i}' for i in range(1, 7)] + [f'X{i}' for i in range(1, 7)],
         'code': ['t', 'b'], 'three': ['H1', 'H2', 'H3', 'X1', 'X2', 'X3'],
         'four': ['H1', 'H2', 'H3', 'H4', 'X1', 'X2', 'X3', 'X4'],
         'cols': ['t', 'H1', 'H2', 'H3', 'X1', 'X2', 'X3'], 'quiz': ['t', 'b']}
# (slide id, kind, field, scenario minute block)
SPEC = [
 ('V01', 'cover', None, '0–10 · Діагностика'), ('V02', 'copy', None, '0–10 · Діагностика'),
 ('V03', 'plan', None, '0–10 · Діагностика'), ('V04', 'result', None, '0–10 · Діагностика'),
 ('V05', 'bul', 'terms', '10–30 · Теорія з даних'), ('V06', 'div', 1, '10–30 · Теорія з даних'),
 ('V07', 'bul', 'c1', '10–30 · Теорія з даних'), ('V08', 'proc', 'proc', '10–30 · Теорія з даних'),
 ('V09', 'bul', 'c2', '10–30 · Теорія з даних'), ('V10', 'three', 'three', '10–30 · Теорія з даних'),
 ('V11', 'code', 'code', '10–30 · Теорія з даних'), ('V12', 'bul', 'c3', '10–30 · Теорія з даних'),
 ('V13', 'div', 2, '30–50 · V01 live demo'), ('V14', 'proc', 'demo', '30–50 · V01 live demo'),
 ('V15', 'code', 'dcode', '30–50 · V01 live demo'), ('V16', 'four', 'four', '50–60 · Типова помилка'),
 ('V17', 'bul', 'check', '50–60 · Типова помилка'), ('V18', 'div', 'break', '60–70 · Перерва'),
 ('V19', 'div', 3, '70–90 · Другий контекст і практика'), ('V20', 'bul', 'ex1', '90–115 · Guided practice'),
 ('V21', 'bul', 'ex2', '90–115 · Guided practice'), ('V22', 'bul', 'ex3', '115–132 · Challenge'),
 ('V23', 'bul', 'ans', '115–132 · Розбір вправ'), ('V24', 'cols', 'cols', '115–132 · Рівні Base / Core / Challenge'),
 ('V25', 'quiz', 'quiz', '132–145 · Розбір Q&A'), ('V26', 'bul', 'qa', '132–145 · Розбір Q&A'),
 ('V27', 'bul', 'hw', '145–150 · Exit'), ('V28', 'bul', 'res', '145–150 · Exit'),
 ('V29', 'bul', 'sum', '145–150 · Exit'),
]
OLD = [f'S{i:02d}_pg' for i in range(1, 17)]
# speaker-notes shape of each v1 source slide (identical in all 15 decks: copies of one skeleton)
NOTES = {'S01': '80', 'S02': '74', 'S03': '64', 'S04': '60', 'S05': '57', 'S06': '54', 'S07': '50',
         'S09': '38', 'S11': '32', 'S12': '24', 'S13': '18', 'S14': '10'}
NP = 'SLIDES_API1620973295_'
SIZE = {('cover', 't'): 36, ('cover', 'n'): 11, ('plan', 'b'): 16, ('result', 'b'): 18, ('div', 'h'): 60,
        ('div', 's'): 30, ('bul', 'b'): 15, ('proc', 'L'): 16, ('proc', 'X'): 15, ('code', 'b'): 13,
        ('three', 'H'): 30, ('three', 'X'): 14, ('four', 'H'): 19, ('four', 'X'): 12, ('cols', 'H'): 22,
        ('cols', 'X'): 13, ('cols', 't'): 30, ('quiz', 'b'): 15}
pt = lambda n: {'magnitude': n, 'unit': 'PT'}


def _lines(t, w):
    n, cur = 1, 0
    for word in t.split():
        add = len(word) + (1 if cur else 0)
        if cur + add > w: n, cur = n + 1, len(word)
        else: cur += add
    return n


def texts(n):
    """slide id -> {role: text}; field data from content.py (v1) + content_v2.py (v2)."""
    d = dict(L[n]); d.update(X[n]); out = {}
    for vid, kind, f, _ in SPEC:
        m = {}
        if kind == 'cover': m = {'t': d['title'], 'n': f'Заняття {n}'}
        elif kind == 'plan': m = {'b': '\n'.join(d['plan'])}
        elif kind == 'result': m = {'b': d['result']}
        elif kind == 'div':
            m = {1: {'h': 'ЧАСТИНА 1', 's': d['d1']}, 2: {'h': 'ЧАСТИНА 2', 's': d['dd']},
                 3: {'h': 'ЧАСТИНА 3', 's': 'Практика: 3 вправи'},
                 'break': {'h': 'ПЕРЕРВА', 's': '10 хв · потім GREEN / YELLOW / RED'}}[f]
        elif kind in ('bul', 'code', 'quiz'):
            t, b = d[f]; m = {'t': t, 'b': '\n'.join(b)}
        elif kind == 'proc':
            t, st = d[f]; m = {'t': t}
            for i, (a, b) in enumerate(st, 1): m[f'L{i}'] = a; m[f'X{i}'] = b
        elif kind == 'three':
            for i, (a, b) in enumerate(d[f], 1): m[f'H{i}'] = a; m[f'X{i}'] = b
        elif kind == 'four':
            for i, (a, b) in enumerate(d[f], 1): m[f'H{i}'] = a; m[f'X{i}'] = b
        elif kind == 'cols':
            t, cols = d[f]; m = {'t': t}
            for i, (a, ls) in enumerate(cols, 1): m[f'H{i}'] = a; m[f'X{i}'] = '\n'.join(ls)
        out[vid] = m
    return out


def size_of(kind, role, txt, n):
    if role == 't':
        if kind == 'cover':
            return 36 if _lines(txt, 27) <= 2 else (30 if _lines(txt, 33) <= 2 else 26)
        if kind == 'proc' and len(txt) <= 25: return 44
        if kind == 'cols': return 30 if len(txt) <= 36 else 26
        return 30 if len(txt) <= 36 else 26
    if kind == 'code' and role == 'b':
        return 12 if max(len(x) for x in txt.split('\n')) > 74 else 13
    if kind == 'three' and role[0] == 'H': return 30 if len(txt) <= 11 else 26
    if kind == 'four' and role[0] == 'H': return 19 if len(txt) <= 11 else 17
    return SIZE.get((kind, role), SIZE.get((kind, role[0])))


def src_sizes(n):
    """font sizes the v1 deck n has now (template defaults + v1 gen.fit_styles overrides)"""
    import gen
    cur = {'S01_t': 36, 'S01_n': 11, 'S03_b': 16, 'S04_b': 18, 'S05_h': 60, 'S05_s': 30, 'S06_b': 15, 'S07_t': 44,
           'S09_b': 13, 'S13_t': 30, 'S14_t': 30, 'S14_b': 15}
    for i in range(1, 7): cur[f'S07_L{i}'] = 16; cur[f'S07_X{i}'] = 15
    for i in range(1, 4): cur[f'S11_H{i}'] = 30; cur[f'S11_X{i}'] = 14; cur[f'S13_H{i}'] = 22; cur[f'S13_X{i}'] = 13
    for i in range(1, 5): cur[f'S12_H{i}'] = 19; cur[f'S12_X{i}'] = 12
    for r in gen.fit_styles(n): u = r['updateTextStyle']; cur[u['objectId']] = u['style']['fontSize']['magnitude']
    return cur


def build(n, part=None):
    """part 'a' = slides V01-V15, 'b' = V16-V29 + delete v1 + reorder; None = all in one batch"""
    T = texts(n); NT = notes(n); cur = src_sizes(n); req = []
    spec = {'a': SPEC[:15], 'b': SPEC[15:], None: SPEC}[part]
    for vid, kind, f, _ in spec:
        src = SRC[kind]
        ids = {src + '_pg': vid + '_pg', NP + NOTES[src]: vid + '_nt'}
        for r in ROLES[kind]: ids[f'{src}_{r}'] = f'{vid}_{r}'
        req.append({'duplicateObject': {'objectId': src + '_pg', 'objectIds': ids}})
        req.append({'insertText': {'objectId': vid + '_nt', 'text': NT[vid]}})
        for r, txt in T[vid].items():
            oid = f'{vid}_{r}'
            req.append({'deleteText': {'objectId': oid, 'textRange': {'type': 'ALL'}}})
            req.append({'insertText': {'objectId': oid, 'text': txt}})
            s = size_of(kind, r, txt, n)
            if s and s != cur.get(f'{src}_{r}'): req.append({'updateTextStyle': {'objectId': oid, 'textRange': {'type': 'ALL'},
                                                  'style': {'fontSize': pt(s)}, 'fields': 'fontSize'}})
        if kind in ('bul', 'plan', 'quiz'):
            req.append({'createParagraphBullets': {'objectId': f'{vid}_b', 'textRange': {'type': 'ALL'},
                        'bulletPreset': 'NUMBERED_DIGIT_ALPHA_ROMAN' if kind == 'quiz' else 'BULLET_DISC_CIRCLE_SQUARE'}})
    if part == 'a': return req
    for s in OLD: req.append({'deleteObject': {'objectId': s}})
    for i, (vid, *_r) in enumerate(SPEC):
        req.append({'updateSlidesPosition': {'slideObjectIds': [vid + '_pg'], 'insertionIndex': i}})
    return req


def notes(n):
    d = dict(L[n]); d.update(X[n]); man = M[n]; hn = d.get('notes', {}); out = {}
    q, a = d['quiz'][1], d['qa'][1]
    auto = {
     'V01': 'Відкрийте за 15 хв до дзвінка: ' + '; '.join(man['C']) + '. Instructor-only матеріали під час показу екрана не відкривати. '
            'Діагностичне питання до групи: «Яке бізнес-питання можна перевірити цим файлом і що в ньому є одним рядком?»',
     'V02': 'Замініть шаблонний текст: ім\'я, прізвище, досвід лектора, 2–3 речення про роботу з даними.',
     'V03': 'Озвучте план і поясніть, що кожен блок закінчується контрольним числом. Наприкінці кожного блоку — GREEN / YELLOW / RED у чаті.',
     'V04': 'Результат заняття за програмою. Поверніться до цього слайда на 145-й хвилині: що з цього вже готово в кожного.',
     'V05': 'Не читайте терміни списком: кожен термін одразу прив\'язуйте до конкретного рядка або контрольного числа V01. '
            'Перевірка: студент пояснює перші два терміни своїми словами.',
     'V06': 'Теорія — 20 хвилин. Кожне твердження підкріплюйте рядком або числом з V01.',
     'V08': 'Що малюємо на порожньому слайді PowerPoint: ' + man['E'],
     'V10': 'Три ідеї, які група має забрати із заняття. Попросіть одного студента переказати кожну своїми словами.',
     'V11': 'Покажіть фрагмент на екрані, а контрольне число попросіть передбачити до запуску.',
     'V13': 'Live demo: перед кожною дією група прогнозує результат у чаті (рядки, унікальні ключі або сума), потім виконуємо і звіряємо.',
     'V17': 'Контрольні точки студенти записують у свій контрольний аркуш. Після слайда — GREEN / YELLOW / RED у чаті.',
     'V18': '10 хвилин перерви. Після повернення — GREEN / YELLOW / RED у чаті за останнім checkpoint. '
            '≤15% RED — далі; 15–30% — один спільний fix; >30% — повернення до checkpoint.',
     'V19': 'На екрані лишається лише умова й критерій приймання. Не копіюйте кліки викладача — перевіряйте себе контрольним числом.',
     'V21': 'Через 10 хв — GREEN / YELLOW / RED. Слайд із відповідями не показуйте до розбору.',
     'V22': 'Challenge: спершу студент описує метод, потім виконує. Обведіть на схемі місце, де найімовірніша помилка.',
     'V23': 'Розбір: спершу 1–2 студенти озвучують свої числа, потім відкрийте відповіді. Правильна відповідь — не лише число, а правило і grain.',
     'V24': 'Base — повторити показаний приклад V01 і отримати контрольний результат. Core — аналогічний кейс на іншому файлі або другому '
            'контексті з evidence. Challenge — запропонувати додатковий контроль і пояснити, яку помилку він ловить.',
     'V25': 'Відповіді в чат протягом 3 хв. Очікувані відповіді: ' + ' | '.join(f'{i}) {x} → {y.split(". ",1)[-1]}' for i, (x, y) in enumerate(zip(q, a), 1)),
     'V26': 'Ціль — не менше 70% коректних відповідей після розбору. Якщо група відстає: ' + man['I'],
     'V27': man['J'] or 'Окремого ДЗ немає.',
     'V28': 'Література після заняття: ' + '; '.join(man['K']),
     'V29': 'Exit: зафіксуйте, що має бути готово до наступного заняття. Виділіть рамкою результат, що стає входом наступного заняття.',
    }
    keymap = {'V07': 'c1', 'V09': 'c2', 'V12': 'c3', 'V14': 'demo', 'V15': 'dcode', 'V16': 'four', 'V20': 'ex'}
    for vid, kind, f, block in SPEC:
        body = hn.get(keymap.get(vid, ''), '') or auto.get(vid, '')
        if vid == 'V16' and hn.get('four'): body = hn['four']
        out[vid] = f'[хв {block}] ' + body
    return out


def notes_requests(n, ids):
    return [{'insertText': {'objectId': ids[v + '_pg'], 'text': t, 'insertionIndex': 0}} for v, t in notes(n).items()]


if __name__ == '__main__':
    cmd, n = sys.argv[1], int(sys.argv[2])
    if cmd == 'build': print(json.dumps(build(n, sys.argv[3] if len(sys.argv) > 3 else None), ensure_ascii=False, separators=(',', ':')))
    elif cmd == 'notes': print(json.dumps(notes_requests(n, json.load(open(sys.argv[3]))), ensure_ascii=False, separators=(',', ':')))
    elif cmd == 'show':
        for k, v in texts(n).items(): print(k, v)
        for k, v in notes(n).items(): print(k, v)
