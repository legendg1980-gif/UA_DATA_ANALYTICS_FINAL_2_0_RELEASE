"""Add v2 (29-slide) theory/practice sections to the course program and instructor manual.

Idempotent: removes previously inserted v2 blocks (marked by heading text) before inserting.
Usage: python3 update_docs.py
"""
import os
import docx
import gen2

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROGRAM = os.path.join(ROOT, 'student/01_Програма/Програма_курсу_FINAL_2_0.docx')
MANUAL = os.path.join(ROOT, 'instructor/02_Instructor_Manual/Instructor_Manual_FINAL_2_0.docx')

P_HEAD = 'Теорія і практика заняття (слайди v2)'
M_HEAD = 'L. Слайди v2: теорія, вправи, відповіді'
M_INTRO = '0.2. Структура презентації v2 (29 слайдів)'

STRUCTURE = [
    'V01 обкладинка · V02 лектор · V03 план · V04 результат за програмою · V05 терміни заняття',
    'Частина 1 «Теорія» (V06–V12): розділювач, два теоретичні блоки, схема процесу, три ключові ідеї, фрагмент коду/таблиці, приклад на даних V01',
    'Частина 2 «Демо» (V13–V17): розділювач, 6 кроків демо, код демо з контрольним числом, 4 типові помилки, контрольні точки демо',
    'V18 перерва · Частина 3 «Практика» (V19–V24): 3 вправи з критерієм приймання, слайд відповідей, рівні Base / Core / Challenge',
    'V25 контрольні питання · V26 відповіді · V27 ДЗ / що далі · V28 література · V29 підсумок',
    'Кожен слайд має нотатки лектора з позначкою хвилин сценарію 150 хв: [хв 0–10], [хв 10–30], [хв 30–50], [хв 50–60], [хв 60–70], [хв 70–90], [хв 90–115], [хв 115–132], [хв 132–145], [хв 145–150].',
    'Слайди V23 (відповіді до вправ) і V26 (відповіді на контрольні питання) не показуються до розбору.',
]


def lines(s):
    return [x for x in (s or '').split('\n') if x.strip()]


def strip_v2(doc, marker):
    """Delete paragraphs from a marker heading until the next Heading 1/2 that is not part of the block."""
    body = doc.element.body
    ps = doc.paragraphs
    i = 0
    removed = 0
    while i < len(ps):
        p = ps[i]
        if p.style.name.startswith('Heading') and p.text.strip() in marker:
            lvl = p.style.name
            j = i + 1
            while j < len(ps) and not (ps[j].style.name.startswith('Heading')
                                        and ps[j].style.name <= lvl
                                        and ps[j].text.strip() not in marker):
                j += 1
            for q in ps[i:j]:
                body.remove(q._element)
                removed += 1
            ps = doc.paragraphs
            continue
        i += 1
    return removed


def lesson_heads(doc):
    """Return {n: index of 'Заняття n.' Heading 1} and list of all Heading 1 indices."""
    res = {}
    h1 = []
    for i, p in enumerate(doc.paragraphs):
        if p.style.name == 'Heading 1':
            h1.append(i)
            t = p.text.strip()
            if t.startswith('Заняття '):
                n = int(t.split()[1].rstrip('.'))
                res[n] = i
    return res, h1


def anchor_after_lesson(doc, n):
    """Paragraph before which the block for lesson n is inserted (next Heading 1), or None (= end)."""
    heads, h1 = lesson_heads(doc)
    start = heads[n]
    nxt = [i for i in h1 if i > start]
    return doc.paragraphs[nxt[0]] if nxt else None


def add(doc, anchor, text, style=None):
    if anchor is None:
        p = doc.add_paragraph(text)
        if style:
            p.style = style
        return p
    return anchor.insert_paragraph_before(text, style=style)


def bullets(doc, anchor, items, style='List Bullet'):
    st = style if style in [s.name for s in doc.styles] else None
    for it in items:
        add(doc, anchor, it if st else '• ' + it, st)


def program_block(doc, n):
    t = gen2.texts(n)
    a = anchor_after_lesson(doc, n)
    add(doc, a, P_HEAD, 'Heading 2')
    add(doc, a, 'Терміни заняття:')
    bullets(doc, a, lines(t['V05']['b']))
    add(doc, a, 'Теоретичні блоки:')
    theory = [t['V07']['t'], t['V09']['t'], t['V12']['t']]
    bullets(doc, a, theory)
    add(doc, a, 'Три ключові ідеї:')
    bullets(doc, a, [f"{t['V10'][f'H{k}']} — {t['V10'][f'X{k}']}" for k in (1, 2, 3)])
    add(doc, a, 'Демонстрація на наскрізному датасеті V01 — контрольні точки:')
    bullets(doc, a, lines(t['V17']['b']))
    add(doc, a, 'Практика (3 вправи з критерієм приймання):')
    for v in ('V20', 'V21', 'V22'):
        add(doc, a, t[v]['t'])
        bullets(doc, a, lines(t[v]['b']))
    add(doc, a, 'Рівні виконання: ' + '; '.join(
        f"{t['V24'][f'H{k}']}: {t['V24'][f'X{k}'].replace(chr(10), ' ')}" for k in (1, 2, 3)))
    add(doc, a, 'Контрольні питання для самоперевірки:')
    bullets(doc, a, lines(t['V25']['b']))


def manual_block(doc, n):
    t = gen2.texts(n)
    nt = gen2.notes(n)
    a = anchor_after_lesson(doc, n)
    add(doc, a, M_HEAD, 'Heading 2')
    add(doc, a, f"Приклад V01 (V12): {t['V12']['t']}")
    bullets(doc, a, lines(t['V12']['b']))
    add(doc, a, f"Демо (V14–V15): {t['V14']['t']}")
    bullets(doc, a, [f"{t['V14'][f'L{k}']}: {t['V14'][f'X{k}']}" for k in range(1, 7)])
    add(doc, a, 'Код / фрагмент демо (V15):')
    from docx.shared import Pt
    for ln in (t['V15']['b'] or '').split('\n'):
        par = add(doc, a, '')
        run = par.add_run(ln if ln.strip() else ' ')
        run.font.name = 'Consolas'
        run.font.size = Pt(9)
        par.paragraph_format.space_after = Pt(0)
    add(doc, a, 'Типові помилки (V16):')
    bullets(doc, a, [f"{t['V16'][f'H{k}']} — {t['V16'][f'X{k}']}" for k in range(1, 5)])
    add(doc, a, 'Контрольні точки демо (V17):')
    bullets(doc, a, lines(t['V17']['b']))
    for v in ('V20', 'V21', 'V22'):
        add(doc, a, f"{t[v]['t']} ({v})")
        bullets(doc, a, lines(t[v]['b']))
    add(doc, a, 'Відповіді до вправ (V23, показувати лише на розборі):')
    bullets(doc, a, lines(t['V23']['b']))
    add(doc, a, 'Контрольні питання і відповіді (V25–V26):')
    q = lines(t['V25']['b'])
    ans = [x.split('. ', 1)[1] if '. ' in x[:4] else x for x in lines(t['V26']['b'])]
    bullets(doc, a, [f"{qq} → {aa}" for qq, aa in zip(q, ans)])
    add(doc, a, 'Нотатки лектора до слайдів (дублюються в Speaker Notes презентації):')
    bullets(doc, a, [f"{k}: {v}" for k, v in sorted(nt.items()) if k in ('V07', 'V09', 'V12', 'V14', 'V15', 'V20', 'V23')])


def manual_intro(doc):
    ps = doc.paragraphs
    first_lesson = next(p for p in ps if p.style.name == 'Heading 1' and p.text.startswith('Заняття 1.'))
    add(doc, first_lesson, M_INTRO, 'Heading 2')
    bullets(doc, first_lesson, STRUCTURE)


def run():
    d = docx.Document(PROGRAM)
    r = strip_v2(d, {P_HEAD})
    for n in range(1, 16):
        program_block(d, n)
    d.save(PROGRAM)
    print('program: removed', r, 'paragraphs:', len(d.paragraphs))

    d = docx.Document(MANUAL)
    r = strip_v2(d, {M_HEAD, M_INTRO})
    manual_intro(d)
    for n in range(1, 16):
        manual_block(d, n)
    d.save(MANUAL)
    print('manual: removed', r, 'paragraphs:', len(d.paragraphs))


if __name__ == '__main__':
    run()
