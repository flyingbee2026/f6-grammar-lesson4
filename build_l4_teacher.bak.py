# -*- coding: utf-8 -*-
"""Build 6B Grammar Lesson 4 - Giving Examples (Teacher's) by editing a copy of
Helen's finalised Lesson 3 teacher's file (same styles, shading, fonts, layout).
Run: python3 build_l4_teacher.py
"""
from copy import deepcopy
import docx
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

SRC = '/root/6b-gram-lesson4/L3_Teachers.docx'
OUT = "/root/6b-gram-lesson4/6B Grammar Lesson 4 - Giving Examples (Teacher's).docx"
RED = 'C00000'

doc = docx.Document(SRC)
body = doc.element.body


# ----------------------------------------------------------------- helpers
def kill_runs(par):
    for r in list(par.runs):
        r._r.getparent().remove(r._r)


def write_runs(par, specs, tmpl_el):
    """specs: list of dicts t/b/i/u/c ; tmpl_el: an <w:r> element to inherit rPr from."""
    kill_runs(par)
    for sp in specs:
        new = deepcopy(tmpl_el)
        for t in new.findall(qn('w:t')):
            new.remove(t)
        par._p.append(new)
        run = par.runs[-1]
        run.text = sp['t']
        run.font.bold = sp.get('b')
        run.font.italic = sp.get('i')
        run.font.underline = sp.get('u')
        rPr = run._r.get_or_add_rPr()
        for c in rPr.findall(qn('w:color')):
            rPr.remove(c)
        if sp.get('c'):
            run.font.color.rgb = RGBColor.from_string(sp['c'])


def fill_cell(cell, specs_list, tmpl_el):
    """specs_list: list of paragraphs, each a list of run specs."""
    while len(cell.paragraphs) < len(specs_list):
        cell._tc.append(deepcopy(cell.paragraphs[-1]._p))
    for i, specs in enumerate(specs_list):
        write_runs(cell.paragraphs[i], specs, tmpl_el)
    for extra in cell.paragraphs[len(specs_list):]:
        extra._p.getparent().remove(extra._p)


def cant_split(tbl, all_rows=True):
    """Stop a table's rows breaking across pages."""
    for row in tbl.rows:
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            el = trPr.makeelement(qn('w:cantSplit'), {})
            trPr.insert(0, el)


# ------------------------------------------------------------- index body
idx = []
for i, child in enumerate(body.iterchildren()):
    if child.tag == qn('w:p'):
        idx.append(('p', i, Paragraph(child, doc)))
    elif child.tag == qn('w:tbl'):
        idx.append(('tbl', i, Table(child, doc)))


def find_p(prefix, nth=0):
    hits = [o for k, i, o in idx if k == 'p' and o.text.strip().startswith(prefix)]
    return hits[nth]


def find_all_p(prefix):
    return [o for k, i, o in idx if k == 'p' and o.text.strip().startswith(prefix)]


def find_table(pred):
    return [o for k, i, o in idx if k == 'tbl' and pred(o)]


# ------------------------------------------------------------- locate parts
title_p = find_p('Grammar & Writing')
partA_box = [t for t in find_table(lambda t: t.rows[0].cells[0].text.strip().startswith('1. In recent'))][0]
partB_head = find_p('PART B:')
vocab = [t for t in find_table(lambda t: [c.text.strip() for c in t.rows[0].cells] == ['#', 'Basic', 'Upgraded'])][0]
basic_box = [t for t in find_table(lambda t: t.rows[0].cells[0].text.strip().startswith('THE BASIC SENTENCES'))][0]
pat_tbls = find_table(lambda t: t.rows[0].cells[0].text.strip().startswith('Pattern '))
assert len(pat_tbls) == 3, len(pat_tbls)
teach_notes = find_all_p('Teaching note:')
design_note = find_p('Teacher\u2019s note on the design:')
partD_intro = find_p('Upgrade each simple sentence')
set_heads = find_all_p('Set ')
set_tbls = find_table(lambda t: 'Use Pattern' in ''.join(c.text for r in t.rows for c in r.cells))
assert len(set_heads) == 6 and len(set_tbls) == 6, (len(set_heads), len(set_tbls))
key_notes = find_all_p('Key note:')
assert len(key_notes) == 6
prompt_box = [t for t in find_table(lambda t: 'A recent survey has found that young people in Hong Kong' in t.rows[0].cells[0].text)][0]
req_head = find_p('Task requirements:')
req1 = find_p('Make ONE suggestion')
chain = find_p('Follow the chain')
models_head = find_p('Teacher\u2019s three suggested paragraphs')
model_heads = find_all_p('Suggested paragraph')
final_note = find_p('Teacher\u2019s note: each paragraph')
assert len(model_heads) == 3
model_bodies = [Paragraph(h._p.getnext(), doc) for h in model_heads]

# templates for new runs (all from Helen's own file)
BLACK = find_p('Upgrade each simple sentence').runs[0]._r
RED_T = teach_notes[0].runs[1]._r           # red, not bold
RED_B = teach_notes[0].runs[0]._r           # red, bold

# =============================================================== 1. title
write_runs(title_p, [{'t': 'Grammar & Writing \u2013 Lesson 4: Giving Examples'}], title_p.runs[0]._r)

# =========================================================== 2. PART A
A_SENTENCES = [
    'There is an urgent need for the government to investing in free sports facilities, as young people is losing interest in outdoor activities.',
    'To improving employees\u2019 mental health, the company should make it a priority to hold a short team meeting every Monday, which in turn help staff to share their workload.',
    'Managers, in collaboration of the human resources team, should encourage their staff taking a proper lunch break away from the desk.',
    'There is an urgent need for the government to raise graduate awareness of the skills that employers need, because many fresh graduates lacks practical experience.',
    'To cut company cost, managers should make it a priority to reduce overtime hours and allowing staff to work from home twice a week.',
    'Admittedly, the elderlies in our community benefit from regular visits; nevertheless, most employees rarely volunteers their free time.',
]
cellA = partA_box.rows[0].cells[0]
for i, s in enumerate(A_SENTENCES):
    par = cellA.paragraphs[i]
    tmpl = par.runs[1]._r if len(par.runs) > 1 else par.runs[0]._r
    write_runs(par, [{'t': '%d. ' % (i + 1), 'b': True}, {'t': s}], tmpl)

A_KEYS = [
    ('There is an urgent need for the government to invest in free sports facilities, as young people are losing interest in outdoor activities.',
     '(a) Verb form: \u201cto investing\u201d \u2192 \u201cto invest\u201d \u2014 Lesson 3, Pattern 1 takes TO + the base verb (\u201cfor the government to invest\u201d).',
     '(b) Agreement: \u201cyoung people is\u201d \u2192 \u201cyoung people are\u201d \u2014 \u201cpeople\u201d is plural.'),
    ('To improve employees\u2019 mental health, the company should make it a priority to hold a short team meeting every Monday, which in turn helps staff to share their workload.',
     '(a) Verb form: \u201cTo improving\u201d \u2192 \u201cTo improve\u201d \u2014 Lesson 3, Pattern 2 opens with TO + the base verb to state the goal.',
     '(b) Agreement: \u201cwhich in turn help\u201d \u2192 \u201chelps\u201d \u2014 the relative clause takes a singular verb (the meeting helps).'),
    ('Managers, in collaboration with the human resources team, should encourage their staff to take a proper lunch break away from the desk.',
     '(a) Preposition: \u201cin collaboration of\u201d \u2192 \u201cin collaboration with\u201d \u2014 Lesson 3, Pattern 3 is a fixed phrase.',
     '(b) Verb pattern: \u201cencourage their staff taking\u201d \u2192 \u201cencourage their staff to take\u201d \u2014 encourage + object + TO + base verb.'),
    ('There is an urgent need for the government to raise graduates\u2019 awareness of the skills that employers need, because many fresh graduates lack practical experience.',
     '(a) Possessive form: \u201cgraduate awareness\u201d \u2192 \u201cgraduates\u2019 awareness\u201d (the awareness belongs to the graduates).',
     '(b) Agreement: \u201cmany fresh graduates lacks\u201d \u2192 \u201clack\u201d \u2014 a plural subject takes the base verb.'),
    ('To cut company costs, managers should make it a priority to reduce overtime hours and allow staff to work from home twice a week.',
     '(a) Noun number: \u201ccompany cost\u201d \u2192 \u201ccompany costs\u201d \u2014 costs are countable here.',
     '(b) Parallel form: \u201cand allowing\u201d \u2192 \u201cand allow\u201d \u2014 the two verbs after \u201cto\u201d must match (\u201cto reduce \u2026 and allow \u2026\u201d).'),
    ('Admittedly, the elderly in our community benefit from regular visits; nevertheless, most employees rarely volunteer their free time.',
     '(a) Word form: \u201cthe elderlies\u201d \u2192 \u201cthe elderly\u201d \u2014 \u201cthe elderly\u201d is already plural (Lesson 3, Set 6).',
     '(b) Agreement: \u201cmost employees rarely volunteers\u201d \u2192 \u201cvolunteer\u201d \u2014 plural subject, base verb.'),
]

key_tbl = deepcopy(vocab._tbl)          # 7 x 3, same grid as the sentence table
st = key_tbl.find(qn('w:tblPr')).find(qn('w:tblStyle'))
st.set(qn('w:val'), 'aff2')
key_tbl_el = key_tbl
for j, h in enumerate(['Item', 'Corrected sentence', 'The TWO errors and why']):
    write_runs(Table(key_tbl_el, doc).rows[0].cells[j].paragraphs[0], [{'t': h, 'b': True}], BLACK)
KT = Table(key_tbl_el, doc)
for r, (corr, e1, e2) in enumerate(A_KEYS, start=1):
    fill_cell(KT.rows[r].cells[0], [[{'t': str(r), 'b': True, 'c': RED}]], RED_B)
    fill_cell(KT.rows[r].cells[1], [[{'t': corr, 'c': RED}], []], RED_T)
    fill_cell(KT.rows[r].cells[2], [[{'t': e1, 'c': RED}], [{'t': e2, 'c': RED}]], RED_T)

head_key = deepcopy(find_p('PART A:')._p)
write_runs(Paragraph(head_key, doc), [{'t': 'TEACHER\u2019S ANSWERS (LESSON 4, PART A)'}],
           find_p('PART A:').runs[0]._r)
partA_box._tbl.addnext(head_key)
head_key.addnext(key_tbl_el)

# =========================================================== 3. PART B
# (Part A instruction line)
instr = find_p('Proofread the following sentences')
write_runs(instr, [{'t': 'Proofread the following sentences. Each sentence contains TWO errors. '
                         'Most of the mistakes come from Lesson 3\u2019s sentence patterns.'}], instr.runs[0]._r)

partB_topic = deepcopy(instr._p)
write_runs(Paragraph(partB_topic, doc),
           [{'t': 'Topic: the workplace.'}], BLACK)
partB_head._p.addnext(partB_topic)

VOCAB = [
    (('Many office workers ', 'have too much work', ' to finish in one day.'), 'struggle with an excessive workload'),
    (('Employees often ', 'feel so tired', ' that they cannot enjoy their weekends.'), 'suffer from burnout'),
    (('Companies can ', 'let employees start and finish work at different times', '.'), 'offer flexible working hours'),
    (('Managers should ', 'praise employees when they do a good job', '.'), 'recognise employees\u2019 contributions'),
    (('Some companies ', 'do not pay employees for the extra hours they work', '.'), 'leave overtime unpaid'),
    (('The company ', 'loses many young employees every year', '.'), 'struggles to retain young talent'),
]
for r, (basic, up) in enumerate(VOCAB, start=1):
    c1 = vocab.rows[r].cells[1]
    tmpl = c1.paragraphs[0].runs[0]._r
    if len(tmpl.findall(qn('w:rPr'))) and [c for c in tmpl.find(qn('w:rPr')).findall(qn('w:u'))]:
        pass
    write_runs(c1.paragraphs[0],
               [{'t': basic[0], 'i': True}, {'t': basic[1], 'i': True, 'u': True}, {'t': basic[2], 'i': True}],
               tmpl)
    c2 = vocab.rows[r].cells[2]
    write_runs(c2.paragraphs[0], [{'t': up, 'b': True}], c2.paragraphs[0].runs[0]._r)

# =========================================================== 4. PART C
C_BOX = ['THE BASIC EXAMPLES STUDENTS USUALLY WRITE',
         'For example, working long hours is bad for employees\u2019 health.',
         'For example, employees who answer messages at night never really rest.',
         'For example, an employee may miss dinner with his family because of overtime.',
         'For example, some governments have made laws to protect workers.']
cbox = basic_box.rows[0].cells[0]
while len(cbox.paragraphs) < len(C_BOX):
    cbox.paragraphs[-1]._p.addnext(deepcopy(cbox.paragraphs[-1]._p))
for i, txt in enumerate(C_BOX):
    par = cbox.paragraphs[i]
    tmpl = par.runs[0]._r
    write_runs(par, [{'t': txt, 'b': True if i == 0 else None}], tmpl)

WAYS = [
    dict(
        head='Way 1 \u2014 Findings (a named study)',
        formula='According to [source] ([year]), [finding].',
        basic='For example, working long hours is bad for employees\u2019 health.',
        up=[{'t': 'According to a 2021 joint study by the World Health Organization and the International Labour Organization', 'b': True},
            {'t': ', employees who work '},
            {'t': '55 hours or more a week', 'b': True},
            {'t': ' run a '},
            {'t': '35% higher risk of stroke', 'b': True},
            {'t': '.'}],
        note='\u201cAccording to\u201d + a NAMED source + the year is the whole formula \u2014 never \u201cstudies show\u201d, '
             'and never a number you have invented. Learn ONE finding for the topic and you own this way; the figures '
             'above are real (WHO and ILO, 2021) and can be quoted as they are.',
        label='Use Way 1'),
    dict(
        head='Way 2 \u2014 Typical Process',
        formula='Consider what happens when + [situation]: + [step 1 \u2192 step 2 \u2192 result].',
        basic='For example, employees who answer messages at night never really rest.',
        up=[{'t': 'Consider what happens when', 'b': True},
            {'t': ' a manager answers one \u201cquick\u201d message at bedtime: the phone stays on the pillow, the mind '
                  'starts drafting tomorrow\u2019s replies, and the hours meant for sleep are spent half at work.'}],
        note='No source and no name needed \u2014 the safety net for students who remember nothing. Keep it to three moves '
             '(first this, then that, so the harm), and make the LAST move the damage your argument is about. It must '
             'stay general: once one named person does one thing, it is Way 3.',
        label='Use Way 2'),
    dict(
        head='Way 3 \u2014 Hypothetical Scene',
        formula='Imagine + [a worker in an ordinary moment] + [what happens next] + [so what].',
        basic='For example, an employee may miss dinner with his family because of overtime.',
        up=[{'t': 'Imagine', 'b': True},
            {'t': ' a junior clerk in Kwun Tong '},
            {'t': 'who', 'b': True},
            {'t': ' leaves the office at nine every evening: by the time he reaches home, his daughter is asleep and '
                  'the dinner his mother kept warm has gone cold.'}],
        note='Give the marker a scene they can SEE: one worker, one place, one object (a phone, a cold dinner, a desk '
             'lamp). Write \u201cImagine a junior clerk \u2026\u201d, never \u201cImagine if you \u2026\u201d \u2014 the essay bans '
             '\u201cyou\u201d. Keep numbers out of the scene: they read as fake statistics.',
        label='Use Way 3'),
    dict(
        head='Way 4 \u2014 Real-world Example',
        formula='In + [month year], + [who did what] + \u2014 evidence that + [point].',
        basic='For example, some governments have made laws to protect workers.',
        up=[{'t': 'In August 2024', 'b': True},
            {'t': ', '},
            {'t': 'Australia gave employees the legal right to ignore work emails and calls after hours', 'b': True},
            {'t': ' \u2014 evidence that even governments now treat constant availability as a problem.'}],
        note='One real event, dated, with the country named, plus the little \u201cevidence that \u2026\u201d clause that ties it '
             'back to your argument. France did the same in January 2017 for companies with fifty or more staff. Never '
             'invent an event, and never quote one you cannot describe in a single sentence.',
        label='Use Way 4'),
]

for n in range(3):
    tb = pat_tbls[n]
    w = WAYS[n]
    row0 = tb.rows[0].cells[0]
    tmpl_b = row0.paragraphs[0].runs[0]._r
    write_runs(row0.paragraphs[0], [{'t': w['head'], 'b': True}], tmpl_b)
    write_runs(row0.paragraphs[1], [{'t': 'Target Formula: ', 'b': True},
                                    {'t': w['formula']}], tmpl_b)
    write_runs(tb.rows[1].cells[1].paragraphs[0], [{'t': w['basic']}],
               tb.rows[1].cells[1].paragraphs[0].runs[0]._r)
    write_runs(tb.rows[2].cells[1].paragraphs[0], w['up'],
               tb.rows[2].cells[1].paragraphs[0].runs[0]._r)
    note = teach_notes[n]
    write_runs(note, [{'t': 'Teaching note: ', 'b': True, 'c': RED}, {'t': w['note'], 'c': RED}],
               note.runs[0]._r)

# fourth block: clone pattern 3's table + its teaching note, insert before the design note
t4 = deepcopy(pat_tbls[2]._tbl)
w = WAYS[3]
T4 = Table(t4, doc)
tmpl_b = T4.rows[0].cells[0].paragraphs[0].runs[0]._r
write_runs(T4.rows[0].cells[0].paragraphs[0], [{'t': w['head'], 'b': True}], tmpl_b)
write_runs(T4.rows[0].cells[0].paragraphs[1], [{'t': 'Target Formula: ', 'b': True}, {'t': w['formula']}], tmpl_b)
write_runs(T4.rows[1].cells[1].paragraphs[0], [{'t': w['basic']}],
           T4.rows[1].cells[1].paragraphs[0].runs[0]._r)
write_runs(T4.rows[2].cells[1].paragraphs[0], w['up'], T4.rows[2].cells[1].paragraphs[0].runs[0]._r)
n4 = deepcopy(teach_notes[2]._p)
write_runs(Paragraph(n4, doc), [{'t': 'Teaching note: ', 'b': True, 'c': RED}, {'t': w['note'], 'c': RED}],
           RED_T)
blank_tmpl = deepcopy(design_note._p)
for r in blank_tmpl.findall(qn('w:r')):
    blank_tmpl.remove(r)
design_note._p.addprevious(deepcopy(blank_tmpl))
design_note._p.addprevious(t4)
design_note._p.addprevious(n4)
design_note._p.addprevious(deepcopy(blank_tmpl))

write_runs(design_note,
           [{'t': 'Teacher\u2019s note on the design: ', 'b': True, 'c': RED},
            {'t': 'today\u2019s four ways are a MENU, not four rules. The marker\u2019s test is simple \u2014 can I picture it, and '
                  'does it prove the point? \u2014 not \u201cdid you cite something?\u201d. Way 1 is the strongest but needs a '
                  'memorised finding; Ways 2 and 3 need no reading at all. ONE way per example: two ways in one '
                  'paragraph reads as two examples and usually loses the round-off. Keep the \u201cbecause\u201d workings in '
                  '[Explain] \u2014 an example illustrates, it does not explain.', 'c': RED}],
           design_note.runs[0]._r)

# =========================================================== 5. PART D
partC_head = find_p('PART C:')
write_runs(partC_head, [{'t': 'PART C: UPGRADING SENTENCE (FOUR WAYS TO GIVE EXAMPLES)', 'b': True}],
           partC_head.runs[0]._r)

write_runs(partD_intro,
           [{'t': 'Rewrite each weak example with the way shown, and use the way exactly as taught in Part C. The '
                  'topics are different from the working-hours topic in Part C. Sets 4\u20136 are more challenging: the '
                  'box gives a topic sentence AND a weak example, so you must rewrite the example and finish the '
                  'paragraph with a [Round off].'}], partD_intro.runs[0]._r)

partD_head = find_p('PART D:')
write_runs(partD_head, [{'t': 'PART D: REWRITING EXERCISE (PRACTICE SETS)', 'b': True}], partD_head.runs[0]._r)

SETS = [
    dict(topic='Keeping Young Staff',
         basic='For example, many young employees leave their jobs after one year.',
         label='Use Way 1',
         ans='According to Gallup\u2019s 2023 State of the Global Workplace report, 51% of employees worldwide are '
             'watching for or actively seeking a new job.',
         note='Any correctly framed finding is accepted (source + year + figure) \u2014 \u201cAccording to Gallup (2023), '
              'more than half of the world\u2019s employees are looking for another job.\u201d The figures here are real.'),
    dict(topic='Recognising Good Work',
         basic='For example, employees who are never praised stop trying.',
         label='Use Way 2',
         ans='Consider what happens when praise never comes: the extra effort goes unnoticed, the employee stops '
             'volunteering for the difficult tasks, and a competent worker quietly becomes an average one.',
         note='Accept any clear three-move process (first \u2026 then \u2026 so that \u2026). The last move must be the '
              'harm \u2014 the point the paragraph is arguing, not a repeat of the topic sentence.'),
    dict(topic='Working from Home',
         basic='For example, some employees find it hard to switch off when they work from home.',
         label='Use Way 3',
         ans='Imagine a sales executive whose office is now the kitchen table: one email is answered while the kettle '
             'boils, another during dinner, and by ten o\u2019clock the laptop is still open in front of the television.',
         note='Accept any scene with one worker, one place and one object. The final detail must show the harm (the '
              'evening has gone), not repeat the topic sentence.'),
    dict(topic='Messages After Office Hours', hard=True,
         ts='The company should stop sending work messages after office hours.',
         basic='For example, some companies have already made rules about this.',
         label='Use Way 4 + [Round off]',
         ans='In January 2017, France required companies with fifty or more staff to agree rules on out-of-hours work '
             'contact \u2014 evidence that even governments now treat constant availability as a problem. A manager who '
             'waits until morning sends the same instruction and keeps a colleague, and that habit costs our company '
             'nothing.',
         note='The dated event replaces the vague generalisation, and the round-off echoes the question\u2019s wording '
              '(\u201cwork messages after office hours\u201d).'),
    dict(topic='Overtime That Is Never Paid', hard=True,
         ts='Unpaid overtime has become normal in many Hong Kong offices.',
         basic='For example, staff work extra hours but get nothing back.',
         label='Use Way 2 + [Round off]',
         ans='Consider what happens when the extra hours are never recorded: the first hour is given \u201cjust this '
             'once\u201d, the second becomes expected, and by the end of the quarter the team is working a seventh day '
             'for the same salary. Overtime that nobody counts is the quietest way a company takes from its staff, '
             'and it is the first thing a new colleague notices.',
         note='Accept any three-move process. The round-off must return to the unpaid hours, not to \u201cover-time in '
              'general\u201d.'),
    dict(topic='Help for New Graduates', hard=True,
         ts='Employers should give fresh graduates more training and support.',
         basic='For example, new staff are given work they cannot do.',
         label='Use Way 3 + [Round off]',
         ans='Imagine a graduate in her first month who is handed a client\u2019s whole account and a one-line brief: '
             'she works until midnight, repeats the same mistakes, and learns only that the office is a lonely '
             'place. Patience with a new colleague in the first year is how a company grows its own experts instead '
             'of buying them.',
         note='The scene must be about the graduates themselves (not about a manager who trains them), and the '
              'round-off must echo the question\u2019s key words.'),
]

for i, (head, tbl, note) in enumerate(zip(set_heads, set_tbls, key_notes)):
    st_ = SETS[i]
    tag = 'Set %d. Topic: %s%s' % (i + 1, st_['topic'], ' (more challenging)' if st_.get('hard') else '')
    write_runs(head, [{'t': tag, 'b': True}], head.runs[0]._r)
    if st_.get('hard'):
        write_runs(tbl.rows[0].cells[0].paragraphs[0], [{'t': 'Topic sentence', 'b': True}], BLACK)
        write_runs(tbl.rows[0].cells[1].paragraphs[0], [{'t': st_['ts']}], tbl.rows[0].cells[1].paragraphs[0].runs[0]._r)
        write_runs(tbl.rows[1].cells[0].paragraphs[0], [{'t': ''}], BLACK)
        write_runs(tbl.rows[1].cells[1].paragraphs[0], [{'t': st_['basic']}], tbl.rows[1].cells[1].paragraphs[0].runs[0]._r)
        write_runs(tbl.rows[2].cells[0].paragraphs[0], [{'t': st_['label'], 'b': True}], BLACK)
        ans_par = tbl.rows[2].cells[1].paragraphs[0]
    else:
        write_runs(tbl.rows[0].cells[0].paragraphs[0], [{'t': 'Weak example', 'b': True}], BLACK)
        write_runs(tbl.rows[0].cells[1].paragraphs[0], [{'t': st_['basic']}], tbl.rows[0].cells[1].paragraphs[0].runs[0]._r)
        write_runs(tbl.rows[1].cells[0].paragraphs[0], [{'t': st_['label'], 'b': True}], BLACK)
        ans_par = tbl.rows[1].cells[1].paragraphs[0]
    write_runs(ans_par, [{'t': st_['ans'], 'b': True, 'c': RED}], ans_par.runs[0]._r)
    write_runs(note, [{'t': 'Key note: ', 'b': True, 'c': RED}, {'t': st_['note'], 'c': RED}], note.runs[0]._r)

# =========================================================== 6. PART E
pc = prompt_box.rows[0].cells[0]
write_runs(pc.paragraphs[0],
           [{'t': 'HKDSE adapted Paper 2 prompt (Workplace Communication \u2014 based on DSE 2018 Paper 2 Part B):'}],
           pc.paragraphs[0].runs[0]._r)
write_runs(pc.paragraphs[2],
           [{'t': 'You are the boss of Reboot Online Company and you have recently received complaints from some staff '
                  'about the number of work-related emails and text messages received out of office. Write a letter to '
                  'staff addressing their complaints.'}], pc.paragraphs[2].runs[0]._r)

write_runs(req1, [{'t': 'Make ONE suggestion (about 100 words) about reducing the number of work-related messages '
                        'staff receive outside office hours, and open your paragraph with ONE of Lesson 3\u2019s three '
                        'patterns as the topic sentence.'}], req1.runs[0]._r)
write_runs(chain, [{'t': 'Follow the chain: [Topic sentence] \u2192 [Explain] \u2192 [Example \u2014 using ONE of '
                         'today\u2019s four ways] \u2192 [Round off]. Underline your example and label it (Way 1 / 2 / 3 '
                         '/ 4), so your teacher can see which way you chose.'}], chain.runs[0]._r)

MODELS = [
    dict(head='Suggested paragraph 1 \u2014 the company\u2019s policy (Pattern 1 + Way 1: findings)',
         body=[('There is an urgent need for Reboot Online to put an out-of-hours contact policy in writing.', None),
               ('The problem is not the rare emergency but the drip of ordinary messages: a question about tomorrow\u2019s '
                'meeting sent at ten o\u2019clock still costs an employee the evening, because a mind that is half at '
                'work never fully rests.', None),
               ('According to a 2021 joint study by the World Health Organization and the International Labour '
                'Organization, staff who work 55 hours or more a week run a 35% higher risk of stroke than those who '
                'work 35 to 40 hours.', None),
               ('Protecting our evenings is therefore the cheapest health insurance this company can buy, and it costs '
                'nothing but a rule.', None)],
         labels=['  [Topic sentence \u2014 Pattern 1]', '  [Explain]', '  [Example \u2014 Way 1: findings]', '  [Round off]']),
    dict(head='Suggested paragraph 2 \u2014 the team meeting (Pattern 2 + Way 4: real-world example)',
         body=[('To give every colleague back their private time, Reboot Online should make it a priority to agree an '
                'out-of-hours contact policy with the whole team.', None),
               ('Vague goodwill does not survive a busy quarter, whereas a written policy tells each manager when a '
                'message really cannot wait until morning and when it must.', None),
               ('In August 2024, Australia gave employees the legal right to ignore work emails and calls after hours '
                '\u2014 evidence that even governments now treat constant availability as a problem.', None),
               ('A company that writes the rule down keeps its best people, and keeping them costs far less than '
                'replacing them.', None)],
         labels=['  [Topic sentence \u2014 Pattern 2]', '  [Explain]', '  [Example \u2014 Way 4: real-world example]',
                 '  [Round off]']),
    dict(head='Suggested paragraph 3 \u2014 the managers (Pattern 3 + Way 3: hypothetical scene)',
         body=[('Managers, in collaboration with the staff they lead, should agree on one hour each evening when nobody '
                'is expected to reply.', None),
               ('Respect is shown in small habits, and a team that knows when it may switch off plans its day better '
                'instead of scattering its work across the night.', None),
               ('Imagine a junior clerk at Reboot Online who leaves the office at seven: because his manager keeps '
                'until morning everything that can wait, he reaches home in time for dinner and returns the next day '
                'with something left to give.', None),
               ('A few quiet hours are a small price for staff who arrive rested, and our company would be foolish to '
                'pay more for less.', None)],
         labels=['  [Topic sentence \u2014 Pattern 3]', '  [Explain]', '  [Example \u2014 Way 3: hypothetical scene]',
                 '  [Round off]']),
]

for mh, mb, m in zip(model_heads, model_bodies, MODELS):
    write_runs(mh, [{'t': m['head'], 'b': True}], mh.runs[0]._r)
    specs = []
    for (txt, _), lab in zip(m['body'], m['labels']):
        specs.append({'t': txt})
        specs.append({'t': lab, 'c': RED})
    write_runs(mb, specs, mb.runs[0]._r)

write_runs(final_note,
           [{'t': 'Teacher\u2019s note: ', 'b': True, 'c': RED},
            {'t': 'each paragraph makes ONE suggestion and uses ONE example way \u2014 the topic sentence is the only '
                  'place where a Lesson 3 pattern is required, and the example is the only place where a way is '
                  'required. Models 2 and 3 need no memorised facts; Model 1 does. If a student hands in Way 2, check '
                  'that the last move states the harm; if a student hands in Way 4, check that the event is real, '
                  'dated and tied back with \u201cevidence that \u2026\u201d.', 'c': RED}],
           final_note.runs[0]._r)

# keep table rows whole on the page
cant_split(KT)
for t in pat_tbls + [T4] + set_tbls:
    cant_split(t)

# =========================================================== 7. marker block
sectPr = body.find(qn('w:sectPr'))
head_tmpl = deepcopy(req_head._p)
bullet_tmpl = deepcopy(partD_intro._p)
BLOCK = [
    ('h', 'What to reward when marking'),
    ('b', 'Content: ONE clear, workable suggestion for cutting out-of-hours messages, developed with a mechanism and '
          'proved with ONE example \u2014 not three ideas.'),
    ('b', 'Language: the topic sentence uses ONE Lesson 3 pattern accurately; the example names its source or builds '
          'its scene, with no invented figures, names or events.'),
    ('b', 'Organisation: [Topic sentence] \u2192 [Explain] \u2192 [Example] \u2192 [Round off], with the round-off echoing '
          'the words of the question (\u201cwork-related emails and text messages received out of office\u201d).'),
    ('h', 'Common errors to watch when marking'),
    ('b', 'Lesson 3 pattern forms: \u201cTo improving \u2026\u201d \u2192 \u201cTo improve \u2026\u201d (base verb after '
          'the goal); \u201cin collaboration of\u201d \u2192 \u201cin collaboration with\u201d; \u201cencourage staff '
          'taking\u201d \u2192 \u201cencourage staff to take\u201d.'),
    ('b', 'Two ways in one example (\u201cAccording to a 2021 study, imagine a clerk \u2026\u201d) \u2192 keep ONE way per '
          'example; an example never tied back to the argument loses the round-off mark; \u201cstudies show \u2026\u201d with '
          'no source, or a figure that cannot be sourced \u2192 rewrite as Way 2 or Way 3.'),
]
def set_spacing(par, before=None, after=None):
    pPr = par._p.get_or_add_pPr()
    sp = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = pPr.makeelement(qn('w:spacing'), {})
        pPr.append(sp)
    if before is not None:
        sp.set(qn('w:before'), str(before))
    if after is not None:
        sp.set(qn('w:after'), str(after))


for kind, txt in BLOCK:
    if kind == 'h':
        el = deepcopy(head_tmpl)
        par = Paragraph(el, doc)
        write_runs(par, [{'t': txt, 'b': True, 'c': RED}], RED_B)
        set_spacing(par, before=140, after=40)
    else:
        el = deepcopy(bullet_tmpl)
        par = Paragraph(el, doc)
        write_runs(par, [{'t': txt}], BLACK)
        set_spacing(par, before=0, after=40)
    sectPr.addprevious(el)

doc.save(OUT)
print('saved ->', OUT)
