"""Φτιάχνει το φύλλο χειροκίνητης αξιολόγησης 20 καρτών (evaluation/aksiologisi.xlsx).

Δείγμα: 20 τυχαίες κάρτες με σταθερό seed=42, μία ανά υπουργείο (ποικιλία).
Μία γραμμή ανά κάρτα και πεδίο: Ζήτημα, Ωφελούμενοι, Στόχοι, Κόστος -> 80 γραμμές.

Χρήση:  python evaluation/make_sheet.py
ΠΡΟΣΟΧΗ: αντικαθιστά το aksiologisi.xlsx· μην το ξανατρέξετε αφού αρχίσετε να βαθμολογείτε.
"""
import json
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / 'evaluation' / 'aksiologisi.xlsx'
N, SEED = 20, 42
SHEET = 'Αξιολόγηση'
REF = f"'{SHEET}'!"

FONT = 'Arial'
HEAD_FILL = PatternFill('solid', fgColor='1F3864')
INPUT_FILL = PatternFill('solid', fgColor='FFF2CC')   # κελιά που συμπληρώνει ο αξιολογητής
NA_FILL = PatternFill('solid', fgColor='EDEDED')      # δεν αφορούν αυτό το πεδίο
BAND_FILL = PatternFill('solid', fgColor='F3F6FA')    # εναλλάξ κάρτες, για να ξεχωρίζουν
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='top', wrap_text=True)

COLS = ['Α/Α κάρτας', 'Τίτλος', 'Σύνδεσμος PDF ΑΣΡ', 'Πεδίο', 'Τι εξήγαγε το LLM',
        'Πηγή/τεκμήρια', 'Βαθμός', 'Πλήθος στοιχείων', 'Σωστά στοιχεία',
        'Λείπει κάτι σημαντικό;', 'Σχόλιο', 'law_id']
WIDTHS = [9, 34, 14, 13, 60, 60, 16, 11, 11, 13, 30, 14]
FIELDS = ['Ζήτημα', 'Ωφελούμενοι', 'Στόχοι', 'Κόστος']


def sample_cards() -> pd.DataFrame:
    with open(ROOT / 'cards.jsonl', encoding='utf-8') as f:
        cards = pd.DataFrame([json.loads(l) for l in f if l.strip()])
    meta = pd.read_parquet(ROOT / 'data' / 'cards_public.parquet',
                           columns=['law_id', 'Τίτλος', 'Υπουργείο', 'Ανάλυση_Συνεπειών_Ρύθμισης'])
    df = meta.merge(cards[['law_id', 'card']], on='law_id')
    df = df.sort_values('law_id').sample(frac=1, random_state=SEED).reset_index(drop=True)
    # ποικιλία: πρώτα μία κάρτα από κάθε υπουργείο, με τη σειρά της τυχαίας ανακατάταξης
    df['_round'] = df.groupby('Υπουργείο').cumcount()
    df['_order'] = range(len(df))
    return df.sort_values(['_round', '_order']).head(N).reset_index(drop=True)


def numbered(items, key):
    return '\n'.join(f'{i}. {it[key]}' for i, it in enumerate(items, 1)) or '(κανένα)'


def field_rows(card):
    """(πεδίο, κείμενο LLM, τεκμήρια, πλήθος ή None)"""
    c = card['cost'] if isinstance(card['cost'], dict) else {}
    cost_text = c.get('summary') or 'Δεν αναφέρεται'
    if c.get('amount'):
        cost_text += f'\nΠοσό: {c["amount"]}'
    return [
        ('Ζήτημα', card['issue'], card.get('issue_evidence', ''), None),
        ('Ωφελούμενοι', numbered(card['beneficiaries'], 'group'),
         numbered(card['beneficiaries'], 'evidence'), len(card['beneficiaries'])),
        ('Στόχοι', numbered(card['goals'], 'goal'), numbered(card['goals'], 'evidence'),
         len(card['goals'])),
        ('Κόστος', cost_text, c.get('evidence', '') or '(καμία)', None),
    ]


def style_header(ws, row, ncols):
    for col in range(1, ncols + 1):
        cell = ws.cell(row, col)
        cell.font = Font(name=FONT, bold=True, color='FFFFFF')
        cell.fill = HEAD_FILL
        cell.alignment = CENTER
        cell.border = BORDER


def build_eval(ws, df):
    ws.append(COLS)
    style_header(ws, 1, len(COLS))
    ws.row_dimensions[1].height = 32

    dv_grade = DataValidation(type='list', formula1='"Σωστό,Μερικώς σωστό,Λάθος"', allow_blank=True,
                              error='Διαλέξτε Σωστό, Μερικώς σωστό ή Λάθος.', errorTitle='Βαθμός')
    dv_yesno = DataValidation(type='list', formula1='"Ναι,Όχι"', allow_blank=True,
                              error='Διαλέξτε Ναι ή Όχι.', errorTitle='Λείπει κάτι σημαντικό;')
    dv_count = DataValidation(type='whole', operator='between', formula1='0', formula2='999',
                              allow_blank=True, error='Γράψτε ακέραιο από 0 έως το πλήθος στοιχείων.',
                              errorTitle='Σωστά στοιχεία')
    for dv in (dv_grade, dv_yesno, dv_count):
        ws.add_data_validation(dv)

    r = 2
    for i, row in df.iterrows():
        card = row['card']
        band = BAND_FILL if i % 2 else None
        for field, text, ev, count in field_rows(card):
            values = [i + 1, row['Τίτλος'], 'Άνοιγμα PDF', field, text, ev,
                      None, count, None, None, None, row['law_id']]
            for col, v in enumerate(values, 1):
                cell = ws.cell(r, col, v)
                cell.font = Font(name=FONT, size=10)
                cell.alignment = WRAP
                cell.border = BORDER
                if band:
                    cell.fill = band
            ws.cell(r, 1).alignment = CENTER
            ws.cell(r, 4).font = Font(name=FONT, size=10, bold=True)
            ws.cell(r, 6).font = Font(name=FONT, size=9, italic=True, color='595959')
            ws.cell(r, 12).font = Font(name=FONT, size=8, color='808080')
            link = ws.cell(r, 3)
            link.hyperlink = row['Ανάλυση_Συνεπειών_Ρύθμισης']
            link.font = Font(name=FONT, size=10, color='0563C1', underline='single')
            ws.cell(r, 8).alignment = CENTER

            if count is None:   # Ζήτημα / Κόστος: βαθμός
                inputs, na = [7], [8, 9, 10]
                dv_grade.add(ws.cell(r, 7))
            else:               # λίστες: σωστά στοιχεία + ελλείψεις
                inputs, na = [9, 10], [7]
                dv_count.add(ws.cell(r, 9))
                dv_yesno.add(ws.cell(r, 10))
            for col in inputs + [11]:
                ws.cell(r, col).fill = INPUT_FILL
                ws.cell(r, col).alignment = CENTER if col != 11 else WRAP
            for col in na:
                ws.cell(r, col).fill = NA_FILL
            r += 1

    for col, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[ws.cell(1, col).column_letter].width = w
    ws.freeze_panes = 'E2'
    ws.auto_filter.ref = f'A1:L{r - 1}'
    ws.cell(1, 9).comment = Comment('Πόσα από τα αριθμημένα στοιχεία είναι σωστά (0 έως «Πλήθος στοιχείων»).',
                                    'Κάρτα Νόμου')
    return r - 1


def build_guide(ws):
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 95
    rows = [
        ('Οδηγίες αξιολόγησης', None),
        (None, None),
        ('Τι αξιολογούμε', f'{N} τυχαίες κάρτες (seed={SEED}, μία ανά υπουργείο). Για κάθε κάρτα '
                           'ελέγχουμε 4 πεδία: Ζήτημα, Ωφελούμενοι, Στόχοι, Κόστος.'),
        ('Πώς', 'Ανοίξτε το PDF της ΑΣΡ (στήλη «Σύνδεσμος PDF ΑΣΡ») και συγκρίνετε με τη στήλη '
                '«Τι εξήγαγε το LLM». Η στήλη «Πηγή/τεκμήρια» δείχνει το απόσπασμα που επικαλείται '
                'το μοντέλο· με Ctrl+F στο PDF το βρίσκετε γρήγορα.'),
        ('Τι συμπληρώνω', 'ΜΟΝΟ τα κίτρινα κελιά. Τα γκρι δεν αφορούν το συγκεκριμένο πεδίο. '
                          'Τα αποτελέσματα υπολογίζονται μόνα τους στο φύλλο «Αποτελέσματα».'),
        (None, None),
        ('Κριτήρια βαθμού', None),
        ('Σωστό', 'Αντιστοιχεί στο κείμενο και δεν λείπει κάτι σημαντικό.'),
        ('Μερικώς σωστό', 'Σωστό, αλλά ελλιπές ή ασαφές.'),
        ('Λάθος', 'Δεν υπάρχει στο κείμενο ή το διαστρεβλώνει.'),
        (None, None),
        ('Ζήτημα και Κόστος', 'Στήλη «Βαθμός»: Σωστό / Μερικώς σωστό / Λάθος (λίστα επιλογής). '
                              'Αν το LLM γράφει «Δεν αναφέρεται» για το κόστος και πράγματι η ΑΣΡ '
                              'δεν λέει τίποτα για κόστος, αυτό είναι Σωστό.'),
        ('Ωφελούμενοι και Στόχοι', '«Πλήθος στοιχείων»: έχει ήδη συμπληρωθεί (πόσα έδωσε το LLM).\n'
                                   '«Σωστά στοιχεία»: πόσα από τα αριθμημένα στοιχεία είναι σωστά '
                                   '(ένα στοιχείο είναι σωστό αν υπάρχει στο κείμενο και δεν '
                                   'διαστρεβλώνεται).\n'
                                   '«Λείπει κάτι σημαντικό;»: Ναι, αν η ΑΣΡ αναφέρει σημαντικό '
                                   'ωφελούμενο/στόχο που το LLM παρέλειψε.'),
        ('Σχόλιο', 'Προαιρετικό: τι ακριβώς ήταν λάθος ή τι έλειπε. Χρήσιμο για την αναφορά.'),
        (None, None),
        ('Παράδειγμα', 'Στόχοι, Πλήθος στοιχείων = 5 → γράφω Σωστά στοιχεία = 4, '
                       'Λείπει κάτι σημαντικό; = Όχι, Σχόλιο = «ο στόχος 3 δεν αναφέρεται στην ΑΣΡ».'),
        (None, None),
        ('Συνολικό ποσοστό', 'Κάθε Ζήτημα και κάθε Κόστος μετρά ως 1 στοιχείο· κάθε ωφελούμενος '
                             'και κάθε στόχος επίσης ως 1. Δίνονται δύο ποσοστά: αυστηρό (μόνο τα '
                             '«Σωστό») και επιεικές (το «Μερικώς σωστό» μετρά μισό).'),
    ]
    for i, (a, b) in enumerate(rows, 1):
        ca, cb = ws.cell(i, 1, a), ws.cell(i, 2, b)
        ca.font = Font(name=FONT, bold=True, size=10)
        cb.font = Font(name=FONT, size=10)
        ca.alignment = cb.alignment = WRAP
    ws['A1'].font = Font(name=FONT, bold=True, size=14)
    ws['A7'].font = Font(name=FONT, bold=True, size=12)
    for r in (8, 9, 10):
        ws.cell(r, 1).fill = INPUT_FILL


def build_results(ws, last):
    def rng(col):
        return f'{REF}${col}$2:${col}${last}'
    F, G, H, I, J = rng('D'), rng('G'), rng('H'), rng('I'), rng('J')

    def put(cell, value, bold=False, fmt=None):
        c = ws[cell]
        c.value = value
        c.font = Font(name=FONT, size=10, bold=bold)
        c.border = BORDER
        c.alignment = Alignment(horizontal='center' if cell[0] != 'A' else 'left',
                                vertical='center', wrap_text=True)
        if fmt:
            c.number_format = fmt

    ws.column_dimensions['A'].width = 34
    for col in 'BCDEFGH':
        ws.column_dimensions[col].width = 15
    ws['A1'] = 'Αποτελέσματα αξιολόγησης'
    ws['A1'].font = Font(name=FONT, bold=True, size=14)
    ws['A2'] = 'Υπολογίζονται αυτόματα από το φύλλο «Αξιολόγηση». Μην γράφετε σε αυτό το φύλλο.'
    ws['A2'].font = Font(name=FONT, italic=True, size=9, color='595959')

    put('A4', 'Κάρτες που έχουν βαθμολογηθεί (Ζήτημα)', bold=True)
    put('B4', '=B7')
    put('C4', f'=COUNTIF({F},"Ζήτημα")')
    put('D4', '=IF(C4=0,"",B4/C4)', fmt='0%')

    # Πίνακας 1: πεδία με βαθμό
    put('A6', 'Ζήτημα και Κόστος', bold=True)
    for col, h in zip('BCDEFGH', ['Βαθμολογήθηκαν', 'Σωστό', 'Μερικώς σωστό', 'Λάθος',
                                   '% Σωστό', '% Μερικώς', '% Λάθος']):
        put(f'{col}6', h, bold=True)
        ws[f'{col}6'].fill = PatternFill('solid', fgColor='D9E1F2')
    for r, field in ((7, 'Ζήτημα'), (8, 'Κόστος')):
        put(f'A{r}', field)
        put(f'B{r}', f'=C{r}+D{r}+E{r}')
        for col, grade in zip('CDE', ['Σωστό', 'Μερικώς σωστό', 'Λάθος']):
            put(f'{col}{r}', f'=COUNTIFS({F},"{field}",{G},"{grade}")')
        for col, src in zip('FGH', 'CDE'):
            put(f'{col}{r}', f'=IF($B{r}=0,"",{src}{r}/$B{r})', fmt='0.0%')

    # Πίνακας 2: πεδία-λίστες
    put('A11', 'Ωφελούμενοι και Στόχοι', bold=True)
    for col, h in zip('BCDEFG', ['Κάρτες βαθμολογημένες', 'Στοιχεία (πλήθος)', 'Σωστά στοιχεία',
                                  '% σωστά στοιχεία', 'Κάρτες με ελλείψεις', '% καρτών με ελλείψεις']):
        put(f'{col}11', h, bold=True)
        ws[f'{col}11'].fill = PatternFill('solid', fgColor='D9E1F2')
    for r, field in ((12, 'Ωφελούμενοι'), (13, 'Στόχοι')):
        put(f'A{r}', field)
        # «βαθμολογημένη» = έχει αριθμό στα Σωστά στοιχεία (και το 0 μετρά)
        put(f'B{r}', f'=COUNTIFS({F},"{field}",{I},">=0")')
        put(f'C{r}', f'=SUMIFS({H},{F},"{field}",{I},">=0")')
        put(f'D{r}', f'=SUMIFS({I},{F},"{field}")')
        put(f'E{r}', f'=IF(C{r}=0,"",D{r}/C{r})', fmt='0.0%')
        put(f'F{r}', f'=COUNTIFS({F},"{field}",{J},"Ναι")')
        answered = f'(F{r}+COUNTIFS({F},"{field}",{J},"Όχι"))'
        put(f'G{r}', f'=IF({answered}=0,"",F{r}/{answered})', fmt='0.0%')

    # Πίνακας 3: συνολικά
    put('A16', 'Συνολικά (όλα τα στοιχεία)', bold=True)
    for col, h in zip('BCDEF', ['Στοιχεία αξιολογημένα', 'Σωστά (αυστηρά)', '% σωστά (αυστηρά)',
                                 'Σωστά (Μερικώς = ½)', '% σωστά (Μερικώς = ½)']):
        put(f'{col}16', h, bold=True)
        ws[f'{col}16'].fill = PatternFill('solid', fgColor='D9E1F2')
    put('A17', 'Όλα τα πεδία')
    put('B17', '=B7+B8+C12+C13')
    put('C17', '=C7+C8+D12+D13')
    put('D17', '=IF(B17=0,"",C17/B17)', fmt='0.0%')
    put('E17', '=C17+0.5*(D7+D8)')
    put('F17', '=IF(B17=0,"",E17/B17)', fmt='0.0%')
    for cell in ('D17', 'F17'):
        ws[cell].font = Font(name=FONT, size=12, bold=True)

    notes = [
        'Σημειώσεις:',
        '• Μετρούν μόνο οι γραμμές που έχουν συμπληρωθεί· τα ποσοστά ενημερώνονται όσο προχωράτε.',
        '• Ζήτημα και Κόστος: 1 στοιχείο ανά κάρτα. Ωφελούμενοι και Στόχοι: κάθε αριθμημένο '
        'στοιχείο μετρά ως 1.',
        '• Αυστηρά: μετρά μόνο το «Σωστό». Μερικώς = ½: το «Μερικώς σωστό» μετρά μισό στοιχείο.',
        '• Οι λίστες δεν έχουν «μερικώς»· οι ελλείψεις φαίνονται χωριστά στο «% καρτών με ελλείψεις».',
    ]
    for i, t in enumerate(notes, 19):
        ws.cell(i, 1, t).font = Font(name=FONT, size=9, bold=(i == 19), color='404040')


def main():
    df = sample_cards()
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    last = build_eval(ws, df)
    build_guide(wb.create_sheet('Οδηγίες'))
    build_results(wb.create_sheet('Αποτελέσματα'), last)
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f'{len(df)} κάρτες, {last - 1} γραμμές -> {OUT.relative_to(ROOT)}')
    print('υπουργεία:', df['Υπουργείο'].nunique())


if __name__ == '__main__':
    main()
