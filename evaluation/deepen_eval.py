"""Βαθύτερη αξιολόγηση: προσθήκες ΠΑΝΩ στο ήδη συμπληρωμένο evaluation/aksiologisi.xlsx.

Δεν σβήνει τίποτα· κρατά όλους τους βαθμούς και τα σχόλια.
1. Φύλλο «Αξιολόγηση»: στήλη «Πόσα λείπουν» (0 όπου «Λείπει κάτι σημαντικό;» = Όχι,
   κενό κίτρινο όπου = Ναι).
2. Νέο φύλλο «Ανάλυση λαθών»: 60 κενές γραμμές με λίστες επιλογής.
3. Φύλλο «Αποτελέσματα»: ποσοστό σωστά / (όσα έδωσε το LLM + όσα λείπουν) και λάθη ανά τύπο.

Χρήση:  python evaluation/deepen_eval.py
Τρέχει μία φορά· αν βρει ήδη τη στήλη «Πόσα λείπουν», σταματά χωρίς αλλαγές.
"""
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
from make_sheet import (BORDER, CENTER, FIELDS, FONT, HEAD_FILL, INPUT_FILL,  # noqa: E402
                        NA_FILL, SHEET, WRAP, C_FIELD, C_MISS)

PATH = Path(__file__).resolve().parent / 'aksiologisi.xlsx'
C_HOWMANY = C_MISS + 1          # η νέα στήλη μπαίνει αμέσως μετά το «Λείπει κάτι σημαντικό;»
ERR = 'Ανάλυση λαθών'
ERR_ROWS = 60
ERR_LAST = 500                  # οι τύποι μετρούν ως εδώ, αν προστεθούν γραμμές
KINDS = ['Λάθος', 'Μερικώς', 'Λείπει']
TYPES = ['Δευτερεύουσα ρύθμιση ως στόχος/ωφελούμενος', 'Διεύρυνση πέρα από το κείμενο',
         'Γενικόλογη ή ασαφής διατύπωση', 'Παράλειψη σημαντικού στοιχείου',
         'Ελλιπής περιγραφή κόστους', 'Άλλο']
SUB_FILL = PatternFill('solid', fgColor='D9E1F2')


def cell(ws, ref, value=None, bold=False, fmt=None, fill=None, align=CENTER, size=10):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = Font(name=FONT, size=size, bold=bold)
    c.border = BORDER
    c.alignment = align
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def add_howmany(ws):
    last = ws.max_row
    widths = {c: ws.column_dimensions[get_column_letter(c)].width for c in range(1, ws.max_column + 1)}
    ws.insert_cols(C_HOWMANY)
    # τα πλάτη δεν μετακινούνται μόνα τους
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c + 1 if c >= C_HOWMANY else c)].width = w
    ws.column_dimensions[get_column_letter(C_HOWMANY)].width = 11
    ws.auto_filter.ref = f'A1:{get_column_letter(ws.max_column)}{last}'

    head = ws.cell(1, C_HOWMANY, 'Πόσα λείπουν')
    head.font = Font(name=FONT, bold=True, color='FFFFFF')
    head.fill, head.alignment, head.border = HEAD_FILL, CENTER, BORDER
    head.comment = Comment('Πόσα σημαντικά στοιχεία παρέλειψε το LLM. 0 όταν δεν λείπει τίποτα.',
                           'Κάρτα Νόμου')

    dv = DataValidation(type='whole', operator='between', formula1='0', formula2='99', allow_blank=True,
                        error='Γράψτε ακέραιο ≥ 0.', errorTitle='Πόσα λείπουν')
    ws.add_data_validation(dv)
    pending = 0
    for r in range(2, last + 1):
        c = ws.cell(r, C_HOWMANY)
        c.border = BORDER
        c.alignment = CENTER
        c.font = Font(name=FONT, size=10)
        if ws.cell(r, C_FIELD).value in ('Ωφελούμενοι', 'Στόχοι'):
            c.fill = INPUT_FILL
            dv.add(c)
            if ws.cell(r, C_MISS).value == 'Όχι':
                c.value = 0
            else:
                pending += 1        # «Ναι» (ή κενό): το συμπληρώνει ο αξιολογητής
        else:
            c.fill = NA_FILL
    return last, pending


def add_error_sheet(wb):
    ws = wb.create_sheet(ERR, index=wb.sheetnames.index('Αποτελέσματα'))
    heads = ['Α/Α κάρτας', 'Πεδίο', 'Στοιχείο (κείμενο)', 'Είδος', 'Τύπος λάθους', 'Σχόλιο']
    for col, (h, w) in enumerate(zip(heads, [10, 14, 60, 12, 42, 50]), 1):
        c = ws.cell(1, col, h)
        c.font = Font(name=FONT, bold=True, color='FFFFFF')
        c.fill, c.alignment, c.border = HEAD_FILL, CENTER, BORDER
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 30

    validations = [
        ('A', DataValidation(type='whole', operator='between', formula1='1', formula2='20',
                             allow_blank=True, error='Αριθμός κάρτας από 1 έως 20.')),
        ('B', DataValidation(type='list', formula1='"' + ','.join(FIELDS) + '"', allow_blank=True)),
        ('D', DataValidation(type='list', formula1='"' + ','.join(KINDS) + '"', allow_blank=True)),
        ('E', DataValidation(type='list', formula1='"' + ','.join(TYPES) + '"', allow_blank=True)),
    ]
    assert len(validations[3][1].formula1) <= 255, 'το Excel δέχεται έως 255 χαρακτήρες σε λίστα'
    for col, dv in validations:
        ws.add_data_validation(dv)
        dv.add(f'{col}2:{col}{ERR_ROWS + 1}')
    for r in range(2, ERR_ROWS + 2):
        for col in range(1, 7):
            c = ws.cell(r, col)
            c.fill, c.border = INPUT_FILL, BORDER
            c.font = Font(name=FONT, size=10)
            c.alignment = WRAP if col in (3, 6) else CENTER
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:F{ERR_ROWS + 1}'


def add_results(ws, last):
    def rng(col):
        return f'{SHEET}!${col}$2:${col}${last}'
    FLD = rng(get_column_letter(C_FIELD))
    HOW = rng(get_column_letter(C_HOWMANY))
    MISS = rng(get_column_letter(C_MISS))
    E_KIND, E_TYPE = f"'{ERR}'!$D$2:$D${ERR_LAST}", f"'{ERR}'!$E$2:$E${ERR_LAST}"

    # δίπλα στο συνολικό ποσοστό χωρίς ελλείψεις (γραμμή 17)
    for ref, h in (('G16', 'Στοιχεία που λείπουν'), ('H16', '% σωστά (με όσα λείπουν)')):
        cell(ws, ref, h, bold=True, fill=SUB_FILL, align=Alignment(horizontal='center',
                                                                  vertical='center', wrap_text=True))
    cell(ws, 'G17', '=SUM(D29:D32)')
    cell(ws, 'H17', '=IF(B17+G17=0,"",C17/(B17+G17))', fmt='0.0%', size=12, bold=True)
    for col in 'GH':
        ws.column_dimensions[col].width = 15

    # ανά πεδίο
    cell(ws, 'A26', 'Ποσοστό με όσα λείπουν: σωστά / (όσα έδωσε το LLM + όσα λείπουν)', bold=True,
         align=Alignment(vertical='center'), size=11).border = None
    heads = ['Πεδίο', 'Σωστά', 'Έδωσε το LLM', 'Λείπουν', 'Σύνολο αναμενόμενων',
             '% σωστά (αυστηρό)', '% σωστά (με όσα λείπουν)', 'Εκκρεμούν «Πόσα λείπουν»']
    for col, h in enumerate(heads):
        cell(ws, f'{"ABCDEFGH"[col]}28', h, bold=True, fill=SUB_FILL,
             align=Alignment(horizontal='center', vertical='center', wrap_text=True))
    ws.row_dimensions[28].height = 45
    # Ζήτημα/Κόστος: 1 στοιχείο ανά κάρτα, δεν μετράμε ελλείψεις (βαθμολογήθηκαν ως σύνολο)
    src = {'Ζήτημα': ('C7', 'B7', None), 'Κόστος': ('C8', 'B8', None),
           'Ωφελούμενοι': ('D12', 'C12', 'Ωφελούμενοι'), 'Στόχοι': ('D13', 'C13', 'Στόχοι')}
    for r, field in enumerate(FIELDS, 29):
        ok, given, lst = src[field]
        cell(ws, f'A{r}', field, align=WRAP)
        cell(ws, f'B{r}', f'={ok}')
        cell(ws, f'C{r}', f'={given}')
        cell(ws, f'D{r}', f'=SUMIFS({HOW},{FLD},"{lst}")' if lst else 0)
        cell(ws, f'H{r}', (f'=COUNTIFS({FLD},"{lst}",{MISS},"Ναι")'
                           f'-COUNTIFS({FLD},"{lst}",{MISS},"Ναι",{HOW},">=0")') if lst else 0)
        cell(ws, f'E{r}', f'=C{r}+D{r}')
        cell(ws, f'F{r}', f'=IF(C{r}=0,"",B{r}/C{r})', fmt='0.0%')
        cell(ws, f'G{r}', f'=IF(E{r}=0,"",B{r}/E{r})', fmt='0.0%')
    r = 33
    cell(ws, f'A{r}', 'Συνολικά', bold=True, align=WRAP)
    for col in 'BCDEH':
        cell(ws, f'{col}{r}', f'=SUM({col}29:{col}32)', bold=True)
    cell(ws, f'F{r}', f'=IF(C{r}=0,"",B{r}/C{r})', fmt='0.0%', bold=True)
    cell(ws, f'G{r}', f'=IF(E{r}=0,"",B{r}/E{r})', fmt='0.0%', bold=True)
    notes = ['• Στο Ζήτημα και στο Κόστος δεν μετρώνται ελλείψεις (1 στοιχείο ανά κάρτα), '
             'άρα τα δύο ποσοστά συμπίπτουν.',
             '• «Εκκρεμούν»: γραμμές με «Λείπει κάτι σημαντικό; = Ναι» όπου το «Πόσα λείπουν» '
             'είναι ακόμη κενό. Όσο δεν είναι 0, το ποσοστό «με όσα λείπουν» δεν είναι τελικό.']
    for i, t in enumerate(notes, 34):
        ws.cell(i, 1, t).font = Font(name=FONT, size=9, color='404040')

    # λάθη ανά τύπο
    cell(ws, 'A37', 'Λάθη ανά τύπο (από το φύλλο «Ανάλυση λαθών»)', bold=True,
         align=Alignment(vertical='center'), size=11).border = None
    for col, h in zip('ABCDEF', ['Τύπος λάθους'] + KINDS + ['Σύνολο', '% του συνόλου']):
        cell(ws, f'{col}39', h, bold=True, fill=SUB_FILL,
             align=Alignment(horizontal='center', vertical='center', wrap_text=True))
    for r, t in enumerate(TYPES, 40):
        cell(ws, f'A{r}', t, align=WRAP)
        for col, kind in zip('BCD', KINDS):
            cell(ws, f'{col}{r}', f'=COUNTIFS({E_TYPE},"{t}",{E_KIND},"{kind}")')
        cell(ws, f'E{r}', f'=SUM(B{r}:D{r})', bold=True)
        cell(ws, f'F{r}', f'=IF($E$46=0,"",E{r}/$E$46)', fmt='0.0%')
    cell(ws, 'A46', 'Σύνολο', bold=True, align=WRAP)
    for col in 'BCDE':
        cell(ws, f'{col}46', f'=SUM({col}40:{col}45)', bold=True)
    cell(ws, 'F46', '=IF($E$46=0,"",E46/$E$46)', fmt='0.0%', bold=True)
    cell(ws, 'A47', f'=IF(COUNTA({E_KIND})>E46,"Προσοχή: "&(COUNTA({E_KIND})-E46)'
                    f'&" γραμμές έχουν Είδος αλλά όχι Τύπο λάθους","")',
         align=Alignment(vertical='center')).border = None
    ws['A47'].font = Font(name=FONT, size=9, bold=True, color='C00000')
    ws.column_dimensions['A'].width = 44


def add_guide(ws):
    start = ws.max_row + 2
    rows = [
        ('Βαθύτερη αξιολόγηση', None),
        ('Πόσα λείπουν', 'Στήλη στο φύλλο «Αξιολόγηση», για Ωφελούμενους και Στόχους: πόσα σημαντικά '
                         'στοιχεία της ΑΣΡ παρέλειψε το LLM. Έχει ήδη 0 όπου «Λείπει κάτι σημαντικό;» '
                         '= Όχι· συμπληρώστε τις γραμμές όπου = Ναι.'),
        ('Ανάλυση λαθών', 'Μία γραμμή για κάθε στοιχείο που ήταν Λάθος, Μερικώς σωστό ή Λείπει: '
                          'αριθμός κάρτας, πεδίο, το στοιχείο (αντιγραφή του κειμένου ή περιγραφή '
                          'αυτού που λείπει), είδος, τύπος λάθους από τη λίστα, σχόλιο.'),
        ('Τύποι λάθους', '\n'.join(f'• {t}' for t in TYPES)),
        ('Νέο ποσοστό', 'Στο φύλλο «Αποτελέσματα»: σωστά / (όσα έδωσε το LLM + όσα λείπουν). '
                        'Μετρά και τις παραλείψεις, άρα είναι αυστηρότερο από το ποσοστό χωρίς ελλείψεις.'),
    ]
    for i, (a, b) in enumerate(rows, start):
        ca, cb = ws.cell(i, 1, a), ws.cell(i, 2, b)
        ca.font = Font(name=FONT, bold=True, size=12 if b is None else 10)
        cb.font = Font(name=FONT, size=10)
        ca.alignment = cb.alignment = WRAP


def main():
    wb = load_workbook(PATH)
    ws = wb[SHEET]
    if any(ws.cell(1, c).value == 'Πόσα λείπουν' for c in range(1, ws.max_column + 1)):
        sys.exit('Η στήλη «Πόσα λείπουν» υπάρχει ήδη· δεν έγινε καμία αλλαγή.')
    assert ws.cell(1, C_MISS).value == 'Λείπει κάτι σημαντικό;', 'απρόσμενη διάταξη στηλών'
    last, pending = add_howmany(ws)
    add_error_sheet(wb)
    add_results(wb['Αποτελέσματα'], last)
    add_guide(wb['Οδηγίες'])
    wb.save(PATH)
    print(f'OK: «Πόσα λείπουν» με {pending} κενά για συμπλήρωση· φύλλο «{ERR}» με {ERR_ROWS} γραμμές.')


if __name__ == '__main__':
    main()
