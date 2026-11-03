"""Δημόσιο αντίγραφο της συμπληρωμένης αξιολόγησης, χωρίς τη στήλη «Πρωτότυπο κείμενο».

Το aksiologisi.xlsx περιέχει κείμενο ΑΣΡ από το gated dataset και δεν ανεβαίνει στο GitHub.
Το αντίγραφο κρατά βαθμούς, σχόλια, τεκμήρια και όλους τους τύπους.

Χρήση:  python evaluation/make_public.py   -> evaluation/aksiologisi_public.xlsx
"""
import re
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.worksheet.cell_range import CellRange

sys.path.insert(0, str(Path(__file__).resolve().parent))
from make_sheet import C_EV, C_FIELD, C_ORIG, C_TEXT, SHEET, row_height  # noqa: E402

ROOT = Path(__file__).resolve().parent
SRC, OUT = ROOT / 'aksiologisi.xlsx', ROOT / 'aksiologisi_public.xlsx'
GONE = C_ORIG                                  # η στήλη που αφαιρείται (G)
# Αξιολόγηση!$H$2  ή  Αξιολόγηση!$H$2:$H$81 (μετακινούνται και τα δύο άκρα)
REF = re.compile(r"((?:'%s'|%s)!\$?)([A-Z]+)(\$?\d+)(?:(:\$?)([A-Z]+)(\$?\d+))?" % (SHEET, SHEET))


def shift_ref(m) -> str:
    out = m[1] + shift_letter(m[2]) + m[3]
    if m[5]:
        out += m[4] + shift_letter(m[5]) + m[6]
    return out


def shift(col: int) -> int:
    return col - 1 if col > GONE else col


def shift_letter(letter: str) -> str:
    return get_column_letter(shift(column_index_from_string(letter)))


def main():
    wb = load_workbook(SRC)
    ws = wb[SHEET]
    assert ws.cell(1, GONE).value == 'Πρωτότυπο κείμενο', 'η στήλη δεν είναι εκεί που περιμέναμε'
    widths = {c: ws.column_dimensions[get_column_letter(c)].width for c in range(1, ws.max_column + 1)}

    ws.delete_cols(GONE)

    # πλάτη στηλών και φίλτρο: μετατόπιση μία θέση αριστερά
    for c, w in widths.items():
        if c != GONE:
            ws.column_dimensions[get_column_letter(shift(c))].width = w
    ws.auto_filter.ref = f'A1:{get_column_letter(ws.max_column)}{ws.max_row}'

    # λίστες επιλογής (Βαθμός, Ναι/Όχι, Σωστά στοιχεία)
    for dv in ws.data_validations.dataValidation:
        ranges = []
        for rng in dv.sqref.ranges:
            cr = CellRange(rng.coord)
            ranges.append(CellRange(min_col=shift(cr.min_col), min_row=cr.min_row,
                                    max_col=shift(cr.max_col), max_row=cr.max_row).coord)
        dv.sqref = ' '.join(ranges)

    # ύψη γραμμών: χωρίς το μεγάλο κείμενο δεν χρειάζονται 409 στιγμές
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = row_height(ws.cell(r, C_TEXT).value, ws.cell(r, C_EV).value)

    # τύποι στα άλλα φύλλα που δείχνουν στο «Αξιολόγηση»
    fixed = 0
    for other in wb.worksheets:
        if other is ws:
            continue
        for row in other.iter_rows():
            for cell in row:
                if cell.data_type == 'f' and isinstance(cell.value, str) and SHEET in cell.value:
                    new = REF.sub(shift_ref, cell.value)
                    fixed += new != cell.value
                    cell.value = new

    # οδηγίες: σημείωση ότι λείπει η στήλη
    guide = wb['Οδηγίες']
    for row in guide.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and 'Πρωτότυπο κείμενο' in cell.value:
                cell.value = ('[Δημόσια έκδοση] Η στήλη «Πρωτότυπο κείμενο» έχει αφαιρεθεί, γιατί περιείχε '
                              'κείμενο από το dataset (gated). Το πλήρες κείμενο βρίσκεται στο PDF της ΑΣΡ '
                              '(«Σύνδεσμος PDF ΑΣΡ»).\n\n' + cell.value)

    assert all(ws.cell(1, c).value != 'Πρωτότυπο κείμενο' for c in range(1, ws.max_column + 1))
    assert ws.cell(1, C_FIELD).value == 'Πεδίο'
    wb.save(OUT)
    print(f'{OUT.name}: {ws.max_column} στήλες, {ws.max_row - 1} γραμμές, {fixed} τύποι διορθώθηκαν')


if __name__ == '__main__':
    main()
