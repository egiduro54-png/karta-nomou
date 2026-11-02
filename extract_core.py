"""Αποκοπή των χρήσιμων τμημάτων της Ανάλυσης Συνεπειών Ρύθμισης (ΑΣΡ).

- core: ερωτήσεις 1-8 (ζήτημα, πρόβλημα, φορείς, στόχοι)
- cost: ο σχολιασμός κάτω από τον πίνακα «Κόστος αξιολογούμενης ρύθμισης»
        και, όπου υπάρχει, η έκθεση του άρθρου 75 παρ. 3 Συντάγματος
        και φράσεις όπως «δεν προκαλείται δαπάνη».

Χρήση:  python extract_core.py data/bills.parquet   -> γράφει cores.parquet
"""
import re
import sys

import pandas as pd

# Απόφαση καθηγητή: εκτός οι διεθνείς συμβάσεις· μένουν «Νόμος» και «Σχέδιο νόμου»
EXCLUDE_TYPES = ["Διεθνής Σύμβαση"]

START = re.compile(r'Ποιο ζήτημα αντιμετωπίζει')
END = re.compile(
    r'(?m)^\s*9\.\s|Ειδικότεροι στόχοι|Σύνοψη στόχων|Πίνακας Τροποποιούμενων'
    r'|Έκθεση γενικών συνεπειών', re.I)
COST_TABLE = re.compile(r'Κόστος αξιολογούμενης ρύθμισης', re.I)
COMMENT = re.compile(r'Σχολιασμός\s*/\s*ποιοτική\s+αποτίμηση', re.I)
ART75 = re.compile(r'ΕΚΘΕΣΗ ΤΟΥ ΑΡΘΡΟΥ 75 ΠΑΡ\.?\s?3|Έκθεση του άρθρου 75 παρ\.?\s?3', re.I)
SPEND = re.compile(r'(?:δαπάνη|επιβάρυνση)\s+(?:σε\s+βάρος|του)\s+(?:του\s+)?(?:κρατικού\s+)?προϋπολογισμ', re.I)
SCAN_LIMIT = 1_500_000  # τα τεράστια κείμενα είναι κυρίως διαβούλευση· δεν χρειάζεται όλο


def sentence_around(text: str, pos: int) -> str:
    a = text.rfind('.', max(0, pos - 300), pos) + 1
    b = text.find('.', pos, pos + 300)
    return text[a:(b + 1 if b != -1 else pos + 300)]

SECTION_END = re.compile(
    r'(?m)^\s*\|?\s*\d{2}\.\s|Κίνδυνοι αξιολογούμενης|ΕΝΟΤΗΤΑ|Έκθεση γενικών|Οφέλη αξιολογούμενης',
    re.I)


def clean(s: str) -> str:
    """Πετάει γραμμές-πίνακες χωρίς κείμενο και κρατά μόνο ό,τι έχει αρκετά γράμματα."""
    lines = [l for l in s.splitlines() if sum(ch.isalpha() for ch in l) >= 3]
    out = '\n'.join(lines).strip(' :\n')
    letters = sum(ch.isalpha() for ch in out)
    return out if letters >= 40 and letters / max(1, len(out)) > 0.5 else ''


MAX_CORE = 40_000      # χαρακτήρες, για να μένει μικρό το prompt
MAX_COST = 6_000


def core(text: str):
    m = START.search(text)
    if not m:
        return None
    e = END.search(text, m.end() + 200)
    stop = e.start() if e and e.start() - m.start() < 60_000 else m.start() + 30_000
    return text[m.start():stop][:MAX_CORE]


def cost(text: str):
    parts = []
    # 1) Σχολιασμοί κάτω από κάθε πίνακα κόστους (μπορεί να είναι πολλοί, ανά Μέρος)
    for t in COST_TABLE.finditer(text):
        c = COMMENT.search(text, t.end(), t.end() + 6_000)
        if c:
            chunk = text[c.end():c.end() + 1_500]
            stop = SECTION_END.search(chunk)
            chunk = clean(chunk[:stop.start()] if stop else chunk)
            if chunk:
                parts.append(chunk)
    # 2) Έκθεση άρθρου 75 παρ. 3 (εκτός αν απλώς παραπέμπει σε συνημμένη)
    a = list(ART75.finditer(text))
    if a:
        snippet = text[a[-1].end():a[-1].end() + 1_500]
        snippet = re.sub(r'^\s*του\s+Συντάγματος', '', snippet, flags=re.I)
        stop = SECTION_END.search(snippet)
        snippet = clean(snippet[:stop.start()] if stop else snippet)
        if snippet and not re.match(r'Βλ\.?\s*[Σσ]υνημμέν', snippet):
            parts.append('[Έκθεση άρθρου 75 παρ. 3] ' + snippet)
    # 3) Φράσεις για δαπάνη σε βάρος του προϋπολογισμού
    seen = set()
    for m in SPEND.finditer(text, 0, SCAN_LIMIT):
        s = re.sub(r'\s+', ' ', sentence_around(text, m.start())).strip()
        if s not in seen:
            seen.add(s)
            parts.append('[Φράση] ' + s)
        if len(seen) >= 5:
            break
    out = '\n---\n'.join(p for p in parts if p)
    return out[:MAX_COST] or None


if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else 'bills.parquet'
    df = pd.read_parquet(src)
    n = len(df)
    df = df[~df['Τύπος'].isin(EXCLUDE_TYPES)].copy()
    print(f'μένουν {len(df)} / {n} γραμμές (εκτός: {", ".join(EXCLUDE_TYPES)})')
    df['core'] = df.md_text.map(core)
    df['cost'] = df.md_text.map(cost)
    print('core βρέθηκε:', df.core.notna().sum(), '/', len(df))
    print('cost βρέθηκε:', df.cost.notna().sum(), '/', len(df))
    print('μήκος core (διάμεσος):', int(df.core.str.len().median()))
    cols = ['law_id', 'Τίτλος', 'Τύπος', 'Υπουργείο', 'Ημερομηνία_κατάθεσης',
            'Ημερομηνία_ψήφισης', 'Φάση_επεξεργασίας', 'Σύνδεσμος_νομοσχεδίου',
            'Ανάλυση_Συνεπειών_Ρύθμισης', 'core', 'cost']
    df[cols].to_parquet('cores.parquet')
    print('-> cores.parquet')
