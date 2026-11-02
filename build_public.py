"""Φτιάχνει το μικρό αρχείο που διαβάζει η εφαρμογή: μεταδεδομένα + κάρτες.

Ενώνει το cards.jsonl (από extract_cards.py) με τα μεταδεδομένα του cores.parquet.
ΔΕΝ κρατά md_text, core ή cost (το dataset είναι gated· ανεβαίνουν μόνο οι κάρτες).

Χρήση:  python build_public.py   -> γράφει data/cards_public.parquet
"""
import argparse
import json
import re

import pandas as pd

META = ['law_id', 'Τίτλος', 'Τύπος', 'Υπουργείο', 'Ημερομηνία_κατάθεσης',
        'Ημερομηνία_ψήφισης', 'Φάση_επεξεργασίας', 'Σύνδεσμος_νομοσχεδίου',
        'Ανάλυση_Συνεπειών_Ρύθμισης']

LAW_NO = re.compile(r'Ν\.\s*(\d+)')
FEK = re.compile(r'ΦΕΚ:\s*(.+)$')


def fix_cost(cost):
    """Σε λίγες κάρτες το κόστος ήρθε ως κομμένο κείμενο JSON· κρατάμε ό,τι διαβάζεται."""
    if isinstance(cost, dict):
        return cost
    try:
        return json.loads(cost)
    except (TypeError, ValueError):
        pass
    s = str(cost)
    m = re.search(r'"summary"\s*:\s*"(.*?)(?:"\s*,|$)', s, re.S)
    summary = m.group(1).strip() if m else 'Δεν αναφέρεται'
    return {'mentioned': '"mentioned":true' in s.replace(' ', ''),
            'summary': summary, 'amount': '', 'evidence': ''}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--cards', default='cards.jsonl')
    ap.add_argument('--cores', default='cores.parquet')
    ap.add_argument('--out', default='data/cards_public.parquet')
    args = ap.parse_args()

    with open(args.cards, encoding='utf-8') as f:
        rows = [json.loads(l) for l in f if l.strip()]
    fixed = 0
    for r in rows:
        if not isinstance(r['card']['cost'], dict):
            r['card']['cost'] = fix_cost(r['card']['cost'])
            fixed += 1
    cards = pd.DataFrame({
        'law_id': [r['law_id'] for r in rows],
        'model': [r['model'] for r in rows],
        'issue': [r['card']['issue'] for r in rows],  # ξεχωριστή στήλη για την αναζήτηση
        'card': [json.dumps(r['card'], ensure_ascii=False) for r in rows],
    }).drop_duplicates('law_id', keep='last')

    meta = pd.read_parquet(args.cores, columns=META)
    df = meta.merge(cards, on='law_id', how='inner')
    phase = df['Φάση_επεξεργασίας'].fillna('')
    df['Αριθμός_νόμου'] = phase.str.extract(LAW_NO)[0]
    df['ΦΕΚ'] = phase.str.extract(FEK)[0].str.strip()
    for c in ['Ημερομηνία_κατάθεσης', 'Ημερομηνία_ψήφισης']:
        df[c] = pd.to_datetime(df[c])

    df.to_parquet(args.out, index=False)
    print(f'{len(df)} κάρτες -> {args.out}')
    if fixed:
        print(f'διορθώθηκε κομμένο κόστος σε {fixed} κάρτες (καλό είναι να ξαναγίνουν)')
    missing = len(cards) - len(df)
    if missing:
        print(f'προσοχή: {missing} κάρτες δεν βρέθηκαν στο {args.cores}')


if __name__ == '__main__':
    main()
