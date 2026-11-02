"""Εξαγωγή «κάρτας νόμου» από κάθε ΑΣΡ με το Claude API.

Προϋπόθεση: έχεις τρέξει πρώτα  python extract_core.py bills.parquet  (φτιάχνει cores.parquet)

Παραδείγματα:
  python extract_cards.py --dry-run --sample 1          # δείχνει το prompt, χωρίς κλήση API
  python extract_cards.py --sample 5                     # 5 νομοσχέδια με το προεπιλεγμένο μοντέλο
  python extract_cards.py --sample 5 --model claude-haiku-5-5 --out cards_haiku.jsonl
  python extract_cards.py --all                          # όλα (συνεχίζει από όπου σταμάτησε)

Το κλειδί διαβάζεται από τη μεταβλητή περιβάλλοντος ANTHROPIC_API_KEY.
"""
import argparse
import json
import os
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd

DEFAULT_MODEL = 'claude-sonnet-5-5'

SYSTEM = (
    'Είσαι βοηθός που διαβάζει Αναλύσεις Συνεπειών Ρύθμισης (ΑΣΡ) ελληνικών νομοσχεδίων '
    'και συμπληρώνει μια δομημένη «κάρτα νόμου». Κανόνες:\n'
    '1. Χρησιμοποίησε ΜΟΝΟ πληροφορίες από το κείμενο που σου δίνεται. Μην προσθέτεις γνώσεις σου.\n'
    '2. Γράφε σε απλά, κατανοητά ελληνικά, για πολίτη χωρίς νομικές γνώσεις.\n'
    '3. Για κάθε στοιχείο δώσε ως «τεκμήριο» ένα ΑΥΤΟΥΣΙΟ σύντομο απόσπασμα (έως 30 λέξεις) '
    'από το κείμενο, αντιγραμμένο λέξη προς λέξη, που το στηρίζει.\n'
    '4. Αν κάτι δεν αναφέρεται, το λες ρητά (π.χ. κόστος: mentioned=false). Ποτέ μην επινοείς.\n'
    '5. Το κείμενο προέρχεται από μετατροπή PDF και μπορεί να έχει θόρυβο (σπασμένους πίνακες, '
    'αριθμούς σελίδων). Αγνόησέ τον. Αν το κείμενο είναι εντελώς αδιάβαστο, βάλε readable=false.\n'
    '6. Κόστος: αν mentioned=false, τότε summary=\'Δεν αναφέρεται\' και evidence κενό. Αν mentioned=true, '
    'περίγραψε τι ισχύει για το νομοσχέδιο ΣΥΝΟΛΙΚΑ, όχι για μία μεμονωμένη απόφαση ή άρθρο. '
    'Αν οι πληροφορίες είναι αντιφατικές ή αποσπασματικές, γράψε το απλά (π.χ. \'Κάποιες διατάξεις '
    'προκαλούν δαπάνη, το ποσό δεν αναφέρεται\').\n'
    '7. Μην αναφέρεσαι ποτέ στο \'κείμενο που δόθηκε\', στα \'αποσπάσματα\' ή στο prompt. '
    'Γράφε σαν να ενημερώνεις απευθείας τον πολίτη.'
)

CARD_TOOL = {
    'name': 'save_card',
    'description': 'Αποθηκεύει την κάρτα νόμου που εξήχθη από την ΑΣΡ.',
    'strict': True,  # το API εγγυάται ότι η απάντηση ακολουθεί ακριβώς το schema
    'input_schema': {
        'type': 'object',
        'additionalProperties': False,
        'properties': {
            'readable': {'type': 'boolean',
                         'description': 'false αν το κείμενο είναι αδιάβαστο (χαλασμένη μετατροπή).'},
            'issue': {'type': 'string',
                      'description': 'Το ζήτημα/πρόβλημα που αντιμετωπίζει η ρύθμιση, σε 2-3 προτάσεις.'},
            'issue_evidence': {'type': 'string', 'description': 'Αυτούσιο απόσπασμα που το στηρίζει.'},
            'beneficiaries': {
                'type': 'array',
                'description': 'Φορείς ή πληθυσμιακές ομάδες που αφορά/ωφελεί η ρύθμιση (έως 8).',
                'items': {'type': 'object',
                          'additionalProperties': False,
                          'properties': {'group': {'type': 'string'},
                                         'evidence': {'type': 'string'}},
                          'required': ['group', 'evidence']}},
            'goals': {
                'type': 'array',
                'description': 'Οι στόχοι της ρύθμισης, ο καθένας σε μία σύντομη φράση (έως 8).',
                'items': {'type': 'object',
                          'additionalProperties': False,
                          'properties': {'goal': {'type': 'string'},
                                         'evidence': {'type': 'string'}},
                          'required': ['goal', 'evidence']}},
            'cost': {
                'type': 'object',
                'additionalProperties': False,
                'properties': {
                    'mentioned': {'type': 'boolean',
                                  'description': 'true αν το κείμενο λέει κάτι για κόστος/δαπάνη '
                                                 '(και το «δεν προκαλείται δαπάνη» μετράει ως true).'},
                    'summary': {'type': 'string',
                                'description': 'Τι λέει για το κόστος σε 1-2 προτάσεις, '
                                               'ή «Δεν αναφέρεται».'},
                    'amount': {'type': 'string',
                               'description': 'Συγκεκριμένο ποσό αν δίνεται (π.χ. «2,5 εκ. ευρώ ετησίως»), αλλιώς κενό.'},
                    'evidence': {'type': 'string'}},
                'required': ['mentioned', 'summary', 'amount', 'evidence']},
        },
        'required': ['readable', 'issue', 'issue_evidence', 'beneficiaries', 'goals', 'cost'],
    },
}


def build_prompt(row) -> str:
    return (
        f'Τίτλος νομοσχεδίου: {row["Τίτλος"]}\n'
        f'Υπουργείο: {row["Υπουργείο"]}\n'
        f'Τύπος: {row["Τύπος"]}\n\n'
        f'<asr_erotiseis_1_8>\n{row["core"]}\n</asr_erotiseis_1_8>\n\n'
        f'<asr_kostos>\n{row["cost"] or "(δεν βρέθηκε κείμενο για το κόστος)"}\n</asr_kostos>\n\n'
        'Συμπλήρωσε την κάρτα νόμου καλώντας το εργαλείο save_card.\n'
        'Απάντησε ΜΟΝΟ με την κλήση του εργαλείου, χωρίς άλλο κείμενο.'
    )


def extract(client, model, row):
    msg = client.messages.create(
        model=model,
        max_tokens=8000,
        system=SYSTEM,
        tools=[CARD_TOOL],
        # τα Sonnet/Opus 5.5 δεν δέχονται 'tool'/'any'· το prompt ζητά ρητά την κλήση
        tool_choice={'type': 'auto'},
        messages=[{'role': 'user', 'content': build_prompt(row)}],
    )
    # μόνο ολοκληρωμένη κλήση εργαλείου· αλλιώς (π.χ. max_tokens) η κάρτα θα ήταν κομμένη.
    # Δεν γράφεται τίποτα και ξαναδοκιμάζεται στο επόμενο τρέξιμο.
    if msg.stop_reason != 'tool_use':
        raise RuntimeError(f'ημιτελής απάντηση (stop_reason={msg.stop_reason})')
    card = next((b.input for b in msg.content if b.type == 'tool_use'), None)
    if card is None:
        raise RuntimeError('δεν κλήθηκε το save_card')
    return {'law_id': row['law_id'], 'model': model, 'card': card,
            'usage': {'input': msg.usage.input_tokens, 'output': msg.usage.output_tokens}}


def pick(df, args):
    df = df[df.core.notna()]
    if args.ids:
        return df[df.law_id.isin(args.ids.split(','))]
    if args.all:
        return df
    # ποικίλο δείγμα: διαφορετικοί τύποι και υπουργεία, σταθερό seed
    d = df.sample(frac=1, random_state=args.seed).drop_duplicates('Υπουργείο')
    is_conv = d['Τύπος'] == 'Διεθνής Σύμβαση'
    conv = d[is_conv].head(max(1, args.sample // 3))
    laws = d[~is_conv].head(args.sample - len(conv))
    return pd.concat([laws, conv])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--cores', default='cores.parquet')
    ap.add_argument('--model', default=DEFAULT_MODEL)
    ap.add_argument('--out', default='cards.jsonl')
    ap.add_argument('--sample', type=int, default=5)
    ap.add_argument('--seed', type=int, default=42)
    ap.add_argument('--ids', help='law_id χωρισμένα με κόμμα')
    ap.add_argument('--all', action='store_true')
    ap.add_argument('--workers', type=int, default=4)
    ap.add_argument('--dry-run', action='store_true')
    args = ap.parse_args()

    df = pd.read_parquet(args.cores)
    rows = pick(df, args)

    done = set()
    if os.path.exists(args.out):
        with open(args.out, encoding='utf-8') as f:
            done = {json.loads(l)['law_id'] for l in f if l.strip()}
    rows = rows[~rows.law_id.isin(done)]
    print(f'Προς επεξεργασία: {len(rows)} (ήδη έτοιμα: {len(done)})')

    if args.dry_run:
        r = rows.iloc[0]
        p = build_prompt(r)
        print(SYSTEM, '\n\n', p[:3000], '\n...\n')
        print(f'~{len(p) // 2:,} tokens εισόδου (χονδρική εκτίμηση για ελληνικά)')
        return

    import anthropic
    client = anthropic.Anthropic()
    tot_in = tot_out = 0
    with open(args.out, 'a', encoding='utf-8') as f, ThreadPoolExecutor(args.workers) as ex:
        futs = {ex.submit(extract, client, args.model, r): r for _, r in rows.iterrows()}
        for i, fut in enumerate(as_completed(futs), 1):
            r = futs[fut]
            try:
                res = fut.result()
            except Exception as e:  # συνεχίζουμε· θα ξαναδοκιμαστεί στο επόμενο τρέξιμο
                print(f'[{i}] ΣΦΑΛΜΑ {r["law_id"]}: {e}', file=sys.stderr)
                continue
            f.write(json.dumps(res, ensure_ascii=False) + '\n')
            f.flush()
            tot_in += res['usage']['input']
            tot_out += res['usage']['output']
            print(f'[{i}/{len(futs)}] {r["Τίτλος"][:70]}')
    print(f'Tokens: {tot_in:,} εισόδου, {tot_out:,} εξόδου -> {args.out}')


if __name__ == '__main__':
    main()
