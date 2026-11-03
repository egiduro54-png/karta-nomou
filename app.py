"""Εφαρμογή «Κάρτα Νόμου» (Streamlit).

Διαβάζει έτοιμες κάρτες από data/cards_public.parquet (φτιάχνεται με build_public.py).
ΔΕΝ καλεί το LLM.

Εκκίνηση:  streamlit run app.py
"""
import json
import unicodedata
from pathlib import Path

import pandas as pd
import streamlit as st

DATA = Path(__file__).parent / 'data' / 'cards_public.parquet'
MODEL_NAME = 'Claude Sonnet 5.5'

st.set_page_config(page_title='Κάρτα Νόμου', page_icon='📜', layout='wide')


def plain(s: str) -> str:
    """Πεζά και χωρίς τόνους, ώστε το «νομος» να βρίσκει και το «Νόμος»."""
    s = unicodedata.normalize('NFD', str(s).lower())
    return ''.join(ch for ch in s if unicodedata.category(ch) != 'Mn')


def esc(s) -> str:
    """Το Streamlit διαβάζει το $ ως μαθηματικό τύπο· το «σβήνουμε»."""
    return str(s or '').replace('$', r'\$')


def date_gr(d) -> str:
    return d.strftime('%d/%m/%Y') if pd.notna(d) else '—'


@st.cache_data
def load() -> pd.DataFrame:
    """Διαβάζει και προετοιμάζει τα δεδομένα μία φορά· τα reruns παίρνουν το έτοιμο αποτέλεσμα."""
    df = pd.read_parquet(DATA)
    df['card'] = df['card'].map(json.loads)
    df['Έτος'] = df['Ημερομηνία_κατάθεσης'].dt.year
    df['Κατάθεση'] = df['Ημερομηνία_κατάθεσης'].map(date_gr)
    df['_search'] = (df['Τίτλος'] + ' ' + df['issue']).map(plain)
    return df.sort_values('Ημερομηνία_κατάθεσης', ascending=False).reset_index(drop=True)


@st.cache_data
def filter_options():
    """Λίστα υπουργείων και εύρος ετών για τα φίλτρα."""
    df = load()
    return (sorted(df['Υπουργείο'].dropna().unique()),
            int(df['Έτος'].min()), int(df['Έτος'].max()))


def evidence(text: str):
    if text:
        st.caption(f'Πηγή: «{esc(text)}»')


def show_card(row):
    card = row['card']
    st.subheader(esc(row['Τίτλος']))

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(f'**Υπουργείο**  \n{esc(row["Υπουργείο"])}')
    c2.markdown(f'**Κατάθεση**  \n{date_gr(row["Ημερομηνία_κατάθεσης"])}')
    c3.markdown(f'**Ψήφιση**  \n{date_gr(row["Ημερομηνία_ψήφισης"])}')
    law = f'ν. {row["Αριθμός_νόμου"]}' if pd.notna(row['Αριθμός_νόμου']) else '—'
    fek = f'  \nΦΕΚ {esc(row["ΦΕΚ"])}' if pd.notna(row['ΦΕΚ']) else ''
    c4.markdown(f'**Νόμος / ΦΕΚ**  \n{law}{fek}')

    l1, l2, _ = st.columns([1, 1, 2])
    l1.link_button('Σελίδα στη Βουλή', row['Σύνδεσμος_νομοσχεδίου'], width='stretch')
    if row['Ανάλυση_Συνεπειών_Ρύθμισης']:
        l2.link_button('PDF Ανάλυσης Συνεπειών', row['Ανάλυση_Συνεπειών_Ρύθμισης'],
                       width='stretch')

    if not card.get('readable', True):
        st.warning('Το κείμενο της Ανάλυσης Συνεπειών δεν ήταν αναγνώσιμο· η κάρτα μπορεί να είναι ελλιπής.')

    st.markdown('#### Ζήτημα')
    st.markdown(esc(card['issue']))
    evidence(card.get('issue_evidence'))

    st.markdown('#### Ωφελούμενοι')
    if card['beneficiaries']:
        for b in card['beneficiaries']:
            st.markdown(f'- {esc(b["group"])}')
            evidence(b.get('evidence'))
    else:
        st.markdown('Δεν αναφέρονται.')

    st.markdown('#### Στόχοι')
    if card['goals']:
        for g in card['goals']:
            st.markdown(f'- {esc(g["goal"])}')
            evidence(g.get('evidence'))
    else:
        st.markdown('Δεν αναφέρονται.')

    st.markdown('#### Κόστος')
    cost = card['cost']
    if cost.get('mentioned'):
        st.markdown(esc(cost['summary']))
        if cost.get('amount'):
            st.markdown(f'**Ποσό:** {esc(cost["amount"])}')
        evidence(cost.get('evidence'))
    else:
        st.markdown('Δεν αναφέρεται.')


# ---------------------------------------------------------------- σελίδα
st.title('📜 Κάρτα Νόμου')
st.markdown('Για κάθε νομοσχέδιο της Βουλής (2020-2026), μια σύντομη κάρτα με το ζήτημα που '
            'αντιμετωπίζει, τους στόχους, ποιους αφορά και τι λέει για το κόστος, βγαλμένη από '
            'την Ανάλυση Συνεπειών Ρύθμισης.')

if not DATA.exists():
    st.error('Δεν βρέθηκε το αρχείο data/cards_public.parquet. Τρέξτε πρώτα: python build_public.py')
    st.stop()

df = load()
all_ministries, y0, y1 = filter_options()

with st.sidebar:
    st.header('Φίλτρα')
    query = st.text_input('Αναζήτηση', placeholder='λέξη στον τίτλο ή στο ζήτημα')
    ministries = st.multiselect('Υπουργείο', all_ministries, placeholder='Όλα τα υπουργεία')
    years = st.slider('Έτος κατάθεσης', y0, y1, (y0, y1))

mask = df['Έτος'].between(*years)
if ministries:
    mask &= df['Υπουργείο'].isin(ministries)
for word in plain(query).split():
    mask &= df['_search'].str.contains(word, regex=False)
res = df[mask].reset_index(drop=True)

st.markdown(f'**Βρέθηκαν {len(res)} νομοσχέδια**' if len(res) != 1 else '**Βρέθηκε 1 νομοσχέδιο**')

if res.empty:
    st.info('Κανένα νομοσχέδιο δεν ταιριάζει με τα φίλτρα. Δοκιμάστε κάτι πιο γενικό.')
else:
    table = res[['Τίτλος', 'Υπουργείο', 'Κατάθεση']]
    st.caption('Κάντε κλικ σε μια γραμμή για να ανοίξει η κάρτα.')
    event = st.dataframe(table, hide_index=True, width='stretch', height=320,
                         on_select='rerun', selection_mode='single-row',
                         column_config={'Τίτλος': st.column_config.TextColumn(width='large')})
    picked = event.selection.rows
    st.divider()
    if picked:
        show_card(res.iloc[picked[0]])
    else:
        st.info('Επιλέξτε ένα νομοσχέδιο από τη λίστα για να δείτε την κάρτα του.')

st.divider()
st.caption(f'Οι κάρτες φτιάχτηκαν αυτόματα με {MODEL_NAME} από τις Αναλύσεις Συνεπειών Ρύθμισης '
           'και μπορεί να περιέχουν λάθη. Για κάθε στοιχείο δίνεται η πηγή του· για το επίσημο '
           'κείμενο ανατρέξτε στη σελίδα της Βουλής.')
