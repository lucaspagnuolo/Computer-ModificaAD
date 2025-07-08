import streamlit as st
import csv
import pandas as pd
import io

# Funzioni utili

def genera_samaccountname(nome, cognome, secondo_nome="", secondo_cognome="", esterno=False) -> str:
    n, sn = nome.strip().lower(), secondo_nome.strip().lower()
    c, sc = cognome.strip().lower(), secondo_cognome.strip().lower()
    suffix = ".ext" if esterno else ""
    limit = 16 if esterno else 20
    cand1 = f"{n}{sn}.{c}{sc}"
    if len(cand1) <= limit:
        return cand1 + suffix
    cand2 = f"{n[:1]}{sn[:1]}.{c}{sc}"
    if len(cand2) <= limit:
        return cand2 + suffix
    base = f"{n[:1]}{sn[:1]}.{c}"
    return base[:limit] + suffix


def build_full_name(cognome, secondo_cognome, nome, secondo_nome, esterno=False) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    full = " ".join(parts)
    return full + (" (esterno)" if esterno else "")

# Configurazione pagina
st.set_page_config(page_title="Genera CSV Computer")
st.title("Genera CSV Computer")

# Input necessari
cognome = st.text_input("Cognome").strip().capitalize()
nome = st.text_input("Nome").strip().capitalize()
secondo_cognome = st.text_input("Secondo Cognome").strip().capitalize()
secondo_nome = st.text_input("Secondo Nome").strip().capitalize()
numero_telefono = st.text_input("Mobile (+39 già inserito)", "").replace(" ", "")
description = st.text_input("PC (nome del computer)", "").strip()

# Generazione CSV
if st.button("Genera CSV Computer"):
    # Costruisci SAM e CN
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""
    comp = description

    # Definisci header e riga
    comp_header = [
        "Computer", "OU", "add_mail", "remove_mail",
        "add_mobile", "remove_mobile",
        "add_userprincipalname", "remove_userprincipalname",
        "disable", "moveToOU"
    ]
    comp_row = [
        comp,
        "",  # OU
        f"{sAM}@consip.it",
        "",  # remove_mail
        f"\"{mobile}\"",  # add_mobile
        "",  # remove_mobile
        f"\"{cn}\"",  # add_userprincipalname
        "",  # remove_userprincipalname
        "",  # disable
        ""   # moveToOU
    ]

    # Anteprima
    st.subheader("Anteprima CSV Computer")
    df_comp = pd.DataFrame([comp_row], columns=comp_header)
    st.dataframe(df_comp)

    # Generazione e download
    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar="\\")
    writer.writerow(comp_header)
    writer.writerow(comp_row)
    buf.seek(0)

    st.download_button(
        label="📥 Scarica CSV Computer",
        data=buf.getvalue(),
        file_name=f"{cognome}_{nome[:1]}_computer.csv",
        mime="text/csv"
    )
    st.success("✅ File CSV Computer generato correttamente")
