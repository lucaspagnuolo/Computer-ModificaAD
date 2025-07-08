import streamlit as st
import csv
import pandas as pd
import io

# Configurazione pagina
st.set_page_config(page_title="Genera CSV Computer")
st.title("Genera CSV Computer")

# Input: indicare utenza e caricare dati
utenza = st.text_input("Indicare Utenza (SamAccountName es. nome.cognome)").strip().lower()
config_file = st.file_uploader(
    "Caricare file Est_Dati (Excel)",
    type=["xlsx", "xls"],
    help="File con colonne: SamAccountName, UserPrincipalName, Name, Mobile"
)
# Input: nome del computer
computer = st.text_input("Computer (nome del computer)", "").strip()

# Verifica input
if not utenza or not config_file:
    st.warning("Per favore inserisci l'utenza, carica il file Est_Dati e indica il nome del computer per procedere.")
    st.stop()

# Lettura dati Excel con engine esplicito
try:
    df = pd.read_excel(io.BytesIO(config_file.read()), engine='openpyxl')
except Exception as e:
    st.error(f"Errore nel caricamento del file: {e}. Assicurati di avere installato openpyxl.")
    st.stop()

# Verifica presenza colonne
required_cols = ["SamAccountName", "UserPrincipalName", "Name", "Mobile"]
if not all(col in df.columns for col in required_cols):
    st.error(f"Il file deve contenere le colonne: {', '.join(required_cols)}.")
    st.stop()

# Ricerca record per SamAccountName
row = df[df["SamAccountName"].str.lower().str.strip() == utenza]
if row.empty:
    st.error(f"Utenza '{utenza}' non trovata in Est_Dati (SamAccountName).")
    st.stop()
record = row.iloc[0]

# Generazione CSV
if st.button("Genera CSV Computer"):
    mail = record["UserPrincipalName"]
    cn = record["Name"]
    mobile = record["Mobile"]

    # Definisci header e riga
    comp_header = [
        "Computer", "OU", "add_mail", "remove_mail",
        "add_mobile", "remove_mobile",
        "add_userprincipalname", "remove_userprincipalname",
        "disable", "moveToOU"
    ]
    comp_row = [
        computer,
        "",  # OU
        mail,
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
        file_name=f"{utenza}_computer.csv",
        mime="text/csv"
    )
    st.success("✅ File CSV Computer generato correttamente")
