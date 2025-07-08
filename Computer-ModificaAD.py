import streamlit as st
import csv
import pandas as pd
import io

# Configurazione pagina
st.set_page_config(page_title="Genera CSV Computer e Utente")
st.title("Genera CSV Computer e Utente")

# Input comuni
utenza = st.text_input("Indicare Utenza (SamAccountName es. nome.cognome)").strip().lower()
config_file = st.file_uploader(
    "Caricare file Est_Dati (Excel)",
    type=["xlsx", "xls"],
    help="File con colonne: SamAccountName, UserPrincipalName, Name, Mobile"
)

# Input: nome del computer (per CSV Computer)
computer = st.text_input("Computer (nome del computer)").strip()
# Input: dati utente (per CSV Utente)
sam_account_user = st.text_input("sAMAccountName Utente").strip()
# Description per CSV Utente viene preso dal campo computer

# Validazione upload
def validate_inputs(require_user=False):
    if not utenza or not config_file:
        st.warning("Per favore inserisci l'utenza e carica il file Est_Dati per procedere.")
        return False
    if not computer:
        st.warning("Per favore indica il nome del computer.")
        return False
    if require_user and not sam_account_user:
        st.warning("Per favore indica lo sAMAccountName per il CSV Utente.")
        return False
    return True

# Lettura dati Excel
def load_df():
    try:
        df = pd.read_excel(io.BytesIO(config_file.read()), engine='openpyxl')
    except Exception as e:
        st.error(f"Errore nel caricamento del file: {e}. Installa openpyxl.")
        return None
    return df

# Controllo colonne
required_cols = ["SamAccountName", "UserPrincipalName", "Name", "Mobile"]

def filter_record(df):
    if not all(col in df.columns for col in required_cols):
        st.error(f"Il file deve contenere le colonne: {', '.join(required_cols)}.")
        return None
    row = df[df["SamAccountName"].str.lower().str.strip() == utenza]
    if row.empty:
        st.error(f"Utenza '{utenza}' non trovata in Est_Dati (SamAccountName).")
        return None
    return row.iloc[0]

# Carica e filtra
if config_file and utenza:
    df = load_df()
    record = filter_record(df) if df is not None else None
else:
    record = None

# Sezione CSV Computer
tab1, tab2 = st.tabs(["CSV Computer", "CSV Utente"])

with tab1:
    if st.button("Genera CSV Computer"):
        if not validate_inputs():
            st.stop()
        if record is None:
            st.stop()
        mail = record["UserPrincipalName"]
        cn = record["Name"]
        mobile = record["Mobile"]
        # Costruisci CSV Computer
        header_comp = ["Computer","OU","add_mail","remove_mail","add_mobile","remove_mobile","add_userprincipalname","remove_userprincipalname","disable","moveToOU"]
        row_comp = [computer, "", mail, "", f"\"{mobile}\"", "", f"\"{cn}\"", "", "", ""]
        st.subheader("Anteprima CSV Computer")
        st.dataframe(pd.DataFrame([row_comp], columns=header_comp))
        buf = io.StringIO()
        writer = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar="\\")
        writer.writerow(header_comp)
        writer.writerow(row_comp)
        buf.seek(0)
        st.download_button("📥 Scarica CSV Computer", buf.getvalue(), file_name=f"{utenza}_computer.csv", mime="text/csv")
        st.success("✅ File CSV Computer generato")

with tab2:
    if st.button("Genera CSV Utente"):
        if not validate_inputs(require_user=True):
            st.stop()
        # Costruisci CSV Utente
        header_user = ['sAMAccountName','Creation','OU','Name','DisplayName','cn','GivenName','Surname','employeeNumber','employeeID','department','Description','passwordNeverExpired','ExpireDate','userprincipalname','mail','mobile','RimozioneGruppo','InserimentoGruppo','disable','moveToOU','telephoneNumber','company']
        desc = computer
        row_user = [sam_account_user,'SI','','','','','','','',sam_account_user,'','',desc,'','',sam_account_user+'@consip.it',sam_account_user+'@consip.it','','','',sam_account_user+'@consip.it','','','']
        st.subheader("Anteprima CSV Utente")
        st.dataframe(pd.DataFrame([row_user], columns=header_user))
        buf2 = io.StringIO()
        writer2 = csv.writer(buf2, quoting=csv.QUOTE_NONE, escapechar="\\")
        writer2.writerow(header_user)
        writer2.writerow(row_user)
        buf2.seek(0)
        st.download_button("📥 Scarica CSV Utente", buf2.getvalue(), file_name=f"{sam_account_user}_utente.csv", mime="text/csv")
        st.success("✅ File CSV Utente generato")
