import streamlit as st
import pandas as pd

# Carica i dati dal file Excel
file_path = "gestionale_tessile_1_final.xlsx"  # Cambia con il percorso corretto
excel_data = pd.ExcelFile(file_path)

# Carico i fogli
clienti_df = pd.read_excel(file_path, sheet_name='Clienti')
fornitori_df = pd.read_excel(file_path, sheet_name='Fornitori')
prodotti_df = pd.read_excel(file_path, sheet_name='Prodotti')
ordini_df = pd.read_excel(file_path, sheet_name='Ordini')
fatture_df = pd.read_excel(file_path, sheet_name='Fatture')

# Funzione per visualizzare i dati
def visualizza_dati(df, nome):
    st.write(f"### {nome}")
    st.dataframe(df)

# Funzione per inserire nuovi dati
def inserisci_cliente():
    st.write("## Inserisci Nuovo Cliente")
    id_cliente = st.text_input("ID Cliente")
    nome_cliente = st.text_input("Nome Cliente")
    indirizzo = st.text_input("Indirizzo")
    persona_contatto = st.text_input("Persona di Contatto")
    telefono = st.text_input("Telefono")
    email = st.text_input("Email")
    partita_iva = st.text_input("Partita IVA")
    codice_sdi = st.text_input("Codice SDI")
    
    if st.button("Aggiungi Cliente"):
        nuovo_cliente = {
            "ID Cliente": id_cliente,
            "Nome": nome_cliente,
            "Indirizzo": indirizzo,
            "Persona di contatto": persona_contatto,
            "Telefono": telefono,
            "Email": email,
            "Partita IVA": partita_iva,
            "codice SDI": codice_sdi
        }
        clienti_df.loc[len(clienti_df)] = nuovo_cliente
        st.success("Cliente aggiunto con successo!")
        # Codice per salvare su Excel o database può essere aggiunto qui

# Inizio dell'applicazione
st.title("Gestione Tessile")

# Seleziona la sezione
sezione = st.sidebar.selectbox(
    "Seleziona una sezione",
    ["Clienti", "Fornitori", "Prodotti", "Ordini", "Fatture", "Aggiungi Cliente"]
)

# Visualizzazione dei dati
if sezione == "Clienti":
    visualizza_dati(clienti_df, "Clienti")
elif sezione == "Fornitori":
    visualizza_dati(fornitori_df, "Fornitori")
elif sezione == "Prodotti":
    visualizza_dati(prodotti_df, "Prodotti")
elif sezione == "Ordini":
    visualizza_dati(ordini_df, "Ordini")
elif sezione == "Fatture":
    visualizza_dati(fatture_df, "Fatture")
elif sezione == "Aggiungi Cliente":
    inserisci_cliente()
