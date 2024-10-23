import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# Percorso del file Excel
file_path = "gestionale_tessile_1_final.xlsx"

def carica_dati():
    clienti_df = pd.read_excel(file_path, sheet_name='Clienti')
    fornitori_df = pd.read_excel(file_path, sheet_name='Fornitori')
    prodotti_df = pd.read_excel(file_path, sheet_name='Prodotti')
    ordini_df = pd.read_excel(file_path, sheet_name='Ordini')
    fatture_df = pd.read_excel(file_path, sheet_name='Fatture')
    return clienti_df, fornitori_df, prodotti_df, ordini_df, fatture_df

def salva_dati(df, sheet_name):
    # Carica il workbook esistente
    book = load_workbook(file_path)
    # Apri un ExcelWriter e assegna il workbook caricato
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        writer.book = book
        # Sovrascrivi solo il foglio specifico
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        writer.save()

# Carico i dati dal file Excel
clienti_df, fornitori_df, prodotti_df, ordini_df, fatture_df = carica_dati()

# Funzione per visualizzare i dati
def visualizza_dati(df, nome):
    st.write(f"### {nome}")
    st.dataframe(df)

# Funzione per inserire nuovi clienti
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
        nuovo_cliente = pd.DataFrame({
            "ID Cliente": [id_cliente],
            "Nome": [nome_cliente],
            "Indirizzo": [indirizzo],
            "Persona di contatto": [persona_contatto],
            "Telefono": [telefono],
            "Email": [email],
            "Partita IVA": [partita_iva],
            "codice SDI": [codice_sdi]
        })
        global clienti_df
        clienti_df = pd.concat([clienti_df, nuovo_cliente], ignore_index=True)
        salva_dati(clienti_df, 'Clienti')
        st.success("Cliente aggiunto con successo!")

# Funzione per inserire nuovi fornitori
def inserisci_fornitore():
    st.write("## Inserisci Nuovo Fornitore")
    id_fornitore = st.text_input("ID Fornitore")
    nome_fornitore = st.text_input("Nome Fornitore")
    indirizzo = st.text_input("Indirizzo")
    persona_contatto = st.text_input("Persona di Contatto")
    telefono = st.text_input("Telefono")
    email = st.text_input("Email")
    partita_iva = st.text_input("Partita IVA")
    
    if st.button("Aggiungi Fornitore"):
        nuovo_fornitore = pd.DataFrame({
            "ID Fornitore": [id_fornitore],
            "Nome Fornitore": [nome_fornitore],
            "Indirizzo": [indirizzo],
            "Persona di Contatto": [persona_contatto],
            "Telefono": [telefono],
            "Email": [email],
            "Partita IVA": [partita_iva]
        })
        global fornitori_df
        fornitori_df = pd.concat([fornitori_df, nuovo_fornitore], ignore_index=True)
        salva_dati(fornitori_df, 'Fornitori')
        st.success("Fornitore aggiunto con successo!")

# Funzione per inserire nuovi prodotti
def inserisci_prodotto():
    st.write("## Inserisci Nuovo Prodotto")
    id_prodotto = st.text_input("ID Prodotto")
    id_fornitore = st.text_input("ID Fornitore")
    descrizione = st.text_input("Descrizione Prodotto")
    unita_misura = st.text_input("Unità di Misura")
    
    if st.button("Aggiungi Prodotto"):
        nuovo_prodotto = pd.DataFrame({
            "ID Prodotto": [id_prodotto],
            "ID Fornitore": [id_fornitore],
            "Descrizione": [descrizione],
            "Unità di misura": [unita_misura]
        })
        global prodotti_df
        prodotti_df = pd.concat([prodotti_df, nuovo_prodotto], ignore_index=True)
        salva_dati(prodotti_df, 'Prodotti')
        st.success("Prodotto aggiunto con successo!")

# Funzione per inserire nuovi ordini
def inserisci_ordine():
    st.write("## Inserisci Nuovo Ordine")
    id_ordine = st.text_input("ID Ordine")
    id_cliente = st.text_input("ID Cliente")
    id_prodotto = st.text_input("ID Prodotto")
    importo = st.number_input("Importo Imponibile (€)", min_value=0.0, format="%.2f")
    iva = st.number_input("IVA (%)", min_value=0.0, max_value=100.0, format="%.2f")
    totale = st.number_input("Totale (€)", min_value=0.0, format="%.2f")
    
    if st.button("Aggiungi Ordine"):
        nuovo_ordine = pd.DataFrame({
            "ID Ordine": [id_ordine],
            "ID Cliente": [id_cliente],
            "ID Prodotto": [id_prodotto],
            "Importo Imponibile (€)": [importo],
            "IVA (%)": [iva],
            "Totale (€)": [totale]
        })
        global ordini_df
        ordini_df = pd.concat([ordini_df, nuovo_ordine], ignore_index=True)
        salva_dati(ordini_df, 'Ordini')
        st.success("Ordine aggiunto con successo!")

# Funzione per inserire nuove fatture
def inserisci_fattura():
    st.write("## Inserisci Nuova Fattura")
    id_fattura = st.text_input("ID Fattura")
    id_ordine = st.text_input("ID Ordine")
    data_fattura = st.date_input("Data Fattura")
    data_incasso = st.date_input("Data Incasso")
    incassata = st.checkbox("Incassata?")
    
    if st.button("Aggiungi Fattura"):
        nuova_fattura = pd.DataFrame({
            "ID Fattura": [id_fattura],
            "ID Ordine": [id_ordine],
            "Data Fattura": [data_fattura],
            "Data Incasso": [data_incasso],
            "Incassata": [incassata]
        })
        global fatture_df
        fatture_df = pd.concat([fatture_df, nuova_fattura], ignore_index=True)
        salva_dati(fatture_df, 'Fatture')
        st.success("Fattura aggiunta con successo!")

# Inizio dell'applicazione
st.title("Gestione Tessile")

# Seleziona la sezione
sezione = st.sidebar.selectbox(
    "Seleziona una sezione",
    ["Clienti", "Fornitori", "Prodotti", "Ordini", "Fatture", "Aggiungi Cliente", "Aggiungi Fornitore", "Aggiungi Prodotto", "Aggiungi Ordine", "Aggiungi Fattura"]
)

# Visualizzazione e inserimento dei dati
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
elif sezione == "Aggiungi Fornitore":
    inserisci_fornitore()
elif sezione == "Aggiungi Prodotto":
    inserisci_prodotto()
elif sezione == "Aggiungi Ordine":
    inserisci_ordine()
elif sezione == "Aggiungi Fattura":
    inserisci_fattura()
