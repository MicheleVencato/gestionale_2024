import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# Percorso del file Excel
file_path = "gestionale_tessile_1_final.xlsx"

def carica_dati():
    try:
        clienti_df = pd.read_excel(file_path, sheet_name='Clienti')
        fornitori_df = pd.read_excel(file_path, sheet_name='Fornitori')
        prodotti_df = pd.read_excel(file_path, sheet_name='Prodotti')
        ordini_df = pd.read_excel(file_path, sheet_name='Ordini')
        fatture_df = pd.read_excel(file_path, sheet_name='Fatture')
        return clienti_df, fornitori_df, prodotti_df, ordini_df, fatture_df
    except Exception as e:
        st.error(f"Errore nel caricamento dei dati: {e}")
        return None, None, None, None, None

def salva_dati(df, sheet_name):
    try:
        # Carica il workbook esistente
        book = load_workbook(file_path)
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        st.success(f"Dati salvati con successo nel foglio '{sheet_name}'!")
    
    except PermissionError:
        st.error("Errore: il file è attualmente in uso da un altro programma. Chiudi il file e riprova.")
    
    except Exception as e:
        st.error(f"Si è verificato un errore durante il salvataggio del file Excel: {e}")

# Funzione per visualizzare i dati
def visualizza_dati(df, nome):
    if df is not None:
        st.write(f"### {nome}")
        st.dataframe(df)
    else:
        st.warning(f"Nessun dato disponibile per {nome}.")

# Funzione per inserire nuove fatture con gestione `KeyError`
def inserisci_fattura():
    st.write("## Inserisci Nuova Fattura")
    
    id_fattura = st.text_input("ID Fattura")
    ordine_selezionato = st.selectbox("Seleziona Ordine", ordini_df['ID Ordine'].dropna().unique() if 'ID Ordine' in ordini_df.columns else [])
    cliente_selezionato = st.selectbox("Seleziona Cliente", clienti_df['Nome'].dropna().unique() if 'Nome' in clienti_df.columns else [])
    fornitore_selezionato = st.selectbox("Seleziona Fornitore", fornitori_df['Nome Fornitore'].dropna().unique() if 'Nome Fornitore' in fornitori_df.columns else [])
    data_fattura = st.date_input("Data Fattura")
    data_scadenza = st.date_input("Data Scadenza")
    data_incasso = st.date_input("Data Incasso")
    metodo_pagamento = st.text_input("Metodo di Pagamento")
    
    # Debug: Visualizza le colonne disponibili in ordini_df
    st.write("Colonne disponibili in 'ordini_df':", list(ordini_df.columns))

    # Controllo per il calcolo dell'importo e gestione del KeyError
    if 'Importo Imponibile (€)' in ordini_df.columns:
        importo = ordini_df.loc[ordini_df['ID Ordine'] == ordine_selezionato, 'Importo Imponibile (€)'].values
    else:
        st.error("Colonna 'Importo Imponibile (€)' non trovata nel foglio Ordini. Verifica il file Excel.")
        importo = [0.0]

    importo = importo[0] if len(importo) > 0 else 0.0

    if 'IVA (%)' in ordini_df.columns:
        iva = ordini_df.loc[ordini_df['ID Ordine'] == ordine_selezionato, 'IVA (%)'].values
    else:
        st.error("Colonna 'IVA (%)' non trovata nel foglio Ordini. Verifica il file Excel.")
        iva = [0.0]

    iva = iva[0] if len(iva) > 0 else 0.0

    totale = importo + (importo * iva / 100)
    
    st.write(f"Importo (€): {importo:.2f}")
    st.write(f"IVA (%): {iva:.2f}")
    st.write(f"Totale (€): {totale:.2f}")
    incassata = st.checkbox("Incassata?")
    
    if st.button("Aggiungi Fattura"):
        nuova_fattura = pd.DataFrame({
            "ID Fattura": [id_fattura],
            "ID Ordine": [ordine_selezionato],
            "Cliente": [cliente_selezionato],
            "Fornitore": [fornitore_selezionato],
            "Data Fattura": [data_fattura],
            "Data Scadenza": [data_scadenza],
            "Data Incasso": [data_incasso],
            "Metodo di Pagamento": [metodo_pagamento],
            "Importo (€)": [importo],
            "IVA (%)": [iva],
            "Totale (€)": [totale],
            "Incassata": [incassata]
        })
        global fatture_df
        fatture_df = pd.concat([fatture_df, nuova_fattura], ignore_index=True)
        salva_dati(fatture_df, 'Fatture')

# Inizio dell'applicazione
st.title("Gestione Tessile")

# Carica i dati una volta all'avvio
clienti_df, fornitori_df, prodotti_df, ordini_df, fatture_df = carica_dati()

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
