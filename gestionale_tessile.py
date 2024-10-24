import streamlit as st
import pandas as pd
from openpyxl import load_workbook

# Percorso del file Excel
file_path = "gestionale_tessile_1_final.xlsx"

# Funzione per caricare i dati dal file Excel
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

# Funzione per salvare i dati nel file Excel
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

# Funzione per modificare i dati esistenti
def modifica_dati(df, sheet_name):
    st.write(f"## Modifica Dati - {sheet_name}")
    if df is not None:
        indice = st.number_input("Seleziona l'indice del record da modificare", min_value=0, max_value=len(df)-1, step=1)
        if st.button("Carica Record"):
            record = df.iloc[indice]
            modifiche = {}
            for col in df.columns:
                modifiche[col] = st.text_input(f"{col}", value=str(record[col]))
            if st.button("Salva Modifiche"):
                for col, value in modifiche.items():
                    df.at[indice, col] = value
                salva_dati(df, sheet_name)
    else:
        st.warning(f"Nessun dato disponibile per {sheet_name}.")

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
            "Codice SDI": [codice_sdi]
        })
        global clienti_df
        clienti_df = pd.concat([clienti_df, nuovo_cliente], ignore_index=True)
        salva_dati(clienti_df, 'Clienti')

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

# Funzione per inserire nuovi prodotti
def inserisci_prodotto():
    st.write("## Inserisci Nuovo Prodotto")
    id_prodotto = st.text_input("ID Prodotto")
    fornitore_selezionato = st.selectbox("Seleziona Fornitore", fornitori_df['Nome Fornitore'].dropna().unique() if 'Nome Fornitore' in fornitori_df.columns else [])
    descrizione = st.text_input("Descrizione Prodotto")
    unita_misura = st.text_input("Unità di Misura")
    note = st.text_input("Note")

    if st.button("Aggiungi Prodotto"):
        nuovo_prodotto = pd.DataFrame({
            "ID Prodotto": [id_prodotto],
            "ID Fornitore": [fornitore_selezionato],
            "Nome Fornitore (Automatico)": [fornitore_selezionato],
            "Descrizione": [descrizione],
            "Unità di misura": [unita_misura],
            "Note": [note]
        })
        global prodotti_df
        prodotti_df = pd.concat([prodotti_df, nuovo_prodotto], ignore_index=True)
        salva_dati(prodotti_df, 'Prodotti')

# Funzione per inserire nuovi ordini
def inserisci_ordine():
    st.write("## Inserisci Nuovo Ordine")
    
    id_ordine = st.text_input("ID Ordine")
    id_cliente = st.text_input("ID Cliente")
    nome_cliente = st.text_input("Nome Cliente")
    id_fornitore = st.text_input("ID Fornitore")
    nome_fornitore = st.text_input("Nome Fornitore")
    id_prodotto = st.text_input("ID Prodotto")
    descrizione_prodotto = st.text_input("Descrizione Prodotto")
    prezzo_per_kg = st.number_input("Prezzo per kg (€)", min_value=0.0, format="%.2f")
    quantita_ordinata = st.number_input("Quantità Ordinata (kg)", min_value=0.0, format="%.2f")
    data_ordine = st.date_input("Data Ordine")
    data_consegna_prevista = st.date_input("Data Consegna Prevista")
    numero_ordine_cliente = st.text_input("Numero Ordine Cliente")
    numero_ordine_fornitore = st.text_input("Numero Ordine Fornitore")
    quantita_consegnata_totale = st.number_input("Quantità Consegnata Totale (kg)", min_value=0.0, format="%.2f")
    quantita_rimanente = st.number_input("Quantità Rimanente (kg)", min_value=0.0, format="%.2f")
    stato_ordine = st.text_input("Stato Ordine")
    note = st.text_input("Note")
    
    if st.button("Aggiungi Ordine"):
        nuovo_ordine = pd.DataFrame({
            "ID Ordine": [id_ordine],
            "ID Cliente": [id_cliente],
            "Nome Cliente (Automatico)": [nome_cliente],
            "ID Fornitore": [id_fornitore],
            "Nome Fornitore (Automatico)": [nome_fornitore],
            "ID Prodotto": [id_prodotto],
            "Descrizione Prodotto (Automatico)": [descrizione_prodotto],
            "Prezzo per kg (€)": [prezzo_per_kg],
            "Quantità Ordinata (kg)": [quantita_ordinata],
            "Data Ordine": [data_ordine],
            "Data Consegna Prevista": [data_consegna_prevista],
            "Numero Ordine Cliente": [numero_ordine_cliente],
            "Numero Ordine Fornitore": [numero_ordine_fornitore],
            "Quantità Consegnata Totale (kg)": [quantita_consegnata_totale],
            "Quantità Rimanente (kg)": [quantita_rimanente],
            "Stato Ordine": [stato_ordine],
            "Note": [note]
        })
        global ordini_df
        ordini_df = pd.concat([ordini_df, nuovo_ordine], ignore_index=True)
        salva_dati(ordini_df, 'Ordini')

# Funzione per inserire nuove fatture
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

    # Calcolo dell'importo e dell'IVA (se non presenti nel dataframe ordini)
    prezzo_unitario = ordini_df.loc[ordini_df['ID Ordine'] == ordine_selezionato, 'Prezzo per kg (€)'].values
    quantita = ordini_df.loc[ordini_df['ID Ordine'] == ordine_selezionato, 'Quantità Ordinata (kg)'].values

    prezzo_unitario = prezzo_unitario[0] if len(prezzo_unitario) > 0 else 0.0
    quantita = quantita[0] if len(quantita) > 0 else 0.0

    importo = prezzo_unitario * quantita
    iva = st.number_input("IVA (%)", min_value=0.0, max_value=100.0, value=22.0, format="%.2f")
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
    ["Clienti", "Fornitori", "Prodotti", "Ordini", "Fatture", "Aggiungi Cliente", "Aggiungi Fornitore", "Aggiungi Prodotto", "Aggiungi Ordine", "Aggiungi Fattura", "Modifica Cliente", "Modifica Fornitore", "Modifica Prodotto", "Modifica Ordine", "Modifica Fattura"]
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
elif sezione == "Modifica Cliente":
    modifica_dati(clienti_df, 'Clienti')
elif sezione == "Modifica Fornitore":
    modifica_dati(fornitori_df, 'Fornitori')
elif sezione == "Modifica Prodotto":
    modifica_dati(prodotti_df, 'Prodotti')
elif sezione == "Modifica Ordine":
    modifica_dati(ordini_df, 'Ordini')
elif sezione == "Modifica Fattura":
    modifica_dati(fatture_df, 'Fatture')
