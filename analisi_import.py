import pandas as pd
import sqlite3
import os
import datetime

# --- CONFIGURAZIONE ---
NOME_FILE = 'tab_ag_15_cli_art_2025.xlsx'  # Nome del file Excel
FOGLIO_DA_LEGGERE = 'Scriptare'            # Nome del foglio pulito
DB_NAME = 'gestionale.db'

def get_db_path():
    if os.path.exists(DB_NAME): return DB_NAME
    elif os.path.exists(os.path.join('instance', DB_NAME)): return os.path.join('instance', DB_NAME)
    else: return None

def pulisci_codice(valore):
    try:
        if pd.isna(valore): return None
        if isinstance(valore, float): valore = int(valore)
        return str(valore).strip()
    except:
        return str(valore).strip()

def analisi_simulata():
    db_path = get_db_path()
    if not db_path:
        print("❌ ERRORE: Database non trovato. Avvia prima 'py app.py'.")
        return

    print(f"📂 Apro il file Excel: {NOME_FILE} (Foglio: {FOGLIO_DA_LEGGERE})...")
    
    # 1. LETTURA FILE EXCEL
    try:
        df = pd.read_excel(NOME_FILE, sheet_name=FOGLIO_DA_LEGGERE, engine='openpyxl')
    except ValueError as ve:
        print(f"❌ ERRORE: Non trovo il foglio '{FOGLIO_DA_LEGGERE}'.")
        return
    except Exception as e:
        print(f"❌ Errore lettura Excel: {e}")
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    clienti_trovati = {}   # codice -> nome
    prodotti_trovati = {}  # codice -> {nome, prezzo}
    ordini_trovati = {}    # chiave (SOLO DATA) -> lista righe

    print("🔍 Analisi righe in corso...")
    
    righe_totali = 0
    righe_scartate = 0

    for index, row in df.iterrows():
        righe_totali += 1
        try:
            # --- A. GESTIONE DATA ---
            raw_data = row.get('DataDoc')
            if isinstance(raw_data, datetime.datetime):
                data_iso = raw_data.strftime('%Y-%m-%d')
            else:
                data_iso = str(raw_data).split(' ')[0].strip()

            # --- B. GESTIONE CAMPI ---
            cli_cod = pulisci_codice(row.get('Cd_CF'))
            cli_nome = str(row.get('CF_Descrizione')).strip()
            
            prod_cod = pulisci_codice(row.get('Cd_AR'))
            prod_nome = str(row.get('DORig_Descrizione')).strip()
            
            try: qta = int(row.get('Qta', 0))
            except: qta = 0

            prezzo_raw = row.get('PrezzoUnitarioV', 0)
            try:
                if isinstance(prezzo_raw, str): prezzo = float(prezzo_raw.replace(',', '.'))
                else: prezzo = float(prezzo_raw)
            except: prezzo = 0.0

            # --- C. CONTROLLI ---
            if not cli_cod or not prod_cod or cli_cod == 'None' or prod_cod == 'None':
                righe_scartate += 1
                continue

            # --- D. SALVATAGGIO ---
            clienti_trovati[cli_cod] = cli_nome
            prodotti_trovati[prod_cod] = {'nome': prod_nome, 'prezzo': prezzo}

            # MODIFICA QUI: Raggruppiamo solo per DATA
            chiave_ordine = data_iso 
            
            if chiave_ordine not in ordini_trovati:
                ordini_trovati[chiave_ordine] = []
            
            ordini_trovati[chiave_ordine].append({
                'cli_cod': cli_cod, # Salviamo chi è il cliente per info
                'prod_cod': prod_cod,
                'prod_nome': prod_nome,
                'qta': qta,
                'prezzo': prezzo
            })

        except Exception as e:
            righe_scartate += 1

    print(f"✅ Righe analizzate: {righe_totali}")
    print(f"🗑️ Righe ignorate: {righe_scartate}")

    # =================================================================
    # REPORT
    # =================================================================
    
    print("\n" + "="*50)
    print("📊 REPORT ANALISI")
    print("="*50)

    # 1. CLIENTI
    cursor.execute("SELECT codice, nome FROM cliente")
    db_clienti = {row[0]: row[1] for row in cursor.fetchall()}
    
    nuovi_clienti = 0
    clienti_conflitto = [] 

    for cod, nome in clienti_trovati.items():
        if cod not in db_clienti:
            nuovi_clienti += 1
        else:
            if db_clienti[cod] != nome:
                clienti_conflitto.append((cod, db_clienti[cod], nome))

    print(f"\n👥 CLIENTI:")
    print(f"   - Totali nel file: {len(clienti_trovati)}")
    print(f"   - NUOVI: {nuovi_clienti}")
    
    print("\n   --- ELENCO COMPLETO CLIENTI NEL FILE ---")
    for cod, nome in sorted(clienti_trovati.items(), key=lambda item: item[1]):
        stato = "🆕 (Nuovo)" if cod not in db_clienti else "✅ (Già presente)"
        print(f"   • [{cod}] {nome} {stato}")
    print("   ---------------------------------------")

    if clienti_conflitto:
        print(f"\n⚠️  NOMI DIVERSI ({len(clienti_conflitto)} casi):")
        for c in clienti_conflitto: 
            print(f"   • {c[0]}: DB='{c[1]}' <--> FILE='{c[2]}'")

    # 2. PRODOTTI
    cursor.execute("SELECT codice, nome FROM prodotto")
    db_prodotti = {row[0]: row[1] for row in cursor.fetchall()}

    nuovi_prodotti = 0
    prodotti_conflitto = []

    for cod, dati in prodotti_trovati.items():
        nome_file = dati['nome']
        if cod not in db_prodotti:
            nuovi_prodotti += 1
        else:
            if db_prodotti[cod].strip().lower() != nome_file.strip().lower():
                prodotti_conflitto.append((cod, db_prodotti[cod], nome_file))

    print(f"\n📦 PRODOTTI:")
    print(f"   - Totali nel file: {len(prodotti_trovati)}")
    print(f"   - NUOVI: {nuovi_prodotti}")

    # --- STAMPA ELENCO PRODOTTI NUOVI ---
    if nuovi_prodotti > 0:
        print("\n   --- ELENCO PRODOTTI NUOVI (Non nel DB) ---")
        for cod, dati in sorted(prodotti_trovati.items(), key=lambda x: x[1]['nome']):
            if cod not in db_prodotti:
                print(f"   • [{cod}] {dati['nome']}")
        print("   ------------------------------------------")

    if prodotti_conflitto:
        print(f"\n⚠️  NOMI DIVERSI ({len(prodotti_conflitto)} casi):")
        print("   (Legenda: DB = Nome attuale nel tuo programma | FILE = Nome nel file Excel)")
        for p in prodotti_conflitto:
            print(f"   • Cod. {p[0]}:")
            print(f"     DB:   {p[1]}")
            print(f"     FILE: {p[2]}")
            print("     ---")

    # 3. ORDINI
    print(f"\n📄 ORDINI DA CREARE:")
    print(f"   - Ordini totali (Giorni di lavoro): {len(ordini_trovati)}")
    
    fatturato = sum(r['qta'] * r['prezzo'] for ord_list in ordini_trovati.values() for r in ord_list)
    print(f"   - Fatturato storico totale: € {fatturato:,.2f}")

    conn.close()
    print("\n🏁 Fine simulazione.")

if __name__ == "__main__":
    analisi_simulata()