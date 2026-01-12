import pandas as pd
import sqlite3
import os
import datetime

# --- CONFIGURAZIONE ---
NOME_FILE = 'tab_ag_15_cli_art_2025.xlsx'
FOGLIO_DA_LEGGERE = 'Scriptare'
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

def importa_dati():
    db_path = get_db_path()
    if not db_path:
        print("❌ ERRORE: Database non trovato. Avvia prima 'py app.py'.")
        return

    print(f"📂 Apro il file: {NOME_FILE}...")
    
    try:
        df = pd.read_excel(NOME_FILE, sheet_name=FOGLIO_DA_LEGGERE, engine='openpyxl')
    except Exception as e:
        print(f"❌ Errore lettura Excel: {e}")
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    print("🚀 Inizio Importazione (Raggruppamento per DATA)...")
    print("ℹ️  I Clienti NON verranno creati. Le righe di clienti mancanti verranno saltate.")
    
    # Contatori
    cnt_prodotti_nuovi = 0
    cnt_prodotti_aggiornati = 0
    cnt_ordini_creati = 0      # Ora corrisponderà ai GIORNI di lavoro
    cnt_righe_inserite = 0
    
    # Tracciamento errori
    clienti_mancanti_set = set()
    righe_saltate_per_cliente = 0

    # Cache ID (Codice -> ID)
    map_clienti = {}  
    map_prodotti = {} 

    # 1. Carico ID esistenti dal DB
    cursor.execute("SELECT codice, id FROM cliente")
    for r in cursor.fetchall(): map_clienti[r[0]] = r[1]

    cursor.execute("SELECT codice, id FROM prodotto")
    for r in cursor.fetchall(): map_prodotti[r[0]] = r[1]

    # Struttura temporanea: Chiave = DATA (Stringa), Valore = Lista di righe (con info cliente)
    ordini_giornalieri = {} 

    # --- FASE 1: Scansione Righe ---
    for index, row in df.iterrows():
        try:
            # Dati
            cli_cod = pulisci_codice(row.get('Cd_CF'))
            prod_cod = pulisci_codice(row.get('Cd_AR'))
            prod_nome = str(row.get('DORig_Descrizione')).strip()
            
            # Data (Chiave fondamentale per raggruppare)
            raw_data = row.get('DataDoc')
            if isinstance(raw_data, datetime.datetime):
                data_iso = raw_data.strftime('%Y-%m-%d')
            else:
                data_iso = str(raw_data).split(' ')[0].strip()

            # Numeri
            try: qta = int(row.get('Qta', 0))
            except: qta = 0
            
            prezzo_raw = row.get('PrezzoUnitarioV', 0)
            try:
                if isinstance(prezzo_raw, str): prezzo = float(prezzo_raw.replace(',', '.'))
                else: prezzo = float(prezzo_raw)
            except: prezzo = 0.0

            if not cli_cod or not prod_cod or cli_cod == 'None' or prod_cod == 'None':
                continue

            # --- CONTROLLO CLIENTE ---
            if cli_cod not in map_clienti:
                clienti_mancanti_set.add(cli_cod)
                righe_saltate_per_cliente += 1
                continue 

            # --- GESTIONE PRODOTTO ---
            if prod_cod not in map_prodotti:
                # NUOVO: Lo creo
                cursor.execute("INSERT INTO prodotto (codice, nome, ingredienti, prezzo, attivo) VALUES (?, ?, ?, ?, ?)", 
                               (prod_cod, prod_nome, "", prezzo, True))
                nuovo_id = cursor.lastrowid
                map_prodotti[prod_cod] = nuovo_id
                cnt_prodotti_nuovi += 1
            else:
                # ESISTENTE: Aggiorno SOLO il prezzo
                cursor.execute("UPDATE prodotto SET prezzo = ? WHERE id = ?", (prezzo, map_prodotti[prod_cod]))
                cnt_prodotti_aggiornati += 1

            # --- PREPARAZIONE ORDINE (Raggruppo SOLO per Data) ---
            if data_iso not in ordini_giornalieri:
                ordini_giornalieri[data_iso] = []
            
            # Aggiungo la riga alla data, ma mi ricordo CHI è il cliente di QUESTA riga
            ordini_giornalieri[data_iso].append({
                'cli_id': map_clienti[cli_cod], # Fondamentale!
                'prod_id': map_prodotti[prod_cod],
                'qta': qta,
                'prezzo_storico': prezzo
            })

        except Exception as e:
            print(f"⚠️ Errore riga {index}: {e}")

    conn.commit()
    print(f"✅ Prodotti processati: {cnt_prodotti_nuovi} Nuovi, {cnt_prodotti_aggiornati} Prezzi aggiornati.")

    # --- FASE 2: Creazione Ordini Giornalieri ---
    print(f"📦 Creazione di {len(ordini_giornalieri)} Ordini Giornalieri...")
    
    # Ordino per data così nel DB sono sequenziali
    for data_cons in sorted(ordini_giornalieri.keys()):
        righe = ordini_giornalieri[data_cons]
        
        # 1. Creo UN SOLO Ordine per questa data
        cursor.execute("""
            INSERT INTO ordine (data_consegna, stato, note, ora_creazione) 
            VALUES (?, ?, ?, ?)
        """, (data_cons, 'inviato', '', '00-00'))
        
        ordine_id = cursor.lastrowid
        cnt_ordini_creati += 1

        # 2. Inserisco tutte le righe (ognuna col suo cliente)
        for r in righe:
            cursor.execute("""
                INSERT INTO dettaglio_ordine (ordine_id, cliente_id, prodotto_id, quantita, prezzo_storico)
                VALUES (?, ?, ?, ?, ?)
            """, (ordine_id, r['cli_id'], r['prod_id'], r['qta'], r['prezzo_storico']))
            cnt_righe_inserite += 1

    conn.commit()
    conn.close()

    print("\n" + "="*40)
    print("🏆 IMPORTAZIONE COMPLETATA")
    print("="*40)
    print(f"📦 Prodotti Nuovi Creati:     {cnt_prodotti_nuovi}")
    print(f"💲 Prezzi Prodotti Aggiornati: {cnt_prodotti_aggiornati}")
    print(f"📅 Ordini (Giorni) Creati:    {cnt_ordini_creati}")
    print(f"📝 Righe Totali Inserite:     {cnt_righe_inserite}")
    print("-" * 40)
    
    if len(clienti_mancanti_set) > 0:
        print(f"⚠️ ATTENZIONE: Saltate {righe_saltate_per_cliente} righe totali.")
        print(f"   Causa: {len(clienti_mancanti_set)} Clienti non trovati nel DB.")
        print("   Codici Clienti Mancanti:")
        print("   " + ", ".join(sorted(list(clienti_mancanti_set))))
    else:
        print("✅ Tutti i clienti del file erano presenti nel DB.")
        
    print("="*40)

if __name__ == "__main__":
    importa_dati()