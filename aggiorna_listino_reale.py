import pandas as pd
import sqlite3
import os

# --- CONFIGURAZIONE ---
NOME_FILE = 'listino_convertito.xlsx'
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

def aggiorna_db():
    db_path = get_db_path()
    if not db_path:
        print("❌ ERRORE: Database non trovato. Avvia prima 'py app.py'.")
        return

    print(f"📂 Lettura file listino: {NOME_FILE}...")
    try:
        df = pd.read_excel(NOME_FILE, engine='openpyxl')
    except Exception as e:
        print(f"❌ Errore lettura Excel: {e}")
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Mappiamo i prodotti esistenti per sapere se fare UPDATE o INSERT
    # Creiamo un dizionario: codice -> id_database
    print("📥 Lettura prodotti esistenti dal DB...")
    cursor.execute("SELECT codice, id FROM prodotto")
    map_prodotti_db = {row[0]: row[1] for row in cursor.fetchall()}

    print("🚀 Inizio aggiornamento...")

    cnt_aggiornati = 0
    cnt_inseriti = 0
    righe_totali = 0

    for index, row in df.iterrows():
        righe_totali += 1
        
        # Recupero dati dal file Excel
        codice = pulisci_codice(row.get('Codice'))
        nome = str(row.get('Nome Prodotto', '')).strip()
        
        try:
            prezzo = float(row.get('Prezzo_Listino', 0))
        except:
            prezzo = 0.0

        if not codice:
            continue

        # --- LOGICA CORE ---
        if codice in map_prodotti_db:
            # CASO 1: ESISTE GIA' -> Aggiorno SOLO il prezzo
            id_prod = map_prodotti_db[codice]
            cursor.execute("UPDATE prodotto SET prezzo = ? WHERE id = ?", (prezzo, id_prod))
            cnt_aggiornati += 1
        else:
            # CASO 2: NUOVO -> Inserisco tutto
            # Ingredienti vuoti (""), attivo = True
            cursor.execute("""
                INSERT INTO prodotto (codice, nome, prezzo, ingredienti, attivo) 
                VALUES (?, ?, ?, ?, ?)
            """, (codice, nome, prezzo, "", True))
            cnt_inseriti += 1

    conn.commit()
    conn.close()

    print("\n" + "="*40)
    print("🏆 AGGIORNAMENTO COMPLETATO")
    print("="*40)
    print(f"📄 Righe Excel lette:      {righe_totali}")
    print(f"🔄 Prezzi Aggiornati:      {cnt_aggiornati}")
    print(f"🆕 Nuovi Prodotti creati:  {cnt_inseriti}")
    print("="*40)
    print("✅ Il database è ora sincronizzato con il listino PDF.")

if __name__ == "__main__":
    aggiorna_db()