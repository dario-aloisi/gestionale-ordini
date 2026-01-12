import pandas as pd
import sqlite3
import os

# --- CONFIGURAZIONE ---
NOME_FILE = 'listino_convertito.xlsx' # Il tuo file Excel pulito
DB_NAME = 'gestionale.db'

def get_db_path():
    if os.path.exists(DB_NAME): return DB_NAME
    elif os.path.exists(os.path.join('instance', DB_NAME)): return os.path.join('instance', DB_NAME)
    else: return None

def pulisci_codice(valore):
    """Pulisce il codice da spazi e formati strani (es. 100.0 diventa 100)"""
    try:
        if pd.isna(valore): return None
        if isinstance(valore, float): valore = int(valore)
        return str(valore).strip()
    except:
        return str(valore).strip()

def analisi_listino():
    db_path = get_db_path()
    if not db_path:
        print("❌ ERRORE: Database non trovato. Avvia prima 'py app.py'.")
        return

    print(f"📂 Lettura del file listino: {NOME_FILE}...")
    
    try:
        # Legge il file Excel
        df = pd.read_excel(NOME_FILE, engine='openpyxl')
        # Se per caso usi il CSV, scommenta la riga sotto e commenta quella sopra:
        # df = pd.read_csv('listino_convertito.csv')
    except Exception as e:
        print(f"❌ Errore lettura file: {e}")
        print("   Assicurati che il file esista e sia chiuso.")
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 1. Recuperiamo tutti i prodotti attuali dal DB
    print("📥 Recupero dati dal Database...")
    cursor.execute("SELECT codice, nome, prezzo FROM prodotto")
    # Creiamo un dizionario: codice -> {nome, prezzo}
    db_prodotti = {row[0]: {'nome': row[1], 'prezzo': row[2]} for row in cursor.fetchall()}
    
    # 2. Iniziamo l'analisi
    print("🔍 Analisi confronto in corso...")
    
    totale_excel = 0
    prodotti_nuovi = []      # Lista di (codice, nome, prezzo)
    prodotti_presenti = 0
    conflitti_nome = []      # Lista di (codice, nome_db, nome_excel)
    
    # Iteriamo sul file Excel
    for index, row in df.iterrows():
        totale_excel += 1
        
        # Mappiamo le colonne in base al tuo file
        # Assumiamo che le colonne si chiamino: "Codice", "Nome Prodotto", "Prezzo_Listino"
        cod_excel = pulisci_codice(row.get('Codice'))
        nome_excel = str(row.get('Nome Prodotto', '')).strip()
        
        try:
            prezzo_excel = float(row.get('Prezzo_Listino', 0))
        except:
            prezzo_excel = 0.0

        if not cod_excel:
            continue

        # --- LOGICA DI CONFRONTO ---
        if cod_excel not in db_prodotti:
            # CASO 1: Il prodotto non esiste nel DB
            prodotti_nuovi.append({
                'cod': cod_excel, 
                'nome': nome_excel, 
                'prezzo': prezzo_excel
            })
        else:
            # CASO 2: Il prodotto esiste già
            prodotti_presenti += 1
            dati_db = db_prodotti[cod_excel]
            
            # Controllo se il nome è diverso (Case insensitive)
            if dati_db['nome'].strip().lower() != nome_excel.strip().lower():
                conflitti_nome.append({
                    'cod': cod_excel,
                    'db_nome': dati_db['nome'],
                    'ex_nome': nome_excel
                })

    conn.close()

    # --- STAMPA DEL REPORT ---
    print("\n" + "="*50)
    print("📊 REPORT ANALISI LISTINO vs DATABASE")
    print("="*50)
    
    print(f"📄 Prodotti totali nel file Excel: {totale_excel}")
    print(f"✅ Prodotti GIÀ PRESENTI nel DB:   {prodotti_presenti}")
    print(f"🆕 Prodotti NUOVI (non nel DB):    {len(prodotti_nuovi)}")
    
    print("-" * 50)
    
    # 1. Analisi Nomi Differenti
    if len(conflitti_nome) > 0:
        print(f"⚠️  ATTENZIONE: {len(conflitti_nome)} Prodotti hanno lo STESSO CODICE ma NOME DIVERSO.")
        print("   (Ecco i primi 10 casi come esempio):")
        for c in conflitti_nome[:10]:
            print(f"   • [{c['cod']}]")
            print(f"     DB:    {c['db_nome']}")
            print(f"     EXCEL: {c['ex_nome']}")
            print("     ---")
        if len(conflitti_nome) > 10:
            print(f"   ...e altri {len(conflitti_nome) - 10} casi.")
    else:
        print("✨ Nessun conflitto di nomi rilevato sui prodotti esistenti.")

    # 2. Anteprima Nuovi Prodotti
    if len(prodotti_nuovi) > 0:
        print("-" * 50)
        print(f"🆕 Esempi di prodotti che verrebbero aggiunti ({len(prodotti_nuovi)} tot):")
        for p in prodotti_nuovi[:5]:
            print(f"   • [{p['cod']}] {p['nome']} (€ {p['prezzo']})")
        if len(prodotti_nuovi) > 5:
            print("   ...")

    print("="*50)
    print("🏁 Analisi terminata. Nessuna modifica applicata.")

if __name__ == "__main__":
    analisi_listino()