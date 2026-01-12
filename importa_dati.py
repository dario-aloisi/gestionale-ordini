import pandas as pd
import sqlite3
import os

# --- CONFIGURAZIONE ---
NOME_FILE = 'listino.ods'       
DB_NAME = 'gestionale.db'       

def get_db_path():
    if os.path.exists(DB_NAME):
        return DB_NAME
    elif os.path.exists(os.path.join('instance', DB_NAME)):
        return os.path.join('instance', DB_NAME)
    else:
        print("❌ ERRORE CRITICO: Database non trovato! Avvia prima 'py app.py'.")
        return None

def pulisci_codice(valore):
    try:
        if pd.isna(valore): return None
        valore = int(valore) # Toglie .0
    except:
        pass
    return str(valore).strip()

def pulisci_nome(valore):
    if pd.isna(valore): return None
    return str(valore).replace('*', '').strip()

def importa_tutto():
    db_path = get_db_path()
    if not db_path: return

    print(f"📂 Apro il file: {NOME_FILE}")
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # ==========================================
    # 1. IMPORTAZIONE PRODOTTI
    # ==========================================
    print("\n" + "="*40)
    print("--- INIZIO IMPORT PRODOTTI ---")
    print("="*40)
    
    try:
        df_prod = pd.read_excel(NOME_FILE, engine='odf', sheet_name='Prodotti', header=None)
        
        count_ok = 0
        count_bad = 0
        count_dup = 0

        for index, row in df_prod.iterrows():
            excel_row_num = index + 1 # Numero riga Excel (perché parte da 1)
            
            raw_codice = row[0]
            raw_nome = row[1]

            codice = pulisci_codice(raw_codice)
            nome = pulisci_nome(raw_nome)

            # A. CONTROLLO DATI SPORCHI
            if not codice or not nome:
                print(f"❌ [RIGA {excel_row_num}] SCARTATA (Dati mancanti/vuoti)")
                print(f"   Originale: '{raw_codice}' - '{raw_nome}'")
                count_bad += 1
                continue 

            try:
                # B. TENTATIVO INSERIMENTO
                cursor.execute("""
                    INSERT OR IGNORE INTO prodotto (codice, nome, ingredienti, prezzo, attivo)
                    VALUES (?, ?, ?, 0.0, 1)
                """, (codice, nome, ""))
                
                if cursor.rowcount > 0:
                    count_ok += 1
                    # print(f"✅ [RIGA {excel_row_num}] Inserito: {codice} - {nome}")
                else:
                    count_dup += 1
                    print(f"⚠️ [RIGA {excel_row_num}] GIÀ PRESENTE (Saltato): {codice} - {nome}")

            except Exception as e:
                print(f"🔥 [RIGA {excel_row_num}] ERRORE DB: {e}")

        print("-" * 30)
        print(f"PRODOTTI TOTALI: {len(df_prod)}")
        print(f"✅ Inseriti: {count_ok}")
        print(f"⚠️ Doppioni: {count_dup}")
        print(f"❌ Scartati: {count_bad}")


    except ValueError:
        print("⚠️ Foglio 'Prodotti' non trovato.")
    except Exception as e:
        print(f"❌ Errore lettura Prodotti: {e}")


    # ==========================================
    # 2. IMPORTAZIONE CLIENTI
    # ==========================================
    print("\n" + "="*40)
    print("--- INIZIO IMPORT CLIENTI ---")
    print("="*40)

    try:
        df_cli = pd.read_excel(NOME_FILE, engine='odf', sheet_name='Clienti', header=None)
        
        count_ok_c = 0
        count_bad_c = 0
        count_dup_c = 0

        for index, row in df_cli.iterrows():
            excel_row_num = index + 1
            
            raw_codice = row[0]
            raw_nome = row[1]

            codice = pulisci_codice(raw_codice)
            nome = pulisci_nome(raw_nome)

            if not codice or not nome:
                print(f"❌ [RIGA {excel_row_num}] SCARTATA (Dati mancanti)")
                print(f"   Originale: '{raw_codice}' - '{raw_nome}'")
                count_bad_c += 1
                continue

            try:
                cursor.execute("""
                    INSERT OR IGNORE INTO cliente (codice, nome, note, attivo)
                    VALUES (?, ?, ?, 1)
                """, (codice, nome, ""))
                
                if cursor.rowcount > 0:
                    count_ok_c += 1
                else:
                    count_dup_c += 1
                    print(f"⚠️ [RIGA {excel_row_num}] GIÀ PRESENTE (Saltato): {codice} - {nome}")

            except Exception as e:
                print(f"🔥 [RIGA {excel_row_num}] ERRORE DB: {e}")

        print("-" * 30)
        print(f"CLIENTI TOTALI: {len(df_cli)}")
        print(f"✅ Inseriti: {count_ok_c}")
        print(f"⚠️ Doppioni: {count_dup_c}")
        print(f"❌ Scartati: {count_bad_c}")

    except ValueError:
        print("⚠️ Foglio 'Clienti' non trovato.")
    except Exception as e:
        print(f"❌ Errore lettura Clienti: {e}")

    conn.commit()
    conn.close()
    print("\n🎉 FINE OPERAZIONI.")

if __name__ == "__main__":
    importa_tutto()

