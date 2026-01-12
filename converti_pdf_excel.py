import pdfplumber
import pandas as pd
import re
import os

# --- CONFIGURAZIONE ---
NOME_FILE_PDF = 'Listino articoli al 14102025.pdf'
NOME_FILE_EXCEL = 'listino_convertito.xlsx'

def pulisci_testo(testo):
    """Rimuove a capo e spazi extra"""
    if testo:
        return str(testo).replace('\n', ' ').strip()
    return ""

def pulisci_prezzo(testo):
    """Trasforma '33,00 €' in 33.0"""
    if not testo: return 0.0
    # Rimuove tutto ciò che non è numero o virgola
    pulito = re.sub(r'[^\d,]', '', str(testo))
    try:
        return float(pulito.replace(',', '.'))
    except:
        return 0.0

def converti_pdf():
    if not os.path.exists(NOME_FILE_PDF):
        print(f"❌ Errore: Non trovo il file '{NOME_FILE_PDF}'")
        return

    print(f"📄 Apertura file: {NOME_FILE_PDF}...")
    
    dati_estratti = []
    
    with pdfplumber.open(NOME_FILE_PDF) as pdf:
        tot_pagine = len(pdf.pages)
        print(f"   Trovate {tot_pagine} pagine.")

        for i, pagina in enumerate(pdf.pages):
            print(f"   Elaborazione pagina {i+1}/{tot_pagine}...", end='\r')
            
            # Estrae la tabella
            table = pagina.extract_table()
            
            if table:
                for riga in table:
                    # Saltiamo l'intestazione se la troviamo (controllando se c'è scritto "Cd AR")
                    # Nota: pdfplumber a volte spacca le righe, controlliamo che la riga abbia dati sensati
                    if not riga or not riga[0] or "Cd AR" in str(riga[0]):
                        continue
                    
                    # Mapping colonne basato sul tuo PDF:
                    # 0: Codice (Cd AR)
                    # 1: Descrizione
                    # 2: Prezzo Listino (LSArticolo Prezzo)
                    # 3: Pezzi per cartone? (NR PZ)
                    # 4: Prezzo Unitario?
                    # ...
                    
                    # Prendiamo solo quello che ci serve
                    codice = pulisci_testo(riga[0])
                    descrizione = pulisci_testo(riga[1])
                    prezzo_listino = pulisci_prezzo(riga[2])
                    
                    # Colonna 3 sembra essere "Pezzi per Cartone" o simili
                    # Colonna 4 sembra il prezzo unitario scontato o netto
                    # Salviamo tutto per sicurezza
                    
                    if codice: # Se c'è un codice, è una riga valida
                        dati_estratti.append({
                            'Codice': codice,
                            'Descrizione': descrizione,
                            'Prezzo_Listino': prezzo_listino,
                            # Aggiungo colonne extra grezze se servono controlli
                            'Colonna_Extra_1': pulisci_testo(riga[3]) if len(riga) > 3 else "",
                            'Colonna_Extra_2': pulisci_testo(riga[4]) if len(riga) > 4 else ""
                        })

    print(f"\n✅ Estrazione completata. Trovati {len(dati_estratti)} articoli.")
    
    # Creiamo il DataFrame e salviamo in Excel
    df = pd.DataFrame(dati_estratti)
    
    print(f"💾 Salvataggio in {NOME_FILE_EXCEL}...")
    df.to_excel(NOME_FILE_EXCEL, index=False)
    
    print("🏆 Fatto! Apri il file Excel e controllalo.")

if __name__ == "__main__":
    converti_pdf()