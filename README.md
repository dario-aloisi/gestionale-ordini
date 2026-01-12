# 📦 Gestione Ordini & Magazzino (v2.0)

Software gestionale web-based sviluppato in Python/Flask per la gestione semplificata di ordini clienti, storico acquisti e analisi dati.

## 🚀 Funzionalità Principali

* **Dashboard Intelligente:** Creazione rapida ordini con suggerimenti basati sullo storico (prodotti preferiti, ultimi acquisti).
* **Gestione CRUD Completa:** Creazione, Lettura, Modifica (con logica avanzata) ed Eliminazione ordini.
* **Storico & Analisi:**
    * Registro tabellare con ricerca e filtri avanzati.
    * Grafici interattivi (Chart.js) per vendite mensili, top clienti e top prodotti.
    * Analisi fatturato e individuazione clienti "dormienti".
* **Export Dati:** Generazione automatica di file Excel (.xlsx) per gli ordini.
* **Interfaccia Responsive:** Ottimizzata per utilizzo desktop e portatile.

## 🛠️ Tecnologie Utilizzate

* **Backend:** Python 3, Flask, SQLAlchemy/SQLite.
* **Frontend:** HTML5, CSS3, JavaScript.
* **Librerie JS:** jQuery, DataTables (tabelle dinamiche), Select2 (ricerca dropdown), Chart.js (grafici), SweetAlert2 (popup), Flatpickr (calendario).

## ⚙️ Installazione

1.  Clona il repository.
2.  Crea un ambiente virtuale:
    ```bash
    python -m venv venv
    ```
3.  Attiva l'ambiente e installa le dipendenze:
    ```bash
    pip install -r requirements.txt
    ```
4.  Avvia l'applicazione:
    ```bash
    python app.py
    ```
5.  Apri il browser su `http://127.0.0.1:5000`.

## 🔒 Note sulla Privacy
Il database contenente i dati sensibili dei clienti (`database.db`) e le chiavi segrete (`.env`) **non sono inclusi** in questo repository per motivi di sicurezza e privacy, come definito nel file `.gitignore`.

---
*Sviluppato da Dario - 2025*
