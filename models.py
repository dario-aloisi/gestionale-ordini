from flask_sqlalchemy import SQLAlchemy

# Inizializziamo l'estensione DB
db = SQLAlchemy()

# Tabella Clienti
class Cliente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codice = db.Column(db.String(30), unique=True, nullable=False)
    nome = db.Column(db.String(150), nullable=False)
    note = db.Column(db.Text, nullable=True)
    attivo = db.Column(db.Boolean, default=True, nullable=False)

    dettagli = db.relationship('DettaglioOrdine', backref='cliente', lazy=True)

# Tabella Prodotti
class Prodotto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    codice = db.Column(db.String(30), unique=True, nullable=False)
    nome = db.Column(db.String(150), nullable=False)
    ingredienti = db.Column(db.Text, nullable=True)
    # Prezzo di Listino Attuale (default 0.0 se non lo sappiamo)
    prezzo = db.Column(db.Float, default=0.0) 
    attivo = db.Column(db.Boolean, default=True, nullable=False)
    
    dettagli = db.relationship('DettaglioOrdine', backref='prodotto', lazy=True)

# Tabella Ordini 
class Ordine(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    data_consegna = db.Column(db.Date, nullable=False)
    note = db.Column(db.Text, nullable=True)
    stato = db.Column(db.String(20), default='inviato') # inviato (o aperto)
    ora_creazione = db.Column(db.String(10), nullable=True)

    righe = db.relationship('DettaglioOrdine', backref='ordine', lazy=True, cascade="all, delete-orphan")

# Tabella Dettaglio (Righe dell'ordine)
class DettaglioOrdine(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ordine_id = db.Column(db.Integer, db.ForeignKey('ordine.id'), nullable=False)
    cliente_id = db.Column(db.Integer, db.ForeignKey('cliente.id'), nullable=False)
    prodotto_id = db.Column(db.Integer, db.ForeignKey('prodotto.id'), nullable=False)
    
    quantita = db.Column(db.Integer, nullable=False)
    # Prezzo al momento dell'ordine (Storico)
    prezzo_storico = db.Column(db.Float, default=0.0)