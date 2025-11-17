from app import db

class Funcionario(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(100), nullable=False)
    setor = db.Column(db.String(50), nullable=True)
    cargo = db.Column(db.String(50), nullable=True)
    email = db.Column(db.String(100), nullable=True)
    telefone = db.Column(db.String(20), nullable=True)
