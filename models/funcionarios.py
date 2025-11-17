from database import db

class Funcionarios(db.Model):
    __tablename__ = 'funcionarios'

    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(120), nullable=False)
    setor = db.Column(db.String(120), nullable=False)
    cargo = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), nullable=False)
    telefone = db.Column(db.String(20), nullable=False)
