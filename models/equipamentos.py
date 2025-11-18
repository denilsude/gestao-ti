from database import db
from sqlalchemy import func

class Equipamento(db.Model):
    __tablename__ = 'equipamentos'

    id = db.Column(db.Integer, primary_key=True)
    
    # --- CAMPO NOVO ---
    patrimonio = db.Column(db.String(100), unique=True, nullable=False)
    
    nome_host = db.Column(db.String(100), nullable=False, unique=True)
    tipo = db.Column(db.String(50), nullable=False)
    serial = db.Column(db.String(100), unique=True)
    modelo = db.Column(db.String(100))
    status = db.Column(db.String(50), nullable=False, default="Disponível")
    
    # --- CAMPO NOVO ---
    comentarios = db.Column(db.Text, nullable=True)

    # --- A LIGAÇÃO MÁGICA ---
    funcionario_id = db.Column(db.Integer, db.ForeignKey('funcionarios.id'), nullable=True)
    funcionario = db.relationship('Funcionarios', back_populates='equipamentos')
    
    # --- Rastreamento ---
    created_at = db.Column(db.DateTime(timezone=True), server_default=func.now())
    updated_at = db.Column(db.DateTime(timezone=True), onupdate=func.now())

    def __repr__(self):
        return f'<Equipamento {self.patrimonio} - {self.nome_host}>'