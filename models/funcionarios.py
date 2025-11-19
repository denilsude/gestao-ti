from database import db

class Funcionarios(db.Model):
    __tablename__ = 'funcionarios'

    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(120), nullable=False)
    setor = db.Column(db.String(120), nullable=False)
    cargo = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120), nullable=False)
    telefone = db.Column(db.String(20), nullable=True)
    
    # --- CAMPOS DE TI ---
    num_pa = db.Column(db.String(10), nullable=True)
    login_ad = db.Column(db.String(50), nullable=True)
    login_sisbr = db.Column(db.String(50), nullable=True)
    status = db.Column(db.String(20), nullable=False, default='Ativo')

    equipamentos = db.relationship('Equipamento', back_populates='funcionario')
    checklists = db.relationship('ChecklistProcesso', back_populates='funcionario', lazy='dynamic')