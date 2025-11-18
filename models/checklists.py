from database import db
from sqlalchemy import func

# Representa um PROCESSO de checklist (Ex: "Demissão do Denilson")
class ChecklistProcesso(db.Model):
    __tablename__ = 'checklist_processos'
    
    id = db.Column(db.Integer, primary_key=True)
    tipo = db.Column(db.String(50), nullable=False) # 'demissao', 'ferias', etc.
    status = db.Column(db.String(50), nullable=False, default='Em Andamento') # Em Andamento, Concluído
    created_at = db.Column(db.DateTime(timezone=True), server_default=func.now())
    
    # --- CAMPO NOVO ---
    comentario_geral = db.Column(db.Text, nullable=True) # O comentário principal

    # --- Relacionamentos ---
    funcionario_id = db.Column(db.Integer, db.ForeignKey('funcionarios.id'), nullable=False)
    funcionario = db.relationship('Funcionarios', back_populates='checklists')

    ti_user_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=False)
    ti_user = db.relationship('User', back_populates='checklists_iniciados')

    items = db.relationship('ChecklistItem', back_populates='processo', cascade="all, delete-orphan")

# Representa um ITEM de checklist (Ex: "Desativar usuário")
class ChecklistItem(db.Model):
    __tablename__ = 'checklist_items'
    
    id = db.Column(db.Integer, primary_key=True)
    descricao = db.Column(db.String(500), nullable=False)
    status = db.Column(db.String(50), nullable=False, default='Pendente') # Pendente, Concluído
    
    # --- O COMENTÁRIO INDIVIDUAL FOI REMOVIDO DA LÓGICA ---
    comentario = db.Column(db.Text, nullable=True) # (Ainda no BD, mas não vamos mais usar)
    
    updated_at = db.Column(db.DateTime(timezone=True), onupdate=func.now())
    
    processo_id = db.Column(db.Integer, db.ForeignKey('checklist_processos.id'), nullable=False)
    processo = db.relationship('ChecklistProcesso', back_populates='items')
    
    ti_concluiu_id = db.Column(db.Integer, db.ForeignKey('users.id'), nullable=True)
    ti_concluiu_user = db.relationship('User', back_populates='items_concluidos')