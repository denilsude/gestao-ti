from database import db
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash

class User(UserMixin, db.Model):
    __tablename__ = "users"

    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)
    
    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

    # --- ADICIONE ESTAS LINHAS NO FINAL DA CLASSE 'User' ---
    checklists_iniciados = db.relationship('ChecklistProcesso', back_populates='ti_user')
    items_concluidos = db.relationship('ChecklistItem', back_populates='ti_concluiu_user')