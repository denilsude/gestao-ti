from flask import Flask, render_template
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin
import os

# Inicialização básica
app = Flask(__name__)
app.config.from_object('config.Config')

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'auth.login'

# Importa as rotas
from routes.auth import auth_bp
app.register_blueprint(auth_bp)

# Modelo de usuário temporário (em memória)
class Usuario(UserMixin):
    def __init__(self, id, email):
        self.id = id
        self.email = email

# Dicionário temporário para manter usuários logados
usuarios = {}

# Função exigida pelo Flask-Login
@login_manager.user_loader
def load_user(user_id):
    """Recarrega o usuário a partir da sessão."""
    return usuarios.get(user_id)

@app.route('/')
def index():
    return render_template('base.html')

if __name__ == '__main__':
    app.run(debug=True)
