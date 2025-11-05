from flask import Blueprint, render_template, request, redirect, url_for, flash, current_app
from flask_login import login_user, UserMixin
from ldap3 import Server, Connection, ALL, NTLM, Tls
import ssl

auth_bp = Blueprint('auth', __name__)

# Usuário mínimo em memória (temporário até adicionarmos DB de usuários)
class Usuario(UserMixin):
    def __init__(self, id, email):
        self.id = id
        self.email = email

# Dicionário dos usuários autenticados (em memória)
usuarios = {}

def user_loader(user_id: str):
    """Chamado pelo Flask-Login para recarregar usuário a partir da sessão."""
    return usuarios.get(user_id)

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email'].strip()
        senha = request.form['senha']

        if autenticar_ad(email, senha):
            user = usuarios.get(email) or Usuario(id=email, email=email)
            usuarios[email] = user
            login_user(user)
            flash('Login bem-sucedido!')
            return redirect(url_for('dashboard.dashboard'))
        else:
            flash('Usuário ou senha inválidos.')
    return render_template('login.html')

def autenticar_ad(username: str, password: str) -> bool:
    """Valida no AD via LDAPS (porta 636)."""
    if not username or not password:
        return False

    server_url = current_app.config.get('LDAP_SERVER', 'ldaps://localhost:636')
    domain = current_app.config.get('LDAP_DOMAIN', '')
    user_dn = f"{domain}\\{username}" if domain else username

    tls_config = Tls(validate=ssl.CERT_NONE, version=ssl.PROTOCOL_TLSv1_2)
    server = Server(server_url, use_ssl=True, get_info=ALL, tls=tls_config)

    try:
        conn = Connection(server, user=user_dn, password=password, authentication=NTLM, auto_bind=True)
        conn.unbind()
        return True
    except Exception as e:
        # Log básico (pode trocar por logging de verdade depois)
        print("Falha AD:", e)
        return False
