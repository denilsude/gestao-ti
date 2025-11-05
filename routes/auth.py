from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_user, logout_user, login_required, UserMixin
from ldap3 import Server, Connection, ALL, NTLM, Tls
import ssl

auth_bp = Blueprint('auth', __name__)

# Modelo mínimo de usuário
class Usuario(UserMixin):
    def __init__(self, id, email):
        self.id = id
        self.email = email

# Simples storage (substituiremos por banco depois)
usuarios = {}

@auth_bp.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email']
        senha = request.form['senha']
        if autenticar_ad(email, senha):
            user = Usuario(id=email, email=email)
            usuarios[email] = user
            login_user(user)
            flash('Login bem-sucedido!')
            return redirect(url_for('index'))
        else:
            flash('Usuário ou senha inválidos.')
    return render_template('login.html')

def autenticar_ad(username, password):
    """Valida no AD via LDAPS (porta 636)."""
    AD_SERVER = 'ldaps://srv-ad01.sicoob.local:636'
    AD_DOMAIN = 'SICOOB'
    USER_DN = f"{AD_DOMAIN}\\{username}"

    tls_config = Tls(validate=ssl.CERT_NONE, version=ssl.PROTOCOL_TLSv1_2)
    server = Server(AD_SERVER, use_ssl=True, get_info=ALL, tls=tls_config)

    try:
        conn = Connection(server, user=USER_DN, password=password, authentication=NTLM, auto_bind=True)
        conn.unbind()
        return True
    except Exception:
        return False
