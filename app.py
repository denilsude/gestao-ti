from flask import Flask, render_template, redirect, url_for
from flask_login import LoginManager, current_user, logout_user, login_required
from database import init_app
from routes.auth import auth_bp, user_loader
from routes.dashboard import dashboard_bp

app = Flask(__name__)
app.config.from_object('config.Config')

# Login Manager
login_manager = LoginManager(app)
login_manager.login_view = 'auth.login'
login_manager.user_loader(user_loader)  # registra o carregador vindo do routes.auth

# Blueprints
app.register_blueprint(auth_bp)
app.register_blueprint(dashboard_bp)

# Inicializa DB e cria tabelas
init_app(app)

@app.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard.dashboard'))
    return render_template('base.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('auth.login'))

if __name__ == '__main__':
    app.run(debug=False)
