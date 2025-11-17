from flask import Flask
from flask_login import LoginManager
from database import db
from models.user import User

# Blueprints
from routes.auth import auth_bp
from routes.dashboard import dashboard_bp
from routes.funcionarios import funcionarios_bp
from routes.checklists import checklists_bp


def create_app():
    app = Flask(__name__)
    app.config.from_object('config.Config')

    # Banco de dados
    db.init_app(app)

    # Login
    login_manager = LoginManager()
    login_manager.login_view = 'auth.login'
    login_manager.init_app(app)

    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    # Registrar rotas
    app.register_blueprint(auth_bp)
    app.register_blueprint(dashboard_bp)
    app.register_blueprint(funcionarios_bp)
    app.register_blueprint(checklists_bp)

    return app


app = create_app()


if __name__ == "__main__":
    app.run(debug=True)
