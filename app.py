from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager
from config import Config




db = SQLAlchemy()
login_manager = LoginManager()

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    db.init_app(app)
    login_manager.init_app(app)
    login_manager.login_view = 'auth.login'

    from routes.auth import auth_bp
    from routes.dashboard import dashboard_bp
    from routes.funcionarios import funcionarios_bp
    from routes.checklists import checklists_bp
    app.register_blueprint(auth_bp)
    app.register_blueprint(dashboard_bp)
    app.register_blueprint(funcionarios_bp)
    app.register_blueprint(checklists_bp)
    return app

app = create_app()

from models import user, funcionarios
with app.app_context():
    db.create_all()
