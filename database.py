from flask_sqlalchemy import SQLAlchemy
from flask import Flask
from config import Config

db = SQLAlchemy()

def init_app(app: Flask):
    app.config.from_object(Config)
    db.init_app(app)
    with app.app_context():
        db.create_all()
