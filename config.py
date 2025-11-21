import os

class Config:
    # A chave secreta será fornecida via Docker para maior segurança
    SECRET_KEY = os.environ.get('SECRET_KEY') or 'uma_chave_secreta_padrao_muito_segura'
    
    # MUDANÇA: Se a variável DATABASE_URL existir (Docker), usa o PostgreSQL. 
    # Senão, usa o SQLite para desenvolvimento local.
    SQLALCHEMY_DATABASE_URI = os.environ.get('DATABASE_URL') or 'sqlite:///gestaoti.db'
    SQLALCHEMY_TRACK_MODIFICATIONS = False