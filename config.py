import os

class Config:
    SECRET_KEY = os.getenv('SECRET_KEY', 'dev-secret')
    SQLALCHEMY_DATABASE_URI = os.getenv('DATABASE_URL', 'sqlite:///gestao.db')
    SQLALCHEMY_TRACK_MODIFICATIONS = False

    # Configurações LDAP (LDAPS:636)
    LDAP_SERVER = os.getenv('LDAP_SERVER', 'ldaps://SEU_SERVIDOR_AD:636')
    LDAP_DOMAIN = os.getenv('LDAP_DOMAIN', 'SICOOB')
    LDAP_BASE_DN = os.getenv('LDAP_BASE_DN', 'DC=sicoob,DC=local')
