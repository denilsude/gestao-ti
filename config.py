import os

class Config:
    SECRET_KEY = 'sua_chave_secreta'
    SQLALCHEMY_DATABASE_URI = 'sqlite:///gestaoti.db'
    SQLALCHEMY_TRACK_MODIFICATIONS = False
