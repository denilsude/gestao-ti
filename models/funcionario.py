from database import db

class Funcionario(db.Model):
    __tablename__ = 'funcionarios'

    id = db.Column(db.Integer, primary_key=True)
    nome = db.Column(db.String(120), nullable=False)
    setor = db.Column(db.String(120))
    cargo = db.Column(db.String(120))
    email = db.Column(db.String(160))
    telefone = db.Column(db.String(40))

    def __repr__(self):
        return f'<Funcionario {self.nome}>'
