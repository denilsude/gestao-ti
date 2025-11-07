from app import app, db
from models.user import User

with app.app_context():
    db.drop_all()
    db.create_all()
    print("Banco criado com sucesso!")

    u = User(username="admin")
    u.set_password("123456")
    db.session.add(u)
    db.session.commit()
    print("Usuário admin criado com sucesso!")
