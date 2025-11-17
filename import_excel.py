import pandas as pd
from app import app
from database import db
from models.funcionarios import Funcionario

ARQUIVO = "data/funcionarios.xlsx"

def pick(row, *candidatos):
    """Tenta pegar o valor de uma entre várias opções de coluna (case-insensitive)."""
    for col in candidatos:
        for real in row.index:
            if real.strip().lower() == col.strip().lower():
                val = row.get(real)
                if pd.notna(val):
                    return str(val).strip()
    return None

def importar_funcionarios():
    with app.app_context():
        df = pd.read_excel(ARQUIVO)
        inseridos = 0

        for _, r in df.iterrows():
            nome = pick(r, 'Nome')
            if not nome:
                continue

            setor = pick(r, 'Setor', 'Departamento')
            cargo = pick(r, 'Cargo', 'Função')
            email = pick(r, 'E-mail', 'Email', 'e_mail')
            telefone = pick(r, 'Telefone', 'Fone', 'Celular')

            f = Funcionario(
                nome=nome,
                setor=setor,
                cargo=cargo,
                email=email,
                telefone=telefone
            )
            db.session.add(f)
            inseridos += 1

        db.session.commit()
        print(f"Importação concluída: {inseridos} registros inseridos.")

if __name__ == "__main__":
    importar_funcionarios()
