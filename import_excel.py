import pandas as pd
from app import app
from database import db
from models.funcionarios import Funcionarios
from models.equipamentos import Equipamento

# Nome do arquivo que você confirmou
ARQUIVO_EXCEL = "base_dados.xlsx"

def limpar(valor):
    """Limpa dados vazios ou espaços extras"""
    if pd.isna(valor):
        return None
    texto = str(valor).strip()
    if texto.lower() in ['nan', 'none', '', '0']:
        return None
    return texto

def importar_dados():
    print(f"--- Lendo arquivo: {ARQUIVO_EXCEL} ---")
    
    try:
        # Lê especificamente a aba 'Ativos'
        df = pd.read_excel(ARQUIVO_EXCEL, sheet_name='Ativos')
    except Exception as e:
        print(f"ERRO ao abrir planilha: {e}")
        return

    # Normaliza colunas para minúsculo (evita erro se estiver 'E-mail' ou 'e-mail')
    df.columns = [str(c).strip().lower() for c in df.columns]
    print(f"Colunas encontradas: {list(df.columns)}")

    total_func = 0
    total_equip = 0

    with app.app_context():
        # Dica: Se quiser limpar o banco antes para começar do zero, 
        # descomente as 2 linhas abaixo:
        # db.drop_all()
        # db.create_all()

        for index, row in df.iterrows():
            try:
                # --- 1. DADOS DO FUNCIONÁRIO ---
                nome = limpar(row.get('nome'))
                if not nome: continue # Pula linha vazia

                # Pega os dados das suas colunas específicas
                email = limpar(row.get('e-mail') or row.get('email'))
                setor = limpar(row.get('setor')) or "Geral"
                cargo = limpar(row.get('cargo')) or "Colaborador"
                telefone = limpar(row.get('telefone'))
                
                # Seus campos personalizados
                num_pa = limpar(row.get('pa'))
                login_ad = limpar(row.get('ad'))
                login_sisbr = limpar(row.get('login')) # Na sua planilha 'Login' parece ser o do SISBR

                # Verifica se já existe pelo Nome
                funcionario = Funcionarios.query.filter_by(nome=nome).first()

                if not funcionario:
                    funcionario = Funcionarios(
                        nome=nome,
                        email=email or "pendente@sicoob.com.br",
                        setor=setor,
                        cargo=cargo,
                        telefone=telefone,
                        num_pa=num_pa,
                        login_ad=login_ad,
                        login_sisbr=login_sisbr
                    )
                    db.session.add(funcionario)
                    db.session.flush() # Gera o ID sem commitar ainda
                    total_func += 1
                    print(f"Funcionario criado: {nome}")
                else:
                    # Atualiza dados se já existir
                    funcionario.num_pa = num_pa
                    funcionario.login_ad = login_ad
                    funcionario.login_sisbr = login_sisbr
                    print(f"Funcionario atualizado: {nome}")

                # --- 2. DADOS DO COMPUTADOR ---
                # Na sua planilha a coluna chama 'computador' (Ex: CAI033246...)
                nome_pc = limpar(row.get('computador'))
                
                if nome_pc:
                    # Verifica se esse PC já existe
                    equipamento = Equipamento.query.filter_by(nome_host=nome_pc).first()
                    
                    if not equipamento:
                        # Cria novo
                        equipamento = Equipamento(
                            nome_host=nome_pc,
                            # Como não vi coluna 'Patrimonio', uso o nome do PC como patrimonio provisório
                            patrimonio=nome_pc, 
                            tipo="Desktop", # Padrão, depois vc ajusta
                            status="Em uso",
                            modelo="Padrão",
                            funcionario_id=funcionario.id # <--- VINCULA AO DONO
                        )
                        db.session.add(equipamento)
                        total_equip += 1
                        print(f"   -> PC vinculado: {nome_pc}")
                    else:
                        # Se o PC já existe, atualizamos o dono dele
                        if equipamento.funcionario_id != funcionario.id:
                            equipamento.funcionario_id = funcionario.id
                            print(f"   -> PC {nome_pc} movido para {nome}")

            except Exception as e:
                print(f"Erro na linha {index}: {e}")

        db.session.commit()
        print("="*30)
        print("IMPORTAÇÃO CONCLUÍDA COM SUCESSO!")
        print(f"Novos Funcionários: {total_func}")
        print(f"Novos Equipamentos: {total_equip}")
        print("="*30)

if __name__ == "__main__":
    importar_dados()