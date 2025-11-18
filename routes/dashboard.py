from flask import Blueprint, render_template
from flask_login import login_required
from models.funcionarios import Funcionarios
from models.equipamentos import Equipamento
from database import db
from sqlalchemy import func

dashboard_bp = Blueprint('dashboard', __name__)

@dashboard_bp.route("/dashboard")
@login_required
def dashboard():
    # --- 1. TOTAIS GERAIS ---
    total_funcionarios = Funcionarios.query.count()
    total_equipamentos = Equipamento.query.count()
    
    # --- 2. ESTOQUE (Disponível vs Em Uso) ---
    # Conta equipamentos agrupados por Tipo e Status
    # Ex: Notebook: 10 (Disp), 50 (Em uso)
    estoque_query = db.session.query(
        Equipamento.tipo,
        Equipamento.status,
        func.count(Equipamento.id)
    ).group_by(Equipamento.tipo, Equipamento.status).all()

    # Processa os dados para ficar fácil no HTML
    # Formato final: {'Notebook': {'total': 60, 'livre': 10}, ...}
    estoque = {}
    for tipo, status, count in estoque_query:
        tipo = tipo or "Outros" # Trata nulos
        if tipo not in estoque:
            estoque[tipo] = {'total': 0, 'disponivel': 0, 'em_uso': 0}
        
        estoque[tipo]['total'] += count
        if status == 'Disponível':
            estoque[tipo]['disponivel'] += count
        else:
            estoque[tipo]['em_uso'] += count

    # --- 3. EQUIPAMENTOS POR SETOR (Top 5) ---
    # Quem tem mais equipamentos?
    setor_query = db.session.query(
        Funcionarios.setor,
        func.count(Equipamento.id)
    ).join(Equipamento).group_by(Funcionarios.setor).order_by(func.count(Equipamento.id).desc()).limit(5).all()

    return render_template("dashboard.html",
                           total_funcionarios=total_funcionarios,
                           total_equipamentos=total_equipamentos,
                           estoque=estoque,
                           top_setores=setor_query)