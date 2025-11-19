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
    # --- 1. TOTAIS E CONTADORES GERAIS ---
    total_funcionarios = Funcionarios.query.filter(Funcionarios.status != 'Desligado').count()
    total_equipamentos = Equipamento.query.count()
    
    # --- 2. MONITOR DE ESTOQUE POR TIPO ---
    TIPOS_ESTOQUE = ['Notebook', 'Desktop', 'Monitor', 'Fone', 'Teclado']
    
    estoque_query = db.session.query(
        Equipamento.tipo,
        Equipamento.status,
        func.count(Equipamento.id)
    ).filter(Equipamento.tipo.in_(TIPOS_ESTOQUE)).group_by(Equipamento.tipo, Equipamento.status).all()

    estoque = {tipo: {'total': 0, 'disponivel': 0, 'em_uso': 0} for tipo in TIPOS_ESTOQUE}
    
    for tipo, status, count in estoque_query:
        if tipo in estoque:
            estoque[tipo]['total'] += count
            if status == 'Disponível':
                estoque[tipo]['disponivel'] += count
            else:
                estoque[tipo]['em_uso'] += count

    # --- 3. COLABORADORES POR DEPARTAMENTO (TOP 5) ---
    setor_query = db.session.query(
        Funcionarios.setor,
        func.count(Funcionarios.id)
    ).filter(Funcionarios.status == 'Ativo').group_by(Funcionarios.setor).order_by(func.count(Funcionarios.id).desc()).limit(5).all()

    return render_template("dashboard.html",
                           total_funcionarios=total_funcionarios,
                           total_equipamentos=total_equipamentos,
                           estoque=estoque,
                           top_setores=setor_query)