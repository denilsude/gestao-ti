from flask import Blueprint, render_template, request
from models.funcionario import Funcionario

dashboard_bp = Blueprint('dashboard', __name__)

@dashboard_bp.route('/dashboard')
def dashboard():
    page = request.args.get('page', 1, type=int)
    busca = request.args.get('busca', '', type=str)

    query = Funcionario.query
    if busca:
        busca_like = f"%{busca}%"
        query = query.filter(
            Funcionario.nome.ilike(busca_like) |
            Funcionario.setor.ilike(busca_like)
        )

    funcionarios = query.order_by(Funcionario.nome).paginate(page=page, per_page=15)
    return render_template('dashboard.html', funcionarios=funcionarios, busca=busca)

@dashboard_bp.route('/checklist/<tipo>')
def checklist(tipo):
    tipos_validos = ['admissao', 'demissao', 'ferias', 'troca_pa']
    if tipo not in tipos_validos:
        return render_template('404.html'), 404
    return render_template(f'checklists/checklist_{tipo}.html', tipo=tipo)
