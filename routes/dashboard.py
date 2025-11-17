from flask import Blueprint, render_template, request
from flask_login import login_required
from models.funcionarios import Funcionarios

dashboard_bp = Blueprint('dashboard', __name__)


@dashboard_bp.route("/dashboard")
@login_required
def dashboard():
    page = request.args.get('page', 1, type=int)
    busca = request.args.get('busca', '', type=str)

    query = Funcionarios.query

    if busca:
        busca_like = f"%{busca}%"
        query = query.filter(
            Funcionarios.nome.ilike(busca_like) |
            Funcionarios.setor.ilike(busca_like)
        )

    funcionarios = query.order_by(Funcionarios.nome).paginate(page=page, per_page=15)

    return render_template("dashboard.html",
                           funcionarios=funcionarios,
                           busca=busca)
