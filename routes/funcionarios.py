from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required
from models.funcionarios import Funcionarios
from database import db
from sqlalchemy import desc

funcionarios_bp = Blueprint("funcionarios", __name__)


@funcionarios_bp.route("/funcionarios")
@login_required
def listar():
    page = request.args.get('page', 1, type=int)
    busca = request.args.get('busca', '', type=str)

    query = Funcionarios.query

    if busca:
        busca_like = f"%{busca}%"
        query = query.filter(
            (Funcionarios.nome.ilike(busca_like)) |
            (Funcionarios.setor.ilike(busca_like)) |
            (Funcionarios.email.ilike(busca_like))
        )
    
    # Filtra por padrão APENAS os ativos/inativos, exclui desligados do filtro principal
    query = query.filter(Funcionarios.status != 'Desligado')

    funcionarios = query.order_by(Funcionarios.nome).paginate(page=page, per_page=15)

    return render_template("funcionarios.html",
                           funcionarios=funcionarios,
                           busca=busca)


@funcionarios_bp.route("/funcionarios/novo", methods=["GET", "POST"])
@login_required
def novo():
    # Rota 'novo' mantida para compatibilidade, mas o fluxo principal é via Admissão
    if request.method == "POST":
        novo = Funcionarios(
            nome=request.form.get("nome"),
            setor=request.form.get("setor"),
            cargo=request.form.get("cargo"),
            email=request.form.get("email"),
        )
        db.session.add(novo)
        db.session.commit()
        flash("Colaborador cadastrado (via formulário rápido)!", "success")
        return redirect(url_for("funcionarios.listar"))

    return render_template("funcionario_form.html", funcionario=None)


@funcionarios_bp.route("/funcionarios/editar/<int:id>", methods=["GET", "POST"])
@login_required
def editar_funcionario(id):
    f = Funcionarios.query.get_or_404(id)

    if request.method == "POST":
        f.nome = request.form.get("nome")
        f.setor = request.form.get("setor")
        f.cargo = request.form.get("cargo")
        f.email = request.form.get("email")
        
        db.session.commit()
        flash("Colaborador atualizado com sucesso!", "success")
        return redirect(url_for("funcionarios.listar"))

    checklists_hist = f.checklists.order_by(desc('created_at')).all()

    return render_template("funcionario_form.html", 
                           funcionario=f,
                           checklists_hist=checklists_hist)


@funcionarios_bp.route("/funcionarios/excluir/<int:id>", methods=["GET", "POST"])
@login_required
def excluir_funcionario(id):
    f = Funcionarios.query.get_or_404(id)

    if request.method == "POST":
        db.session.delete(f)
        db.session.commit()
        flash("Colaborador excluído permanentemente!", "danger")
        return redirect(url_for("funcionarios.listar"))

    return render_template("funcionario_delete.html", funcionario=f)