from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
from database import db
from models.funcionario import Funcionario

funcionarios_bp = Blueprint("funcionarios", __name__)

@funcionarios_bp.route("/funcionarios/novo", methods=["GET", "POST"])
@login_required
def novo_funcionario():
    if request.method == "POST":
        nome = request.form.get("nome")
        email = request.form.get("email")
        cargo = request.form.get("cargo")
        setor = request.form.get("setor")

        novo = Funcionario(nome=nome, email=email, cargo=cargo, setor=setor)
        db.session.add(novo)
        db.session.commit()
        flash("Funcionário adicionado com sucesso!", "success")
        return redirect(url_for("dashboard.dashboard"))

    return render_template("funcionario_form.html", titulo="Novo Funcionário")

@funcionarios_bp.route("/funcionarios/<int:id>/editar", methods=["GET", "POST"])
@login_required
def editar_funcionario(id):
    funcionario = Funcionario.query.get_or_404(id)

    if request.method == "POST":
        funcionario.nome = request.form.get("nome")
        funcionario.email = request.form.get("email")
        funcionario.cargo = request.form.get("cargo")
        funcionario.setor = request.form.get("setor")
        db.session.commit()
        flash("Funcionário atualizado com sucesso!", "success")
        return redirect(url_for("dashboard.dashboard"))

    return render_template("funcionario_form.html", funcionario=funcionario, titulo="Editar Funcionário")

@funcionarios_bp.route("/funcionarios/<int:id>/excluir", methods=["GET", "POST"])
@login_required
def excluir_funcionario(id):
    funcionario = Funcionario.query.get_or_404(id)

    if request.method == "POST":
        db.session.delete(funcionario)
        db.session.commit()
        flash("Funcionário excluído com sucesso!", "success")
        return redirect(url_for("dashboard.dashboard"))

    return render_template("funcionario_delete.html", funcionario=funcionario)
