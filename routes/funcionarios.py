from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required
from models.funcionarios import Funcionarios
from app import db

funcionarios_bp = Blueprint("funcionarios", __name__)


@funcionarios_bp.route("/funcionarios")
@login_required
def listar():
    funcionarios = Funcionarios.query.order_by(Funcionarios.nome).all()
    return render_template("funcionarios.html", funcionarios=funcionarios)


@funcionarios_bp.route("/funcionarios/novo", methods=["GET", "POST"])
@login_required
def novo():
    if request.method == "POST":
        novo = Funcionarios(
            nome=request.form.get("nome"),
            setor=request.form.get("setor"),
            cargo=request.form.get("cargo"),
            email=request.form.get("email"),
            telefone=request.form.get("telefone")
        )
        db.session.add(novo)
        db.session.commit()
        flash("Funcionário cadastrado com sucesso!", "success")
        return redirect(url_for("funcionarios.listar"))

    return render_template("funcionario_form.html")
