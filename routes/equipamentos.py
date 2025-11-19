from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required
from models.funcionarios import Funcionarios
from models.equipamentos import Equipamento
from database import db

equipamentos_bp = Blueprint("equipamentos", __name__)


@equipamentos_bp.route("/equipamentos")
@login_required
def listar():
    page = request.args.get('page', 1, type=int)
    busca = request.args.get('busca', '', type=str)

    query = Equipamento.query

    if busca:
        busca_like = f"%{busca}%"
        query = query.filter(
            (Equipamento.nome_host.ilike(busca_like)) |
            (Equipamento.modelo.ilike(busca_like)) |
            (Equipamento.serial.ilike(busca_like)) |
            (Equipamento.tipo.ilike(busca_like)) |
            (Equipamento.patrimonio.ilike(busca_like))
        )

    equipamentos = query.order_by(Equipamento.patrimonio).paginate(page=page, per_page=15)
    
    return render_template("equipamentos.html",
                           equipamentos=equipamentos,
                           busca=busca)


@equipamentos_bp.route("/equipamentos/novo", methods=["GET", "POST"])
@login_required
def novo():
    vincular_id = request.args.get('vincular', None, type=int)

    if request.method == "POST":
        func_id = request.form.get("funcionario_id")
        
        if func_id == "":
            func_id = None
            status = "Disponível"
        else:
            status = "Em uso"

        novo = Equipamento(
            patrimonio=request.form.get("patrimonio"),
            nome_host=request.form.get("nome_host"),
            tipo=request.form.get("tipo"),
            serial=request.form.get("serial"),
            modelo=request.form.get("modelo"),
            status=status,
            comentarios=request.form.get("comentarios"),
            funcionario_id=func_id
        )
        db.session.add(novo)
        db.session.commit()
        flash("Equipamento cadastrado com sucesso!", "success")
        return redirect(url_for("equipamentos.listar"))

    funcionarios = Funcionarios.query.filter(Funcionarios.status != 'Desligado').order_by(Funcionarios.nome).all()
    return render_template("equipamento_form.html", 
                           equipamento=None, 
                           funcionarios=funcionarios,
                           vincular_id=vincular_id)


@equipamentos_bp.route("/equipamentos/editar/<int:id>", methods=["GET", "POST"])
@login_required
def editar_equipamento(id):
    eq = Equipamento.query.get_or_404(id)

    if request.method == "POST":
        func_id = request.form.get("funcionario_id")
        
        if func_id == "":
            func_id = None
            eq.status = "Disponível"
        else:
            eq.status = "Em uso"

        eq.patrimonio = request.form.get("patrimonio")
        eq.nome_host = request.form.get("nome_host")
        eq.tipo = request.form.get("tipo")
        eq.serial = request.form.get("serial")
        eq.modelo = request.form.get("modelo")
        eq.funcionario_id = func_id
        eq.comentarios = request.form.get("comentarios")
        
        db.session.commit()
        flash("Equipamento atualizado com sucesso!", "success")
        return redirect(url_for("equipamentos.listar"))

    funcionarios = Funcionarios.query.filter(Funcionarios.status != 'Desligado').order_by(Funcionarios.nome).all()
    return render_template("equipamento_form.html", 
                           equipamento=eq, 
                           funcionarios=funcionarios)


@equipamentos_bp.route("/equipamentos/excluir/<int:id>", methods=["GET", "POST"])
@login_required
def excluir_equipamento(id):
    eq = Equipamento.query.get_or_404(id)

    if request.method == "POST":
        db.session.delete(eq)
        db.session.commit()
        flash("Equipamento excluído permanentemente!", "danger")
        return redirect(url_for("equipamentos.listar"))

    return render_template("equipamento_delete.html", equipamento=eq)


@equipamentos_bp.route("/equipamentos/renomear/<int:id>", methods=["POST"])
@login_required
def renomear_host(id):
    eq = Equipamento.query.get_or_404(id)
    novo_nome = request.form.get("novo_nome_host")

    if novo_nome:
        eq.nome_host = novo_nome
        db.session.commit()
        flash(f"Nome do Host para {eq.patrimonio} alterado para {novo_nome}!", "success")
    else:
        flash("O novo nome do Host não pode ser vazio.", "danger")

    return redirect(url_for('equipamentos.listar'))