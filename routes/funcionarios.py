from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required
from models.funcionarios import Funcionario
from app import db

funcionarios_bp = Blueprint('funcionarios', __name__)

@funcionarios_bp.route('/funcionarios')
@login_required
def listar_funcionarios():
    funcionarios = Funcionario.query.all()
    return render_template('funcionarios.html', funcionarios=funcionarios)

@funcionarios_bp.route('/funcionarios/novo', methods=['GET', 'POST'])
@login_required
def novo_funcionario():
    if request.method == 'POST':
        nome = request.form['nome']
        setor = request.form['setor']
        cargo = request.form['cargo']
        email = request.form['email']
        telefone = request.form['telefone']

        novo = Funcionario(nome=nome, setor=setor, cargo=cargo, email=email, telefone=telefone)
        db.session.add(novo)
        db.session.commit()

        flash('Funcionário adicionado com sucesso!', 'success')
        return redirect(url_for('funcionarios.listar_funcionarios'))

    return render_template('funcionario_form.html')
