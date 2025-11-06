from flask import Blueprint, render_template
#from flask_login import login_required
from models.funcionario import Funcionario

dashboard_bp = Blueprint('dashboard', __name__)

@dashboard_bp.route('/dashboard')
#@login_required  #desativado para teste
def dashboard():
    funcionarios = Funcionario.query.all()
    return render_template('dashboard.html', funcionarios=funcionarios)
