from flask import Blueprint, render_template
from flask_login import login_required
# --- Importe os modelos ---
from models.funcionarios import Funcionarios
from models.equipamentos import Equipamento
from database import db # Importa o db

dashboard_bp = Blueprint('dashboard', __name__)

@dashboard_bp.route("/dashboard")
@login_required
def dashboard():
    # --- Faz a contagem no banco ---
    total_funcionarios = Funcionarios.query.count()
    total_equipamentos = Equipamento.query.count()
    
    # --- Envia os totais para o template ---
    return render_template("dashboard.html",
                           total_funcionarios=total_funcionarios,
                           total_equipamentos=total_equipamentos)