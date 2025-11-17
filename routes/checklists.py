from flask import Blueprint, render_template
from flask_login import login_required

checklists_bp = Blueprint('checklists', __name__)

@checklists_bp.route('/checklist/admissao')
@login_required
def checklist_admissao():
    return render_template('checklist_admissao.html')

@checklists_bp.route('/checklist/demissao')
@login_required
def checklist_demissao():
    return render_template('checklist_demissao.html')

@checklists_bp.route('/checklist/ferias')
@login_required
def checklist_ferias():
    return render_template('checklist_ferias.html')

@checklists_bp.route('/checklist/troca_pa')
@login_required
def checklist_troca_pa():
    return render_template('checklist_troca_pa.html')
