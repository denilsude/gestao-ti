from flask import Blueprint, render_template

checklists_bp = Blueprint("checklists", __name__)


@checklists_bp.route("/checklists/<tipo>")
def checklist(tipo):
    tipos = ["admissao", "demissao", "ferias", "troca_pa"]

    if tipo not in tipos:
        return "Checklist não encontrado", 404

    return render_template(f"checklists/checklist_{tipo}.html")
