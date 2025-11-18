from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
from models.funcionarios import Funcionarios
from models.checklists import ChecklistProcesso, ChecklistItem
from database import db
import datetime

checklists_bp = Blueprint("checklists", __name__)

# --- DEFINIÇÕES DOS NOSSOS CHECKLISTS ---
# Agora incluindo 'admissao'
CHECKLIST_DEFINICOES = {
    'admissao': [
        "Definir nome de usuário e e-mail com RH",
        "Criar usuário no Active Directory",
        "Criar e-mail corporativo",
        "Preparar e configurar computador",
        "Vincular patrimônio ao novo usuário",
        "Configurar acessos (sistemas, pastas)"
    ],
    'demissao': [
        "Desativar usuário",
        "Revogar acessos",
        "Recolher equipamentos",
        "Registrar devolução de ativos",
        "Remover do e-mail / grupos / AD"
    ],
    'ferias': [
        "Bloquear acesso de internet caso necessário",
        "Redirecionar e-mail temporariamente",
        "Ajustar escala de acesso"
    ],
    'troca_pa': [
        "Mudança de cargo",
        "Inativar usuário SIBR",
        "Criar novo acesso no SIBR",
        "Enviar termo para diretoria",
        "Alterar operador CCS",
        "Mudar cargo no Teams",
        "Alterar usuário no AD (políticas)",
        "Alterar grupos do Outlook",
        "Alterar acesso da Biometria",
        "Mudar patrimônio do computador/equipamentos para nova agência"
    ]
}


@checklists_bp.route("/checklists")
@login_required
def listar_ativos():
    """
    Página principal de Checklists:
    1. Mostra processos em andamento.
    2. Mostra formulários para iniciar novos processos.
    """
    processos = ChecklistProcesso.query.filter_by(status='Em Andamento').order_by(ChecklistProcesso.created_at.desc()).all()
    
    # Pega funcionários para o dropdown de "existentes"
    funcionarios = Funcionarios.query.order_by(Funcionarios.nome).all()
    
    return render_template("checklists_ativos.html", 
                           processos=processos,
                           funcionarios=funcionarios)


def _criar_processo(tipo, funcionario_id, ti_user_id):
    """Função helper interna para criar o processo e seus itens."""
    tarefas = CHECKLIST_DEFINICOES.get(tipo)
    if not tarefas:
        return None, "Checklist desse tipo não configurado"

    # Cria o "Cabeçalho"
    novo_processo = ChecklistProcesso(
        tipo=tipo,
        funcionario_id=funcionario_id,
        ti_user_id=ti_user_id
    )
    db.session.add(novo_processo)
    
    # Cria os "Itens"
    for tarefa_desc in tarefas:
        novo_item = ChecklistItem(
            processo=novo_processo,
            descricao=tarefa_desc
        )
        db.session.add(novo_item)
    
    db.session.commit()
    return novo_processo, None


@checklists_bp.route("/checklists/iniciar/existente", methods=["POST"])
@login_required
def iniciar_processo_existente():
    """Cria um checklist para um funcionário que JÁ EXISTE."""
    
    tipo = request.form.get("tipo")
    funcionario_id = request.form.get("funcionario_id")

    if not funcionario_id or not tipo:
        flash("Você deve selecionar um tipo de checklist e um funcionário.", "danger")
        return redirect(url_for('checklists.listar_ativos'))

    processo_existente = ChecklistProcesso.query.filter_by(
        funcionario_id=funcionario_id, 
        tipo=tipo, 
        status='Em Andamento'
    ).first()

    if processo_existente:
        flash("Já existe um checklist desse tipo em andamento para este funcionário.", "warning")
        return redirect(url_for('checklists.ver_processo', processo_id=processo_existente.id))

    novo_processo, erro = _criar_processo(tipo, funcionario_id, current_user.id)
    if erro:
        flash(erro, "danger")
        return redirect(url_for('checklists.listar_ativos'))

    flash(f"Checklist de {tipo} iniciado para {novo_processo.funcionario.nome}!", "success")
    return redirect(url_for('checklists.ver_processo', processo_id=novo_processo.id))


@checklists_bp.route("/checklists/iniciar/admissao", methods=["POST"])
@login_required
def iniciar_processo_admissao():
    """
    FLUXO DE ADMISSÃO:
    1. Cria um funcionário "placeholder".
    2. Inicia o checklist de 'admissao' para ele.
    """
    nome_novo_funcionario = request.form.get("nome_novo_funcionario")
    if not nome_novo_funcionario:
        flash("O nome do novo funcionário é obrigatório.", "danger")
        return redirect(url_for('checklists.listar_ativos'))
        
    # 1. Cria o funcionário placeholder
    novo_funcionario = Funcionarios(
        nome=nome_novo_funcionario,
        setor="Aguardando Admissão",
        cargo="Aguardando Admissão",
        email="a.definir@empresa.com" # (será atualizado pelo checklist)
    )
    db.session.add(novo_funcionario)
    db.session.commit() # Commit para que o 'novo_funcionario' tenha um ID

    # 2. Inicia o checklist para este novo ID
    novo_processo, erro = _criar_processo('admissao', novo_funcionario.id, current_user.id)
    if erro:
        flash(erro, "danger")
        return redirect(url_for('checklists.listar_ativos'))
    
    flash(f"Checklist de Admissão iniciado para {novo_funcionario.nome}!", "success")
    return redirect(url_for('checklists.ver_processo', processo_id=novo_processo.id))


@checklists_bp.route("/checklists/processo/<int:processo_id>", methods=["GET", "POST"])
@login_required
def ver_processo(processo_id):
    """Mostra o checklist interativo (com checkboxes)."""
    
    processo = ChecklistProcesso.query.get_or_404(processo_id)

    if request.method == "POST":
        comentario_geral = request.form.get("comentario_geral")
        itens_pendentes = False

        for item in processo.items:
            # Não pegamos mais o comentário individual
            
            # Verifica o status do checkbox
            status_item = request.form.get(f"check_{item.id}")
            if status_item == 'on': # Caixa marcada
                if item.status == 'Pendente':
                    item.status = 'Concluído'
                    item.ti_concluiu_id = current_user.id
                    item.updated_at = datetime.datetime.now(datetime.timezone.utc)
            else: # Caixa não marcada
                item.status = 'Pendente'
                item.ti_concluiu_id = None
            
            if item.status == 'Pendente':
                itens_pendentes = True

        # SUA LÓGICA: Se itens estiverem pendentes, o comentário geral é obrigatório
        if itens_pendentes and not comentario_geral:
            flash("Você deve preencher um comentário geral se houver itens pendentes.", "danger")
            db.session.rollback() # Desfaz as mudanças
            return redirect(url_for('checklists.ver_processo', processo_id=processo_id))

        # Salva o comentário geral no processo
        processo.comentario_geral = comentario_geral
            
        # Atualiza o status geral do processo
        if not itens_pendentes:
            processo.status = "Concluído"
            flash("Checklist concluído com sucesso!", "success")
        else:
            processo.status = "Em Andamento"
            flash("Progresso salvo!", "info")

        db.session.commit()
        
        if processo.status == "Concluído":
             return redirect(url_for('checklists.listar_ativos'))
        
        return redirect(url_for('checklists.ver_processo', processo_id=processo_id))
        
    # Se for GET, apenas mostra a página
    # Passamos o comentário geral salvo para o textarea
    return render_template("checklist_processo.html", 
                           processo=processo,
                           comentario_geral=processo.comentario_geral or '')