from flask import Blueprint, render_template, request, redirect, url_for, flash
from flask_login import login_required, current_user
from models.funcionarios import Funcionarios
from models.checklists import ChecklistProcesso, ChecklistItem
from database import db
import datetime

checklists_bp = Blueprint("checklists", __name__)

# --- ACESSOS QUE PRECISAM DE ASSINATURA ---
ACESSOS_RESTRITOS = [
    'GRUPO_OPERADOR_DE_CREDITO',
    'BASE_ALCADA TECNICA',
    'BASE_CREDITO_COMITE I'
]

# --- REGRAS DE ACESSO (Copiadas do seu texto) ---
ACESSOS_POR_CARGO = {
    'agente_relacionamento': {
        'nome': 'Agente de Relacionamento',
        'perfil_sisbr': 'Atendimento 1,3 , Credito 1, Investimentos 1',
        'acessos': ['PLAT_GED_CADASTRO_CAPES', 'INSERIR_DOCUMENTO_GED', 'PORTAL DE GENTE_GESTAO DA PERFORMANCE', 'GRUPO_OPERADOR_DE_CREDITO', 'EDUCACAO CORPORATIVA', 'FLUXO_CAPACIDADE DE PAGAMENTO_ENCAMINHAR', 'FLUXO_ASSINATURA_ELETRONICA_ANEXAR', 'FLUXO_ASSINATURA_ELETRONICA_REVISAR', 'NOVO_SINS_VOLUNTARIO', 'SAA_ATENDIMENTO I']
    },
    'agente_atendimento': {
        'nome': 'Agente de Atendimento',
        'perfil_sisbr': 'Atendimento 1,3 , Credito 1, Investimentos 1',
        'acessos': ['PLAT_GED_CADASTRO_CAPES', 'INSERIR_DOCUMENTO_GED', 'PORTAL DE GENTE_GESTAO DA PERFORMANCE', 'GRUPO_OPERADOR_DE_CREDITO', 'EDUCACAO CORPORATIVA', 'FLUXO_CAPACIDADE DE PAGAMENTO_ENCAMINHAR', 'FLUXO_ASSINATURA_ELETRONICA_ANEXAR', 'FLUXO_ASSINATURA_ELETRONICA_REVISAR', 'NOVO_SINS_VOLUNTARIO', 'SAA_ATENDIMENTO I']
    },
    'caixa': {
        'nome': 'Caixa',
        'perfil_sisbr': 'Atendimento 1 e 2 . Crédito 1',
        'acessos': ['EDUCACAO CORPORATIVA', 'INSERIR_DOCUMENTO_GED', 'PLAT_GED_CADASTRO_CAPES', 'NOVO_SINS_VOLUNTARIO', 'PORTAL DE GENTE_GESTAO DA PERFORMANCE']
    },
    'gerente': {
        'nome': 'Gerente',
        'perfil_sisbr': 'Atendimento 1,3 , Credito 1, Investimentos 1, Retaguarda 2, PLD 4',
        'acessos': ['PLAT_GED_CADASTRO_CAPES', 'INSERIR_DOCUMENTO_GED', 'NOVO_SINS_VOLUNTARIO', 'PORTAL DE GENTE_GESTAO DA PERFORMANCE', 'BASE_CREDITO_COMITE I', 'GRUPO_OPERADOR_DE_CREDITO', 'EDUCACAO CORPORATIVA', 'FLUXO_ASSINATURA_ELETRONICA_ANEXAR', 'FLUXO_ASSINATURA_ELETRONICA_REVISAR', 'FLUXO_CAPACIDADE DE PAGAMENTO_ENCAMINHAR', 'FLUXO_PLDPCF_TESTE_CADASTRO_COOP_PARECER', 'FLUXO_PLDPCF_TESTE_CADASTRO_COOPERATIVA', 'PCF_ANALISTA_ATIVIDADE', 'PCF_RESPONSAVEL_ATIVIDADE', 'SAA_ATENDIMENTO I']
    },
    'gerente_atendimento': {
        'nome': 'Gerente de Atendimento',
        'perfil_sisbr': 'Atendimento 1,3 , Credito 1, Investimentos 1, Retaguarda 2, PLD 2,3,4',
        'acessos': ['PLAT_GED_CADASTRO_CAPES', 'INSERIR_DOCUMENTO_GED', 'NOVO_SINS_VOLUNTARIO', 'PORTAL DE GENTE_GESTAO DA PERFORMANCE', 'BASE_CREDITO_COMITE I', 'GRUPO_OPERADOR_DE_CREDITO', 'EDUCACAO CORPORATIVA', 'FLUXO_ASSINATURA_ELETRONICA_ANEXAR', 'FLUXO_ASSINATURA_ELETRONICA_REVISAR', 'FLUXO_CAPACIDADE DE PAGAMENTO_ENCAMINHAR', 'FLUXO_PLDPCF_TESTE_CADASTRO_COOP_PARECER', 'FLUXO_PLDPCF_TESTE_CADASTRO_COOPERATIVA', 'PCF_ANALISTA_ATIVIDADE', 'PCF_RESPONSAVEL_ATIVIDADE', 'PLAT_GED_CADASTRO_CAPES', 'INSERIR_DOCUMENTO_GED', 'SAA_ATENDIMENTO I', 'SAA_ATENDIMENTO II']
    },
    'recuperacao_credito': {
        'nome': 'Recuperação de Crédito',
        'perfil_sisbr': 'ATENDIMENTO 1 , CREDITO 3, FINANCEIRO 1',
        'acessos': ['GRUPO_OPERADOR_DE_CREDITO', 'BASE_ALCADA TECNICA', 'NOVO_SINS_VOLUNTARIO', 'PORTAL DE GENTE_GESTAO DA PERFORMANCE', 'INSERIR_DOCUMENTO_GED', 'EDUCACAO CORPORATIVA', 'FLUXO_PROCESSO_MANUTENCAO_GARANTIA', 'FLUXO_PROP_DIST_RURAL_CCL_DOCUMENTACAO', 'FLUXO_PROP_DIST_RURAL_CCL_FORMALIZACAO', 'FLUXO_PRORROGACAO_OPERACOES_CREDITO', 'FLUXO_PRORROGACAO_OPERACOES_CREDITO_APROVACO', 'SAA_ATENDIMENTO I']
    },
    'produtos': {
        'nome': 'Produtos',
        'perfil_sisbr': 'Padrão do Cargo',
        'acessos': ['FLUXO_ATRIBUICAO DE LIMITE DE CAMBIO_EMITIR CARTA', 'FLUXO_ATRIBUICAO DE LIMITE DE CAMBIO_INCLUIR', 'FLUXO_ATRIBUICAO DE LIMITE DE CAMBIO_RECEPCIONAR APROVACAO', 'FLUXO_ATRIBUICAO DE LIMITE DE CAMBIO_RECEPCIONAR REPROVACAO', 'FLUXO_ATRIBUICAO DE LIMITE DE CAMBIO_RECEPCIONAR VALIDACAO DA CARTA FIANCA', 'FLUXO_CAPACIDADE DE PAGAMENTO_ENCAMINHAR', 'FLUXO_DESCREDENCIAMENTO POR INATIVIDADE SIPAG', 'FLUXO_DOCUMENTOS_CAMBIO', 'INSERIR_DOCUMENTO_GED', 'PLAT_GED_CADASTRO_CAPES', 'RISCO_SOCIOAMBIENTAL_ALERTAS', 'EDUCACAO CORPORATIVA', 'NOVO_SINS_VOLUNTARIO', 'PORTAL DE GENTE_GESTAO DA PERFORMANCE']
    },
    'credito': {
        'nome': 'Crédito',
        'perfil_sisbr': 'ATENDIMENTO 2, CREDITO 2,4,6',
        'acessos': ['BASE_ALCADA TECNICA', 'BASE_CREDITO_COMITE I', 'GRUPO_OPERADOR_DE_CREDITO', 'INSERIR_DOCUMENTO_GED', 'FLUXO_APROVA_ALTER_LIMITE_CHEQUE_ESPEC_CADASTRO', 'FLUXO_ASSINATURA_ELETRONICA_ANEXAR', 'FLUXO_ASSINATURA_ELETRONICA_REVISAR', 'FLUXO_ATRIBUICAO DE LIMITE DE CAMBIO_EMITIR CARTA PARA O TOMADOR', 'FLUXO_ATRIBUICAO DE LIMITE DE CAMBIO_INCLUIR', 'FLUXO_ATRIBUICAO DE LIMITE DE CAMBIO_RECEPCIONAR APROVACAO', 'FLUXO_ATRIBUICAO DE LIMITE DE CAMBIO_RECEPCIONAR REPROVACAO', 'FLUXO_CREDITO_CCL_LIMITE_GLOBAL_ALTERACAO', 'FLUXO_CREDITO_CCL_LIMITE_GLOBAL_APROVACAO', 'FLUXO_DOCUMENTOS_CAMBIO', 'FLUXO_PROCESSO_MANUTENCAO_GARANTIA', 'FLUXO_PROP_DIST_RURAL_CCL_DOCUMENTACAO', 'FLUXO_PROP_DIST_RURAL_CCL_FORMALIZACAO', 'FLUXO_PRORROGACAO_OPERACOES_CREDITO', 'FLUXO_PRORROGACAO_OPERACOES_CREDITO_APROVACAO', 'PLAT_GED_ALCADA_CRL_1', 'PLAT_GED_ALCADA_CRL_2', 'PLAT_GED_ALCADA_CRL_3', 'PLAT_GED_ALCADA_CRL_4', 'EDUCACAO CORPORATIVA', 'NOVO_SINS_VOLUNTARIO', 'PORTAL DE GENTE_GESTAO DA PERFORMANCE']
    },
    'administrativo': {
        'nome': 'Administrativo (Jovem/Estágio)',
        'perfil_sisbr': 'Atendimento 1,2,3 Crédito 1 e Financeiro 1',
        'acessos': ['FLUXO_SGE_APROVACAO_REGISTRO_DOCUMENTO_CSC_APROVACAO_REPROVACAO', 'FLUXO_SGE_APROVACAO_REGISTRO_DOCUMENTO_CSC_REGISTRO', 'INSERIR_DOCUMENTO_GED', 'PLAT_GED_CADASTRO_CAPES', 'EDUCACAO CORPORATIVA', 'NOVO_SINS_VOLUNTARIO', 'PORTAL DE GENTE_GESTAO DA PERFORMANCE']
    },
    'riscos': {
        'nome': 'Riscos',
        'perfil_sisbr': 'Padrão do Cargo',
        'acessos': ['CONTROLES INTERNOS_NOTIFICACAO', 'EDUCACAO CORPORATIVA', 'FLUXO_PLD_ANALISTA', 'FLUXO_PLD_RESPONSAVEL', 'FLUXO_PLDPCF_TESTE_CADASTRO_COOP_PARECER', 'FLUXO_PLDPCF_TESTE_CADASTRO_COOPERATIVA', 'FLUXO_SAC_ATENDIMENTO_COOPERATIVA', 'GED_FUNDO_RESSARCIMENTO_PERDAS_OPERACIONAIS', 'GED_RESSARCIMENTO_SOLICITANTE', 'INSERIR_DOCUMENTO_GED', 'PAINEL_ADERENCIA_DECLARACAO_PCF', 'PCF_ANALISTA_ATIVIDADE', 'PCF_PAINEL ADERENCIA', 'PCF_RESPONSAVEL_ATIVIDADE', 'NOVO_SINS_VOLUNTARIO', 'PORTAL DE GENTE_GESTAO DA PERFORMANCE']
    }
}

# --- TAREFAS GERAIS ---
CHECKLIST_DEFINICOES = {
    'admissao': [
        "Validar dados pessoais e documentos no RH",
        "Criar usuário no AD (conforme padrão)",
        "Criar e-mail corporativo",
        "Configurar computador (ingressar no domínio, antivírus)",
        "Vincular patrimônio",
        "Cadastrar biometria"
    ],
    'demissao': ["Desativar AD", "Bloquear E-mail", "Inativar SISBR", "Recolher Crachá/Token", "Recolher Equipamentos"],
    'ferias': ["Configurar resposta automática", "Bloquear VPN"],
    'troca_pa': ["Alterar unidade no AD", "Migrar arquivos", "Mudar patrimônio"]
}

# --- GERADOR DE LOGIN INTELIGENTE ---
def gerar_sugestoes_login(nome_completo, pa):
    partes = nome_completo.lower().split()
    if len(partes) < 2: return "", ""

    primeiro = partes[0]
    ultimo = partes[-1]
    
    # AD: denilson.silva
    sugestao_ad = f"{primeiro}.{ultimo}"

    # SISBR: DENILSONO3246_PA (Primeiro + 1ª letra do ultimo nome, para garantir unicidade)
    # Ou conforme sua regra: primeiro + letra do segundo nome
    segundo_nome = partes[1] if len(partes) > 1 else ultimo
    letra = segundo_nome[0]
    pa_fmt = pa.zfill(2)
    sugestao_sisbr = f"{primeiro}{letra}3246_{pa_fmt}".upper()
    
    return sugestao_ad, sugestao_sisbr

# --- ROTAS ---

@checklists_bp.route("/checklists")
@login_required
def listar_ativos():
    processos = ChecklistProcesso.query.filter_by(status='Em Andamento').order_by(ChecklistProcesso.created_at.desc()).all()
    funcionarios = Funcionarios.query.filter(Funcionarios.status != 'Desligado').order_by(Funcionarios.nome).all()
    return render_template("checklists_ativos.html", processos=processos, funcionarios=funcionarios, cargos=ACESSOS_POR_CARGO)

@checklists_bp.route("/checklists/iniciar/admissao", methods=["POST"])
@login_required
def iniciar_processo_admissao():
    nome = request.form.get("nome_novo_funcionario")
    cargo_key = request.form.get("cargo_selecionado")
    num_pa = request.form.get("num_pa")

    if not nome or not cargo_key:
        flash("Nome e Cargo obrigatórios.", "danger")
        return redirect(url_for('checklists.listar_ativos'))
    
    s_ad, s_sisbr = gerar_sugestoes_login(nome, num_pa or "00")
    dados_cargo = ACESSOS_POR_CARGO.get(cargo_key)
    nome_cargo = dados_cargo['nome'] if dados_cargo else cargo_key

    # Cria funcionário
    novo_func = Funcionarios(
        nome=nome, setor=f"PA {num_pa}", cargo=nome_cargo, email=f"{s_ad}@sicoob.com.br",
        num_pa=num_pa, login_ad=s_ad, login_sisbr=s_sisbr, status='Ativo'
    )
    db.session.add(novo_func)
    db.session.commit()

    # Cria Checklist e Itens
    proc = ChecklistProcesso(tipo='admissao', funcionario_id=novo_func.id, ti_user_id=current_user.id)
    db.session.add(proc)
    
    tarefas = CHECKLIST_DEFINICOES['admissao'].copy()
    if dados_cargo:
        tarefas.append(f"Configurar Perfil SISBR: {dados_cargo['perfil_sisbr']}")
        for ac in dados_cargo['acessos']:
            t = f"Liberar: {ac}"
            if ac in ACESSOS_RESTRITOS: t += " (⚠️ REQUER TERMO)"
            tarefas.append(t)
            
    for t in tarefas:
        db.session.add(ChecklistItem(processo=proc, descricao=t))
    
    db.session.commit()
    
    flash(f"Admissão iniciada para {nome}!", "success")
    return redirect(url_for('checklists.ver_processo', processo_id=proc.id))

@checklists_bp.route("/checklists/iniciar/existente", methods=["POST"])
@login_required
def iniciar_processo_existente():
    tipo = request.form.get("tipo")
    fid = request.form.get("funcionario_id")
    
    if not fid or not tipo: return redirect(url_for('checklists.listar_ativos'))

    proc = ChecklistProcesso(tipo=tipo, funcionario_id=fid, ti_user_id=current_user.id)
    db.session.add(proc)
    for t in CHECKLIST_DEFINICOES.get(tipo, []):
        db.session.add(ChecklistItem(processo=proc, descricao=t))
    
    # Atualiza status se for demissão
    if tipo == 'demissao':
        f = Funcionarios.query.get(fid)
        f.status = 'Inativo'
    
    db.session.commit()
    flash("Checklist iniciado!", "success")
    return redirect(url_for('checklists.ver_processo', processo_id=proc.id))

@checklists_bp.route("/checklists/processo/<int:processo_id>", methods=["GET", "POST"])
@login_required
def ver_processo(processo_id):
    processo = ChecklistProcesso.query.get_or_404(processo_id)
    funcionario = processo.funcionario

    if request.method == "POST":
        # Salva Credenciais (apenas admissão)
        if processo.tipo == 'admissao':
            if request.form.get("login_ad"): funcionario.login_ad = request.form.get("login_ad")
            if request.form.get("login_sisbr"): funcionario.login_sisbr = request.form.get("login_sisbr")
            if request.form.get("email"): funcionario.email = request.form.get("email")
            db.session.add(funcionario)

        # Salva Itens
        comentario = request.form.get("comentario_geral")
        pendentes = False
        for item in processo.items:
            checado = request.form.get(f"check_{item.id}")
            if checado and item.status == 'Pendente':
                item.status = 'Concluído'
                item.ti_concluiu_id = current_user.id
                item.updated_at = datetime.datetime.now(datetime.timezone.utc)
            elif not checado:
                item.status = 'Pendente'
                item.ti_concluiu_id = None
            if item.status == 'Pendente': pendentes = True

        processo.comentario_geral = comentario
        processo.status = "Em Andamento" if pendentes else "Concluído"
        
        db.session.commit()
        
        if processo.status == "Concluído":
            if processo.tipo == 'demissao':
                funcionario.status = 'Desligado'
                db.session.commit()
            flash("Checklist finalizado!", "success")
            return redirect(url_for('checklists.listar_ativos'))
        
        flash("Salvo!", "success")
        return redirect(url_for('checklists.ver_processo', processo_id=processo.id))

    return render_template("checklist_processo.html", processo=processo, comentario_geral=processo.comentario_geral or '')

@checklists_bp.route("/checklists/termo/<int:funcionario_id>")
@login_required
def gerar_termo(funcionario_id):
    # (Mesma lógica de geração de termo que já fizemos)
    # Vou resumir para caber, mas use a versão completa anterior se tiver
    funcionario = Funcionarios.query.get_or_404(funcionario_id)
    grupos = []
    for p in funcionario.checklists:
        for i in p.items:
            for r in ACESSOS_RESTRITOS:
                if r in i.descricao:
                    grupos.append({'grupo': r, 'pa': funcionario.num_pa})
    
    grupos = [dict(t) for t in {tuple(d.items()) for d in grupos}]
    
    if not grupos:
        flash("Nenhum acesso restrito para gerar termo.", "info")
        return redirect(url_for('funcionarios.editar_funcionario', id=funcionario.id))

    dados_coop = {'razao': 'Cooperativa de Crédito Sicoob', 'cidade': 'Goiânia', 'diretor': 'Cassio Mendes', 'cpf_diretor': '...'}
    return render_template("termo_impressao.html", funcionario=funcionario, grupos=grupos, empresa=dados_coop, hoje=datetime.date.today(), solicitante=current_user)