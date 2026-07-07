import re
import datetime as dt
from pathlib import Path
from io import BytesIO
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter
import requests
import base64

# -----------------------  Configurações  -----------------------
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR

CAMINHO_SALAS = DATA_DIR / "SALAS - COPIA.xlsx"
CAMINHO_DISCIPLINAS = DATA_DIR / "dados_disciplinas.xlsx"

DIAS_SEMANA = ["SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"]
INDICE_DIAS = {d: i for i, d in enumerate(DIAS_SEMANA)}

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

# ===============================
# FUNÇÕES DE COMMIT NO GITHUB
# ===============================

def commit_dados_disciplinas(df, mensagem=None):
    """Salva o DataFrame diretamente no repositório do GitHub via API."""
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo = st.secrets["REPO_NAME"]
        branch = st.secrets.get("BRANCH", "main")
    except KeyError as e:
        st.error(f"❌ Secret não configurado: {e}. Vá em Settings → Secrets no Streamlit Cloud.")
        return False

    caminho_arquivo = "dados_disciplinas.xlsx"

    if mensagem is None:
        mensagem = f"Atualiza alocação de salas - {dt.datetime.now().strftime('%d/%m/%Y %H:%M')}"

    buffer = BytesIO()
    df.to_excel(buffer, index=False, engine='openpyxl')
    conteudo_base64 = base64.b64encode(buffer.getvalue()).decode()

    url_api = f"https://api.github.com/repos/{repo}/contents/{caminho_arquivo}"
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json"
    }

    resp_get = requests.get(url_api, headers=headers, params={"ref": branch})

    sha_atual = None
    if resp_get.status_code == 200:
        sha_atual = resp_get.json().get("sha")
    elif resp_get.status_code == 404:
        pass
    else:
        st.error(f"Erro ao buscar arquivo no GitHub: {resp_get.status_code}")
        return False

    payload = {
        "message": mensagem,
        "content": conteudo_base64,
        "branch": branch
    }
    if sha_atual:
        payload["sha"] = sha_atual

    resp_put = requests.put(url_api, headers=headers, json=payload)

    if resp_put.status_code in [200, 201]:
        return True
    else:
        st.error(f"Erro no commit: {resp_put.status_code} - {resp_put.text}")
        return False


# ===============================
# FUNÇÕES DE LEITURA E PROCESSAMENTO
# ===============================

@st.cache_data(show_spinner=False)
def carregar_dados():
    """Carrega os dados de salas e turmas do repositório."""
    if not CAMINHO_DISCIPLINAS.exists():
        st.error(f"❌ Arquivo de disciplinas não encontrado em: {CAMINHO_DISCIPLINAS}")
        st.stop()

    df_turmas = pd.read_excel(CAMINHO_DISCIPLINAS, sheet_name="salas")

    # Extrair lista de salas e capacidades do próprio arquivo de disciplinas
    salas_info = {}
    for _, row in df_turmas.iterrows():
        sala = str(row.get("SALA") or "").strip()
        if not sala:
            continue
        alunos = row.get("ALUNOS")
        try:
            alunos = int(alunos) if pd.notna(alunos) else 0
        except:
            alunos = 0
        if sala not in salas_info or salas_info[sala] < alunos:
            salas_info[sala] = alunos

    # Se houver arquivo de salas separado, usa ele para complementar
    if CAMINHO_SALAS.exists():
        try:
            df_salas_arq = pd.read_excel(CAMINHO_SALAS)
            for _, row in df_salas_arq.iterrows():
                nome = str(row.get("SALAS") or row.get("SALA") or row.get("NOME") or "").strip()
                cap = row.get("CAPACIDADE")
                try:
                    cap = int(cap) if pd.notna(cap) else 0
                except:
                    cap = 0
                if nome and nome not in salas_info:
                    salas_info[nome] = cap
                elif nome and cap > salas_info.get(nome, 0):
                    salas_info[nome] = cap
        except Exception:
            pass

    return df_turmas, salas_info


def str_to_time(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    if isinstance(s, dt.time):
        return s
    s = str(s).strip()
    for fmt in ("%H:%M:%S", "%H:%M", "%H.%M"):
        try:
            return dt.datetime.strptime(s, fmt).time()
        except Exception:
            pass
    s2 = re.sub(r'[^0-9:]', '', s)
    try:
        return dt.datetime.strptime(s2, "%H:%M").time()
    except Exception:
        return None


def time_to_minutes(t):
    if t is None:
        return 0
    return t.hour * 60 + t.minute


def intervals_overlap(a_start, a_end, b_start, b_end):
    a_s = time_to_minutes(str_to_time(a_start))
    a_e = time_to_minutes(str_to_time(a_end))
    b_s = time_to_minutes(str_to_time(b_start))
    b_e = time_to_minutes(str_to_time(b_end))
    return max(a_s, b_s) < min(a_e, b_e)


def re_split_days(s: str):
    parts = re.split(r'[;,/\\]+|\s{2,}|\s', s)
    return [p for p in parts if p]


def criar_lista_salas(salas_info: dict):
    salas = []
    for nome, capacidade in salas_info.items():
        salas.append({
            "NOME": nome,
            "CAPACIDADE": capacidade,
            "DATAS": set(),
            "HORARIOS_OCUPADOS": set(),
            "HORARIOS_OCUPADOS_SEMANA": {d: [] for d in DIAS_SEMANA},
            "RESERVAS": []
        })
    return salas


def gerar_datas(df_turmas):
    try:
        data_inicio = list(map(int, str(df_turmas.iloc[0, 13]).split(",")))
        data_final = list(map(int, str(df_turmas.iloc[0, 14]).split(",")))
        return pd.date_range(dt.date(*data_inicio), dt.date(*data_final))
    except Exception:
        try:
            col0 = df_turmas.columns[0]
            min_date = pd.to_datetime(df_turmas[col0]).min().date()
            max_date = pd.to_datetime(df_turmas[col0]).max().date()
            return pd.date_range(min_date, max_date)
        except Exception:
            hoje = dt.date.today()
            return pd.date_range(hoje, hoje)


def processar_alocacoes(df_turmas: pd.DataFrame, todas_as_datas, salas_ct: list):
    registros = []
    for _, aloc in df_turmas.iterrows():
        status = str(aloc.get("STATUS") or "").strip()
        if status.upper() != "ALOCADA":
            continue
        sala = str(aloc.get("SALA") or aloc.get("SALAS") or "").strip()
        if not sala:
            continue
        dias_raw = str(aloc.get("DIAS") or "").strip()
        if not dias_raw:
            continue
        dias_tokens = [t.strip().upper() for t in re_split_days(dias_raw)]
        dias_validos = [d for d in dias_tokens if d in INDICE_DIAS]
        if not dias_validos:
            continue
        inicio_raw = aloc.get("HORARIO INICIO") or aloc.get("HORARIO") or aloc.get("HORÁRIO INICIO")
        fim_raw = aloc.get("HORARIO FINAL") or aloc.get("HORÁRIO FINAL") or aloc.get("HORARIO_FIM")
        inicio_t = str_to_time(inicio_raw)
        fim_t = str_to_time(fim_raw)
        descricao = (
            f"{aloc.get('CODIGO') or ''} - "
            f"{aloc.get('DISCIPLINA') or ''} - "
            f"{aloc.get('TURMA') or ''} - "
            f"{aloc.get('PROFESSOR') or ''}"
        )

        indices = [INDICE_DIAS[d] for d in dias_validos]
        datas = todas_as_datas[todas_as_datas.dayofweek.isin(indices)]
        registros.append({
            "CURSO": aloc.get("CURSO"),
            "CODIGO": aloc.get("CODIGO"),
            "SALA": sala,
            "DISCIPLINA": aloc.get("DISCIPLINA"),
            "TURMA": aloc.get("TURMA"),
            "DIAS": ",".join(dias_validos),
            "HORARIO_INICIO": inicio_t.strftime("%H:%M") if inicio_t else None,
            "HORARIO_FINAL": fim_t.strftime("%H:%M") if fim_t else None,
            "HORARIOS_RAW": aloc.get("HORARIO") or aloc.get("HORÁRIO") or "",
            "ALUNOS": aloc.get("ALUNOS") or 0,
            "PROFESSOR": aloc.get("PROFESSOR"),
            "CAPACIDADE": next((s["CAPACIDADE"] for s in salas_ct if s["NOME"] == sala), None),
            "DATAS": datas,
            "DESCRICAO": descricao,
            "TIPO": "DISCIPLINA"
        })

        sala_obj = next((s for s in salas_ct if s["NOME"] == sala), None)
        if sala_obj:
            for d in dias_validos:
                if inicio_t and fim_t:
                    sala_obj["HORARIOS_OCUPADOS_SEMANA"][d].append((
                        inicio_t.strftime("%H:%M"), fim_t.strftime("%H:%M"), descricao
                    ))
                    sala_obj["HORARIOS_OCUPADOS"].add(f"{inicio_t.strftime('%H:%M')} - {fim_t.strftime('%H:%M')}")
                else:
                    raw = str(aloc.get("HORARIO") or "")
                    blocos = [b.strip() for b in raw.split(",") if b.strip()]
                    for bloco in blocos:
                        try:
                            parts = bloco.split()
                            dia = parts[0].upper()
                            horas = parts[1]
                            h1, h2 = horas.split("-")
                            if dia in DIAS_SEMANA:
                                sala_obj["HORARIOS_OCUPADOS_SEMANA"][dia].append((h1, h2, descricao))
                                sala_obj["HORARIOS_OCUPADOS"].add(f"{h1} - {h2}")
                        except Exception:
                            continue
    return pd.DataFrame(registros)


def verificar_disponibilidade(salas_ct, data_ini, data_fim, dias_evento, h_ini, h_fim, capacidade, bloco_pref=None, sala_pref=None):
    """Verifica quais salas estão disponíveis para o evento solicitado."""
    mapping = {'MONDAY': 'SEGUNDA', 'TUESDAY': 'TERÇA', 'WEDNESDAY': 'QUARTA',
               'THURSDAY': 'QUINTA', 'FRIDAY': 'SEXTA', 'SATURDAY': 'SÁBADO', 'SUNDAY': 'DOMINGO'}

    inicio_str = h_ini.strftime("%H:%M")
    fim_str = h_fim.strftime("%H:%M")

    # Preparar datas a verificar
    if data_fim and dias_evento:
        datas_a_verificar = pd.date_range(data_ini, data_fim, freq='D')
        datas_a_verificar = [d for d in datas_a_verificar 
                             if mapping.get(d.strftime("%A").upper(), d.strftime("%A").upper()) in dias_evento]
    else:
        datas_a_verificar = [data_ini]

    salas_disponiveis = []

    for sala in salas_ct:
        # Filtro por bloco
        if bloco_pref and bloco_pref != "Qualquer bloco":
            if not sala["NOME"].startswith(bloco_pref):
                continue

        # Filtro por sala específica
        if sala_pref and sala_pref != "Qualquer sala":
            if sala["NOME"] != sala_pref:
                continue

        # Verificar capacidade
        if capacidade > sala["CAPACIDADE"]:
            continue

        # Verificar conflitos de horário
        tem_conflito = False
        conflitos_detalhes = []

        for data in datas_a_verificar:
            dia_port = mapping.get(data.strftime("%A").upper(), data.strftime("%A").upper())
            for a, b, desc in sala["HORARIOS_OCUPADOS_SEMANA"].get(dia_port, []):
                if intervals_overlap(a, b, inicio_str, fim_str):
                    tem_conflito = True
                    conflitos_detalhes.append({
                        "data": data.strftime("%d/%m"),
                        "dia": dia_port,
                        "inicio": a,
                        "fim": b,
                        "ocupacao": desc
                    })

        if not tem_conflito:
            salas_disponiveis.append({
                "SALA": sala["NOME"],
                "CAPACIDADE": sala["CAPACIDADE"],
                "BLOCO": sala["NOME"].split("-")[0] if "-" in sala["NOME"] else sala["NOME"]
            })

    return salas_disponiveis, datas_a_verificar


# ===============================
# INTERFACE STREAMLIT
# ===============================

def interface_usuario(salas_ct, df_processado, df_turmas_raw):

    st.markdown("""
    <style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1a5276;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #5d6d7e;
        margin-bottom: 2rem;
    }
    .card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 1.5rem;
        border-left: 4px solid #1a5276;
        margin-bottom: 1rem;
    }
    .sala-card {
        background-color: #e8f6f3;
        border-radius: 8px;
        padding: 1rem;
        border-left: 4px solid #1abc9c;
        margin-bottom: 0.5rem;
    }
    .sala-card-ocupada {
        background-color: #fdedec;
        border-radius: 8px;
        padding: 1rem;
        border-left: 4px solid #e74c3c;
        margin-bottom: 0.5rem;
    }
    .info-box {
        background-color: #eaf2f8;
        border-radius: 8px;
        padding: 1rem;
        margin: 1rem 0;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="main-header">🏫 Sistema de Alocação de Salas – CT</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Solicite a reserva de salas informando os dados do seu evento</div>', unsafe_allow_html=True)

    # Inicializar session_state
    if 'reservas_lista' not in st.session_state:
        st.session_state.reservas_lista = []
    if 'df_completo' not in st.session_state:
        st.session_state.df_completo = df_processado.copy()
    if 'salas_disponiveis' not in st.session_state:
        st.session_state.salas_disponiveis = []
    if 'evento_dados' not in st.session_state:
        st.session_state.evento_dados = {}

    # ==========================================
    # ETAPA 1: DADOS DO EVENTO
    # ==========================================
    st.markdown("### Dados do Evento")

    with st.container():
        col1, col2 = st.columns(2)
        with col1:
            evento = st.text_input("Nome do evento:*", placeholder="Ex: Reunião de Coordenação")
            nome_solicitante = st.text_input("Seu nome:*", placeholder="Ex: João Silva")
        with col2:
            email = st.text_input("E-mail para contato:", placeholder="exemplo@ufc.br")
            capacidade = st.number_input("Número de participantes:*", min_value=1, value=30, step=1)

    st.divider()

    # ==========================================
    # ETAPA 2: PERÍODO DO EVENTO
    # ==========================================
    st.markdown("### Período do Evento")

    with st.container():
        col1, col2 = st.columns(2)
        with col1:
            data_ini = st.date_input("Data inicial:*", value=dt.date(2026, 3, 2), min_value=dt.date(2026, 1, 1))
        with col2:
            usa_fim = st.selectbox("Evento recorrente?", ["NÃO - Evento único", "SIM - Evento recorrente"])

        if usa_fim == "SIM - Evento recorrente":
            col1, col2 = st.columns(2)
            with col1:
                data_fim = st.date_input("Data final:*", value=dt.date(2026, 7, 7), min_value=data_ini)
            with col2:
                dias_evento = st.multiselect(
                    "Dias da semana:*",
                    DIAS_SEMANA,
                    default=["SEGUNDA"],
                    help="Selecione os dias em que o evento ocorrerá"
                )
        else:
            data_fim = None
            dias_evento = [mapping_dia_semana(data_ini)]

    st.divider()

    # ==========================================
    # ETAPA 3: HORÁRIO DO EVENTO
    # ==========================================
    st.markdown("### Horário do Evento")

    with st.container():
        col1, col2 = st.columns(2)
        with col1:
            h_ini = st.time_input("Horário de início:*", value=dt.time(8, 0))
        with col2:
            h_fim = st.time_input("Horário de término:*", value=dt.time(10, 0))

    st.divider()

    # ==========================================
    # ETAPA 4: PREFERÊNCIA DE SALA (OPCIONAL)
    # ==========================================
    st.markdown("### Preferência de Sala (Opcional)")

    with st.container():
        # Extrair blocos únicos
        blocos = sorted(set(s["NOME"].split("-")[0] for s in salas_ct if "-" in s["NOME"]))
        blocos = ["Qualquer bloco"] + blocos

        col1, col2 = st.columns(2)
        with col1:
            bloco_pref = st.selectbox("Bloco de preferência:", blocos)
        with col2:
            if bloco_pref != "Qualquer bloco":
                salas_bloco = ["Qualquer sala"] + sorted([s["NOME"] for s in salas_ct if s["NOME"].startswith(bloco_pref)])
            else:
                salas_bloco = ["Qualquer sala"] + sorted([s["NOME"] for s in salas_ct])
            sala_pref = st.selectbox("Sala específica:", salas_bloco)

    st.divider()

    # ==========================================
    # BOTÃO: BUSCAR SALAS DISPONÍVEIS
    # ==========================================
    col_btn1, col_btn2, col_btn3 = st.columns([2, 1, 2])
    with col_btn2:
        buscar = st.button("🔍 Buscar Salas Disponíveis", type="primary", use_container_width=True)

    if buscar:
        # Validações
        erros = []
        if not evento or not str(evento).strip():
            erros.append("• Nome do evento é obrigatório")
        if not nome_solicitante or not str(nome_solicitante).strip():
            erros.append("• Nome do solicitante é obrigatório")
        if capacidade <= 0:
            erros.append("• Número de participantes deve ser maior que zero")
        if usa_fim == "SIM - Evento recorrente" and (not data_fim or not dias_evento):
            erros.append("• Para eventos recorrentes, informe data final e dias da semana")
        if time_to_minutes(h_fim) <= time_to_minutes(h_ini):
            erros.append("• Horário de término deve ser posterior ao horário de início")

        if erros:
            st.error("**⚠️ Por favor, corrija os seguintes erros:**\n" + "\n".join(erros))
        else:
            with st.spinner("Verificando disponibilidade de salas..."):
                salas_disp, datas_verif = verificar_disponibilidade(
                    salas_ct, data_ini, data_fim, dias_evento, h_ini, h_fim, capacidade,
                    bloco_pref, sala_pref
                )
                st.session_state.salas_disponiveis = salas_disp
                st.session_state.evento_dados = {
                    "evento": evento,
                    "nome": nome_solicitante,
                    "email": email,
                    "capacidade": capacidade,
                    "data_ini": data_ini,
                    "data_fim": data_fim,
                    "dias_evento": dias_evento,
                    "h_ini": h_ini,
                    "h_fim": h_fim,
                    "datas_verif": datas_verif,
                    "bloco_pref": bloco_pref,
                    "sala_pref": sala_pref
                }

    # ==========================================
    # RESULTADO: SALAS DISPONÍVEIS
    # ==========================================
    if st.session_state.salas_disponiveis:
        st.markdown("---")
        st.markdown("### ✅ Salas Disponíveis")

        dados = st.session_state.evento_dados
        periodo_str = f"{dados['data_ini'].strftime('%d/%m/%Y')}"
        if dados['data_fim']:
            periodo_str += f" a {dados['data_fim'].strftime('%d/%m/%Y')}"
        dias_str = ", ".join(dados['dias_evento'])

        st.markdown(f"""
        <div class="info-box">
        <b>📌 Evento:</b> {dados['evento']} | <b>👥 Participantes:</b> {dados['capacidade']}<br>
        <b>📅 Período:</b> {periodo_str} | <b>📆 Dias:</b> {dias_str}<br>
        <b>⏰ Horário:</b> {dados['h_ini'].strftime('%H:%M')} - {dados['h_fim'].strftime('%H:%M')}
        </div>
        """, unsafe_allow_html=True)

        salas_disp = st.session_state.salas_disponiveis

        if not salas_disp:
            st.warning("⚠️ **Nenhuma sala disponível** para os critérios informados. Tente ajustar a capacidade, horário ou período.")
        else:
            st.success(f"**{len(salas_disp)} sala(s) encontrada(s) disponível(is)**")

            # Agrupar por bloco
            salas_por_bloco = {}
            for s in salas_disp:
                bloco = s["BLOCO"]
                if bloco not in salas_por_bloco:
                    salas_por_bloco[bloco] = []
                salas_por_bloco[bloco].append(s)

            for bloco in sorted(salas_por_bloco.keys()):
                with st.expander(f"🏢 Bloco {bloco} ({len(salas_por_bloco[bloco])} sala(s))", expanded=True):
                    cols = st.columns(3)
                    for idx, sala in enumerate(salas_por_bloco[bloco]):
                        with cols[idx % 3]:
                            st.markdown(f"""
                            <div class="sala-card">
                            <b>📍 {sala['SALA']}</b><br>
                            <span style="color:#1a5276;">👥 Capacidade: {sala['CAPACIDADE']} lugares</span><br>
                            <span style="color:#27ae60;">✅ Disponível</span>
                            </div>
                            """, unsafe_allow_html=True)

                            if st.button(f"✅ Reservar {sala['SALA']}", key=f"btn_reservar_{sala['SALA']}", use_container_width=True):
                                realizar_reserva(sala['SALA'], salas_ct, df_processado)

    elif buscar and not st.session_state.salas_disponiveis:
        # Já mostrou warning acima
        pass

    # ==========================================
    # RESERVAS REALIZADAS NA SESSÃO
    # ==========================================
    if len(st.session_state.reservas_lista) > 0:
        st.markdown("---")
        st.markdown("### 📋 Reservas Realizadas nesta Sessão")
        df_preview = pd.DataFrame(st.session_state.reservas_lista)
        st.dataframe(
            df_preview[["SALA", "DISCIPLINA", "DIAS", "HORARIO_INICIO", "HORARIO_FINAL", "ALUNOS", "PROFESSOR"]],
            use_container_width=True,
            hide_index=True
        )

        # Download e GitHub
        col1, col2 = st.columns(2)
        with col1:
            if st.button("☁️ Salvar no GitHub", type="primary"):
                with st.spinner("Fazendo commit no GitHub..."):
                    sucesso = commit_dados_disciplinas(st.session_state.df_completo)
                    if sucesso:
                        st.success("✅ Arquivo atualizado no GitHub com sucesso!")
                        st.balloons()
                    else:
                        st.error("❌ Falha ao salvar no GitHub. Verifique o token e as permissões.")

        with col2:
            buf_df = BytesIO()
            st.session_state.df_completo.to_excel(buf_df, index=False, engine='openpyxl')
            buf_df.seek(0)
            st.download_button(
                "📥 Baixar cópia local (backup)",
                data=buf_df,
                file_name="dados_disciplinas_backup.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


def mapping_dia_semana(data):
    mapping = {'MONDAY': 'SEGUNDA', 'TUESDAY': 'TERÇA', 'WEDNESDAY': 'QUARTA',
               'THURSDAY': 'QUINTA', 'FRIDAY': 'SEXTA', 'SATURDAY': 'SÁBADO', 'SUNDAY': 'DOMINGO'}
    return mapping.get(data.strftime("%A").upper(), data.strftime("%A").upper())


def realizar_reserva(sala_escolhida, salas_ct, df_processado):
    """Realiza a reserva da sala selecionada."""
    dados = st.session_state.evento_dados
    sala_info = next((s for s in salas_ct if s["NOME"] == sala_escolhida), None)

    if not sala_info:
        st.error("Erro: Sala não encontrada.")
        return

    inicio_str = dados['h_ini'].strftime("%H:%M")
    fim_str = dados['h_fim'].strftime("%H:%M")
    mapping = {'MONDAY': 'SEGUNDA', 'TUESDAY': 'TERÇA', 'WEDNESDAY': 'QUARTA',
               'THURSDAY': 'QUINTA', 'FRIDAY': 'SEXTA', 'SATURDAY': 'SÁBADO', 'SUNDAY': 'DOMINGO'}

    nome_evento = dados['evento'].strip()
    desc = f"RESERVA_MANUAL - {nome_evento}"

    # Adiciona à sala em memória
    datas_list = []
    for data in dados['datas_verif']:
        dia_port = mapping.get(data.strftime("%A").upper(), data.strftime("%A").upper())
        sala_info["RESERVAS"].append((data, inicio_str, fim_str, desc))
        sala_info["HORARIOS_OCUPADOS_SEMANA"].setdefault(dia_port, []).append(
            (inicio_str, fim_str, desc))
        sala_info["HORARIOS_OCUPADOS"].add(f"{inicio_str} - {fim_str}")
        datas_list.append(data)

    # Cria registro da reserva para o DataFrame
    nova_reserva = {
        "CURSO": "RESERVA",
        "CODIGO": "MANUAL",
        "SALA": sala_escolhida,
        "DISCIPLINA": nome_evento,
        "TURMA": "N/A",
        "DIAS": ",".join([mapping.get(d.strftime("%A").upper(), d.strftime("%A").upper()) for d in dados['datas_verif']]),
        "HORARIO_INICIO": inicio_str,
        "HORARIO_FINAL": fim_str,
        "ALUNOS": dados['capacidade'],
        "PROFESSOR": dados['nome'],
        "CAPACIDADE": sala_info["CAPACIDADE"],
        "DATAS": datas_list,
        "DESCRICAO": desc,
        "TIPO": "RESERVA_MANUAL"
    }

    st.session_state.reservas_lista.append(nova_reserva)
    df_reservas = pd.DataFrame(st.session_state.reservas_lista)
    st.session_state.df_completo = pd.concat([df_processado, df_reservas], ignore_index=True)

    # Limpa salas disponíveis para forçar nova busca
    st.session_state.salas_disponiveis = []

    st.success(f"✅ Sala **{sala_escolhida}** reservada com sucesso para o evento {nome_evento}!")
    st.info(f"📅 {len(datas_list)} dia(s) reservado(s). As reservas serão incluídas no arquivo de dados.")
    st.rerun()


# -----------------------  Main  -----------------------
def main():
    st.set_page_config(
        page_title="Sistema de Alocação de Salas - CT",
        page_icon="🏫",
        layout="wide"
    )

    with st.spinner("Carregando dados..."):
        df_turmas_raw, salas_info = carregar_dados()
        salas_ct = criar_lista_salas(salas_info)
        todas_as_datas = gerar_datas(df_turmas_raw)
        df_dados = processar_alocacoes(df_turmas_raw, todas_as_datas, salas_ct)

    interface_usuario(salas_ct, df_dados, df_turmas_raw)

if __name__ == "__main__":
    main()