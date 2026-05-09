import datetime as dt
from pathlib import Path
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter
import smtplib  # [ADICIONADO]
from email.mime.multipart import MIMEMultipart  # [ADICIONADO]
from email.mime.text import MIMEText  # [ADICIONADO]


# ===============================
# CONFIGURAÇÕES GERAIS
# ===============================

# Caminhos relativos dentro do repositório
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR

CAMINHO_SALAS = DATA_DIR / "SALAS - COPIA.xlsx"
CAMINHO_DISCIPLINAS = DATA_DIR / "Resultados_Gerais.xlsx"
OUTPUT_DIR = BASE_DIR / "resultados"

DIAS_SEMANA = ["SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"]
INDICE_DIAS = {d: i for i, d in enumerate(DIAS_SEMANA)}


# ===============================
# FUNÇÕES DE LEITURA E PROCESSAMENTO
# ===============================

@st.cache_data(show_spinner=False)
def carregar_dados():
    """Carrega os dados de salas e turmas do repositório."""
    if not CAMINHO_SALAS.exists():
        st.error(f"❌ Arquivo de salas não encontrado em: {CAMINHO_SALAS}")
        st.stop()

    if not CAMINHO_DISCIPLINAS.exists():
        st.error(f"❌ Arquivo de disciplinas não encontrado em: {CAMINHO_DISCIPLINAS}")
        st.stop()

    df_salas = pd.read_excel(CAMINHO_SALAS)
    df_turmas = pd.read_excel(CAMINHO_DISCIPLINAS)
    return df_salas, df_turmas


def criar_lista_salas(df_salas):
    """Cria estrutura de salas com capacidade e controle de horários."""
    return [
        {
            "NOME": row["SALAS"],
            "CAPACIDADE": row["CAPACIDADE"],
            "DATAS": set(),
            "HORARIOS_OCUPADOS": set(),
        }
        for _, row in df_salas.iterrows()
    ]


def gerar_datas(df_turmas):
    """Gera todas as datas entre o início e o fim definidos na planilha."""
    data_inicio = list(map(int, df_turmas.iloc[0, 13].split(",")))
    data_final = list(map(int, df_turmas.iloc[0, 14].split(",")))
    return pd.date_range(dt.date(*data_inicio), dt.date(*data_final))


def processar_alocacoes(df_turmas, todas_as_datas, salas_ct):
    """Processa as turmas e cria DataFrame com dados das disciplinas."""
    dados = []

    for _, aloc in df_turmas.iterrows():
        if aloc.get("STATUS") != "Alocada":
            continue

        sala = aloc["SALA"]
        dias = aloc.get("DIAS")
        if pd.isna(dias):
            continue

        capacidade = next(
            (s["CAPACIDADE"] for s in salas_ct if s["NOME"] == sala),
            None
        )

        dias_lista = dias.split()
        indices = [INDICE_DIAS.get(dia) for dia in dias_lista if dia in INDICE_DIAS]
        datas = todas_as_datas[todas_as_datas.dayofweek.isin(indices)]

        dados.append({
            "CURSO": aloc["CURSO"],
            "CODIGO": aloc["CODIGO"],
            "SALA": sala,
            "DISCIPLINA": aloc["DISCIPLINA"],
            "TURMA": aloc["TURMA"],
            "DIAS": dias,
            "HORARIO INICIO": aloc["HORARIO INICIO"],
            "HORARIO FINAL": aloc["HORARIO FINAL"],
            "HORARIOS": aloc["HORARIO"],
            "ALUNOS": aloc["ALUNOS"],
            "PROFESSOR": aloc["PROFESSOR"],
            "CAPACIDADE": capacidade,
            "DATAS": datas,
        })

        for s in salas_ct:
            if s["NOME"] == sala:
                s["DATAS"].update(datas)
                s["HORARIOS_OCUPADOS"].add(aloc["HORARIO"])
                s["HORARIO INICIO"].add(aloc["HORARIO INICIO"])
                s["HORARIO FINAL"].add(aloc["HORARIO FINAL"])

    return pd.DataFrame(dados)


def exportar_dados(df):
    """Exporta o DataFrame processado para bytes Excel e também salva localmente."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    caminho = OUTPUT_DIR / "dados_disciplinas.xlsx"
    df.to_excel(caminho, index=False)

    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer, caminho


def gerar_intervalos(inicio, fim, meio):
    horarios_intermediarios = []
    horario_atual = inicio
    while horario_atual <= fim:
        horarios_intermediarios.append(horario_atual)
        horario_atual += meio
    return horarios_intermediarios


# [ADICIONADO] Função para enviar e-mail
def enviar_email_solicitacao(nome, email_cliente, evento, sala, data, h_ini, h_fim, capacity):
    """Envia e-mail de confirmação da solicitação de sala."""
    remetente = "reservasalact-naoresponda@ufc.br"
    destinatario = "reservasalact@ufc.br"
    senha = "rmqz ohnf oppx zpwo"
    assunto = "Teste - Solicitação de Salas"
    
    corpo = f"""
    <html>
    <body>
        <h3>Solicitação de Reserva de Sala - CT</h3>
        <p><strong>Nome do solicitante:</strong> {nome}</p>
        <p><strong>E-mail:</strong> {email_cliente}</p>
        <p><strong>Evento:</strong> {evento}</p>
        <p><strong>Sala:</strong> {sala}</p>
        <p><strong>Data:</strong> {data.strftime('%d/%m/%Y')}</p>
        <p><strong>Horário:</strong> {h_ini.strftime('%H:%M')} - {h_fim.strftime('%H:%M')}</p>
        <p><strong>Capacidade solicitada:</strong> {capacity}</p>
        <hr>
        <p><em>Este é um e-mail automático. Por favor, não responda.</em></p>
    </body>
    </html>
    """
    
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto
    msg.attach(MIMEText(corpo, 'html'))
    
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        server.sendmail(remetente, destinatario, msg.as_string())
        server.quit()
        return True
    except Exception as e:
        st.error(f"❌ Erro ao enviar e-mail: {str(e)}")
        return False


# ===============================
# INTERFACE STREAMLIT
# ===============================

def interface_interativa(salas_ct, df_processado):
    """Interface para seleção de bloco, sala, data e horário + download."""
    st.header("🎯 Solicitação de Sala")

    # [ADICIONADO] Campos de entrada para e-mail (mesmas informações do segundo código)
    evento = st.text_input("Digite o nome do evento:")
    nome = st.text_input("Digite seu nome:")
    email_cliente = st.text_input("Digite seu email:")
    capacity = st.number_input("Capacidade:", min_value=0, value=0)

    # Extrai blocos únicos (apenas a primeira parte do nome da sala)
    blocos = ["707","717","726","727"]
    bloco_selecionado = st.selectbox("Selecione o bloco:", blocos)

    # Filtra salas do bloco escolhido
    salas_filtradas = [s["NOME"] for s in salas_ct if s["NOME"].startswith(bloco_selecionado)]
    sala_escolhida = st.selectbox("Selecione a sala:", salas_filtradas)

    data_escolhida = st.date_input("Selecione a data:")
    horario_inicio = st.time_input("Horário de início:")
    horario_fim = st.time_input("Horário de término:")

    sala_info = next((s for s in salas_ct if s["NOME"] == sala_escolhida), None)

    if sala_info:
        if sala_info["HORARIOS_OCUPADOS"]:
            st.info(f"🕓 Horários ocupados: {', '.join(sorted(sala_info['HORARIOS_OCUPADOS']))}")
        else:
            st.success("✅ Nenhum horário ocupado encontrado para esta sala.")

    if st.button("📅 Solicitar Sala"):
        if not sala_info:
            st.error("Sala não encontrada.")
            return

        # [ADICIONADO] Verifica capacidade
        if capacity > sala_info["CAPACIDADE"]:
            ociosidade = (-1) * (sala_info["CAPACIDADE"] - capacity)
            st.error(f"❌ Conflito de Ociosidade: Capacidade excedida em {abs(ociosidade)} alunos (Capacidade da sala: {sala_info['CAPACIDADE']}, Solicitado: {capacity})")
            return

        conflito = any(
            horario_inicio.strftime("%H:%M") in h or horario_fim.strftime("%H:%M") in h
            for h in sala_info["HORARIOS_OCUPADOS"]
        )
        intervalo = dt.timedelta(minutes=1)
        ini = sala_info["HORARIOS INICIO"]
        f = sala_info["HORARIO FINAL"]
        horario_intervalo = gerar_intervalos(ini, f, intervalo)
        amostra = [True if horario_inicio in h or horario_fim in h else False for h in horario_intervalo]
        conflito_2 = any(amostra)

        if conflito or conflito_2:
            st.error("❌ A sala está ocupada no horário selecionado.")
        else:
            st.success(f"✅ Solicitação registrada para **{sala_escolhida}** em {data_escolhida} "
                       f"({horario_inicio.strftime('%H:%M')}–{horario_fim.strftime('%H:%M')})")
            sala_info["HORARIOS_OCUPADOS"].add(f"{horario_inicio.strftime('%H:%M')} - {horario_fim.strftime('%H:%M')}")

            # [ADICIONADO] Envia e-mail após confirmação bem-sucedida
            nome_evento = evento.strip() if evento and str(evento).strip() else "Evento Manual"
            email_enviado = enviar_email_solicitacao(
                nome=nome,
                email_cliente=email_cliente,
                evento=nome_evento,
                sala=sala_escolhida,
                data=data_escolhida,
                h_ini=horario_inicio,
                h_fim=horario_fim,
                capacity=capacity
            )
            if email_enviado:
                st.info("📧 E-mail de confirmação enviado com sucesso!")

    # Botão de download
    st.download_button(
        label="📥 Baixar Excel Processado",
        data=exportar_dados(df_processado)[0],
        file_name="dados_disciplinas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ===============================
# APP PRINCIPAL
# ===============================

def main():
    st.title("🏫 Sistema de Alocação de Salas – CT")

    with st.spinner("Carregando dados..."):
        df_salas, df_turmas = carregar_dados()
        salas_ct = criar_lista_salas(df_salas)
        todas_as_datas = gerar_datas(df_turmas)
        df_dados = processar_alocacoes(df_turmas, todas_as_datas, salas_ct)

    st.success("✅ Dados carregados e processados com sucesso!")
    st.divider()
    interface_interativa(salas_ct, df_dados)


if __name__ == "__main__":
    main()
