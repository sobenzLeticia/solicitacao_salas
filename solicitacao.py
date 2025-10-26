import datetime as dt
from pathlib import Path
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter


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

# ===============================
# INTERFACE STREAMLIT
# ===============================

def interface_interativa(salas_ct, df_processado):
    """Interface para seleção de bloco, sala, data e horário + download."""
    st.header("🎯 Solicitação de Sala")

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
            st.stop()

        # --- NOVA LÓGICA DE CONFLITO ---
        def _para_minutos(t: dt.time) -> int:
            return t.hour * 60 + t.minute

        inicio_sol = _para_minutos(horario_inicio)
        fim_sol = _para_minutos(horario_fim)

        conflito = False
        for h_ocup in sala_info["HORARIOS_OCUPADOS"]:
            # h_ocup pode estar no formato "HH:MM - HH:MM"
            try:
                h_ini_str, h_fim_str = h_ocup.split(" - ")
                ini_ocup = _para_minutos(dt.time.fromisoformat(h_ini_str))
                fim_ocup = _para_minutos(dt.time.fromisoformat(h_fim_str))
            except ValueError:
                # caso ainda esteja no formato antigo "HH:MM - HH:MM" sem o split
                continue

            # detecta sobreposição
            if inicio_sol < fim_ocup and fim_sol > ini_ocup:
                conflito = True
                break

        if conflito:
            st.error("❌ A sala está ocupada no horário selecionado.")
        else:
            st.success(f"✅ Solicitação registrada para **{sala_escolhida}** em {data_escolhida} "
                       f"({horario_inicio.strftime('%H:%M')}–{horario_fim.strftime('%H:%M')})")
            sala_info["HORARIOS_OCUPADOS"].add(f"{horario_inicio.strftime('%H:%M')} - {horario_fim.strftime('%H:%M')}")

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
