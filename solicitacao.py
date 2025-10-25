import datetime as dt
from pathlib import Path

import pandas as pd
import streamlit as st
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


def exportar_dados(df, output_dir, nome="dados_disciplinas.xlsx"):
    """Exporta os dados processados para Excel."""
    output_dir.mkdir(parents=True, exist_ok=True)
    caminho = output_dir / nome
    df.to_excel(caminho, index=False)
    st.success(f"📂 Arquivo exportado: `{caminho}`")
    return caminho


# ===============================
# INTERFACE STREAMLIT
# ===============================

def interface_interativa(salas_ct):
    """Interface para seleção de bloco, sala, data e horário."""
    st.header("🎯 Solicitação de Sala")

    blocos = sorted(set(s["NOME"].split()[0] for s in salas_ct))
    bloco_selecionado = st.selectbox("Selecione o bloco:", blocos)

    salas_disponiveis = [s["NOME"] for s in salas_ct if bloco_selecionado in s["NOME"]]
    sala_escolhida = st.selectbox("Selecione a sala:", salas_disponiveis)

    data_escolhida = st.date_input("Selecione a data:")
    horario_inicio = st.time_input("Horário de início:")
    horario_fim = st.time_input("Horário de término:")

    sala_info = next((s for s in salas_ct if s["NOME"] == sala_escolhida), None)
    if sala_info:
        if sala_info["HORARIOS_OCUPADOS"]:
            st.info(f"🕓 Horários ocupados: {', '.join(sala_info['HORARIOS_OCUPADOS'])}")
        else:
            st.success("✅ Nenhum horário ocupado encontrado para esta sala.")

    if st.button("📅 Solicitar Sala"):
        if not sala_info:
            st.error("Sala não encontrada.")
            return

        conflito = any(
            horario_inicio.strftime("%H:%M") in h or horario_fim.strftime("%H:%M") in h
            for h in sala_info["HORARIOS_OCUPADOS"]
        )

        if conflito:
            st.error("❌ A sala está ocupada no horário selecionado.")
        else:
            st.success(f"✅ Solicitação registrada para **{sala_escolhida}** em {data_escolhida} "
                       f"({horario_inicio.strftime('%H:%M')}–{horario_fim.strftime('%H:%M')})")
            sala_info["HORARIOS_OCUPADOS"].add(f"{horario_inicio.strftime('%H:%M')} - {horario_fim.strftime('%H:%M')}")


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

    st.success("✅ Dados carregados com sucesso!")
    exportar_dados(df_dados, OUTPUT_DIR)

    st.divider()
    interface_interativa(salas_ct)


if __name__ == "__main__":
    main()
