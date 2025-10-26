# app.py
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
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR

CAMINHO_SALAS      = DATA_DIR / "SALAS - COPIA.xlsx"
CAMINHO_DISCIPLINAS = DATA_DIR / "Resultados_Gerais.xlsx"
OUTPUT_DIR          = BASE_DIR / "resultados"

DIAS_SEMANA = ["SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"]
INDICE_DIAS = {d: i for i, d in enumerate(DIAS_SEMANA)}


# ===============================
# FUNÇÕES AUXILIARES
# ===============================
@st.cache_data(show_spinner=False)
def carregar_dados():
    """Carrega salas e turmas."""
    for arq, nome in ((CAMINHO_SALAS, "salas"), (CAMINHO_DISCIPLINAS, "disciplinas")):
        if not arq.exists():
            st.error(f"❌ Arquivo de {nome} não encontrado: {arq}")
            st.stop()
    df_salas  = pd.read_excel(CAMINHO_SALAS)
    df_turmas = pd.read_excel(CAMINHO_DISCIPLINAS)
    return df_salas, df_turmas


def normalizar_colunas(df):
    """Tira acentos, espaços duplos e deixa maiúscula."""
    df.columns = (
        df.columns
          .str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("utf-8")
          .str.upper()
          .str.strip()
          .str.replace("  ", " ")
    )
    return df


def criar_lista_salas(df_salas):
    """Cria estrutura de salas."""
    df_salas = normalizar_colunas(df_salas.copy())
    if "SALAS" not in df_salas.columns or "CAPACIDADE" not in df_salas.columns:
        st.error("Planilha de salas deve conter colunas SALAS e CAPACIDADE")
        st.stop()
    return [
        {
            "NOME": str(row["SALAS"]).strip(),
            "CAPACIDADE": int(row["CAPACIDADE"]),
            "DATAS": set(),
            "HORARIOS_OCUPADOS": set(),          # strings "HH:MM–HH:MM"
            "HORARIOS_INICIO": set(),            # datetime.time
            "HORARIOS_FIM": set(),               # datetime.time
        }
        for _, row in df_salas.iterrows()
    ]


def extrair_datas(df_turmas):
    """Lê datas de início/fim da planilha."""
    df_turmas = normalizar_colunas(df_turmas)
    try:
        inicio_str = str(df_turmas.iloc[0]["DATA INICIO"]).strip()
        fim_str    = str(df_turmas.iloc[0]["DATA FINAL"]).strip()
        inicio = dt.date(*map(int, inicio_str.split(",")))
        fim    = dt.date(*map(int, fim_str.split(",")))
    except Exception as e:
        st.error(f"Erro ao ler datas de início/fim: {e}")
        st.stop()
    return pd.date_range(inicio, fim)


def hora_to_min(t: dt.time) -> int:
    return t.hour * 60 + t.minute


def min_to_hora(m: int) -> dt.time:
    return dt.time(m // 60, m % 60)


def sobrepoe(h1_ini, h1_fim, h2_ini, h2_fim) -> bool:
    """Verifica se dois intervalos de tempo se sobrepõem."""
    t1_ini, t1_fim = hora_to_min(h1_ini), hora_to_min(h1_fim)
    t2_ini, t2_fim = hora_to_min(h2_ini), hora_to_min(h2_fim)
    return t1_ini < t2_fim and t2_ini < t1_fim


def processar_alocacoes(df_turmas, todas_as_datas, salas_ct):
    """Processa turmas já alocadas."""
    df_turmas = normalizar_colunas(df_turmas.copy())
    dados = []

    for _, aloc in df_turmas.iterrows():
        if str(aloc.get("STATUS", "")).strip().upper() != "ALOCADA":
            continue

        sala = str(aloc["SALA"]).strip()
        dias = str(aloc.get("DIAS", "")).strip().upper()
        if not dias:
            continue

        # encontra capacidade
        capacidade = next((s["CAPACIDADE"] for s in salas_ct if s["NOME"] == sala), None)
        if capacidade is None:
            continue

        # Converte dias para índices 0-6
        dias_lista = dias.split()
        indices = [INDICE_DIAS[d] for d in dias_lista if d in INDICE_DIAS]
        datas   = todas_as_datas[todas_as_datas.dayofweek.isin(indices)]

        # Converte horários
        try:
            h_ini = pd.to_datetime(aloc["HORARIO INICIO"], format="%H:%M").time()
            h_fim = pd.to_datetime(aloc["HORARIO FINAL"], format="%H:%M").time()
        except Exception:
            continue

        dados.append({
            "CURSO": aloc.get("CURSO", ""),
            "CODIGO": aloc.get("CODIGO", ""),
            "SALA": sala,
            "DISCIPLINA": aloc.get("DISCIPLINA", ""),
            "TURMA": aloc.get("TURMA", ""),
            "DIAS": dias,
            "HORARIO INICIO": h_ini,
            "HORARIO FINAL": h_fim,
            "HORARIOS": aloc.get("HORARIO", ""),
            "ALUNOS": aloc.get("ALUNOS", ""),
            "PROFESSOR": aloc.get("PROFESSOR", ""),
            "CAPACIDADE": capacidade,
            "DATAS": datas,
        })

        # atualiza sala
        for s in salas_ct:
            if s["NOME"] == sala:
                s["DATAS"].update(datas)
                s["HORARIOS_OCUPADOS"].add(f"{h_ini.strftime('%H:%M')}–{h_fim.strftime('%H:%M')}")
                s["HORARIOS_INICIO"].add(h_ini)
                s["HORARIOS_FIM"].add(h_fim)

    return pd.DataFrame(dados)


def exportar_dados(df):
    """Retorna Excel em bytes e salva localmente."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    caminho = OUTPUT_DIR / "dados_disciplinas.xlsx"
    df.to_excel(caminho, index=False)
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer, caminho


# ===============================
# INTERFACE STREAMLIT
# ===============================
def interface_interativa(salas_ct, df_processado):
    st.header("🎯 Solicitação de Sala")

    blocos = sorted({s["NOME"][:3] for s in salas_ct})
    bloco_selec = st.selectbox("Selecione o bloco:", blocos)

    salas_filtradas = [s["NOME"] for s in salas_ct if s["NOME"].startswith(bloco_selec)]
    sala_escolha = st.selectbox("Selecione a sala:", salas_filtradas)

    data_escolha = st.date_input("Selecione a data:")
    h_ini_sel = st.time_input("Horário de início:", value=dt.time(8, 0))
    h_fim_sel = st.time_input("Horário de término:", value=dt.time(10, 0))

    sala_info = next((s for s in salas_ct if s["NOME"] == sala_escolha), None)

    if sala_info and sala_info["HORARIOS_OCUPADOS"]:
        st.info("🕓 Horários ocupados: " + ", ".join(sorted(sala_info["HORARIOS_OCUPADOS"])))
    else:
        st.success("✅ Nenhum horário ocupado encontrado para esta sala.")

    if st.button("📅 Solicitar Sala"):
        if not sala_info:
            st.error("Sala não encontrada.")
            return

        # verifica conflito
        conflito = any(
            sobrepoe(h_ini_sel, h_fim_sel, h_ini_ex, h_fim_ex)
            for h_ini_ex, h_fim_ex in zip(sala_info["HORARIOS_INICIO"], sala_info["HORARIOS_FIM"])
        )
        if conflito:
            st.error("❌ A sala está ocupada no horário selecionado.")
        else:
            st.success(
                f"✅ Solicitação registrada para **{sala_escolha}** em {data_escolha} "
                f"({h_ini_sel.strftime('%H:%M')}–{h_fim_sel.strftime('%H:%M')})"
            )
            # registra
            sala_info["HORARIOS_OCUPADOS"].add(f"{h_ini_sel.strftime('%H:%M')}–{h_fim_sel.strftime('%H:%M')}")
            sala_info["HORARIOS_INICIO"].add(h_ini_sel)
            sala_info["HORARIOS_FIM"].add(h_fim_sel)

    # download
    buffer, _ = exportar_dados(df_processado)
    st.download_button(
        label="📥 Baixar Excel Processado",
        data=buffer,
        file_name="dados_disciplinas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ===============================
# MAIN
# ===============================
def main():
    st.title("🏫 Sistema de Alocação de Salas – CT")
    with st.spinner("Carregando dados..."):
        df_salas, df_turmas = carregar_dados()
        salas_ct = criar_lista_salas(df_salas)
        datas = extrair_datas(df_turmas)
        df_dados = processar_alocacoes(df_turmas, datas, salas_ct)

    st.success("✅ Dados carregados e processados com sucesso!")
    st.divider()
    interface_interativa(salas_ct, df_dados)


if __name__ == "__main__":
    main()
