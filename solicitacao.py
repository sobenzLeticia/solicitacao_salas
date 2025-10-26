# app.py
import datetime as dt
from pathlib import Path
import pandas as pd
import streamlit as st
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter
from collections import defaultdict

# ---------- CONFIGURAÇÕES GERAIS ----------
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR

CAMINHO_SALAS       = DATA_DIR / "SALAS - COPIA.xlsx"
CAMINHO_DISCIPLINAS = DATA_DIR / "Resultados_Gerais.xlsx"
OUTPUT_DIR          = BASE_DIR / "resultados"

DIAS_SEMANA = ["SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"]
INDICE_DIAS = {d: i for i, d in enumerate(DIAS_SEMANA)}
HORAS_MINUTOS = [f"{h:02d}:{m:02d}" for h in range(7, 23) for m in (0, 30)]

# ---------- FUNÇÕES AUXILIARES ----------
@st.cache_data(show_spinner=False)
def carregar_dados():
    for arq, nome in ((CAMINHO_SALAS, "salas"), (CAMINHO_DISCIPLINAS, "disciplinas")):
        if not arq.exists():
            st.error(f"❌ Arquivo de {nome} não encontrado: {arq}")
            st.stop()
    df_salas  = pd.read_excel(CAMINHO_SALAS)
    df_turmas = pd.read_excel(CAMINHO_DISCIPLINAS)
    return df_salas, df_turmas

def normalizar_colunas(df):
    df.columns = (
        df.columns
          .str.normalize("NFKD").str.encode("ascii", errors="ignore").str.decode("utf-8")
          .str.upper()
          .str.strip()
          .str.replace("  ", " ")
    )
    return df

def criar_lista_salas(df_salas):
    df_salas = normalizar_colunas(df_salas.copy())
    if "SALAS" not in df_salas.columns or "CAPACIDADE" not in df_salas.columns:
        st.error("Planilha de salas deve conter colunas SALAS e CAPACIDADE")
        st.stop()
    return [
        {
            "NOME": str(row["SALAS"]).strip(),
            "CAPACIDADE": int(row["CAPACIDADE"]),
            "DATAS": set(),
            "HORARIOS_OCUPADOS": set(),
            "HORARIOS_INICIO": set(),
            "HORARIOS_FIM": set(),
        }
        for _, row in df_salas.iterrows()
    ]

def extrair_datas(df_turmas):
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
    t1_ini, t1_fim = hora_to_min(h1_ini), hora_to_min(h1_fim)
    t2_ini, t2_fim = hora_to_min(h2_ini), hora_to_min(h2_fim)
    return t1_ini < t2_fim and t2_ini < t1_fim

def processar_alocacoes(df_turmas, todas_as_datas, salas_ct):
    df_turmas = normalizar_colunas(df_turmas.copy())
    dados = []

    for _, aloc in df_turmas.iterrows():
        if str(aloc.get("STATUS", "")).strip().upper() != "ALOCADA":
            continue

        sala = str(aloc["SALA"]).strip()
        dias = str(aloc.get("DIAS", "")).strip().upper()
        if not dias:
            continue

        capacidade = next((s["CAPACIDADE"] for s in salas_ct if s["NOME"] == sala), None)
        if capacidade is None:
            continue

        dias_lista = dias.split()
        indices = [INDICE_DIAS[d] for d in dias_lista if d in INDICE_DIAS]
        datas   = todas_as_datas[todas_as_datas.dayofweek.isin(indices)]

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
            "HORARIO": aloc.get("HORARIO", ""),
            "ALUNOS": aloc.get("ALUNOS", ""),
            "PROFESSOR": aloc.get("PROFESSOR", ""),
            "CAPACIDADE": capacidade,
            "DATAS": datas,
        })

        for s in salas_ct:
            if s["NOME"] == sala:
                s["DATAS"].update(datas)
                s["HORARIOS_OCUPADOS"].add(f"{h_ini.strftime('%H:%M')}–{h_fim.strftime('%H:%M')}")
                s["HORARIOS_INICIO"].add(h_ini)
                s["HORARIOS_FIM"].add(h_fim)

    return pd.DataFrame(dados)

def exportar_dados(df):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    caminho = OUTPUT_DIR / "dados_disciplinas.xlsx"
    df.to_excel(caminho, index=False)
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer, caminho

# ---------- GERAÇÃO DA PLANILHA DE HORÁRIOS ----------
def gera_excel_horarios(df_processado, salas_ct):
    def split_horario(horario_completo):
        partes = horario_completo.split()
        if len(partes) < 2 or "-" not in partes[1]:
            return []
        dia, hora_str = partes[0], partes[1]
        hi_str, hf_str = hora_str.split("-")
        hi = dt.datetime.strptime(hi_str, "%H:%M:%S")
        intervals = [
            f"{dia} {hi.strftime('%H:%M')} - {(hi.replace(minute=30)).strftime('%H:%M')}",
            f"{dia} {(hi.replace(minute=30)).strftime('%H:%M')} - {dt.datetime.strptime(hf_str, '%H:%M:%S').strftime('%H:%M')}"
        ]
        return intervals

    horarios_por_sala = defaultdict(lambda: defaultdict(dict))
    for aloc in df_processado.to_dict("records"):
        if not aloc.get("SALA"):
            continue
        sala_nome = aloc["SALA"]
        info = f"{aloc['CODIGO']} - {aloc['DISCIPLINA']} - {aloc['TURMA']} - {aloc['PROFESSOR']}"
        for bloco in aloc.get("HORARIO", "").split(","):
            bloco = bloco.strip()
            if not bloco:
                continue
            dia = bloco.split()[0]
            for intervalo in split_horario(bloco):
                _, horario_fmt = intervalo.split(" ", 1)
                horarios_por_sala[sala_nome][dia][horario_fmt] = info

    borda = Border(left=Side(style='thin'), right=Side(style='thin'),
                   top=Side(style='thin'), bottom=Side(style='thin'))
    alinh_centro = Alignment(horizontal='center', vertical='center', wrap_text=True)
    font_padrao = Font(size=10)

    wb = Workbook()
    wb.remove(wb.active)

    for sala in salas_ct:
        sala_nome = sala["NOME"]
        ws = wb.create_sheet(title=f"Horário {sala_nome}"[:31])

        # Cabeçalho
        info = f"Centro de Tecnologia | {sala_nome} | Capacidade: {sala['CAPACIDADE']}"
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(DIAS_SEMANA)+1)
        c = ws.cell(row=1, column=1, value=info)
        c.font = Font(bold=True, size=12)
        c.alignment = Alignment(horizontal='center', vertical='center')

        ws.cell(row=2, column=1, value="Horário").font = Font(bold=True)
        for col, dia in enumerate(DIAS_SEMANA, start=2):
            ws.cell(row=2, column=col, value=dia).font = Font(bold=True)

        for row, hora in enumerate(HORAS_MINUTOS, start=3):
            ws.cell(row=row, column=1, value=hora)

        # Preenche dados
        if sala_nome in horarios_por_sala:
            for dia, horarios in horarios_por_sala[sala_nome].items():
                try:
                    col = DIAS_SEMANA.index(dia) + 2
                except ValueError:
                    continue
                for hora, info in horarios.items():
                    if hora in HORAS_MINUTOS:
                        row_idx = HORAS_MINUTOS.index(hora) + 3
                        ws.cell(row=row_idx, column=col, value=info)

        # Merge células contíguas iguais (simples)
        for col in range(2, len(DIAS_SEMANA)+2):
            start_row = 3
            curr = ws.cell(row=3, column=col).value
            for row in range(3, len(HORAS_MINUTOS)+3):
                val = ws.cell(row=row, column=col).value
                if val != curr:
                    if curr not in (None, "") and row-1 > start_row:
                        ws.merge_cells(start_row=start_row, start_column=col, end_row=row-1, end_column=col)
                    start_row = row
                    curr = val
            if curr not in (None, "") and len(HORAS_MINUTOS)+2 > start_row:
                ws.merge_cells(start_row=start_row, start_column=col, end_row=len(HORAS_MINUTOS)+2, end_column=col)

        # Estilo
        for row in ws.iter_rows(min_row=1, max_row=len(HORAS_MINUTOS)+2,
                                min_col=1, max_col=len(DIAS_SEMANA)+1):
            for cell in row:
                cell.border = borda
                cell.alignment = alinh_centro
                cell.font = font_padrao
        for col in range(1, len(DIAS_SEMANA)+2):
            ws.column_dimensions[get_column_letter(col)].width = 20

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ---------- INTERFACE STREAMLIT ----------
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
            sala_info["HORARIOS_OCUPADOS"].add(f"{h_ini_sel.strftime('%H:%M')}–{h_fim_sel.strftime('%H:%M')}")
            sala_info["HORARIOS_INICIO"].add(h_ini_sel)
            sala_info["HORARIOS_FIM"].add(h_fim_sel)

    if st.button("📥 Gerar planilha de horários"):
        buffer = gera_excel_horarios(df_processado, salas_ct)
        st.download_button(
            label="📥 Baixar horários das salas",
            data=buffer,
            file_name="horarios_salas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ---------- MAIN ----------
def main():
    st.set_page_config(page_title="Alocador de Salas", layout="wide")
    st.title("Alocador de Salas – CT")
    df_salas, df_turmas = carregar_dados()
    salas_ct = criar_lista_salas(df_salas)
    datas = extrair_datas(df_turmas)
    df_processado = processar_alocacoes(df_turmas, datas, salas_ct)

    tab1, tab2 = st.tabs(["Solicitação de Sala", "Dados Processados"])
    with tab1:
        interface_interativa(salas_ct, df_processado)
    with tab2:
        st.subheader("Pré-visualização dos dados processados")
        st.dataframe(df_processado)
        buffer, _ = exportar_dados(df_processado)
        st.download_button(
            label="📥 Baixar Excel Processado",
            data=buffer,
            file_name="dados_disciplinas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
