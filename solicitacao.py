import re
import datetime as dt
from pathlib import Path
from io import BytesIO
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter

# -----------------------
# Configurações
# -----------------------
BASE_DIR = Path(__file__).parent
CAMINHO_SALAS = BASE_DIR / "SALAS - COPIA.xlsx"
CAMINHO_DISCIPLINAS = BASE_DIR / "Resultados_Gerais.xlsx"
OUTPUT_DIR = BASE_DIR / "resultados"
OUTPUT_DIR.mkdir(exist_ok=True)

DIAS_SEMANA = ["SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"]
INDICE_DIAS = {d: i for i, d in enumerate(DIAS_SEMANA)}

# -----------------------
# Utils horário
# -----------------------
def str_to_time(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    if isinstance(s, dt.time):
        return s
    s = str(s).strip()
    # tenta formatos comuns
    for fmt in ("%H:%M:%S", "%H:%M", "%H.%M"):
        try:
            return dt.datetime.strptime(s, fmt).time()
        except Exception:
            pass
    # tentar extrair dígitos
    s2 = re.sub(r'[^0-9:]', '', s)
    try:
        return dt.datetime.strptime(s2, "%H:%M").time()
    except Exception:
        return None

def normalize_interval(start, end):
    t1 = str_to_time(start)
    t2 = str_to_time(end)
    if not t1 or not t2:
        return None
    return f"{t1.strftime('%H:%M')} - {t2.strftime('%H:%M')}"

def time_to_minutes(t):
    return t.hour * 60 + t.minute

def intervals_overlap(a_start, a_end, b_start, b_end):
    a_s = time_to_minutes(str_to_time(a_start))
    a_e = time_to_minutes(str_to_time(a_end))
    b_s = time_to_minutes(str_to_time(b_start))
    b_e = time_to_minutes(str_to_time(b_end))
    return max(a_s, b_s) < min(a_e, b_e)

def gerar_intervalos(inicio: dt.time, fim: dt.time, passo: dt.timedelta):
    """Gera lista de dt.time de inicio até fim (exclusivo fim) com passo."""
    if inicio is None or fim is None:
        return []
    cur = dt.datetime.combine(dt.date.today(), inicio)
    end_dt = dt.datetime.combine(dt.date.today(), fim)
    res = []
    while cur < end_dt:
        res.append(cur.time())
        cur += passo
    return res

def exportar_dados(df: pd.DataFrame):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    caminho = OUTPUT_DIR / "dados_disciplinas.xlsx"
    df.to_excel(caminho, index=False)
    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return buffer, caminho

# -----------------------
# Leitura e processamento
# -----------------------
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

def criar_lista_salas(df_salas: pd.DataFrame):
    salas = []
    for _, row in df_salas.iterrows():
        nome = str(row.get("SALAS") or row.get("SALA") or row.get("NOME") or "").strip()
        capacidade = int(row.get("CAPACIDADE") or 0)
        salas.append({
            "NOME": nome,
            "CAPACIDADE": capacidade,
            "DATAS": set(),
            # HORARIOS_OCUPADOS armazena strings 'HH:MM - HH:MM'
            "HORARIOS_OCUPADOS": set(),
            # HORARIOS_OCUPADOS_SEMANA: dict dia -> list de (inicio_str, fim_str, descricao)
            "HORARIOS_OCUPADOS_SEMANA": {d: [] for d in DIAS_SEMANA},
            # RESERVAS manuais (data, inicio_str, fim_str, descricao)
            "RESERVAS": []
        })
    return salas

def re_split_days(s: str):
    parts = re.split(r'[;,/\\]+|\s{2,}|\s', s)
    return [p for p in parts if p]

def gerar_datas(df_turmas):
    # tenta extrair colunas 13/14 como antes, com fallback
    try:
        data_inicio = list(map(int, str(df_turmas.iloc[0, 13]).split(",")))
        data_final = list(map(int, str(df_turmas.iloc[0, 14]).split(",")))
        return pd.date_range(dt.date(*data_inicio), dt.date(*data_final))
    except Exception:
        # fallback simples: usa min/max da primeira coluna interpretada como data
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
                # descrição da alocação (segura contra escapes)
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
            "HORARIO_INICIO": inicio_t,
            "HORARIO_FINAL": fim_t,
            "HORARIOS_RAW": aloc.get("HORARIO") or aloc.get("HORÁRIO") or "",
            "ALUNOS": aloc.get("ALUNOS") or 0,
            "PROFESSOR": aloc.get("PROFESSOR"),
            "CAPACIDADE": next((s["CAPACIDADE"] for s in salas_ct if s["NOME"] == sala), None),
            "DATAS": datas,
            "DESCRICAO": descricao
        })

        sala_obj = next((s for s in salas_ct if s["NOME"] == sala), None)
        if sala_obj:
            for d in dias_validos:
                # se há início/fim válidos, adiciona intervalos por dia
                if inicio_t and fim_t:
                    sala_obj["HORARIOS_OCUPADOS_SEMANA"][d].append((
                        inicio_t.strftime("%H:%M"), fim_t.strftime("%H:%M"), descricao
                    ))
                    sala_obj["HORARIOS_OCUPADOS"].add(f"{inicio_t.strftime('%H:%M')} - {fim_t.strftime('%H:%M')}")
                else:
                    # tenta extrair da coluna HORARIOS_RAW (ex: 'SEGUNDA 07:00-08:00, TERÇA 09:00-10:00')
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

# -----------------------
# Cria workbook por sala
# -----------------------
def criar_workbook_horario_sala(sala_obj):
    horas_minutos = []
    for h in range(7, 22):
        horas_minutos.append(f"{h:02d}:00 - {h:02d}:30")
        horas_minutos.append(f"{h:02d}:30 - {h+1:02d}:00")

    wb = Workbook()
    ws = wb.active
    ws.title = sala_obj["NOME"][:31]

    dias = DIAS_SEMANA
    info_sala = f"Centro de Tecnologia | {sala_obj['NOME']} | Capacidade: {sala_obj['CAPACIDADE']}"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(dias)+1)
    cell_info = ws.cell(row=1, column=1, value=info_sala)
    cell_info.font = Font(bold=True, size=12)
    cell_info.alignment = Alignment(horizontal='center', vertical='center')

    ws.cell(row=2, column=1, value="Horário").font = Font(bold=True)
    for col, dia in enumerate(dias, start=2):
        ws.cell(row=2, column=col, value=dia).font = Font(bold=True)

    for row, hora in enumerate(horas_minutos, start=3):
        ws.cell(row=row, column=1, value=hora)

    # ---------- preenche disciplinas + reservas ----------
    for col, dia in enumerate(dias, start=2):
        ocupados = sala_obj["HORARIOS_OCUPADOS_SEMANA"].get(dia, [])
        for inicio, fim, desc in ocupados:
            t_start = str_to_time(inicio)
            t_end   = str_to_time(fim)
            if not t_start or not t_end:
                continue
            cur = dt.datetime.combine(dt.date.today(), t_start)
            fim_dt = dt.datetime.combine(dt.date.today(), t_end)
            while cur < fim_dt:
                nxt = cur + dt.timedelta(minutes=30)
                label = f"{cur.time().strftime('%H:%M')} - {nxt.time().strftime('%H:%M')}"
                try:
                    row_idx = horas_minutos.index(label) + 3
                except ValueError:
                    cur = nxt
                    continue

                # ---------- monta texto da célula ----------
                # desc já está no formato "EVENTO" ou "CÓDIGO - DISCIPLINA - TURMA - PROFESSOR"
                # Se for uma reserva manual, procuramos a data correspondente
                if desc.startswith("RESERVA_MANUAL") or desc == desc.strip():
                    # procuramos a reserva que tem esse horário para pegar as datas
                    for r in sala_obj["RESERVAS"]:
                        r_data, r_ini, r_fim, r_desc = r
                        if r_ini == inicio and r_fim == fim and r_desc == desc:
                            if r_data != r_data:          # comparar só o dia
                                continue
                            data_ini_fmt = r_data.strftime("%d/%m")
                            # verifica se existe data diferente (intervalo)
                            datas_reserva = [d for d, i, f, dscr in sala_obj["RESERVAS"]
                                             if i == inicio and f == fim and dscr == desc]
                            if len({t[0] for t in datas_reserva}) > 1:
                                data_fim_fmt = max(d[0] for d in datas_reserva).strftime("%d/%m")
                                texto_celula = f"{desc} – {data_ini_fmt} – {data_fim_fmt}"
                            else:
                                texto_celula = desc
                            break
                    else:
                        texto_celula = desc
                else:
                    texto_celula = desc

                ws.cell(row=row_idx, column=col, value=texto_celula)
                cur = nxt

    # ---------- mescla células iguais ----------
    for col in range(2, len(dias) + 2):
        start_row = 3
        cur_val = ws.cell(row=3, column=col).value
        for row in range(3, len(horas_minutos) + 3):
            val = ws.cell(row=row, column=col).value
            if val != cur_val:
                if cur_val not in (None, "") and row - 1 >= start_row:
                    ws.merge_cells(start_row=start_row, start_column=col, end_row=row - 1, end_column=col)
                start_row = row
                cur_val = val
        if cur_val not in (None, "") and start_row <= len(horas_minutos) + 2:
            ws.merge_cells(start_row=start_row, start_column=col, end_row=len(horas_minutos) + 2, end_column=col)

    # ---------- estilo ----------
    thin = Side(style="thin")
    borda = Border(left=thin, right=thin, top=thin, bottom=thin)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fonte = Font(size=10)
    for row in ws.iter_rows(min_row=1, max_row=len(horas_minutos)+2, min_col=1, max_col=len(dias)+1):
        for cell in row:
            cell.border = borda
            cell.alignment = align
            cell.font = fonte
    for col in range(1, len(dias)+2):
        ws.column_dimensions[get_column_letter(col)].width = 25

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer
    return wb


# -----------------------
# Interface Streamlit
# -----------------------
def interface_interativa(salas_ct, df_processado):
    st.header("🎯 Solicitação de Sala")

    # ---------- Dados da reserva ----------
    evento        = st.text_input("Digite o nome do evento:")
    blocos        = sorted({s["NOME"][:3] for s in salas_ct if s["NOME"]})
    bloco_sel     = st.selectbox("Selecione o bloco:", blocos)
    salas_filt    = [s["NOME"] for s in salas_ct if s["NOME"].startswith(bloco_sel)]
    sala_escolhida= st.selectbox("Selecione a sala:", salas_filt)

    col1, col2 = st.columns(2)
    with col1:
        data_ini = st.date_input("Data inicial:", key="dt_ini")
    with col2:
        usa_fim  = st.selectbox("Data final (opcional):", ["NÃO","SIM"], key="sn_fim")
    if usa_fim == "SIM":
        data_fim = st.date_input("Data final:", key="dt_fim")
        # >>> escolher dias da semana <<<
        dias_evento = st.multiselect("Dias da semana que o evento ocorrerá:",
                                     DIAS_SEMANA,
                                     default=["SEGUNDA"])
    else:
        data_fim  = None
        dias_evento = None

    h_ini = st.time_input("Horário de início:",  key="h_ini")
    h_fim = st.time_input("Horário de término:", key="h_fim")

    # ---------- Sala escolhida ----------
    sala_info = next((s for s in salas_ct if s["NOME"] == sala_escolhida), None)
    if sala_info is None:
        st.error("Sala não encontrada.")
        return

    # ---------- Horários ocupados (apenas visualização) ----------
    st.subheader("Horários ocupados (por dia)")
    for dia in DIAS_SEMANA:
        ocu = sala_info["HORARIOS_OCUPADOS_SEMANA"].get(dia, [])
        st.write(f"**{dia}**: " + (", ".join([f"{a}-{b} ({c})" for a, b, c in ocu]) if ocu else "Nenhum"))

    # ---------- Botão ÚNICO de solicitação ----------
    if st.button("📅 Solicitar Sala", key="btn_solicitar"):
        inicio_str = h_ini.strftime("%H:%M")
        fim_str    = h_fim.strftime("%H:%M")
        mapping    = {'MONDAY':'SEGUNDA','TUESDAY':'TERÇA','WEDNESDAY':'QUARTA',
                      'THURSDAY':'QUINTA','FRIDAY':'SEXTA','SATURDAY':'SÁBADO','SUNDAY':'DOMINGO'}

        # ---------- gera lista de datas que serão verificadas ----------
        if usa_fim == "SIM" and data_fim and dias_evento:
            # intervalo + filtro de dias da semana
            todas_as_datas_evento = pd.date_range(data_ini, data_fim, freq='D') \
                                      .to_series() \
                                      .map(lambda d: mapping.get(d.strftime("%A").upper(),
                                                                 d.strftime("%A").upper())) \
                                      .isin(dias_evento)
            datas_a_verificar = pd.date_range(data_ini, data_fim, freq='D')[todas_as_datas_evento]
        else:
            # reserva única
            datas_a_verificar = [data_ini]

        # ---------- verifica conflito em CADA data/horário ----------
        conflitos = []
        for data in datas_a_verificar:
            dia_port = mapping.get(data.strftime("%A").upper(), data.strftime("%A").upper())
            for a, b, desc in sala_info["HORARIOS_OCUPADOS_SEMANA"].get(dia_port, []):
                if intervals_overlap(a, b, inicio_str, fim_str):
                    conflitos.append((data.strftime("%d/%m"), a, b, desc))

        if conflitos:
            st.error("❌ Conflitos encontrados:\n" +
                     "\n".join([f"{dt} {a}-{b} ({d})" for dt, a, b, d in conflitos]))
        else:
            desc = evento.strip() if evento and str(evento).strip() else "RESERVA_MANUAL"

            # ---------- grava todas as datas livres ----------
            for data in datas_a_verificar:
                dia_port = mapping.get(data.strftime("%A").upper(), data.strftime("%A").upper())
                sala_info["RESERVAS"].append((data, inicio_str, fim_str, desc))
                sala_info["HORARIOS_OCUPADOS_SEMANA"].setdefault(dia_port, []).append(
                    (inicio_str, fim_str, desc))
                sala_info["HORARIOS_OCUPADOS"].add(f"{inicio_str} - {fim_str}")

            st.success(f"✅ Evento registrado em {len(datas_a_verificar)} dia(s).")
                        
    
    # ---------- Download Excel da sala ----------
    st.divider()
    if st.download_button("📥 Baixar Excel (Sala)",
                      data=criar_workbook_horario_sala(sala_info),
                      file_name=f"horario_{sala_escolhida}.xlsx",
                      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
        pass   # só para manter a lógica do if

    # ---------- Download geral ----------
    st.divider()
    st.subheader("Exportar dados processados (todas as turmas)")
    buf_df = BytesIO()
    df_processado.to_excel(buf_df, index=False)
    buf_df.seek(0)
    st.download_button("📥 Baixar dados_disciplinas.xlsx", data=buf_df,
                       file_name="dados_disciplinas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# -----------------------
# Main
# -----------------------
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
