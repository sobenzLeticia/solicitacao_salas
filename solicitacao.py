import datetime as dt
from pathlib import Path
from io import BytesIO
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter

# ===============================
# CONFIGURAÇÕES GERAIS
# ===============================
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR

CAMINHO_SALAS = DATA_DIR / "SALAS - COPIA.xlsx"
CAMINHO_DISCIPLINAS = DATA_DIR / "Resultados_Gerais.xlsx"
OUTPUT_DIR = BASE_DIR / "resultados"
OUTPUT_DIR.mkdir(exist_ok=True)

DIAS_SEMANA = ["SEGUNDA", "TERÇA", "QUARTA", "QUINTA", "SEXTA", "SÁBADO"]
INDICE_DIAS = {d: i for i, d in enumerate(DIAS_SEMANA)}

# ===============================
# UTILITÁRIOS DE HORÁRIO
# ===============================

def str_to_time(s: str) -> dt.time:
    """Tenta converter strings para dt.time aceitando vários formatos."""
    if pd.isna(s):
        return None
    s = str(s).strip()
    for fmt in ("%H:%M", "%H:%M:%S", "%H.%M"):
        try:
            return dt.datetime.strptime(s, fmt).time()
        except Exception:
            pass
    # tenta extrair apenas os dois primeiros números
    try:
        parts = ''.join(ch if ch.isdigit() or ch == ':' else '' for ch in s)
        return dt.datetime.strptime(parts, "%H:%M").time()
    except Exception:
        return None


def normalize_interval(start, end) -> str:
    """Recebe dt.time ou strings e retorna 'HH:MM - HH:MM'"""
    t1 = start if isinstance(start, dt.time) else str_to_time(start)
    t2 = end if isinstance(end, dt.time) else str_to_time(end)
    if not t1 or not t2:
        return None
    return f"{t1.strftime('%H:%M')} - {t2.strftime('%H:%M')}"


def time_to_minutes(t: dt.time) -> int:
    return t.hour * 60 + t.minute


def intervals_overlap(a_start: str, a_end: str, b_start: str, b_end: str) -> bool:
    a_s = time_to_minutes(str_to_time(a_start))
    a_e = time_to_minutes(str_to_time(a_end))
    b_s = time_to_minutes(str_to_time(b_start))
    b_e = time_to_minutes(str_to_time(b_end))
    return max(a_s, b_s) < min(a_e, b_e)

# ===============================
# LEITURA E PROCESSAMENTO
# ===============================

@st.cache_data(show_spinner=False)
def carregar_dados():
    """Carrega dados de salas e turmas. Lança erro amigável se não achar arquivos."""
    if not CAMINHO_SALAS.exists():
        st.error(f"Arquivo de salas não encontrado: {CAMINHO_SALAS}")
        st.stop()
    if not CAMINHO_DISCIPLINAS.exists():
        st.error(f"Arquivo de disciplinas não encontrado: {CAMINHO_DISCIPLINAS}")
        st.stop()

    df_salas = pd.read_excel(CAMINHO_SALAS)
    df_turmas = pd.read_excel(CAMINHO_DISCIPLINAS)
    return df_salas, df_turmas


def criar_lista_salas(df_salas: pd.DataFrame):
    """Cria estrutura de salas a partir do DataFrame de salas."""
    salas = []
    for _, row in df_salas.iterrows():
        nome = str(row.get('SALAS') or row.get('SALA') or row.get('NOME') or '')
        capacidade = int(row.get('CAPACIDADE') or 0)
        salas.append({
            'NOME': nome.strip(),
            'CAPACIDADE': capacidade,
            # Lista de reservas como tuplas (data, inicio, fim, descricao)
            'RESERVAS': [],
            # Horários ocupados por semana (dia -> list de (inicio,fim,descricao))
            'HORARIOS_OCUPADOS_SEMANA': {d: [] for d in DIAS_SEMANA},
        })
    return salas


def gerar_datas(df_turmas: pd.DataFrame):
    """Tenta extrair data início/fim. Se houver colunas nomeadas usa-as, senão tenta heurística."""
    # procura colunas comuns
    cols = [c.upper() for c in df_turmas.columns]
    if 'DATA INICIO' in cols or 'DATA_INICIO' in cols or 'DATAINICIO' in cols:
        # pega pelo nome original
        for c in df_turmas.columns:
            if str(c).upper().replace('_', ' ') == 'DATA INICIO':
                data_inicio = pd.to_datetime(df_turmas[c].iloc[0]).date()
                break
    else:
        # fallback para o comportamento original (colunas 13/14 contendo 'YYYY,MM,DD')
        try:
            start_list = list(map(int, str(df_turmas.iloc[0, 13]).split(',')))
            data_inicio = dt.date(*start_list)
        except Exception:
            data_inicio = pd.to_datetime(df_turmas.iloc[:, 0].min()).date()

    if 'DATA FINAL' in cols or 'DATA_FINAL' in cols or 'DATAFINAL' in cols:
        for c in df_turmas.columns:
            if str(c).upper().replace('_', ' ') == 'DATA FINAL':
                data_final = pd.to_datetime(df_turmas[c].iloc[0]).date()
                break
    else:
        try:
            end_list = list(map(int, str(df_turmas.iloc[0, 14]).split(',')))
            data_final = dt.date(*end_list)
        except Exception:
            data_final = pd.to_datetime(df_turmas.iloc[:, 0].max()).date()

    if data_final < data_inicio:
        data_final = data_inicio

    return pd.date_range(start=data_inicio, end=data_final)


def processar_alocacoes(df_turmas: pd.DataFrame, todas_as_datas: pd.DatetimeIndex, salas_ct: list):
    """Gera lista de alocações normalizadas e preenche HORARIOS_OCUPADOS_SEMANA nas salas."""
    registros = []

    for _, aloc in df_turmas.iterrows():
        status = str(aloc.get('STATUS') or '').strip()
        if status.upper() != 'ALOCADA':
            continue

        sala = str(aloc.get('SALA') or aloc.get('SALAS') or '').strip()
        if not sala:
            continue

        dias_raw = str(aloc.get('DIAS') or '').strip()
        if not dias_raw:
            continue

        # Normaliza dias: aceita espaço, vírgula ou ';' como separadores
        dias_tokens = [t.strip().upper() for t in re_split_days(dias_raw)]
        dias_validos = [d for d in dias_tokens if d in INDICE_DIAS]
        if not dias_validos:
            continue

        inicio = aloc.get('HORARIO INICIO') or aloc.get('HORARIO') or aloc.get('HORÁRIO INICIO')
        fim = aloc.get('HORARIO FINAL') or aloc.get('HORÁRIO FINAL') or aloc.get('HORARIO_FIM')
        intervalo_normal = normalize_interval(inicio, fim)
        if not intervalo_normal:
            # tenta extrair de uma coluna HORARIO no formato 'SEGUNDA 07:00-08:00, ...'
            # nesse caso processaremos abaixo como lista de blocos
            pass

        codigo = aloc.get('CODIGO') or ''
        disciplina = aloc.get('DISCIPLINA') or ''
        turma = aloc.get('TURMA') or ''
        professor = aloc.get('PROFESSOR') or ''
        alunos = aloc.get('ALUNOS') or 0

        # Desc opcional
        descricao = f"{codigo} - {disciplina} - {turma} - {professor}"

        # datas do semestre/periodo para esses dias
        indices = [INDICE_DIAS[d] for d in dias_validos]
        datas = todas_as_datas[todas_as_datas.dayofweek.isin(indices)]

        registro = {
            'CURSO': aloc.get('CURSO'),
            'CODIGO': codigo,
            'SALA': sala,
            'DISCIPLINA': disciplina,
            'TURMA': turma,
            'DIAS': ','.join(dias_validos),
            'HORARIO_INICIO': str_to_time(inicio),
            'HORARIO_FINAL': str_to_time(fim),
            'HORARIOS_RAW': aloc.get('HORARIO') or aloc.get('HORÁRIO') or '',
            'ALUNOS': alunos,
            'PROFESSOR': professor,
            'DATAS': datas,
            'DESCRICAO': descricao,
        }
        registros.append(registro)

        # Atualiza a sala correspondente
        sala_obj = next((s for s in salas_ct if s['NOME'] == sala), None)
        if sala_obj:
            # Preenche reservas semanais por dia
            for d in dias_validos:
                if registro['HORARIO_INICIO'] and registro['HORARIO_FINAL']:
                    sala_obj['HORARIOS_OCUPADOS_SEMANA'][d].append((registro['HORARIO_INICIO'].strftime('%H:%M'),
                                                                     registro['HORARIO_FINAL'].strftime('%H:%M'),
                                                                     descricao))
                else:
                    # Se não existirem colunas separadas de início/fim, tenta extrair da coluna HORARIOS_RAW
                    raw = registro['HORARIOS_RAW']
                    # espera blocos como 'SEGUNDA 07:00-08:00, TERÇA 09:00-10:00'
                    blocos = [b.strip() for b in str(raw).split(',') if b.strip()]
                    for bloco in blocos:
                        try:
                            parts = bloco.split()
                            dia = parts[0].upper()
                            hora_part = parts[1]
                            h_start, h_end = hora_part.split('-')
                            if dia in DIAS_SEMANA:
                                sala_obj['HORARIOS_OCUPADOS_SEMANA'][dia].append((h_start, h_end, descricao))
                        except Exception:
                            continue

    df_reg = pd.DataFrame(registros)
    return df_reg

# função auxiliar usada acima
import re

def re_split_days(s: str):
    # separa por vírgula, ponto e vírgula, barra, ou espaços múltiplos
    parts = re.split(r'[;,/\\]+|\s{2,}|\s', s)
    return [p for p in parts if p]

# ===============================
# GERAÇÃO DO EXCEL (por sala)
# ===============================

def criar_workbook_horario_sala(sala_obj, horas_minutos=None):
    """Cria um Workbook do Excel contendo apenas o quadro semanal da sala."""
    if horas_minutos is None:
        horas_minutos = []
        for h in range(7, 22):
            horas_minutos.append(f"{h:02d}:00 - {h:02d}:30")
            horas_minutos.append(f"{h:02d}:30 - {h+1:02d}:00")

    wb = Workbook()
    ws = wb.active
    ws.title = sala_obj['NOME'][:31]

    dias = DIAS_SEMANA

    # Cabeçalho
    info_sala = f"Centro de Tecnologia | {sala_obj['NOME']} | Capacidade: {sala_obj['CAPACIDADE']}"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(dias) + 1)
    cell_info = ws.cell(row=1, column=1, value=info_sala)
    cell_info.font = Font(bold=True, size=12)
    cell_info.alignment = Alignment(horizontal='center', vertical='center')

    ws.cell(row=2, column=1, value='Horário').font = Font(bold=True)
    for col, dia in enumerate(dias, start=2):
        ws.cell(row=2, column=col, value=dia).font = Font(bold=True)

    for row, hora in enumerate(horas_minutos, start=3):
        ws.cell(row=row, column=1, value=hora)

    # Preenche horários ocupados
    for col, dia in enumerate(dias, start=2):
        ocupados = sala_obj['HORARIOS_OCUPADOS_SEMANA'].get(dia, [])
        for inicio, fim, desc in ocupados:
            # preenche todos os intervalos 30min que estiverem dentro do bloco
            t_start = str_to_time(inicio)
            t_end = str_to_time(fim)
            if not t_start or not t_end:
                continue
            cur = dt.datetime.combine(dt.date.today(), t_start)
            fim_dt = dt.datetime.combine(dt.date.today(), t_end)
            while cur < fim_dt:
                next_dt = cur + dt.timedelta(minutes=30)
                label = f"{cur.time().strftime('%H:%M')} - {next_dt.time().strftime('%H:%M')}"
                try:
                    row_idx = horas_minutos.index(label) + 3
                except ValueError:
                    cur = next_dt
                    continue
                ws.cell(row=row_idx, column=col, value=desc)
                cur = next_dt

    # Mesclar células verticalmente onde o mesmo valor se repete
    for col in range(2, len(dias) + 2):
        start_row = 3
        current_value = ws.cell(row=3, column=col).value
        for row in range(3, len(horas_minutos) + 3):
            value = ws.cell(row=row, column=col).value
            if value != current_value:
                if current_value not in (None, '') and row - 1 >= start_row:
                    ws.merge_cells(start_row=start_row, start_column=col, end_row=row - 1, end_column=col)
                start_row = row
                current_value = value
        if current_value not in (None, '') and start_row <= len(horas_minutos) + 2:
            ws.merge_cells(start_row=start_row, start_column=col, end_row=len(horas_minutos) + 2, end_column=col)

    # Estilos básicos
    thin = Side(style='thin')
    borda_fina = Border(left=thin, right=thin, top=thin, bottom=thin)
    alinhamento_centro = Alignment(horizontal='center', vertical='center', wrap_text=True)
    fonte_padrao = Font(size=10)

    for row in ws.iter_rows(min_row=1, max_row=len(horas_minutos) + 2, min_col=1, max_col=len(dias) + 1):
        for cell in row:
            cell.border = borda_fina
            cell.alignment = alinhamento_centro
            cell.font = fonte_padrao

    for col in range(1, len(dias) + 2):
        ws.column_dimensions[get_column_letter(col)].width = 25

    return wb

# ===============================
# INTERFACE STREAMLIT
# ===============================

def interface_interativa(salas_ct, df_processado):
    st.header("🎯 Solicitação de Sala")

    # Detecta blocos automaticamente a partir dos prefixos das salas (3 primeiros chars)
    blocos_detectados = sorted({s['NOME'][:3] for s in salas_ct if s['NOME']})
    if not blocos_detectados:
        st.error('Nenhuma sala encontrada nas planilhas.')
        return

    bloco_selecionado = st.selectbox("Selecione o bloco:", blocos_detectados)
    salas_filtradas = [s['NOME'] for s in salas_ct if s['NOME'].startswith(bloco_selecionado)]
    sala_escolhida = st.selectbox("Selecione a sala:", salas_filtradas)

    data_escolhida = st.date_input("Selecione a data:")
    horario_inicio = st.time_input("Horário de início:")
    horario_fim = st.time_input("Horário de término:")

    sala_info = next((s for s in salas_ct if s['NOME'] == sala_escolhida), None)
    if not sala_info:
        st.error('Sala não encontrada.')
        return

    # Mostrar horários semanais ocupados
    st.subheader('Horários ocupados (por dia da semana)')
    for dia in DIAS_SEMANA:
        ocup = sala_info['HORARIOS_OCUPADOS_SEMANA'].get(dia, [])
        if ocup:
            st.write(f"**{dia}**: ", ', '.join([f"{a}-{b} ({c})" for a, b, c in ocup]))
        else:
            st.write(f"**{dia}**: Nenhum")

    # Verifica conflito
    if st.button('📅 Solicitar Sala'):
        inicio_str = horario_inicio.strftime('%H:%M')
        fim_str = horario_fim.strftime('%H:%M')
        dia_semana = data_escolhida.strftime('%A').upper()
        # mapeia inglês para português se necessário
        mapping = {
            'MONDAY': 'SEGUNDA', 'TUESDAY': 'TERÇA', 'WEDNESDAY': 'QUARTA',
            'THURSDAY': 'QUINTA', 'FRIDAY': 'SEXTA', 'SATURDAY': 'SÁBADO', 'SUNDAY': 'DOMINGO'
        }
        dia_semana = mapping.get(dia_semana, dia_semana)

        conflitos = []
        for a, b, desc in sala_info['HORARIOS_OCUPADOS_SEMANA'].get(dia_semana, []):
            if intervals_overlap(a, b, inicio_str, fim_str):
                conflitos.append((a, b, desc))

        if conflitos:
            st.error('❌ A sala está ocupada no horário selecionado: ' + ', '.join([f"{a}-{b} ({c})" for a, b, c in conflitos]))
        else:
            # registra reserva (em memória apenas)
            sala_info['RESERVAS'].append((data_escolhida, inicio_str, fim_str, 'RESERVA MANUAL'))
            st.success(f"✅ Solicitação registrada para {sala_escolhida} em {data_escolhida} ({inicio_str} - {fim_str})")

    st.divider()

    # Gerar e permitir download do Excel apenas para a sala selecionada
    if st.button('📥 Gerar Excel da Sala Selecionada'):
        wb = criar_workbook_horario_sala(sala_info)
        buffer = BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        st.download_button("Baixar Excel (Sala)", data=buffer, file_name=f"horario_{sala_escolhida}.xlsx", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # Também permitir download do DataFrame processado completo
    st.divider()
    st.subheader('Exportar dados processados (todas as turmas)')
    buffer_df = BytesIO()
    df_processado.to_excel(buffer_df, index=False)
    buffer_df.seek(0)
    st.download_button('📥 Baixar dados_disciplinas.xlsx', data=buffer_df, file_name='dados_disciplinas.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ===============================
# APP PRINCIPAL
# ===============================

def main():
    st.title('🏫 Sistema de Alocação de Salas – CT')
    with st.spinner('Carregando dados...'):
        df_salas, df_turmas = carregar_dados()
        salas_ct = criar_lista_salas(df_salas)
        todas_as_datas = gerar_datas(df_turmas)
        df_dados = processar_alocacoes(df_turmas, todas_as_datas, salas_ct)

    st.success('✅ Dados carregados e processados com sucesso!')
    st.divider()
    interface_interativa(salas_ct, df_dados)


if __name__ == '__main__':
    main()
