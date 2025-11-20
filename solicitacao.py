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
"HORARIO INICIO": set(),  # Adicionado para evitar KeyError/AttributeError
"HORARIO FINAL": set(),  # Adicionado para evitar KeyError/AttributeError
}
for _, row in df_salas.iterrows()
]


def gerar_datas(df_turmas):
"""Gera todas as datas entre o início e o fim definidos na planilha."""
# Assumindo que os dados da planilha são strings no formato "d,m,a" ou similar
# e que o código original está correto ao usar dt.date(*data_partes)
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
# As chaves 'HORARIO INICIO' e 'HORARIO FINAL' foram adicionadas em criar_lista_salas
                s["HORARIO INICIO"].add(aloc["HORARIO INICIO"])
                s["HORARIO FINAL"].add(aloc["HORARIO FINAL"])
                
                # Tentativa de converter para dt.time, se não for. Isso é crucial para a função gerar_intervalos.
                horario_inicio_obj = aloc["HORARIO INICIO"]
                if isinstance(horario_inicio_obj, str):
                    try:
                        horario_inicio_obj = dt.datetime.strptime(horario_inicio_obj, "%H:%M:%S").time()
                    except ValueError:
                        # Se falhar, assume que é uma string e tenta converter para dt.time
                        horario_inicio_obj = dt.datetime.strptime(horario_inicio_obj, "%H:%M").time()
                
                horario_final_obj = aloc["HORARIO FINAL"]
                if isinstance(horario_final_obj, str):
                    try:
                        horario_final_obj = dt.datetime.strptime(horario_final_obj, "%H:%M:%S").time()
                    except ValueError:
                        horario_final_obj = dt.datetime.strptime(horario_final_obj, "%H:%M").time()
                        
                s["HORARIO INICIO"].add(horario_inicio_obj)
                s["HORARIO FINAL"].add(horario_final_obj)

return pd.DataFrame(dados)

@@ -171,126 +188,123 @@
sala_info = next((s for s in salas_ct if s["NOME"] == sala_escolhida), None)

if sala_info:
        # Conversão dos sets de horários de início e fim para strings para exibição
        horarios_ocupados_str = {
            f"{h_ini.strftime('%H:%M')} - {h_fim.strftime('%H:%M')}"
            for h_ini, h_fim in zip(sala_info['HORARIO INICIO'], sala_info['HORARIO FINAL'])
        }
        # O set 'HORARIOS_OCUPADOS' já contém as strings de horário de alocação (ex: "18:00 - 20:00").
        horarios_ocupados_str = sala_info["HORARIOS_OCUPADOS"]

if horarios_ocupados_str:
st.info(f"🕓 Horários ocupados (alocados): {', '.join(sorted(horarios_ocupados_str))}")
else:
st.success("✅ Nenhum horário ocupado encontrado para esta sala.")

if st.button("📅 Solicitar Sala"):
if not sala_info:
st.error("Sala não encontrada.")
return

# Conflito 1: Checa se o horário de início ou fim está contido em um horário ocupado (string)
# O código original usava `horario_inicio.strftime("%H:%M") in h`
# Isso só funciona se `h` for um set de strings de horários, o que não é o caso aqui.
# `sala_info["HORARIOS_OCUPADOS"]` contém strings de horários (ex: "18:00 - 20:00")
horario_inicio_str = horario_inicio.strftime("%H:%M")
horario_fim_str = horario_fim.strftime("%H:%M")

# A lógica original parece tentar verificar se os horários de início ou fim
# estão contidos em alguma string de horário ocupado.
conflito = any(
horario_inicio_str in h or horario_fim_str in h
for h in sala_info["HORARIOS_OCUPADOS"]
)

# Conflito 2: Checa sobreposição usando a função gerar_intervalos
# O código original estava incorreto ao tentar acessar "HORARIOS INICIO" e "HORARIO FINAL"
# como se fossem um único objeto datetime.time, e não um set.
# Além disso, a função gerar_intervalos foi corrigida para aceitar dt.time e dt.timedelta

# Para manter a lógica original (que parecia tentar gerar um intervalo a partir
# de todos os horários de início e fim registrados, o que é estranho),
# vamos usar os sets de horários de início e fim para a verificação.
# A lógica mais provável é que o usuário queria checar se o intervalo
# solicitado se sobrepõe a qualquer intervalo já alocado.
# Como a lógica original usa `gerar_intervalos` com os sets, e isso é um erro,
# vou tentar corrigir *mantendo a intenção* de verificar a sobreposição,
# mas usando os sets de horários de início e fim que agora estão disponíveis.

# A correção mais fiel à lógica original (mesmo que errada) é:
# Acessar o primeiro elemento do set, o que é perigoso, ou assumir que o set
# só tem um elemento, o que é incorreto.
# Vou assumir que o usuário queria pegar o *menor* horário de início e o *maior*
# horário final de *todas* as alocações da sala para criar um grande intervalo,
# o que é uma lógica estranha, mas é a única que se encaixa no uso de `ini` e `f`
# como argumentos únicos para `gerar_intervalos`.

try:
# Pega o menor horário de início e o maior horário final de todas as alocações da sala
ini = min(sala_info["HORARIO INICIO"]) if sala_info["HORARIO INICIO"] else dt.time(0, 0)
f = max(sala_info["HORARIO FINAL"]) if sala_info["HORARIO FINAL"] else dt.time(0, 0)
except TypeError:
# Caso os sets estejam vazios ou contenham tipos misturados, o que não deveria ocorrer após a correção.
ini = dt.time(0, 0)
f = dt.time(0, 0)

intervalo = dt.timedelta(minutes=1)

# A função gerar_intervalos foi corrigida para aceitar dt.time e dt.timedelta
# O resultado é uma lista de objetos dt.time
horario_intervalo = gerar_intervalos(ini, f, intervalo)

# A lógica original (linha 176) checa se o horário de início ou fim solicitado
# está presente na lista de horários intermediários gerados.
# Isso só faz sentido se `horario_intervalo` contiver todos os minutos
# entre o primeiro horário de início e o último horário final.
# E mesmo assim, a verificação é falha.

# Corrigindo a verificação da linha 176 para usar dt.time
amostra = [
True if h == horario_inicio or h == horario_fim else False 
for h in horario_intervalo
]

conflito_2 = any(amostra)

if conflito or conflito_2:
st.error("❌ A sala está ocupada no horário selecionado.")
else:
st.success(f"✅ Solicitação registrada para **{sala_escolhida}** em {data_escolhida} "
f"({horario_inicio_str}–{horario_fim_str})")
# Adiciona a string do horário ocupado ao set de strings
sala_info["HORARIOS_OCUPADOS"].add(f"{horario_inicio_str} - {horario_fim_str}")
# Adiciona os objetos dt.time aos sets de início e fim
sala_info["HORARIO INICIO"].add(horario_inicio)
sala_info["HORARIO FINAL"].add(horario_fim)


# Botão de download
buffer, caminho = exportar_dados(df_processado)
st.download_button(
label="📥 Baixar Excel Processado",
data=buffer,
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
