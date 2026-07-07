import streamlit as size_config  # Apenas para o padrão
import streamlit as st
import interface_aluno as int_aluno
import solicitacao_salas_2 as sol_salas

# 1. Configuração inicial da página e do estado da sessão
st.set_page_config(page_title="Sistema de Login", page_icon="🔒", layout="centered")

if "logado" not in st.session_state:
    st.session_state.logado = False
if "perfil" not in st.session_state:
    st.session_state.perfil = None

# Senha do administrador (em produção, usar st.secrets)
SENHA_ADMIN_CORRETA = "admin123"

# Função para deslogar
def deslogar():
    st.session_state.logado = False
    st.session_state.perfil = None
    st.rerun()

# --- TELA DE LOGIN ---
if not st.session_state.logado:
    st.title("Bem-vindo! ")
    st.subheader("Por favor, selecione como deseja acessar o sistema:")

    # Seleção do perfil
    perfil = st.radio("Escolha seu perfil:", ["Estudante", "Administrador"])

    # Se for Administrador, mostra o campo de senha
    if perfil == "Administrador":
        senha = st.text_input("Digite a senha de Administrador:", type="password")
        
        if st.button("Acessar como Admin"):
            if senha == SENHA_ADMIN_CORRETA:
                st.session_state.logado = True
                st.session_state.perfil = "Administrador"
                st.success("Acesso autorizado!")
                st.rerun()
            else:
                st.error("Senha incorreta. Tente novamente.")
                
    # Se for Estudante, entra direto sem senha
    elif perfil == "Estudante":
        if st.button("Acessar como Estudante"):
            st.session_state.logado = True
            st.session_state.perfil = "Estudante"
            st.success("Bem-vindo, estudante!")
            st.rerun()

# --- TELAS APÓS O LOGIN ---
else:
    # Botão de Logout fixo no topo/lateral para facilitar
    st.sidebar.button("Sair / Logout", on_click=deslogar)

    if st.session_state.perfil == "Estudante":
        # --- TELA DO ESTUDANTE ---
        st.title("🎓 Portal do Estudante")
        st.write("Bem-vindo ao Sistema de Solicitação de Salas")
        int_aluno.main()
        # Conteúdo fictício do estudante
        #st.info("Você tem 2 tarefas pendentes para esta semana.")
        #st.metric(label="Sua Frequência Geral", value="92%")

    elif st.session_state.perfil == "Administrador":
        # --- TELA DO ADMINISTRADOR ---
        st.title("🛡️ Painel do Administrador")
        st.write("Bem-vindo ao painel de controle. Você tem acesso total ao sistema.")
        
        sol_salas.main()

        # Conteúdo fictício do admin
        st.warning("Área Restrita: Gerenciamento de usuários, testes e relatórios.")
        st.columns(3)[0].metric(label="Total de Alunos", value="1,240")
        st.columns(3)[1].metric(label="Novos Cadastros", value="+45")