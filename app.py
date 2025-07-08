import streamlit as st
import pandas as pd
import datetime
import os
from fpdf import FPDF
from PIL import Image
import matplotlib.pyplot as plt
import smtplib
from email.message import EmailMessage

# ========== CONFIGURAÇÕES ==========
st.set_page_config(page_title="Inspeções Termográficas PPTM")
STORAGE_PATH = "registros"
IMAGEM_PATH = "imagens"
LOGO_PATH = "logo.png"
OS_FILE = "os's.xlsx"
HISTORICO_FILE = "historico_inspecoes.xlsx"
EMAIL_REMETENTE = "lukinhamala6@gmail.com" 
EMAIL_SENHA = "sjdup ifgr lscq pnxv"

# ========== CARREGAMENTO ==========
@st.cache_data
def carregar_ordens():
    df = pd.read_excel(OS_FILE)
    df = df[['Nº OS Protheus', 'Descrição', 'Bem']].dropna()
    return df

df_os = carregar_ordens()

def buscar_dados_os(numero_os):
    resultado = df_os[df_os['Nº OS Protheus'] == numero_os]
    if not resultado.empty:
        return resultado.iloc[0]['Descrição'], resultado.iloc[0]['Bem']
    return "Ordem não encontrada.", ""

def salvar_historico(data, os, bem, r, s, t, temp):
    if os.path.exists(HISTORICO_FILE):
        df = pd.read_excel(HISTORICO_FILE)
    else:
        df = pd.DataFrame(columns=["Data", "Numero_OS", "Codigo_BEM", "Corrente R", "Corrente S", "Corrente T", "Temperatura"])
    novo = pd.DataFrame([{
        "Data": data,
        "Numero_OS": os,
        "Codigo_BEM": bem,
        "Corrente R": r,
        "Corrente S": s,
        "Corrente T": t,
        "Temperatura": temp
    }])
    df = pd.concat([df, novo], ignore_index=True)
    df.to_excel(HISTORICO_FILE, index=False)

def gerar_grafico_historico(bem):
    if not os.path.exists(HISTORICO_FILE):
        return None
    df = pd.read_excel(HISTORICO_FILE)
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df[df["Codigo_BEM"] == bem]
    if df.empty:
        return None

    fig, ax = plt.subplots()
    ax.plot(df["Data"], df["Corrente R"], label="Fase R", marker="o")
    ax.plot(df["Data"], df["Corrente S"], label="Fase S", marker="o")
    ax.plot(df["Data"], df["Corrente T"], label="Fase T", marker="o")
    ax.plot(df["Data"], df["Temperatura"], label="Temperatura", marker="o")

    for i, row in df.iterrows():
        ax.annotate(f'{row["Corrente R"]:.1f}', (row["Data"], row["Corrente R"]), xytext=(0,10), textcoords='offset points', fontsize=7, ha="center")
        ax.annotate(f'{row["Corrente S"]:.1f}', (row["Data"], row["Corrente S"]), xytext=(0,-12), textcoords='offset points', fontsize=7, ha="center")
        ax.annotate(f'{row["Corrente T"]:.1f}', (row["Data"], row["Corrente T"]), xytext=(0,15), textcoords='offset points', fontsize=7, ha="center")
        ax.annotate(f'{row["Temperatura"]:.1f}', (row["Data"], row["Temperatura"]), xytext=(0,-20), textcoords='offset points', fontsize=7, ha="center")

    ax.set_title(f"Histórico BEM - {bem}")
    ax.set_xlabel("Data")
    ax.set_ylabel("Valores")
    ax.legend()
    ax.grid(True)
    plt.tight_layout()

    caminho = f"{IMAGEM_PATH}/grafico_bem_{bem}.png"
    plt.savefig(caminho)
    plt.close()
    return caminho

def enviar_email_com_anexo(destinatario, assunto, corpo, anexo_path):
    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = destinatario
    msg.set_content(corpo)

    with open(anexo_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename=os.path.basename(anexo_path))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_REMETENTE, EMAIL_SENHA)
        smtp.send_message(msg)

def gerar_pdf_tecnico(numero_os, descricao_os, data_inspecao, executante, observacoes,
                      corrente_r, corrente_s, corrente_t,
                      limpeza_status, limpeza_obs,
                      reaperto_status, reaperto_obs,
                      temperatura, encontrou_anomalia, detalhe_anomalia,
                      imagens_paths, caminho_grafico):
    pdf = FPDF()
    pdf.add_page()

    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=10, y=8, w=30)
    pdf.set_xy(50, 10)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, f"Termografia OS = {numero_os}", ln=True)
    pdf.ln(20)

    def titulo(txt): pdf.set_font("Arial", "B", 12); pdf.cell(0, 10, txt, ln=True); pdf.ln(2)
    def texto(txt): pdf.set_font("Arial", "", 11); pdf.multi_cell(0, 8, txt); pdf.ln(2)

    titulo("Dados da Inspeção")
    texto(f"Número da OS: {numero_os}")
    texto(f"Descrição: {descricao_os}")
    texto(f"Data: {data_inspecao}")
    texto(f"Executante: {executante}")

    titulo("Checklist Técnico")
    texto(f"Corrente Fase R: {corrente_r} A")
    texto(f"Corrente Fase S: {corrente_s} A")
    texto(f"Corrente Fase T: {corrente_t} A")

    texto(f"Limpeza: {'OK' if limpeza_status == 'OK' else 'Não OK'}")
    if limpeza_status == "Não OK": texto(f"Motivo: {limpeza_obs}")
    texto(f"Reaperto: {'OK' if reaperto_status == 'OK' else 'Não OK'}")
    if reaperto_status == "Não OK": texto(f"Motivo: {reaperto_obs}")
    texto(f"Temperatura: {temperatura} °C")

    titulo("Anomalia Encontrada")
    texto("Sim" if encontrou_anomalia == "Sim" else "Não")
    if encontrou_anomalia == "Sim": texto(f"Detalhes: {detalhe_anomalia}")

    titulo("Observações")
    texto(observacoes if observacoes else "-")

    if imagens_paths:
        titulo("Imagens da Inspeção")
        for img_path in imagens_paths:
            if os.path.exists(img_path):
                pdf.image(img_path, w=100)
                pdf.ln(5)

    if caminho_grafico and os.path.exists(caminho_grafico):
        titulo("Histórico por BEM")
        pdf.image(caminho_grafico, w=180)
        pdf.ln(5)

    os.makedirs(STORAGE_PATH, exist_ok=True)
    nome_pdf = f"{STORAGE_PATH}/Ordem_{int(numero_os)}_{data_inspecao}.pdf"
    pdf.output(nome_pdf)
    return nome_pdf

# ========== INTERFACE ==========
st.title("📋 Termografias PPTM")

st.subheader("1️⃣ Identificação da Ordem")
numero_os = st.number_input("Número da Ordem de Serviço", step=1)
descricao_os, codigo_bem = buscar_dados_os(numero_os)
st.text_area("Descrição da Ordem", value=descricao_os, height=80)

st.subheader("2️⃣ Dados da Inspeção")
data_inspecao = st.date_input("Data da Inspeção", value=datetime.date.today())
executante = st.text_input("Nome do Executante")
emails_destinatarios = st.text_input("E-mails para cópia posterior (não enviados)", placeholder="ex: jose@energiapecem.com, maria@energiapecem.com")

st.subheader("3️⃣ Dados Técnicos")
col1, col2, col3 = st.columns(3)
with col1: corrente_r = st.number_input("Corrente Fase R (A)", step=0.1)
with col2: corrente_s = st.number_input("Corrente Fase S (A)", step=0.1)
with col3: corrente_t = st.number_input("Corrente Fase T (A)", step=0.1)

col4, col5 = st.columns(2)
with col4:
    limpeza_status = st.selectbox("Limpeza", ["OK", "Não OK"])
    limpeza_obs = ""
    if limpeza_status == "Não OK":
        limpeza_obs = st.text_area("Motivo da Limpeza Não OK")
with col5:
    reaperto_status = st.selectbox("Reaperto", ["OK", "Não OK"])
    reaperto_obs = ""
    if reaperto_status == "Não OK":
        reaperto_obs = st.text_area("Motivo do Reaperto Não OK")

temperatura = st.number_input("Temperatura (°C)", step=0.1)

st.subheader("4️⃣ Anomalias e Observações")
encontrou_anomalia = st.radio("Anomalia Encontrada?", ["Não", "Sim"])
anomalia_detalhes = ""
if encontrou_anomalia == "Sim":
    anomalia_detalhes = st.text_area("Descreva a anomalia")

observacoes = st.text_area("Observações Gerais")
st.subheader("5️⃣ Imagens")
imagens = st.file_uploader("Imagens da Inspeção", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

if st.button("Salvar e Gerar Relatório"):
    os.makedirs(IMAGEM_PATH, exist_ok=True)
    caminhos_imgs = []
    for i, img in enumerate(imagens):
        caminho = f"{IMAGEM_PATH}/OS_{int(numero_os)}_img_{i+1}.png"
        image = Image.open(img)
        image.save(caminho)
        caminhos_imgs.append(caminho)

    salvar_historico(data_inspecao, numero_os, codigo_bem, corrente_r, corrente_s, corrente_t, temperatura)
    caminho_grafico = gerar_grafico_historico(codigo_bem)

    nome_pdf = gerar_pdf_tecnico(
        numero_os, descricao_os, data_inspecao, executante, observacoes,
        corrente_r, corrente_s, corrente_t,
        limpeza_status, limpeza_obs,
        reaperto_status, reaperto_obs,
        temperatura,
        encontrou_anomalia, anomalia_detalhes,
        caminhos_imgs,
        caminho_grafico
    )

    corpo_email = f"""
Segue em anexo o relatório da inspeção termográfica da OS {numero_os}.

📅 Data: {data_inspecao}
👷 Executante: {executante}

📧 E-mails informados para cópia posterior:
{emails_destinatarios if emails_destinatarios else '-'}
"""
    enviar_email_com_anexo(
        destinatario="lucas.lima@energiapecem.com",
        assunto=f"Relatório Termografia - OS {numero_os}",
        corpo=corpo_email,
        anexo_path=nome_pdf
    )

    st.success("✅ Relatório gerado e enviado com sucesso!")
    st.rerun()
