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
HISTORICO_PATH = "historico.xlsx"
GRAFICO_PATH = "graficos"

# ========== CARREGAMENTO PLANILHA ==========
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
    return "Ordem não encontrada.", "Desconhecido"

# ========== EMAIL ==========
def enviar_email_com_anexo(destinatario, assunto, corpo, anexo_path):
    remetente = "lukinhamala6@gmail.com"
    senha = "jdup ifgr lscq pnxv"
    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = remetente
    msg["To"] = destinatario
    msg.set_content(corpo)

    with open(anexo_path, "rb") as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=os.path.basename(anexo_path))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(remetente, senha)
        smtp.send_message(msg)

# ========== GERA GRÁFICO ==========
def gerar_grafico_bem(codigo_bem):
    if not os.path.exists(HISTORICO_PATH):
        return None
    df = pd.read_excel(HISTORICO_PATH)
    df = df[df["Bem"] == codigo_bem]
    if df.empty:
        return None
    df["Data"] = pd.to_datetime(df["Data"])
    df = df.sort_values("Data")
    plt.figure(figsize=(9, 4))

    def plotar_linha(x, y, label, cor, offset_y=5, offset_x_days=0):
        plt.plot(x, y, label=label, marker="o", color=cor)
        for i in range(len(x)):
            x_pos = x.iloc[i] + pd.Timedelta(days=offset_x_days)
            y_pos = y.iloc[i] + offset_y
            plt.text(x_pos, y_pos, f"{y.iloc[i]:.1f}", fontsize=8, ha="center", va="bottom", color=cor,
                     bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.9))

    plotar_linha(df["Data"], df["Corrente R"], "Corrente R", "blue", offset_y=4, offset_x_days=-0.3)
    plotar_linha(df["Data"], df["Corrente S"], "Corrente S", "orange", offset_y=8, offset_x_days=0)
    plotar_linha(df["Data"], df["Corrente T"], "Corrente T", "green", offset_y=12, offset_x_days=0.3)
    plotar_linha(df["Data"], df["Temperatura"], "Temperatura (°C)", "red", offset_y=16, offset_x_days=0.6)

    plt.title(f"Histórico do bem: {codigo_bem}")
    plt.xlabel("Data")
    plt.ylabel("Valor")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.legend()

    os.makedirs(GRAFICO_PATH, exist_ok=True)
    caminho = f"{GRAFICO_PATH}/grafico_bem_{codigo_bem}.png"
    plt.savefig(caminho)
    plt.close()
    return caminho

# ========== PDF ==========
def gerar_pdf_tecnico(numero_os, descricao_os, data_inspecao, executante, observacoes,
                      corrente_r, corrente_s, corrente_t,
                      limpeza_status, limpeza_obs,
                      reaperto_status, reaperto_obs,
                      temperatura,
                      encontrou_anomalia, detalhe_anomalia,
                      imagens_paths, grafico_path):
    pdf = FPDF()
    pdf.add_page()

    if os.path.exists(LOGO_PATH):
        pdf.image(LOGO_PATH, x=10, y=8, w=30)
    pdf.set_xy(50, 10)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, f"Relatório Termografia - OS {numero_os}", ln=True, align="L")
    pdf.ln(20)

    def titulo(txt):
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, txt, ln=True)
        pdf.ln(2)

    def texto(txt):
        pdf.set_font("Arial", "", 11)
        pdf.multi_cell(0, 8, txt)
        pdf.ln(1)

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
    if limpeza_status == "Não OK":
        texto(f"Motivo: {limpeza_obs}")
    texto(f"Reaperto: {'OK' if reaperto_status == 'OK' else 'Não OK'}")
    if reaperto_status == "Não OK":
        texto(f"Motivo: {reaperto_obs}")
    texto(f"Temperatura: {temperatura} °C")

    titulo("Anomalia Encontrada")
    if encontrou_anomalia == "Sim":
        texto(f"Sim - {detalhe_anomalia}")
    else:
        texto("Não")

    titulo("Observações")
    texto(observacoes or "-")

    if imagens_paths:
        titulo("Imagens da Inspeção")
        for img_path in imagens_paths:
            if os.path.exists(img_path):
                pdf.image(img_path, w=100)
                pdf.ln(5)

    if grafico_path and os.path.exists(grafico_path):
        titulo("Histórico do Bem")
        pdf.image(grafico_path, w=180)
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

st.subheader("3️⃣ Dados Técnicos")
col1, col2, col3 = st.columns(3)
with col1:
    corrente_r = st.number_input("Corrente Fase R (A)", step=0.1)
with col2:
    corrente_s = st.number_input("Corrente Fase S (A)", step=0.1)
with col3:
    corrente_t = st.number_input("Corrente Fase T (A)", step=0.1)

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

if st.button("Gerar e Enviar Relatório"):
    os.makedirs(IMAGEM_PATH, exist_ok=True)
    caminhos_imgs = []

    for i, img in enumerate(imagens):
        caminho = f"{IMAGEM_PATH}/OS_{int(numero_os)}_img_{i+1}.png"
        image = Image.open(img)
        image.save(caminho)
        caminhos_imgs.append(caminho)

    grafico_path = gerar_grafico_bem(codigo_bem)

    nome_pdf = gerar_pdf_tecnico(
        numero_os, descricao_os, data_inspecao, executante, observacoes,
        corrente_r, corrente_s, corrente_t,
        limpeza_status, limpeza_obs,
        reaperto_status, reaperto_obs,
        temperatura,
        encontrou_anomalia, anomalia_detalhes,
        caminhos_imgs, grafico_path
    )

    enviar_email_com_anexo(
        destinatario="lucas.lima@energiapecem.com",
        assunto=f"Novo Relatório Termografia - OS {numero_os}",
        corpo="Segue em anexo o relatório gerado pelo sistema de inspeções.",
        anexo_path=nome_pdf
    )

    st.success("✅ Relatório gerado e enviado com sucesso!")
    st.rerun()
