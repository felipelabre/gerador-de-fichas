import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches
from datetime import datetime
import tempfile
import os

st.set_page_config(
    page_title="Gerador de Fichas",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Gerador de Fichas")

st.markdown("Faça upload da logo (opcional) e do CSV das inscrições.")

logo_file = st.file_uploader(
    "Logo do Acampamento (opcional)",
    type=["png", "jpg", "jpeg"]
)

csv_file = st.file_uploader(
    "Arquivo CSV das Inscrições",
    type=["csv"]
)

titulo_acampamento = st.text_input(
    "Título do Acampamento",
    value="Ficha do Acampante"
)


def calcular_idade(data_nascimento):
    try:
        nascimento = pd.to_datetime(data_nascimento, dayfirst=True)
        hoje = datetime.today()

        idade = (
            hoje.year
            - nascimento.year
            - (
                (hoje.month, hoje.day)
                < (nascimento.month, nascimento.day)
            )
        )

        return idade

    except:
        return ""


def buscar_coluna(df, possibilidades):
    for col in df.columns:
        nome = col.lower().strip()

        for p in possibilidades:
            if p in nome:
                return col

    return None


if csv_file:

    try:
        df = pd.read_csv(csv_file, sep=";")

        st.success(f"{len(df)} inscrições carregadas.")

        if st.button("Gerar Fichas"):

            doc = Document()

            logo_temp = None

            if logo_file:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                    tmp.write(logo_file.read())
                    logo_temp = tmp.name

            col_nome = buscar_coluna(df, ["nome"])
            col_cidade = buscar_coluna(df, ["cidade"])
            col_telefone = buscar_coluna(df, ["telefone", "celular", "whatsapp"])
            col_nascimento = buscar_coluna(df, ["nascimento"])
            col_camiseta = buscar_coluna(df, ["camiseta"])
            col_ministerios = buscar_coluna(df, ["minist"])
            col_dificuldade = buscar_coluna(df, ["dificuldade"])
            col_serviu = buscar_coluna(df, ["serviu"])
            col_pastoral = buscar_coluna(df, ["pastoral", "movimento"])
            col_sacramentos = buscar_coluna(df, ["sacramento"])

            for i, row in df.iterrows():

                if i > 0:
                    doc.add_page_break()

                if logo_temp:
                    doc.add_picture(logo_temp, width=Inches(2))

                doc.add_heading(titulo_acampamento, level=1)

                nome = row[col_nome] if col_nome else ""
                cidade = row[col_cidade] if col_cidade else ""
                telefone = row[col_telefone] if col_telefone else ""
                nascimento = row[col_nascimento] if col_nascimento else ""
                camiseta = row[col_camiseta] if col_camiseta else ""
                ministerios = row[col_ministerios] if col_ministerios else ""
                dificuldade = row[col_dificuldade] if col_dificuldade else ""
                serviu = row[col_serviu] if col_serviu else ""
                pastoral = row[col_pastoral] if col_pastoral else ""
                sacramentos = row[col_sacramentos] if col_sacramentos else ""

                idade = calcular_idade(nascimento)

                doc.add_paragraph(f"Nome: {nome}")
                doc.add_paragraph(f"Cidade: {cidade}")
                doc.add_paragraph(f"Telefone: {telefone}")
                doc.add_paragraph(
                    f"Data de nascimento / Idade: {nascimento} / {idade}"
                )
                doc.add_paragraph(f"Tamanho da camiseta: {camiseta}")
                doc.add_paragraph(f"Ministérios desejados: {ministerios}")
                doc.add_paragraph(
                    f"Dificuldade em qual ministério: {dificuldade}"
                )
                doc.add_paragraph(
                    f"Já serviu em acampamento? Qual(is) ministério(s)?: {serviu}"
                )
                doc.add_paragraph(
                    f"Participa de alguma Pastoral ou Movimento?: {pastoral}"
                )
                doc.add_paragraph(f"Sacramentos: {sacramentos}")

                doc.add_paragraph("")
                doc.add_paragraph("FOTO")
                doc.add_paragraph("")
                doc.add_paragraph("")
                doc.add_paragraph("")
                doc.add_paragraph("")
                doc.add_paragraph("")
                doc.add_paragraph("____________________________________")

            output_path = "fichas_acampamento.docx"
            doc.save(output_path)

            with open(output_path, "rb") as file:
                st.download_button(
                    "📥 Baixar Arquivo Word",
                    file,
                    file_name="fichas_acampamento.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            if logo_temp and os.path.exists(logo_temp):
                os.remove(logo_temp)

    except Exception as e:
        st.error(f"Erro ao processar CSV: {e}")
