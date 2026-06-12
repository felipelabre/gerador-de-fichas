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


if csv_file:

    try:

        df = pd.read_csv(csv_file, sep=";", encoding="utf-8-sig")

        st.success(f"{len(df)} inscrições carregadas.")

        if st.button("Gerar Fichas"):

            doc = Document()

            logo_temp = None

            if logo_file:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                    tmp.write(logo_file.read())
                    logo_temp = tmp.name

            col_nome = "Nome"
            col_cidade = "Cidade"
            col_telefone = "Celular"
            col_nascimento = "Data de Nascimento"

            col_camiseta = "Qual o tamanho da sua camiseta?"

            col_ministerios = (
                "A coordenação do acampamento é a responsável por montar as equipes "
                "de trabalho. Mas, gostaríamos de receber a sua opinião. "
                "Assinale até 3 ministérios que você gostaria de servir. "
            )

            col_serviu = "Já serviu em acampamentos? Se sim, quais ministérios?"

            col_pastoral = (
                "Participa de alguma Pastoral ou Movimento? Se sim, qual:"
            )

            col_sacramentos = (
                "Quais Sacramentos possui (batismo, eucaristia, crisma, matrimônio e ordem)?"
            )

            for i, row in df.iterrows():

                if i > 0:
                    doc.add_page_break()

                if logo_temp:
                    doc.add_picture(logo_temp, width=Inches(2))

                doc.add_heading(titulo_acampamento, level=1)

                nome = row.get(col_nome, "")
                cidade = row.get(col_cidade, "")
                telefone = row.get(col_telefone, "")
                nascimento = row.get(col_nascimento, "")
                camiseta = row.get(col_camiseta, "")
                ministerios = row.get(col_ministerios, "")
                serviu = row.get(col_serviu, "")
                pastoral = row.get(col_pastoral, "")
                sacramentos = row.get(col_sacramentos, "")

                idade = calcular_idade(nascimento)

                doc.add_paragraph(f"Nome: {nome}")
                doc.add_paragraph(f"Cidade: {cidade}")
                doc.add_paragraph(f"Telefone: {telefone}")
                doc.add_paragraph(
                    f"Data de nascimento / Idade: {nascimento} - ({idade} anos)"
                )
                doc.add_paragraph(f"Tamanho da camiseta: {camiseta}")
                doc.add_paragraph(f"Ministérios desejados: {ministerios}")
                doc.add_paragraph(f"Já serviu em acampamentos: {serviu}")
                doc.add_paragraph(
                    f"Participa de alguma Pastoral ou Movimento: {pastoral}"
                )
                doc.add_paragraph(f"Sacramentos: {sacramentos}")

                doc.add_paragraph("")
                doc.add_paragraph("FOTO")
                doc.add_paragraph("")
                doc.add_paragraph("")
                doc.add_paragraph("")
                doc.add_paragraph("")
                doc.add_paragraph("")
                doc.add_paragraph(
                    "________________________________________"
                )

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
