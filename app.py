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
    value="Ficha de Inscrição do Servo - Resumida"
)


def calcular_idade(data_nascimento):
    try:
        nascimento = pd.to_datetime(
            data_nascimento,
            dayfirst=True,
            errors="coerce"
        )

        if pd.isna(nascimento):
            return ""

        hoje = datetime.today()

        idade = hoje.year - nascimento.year

        if (hoje.month, hoje.day) < (nascimento.month, nascimento.day):
            idade -= 1

        return idade

    except:
        return ""


if csv_file:

    try:

        try:
            df = pd.read_csv(
                csv_file,
                sep=";",
                encoding="utf-8-sig"
            )
        except:
            csv_file.seek(0)

            df = pd.read_csv(
                csv_file,
                sep=";",
                encoding="latin1"
            )

        st.success(f"{len(df)} inscrições carregadas.")

        if st.button("Gerar Fichas"):

            doc = Document()

            logo_temp = None

            if logo_file:
                with tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=".png"
                ) as tmp:
                    tmp.write(logo_file.read())
                    logo_temp = tmp.name

            col_nome = "Nome"
            col_cidade = "Cidade"
            col_telefone = "Celular"
            col_nascimento = "Data de Nascimento"

            col_camiseta = "Qual o tamanho da sua camiseta?"

            col_ministerios = (
                "A coordenação do acampamento é a responsável por montar as equipes de trabalho. "
                "Mas, gostaríamos de receber a sua opinião. Assinale até 3 ministérios que você gostaria de servir. "
            )

            col_serviu = (
                "Já serviu em acampamentos? Se sim, quais ministérios?"
            )

            col_pastoral = (
                "Participa de alguma Pastoral ou Movimento? Se sim, qual:"
            )

            col_sacramentos = (
                "Quais Sacramentos possui (batismo, eucaristia, crisma, matrimônio e ordem)?"
            )

            for i, row in df.iterrows():

                if i > 0:
                    doc.add_page_break()

                # Logo centralizada
                if logo_temp:
                    p_logo = doc.add_paragraph()
                    p_logo.alignment = 1

                    run_logo = p_logo.add_run()
                    run_logo.add_picture(
                        logo_temp,
                        width=Inches(2.2)
                    )

                # Título centralizado
                titulo = doc.add_paragraph()
                titulo.alignment = 1

                run_titulo = titulo.add_run(titulo_acampamento)
                run_titulo.bold = True

                # Dados
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

                # Nome grande
                nome_p = doc.add_paragraph()
                nome_p.alignment = 1

                run_nome = nome_p.add_run(str(nome))
                run_nome.bold = True

                doc.add_paragraph("")

                # Tabela
                tabela = doc.add_table(rows=0, cols=2)
                tabela.style = "Table Grid"

                if idade != "":
                    nascimento_texto = f"{nascimento} ({idade} anos)"
                else:
                    nascimento_texto = str(nascimento)

                campos = [
                    ("Cidade", cidade),
                    ("Telefone", telefone),
                    ("Nascimento / Idade", nascimento_texto),
                    ("Tamanho Camiseta", camiseta),
                    ("Ministérios Desejados", ministerios),
                    ("Já Serviu", serviu),
                    ("Pastoral / Movimento", pastoral),
                    ("Sacramentos", sacramentos)
                ]

                for campo, valor in campos:
                    linha = tabela.add_row().cells
                    linha[0].text = str(campo)
                    linha[1].text = str(valor)

                doc.add_paragraph("")

                # Espaço para foto
                foto = doc.add_table(rows=1, cols=1)
                foto.style = "Table Grid"

                celula = foto.cell(0, 0)
                celula.text = "\n\n\n\n\nFOTO\n\n\n\n\n"

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
