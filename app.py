import io
import os
import tempfile
from datetime import datetime
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt  # Importado o Pt para controlar o tamanho

st.set_page_config(
    page_title="Gerador de Fichas", page_icon="📄", layout="centered"
)

st.title("📄 Gerador de Fichas")

logo_file = st.file_uploader(
    "Logo do Acampamento (opcional)", type=["png", "jpg", "jpeg"]
)

csv_file = st.file_uploader("Arquivo CSV das Inscrições", type=["csv"])

titulo_acampamento = st.text_input(
    "Título do Acampamento", value="Ficha de Inscrição do Servo - Resumida"
)


def calcular_idade(data_nascimento):
    try:
        nascimento = pd.to_datetime(
            data_nascimento, dayfirst=True, errors="coerce"
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


if "word_file" not in st.session_state:
    st.session_state.word_file = None

if csv_file:
    try:
        try:
            df = pd.read_csv(csv_file, sep=";", encoding="utf-8-sig")
        except:
            csv_file.seek(0)
            df = pd.read_csv(csv_file, sep=";", encoding="latin1")

        st.success(f"{len(df)} inscrições carregadas.")

        if st.button("Gerar Fichas"):
            with st.spinner("Gerando o documento Word com fontes ampliadas..."):
                doc = Document()
                logo_temp = None

                if logo_file:
                    with tempfile.NamedTemporaryFile(
                        delete=False, suffix=".png"
                    ) as tmp:
                        tmp.write(logo_file.getvalue())
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
                col_sacramentos = "Quais Sacramentos possui (batismo, eucaristia, crisma, matrimônio e ordem)?"

                for i, row in df.iterrows():
                    if i > 0:
                        doc.add_page_break()

                    # Logo centralizada
                    if logo_temp:
                        p_logo = doc.add_paragraph()
                        p_logo.alignment = 1
                        run_logo = p_logo.add_run()
                        run_logo.add_picture(logo_temp, width=Inches(2.2))

                    # Título centralizado (Tamanho 14)
                    titulo = doc.add_paragraph()
                    titulo.alignment = 1
                    run_titulo = titulo.add_run(titulo_acampamento)
                    run_titulo.bold = True
                    run_titulo.font.size = Pt(14)

                    # Dados do Servo
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

                    # Nome bem grande e destacado (Tamanho 24)
                    nome_p = doc.add_paragraph()
                    nome_p.alignment = 1
                    run_nome = nome_p.add_run(str(nome).upper())  # Coloquei em MAIÚSCULO para destacar ainda mais
                    run_nome.bold = True
                    run_nome.font.size = Pt(24)

                    doc.add_paragraph("")

                    # Tabela de Dados Gerais
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
                        ("Sacramentos", sacramentos),
                    ]

                    for campo, valor in campos:
                        linha = tabela.add_row().cells
                        
                        # Coluna 0: Nome do Campo (Tamanho 13 + Negrito)
                        p_campo = linha[0].paragraphs[0]
                        run_campo = p_campo.add_run(str(campo))
                        run_campo.bold = True
                        run_campo.font.size = Pt(13)
                        
                        # Coluna 1: Valor do Campo (Tamanho 12)
                        p_valor = linha[1].paragraphs[0]
                        run_valor = p_valor.add_run(str(valor))
                        run_valor.font.size = Pt(12)

                    doc.add_paragraph("")

                    # Espaço para foto
                    foto = doc.add_table(rows=1, cols=1)
                    foto.style = "Table Grid"
                    celula = foto.cell(0, 0)
                    
                    p_foto = celula.paragraphs[0]
                    p_foto.alignment = 1  # Centraliza o texto "FOTO"
                    run_foto = p_foto.add_run("\n\n\n\n\nFOTO\n\n\n\n\n")
                    run_foto.font.size = Pt(12)

                    doc.add_paragraph("")
                    doc.add_paragraph(
                        "________________________________________"
                    )

                bio = io.BytesIO()
                doc.save(bio)
                bio.seek(0)

                st.session_state.word_file = bio.read()

                if logo_temp and os.path.exists(logo_temp):
                    os.remove(logo_temp)

        if st.session_state.word_file is not None:
            st.download_button(
                label="📥 Baixar Arquivo Word",
                data=st.session_state.word_file,
                file_name="fichas_acampamento.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    except Exception as e:
        st.error(f"Erro ao processar CSV: {e}")
