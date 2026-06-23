import io
import os
import tempfile
from datetime import datetime
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, Cm

st.set_page_config(
    page_title="Gerador de Fichas", page_icon="📄", layout="centered"
)

st.title("📄 Gerador de Fichas (Foco em Estabilidade)")

logo_file = st.file_uploader(
    "1. Logo do Acampamento (opcional)", type=["png", "jpg", "jpeg"]
)

csv_file = st.file_uploader("2. Arquivo CSV das Inscrições", type=["csv"])

fotos_files = st.file_uploader(
    "3. Selecione as Fotos dos Servos (Opcional)",
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True
)

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
        # Lendo o CSV
        try:
            df = pd.read_csv(csv_file, sep=";", encoding="utf-8-sig")
        except:
            csv_file.seek(0)
            df = pd.read_csv(csv_file, sep=";", encoding="latin1")

        # Remove linhas completamente vazias que o Excel costuma gerar no fim do arquivo
        df = df.dropna(how='all')
        
        st.success(f"{len(df)} inscrições válidas carregadas.")

        # Criando o mapeamento de fotos de forma segura
        dicionario_fotos = {}
        if fotos_files:
            for foto in fotos_files:
                if foto.name:
                    nome_sem_extensao = os.path.splitext(foto.name)[0].strip().lower()
                    dicionario_fotos[nome_sem_extensao] = foto
            st.info(f"{len(dicionario_fotos)} fotos processadas e prontas para cruzamento.")

        if st.button("Gerar Fichas"):
            # Variável de controle para sabermos onde o erro aconteceu se algo der errado
            registro_atual = "Início do Processamento"
            
            try:
                with st.spinner("Gerando fichas... Isso pode levar um tempinho devido ao volume de fotos."):
                    doc = Document()
                    
                    # Margens estreitas (1.5 cm)
                    sections = doc.sections
                    for section in sections:
                        section.top_margin = Cm(1.5)
                        section.bottom_margin = Cm(1.5)
                        section.left_margin = Cm(1.5)
                        section.right_margin = Cm(1.5)

                    logo_temp = None
                    if logo_file:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
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
                    col_serviu = "Já serviu em acampamentos? Se sim, quais ministérios?"
                    col_pastoral = "Participa de alguma Pastoral ou Movimento? Se sim, qual:"
                    col_sacramentos = "Quais Sacramentos possui (batismo, eucaristia, crisma, matrimônio e ordem)?"

                    for i, row in df.iterrows():
                        # Proteção contra linhas sem nome preenchido
                        if pd.isna(row.get(col_nome)) or str(row.get(col_nome)).strip() == "":
                            continue

                        nome = str(row.get(col_nome, "")).strip()
                        registro_atual = f"Servo da linha {i+1}: {nome}"

                        if i > 0:
                            doc.add_page_break()

                        # 1. Logo
                        if logo_temp:
                            p_logo = doc.add_paragraph()
                            p_logo.alignment = 1
                            run_logo = p_logo.add_run()
                            run_logo.add_picture(logo_temp, width=Inches(1.8))

                        # 2. Título
                        titulo = doc.add_paragraph()
                        titulo.alignment = 1
                        run_titulo = titulo.add_run(titulo_acampamento)
                        run_titulo.bold = True
                        run_titulo.font.size = Pt(13)

                        # Dados do contrato
                        cidade = row.get(col_cidade, "") if pd.notna(row.get(col_cidade)) else ""
                        telefone = row.get(col_telefone, "") if pd.notna(row.get(col_telefone)) else ""
                        nascimento = row.get(col_nascimento, "") if pd.notna(row.get(col_nascimento)) else ""
                        camiseta = row.get(col_camiseta, "") if pd.notna(row.get(col_camiseta)) else ""
                        ministerios = row.get(col_ministerios, "") if pd.notna(row.get(col_ministerios)) else ""
                        serviu = row.get(col_serviu, "") if pd.notna(row.get(col_serviu)) else ""
                        pastoral = row.get(col_pastoral, "") if pd.notna(row.get(col_pastoral)) else ""
                        sacramentos = row.get(col_sacramentos, "") if pd.notna(row.get(col_sacramentos)) else ""

                        idade = calcular_idade(nascimento)

                        # 3. Nome
                        nome_p = doc.add_paragraph()
                        nome_p.alignment = 1
                        run_nome = nome_p.add_run(nome.upper())
                        run_nome.bold = True
                        run_nome.font.size = Pt(24)

                        doc.add_paragraph("")

                        # 4. Tabela
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
                            p_campo = linha[0].paragraphs[0]
                            run_campo = p_campo.add_run(str(campo))
                            run_campo.bold = True
                            run_campo.font.size = Pt(11)
                            
                            p_valor = linha[1].paragraphs[0]
                            run_valor = p_valor.add_run(str(valor))
                            run_valor.font.size = Pt(11)

                        p_espaco = doc.add_paragraph()
                        p_espaco.paragraph_format.space_before = Pt(8)

                        # 5. Foto
                        foto = doc.add_table(rows=1, cols=1)
                        foto.style = "Table Grid"
                        celula = foto.cell(0, 0)
                        p_foto = celula.paragraphs[0]
                        p_foto.alignment = 1

                        nome_chave = nome.lower()
                        
                        if nome_chave in dicionario_fotos:
                            foto_file_obj = dicionario_fotos[nome_chave]
                            ext = os.path.splitext(foto_file_obj.name)[1]
                            
                            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as f_tmp:
                                f_tmp.write(foto_file_obj.getvalue())
                                foto_temp_path = f_tmp.name
                            
                            run_foto = p_foto.add_run()
                            run_foto.add_picture(foto_temp_path, width=Inches(1.6))
                            
                            if os.path.exists(foto_temp_path):
                                os.remove(foto_temp_path)
                        else:
                            run_foto = p_foto.add_run("\n\n\n\nFOTO NÃO ENCONTRADA\n\n\n\n")
                            run_foto.font.size = Pt(11)

                        p_linha = doc.add_paragraph()
                        p_linha.alignment = 1
                        p_linha.paragraph_format.space_before = Pt(8)
                        p_linha.add_run("________________________________________")

                    # Salvando em memória
                    bio = io.BytesIO()
                    doc.save(bio)
                    bio.seek(0)

                    st.session_state.word_file = bio.read()

                    if logo_temp and os.path.exists(logo_temp):
                        os.remove(logo_temp)
                        
                st.success("✨ Todas as fichas foram geradas com sucesso!")

            except Exception as inner_error:
                st.error(f"Erro detectado em: **{registro_atual}**")
                st.error(f"Detalhes do erro técnico: {inner_error}")

        if st.session_state.word_file is not None:
            st.download_button(
                label="📥 Baixar Arquivo Word",
                data=st.session_state.word_file,
                file_name="fichas_acampamento.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    except Exception as e:
        st.error(f"Erro geral no processamento dos arquivos: {e}")
 
