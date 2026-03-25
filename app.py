import streamlit as st
from docx import Document
from docx.shared import Inches
from openpyxl import load_workbook
from openpyxl.styles import Font
import tempfile

st.set_page_config(page_title="SISGÉN PRO", layout="centered")

# -------- DISEÑO NASA --------
st.markdown("""
<style>
html, body, [class*="css"] {
    background: linear-gradient(135deg, #0b1f3a, #000814) !important;
}

[data-testid="stAppViewContainer"] {
    background-image: url("https://images.pexels.com/photos/373543/pexels-photo-373543.jpeg") !important;
    background-size: cover !important;
    background-position: center !important;
}

[data-testid="stAppViewContainer"]::before {
    content: "";
    position: fixed;
    width: 100%;
    height: 100%;
    background: rgba(0,0,0,0.6);
    z-index: -1;
}

.block-container {
    background-color: rgba(0,0,0,0.6);
    padding: 2rem;
    border-radius: 12px;
}

h1 {
    color: #4da6ff;
    text-align: center;
}

label {
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

# -------- LOGO --------
st.image("logo_sisgen.png", width=180)

st.title("Motor Documental SISGÉN")

# -------- INPUTS --------
empresa = st.text_input("Empresa")
representante = st.text_input("Representante")
direccion = st.text_input("Dirección")
fecha = st.text_input("Fecha")

logo = st.file_uploader("Sube el logo del cliente", type=["png", "jpg", "jpeg"])
archivo = st.file_uploader("Sube archivo Word o Excel", type=["docx", "xlsx"])

# -------- WORD --------
def reemplazar_word(parrafo):
    texto = "".join(run.text for run in parrafo.runs)

    cambios = {
        "EDIFICIO ONCE 94 - PROPIEDAD HORIZONTAL": empresa,
        "OFELIA CORZO PINILLA": representante,
        "CALLE 10 # 20-30": direccion,
        "12/11/2025": fecha
    }

    for clave, valor in cambios.items():
        if clave in texto:
            texto = texto.replace(clave, valor)

            for run in parrafo.runs:
                run.text = ""

            run = parrafo.add_run(texto)
            run.bold = True
            return

# -------- EXCEL --------
def reemplazar_excel(ruta):
    wb = load_workbook(ruta)

    cambios = {
        "EDIFICIO ONCE 94 - PROPIEDAD HORIZONTAL": empresa,
        "OFELIA CORZO PINILLA": representante,
        "CALLE 10 # 20-30": direccion,
        "12/11/2025": fecha
    }

    for hoja in wb.worksheets:
        for fila in hoja.iter_rows():
            for celda in fila:
                if celda.value:
                    texto = str(celda.value)

                    for clave, valor in cambios.items():
                        if clave in texto:
                            celda.value = texto.replace(clave, valor)
                            celda.font = Font(bold=True)

    temp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    wb.save(temp.name)
    return temp.name

# -------- BOTÓN --------
if archivo and st.button("Generar documento"):

    if archivo.name.endswith(".docx"):

        doc = Document(archivo)

        for p in doc.paragraphs:
            reemplazar_word(p)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        reemplazar_word(p)

        if logo:
            temp_logo = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            temp_logo.write(logo.read())

            for section in doc.sections:
                header = section.header
                if header.paragraphs:
                    header.paragraphs[0].clear()
                    run = header.paragraphs[0].add_run()
                    run.add_picture(temp_logo.name, width=Inches(1.5))

        temp_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(temp_docx.name)

        st.success("Documento Word generado")

        with open(temp_docx.name, "rb") as f:
            st.download_button("📄 Descargar Word", f)

    elif archivo.name.endswith(".xlsx"):

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_file.write(archivo.read())

        resultado_excel = reemplazar_excel(temp_file.name)

        st.success("Excel generado")

        with open(resultado_excel, "rb") as f:
            st.download_button("📊 Descargar Excel", f)
