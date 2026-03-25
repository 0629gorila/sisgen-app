import streamlit as st
from docx import Document
from docx.shared import Inches
from openpyxl import load_workbook
from openpyxl.styles import Font
import tempfile
import re

st.set_page_config(page_title="SISGÉN AI", layout="centered")

# -------- LOGIN --------
usuarios = {
    "admin": "1234",
    "sisgen": "2026"
}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.markdown("<h1 style='color:white;text-align:center;'>Acceso SISGÉN AI</h1>", unsafe_allow_html=True)

    user = st.text_input("Usuario")
    password = st.text_input("Contraseña", type="password")

    if st.button("Ingresar"):
        if user in usuarios and usuarios[user] == password:
            st.session_state.autenticado = True
        else:
            st.error("Credenciales incorrectas")

    st.stop()

# -------- DISEÑO FUTURISTA --------
st.markdown("""
<style>
body {
    background: black;
}

.stApp {
    background: radial-gradient(circle at center, #0f2027, #203a43, #000000);
}

[data-testid="stAppViewContainer"] {
    background-image: linear-gradient(120deg, rgba(0,255,255,0.08), rgba(0,0,0,1));
}

.block-container {
    background: rgba(0,0,0,0.8);
    border-radius: 15px;
    padding: 2rem;
    box-shadow: 0 0 25px rgba(0,255,255,0.3);
}

h1, h2, h3, label {
    color: #00ffff !important;
    text-shadow: 0 0 5px #00ffff;
}

button {
    background-color: #00ffff !important;
    color: black !important;
    border-radius: 10px !important;
}

</style>
""", unsafe_allow_html=True)

# -------- LOGO --------
st.image("logo_sisgen.png", width=150)

st.markdown("<h1 style='text-align:center;'>Motor Inteligente SISGÉN</h1>", unsafe_allow_html=True)

# -------- INPUTS --------
empresa = st.text_input("Empresa")
representante = st.text_input("Representante")
direccion = st.text_input("Dirección")
correo = st.text_input("Correo electrónico")
fecha = st.text_input("Fecha")

logo = st.file_uploader("Logo del cliente (OBLIGATORIO)", type=["png", "jpg", "jpeg"])
archivo = st.file_uploader("Documento Word o Excel", type=["docx", "xlsx"])

# -------- GENERAR --------
if st.button("Generar documento"):

    if not archivo:
        st.error("Debe subir un documento")
        st.stop()

    if not logo:
        st.error("Debe subir el logo")
        st.stop()

    if not empresa or not representante or not direccion or not correo:
        st.error("Complete todos los campos")
        st.stop()

    # -------- WORD INTELIGENTE --------
    if archivo.name.endswith(".docx"):

        doc = Document(archivo)

        def es_correo(texto):
            return re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", texto)

        def es_fecha(texto):
            return re.search(r"\d{1,2}/\d{1,2}/\d{2,4}", texto)

        def es_nombre(texto):
            return texto.isupper() and len(texto.split()) >= 2

        def procesar(texto):
            if not texto:
                return texto

            if es_correo(texto):
                return correo

            elif es_fecha(texto):
                return fecha

            elif es_nombre(texto):
                return representante

            elif any(p in texto.lower() for p in ["conjunto", "residencial", "ph"]):
                return empresa

            return texto

        # PÁRRAFOS
        for p in doc.paragraphs:
            p.text = procesar(p.text)

        # TABLAS
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        p.text = procesar(p.text)

        # LOGO EN ENCABEZADO
        try:
            temp_logo = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            temp_logo.write(logo.read())

            for section in doc.sections:
                header = section.header
                header.paragraphs[0].clear()

                run = header.paragraphs[0].add_run()
                run.add_picture(temp_logo.name, width=Inches(1.5))

                header.paragraphs[0].add_run(f"   {empresa}")

        except:
            st.error("Error en el logo")
            st.stop()

        # PIE
        for section in doc.sections:
            footer = section.footer
            footer.paragraphs[0].clear()
            footer.paragraphs[0].add_run(f"Fecha de generación: {fecha}")

        temp_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(temp_docx.name)

        st.success("Documento generado correctamente")

        with open(temp_docx.name, "rb") as f:
            st.download_button("Descargar Word", f)

    # -------- EXCEL --------
    elif archivo.name.endswith(".xlsx"):

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        temp_file.write(archivo.read())

        wb = load_workbook(temp_file.name)

        for hoja in wb.worksheets:
            for fila in hoja.iter_rows():
                for celda in fila:
                    if celda.value:
                        texto = str(celda.value)

                        if "@" in texto:
                            texto = correo
                        elif "/" in texto:
                            texto = fecha

                        celda.value = texto
                        celda.font = Font(bold=True)

        temp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        wb.save(temp_excel.name)

        st.success("Excel generado correctamente")

        with open(temp_excel.name, "rb") as f:
            st.download_button("Descargar Excel", f)
