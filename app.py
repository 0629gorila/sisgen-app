import streamlit as st
from docx import Document
from docx.shared import Inches
from openpyxl import load_workbook
from openpyxl.styles import Font
import tempfile

st.set_page_config(page_title="SISGÉN PRO", layout="centered")

# -------- LOGIN --------
usuarios = {
    "admin": "1234",
    "sisgen": "2026"
}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("Acceso Plataforma SISGÉN")

    user = st.text_input("Usuario")
    password = st.text_input("Contraseña", type="password")

    if st.button("Ingresar"):
        if user in usuarios and usuarios[user] == password:
            st.session_state.autenticado = True
        else:
            st.error("Credenciales incorrectas")

    st.stop()

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

h1, h2, h3, label {
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

    # -------- WORD --------
    if archivo.name.endswith(".docx"):

        doc = Document(archivo)

        def procesar(texto):
            if texto:
                texto = str(texto)
                texto = texto.replace("TORRE AZUL", empresa)
                texto = texto.replace("CONJUNTO", empresa)
                texto = texto.replace("RESIDENCIAL", empresa)
                texto = texto.replace("OFELIA CORZO PINILLA", representante)
                texto = texto.replace("correo@ejemplo.com", correo)

                if "fecha" in texto.lower():
                    texto = fecha

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
            st.error("El logo no es válido")
            st.stop()

        # PIE DE PÁGINA
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
                        texto = texto.replace("TORRE AZUL", empresa)
                        texto = texto.replace("OFELIA CORZO PINILLA", representante)
                        texto = texto.replace("correo@ejemplo.com", correo)

                        if "fecha" in texto.lower():
                            texto = fecha

                        celda.value = texto
                        celda.font = Font(bold=True)

        temp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        wb.save(temp_excel.name)

        st.success("Excel generado correctamente")

        with open(temp_excel.name, "rb") as f:
            st.download_button("Descargar Excel", f)
