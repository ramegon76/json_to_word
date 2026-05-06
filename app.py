"""
SIM → Word  |  Convertidor de Metadatos del INEGI
App Streamlit — sube un JSON del SIM y descarga el Word formateado.
"""

import json
import io
import streamlit as st
from converter import build_document

# ─── Configuración de página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="SIM → Word | INEGI",
    page_icon="📄",
    layout="centered",
)

# ─── Estilos ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Barra superior con color institucional */
    [data-testid="stHeader"] { background-color: #003057; }

    /* Títulos */
    h1 { color: #003057 !important; }
    h3 { color: #003057 !important; }

    /* Botón de descarga */
    [data-testid="stDownloadButton"] > button {
        background-color: #003057;
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
        width: 100%;
    }
    [data-testid="stDownloadButton"] > button:hover {
        background-color: #00467a;
    }

    /* Zona de carga */
    [data-testid="stFileUploader"] {
        border: 2px dashed #003057;
        border-radius: 8px;
        padding: 1rem;
    }

    /* Pie de página */
    .footer {
        text-align: center;
        color: #888;
        font-size: 0.8rem;
        margin-top: 3rem;
        padding-top: 1rem;
        border-top: 1px solid #eee;
    }
</style>
""", unsafe_allow_html=True)

# ─── Encabezado ───────────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 5])
with col1:
    st.image(
        "https://www.inegi.org.mx/contenidos/imagen/inegi_logo.png",
        width=80,
    )
with col2:
    st.title("Convertidor de Metadatos SIM → Word")

st.markdown(
    "Sube el archivo **JSON** exportado del Sistema Integrador de Metadatos (SIM) "
    "y obtén el documento Word con formato institucional listo para revisión."
)
st.divider()

# ─── Zona de carga ────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "📂 Selecciona o arrastra aquí el archivo JSON del SIM",
    type=["json"],
    help="Archivo exportado directamente desde el SIM. Formato .json",
)

if uploaded_file is not None:
    # Leer y validar JSON
    try:
        raw = uploaded_file.read().decode("utf-8-sig")
        data = json.loads(raw)
    except Exception as e:
        st.error(f"❌ El archivo no es un JSON válido: {e}")
        st.stop()

    # Mostrar info básica del proceso
    nombre     = data.get("nombreInstancia", "Sin nombre")
    complemento = data.get("complementoNombre", "")
    nombre_corto = data.get("nombreCortoInstancia", "")

    with st.container(border=True):
        st.markdown("**Proceso de Producción identificado:**")
        st.markdown(f"**{nombre}**")
        if complemento:
            st.markdown(f"{complemento}")
        if nombre_corto:
            st.caption(f"Clave: {nombre_corto}")

    # Generar Word en memoria
    with st.spinner("Generando documento Word..."):
        try:
            doc = build_document(data)
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            word_bytes = buffer.getvalue()
        except Exception as e:
            st.error(f"❌ Error al generar el documento: {e}")
            st.stop()

    # Nombre del archivo de salida
    output_filename = (
        nombre_corto.replace(" ", "_") + ".docx"
        if nombre_corto
        else uploaded_file.name.replace(".json", ".docx")
    )

    st.success("✅ Documento generado correctamente.")

    st.download_button(
        label="⬇️ Descargar documento Word",
        data=word_bytes,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

# ─── Pie de página ────────────────────────────────────────────────────────────
st.markdown(
    '<div class="footer">Sistema Nacional de Información Estadística y Geográfica · INEGI · '
    'Norma Técnica para la elaboración de Metadatos de los Procesos de Producción</div>',
    unsafe_allow_html=True,
)
