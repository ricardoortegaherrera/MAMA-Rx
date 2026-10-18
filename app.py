import io
import pandas as pd
import streamlit as st

# ------------------------------------------------------------------------------
# 1. CONFIGURACIÓN E INICIALIZACIÓN DE LA SESIÓN
# ------------------------------------------------------------------------------
st.set_page_config(page_title="Validación de Caso por Campos", page_icon="📋", layout="wide")

st.title("📋 Procesador de Informes y Validación de Datos")

# Claves de estado de la sesión
if "pdf_uploader_key" not in st.session_state:
    st.session_state.pdf_uploader_key = 0

if "excel_acumulado" not in st.session_state:
    st.session_state.excel_acumulado = None

# Guardará la interpretación leída (diccionario con cada columna)
if "datos_interpretados" not in st.session_state:
    st.session_state.datos_interpretados = None


def convertir_df_a_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Registro_Pacientes")
    return output.getvalue()


# ------------------------------------------------------------------------------
# 2. ENTRADA DE ARCHIVOS Y ESTADO DEL EXCEL
# ------------------------------------------------------------------------------
col_pdf, col_excel = st.columns(2)

with col_pdf:
    st.subheader("1. Cargar Informes PDF")
    archivos_pdf = st.file_uploader(
        "Selecciona los documentos PDF del paciente:",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"pdf_uploader_{st.session_state.pdf_uploader_key}"
    )

with col_excel:
    st.subheader("2. Archivo Excel Registro")
    if st.session_state.excel_acumulado is not None:
        st.success(f"📊 Excel en memoria con {len(st.session_state.excel_acumulado)} caso(s) acumulado(s).")
    else:
        excel_subido = st.file_uploader(
            "Cargar Excel base existente (Opcional):",
            type=["xlsx"],
            key="excel_inicial_uploader"
        )
        if excel_subido is not None:
            st.session_state.excel_acumulado = pd.read_excel(excel_subido)
            st.rerun()

st.divider()

# ------------------------------------------------------------------------------
# 3. PASO 1: LEER PDF E INTERPRETAR COLUMNAS
# ------------------------------------------------------------------------------
if st.button("🔍 Leer y Analizar Informe PDF", type="primary", use_container_width=True):
    if archivos_pdf:
        with st.spinner("Procesando documento y leyendo campos por columna..."):
            # ==================================================================
            # AQUÍ SE EJECUTA LA EXTRACCIÓN DE TUS COLUMNAS DESDE EL PDF
            # Reemplaza este diccionario por el resultado real de tu lectura/IA
            # ==================================================================
            st.session_state.datos_interpretados = {
                "Nombre_Paciente": "Juan Pérez González",
                "DNI_NIE": "12345678X",
                "Fecha_Informe": "2026-10-18",
                "Diagnostico_Principal": "Hipertensión Arterial Primaria",
                "Tratamiento_Recomendado": "Enalapril 10mg / 24h",
                "Observaciones": "Sin alergias conocidas"
            }
            st.rerun()
    else:
        st.warning("⚠️ Por favor, sube al menos un informe PDF.")


# ------------------------------------------------------------------------------
# 4. PASO 2: FORMULARIO DE REVISIÓN / VISTO BUENO CAMPO POR CAMPO
# ------------------------------------------------------------------------------
if st.session_state.datos_interpretados is not None:
    st.subheader("🔍 Revisión de Datos Extraídos (Columna por Columna)")
    st.info("Comprueba la interpretación de cada columna. Puedes corregir cualquier texto antes de dar el visto bueno final.")

    # Formulario interactivo campo por campo
    with st.form("formulario_revision_caso"):
        col_f1, col_f2 = st.columns(2)

        datos = st.session_state.datos_interpretados

        # Columna 1
        with col_f1:
            val_nombre = st.text_input("Columna: Nombre Paciente", value=datos.get("Nombre_Paciente", ""))
            val_dni = st.text_input("Columna: DNI / NIE", value=datos.get("DNI_NIE", ""))
            val_fecha = st.text_input("Columna: Fecha Informe", value=datos.get("Fecha_Informe", ""))

        # Columna 2
        with col_f2:
            val_diag = st.text_input("Columna: Diagnóstico Principal", value=datos.get("Diagnostico_Principal", ""))
            val_trat = st.text_area("Columna: Tratamiento Recomendado", value=datos.get("Tratamiento_Recomendado", ""))
            val_obs = st.text_area("Columna: Observaciones", value=datos.get("Observaciones", ""))

        # Botón de confirmación dentro del formulario
        visto_bueno = st.form_submit_button("✅ Dar Visto Bueno y Guardar en Excel", use_container_width=True)

        if visto_bueno:
            # Construimos la fila con los valores revisados/corregidos
            fila_
