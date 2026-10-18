import io
import pandas as pd
import streamlit as st

# ------------------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y ESTADO
# ------------------------------------------------------------------------------
st.set_page_config(page_title="Procesador de Informes de Pacientes", page_icon="📑", layout="wide")

st.title("📑 Procesador de Informes de Pacientes")

if "pdf_uploader_key" not in st.session_state:
    st.session_state.pdf_uploader_key = 0

if "excel_acumulado" not in st.session_state:
    st.session_state.excel_acumulado = None


def convertir_df_a_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Registro_Pacientes")
    return output.getvalue()


# ------------------------------------------------------------------------------
# ENTRADA DE ARCHIVOS
# ------------------------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Informes PDF")
    archivos_pdf = st.file_uploader(
        "INFORMES PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"pdf_uploader_{st.session_state.pdf_uploader_key}"
    )

with col2:
    st.subheader("2. Archivo Excel Registro (Opcional)")
    if st.session_state.excel_acumulado is not None:
        st.success("Excel de Registro cargado.")
        st.dataframe(st.session_state.excel_acumulado.head())
    else:
        excel_subido = st.file_uploader("Archivo Excel Registro (Opcional)", type=["xlsx"])
        if excel_subido is not None:
            st.session_state.excel_acumulado = pd.read_excel(excel_subido)

st.divider()

# ------------------------------------------------------------------------------
# PROCESAMIENTO Y ACCIONES
# ------------------------------------------------------------------------------
if st.button("Procesar y Actualizar Excel", type="primary", use_container_width=True):
    if archivos_pdf:
        # Lógica de extracción de datos
        nuevos_datos = pd.DataFrame([{
            "Paciente": "Nuevo Paciente", 
            "Archivos_Procesados": len(archivos_pdf)
        }])
        
        if st.session_state.excel_acumulado is not None:
            st.session_state.excel_acumulado = pd.concat(
                [st.session_state.excel_acumulado, nuevos_datos], 
                ignore_index=True
            )
        else:
            st.session_state.excel_acumulado = nuevos_datos
            
        st.success("¡Datos extraídos y Excel actualizado correctamente!")
    else:
        st.warning("Por favor, sube al menos un informe PDF.")

if st.button("Reiniciar / Nuevo paciente", use_container_width=True):
    st.session_state.pdf_uploader_key += 1
    st.rerun()
