import io
import pandas as pd
import streamlit as st

# ------------------------------------------------------------------------------
# 1. CONFIGURACIÓN E INICIALIZACIÓN DE LA SESIÓN
# ------------------------------------------------------------------------------
st.set_page_config(page_title="Procesador de Informes de Pacientes", page_icon="📑", layout="wide")

st.title("📑 Procesador de Informes de Pacientes")

# Claves de estado
if "pdf_uploader_key" not in st.session_state:
    st.session_state.pdf_uploader_key = 0

if "excel_acumulado" not in st.session_state:
    st.session_state.excel_acumulado = None

# Nuevo estado para guardar la vista previa temporal que puedes editar
if "datos_previsualizacion" not in st.session_state:
    st.session_state.datos_previsualizacion = None


def convertir_df_a_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Registro_Pacientes")
    return output.getvalue()


# ------------------------------------------------------------------------------
# 2. ENTRADA DE ARCHIVOS
# ------------------------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Informes PDF (Caso Actual)")
    archivos_pdf = st.file_uploader(
        "Sube los informes PDF del paciente:",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"pdf_uploader_{st.session_state.pdf_uploader_key}",
    )

with col2:
    st.subheader("2. Archivo Excel Registro (Base u Opcional)")
    if st.session_state.excel_acumulado is not None:
        st.success(f"📊 Excel activo con {len(st.session_state.excel_acumulado)} registros acumulados.")
        with st.expander("Ver contenido del Excel acumulado"):
            st.dataframe(st.session_state.excel_acumulado)
    else:
        excel_subido = st.file_uploader(
            "Carga un Excel base existente (opcional):",
            type=["xlsx"],
            key="excel_inicial_uploader"
        )
        if excel_subido is not None:
            st.session_state.excel_acumulado = pd.read_excel(excel_subido)
            st.rerun()

st.divider()

# ------------------------------------------------------------------------------
# 3. PASO 1: LEER PDF Y EXTRAER A PREVISUALIZACIÓN
# ------------------------------------------------------------------------------
if st.button("🔍 1. Leer y Previsualizar PDF", type="primary", use_container_width=True):
    if archivos_pdf:
        with st.spinner("Leyendo informe y preparando datos para revisión..."):
            # ==================================================================
            # AQUÍ SE EJECUTA LA EXTRACCIÓN DE TU PDF
            # Reemplaza esta simulación por la llamada a tu función/IA de lectura
            # ==================================================================
            nombres_archivos = ", ".join([f.name for f in archivos_pdf])
            
            # Generamos la fila extraída como DataFrame borrador
            df_extraido = pd.DataFrame([{
                "Paciente": "Nombre Paciente (Editar aquí)",
                "Fecha": pd.Timestamp.now().strftime("%Y-%m-%d"),
                "Archivos": nombres_archivos,
                "Observaciones": "Revisar datos extraídos"
            }])
            
            st.session_state.datos_previsualizacion = df_extraido
            st.rerun()
    else:
        st.warning("⚠️ Por favor, sube al menos un archivo PDF primero.")


# ------------------------------------------------------------------------------
# 4. PASO 2: EDITAR / CORREGIR DATOS Y CONFIRMAR EN EL EXCEL
# ------------------------------------------------------------------------------
if st.session_state.datos_previsualizacion is not None:
    st.info("✏️ **Previsualización del nuevo caso:** Haz doble clic sobre cualquier celda para corregir la información antes de guardar.")
    
    # Editor interactivo para corregir celdas directamente en la pantalla
    df_editado = st.data_editor(
        st.session_state.datos_previsualizacion,
        num_rows="dynamic",
        use_container_width=True,
        key="editor_datos_caso"
    )

    col_conf1, col_conf2 = st.columns([2, 1])
    
    with col_conf1:
        if st.button("✅ 2. Confirmar y Añadir este caso al Excel", use_container_width=True):
            # Guardamos los datos corregidos en el Excel acumulado
            if st.session_state.excel_acumulado is not None:
                st.session_state.excel_acumulado = pd.concat(
                    [st.session_state.excel_acumulado, df_editado],
                    ignore_index=True
                )
            else:
                st.session_state.excel_acumulado = df_editado
            
            # Limpiamos la previsualización del borrador
            st.session_state.datos_previsualizacion = None
            st.success("¡Caso guardado correctamente en la tabla acumulada!")
            st.rerun()
            
    with col_conf2:
        if st.button("❌ Descartar este borrador", use_container_width=True):
            st.session_state.datos_previsualizacion = None
            st.rerun()

st.divider()

# ------------------------------------------------------------------------------
# 5. DESCARGAR Y REINICIAR
# ------------------------------------------------------------------------------
col_pie1, col_pie2 = st.columns(2)

with col_pie1:
    if st.session_state.excel_acumulado is not None:
        excel_bytes = convertir_df_a_excel(st.session_state.excel_acumulado)
        st.download_button(
            label="📥 Descargar Excel Actualizado",
            data=excel_bytes,
            file_name="Registro_Pacientes_Actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="boton_descarga_final"
        )
    else:
        st.button("📥 Descargar Excel Actualizado", disabled=True, use_container_width=True)

with col_pie2:
    if st.button("🔄 Reiniciar / Nuevo paciente", use_container_width=True):
        # Limpia los PDFs y la vista previa borrador, pero conserva el Excel acumulado
        st.session_state.pdf_uploader_key += 1
        st.session_state.datos_previsualizacion = None
        st.rerun()
