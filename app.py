import io
import pandas as pd
import streamlit as st

# ------------------------------------------------------------------------------
# 1. CONFIGURACIÓN DE LA PÁGINA E INICIALIZACIÓN DEL ESTADO (SESSION STATE)
# ------------------------------------------------------------------------------
st.set_page_config(page_title="Procesador de Informes de Pacientes", page_icon="📑", layout="wide")

st.title("📑 Procesador de Informes de Pacientes y Registro Excel")

# Clave dinámica para reiniciar el selector de PDFs sin perder el estado global
if "pdf_uploader_key" not in st.session_state:
    st.session_state.pdf_uploader_key = 0

# DataFrame acumulado que guardará todos los casos procesados
if "excel_acumulado" not in st.session_state:
    st.session_state.excel_acumulado = None


# ------------------------------------------------------------------------------
# 2. FUNCIÓN AUXILIAR PARA GENERAR EL EXCEL EN MEMORIA
# ------------------------------------------------------------------------------
def convertir_df_a_excel(df):
    """Convierte un DataFrame de pandas a bytes binarios de un archivo .xlsx"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Registro_Pacientes")
    return output.getvalue()


# ------------------------------------------------------------------------------
# 3. INTERFAZ DE USUARIO - ENTRADA DE ARCHIVOS
# ------------------------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Informes PDF (Caso Actual)")
    # El selector de PDFs usa la clave dinámica para ser limpiado al reiniciar
    archivos_pdf = st.file_uploader(
        "Sube los informes PDF del paciente:",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"pdf_uploader_{st.session_state.pdf_uploader_key}",
    )

with col2:
    st.subheader("2. Archivo Excel Registro (Base u Opcional)")
    
    # Si ya hay un Excel acumulado en la sesión, mostramos su estado actual
    if st.session_state.excel_acumulado is not None:
        st.success(f"📊 Excel activo con {len(st.session_state.excel_acumulado)} registros.")
        with st.expander("Ver vista previa del Excel acumulado"):
            st.dataframe(st.session_state.excel_acumulado.tail(5))
    else:
        # Si es la primera vez, permite cargar un Excel existente desde el equipo
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
# 4. PROCESAMIENTO Y ACCIONES
# ------------------------------------------------------------------------------
col_acciones1, col_acciones2, col_acciones3 = st.columns(3)

# --- BOTÓN 1: PROCESAR PACIENTE ACTUAL ---
with col_acciones1:
    if st.button("⚡ Procesar y Actualizar Excel", use_container_width=True):
        if archivos_pdf:
            with st.spinner("Procesando información del PDF..."):
                # ==============================================================
                # AQUÍ IRÍA TU LÓGICA / API PARA EXTRAER LOS DATOS DEL PDF
                # Ejemplo simulado de extracción de datos:
                # ==============================================================
                nombres_archivos = ", ".join([f.name for f in archivos_pdf])
                nuevos_datos = pd.DataFrame([{
                    "Fecha_Procesamiento": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
                    "Archivos_Informes": nombres_archivos,
                    "Estado": "Completado"
                }])

                # Concatenamos con el Excel acumulado o creamos uno nuevo
                if st.session_state.excel_acumulado is not None:
                    st.session_state.excel_acumulado = pd.concat(
                        [st.session_state.excel_acumulado, nuevos_datos],
                        ignore_index=True
                    )
                else:
                    st.session_state.excel_acumulado = nuevos_datos

                st.success("¡Datos extraídos y añadidos al Excel!")
                st.rerun()
        else:
            st.warning("⚠️ Por favor, sube al menos un informe PDF.")

# --- BOTÓN 2: DESCARGAR EXCEL ACTUALIZADO ---
with col_acciones2:
    if st.session_state.excel_acumulado is not None:
        excel_bytes = convertir_df_a_excel(st.session_state.excel_acumulado)
        st.download_button(
            label="📥 Descargar Excel Actualizado",
            data=excel_bytes,
            file_name="Registro_Pacientes_Actualizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="boton_descarga_excel",
        )
    else:
        st.button("📥 Descargar Excel Actualizado", disabled=True, use_container_width=True)

# --- BOTÓN 3: REINICIAR / NUEVO PACIENTE ---
with col_acciones3:
    if st.button("🔄 Reiniciar / Nuevo paciente", use_container_width=True):
        # 1. Cambiamos la clave para vaciar únicamente la casilla de subida de PDFs
        st.session_state.pdf_uploader_key += 1
        # 2. Recargamos la interfaz sin borrar st.session_state.excel_acumulado
        st.rerun()
