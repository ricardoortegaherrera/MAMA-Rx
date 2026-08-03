import os
import re
import json
import pandas as pd
import streamlit as st
from pypdf import PdfReader
import google.generativeai as genai

# Configuración de página Streamlit
st.set_page_config(
    page_title="Extractor Biopsias Mama - Unidad Radiología",
    page_icon="🩺",
    layout="wide"
)

st.title("🩺 Gestor de Biopsias de Mama - Extracción y Registro")
st.markdown("""
Esta aplicación procesa los informes PDF (`...RX`, `...INCIS`, `...ESCIS`), extrae la información estructurada
según las reglas médicas de la unidad y permite revisar y descargar la tabla actualizada para tu Excel.
""")

# Obtener clave de secrets si existe como valor por defecto
default_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        default_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

# Muestra siempre la casilla en la barra lateral
st.sidebar.header("⚙️ Configuración")
api_key = st.sidebar.text_input("Ingresa tu clave de API Gemini:", value=default_key, type="password")
selected_year = st.sidebar.selectbox("Año de destino para la hoja:", ["2026", "2025", "2024"])

SYSTEM_INSTRUCTIONS = """
Eres un asistente médico especializado en radiología y senología mamaria. Tu función es extraer exactamente los campos indicados a partir de los informes PDF proporcionados.

NORMAS Y REGLAS DE EXTRACCIÓN OBLIGATORIAS:

De los informes finalizados en RX / Rx / rx (Informe Radiológico):
- NH: Número de historia clínica del informe.
- NUHSA: Número de identificación andaluza de la paciente.
- NOMBRE: Nombre y apellidos en MAYÚSCULAS, sin comas.
- EDAD: Edad reflejada en el informe radiológico.
- FECHA_BIOPSIA: Fecha de la biopsia mamaria (DD/MM/AAAA).
- SCREENING: "Sí" o "No" (marcar "PDPCM" si aparece esa sigla).
- CLINICA: Opciones del desplegable o "Asintomática" si no consta síntoma.
- ANTECEDENTES_MISMA_MAMA: "Sí" o "No" (en duda poner "No").
- ANTECEDENTES_CONTRALATERAL: "Sí" o "No" (en duda poner "No").
- ASPECTO_MAMOGRAFICO: Descriptor mamográfico. Si solo hay eco, indicar "No mamografía".
- MAMOGRAFIA_CON_CONTRASTE: "Sí" o "No" (MCE / MC).
- BAV: "Sí" o "No" (Biopsia asistida por vacío / estereotaxia).
- TAMAÑO_RX: Tamaño máximo en mamografía. Si no se describe en MX pero sí en Eco, usar Eco. Si solo en RM, usar RM.
- MULTICENTRICO: "Sí" o "No" (múltiples lesiones confirmadas en cuadrantes distintos).
- MULTIFOCAL: "Sí" o "No" (varias lesiones en el mismo cuadrante).
- BIRADS: Descriptor BIRADS (ej. IV, V).
- RM: "Sí" o "No" (indicar si aporta más datos o "No realizada" si no consta).
- ECO_AXILA_ACTO_UNICO: "Sí" o "No".
- SOSPECHA_ECO_AXILAR: "Sí" o "No" (según informe radiológico).
- PAAF_AXILA: "Sí" o "No" (con independencia del resultado).
- BAG_AXILA: "Sí" o "No" (con independencia del resultado).

De los informes finalizados en INCIS / Incis / incis (Informe Anatomopatológico):
- RESULTADO_AP_MAMA: Resultado AP de la lesión mamaria.
- SUBCLASIFICACION_MOLECULAR: Marcadores moleculares (Luminal, HER2, Triple Negativo, etc.).
- ESTADIO_ECOGRAFICO_AXILAR: N0, N1 o N2 (N1/N2 si biopsia/PAAF axilar es positiva).
- RESULTADO_BIOPSIA_AXILAR: Positivo o Negativo para metástasis.

De los informes finalizados en ESCIS / Escis / escis (Hoja Quirúrgica / Escisional):
- GANGLIO_CENTINELA: "Sí" o "No".
- RESULTADO_GANGLIO_CENTINELA: "Sí" si es positivo, "MICROMETÁSTASIS" si consta como tal. Células aisladas = Negativo.
- VACIAMIENTO_AXILAR: "Sí" o "No".
- RESULTADO_VACIAMIENTO: Positivo / Negativo.
- Nota: Si el resultado AP final quirúrgico difiere del incisional, predomina el quirúrgico.

DEBES DEVOLVER UN JSON ESTRICTO CON ESTOS CAMPOS PARA CADA PACIENTE ANALIZADA.
"""

uploaded_files = st.file_uploader("Sube o arrastra aquí los PDFs de la paciente (RX, INCIS, ESCIS):", type=["pdf"], accept_multiple_files=True)

if uploaded_files and api_key:
    if st.button("🚀 Procesar Documentos y Extraer Datos"):
        with st.spinner("Analizando documentos con el modelo médico de Gemini..."):
            try:
                texts = {}
                for f in uploaded_files:
                    reader = PdfReader(f)
                    text = "\n".join([page.extract_text() or "" for page in reader.pages])
                    texts[f.name] = text

                clean_api_key = api_key.strip()
                
                # Forzamos la versión v1 de la API para evitar conflictos de v1beta
                genai.configure(api_key=clean_api_key)

                model = genai.GenerativeModel(
                    model_name="models/gemini-1.5-flash-002", # Nombre con prefijo explícito de endpoint
                    system_instruction=SYSTEM_INSTRUCTIONS,
                    generation_config={"response_mime_type": "application/json"}
                )

                prompt = f"Extrae los datos de los siguientes informes médicos siguiendo estrictamente las instrucciones del sistema:\n\n{json.dumps(texts, ensure_ascii=False)}"
                response = model.generate_content(prompt)
                st.session_state["result_json"] = response.text

            except Exception as e:
                st.error(f"⚠️ Error al conectar con la API de Gemini: {str(e)}")

if "result_json" in st.session_state:
    st.subheader("📋 Previsualización y Edición del Caso Extraído")
    st.info("Comprueba los datos antes de exportar. Puedes hacer doble clic en cualquier celda para corregir datos manualmente si hiciera falta.")
    
    try:
        data = json.loads(st.session_state["result_json"])
        if isinstance(data, dict):
            df_preview = pd.DataFrame([data])
        else:
            df_preview = pd.DataFrame(data)
            
        edited_df = st.data_editor(df_preview, num_rows="dynamic", use_container_width=True)
        
        csv = edited_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Descargar datos como CSV (para Excel)",
            data=csv,
            file_name=f"extraccion_biopsia_{selected_year}.csv",
            mime="text/csv"
        )
    except Exception as e:
        st.error(f"Error al procesar la respuesta JSON: {e}")
