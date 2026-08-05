import os
import re
import json
import io
import pandas as pd
import streamlit as st
from pypdf import PdfReader
import google.generativeai as genai
import openpyxl

# Configuración de página Streamlit
st.set_page_config(
    page_title="Extractor Biopsias Mama - Unidad Radiología",
    page_icon="🩺",
    layout="wide"
)

st.title("🩺 Gestor de Biopsias de Mama - Extracción y Registro")
st.markdown("""
Esta aplicación procesa los informes PDF (`...RX`, `...INCIS`, `...ESCIS`), extrae la información estructurada
según las reglas médicas de la unidad y anexa los datos directamente a tu libro de Excel original.
""")

# Obtener clave de secrets si existe como valor por defecto
default_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        default_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

# Barra lateral de configuración
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
- TAMAÑO_RX: Tamaño máximo de la lesión expresado OBLIGATORIAMENTE EN CENTÍMETROS (cm). Si en el informe viene en milímetros (mm), conviértelo a centímetros (ejemplo: 15 mm -> 1.5 cm o 1.5). Si no se describe en MX pero sí en Eco, usar Eco. Si solo en RM, usar RM.
- MULTICENTRICO: "Sí" o "No" (múltiples lesiones confirmadas en cuadrantes distintos).
- MULTIFOCAL: "Sí" o "No" (varias lesiones en el mismo cuadrante).
- BIRADS: Formato OBLIGATORIO escribiendo siempre 'BIRADS' seguido de un espacio y el número correspondiente (Opciones válidas: 'BIRADS 2', 'BIRADS 3', 'BIRADS 4', 'BIRADS 5'). Ejemplo: si pone V o 5, debes escribir 'BIRADS 5'.
- RM: "Sí" o "No" (indicar si aporta más datos o "No realizada" si no consta).
- ECO_AXILA_ACTO_UNICO: "Sí" o "No".
- SOSPECHA_ECO_AXILAR: "Sí" o "No" (según informe radiológico).
- PAAF_AXILA: "Sí" o "No" (con independencia del resultado).
- BAG_AXILA: "Sí" o "No" (con independencia del resultado).

De los informes finalizados en INCIS / Incis / incis (Informe Anatomopatológico):
- RESULTADO_AP_MAMA: Resultado AP de la lesión mamaria. Si se trata de un carcinoma ductal infiltrante, se debe especificar el GRADO HISTOLÓGICO si está disponible en el informe (ejemplo: 'Carcinoma ductal infiltrante G2' o 'Carcinoma ductal inteligente Grado 2').
- SUBCLASIFICACION_MOLECULAR: Marcadores moleculares (Luminal, HER2, Triple Negativo, etc.).
- ESTADIO_ECOGRAFICO_AXILAR: N0, N1 o N2 (N1/N2 si biopsia/PAAF axilar es positiva).
- RESULTADO_BIOPSIA_AXILAR: Positivo o Negativo para metástasis.

De los informes finalizados en ESCIS / Escis / escis (Hoja Quirúrgica / Escisional):
- GANGLIO_CENTINELA: "Sí" o "No".
- RESULTADO_GANGLIO_CENTINELA: "Sí" si es positivo, "MICROMETÁSTASIS" si consta como tal. Células aisladas = Negativo.
- VACIAMIENTO_AXILAR: "Sí" o "No".
- RESULTADO_VACIAMIENTO: Si VACIAMIENTO_AXILAR es "No", la columna RESULTADO_VACIAMIENTO debe quedar COMPLETAMENTE VACÍA (null/sin texto). Si se hizo ("Sí"), indicar Positivo o Negativo.
- Nota: Si el resultado AP final quirúrgico difiere del incisional, predomina el quirúrgico.

REGLAS ESTRUCTURALES DE SALIDA:
- Devuelve ÚNICAMENTE un objeto JSON válido.
"""

col1, col2 = st.columns(2)

with col1:
    st.markdown("### 1. Archivo Excel Registro (Opcional)")
    excel_file = st.file_uploader("Sube tu Excel máster existente (.xlsx):", type=["xlsx"])

with col2:
    st.markdown("### 2. Informes PDF de la paciente")
    uploaded_files = st.file_uploader("Sube los PDFs (RX, INCIS, ESCIS):", type=["pdf"], accept_multiple_files=True)

if uploaded_files and api_key:
    if st.button("🚀 Procesar Documentos y Extraer Datos"):
        with st.spinner("Analizando documentos con el modelo médico de Gemini..."):
            try:
                clean_api_key = api_key.strip()
                genai.configure(api_key=clean_api_key)

                texts = {}
                for f in uploaded_files:
                    reader = PdfReader(f)
                    text = "\n".join([page.extract_text() or "" for page in reader.pages])
                    texts[f.name] = text

                # Selección de modelo válido y configuración JSON
                selected_model = "gemini-2.0-flash"

                model = genai.GenerativeModel(
                    model_name=selected_model,
                    system_instruction=SYSTEM_INSTRUCTIONS,
                    generation_config={
                        "temperature": 0.0,
                        "response_mime_type": "application/json"
                    }
                )

                prompt = f"Extrae los datos de los siguientes informes médicos siguiendo estrictamente las instrucciones del sistema:\n\n{json.dumps(texts, ensure_ascii=False)}"
                response = model.generate_content(prompt)

                st.session_state["result_json"] = response.text
                st.success("✅ ¡Extracción completada! Revisa los datos en la tabla inferior.")

            except Exception as e:
                st.error(f"⚠️ Error al conectar con la API de Gemini: {str(e)}")

# Sección de revisión y exportación
if "result_json" in st.session_state:
    st.markdown("---")
    st.subheader("📋 Previsualización y Edición del Caso Extraído")
    st.info("👇 **Revisa aquí los datos extraídos.** Puedes modificar cualquier celda directamente en la tabla.")
    
    try:
        raw_json = st.session_state["result_json"]
        
        # Limpieza segura de marcas de bloque JSON Markdown
        cleaned_json = re.sub(r"^```json\s*", "", raw_json, flags=re.MULTILINE)
        cleaned_json = re.sub(r"^```\s*", "", cleaned_json, flags=re.MULTILINE).strip()
        
        data = json.loads(cleaned_json)
        
        if isinstance(data, dict):
            new_df = pd.DataFrame([data])
        else:
            new_df = pd.DataFrame(data)
            
        edited_df = st.data_editor(new_df, num_rows="dynamic", use_container_width=True)
        
        st.markdown("---")
        
        # Si se subió un Excel previo
        if excel_file is not None:
            excel_bytes = excel_file.getvalue()
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
            
            target_sheet_name = str(selected_year).strip()
            
            # Si no existe la pestaña del año seleccionado, la creamos
            if target_sheet_name not in wb.sheetnames:
                ws = wb.create_sheet(title=target_sheet_name)
                # Escribir cabeceras
                ws.append(list(edited_df.columns))
            else:
                ws = wb[target_sheet_name]
            
            # Anexar las filas editadas
            for _, row in edited_df.iterrows():
                ws.append(row.tolist())
            
            # Guardar en buffer de salida
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            final_excel = output_buffer.getvalue()

            st.download_button(
                label=f"📥 Confirmar Visto Bueno y Descargar Excel Completo Actualizado ({selected_year})",
                data=final_excel,
                file_name=f"Registro_Biopsias_Actualizado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            # Si no subió Excel, se descarga como CSV
            csv = edited_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Confirmar Visto Bueno y Descargar Caso como CSV",
                data=csv,
                file_name=f"caso_extraido_{selected_year}.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"Error al procesar los datos: {e}")
