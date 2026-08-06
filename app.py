import os
import re
import json
import io
import requests
import pandas as pd
import streamlit as st
from pypdf import PdfReader
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
- FECHA_BIOPSIA: Fecha de la realización/firma de la prueba en el informe radiológico (DD/MM/AAAA). Ignorar fechas de laboratorio de anatomía patológica.
- SCREENING: Marcar OBLIGATORIAMENTE "Sí" o "No". Si procede de screening, PDPCM, cribado o programa de detección precoz, pon siempre "Sí".
- MAMOGRAFIAS_PREVIAS_SCREENING: Dejar SIEMPRE VACÍO ("").
- ALGUNA_MAMOGRAFIA_PREVIA_SCREENING: Dejar SIEMPRE VACÍO ("").
- CLINICA: Opciones del desplegable o "Asintomática" si no consta síntoma.
- ANTECEDENTES_MISMA_MAMA: "Sí" o "No" (en duda poner "No").
- ANTECEDENTES_CONTRALATERAL: "Sí" o "No" (en duda poner "No").
- ASPECTO_MAMOGRAFICO: Descriptor mamográfico. Si solo hay eco, indicar "No mamografía".
- MAMOGRAFIA_CON_CONTRASTE: "Sí" o "No" (MCE / MC).
- BAV: "Sí" o "No" (Biopsia asistida por vacío / estereotaxia).
- TAMAÑO_RX: Tamaño máximo de la lesión expresado OBLIGATORIAMENTE EN CENTÍMETROS (cm). Si en el informe viene en milímetros (mm), conviértelo a centímetros (ejemplo: 18 mm -> 1.8). Si no se describe en MX pero sí en Eco, usar Eco.
- MULTICENTRICO: "Sí" o "No".
- MULTIFOCAL: "Sí" o "No".
- BIRADS: Formato OBLIGATORIO escribiendo siempre 'BIRADS' seguido de un espacio y el número correspondiente (ejemplo: 'BIRADS 5').
- RM: "Sí" o "No" (indicar si aporta más datos o "No realizada" si no consta).
- ECO_AXILA_ACTO_UNICO: "Sí" o "No".
- SOSPECHA_ECO_AXILAR: "Sí" o "No". Si no se aprecian adenopatías sospechosas en ecografía, marcar "No".
- PAAF_AXILA: "Sí" o "No".
- BAG_AXILA: "Sí" o "No".

De los informes finalizados en INCIS / Incis / incis (Informe Anatomopatológico):
- RESULTADO_AP_MAMA: Resultado AP de la lesión mamaria. Si se trata de un carcinoma ductal infiltrante, especificar el GRADO HISTOLÓGICO (ejemplo: 'Carcinoma ductal infiltrante G2').
- SUBCLASIFICACION_MOLECULAR: Marcadores moleculares (Luminal A, Luminal B, HER2, Triple Negativo, etc.).
- ESTADIO_ECOGRAFICO_AXILAR: N0, N1 o N2.
- RESULTADO_BIOPSIA_AXILAR: "Sí" o "No" / "No realizada". Si la axila fue negativa en ecografía o no fue biopsiada, poner "No".

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

def get_available_models(clean_key):
    list_url = f"https://generativelanguage.googleapis.com/v1beta/models?key={clean_key}"
    try:
        res = requests.get(list_url)
        if res.status_code == 200:
            models_data = res.json().get("models", [])
            valid_models = []
            for m in models_data:
                methods = m.get("supportedGenerationMethods", [])
                if "generateContent" in methods:
                    name = m.get("name", "").replace("models/", "")
                    if "gemini" in name:
                        valid_models.append(name)
            return valid_models
    except Exception:
        pass
    return ["gemini-1.5-flash-latest", "gemini-2.0-flash", "gemini-1.5-pro-latest", "gemini-1.5-flash"]

def get_first_empty_row(ws, check_col=2, start_row=2):
    """Encuentra la primera fila verdaderamente vacía basándose en la Columna B (NUHSA)."""
    r = start_row
    while True:
        val = ws.cell(r, check_col).value
        if val is None or str(val).strip() == "":
            return r
        r += 1

if uploaded_files and api_key:
    if st.button("🚀 Procesar Documentos y Extraer Datos"):
        with st.spinner("Analizando documentos..."):
            try:
                clean_api_key = api_key.strip()
                texts = {}
                for f in uploaded_files:
                    reader = PdfReader(f)
                    text = "\n".join([page.extract_text() or "" for page in reader.pages])
                    texts[f.name] = text

                prompt_text = f"Extrae los datos de los siguientes informes médicos siguiendo estrictamente las instrucciones:\n\n{json.dumps(texts, ensure_ascii=False)}"

                model_candidates = get_available_models(clean_api_key)
                
                extracted_text = None
                successful_model = None

                for model_name in model_candidates:
                    for api_ver in ["v1beta", "v1"]:
                        url = f"https://generativelanguage.googleapis.com/{api_ver}/models/{model_name}:generateContent?key={clean_api_key}"
                        headers = {"Content-Type": "application/json"}
                        payload = {
                            "contents": [{"parts": [{"text": prompt_text}]}],
                            "systemInstruction": {"parts": [{"text": SYSTEM_INSTRUCTIONS}]},
                            "generationConfig": {"temperature": 0.0, "responseMimeType": "application/json"}
                        }

                        response = requests.post(url, headers=headers, json=payload)
                        res_json = response.json()

                        if response.status_code == 200 and "candidates" in res_json:
                            try:
                                extracted_text = res_json["candidates"][0]["content"]["parts"][0]["text"]
                                successful_model = f"{model_name} ({api_ver})"
                                break
                            except (KeyError, IndexError):
                                continue
                    if extracted_text:
                        break

                if not extracted_text:
                    st.error("⚠️ No se pudo conectar con ningún modelo.")
                    st.stop()

                st.session_state["result_json"] = extracted_text
                st.success(f"✅ ¡Extracción completada con éxito!")

            except Exception as e:
                st.error(f"⚠️ Error general al procesar: {str(e)}")

# Sección de revisión y exportación
if "result_json" in st.session_state:
    st.markdown("---")
    st.subheader("📋 Previsualización y Edición del Caso Extraído")
    
    try:
        raw_json = st.session_state["result_json"]
        cleaned_json = re.sub(r"^```json\s*", "", raw_json, flags=re.MULTILINE)
        cleaned_json = re.sub(r"^```\s*", "", cleaned_json, flags=re.MULTILINE).strip()
        
        data = json.loads(cleaned_json)
        
        if isinstance(data, dict):
            new_df = pd.DataFrame([data])
        else:
            new_df = pd.DataFrame(data)
        
        # Garantizar que las columnas G y H estén presentes y vacías
        if "MAMOGRAFIAS_PREVIAS_SCREENING" not in new_df.columns:
            new_df.insert(6, "MAMOGRAFIAS_PREVIAS_SCREENING", "")
        else:
            new_df["MAMOGRAFIAS_PREVIAS_SCREENING"] = ""

        if "ALGUNA_MAMOGRAFIA_PREVIA_SCREENING" not in new_df.columns:
            new_df.insert(7, "ALGUNA_MAMOGRAFIA_PREVIA_SCREENING", "")
        else:
            new_df["ALGUNA_MAMOGRAFIA_PREVIA_SCREENING"] = ""
            
        edited_df = st.data_editor(new_df, num_rows="dynamic", use_container_width=True)
        
        st.markdown("---")
        
        if excel_file is not None:
            excel_bytes = excel_file.getvalue()
            wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
            
            target_sheet_name = str(selected_year).strip()
            
            if target_sheet_name not in wb.sheetnames:
                ws = wb.create_sheet(title=target_sheet_name)
            else:
                ws = wb[target_sheet_name]
            
            # Buscamos la fila vacía en la Columna B (NUHSA)
            target_row = get_first_empty_row(ws, check_col=2, start_row=2)
            
            # Escribir fila a fila directamente en las celdas contiguas
            for _, row in edited_df.iterrows():
                row_values = row.tolist()
                for col_idx, val in enumerate(row_values, start=1):
                    # Forzar vacías las columnas 7 (G) y 8 (H)
                    if col_idx in [7, 8]:
                        ws.cell(row=target_row, column=col_idx, value=None)
                    else:
                        ws.cell(row=target_row, column=col_idx, value=val)
                target_row += 1

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
            csv = edited_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Confirmar Visto Bueno y Descargar Caso como CSV",
                data=csv,
                file_name=f"caso_extraido_{selected_year}.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"Error al procesar los datos: {e}")
