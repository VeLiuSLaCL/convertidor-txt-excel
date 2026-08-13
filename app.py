import streamlit as st
import pandas as pd
import io

# Configuración básica de la página
st.set_page_config(page_title="Procesador de Reportes TXT", page_icon="⚙️", layout="wide")

st.title("⚙️ Procesador de Reportes TXT a Excel")
st.caption("Paso 1: Lectura de ancho fijo y conversión a Excel limpia.")

# Cargar archivo
uploaded_file = st.file_uploader("📂 Arrastra tu archivo .txt aquí", type=["txt", "log"])

# Definición de coordenadas exactas para archivo de Ancho Fijo
colspecs = [
    (0, 21),    # No. Contrato
    (21, 42),   # Cuenta de cheques
    (42, 75),   # Conjunto
    (75, 85),   # Etapa
    (85, 219),  # Domicilio principal
    (219, 271), # Propiedad Nomenclatura
    (271, 287), # Valor vivienda
    (287, 303), # Capital
    (303, 319), # Intereses
    (319, 335), # Saldo a pagar
    (335, 353), # Estatus vivienda
    (353, 368), # % Avance de obra
    (368, 381), # % Ministrado
    (381, 398), # Valor ministrado
    (398, 415)  # Fecha ultima modificación
]

headers = [
    "No. Contrato", "Cuenta de cheques", "Conjunto", "Etapa", "Domicilio principal",
    "Propiedad Nomenclatura", "Valor vivienda", "Capital", "Intereses", "Saldo a pagar",
    "Estatus vivienda", "% Avance de obra", "% Ministrado", "Valor ministrado", "Fecha ultima modificación"
]

if uploaded_file is not None:
    st.info(f"📄 Archivo cargado: **{uploaded_file.name}** ({round(uploaded_file.size/1024, 2)} KB)")
    
    if st.button("🚀 Procesar y Convertir a Excel", type="primary"):
        with st.spinner("Decodificando estructura del archivo..."):
            try:
                # Convertir el archivo a un buffer de texto directo para evitar cortes de red
                bytes_data = uploaded_file.getvalue()
                string_io = io.StringIO(bytes_data.decode("utf-8", errors="ignore"))
                
                # Lectura pandas
                df = pd.read_fwf(
                    string_io,
                    colspecs=colspecs,
                    header=None,
                    names=headers,
                    skiprows=2
                )
                
                # Limpiar espacios en blanco innecesarios
                df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
                
                st.success("¡Estructura leída correctamente!")
                
                # Vista previa
                st.subheader("📊 Vista previa (Primeros 10 registros)")
                st.dataframe(df.head(10), use_container_width=True)

                # Exportación a Excel en memoria
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Reporte')
                    
                    # Formato de ancho automático de celdas
                    worksheet = writer.sheets['Reporte']
                    for col in worksheet.columns:
                        max_len = max(len(str(cell.value or '')) for cell in col)
                        col_letter = col[0].column_letter
                        worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

                output.seek(0)

                # Botón para descargar el Excel resultante
                st.download_button(
                    label="📥 Descargar Excel Generado (.xlsx)",
                    data=output,
                    file_name="Reporte_Convertido.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Error durante el procesamiento: {e}")
