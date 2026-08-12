import streamlit as st
import pandas as pd
import io

# Configuración de la página
st.set_page_config(page_title="Convertidor TXT a Excel", page_icon="📊", layout="wide")

st.title("📊 Convertidor de Reportes TXT a Excel")
st.markdown("Herramienta para procesar el archivo `.txt` periódico y convertirlo a un Excel estructurado sin espacios sobrantes.")

# Cargar archivo
uploaded_file = st.file_uploader("📂 Selecciona o arrastra tu archivo .txt aquí", type=["txt"])

# Coordenadas exactas de ancho fijo (inicio, fin) para cada columna
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
    try:
        # Lectura del archivo de ancho fijo
        df = pd.read_fwf(
            uploaded_file,
            colspecs=colspecs,
            header=None,
            names=headers,
            skiprows=2
        )
        
        # Limpieza de espacios al inicio/final en texto
        df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
        
        st.success("✅ Archivo procesado correctamente.")
        
        # Vista previa en la interfaz
        st.subheader("👀 Vista previa de los datos")
        st.dataframe(df.head(10), use_container_width=True)

        # Generar el archivo Excel descargable
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Reporte')
            
            # Auto-ajustar el ancho de las columnas dentro del propio Excel generado
            worksheet = writer.sheets['Reporte']
            for col in worksheet.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = col[0].column_letter
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

        output.seek(0)

        # Botón para descargar
        st.download_button(
            label="📥 Descargar Excel limpio (.xlsx)",
            data=output,
            file_name="Reporte_Convertido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Ocurrió un error al procesar el archivo: {e}")