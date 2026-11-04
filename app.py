import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Estudio Cherubin - Resultado Período",
    layout="wide"
)

st.title("Estudio Cherubin - Resultado Período")

def parsear_txt_ventas(file_bytes):
    """
    Lee el archivo TXT delimitado por tabulaciones, limpia encabezados y convierte tipos de datos.
    """
    for encoding in ['utf-8', 'latin-1', 'cp1252']:
        try:
            content = file_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        content = file_bytes.decode('utf-8', errors='ignore')

    lines = [line for line in content.splitlines() if line.strip() and not line.strip().startswith('Concepto')]
    
    if not lines:
        raise ValueError("El archivo TXT está vacío o no tiene el formato esperado.")

    df = pd.read_csv(io.StringIO('\n'.join(lines)), sep='\t', dtype=str)
    df.columns = [col.strip() for col in df.columns]

    columnas_numericas = ['Neto Gravado', 'Exento', 'No Gravado', 'No Alcanzado', 'IVA', 'Total del concepto']
    for col in columnas_numericas:
        if col in df.columns:
            df[col] = (
                df[col]
                .fillna('0')
                .str.replace('.', '', regex=False)
                .str.replace(',', '.', regex=False)
                .astype(float)
            )
        else:
            df[col] = 0.0

    if '%Alícuota' in df.columns:
        df['%Alícuota'] = df['%Alícuota'].fillna('0').str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
    elif '%Alcuota' in df.columns:
        df['%Alícuota'] = df['%Alcuota'].fillna('0').str.replace('.', '', regex=False).str.replace(',', '.', regex=False).astype(float)
        
    return df

uploaded_file = st.file_uploader("Selecciona el archivo VENTAS.TXT", type=["txt"])

if uploaded_file is not None:
    try:
        df = parsear_txt_ventas(uploaded_file.getvalue())
        
        st.success(f"Archivo cargado correctamente. Total de registros: {len(df)}")
        st.dataframe(df.head(10))

        # Generar Excel en memoria
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Ventas Octubre')
        output_excel.seek(0)

        # Generar TXT Cabecera
        output_cabecera = io.StringIO()
        for _, row in df.iterrows():
            linea = f"{row.get('Fecha de la operación', '')}\t{row.get('Comprobante', '')}\t{row.get('Letra', '')}\t{row.get('Punto de Venta', '')}\t{row.get('Nº de comprobante', row.get('N de comprobante', ''))}\t{row.get('Total del concepto', 0.0):.2f}\n"
            output_cabecera.write(linea)

        # Generar TXT Detalle
        output_detalle = io.StringIO()
        for _, row in df.iterrows():
            linea = f"{row.get('Nº de comprobante', row.get('N de comprobante', ''))}\t{row.get('Neto Gravado', 0.0):.2f}\t{row.get('%Alícuota', 0.0):.2f}\t{row.get('IVA', 0.0):.2f}\n"
            output_detalle.write(linea)

        st.subheader("Descargas")
        col1, col2, col3 = st.columns(3)

        with col1:
            st.download_button(
                label="📊 Descargar Excel",
                data=output_excel.getvalue(),
                file_name="VENTAS_PROCESADO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col2:
            st.download_button(
                label="📄 TXT Cabecera",
                data=output_cabecera.getvalue().encode('utf-8'),
                file_name="CABECERA_VENTAS.TXT",
                mime="text/plain"
            )

        with col3:
            st.download_button(
                label="📄 TXT Detalle",
                data=output_detalle.getvalue().encode('utf-8'),
                file_name="DETALLE_VENTAS.TXT",
                mime="text/plain"
            )

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")
