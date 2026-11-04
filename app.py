import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Estudio Cherubin - DDJJ IVA & Control",
    layout="wide"
)

st.title("Estudio Cherubin - Control y Liquidación de IVA")

def parsear_txt_comprobantes(file_bytes):
    """
    Lee un archivo TXT delimitado por tabulaciones (Ventas o Compras),
    limpia encabezados y convierte columnas numéricas.
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
        raise ValueError("El archivo está vacío o no tiene un formato válido.")

    df = pd.read_csv(io.StringIO('\n'.join(lines)), sep='\t', dtype=str)
    df.columns = [col.strip() for col in df.columns]

    # Limpieza de columnas de importes
    columnas_numericas = ['Neto Gravado', 'Exento', 'No Gravado', 'No Alcanzado', 'IVA', 'Total del concepto', 'Total']
    for col in columnas_numericas:
        if col in df.columns:
            df[col] = (
                df[col]
                .fillna('0')
                .str.replace('.', '', regex=False)
                .str.replace(',', '.', regex=False)
                .astype(float)
            )

    # Limpieza de alícuotas
    col_alic = '%Alícuota' if '%Alícuota' in df.columns else ('%Alcuota' if '%Alcuota' in df.columns else None)
    if col_alic:
        df['%Alícuota'] = (
            df[col_alic]
            .fillna('0')
            .str.replace('.', '', regex=False)
            .str.replace(',', '.', regex=False)
            .astype(float)
        )
    else:
        df['%Alícuota'] = 0.0

    # Determinar el signo según el tipo de comprobante (Notas de Crédito restan)
    if 'Comprobante' in df.columns:
        df['Signo'] = df['Comprobante'].apply(lambda x: -1.0 if 'Nota de Crédito' in str(x) or 'Nota de Crdito' in str(x) else 1.0)
    else:
        df['Signo'] = 1.0

    return df

# --- INTERFAZ DE CARGA ---
col_u1, col_u2 = st.columns(2)

with col_u1:
    file_ventas = st.file_uploader("📂 Cargar VENTAS.TXT", type=["txt", "TXT", "xlsx", "csv"])

with col_u2:
    file_compras = st.file_uploader("📂 Cargar COMPRAS.TXT", type=["txt", "TXT", "xlsx", "csv"])

df_v = None
df_c = None

if file_ventas is not None:
    try:
        if file_ventas.name.lower().endswith('.txt'):
            df_v = parsear_txt_comprobantes(file_ventas.getvalue())
        else:
            df_v = pd.read_excel(file_ventas)
        st.success(f"Ventas cargadas: {len(df_v)} registros")
    except Exception as e:
        st.error(f"Error en Ventas: {e}")

if file_compras is not None:
    try:
        if file_compras.name.lower().endswith('.txt'):
            df_c = parsear_txt_comprobantes(file_compras.getvalue())
        else:
            df_c = pd.read_excel(file_compras)
        st.success(f"Compras cargadas: {len(df_c)} registros")
    except Exception as e:
        st.error(f"Error en Compras: {e}")

st.markdown("---")

# --- PANEL PRINCIPAL DE LIQUIDACIÓN ---
if df_v is not None or df_c is not None:
    st.header("📊 Posición Mensual de IVA")

    # Cálculos Ventas (Débito Fiscal)
    neto_v = 0.0
    debito_iva = 0.0
    if df_v is not None:
        neto_v = (df_v['Neto Gravado'] * df_v['Signo']).sum() if 'Neto Gravado' in df_v.columns else 0.0
        debito_iva = (df_v['IVA'] * df_v['Signo']).sum() if 'IVA' in df_v.columns else 0.0

    # Cálculos Compras (Crédito Fiscal)
    neto_c = 0.0
    credito_iva = 0.0
    if df_c is not None:
        neto_c = (df_c['Neto Gravado'] * df_c['Signo']).sum() if 'Neto Gravado' in df_c.columns else 0.0
        credito_iva = (df_c['IVA'] * df_c['Signo']).sum() if 'IVA' in df_c.columns else 0.0

    saldo_tecnico = debito_iva - credito_iva

    # Métricas Resumen
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Débito Fiscal (Ventas)", f"$ {debito_iva:,.2f}")
    m2.metric("Crédito Fiscal (Compras)", f"$ {credito_iva:,.2f}")
    m3.metric("Saldo Técnico", f"$ {abs(saldo_tecnico):,.2f}", 
              delta=f"{'A Pagar' if saldo_tecnico > 0 else 'A Favor'}",
              delta_color="inverse" if saldo_tecnico > 0 else "normal")
    m4.metric("Neto Facturado Ventas", f"$ {neto_v:,.2f}")

    st.markdown("---")

    # --- PESTAÑAS DE DETALLE Y DESGLOSE ---
    tab1, tab2, tab3 = st.columns(3)

    st.subheader("📋 Detalle de Archivos Cargados")
    t_ventas, t_compras = st.tabs(["Ventas", "Compras"])

    with t_ventas:
        if df_v is not None:
            st.dataframe(df_v, use_container_width=True)
            
            # Exportar Excel de Ventas
            out_v = io.BytesIO()
            with pd.ExcelWriter(out_v, engine='openpyxl') as w:
                df_v.to_excel(w, index=False, sheet_name='Ventas')
            st.download_button("📥 Descargar Ventas procesadas (Excel)", data=out_v.getvalue(), file_name="VENTAS_PROCESADO.xlsx")
        else:
            st.info("Sin archivo de Ventas cargado.")

    with t_compras:
        if df_c is not None:
            st.dataframe(df_c, use_container_width=True)
            
            # Exportar Excel de Compras
            out_c = io.BytesIO()
            with pd.ExcelWriter(out_c, engine='openpyxl') as w:
                df_c.to_excel(w, index=False, sheet_name='Compras')
            st.download_button("📥 Descargar Compras procesadas (Excel)", data=out_c.getvalue(), file_name="COMPRAS_PROCESADO.xlsx")
        else:
            st.info("Sin archivo de Compras cargado.")

else:
    st.info("Cargá el archivo **VENTAS.TXT** y/o **COMPRAS.TXT** para visualizar la DDJJ y el análisis de la posición mensual.")
