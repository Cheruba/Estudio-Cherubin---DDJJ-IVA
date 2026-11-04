import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Estudio Cherubin - Liquidación DDJJ IVA F.2002", layout="wide")

st.title("📊 Estudio Cherubin - Liquidación de IVA (F.2002)")
st.caption("Procesamiento automático según esquema de Ventas, Compras y Determinación F.2002")

# --- FUNCIÓN ROBUSTA PARA LEER EXCEL EN MEMORIA ---
def cargar_excel_seguro(uploaded_file):
    contenido = uploaded_file.read()
    filename = uploaded_file.name.lower()
    
    df = None
    if filename.endswith('.xlsx'):
        try:
            df = pd.read_excel(io.BytesIO(contenido), engine='openpyxl', header=None)
        except Exception:
            pass

    if df is None:
        try:
            df = pd.read_excel(io.BytesIO(contenido), engine='xlrd', header=None)
        except Exception:
            pass

    if df is None:
        try:
            df = pd.read_excel(io.BytesIO(contenido), engine='openpyxl', header=None)
        except Exception:
            pass

    if df is None:
        try:
            dfs = pd.read_html(io.BytesIO(contenido), header=None)
            if dfs:
                df = dfs[0]
        except Exception:
            pass

    if df is None:
        raise ValueError(f"No se pudo leer el archivo '{uploaded_file.name}'.")

    # Buscar automáticamente en qué fila están los encabezados reales (donde esté 'IVA' o 'Comprobante' o 'Neto')
    header_row = 0
    for idx, row in df.iterrows():
        row_str = " ".join([str(val) for val in row.values if pd.notna(val)]).lower()
        if 'neto' in row_str or 'alícuota' in row_str or 'comprobante' in row_str or 'iva' in row_str:
            header_row = idx
            break

    df.columns = [str(c).strip() for c in df.iloc[header_row].values]
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    return df

def clean_num(val):
    if pd.isna(val): 
        return 0.0
    if isinstance(val, (int, float)): 
        return float(val)
    s = str(val).strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except Exception:
        return 0.0

# --- ARCHIVOS SIDEBAR ---
st.sidebar.header("📁 Cargar Archivos del Período")
file_ventas = st.sidebar.file_uploader("1. Listado de VENTAS (.xls / .xlsx)", type=["xls", "xlsx"])
file_compras = st.sidebar.file_uploader("2. Listado de COMPRAS (.xls / .xlsx)", type=["xls", "xlsx"])

st.sidebar.subheader("💳 Percepciones y Anticipos (Opcional)")
percepciones_input = st.sidebar.number_input("Percepciones del Período ($)", min_value=0.0, value=0.0, step=100.0)
retenciones_input = st.sidebar.number_input("Retenciones del Período ($)", min_value=0.0, value=0.0, step=100.0)
anticipos_input = st.sidebar.number_input("Anticipos Pagados ($)", min_value=0.0, value=0.0, step=100.0)

btn_procesar = st.sidebar.button("🚀 LIQUIDAR IVA (F.2002)", use_container_width=True)

if btn_procesar and file_ventas and file_compras:
    with st.spinner("Procesando planillas y agrupando según esquema DDJJ..."):
        try:
            # 1. VENTAS
            df_v = cargar_excel_seguro(file_ventas)
            
            # Buscar columnas dinámicamente
            col_iva_v = [c for c in df_v.columns if c.lower() == 'iva']
            col_cbte_v = [c for c in df_v.columns if 'comprobante' in c.lower()]
            
            col_iva_v_name = col_iva_v[0] if col_iva_v else 'IVA'
            col_cbte_v_name = col_cbte_v[0] if col_cbte_v else df_v.columns[0]

            # Convertir columna IVA a numérico
            if col_iva_v_name in df_v.columns:
                df_v[col_iva_v_name] = df_v[col_iva_v_name].apply(clean_num)
            else:
                df_v[col_iva_v_name] = 0.0

            # Filtrar filas de totales o encabezados vacíos (donde IVA == 0 y cbte no es válido)
            df_v_valid = df_v[df_v[col_iva_v_name] != 0.0].copy()

            # Separar Facturas y Notas de Crédito
            v_nc = df_v_valid[df_v_valid[col_cbte_v_name].astype(str).str.contains('Nota de Crédito|NC', case=False, na=False)]
            v_facturas = df_v_valid[~df_v_valid[col_cbte_v_name].astype(str).str.contains('Nota de Crédito|NC', case=False, na=False)]

            # 2. COMPRAS
            df_c = cargar_excel_seguro(file_compras)

            col_iva_c = [c for c in df_c.columns if c.lower() == 'iva']
            col_cbte_c = [c for c in df_c.columns if 'comprobante' in c.lower()]

            col_iva_c_name = col_iva_c[0] if col_iva_c else 'IVA'
            col_cbte_c_name = col_cbte_c[0] if col_cbte_c else df_c.columns[0]

            if col_iva_c_name in df_c.columns:
                df_c[col_iva_c_name] = df_c[col_iva_c_name].apply(clean_num)
            else:
                df_c[col_iva_c_name] = 0.0

            df_c_valid = df_c[df_c[col_iva_c_name] != 0.0].copy()

            c_nc = df_c_valid[df_c_valid[col_cbte_c_name].astype(str).str.contains('Nota de Crédito|NC', case=False, na=False)]
            c_facturas = df_c_valid[~df_c_valid[col_cbte_c_name].astype(str).str.contains('Nota de Crédito|NC', case=False, na=False)]

            # 3. CÁLCULOS DDJJ F.2002
            iva_v_facturas = float(v_facturas[col_iva_v_name].sum())
            iva_v_nc = float(v_nc[col_iva_v_name].sum())

            iva_c_facturas = float(c_facturas[col_iva_c_name].sum())
            iva_c_nc = float(c_nc[col_iva_c_name].sum())

            # Débito = IVA Ventas Facturas - IVA NC Compras
            debito_fiscal = iva_v_facturas - abs(iva_c_nc)
            
            # Crédito = IVA Compras Facturas - IVA NC Ventas
            credito_fiscal = iva_c_facturas - abs(iva_v_nc)

            saldo_tecnico = debito_fiscal - credito_fiscal
            total_pagos_cuenta = percepciones_input + retenciones_input + anticipos_input
            saldo_afip = saldo_tecnico - total_pagos_cuenta

            # --- VISTA EN PANTALLA ---
            st.header("📋 Borrador Liquidación IVA (F.2002)")
            col1, col2, col3 = st.columns(3)

            col1.metric("Débito Fiscal Total", f"${debito_fiscal:,.2f}")
            col2.metric("Crédito Fiscal Total", f"${credito_fiscal:,.2f}")
            col3.metric(
                "Saldo Técnico", 
                f"${abs(saldo_tecnico):,.2f}", 
                delta="A Favor AFIP" if saldo_tecnico > 0 else "A Favor Contribuyente",
                delta_color="inverse" if saldo_tecnico > 0 else "normal"
            )

            st.markdown("---")
            st.subheader("💵 Determinación Final del Impuesto")

            res_df = pd.DataFrame([
                {"Concepto": "Total Débito Fiscal", "Monto": f"${debito_fiscal:,.2f}"},
                {"Concepto": "Total Crédito Fiscal", "Monto": f"${credito_fiscal:,.2f}"},
                {"Concepto": "Saldo Técnico a Favor de AFIP", "Monto": f"${max(0.0, saldo_tecnico):,.2f}"},
                {"Concepto": "Percepciones del Período", "Monto": f"${percepciones_input:,.2f}"},
                {"Concepto": "Retenciones del Período", "Monto": f"${retenciones_input:,.2f}"},
                {"Concepto": "Anticipos Pagados", "Monto": f"${anticipos_input:,.2f}"},
                {"Concepto": "SALDO A PAGAR A AFIP", "Monto": f"${max(0.0, saldo_afip):,.2f}"}
            ])

            st.table(res_df)

            # --- BOTÓN DESCARGA EXCEL ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_v.to_excel(writer, sheet_name='Ventas', index=False)
                df_c.to_excel(writer, sheet_name='Compras', index=False)
                res_df.to_excel(writer, sheet_name='DDJJ_IVA', index=False)

            st.download_button(
                label="📥 Descargar Liquidación Consolidada en Excel",
                data=output.getvalue(),
                file_name="liquidacion_ddjj_iva_f2002.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Error al procesar los archivos: {e}")
else:
    st.info("👈 Cargá en la barra lateral tus archivos de Ventas y Compras para generar la liquidación F.2002.")
