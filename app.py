import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Estudio Cherubin - Liquidación DDJJ IVA F.2002", layout="wide")

st.title("📊 Estudio Cherubin - Liquidación de IVA (F.2002)")
st.caption("Procesamiento automático según esquema de Ventas, Compras y Determinación F.2002")

# --- FUNCIÓN ROBUSTA PARA LEER EXCEL EN MEMORIA ---
def cargar_excel_seguro(uploaded_file, header=1):
    contenido = uploaded_file.read()
    filename = uploaded_file.name.lower()
    
    if filename.endswith('.xlsx'):
        try:
            return pd.read_excel(io.BytesIO(contenido), engine='openpyxl', header=header)
        except Exception:
            pass
            
    try:
        return pd.read_excel(io.BytesIO(contenido), engine='xlrd', header=header)
    except Exception:
        pass

    try:
        return pd.read_excel(io.BytesIO(contenido), engine='openpyxl', header=header)
    except Exception:
        pass

    try:
        dfs = pd.read_html(io.BytesIO(contenido), header=header)
        if dfs:
            return dfs[0]
    except Exception:
        pass

    raise ValueError(f"No se pudo determinar el formato del archivo '{uploaded_file.name}'.")

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
            df_v_raw = cargar_excel_seguro(file_ventas, header=1)
            df_v_raw.columns = [str(c).strip() for c in df_v_raw.columns]
            
            col_cbte_v = [c for c in df_v_raw.columns if 'comprobante' in c.lower()]
            col_fecha_v = [c for c in df_v_raw.columns if 'fecha' in c.lower()]
            
            if col_cbte_v and col_fecha_v:
                df_v = df_v_raw[df_v_raw[col_fecha_v[0]].notnull() & df_v_raw[col_cbte_v[0]].astype(str).str.contains('Factura|Nota de Crédito', case=False, na=False)].copy()
            else:
                df_v = df_v_raw.copy()

            for c in ['Neto Gravado', 'Exento', 'No Gravado', 'IVA', 'Total del Comprobante', 'Total del concepto']:
                if c in df_v.columns:
                    df_v[c] = df_v[c].apply(clean_num)

            # 2. COMPRAS
            df_c_raw = cargar_excel_seguro(file_compras, header=1)
            df_c_raw.columns = [str(c).strip() for c in df_c_raw.columns]

            col_cbte_c = [c for c in df_c_raw.columns if 'comprobante' in c.lower()]
            col_fecha_c = [c for c in df_c_raw.columns if 'fecha' in c.lower()]

            if col_cbte_c and col_fecha_c:
                df_c = df_c_raw[df_c_raw[col_fecha_c[0]].notnull() & df_c_raw[col_cbte_c[0]].astype(str).str.contains('Factura|Nota de Crédito|Liq Compra Directa', case=False, na=False)].copy()
            else:
                df_c = df_c_raw.copy()

            for c in ['Neto Gravado', 'Subtotal Exento', 'No Gravado', 'IVA', 'Subtotal de Percepciones', 'Total del Comprobante']:
                if c in df_c.columns:
                    df_c[c] = df_c[c].apply(clean_num)

            # 3. CÁLCULOS DDJJ F.2002
            col_tipo_v = col_cbte_v[0] if col_cbte_v else 'Comprobante'
            col_tipo_c = col_cbte_c[0] if col_cbte_c else 'Comprobante'

            v_facturas = df_v[~df_v[col_tipo_v].astype(str).str.contains('Nota de Crédito', case=False)]
            v_nc = df_v[df_v[col_tipo_v].astype(str).str.contains('Nota de Crédito', case=False)]

            c_facturas = df_c[~df_c[col_tipo_c].astype(str).str.contains('Nota de Crédito', case=False)]
            c_nc = df_c[df_c[col_tipo_c].astype(str).str.contains('Nota de Crédito', case=False)]

            # Asegurar conversión a float explícita para evitar tipos 'str' en sum()
            iva_v_facturas = float(pd.to_numeric(v_facturas['IVA'], errors='coerce').sum()) if 'IVA' in v_facturas.columns else 0.0
            iva_v_nc = float(pd.to_numeric(v_nc['IVA'], errors='coerce').sum()) if 'IVA' in v_nc.columns else 0.0

            iva_c_facturas = float(pd.to_numeric(c_facturas['IVA'], errors='coerce').sum()) if 'IVA' in c_facturas.columns else 0.0
            iva_c_nc = float(pd.to_numeric(c_nc['IVA'], errors='coerce').sum()) if 'IVA' in c_nc.columns else 0.0

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
