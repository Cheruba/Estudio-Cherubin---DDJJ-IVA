import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Estudio Cherubin - Liquidación DDJJ IVA F.2002", layout="wide")

st.title("📊 Estudio Cherubin - Liquidación de IVA (F.2002)")
st.caption("Procesamiento automático según esquema de Ventas, Compras y Determinación F.2002")

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
        # 1. CARGA Y LIMPIEZA DE VENTAS
        df_v_raw = pd.read_excel(file_ventas, engine='xlrd' if file_ventas.name.endswith('.xls') else None, header=1)
        df_v = df_v_raw[df_v_raw['Fecha Contable'].notnull() & df_v_raw['Comprobante'].astype(str).str.contains('Factura|Nota de Crédito', case=False, na=False)].copy()

        def clean_num(val):
            if pd.isna(val): return 0.0
            if isinstance(val, (int, float)): return float(val)
            return float(str(val).strip().replace('.', '').replace(',', '.'))

        for c in ['Neto Gravado', 'Exento', 'No Gravado', 'IVA', 'Total del Comprobante']:
            if c in df_v.columns:
                df_v[c] = df_v[c].apply(clean_num)

        # Crear Claves Auxiliares de Ventas (Cond, Actividad, Tipo, F, Detalle)
        df_v['Cond'] = df_v['Situación IVA'].apply(lambda x: 1 if 'Responsable Inscripto' in str(x) else 2)
        df_v['Actividad'] = df_v['Cond'].apply(lambda x: 101012 if x == 1 else 472130)
        df_v['Tipo'] = df_v['Cond'].apply(lambda x: 'RI' if x == 1 else 'CF')
        df_v['F'] = df_v['Letra'].astype(str) + df_v['%Alícuota'].astype(str).str.replace('.0', '', regex=False).str.replace('.', ',', regex=False)
        df_v['Detalle'] = df_v['Actividad'].astype(str) + df_v['F'] + df_v['Tipo'] + df_v['Comprobante'].astype(str).str.strip()

        # 2. CARGA Y LIMPIEZA DE COMPRAS
        df_c_raw = pd.read_excel(file_compras, engine='xlrd' if file_compras.name.endswith('.xls') else None, header=1)
        df_c = df_c_raw[df_c_raw['Fecha Contable'].notnull() & df_c_raw['Comprobante'].astype(str).str.contains('Factura|Nota de Crédito|Liq Compra Directa', case=False, na=False)].copy()

        for c in ['Neto Gravado', 'Subtotal Exento', 'No Gravado', 'IVA', 'Subtotal de Percepciones', 'Total del Comprobante']:
            if c in df_c.columns:
                df_c[c] = df_c[c].apply(clean_num)

        df_c['Clave_Compra'] = df_c['Comprobante'].astype(str).str.strip() + df_c['Letra'].astype(str).str.strip() + df_c['%Alícuota'].astype(str).str.replace('.0', '', regex=False).str.replace('.', ',', regex=False)

        # 3. CÁLCULO DE DÉBITO Y CRÉDITO FISCAL SEGÚN FÓRMULAS F.2002
        v_facturas = df_v[~df_v['Comprobante'].astype(str).str.contains('Nota de Crédito', case=False)]
        v_nc = df_v[df_v['Comprobante'].astype(str).str.contains('Nota de Crédito', case=False)]

        c_facturas = df_c[~df_c['Comprobante'].astype(str).str.contains('Nota de Crédito', case=False)]
        c_nc = df_c[df_c['Comprobante'].astype(str).str.contains('Nota de Crédito', case=False)]

        # F.2002: Débito = IVA Ventas Facturas - IVA NC Compras
        debito_fiscal = v_facturas['IVA'].sum() - abs(c_nc['IVA'].sum())
        
        # F.2002: Crédito = IVA Compras Facturas - IVA NC Ventas
        credito_fiscal = c_facturas['IVA'].sum() - abs(v_nc['IVA'].sum())

        saldo_tecnico = debito_fiscal - credito_fiscal
        total_pagos_cuenta = percepciones_input + retenciones_input + anticipos_input
        saldo_afip = saldo_tecnico - total_pagos_cuenta

        # --- MOSTRAR RESULTADOS EN PANTALLA ---
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

        # --- BOTÓN PARA DESCARGAR EXCEL COMPLETO ---
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
else:
    st.info("👈 Cargá en la barra lateral tus archivos de Ventas y Compras para generar la liquidación F.2002.")
