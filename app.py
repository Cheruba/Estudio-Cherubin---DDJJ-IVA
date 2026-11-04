import io
import os
import re
import pandas as pd
from flask import Flask, render_template_string, request, send_file

app = Flask(__name__)

# Configuración de límites y carpetas
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Procesador de Comprobantes TXT - Estudio Contable</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f8f9fa; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .container { max-width: 800px; margin-top: 50px; }
        .card { border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); border: none; }
        .card-header { background-color: #0d6efd; color: white; border-top-left-radius: 12px !important; border-top-right-radius: 12px !important; }
        .btn-primary { background-color: #0d6efd; border: none; padding: 10px 24px; border-radius: 8px; font-weight: 600; }
        .btn-primary:hover { background-color: #0b5ed7; }
        .file-input-wrapper { border: 2px dashed #cbd5e1; border-radius: 8px; padding: 30px; text-align: center; background: #fff; cursor: pointer; transition: all 0.2s; }
        .file-input-wrapper:hover { border-color: #0d6efd; background: #f0f7ff; }
    </style>
</head>
<body>
    <div class="container">
        <div class="card">
            <div class="card-header text-center py-3">
                <h4 class="mb-0">Procesador de Ventas (TXT)</h4>
            </div>
            <div class="card-body p-4">
                {% if error %}
                    <div class="alert alert-danger" role="alert">
                        {{ error }}
                    </div>
                {% endif %}

                <form method="POST" enctype="multipart/form-data">
                    <div class="mb-4">
                        <label for="archivo_txt" class="form-label fw-bold">Selecciona el archivo VENTAS.TXT:</label>
                        <input type="file" class="form-control" id="archivo_txt" name="archivo_txt" accept=".txt,.TXT" required>
                    </div>

                    <div class="d-grid gap-2">
                        <button type="submit" name="accion" value="excel" class="btn btn-success py-2">
                            📊 Exportar Excel Formateado
                        </button>
                        <button type="submit" name="accion" value="txt_cabecera" class="btn btn-outline-primary py-2">
                            📄 Generar TXT Cabecera (AFIP/ARCA)
                        </button>
                        <button type="submit" name="accion" value="txt_detalle" class="btn btn-outline-secondary py-2">
                            📄 Generar TXT Detalle (AFIP/ARCA)
                        </button>
                    </div>
                </form>
            </div>
        </div>
    </div>
</body>
</html>
"""

def parsear_txt_ventas(file_stream):
    """
    Lee el archivo TXT delimitado por tabulaciones, limpia encabezados y convierte tipos de datos.
    """
    raw_bytes = file_stream.read()
    
    # Intentar decodificar con distintos encodings comunes en Argentina / Windows
    for encoding in ['utf-8', 'latin-1', 'cp1252']:
        try:
            content = raw_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        content = raw_bytes.decode('utf-8', errors='ignore')

    lines = [line for line in content.splitlines() if line.strip() and not line.strip().startswith('Concepto')]
    
    if not lines:
        raise ValueError("El archivo TXT está vacío o no tiene el formato esperado.")

    # Cargar en dataframe usando tabulación como separador
    df = pd.read_csv(io.StringIO('\n'.join(lines)), sep='\t', dtype=str)
    
    # Limpiar nombres de columnas
    df.columns = [col.strip() for col in df.columns]

    # Mapeo y limpieza de columnas numéricas clave
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
        df['%Alícuota'] = (
            df['%Alícuota']
            .fillna('0')
            .str.replace('.', '', regex=False)
            .str.replace(',', '.', regex=False)
            .astype(float)
        )
    elif '%Alcuota' in df.columns:
        df['%Alícuota'] = (
            df['%Alcuota']
            .fillna('0')
            .str.replace('.', '', regex=False)
            .str.replace(',', '.', regex=False)
            .astype(float)
        )
        
    return df

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET':
        return render_template_string(HTML_TEMPLATE)

    if 'archivo_txt' not in request.files:
        return render_template_string(HTML_TEMPLATE, error="No se seleccionó ningún archivo.")

    file = request.files['archivo_txt']
    if file.filename == '':
        return render_template_string(HTML_TEMPLATE, error="Nombre de archivo inválido.")

    try:
        df = parsear_txt_ventas(file.stream)
        accion = request.form.get('accion', 'excel')

        if accion == 'excel':
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Ventas Octubre')
            output.seek(0)
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                download_name='VENTAS_PROCESADO.xlsx'
            )

        elif accion == 'txt_cabecera':
            # Exportación de archivo de cabecera formateado
            output_txt = io.StringIO()
            for _, row in df.iterrows():
                linea = f"{row.get('Fecha de la operación', '')}\t{row.get('Comprobante', '')}\t{row.get('Letra', '')}\t{row.get('Punto de Venta', '')}\t{row.get('Nº de comprobante', row.get('N de comprobante', ''))}\t{row.get('Total del concepto', 0.0):.2f}\n"
                output_txt.write(linea)
            
            mem = io.BytesIO(output_txt.getvalue().encode('utf-8'))
            return send_file(
                mem,
                mimetype='text/plain',
                as_attachment=True,
                download_name='CABECERA_VENTAS.TXT'
            )

        elif accion == 'txt_detalle':
            # Exportación de archivo de detalle formateado
            output_txt = io.StringIO()
            for _, row in df.iterrows():
                linea = f"{row.get('Nº de comprobante', row.get('N de comprobante', ''))}\t{row.get('Neto Gravado', 0.0):.2f}\t{row.get('%Alícuota', 0.0):.2f}\t{row.get('IVA', 0.0):.2f}\n"
                output_txt.write(linea)

            mem = io.BytesIO(output_txt.getvalue().encode('utf-8'))
            return send_file(
                mem,
                mimetype='text/plain',
                as_attachment=True,
                download_name='DETALLE_VENTAS.TXT'
            )

    except Exception as e:
        return render_template_string(HTML_TEMPLATE, error=f"Error al procesar el archivo TXT: {str(e)}")

if __name__ == '__main__':
    # Configuración para ejecución local o en Render
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
