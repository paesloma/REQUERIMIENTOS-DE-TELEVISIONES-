import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Pedidos TCL", layout="wide")

# --- LÓGICA DE TRANSFORMACIÓN DE MODELO ---
def limpiar_modelo(nombre_producto):
    if pd.isna(nombre_producto):
        return "S/N"
    temp = re.sub(r'TELEVISOR|TELEVISION', '', str(nombre_producto), flags=re.IGNORECASE).strip()
    match = re.search(r'([A-Z0-9]+)', temp)
    if match:
        modelo_base = match.group(1)
        return f"PL-{modelo_base[:5]}"
    return "OTROS"

# --- INTERFAZ DE USUARIO ---
st.title("📊 Generador de Pedidos - Formato Requerido")

st.subheader("1. Cargar Base de Datos (.xlsx o .xls)")
uploaded_file = st.file_uploader("Sube el archivo de control de órdenes", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Leer el archivo Excel y limpiar nombres de columnas
        df_master = pd.read_excel(uploaded_file)
        df_master.columns = [str(c).strip() for c in df_master.columns]
        
        # BUSCAR COLUMNA DE ORDEN (flexible)
        posibles_nombres = ['#Orden', 'ORDEN', '# ORDEN', 'Nro Orden', 'Orden']
        col_orden = next((c for c in posibles_nombres if c in df_master.columns), None)

        if col_orden:
            # Limpiar datos de la columna de orden
            df_master[col_orden] = df_master[col_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            st.success(f"✅ Base de datos cargada. Columna detectada: '{col_orden}'")

            st.divider()
            st.subheader("2. Ingrese Lista de Órdenes para el Pedido")
            input_ordenes = st.text_area(
                "Pegue aquí las órdenes (una por línea):",
                placeholder="28712\n28711",
                height=200
            )

            if input_ordenes:
                lista_solicitada = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                df_resultado = df_master[df_master[col_orden].isin(lista_solicitada)].copy()

                if not df_resultado.empty:
                    # Crear columnas necesarias si no existen
                    if 'Serie' not in df_resultado.columns: df_resultado['Serie'] = "N/A"
                    if 'Técnico' not in df_resultado.columns: df_resultado['Técnico'] = "N/A"
                    if 'Repuestos' not in df_resultado.columns: df_resultado['Repuestos'] = "N/A"
                    
                    # Generar Código PL
                    df_resultado['CODIGO_PL'] = df_resultado['Producto'].apply(limpiar_modelo)
                    
                    # RENOMBRAR Y ORDENAR SEGÚN REQUERIMIENTO
                    df_final = df_resultado.rename(columns={
                        col_orden: 'ORDEN',
                        'Serie': 'SERIE',
                        'Producto': 'MODELO',
                        'Técnico': 'TALLER DE PROCEDENCIA',
                        'Repuestos': 'REPUESTO'
                    })[['ORDEN', 'SERIE', 'MODELO', 'TALLER DE PROCEDENCIA', 'REPUESTO', 'CODIGO_PL']]

                    st.subheader("3. Vista Previa del Nuevo Archivo")
                    st.dataframe(df_final, use_container_width=True)

                    # GENERACIÓN DEL EXCEL
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='PEDIDO_TCL')
                    buffer.seek(0)

                    st.download_button(
                        label="📥 DESCARGAR EXCEL (ORDEN SOLICITADO)",
                        data=buffer,
                        file_name=f"Pedido_Formato_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                else:
                    st.warning("⚠️ No se encontraron las órdenes en el archivo cargado. Verifique que los números coincidan.")
        else:
            st.error(f"❌ No se encontró la columna '#Orden'. Las columnas detectadas son: {list(df_master.columns)}")

    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
else:
    st.info("Cargue el archivo Excel para activar el procesador.")
