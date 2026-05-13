import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Gestor de Pedidos TCL", layout="wide")

# --- LÓGICA DE TRANSFORMACIÓN ---
def limpiar_modelo(nombre_producto):
    if pd.isna(nombre_producto):
        return "S/N"
    # Eliminar "TELEVISOR" o "TELEVISION"
    temp = re.sub(r'TELEVISOR|TELEVISION', '', str(nombre_producto), flags=re.IGNORECASE).strip()
    # Buscar bloque alfanumérico para el modelo
    match = re.search(r'([A-Z0-9]+)', temp)
    if match:
        modelo_base = match.group(1)
        # Prefijo PL + 5 dígitos
        return f"PL-{modelo_base[:5]}"
    return "OTROS"

# --- INTERFAZ DE USUARIO ---
st.title("📦 Carga y Generación de Pedidos")

# 1. CARGA DEL ARCHIVO CSV DESDE LA APP
st.subheader("1. Sube tu base de datos (Archivo CSV)")
uploaded_file = st.file_uploader("Selecciona el archivo CSV de Órdenes", type=["csv"])

if uploaded_file is not None:
    try:
        # Leer el archivo subido
        df_master = pd.read_csv(uploaded_file)
        
        # Estandarizar columna de orden
        if '#Orden' in df_master.columns:
            df_master['#Orden'] = df_master['#Orden'].astype(str).str.strip()
        
        st.success("✅ Base de datos cargada correctamente.")

        # 2. INGRESO MANUAL DE ÓRDENES
        st.divider()
        st.subheader("2. Ingreso de Órdenes para Pedido")
        input_ordenes = st.text_area(
            "Pega aquí los números de orden (una por línea):",
            placeholder="Ejemplo:\n28712\n28711",
            height=200
        )

        if input_ordenes:
            lista_solicitada = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
            
            # Filtrar
            df_resultado = df_master[df_master['#Orden'].isin(lista_solicitada)].copy()

            if not df_resultado.empty:
                # Aplicar regla PL-XXXXX
                df_resultado['CODIGO_PEDIDO'] = df_resultado['Producto'].apply(limpiar_modelo)
                
                # Columnas finales
                columnas = ['#Orden', 'CODIGO_PEDIDO', 'Producto', 'Fecha', 'Repuestos', 'Estado']
                df_final = df_resultado[columnas]

                st.subheader("3. Vista Previa del Pedido")
                st.dataframe(df_final, use_container_width=True)

                # --- BOTÓN DESCARGA EXCEL ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Pedido')
                buffer.seek(0)

                st.download_button(
                    label="📥 Descargar Pedido en EXCEL",
                    data=buffer,
                    file_name=f"pedido_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.warning("⚠️ No se encontraron las órdenes ingresadas en el archivo subido.")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
else:
    st.info("Por favor, sube el archivo CSV para comenzar.")
