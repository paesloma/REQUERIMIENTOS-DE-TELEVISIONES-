import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Gestor de Pedidos Excel", layout="wide")

# --- LÓGICA DE TRANSFORMACIÓN DE MODELO ---
def limpiar_modelo(nombre_producto):
    if pd.isna(nombre_producto):
        return "S/N"
    # Eliminar "TELEVISOR" o "TELEVISION"
    temp = re.sub(r'TELEVISOR|TELEVISION', '', str(nombre_producto), flags=re.IGNORECASE).strip()
    # Buscar bloque alfanumérico para identificar el modelo técnico
    match = re.search(r'([A-Z0-9]+)', temp)
    if match:
        modelo_base = match.group(1)
        # Formato solicitado: Prefijo PL + los primeros 5 caracteres
        return f"PL-{modelo_base[:5]}"
    return "OTROS"

# --- INTERFAZ DE USUARIO ---
st.title("📦 Procesador de Pedidos desde Excel")

# 1. CARGA DEL ARCHIVO EXCEL DESDE LA APP
st.subheader("1. Sube tu base de datos (Archivo .xlsx o .xls)")
uploaded_file = st.file_uploader("Selecciona el archivo de Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # LEER EXCEL (Se usa engine='openpyxl' para .xlsx)
        df_master = pd.read_excel(uploaded_file)
        
        # Estandarizar la columna de orden a texto para evitar decimales (.0)
        col_orden = '#Orden' # Ajustar si el nombre varía en tu archivo
        if col_orden in df_master.columns:
            df_master[col_orden] = df_master[col_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        st.success("✅ Archivo de Excel cargado correctamente.")

        # 2. INGRESO MANUAL DE ÓRDENES
        st.divider()
        st.subheader("2. Órdenes para generar pedido")
        input_ordenes = st.text_area(
            "Pega aquí los números de orden de tu cuadro de Excel (uno por línea):",
            placeholder="Ejemplo:\n28712\n28711",
            height=200
        )

        if input_ordenes:
            # Limpiar la lista ingresada por el usuario
            lista_solicitada = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
            
            # Filtrar las órdenes solicitadas en la base de datos subida
            df_resultado = df_master[df_master[col_orden].isin(lista_solicitada)].copy()

            if not df_resultado.empty:
                # Aplicar la regla de código PL
                df_resultado['CODIGO_PEDIDO'] = df_resultado['Producto'].apply(limpiar_modelo)
                
                # Seleccionar columnas para el reporte final
                # Ajusta estos nombres según las columnas reales de tu Excel
                columnas = [col_orden, 'CODIGO_PEDIDO', 'Producto', 'Repuestos', 'Estado']
                columnas_existentes = [c for c in columnas if c in df_resultado.columns]
                df_final = df_resultado[columnas_existentes]

                st.subheader("3. Vista Previa del Pedido Generado")
                st.dataframe(df_final, use_container_width=True)

                # --- BOTÓN PARA DESCARGAR EL NUEVO EXCEL ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Pedido')
                buffer.seek(0)

                st.download_button(
                    label="📥 Descargar Lista de Pedido (.xlsx)",
                    data=buffer,
                    file_name=f"pedido_generado_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.warning("⚠️ No se encontró ninguna de las órdenes ingresadas en el archivo de Excel subido.")

    except Exception as e:
        st.error(f"Error al procesar el archivo Excel: {e}")
else:
    st.info("Por favor, sube un archivo Excel (.xlsx o .xls) para comenzar.")
