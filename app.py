import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Control de Órdenes y Pedidos", layout="wide")

# --- LÓGICA DE PROCESAMIENTO ---

def limpiar_modelo(nombre_producto):
    """
    Transforma el nombre del producto: elimina 'TELEVISOR', 
    añade prefijo 'PL-' y toma los primeros 5 caracteres del modelo.
    """
    if pd.isna(nombre_producto):
        return "S/N"
    
    # Eliminar la palabra TELEVISOR o TELEVISION
    temp = re.sub(r'TELEVISOR|TELEVISION', '', str(nombre_producto), flags=re.IGNORECASE).strip()
    
    # Buscar el primer bloque alfanumérico que represente el modelo
    match = re.search(r'([A-Z0-9]+)', temp)
    if match:
        modelo_base = match.group(1)
        # Formato PL + 5 caracteres
        return f"PL-{modelo_base[:5]}"
    
    return "OTROS"

@st.cache_data
def load_master_data(file_path):
    """Carga la base de datos principal y estandariza la columna de orden."""
    df = pd.read_csv(file_path)
    # Asegurar que #Orden sea tratado como texto para evitar problemas con ceros o formatos
    if '#Orden' in df.columns:
        df['#Orden'] = df['#Orden'].astype(str).str.strip()
    return df

# --- INTERFAZ DE USUARIO ---

st.title("📦 Sistema de Gestión de Pedidos - 2026")

try:
    # Ruta del archivo (Asegúrate de que coincida con tu archivo en GitHub)
    base_path = "CONTROL ORDENES 2026 TVS TCL.xlsx - ORDENES 2026.csv"
    df_master = load_master_data(base_path)

    # Tabs para organizar las funciones
    tab1, tab2 = st.tabs(["Generar Pedido por Lote", "Vista General de Base"])

    with tab1:
        st.subheader("1. Ingreso de Órdenes (Copia y pega desde Excel)")
        input_ordenes = st.text_area(
            "Pega aquí los números de orden (uno por línea):",
            placeholder="Ejemplo:\n28712\n28711\n28710",
            height=250
        )

        if input_ordenes:
            # Procesar la lista ingresada
            lista_solicitada = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
            
            # Filtrar en la base de datos
            df_resultado = df_master[df_master['#Orden'].isin(lista_solicitada)].copy()

            if not df_resultado.empty:
                # Aplicar transformación de modelos
                df_resultado['CODIGO_PEDIDO'] = df_resultado['Producto'].apply(limpiar_modelo)
                
                # Seleccionar y ordenar columnas para el reporte
                columnas_reporte = ['#Orden', 'CODIGO_PEDIDO', 'Producto', 'Fecha', 'Repuestos', 'Técnico', 'Estado']
                df_final = df_resultado[columnas_reporte]

                st.success(f"✅ Se encontraron {len(df_final)} órdenes de las {len(lista_solicitada)} ingresadas.")
                
                # Mostrar vista previa
                st.dataframe(df_final, use_container_width=True)

                # --- LÓGICA DE DESCARGA EXCEL ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='Pedido_Generado')
                
                buffer.seek(0)

                st.download_button(
                    label="📥 Descargar Pedido en EXCEL (.xlsx)",
                    data=buffer,
                    file_name=f"pedido_tcl_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.error("❌ Ningún número de orden coincide con la base de datos.")

    with tab2:
        st.subheader("Explorador de Datos")
        st.write("Muestra las primeras 100 filas de la base de datos cargada.")
        st.dataframe(df_master.head(100), use_container_width=True)

except FileNotFoundError:
    st.error("⚠️ No se encontró el archivo CSV. Verifica que el nombre en GitHub sea exacto.")
except Exception as e:
    st.error(f"Ocurrió un error inesperado: {e}")
