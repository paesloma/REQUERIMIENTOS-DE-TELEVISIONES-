import streamlit as st
import pandas as pd
import re
from datetime import datetime

# --- Configuración y Limpieza ---
st.set_page_config(page_title="Generador de Pedidos", layout="wide")

def limpiar_modelo(nombre_producto):
    if pd.isna(nombre_producto):
        return "S/N"
    # Quitar "TELEVISOR", extraer el código y poner prefijo PL
    temp = re.sub(r'TELEVISOR|TELEVISION', '', str(nombre_producto), flags=re.IGNORECASE).strip()
    match = re.search(r'([A-Z0-9]+)', temp)
    if match:
        modelo_base = match.group(1)
        return f"PL-{modelo_base[:5]}"
    return "OTROS"

@st.cache_data
def load_base_datos(file_path):
    df = pd.read_csv(file_path)
    # Estandarizar nombre de columna de orden (ajustar según tu CSV real)
    if '#Orden' in df.columns:
        df['#Orden'] = df['#Orden'].astype(str).str.strip()
    return df

# --- Interfaz de Usuario ---
st.title("📦 Generador de Pedidos por Lote")
st.markdown("""
### Instrucciones:
1. Pega o escribe los números de orden en el cuadro de abajo (uno por línea).
2. El sistema buscará la información y generará el formato **PL-XXXXX**.
3. Descarga el archivo final con el botón.
""")

try:
    # Cargar la base de datos principal
    base_path = "CONTROL ORDENES 2026 TVS TCL.xlsx - ORDENES 2026.csv"
    df_master = load_base_datos(base_path)

    # --- CUADRO DE ENTRADA TIPO EXCEL (MANUAL) ---
    st.subheader("1. Ingreso Manual de Órdenes")
    input_ordenes = st.text_area(
        "Pega aquí los números de orden (una por fila):",
        placeholder="Ejemplo:\n28712\n28711\n28710",
        height=200
    )

    if input_ordenes:
        # Convertir el texto ingresado en una lista limpia
        lista_ordenes = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
        
        # Filtrar la base de datos solo por las órdenes solicitadas
        df_filtrado = df_master[df_master['#Orden'].isin(lista_ordenes)].copy()

        if not df_filtrado.empty:
            # Aplicar la regla del código PL a los resultados encontrados
            df_filtrado['CODIGO_PEDIDO'] = df_filtrado['Producto'].apply(limpiar_modelo)
            
            # Reordenar columnas para que el código PL sea visible al inicio
            columnas_finales = ['#Orden', 'CODIGO_PEDIDO', 'Producto', 'Fecha', 'Repuestos', 'Estado']
            df_mostrar = df_filtrado[columnas_finales]

            st.success(f"✅ Se encontraron {len(df_filtrado)} órdenes de las {len(lista_ordenes)} ingresadas.")
            
            # --- VISTA PREVIA Y DESCARGA ---
            st.subheader("2. Vista Previa del Pedido")
            st.dataframe(df_mostrar, use_container_width=True)

            csv_pedido = df_mostrar.to_csv(index=False).encode('utf-8')
            
            st.download_button(
                label="📥 Descargar Lista de Pedido (CSV)",
                data=csv_pedido,
                file_name=f"pedido_generado_{datetime.now().strftime('%H%M%S')}.csv",
                mime="text/csv",
                type="primary"
            )
        else:
            st.error("❌ No se encontró ninguna de las órdenes ingresadas en la base de datos.")

except Exception as e:
    st.error(f"Hubo un error al procesar el archivo: {e}")
