import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Generador de Pedidos TCL", layout="wide")

def limpiar_modelo(nombre_producto):
    if pd.isna(nombre_producto): return "S/N"
    temp = re.sub(r'TELEVISOR|TELEVISION', '', str(nombre_producto), flags=re.IGNORECASE).strip()
    match = re.search(r'([A-Z0-9]+)', temp)
    if match:
        modelo_base = match.group(1)
        return f"PL-{modelo_base[:5]}"
    return "OTROS"

# --- INTERFAZ ---
st.title("📊 Generador de Pedidos - Configuración de Columnas")

st.subheader("1. Cargar Base de Datos de Excel")
uploaded_file = st.file_uploader("Sube el archivo de control (.xlsx o .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Leer Excel y limpiar nombres de columnas
        df_master = pd.read_excel(uploaded_file)
        columnas_reales = list(df_master.columns)
        
        st.info("Seleccione a qué columna corresponde cada dato solicitado:")
        
        # Selección manual de columnas para evitar el error de "KeyError"
        col1, col2, col3 = st.columns(3)
        with col1:
            sel_orden = st.selectbox("Columna de ORDEN:", columnas_reales, index=0)
            sel_serie = st.selectbox("Columna de SERIE:", columnas_reales, index=0)
        with col2:
            sel_modelo = st.selectbox("Columna de MODELO (Producto):", columnas_reales, index=0)
            sel_taller = st.selectbox("Columna de TALLER DE PROCEDENCIA:", columnas_reales, index=0)
        with col3:
            sel_repuesto = st.selectbox("Columna de REPUESTO:", columnas_reales, index=0)

        st.divider()
        st.subheader("2. Selección de Órdenes")
        input_ordenes = st.text_area("Pega aquí los números de orden (uno por línea):", height=150)

        if st.button("🚀 Procesar Pedido", type="primary"):
            if input_ordenes:
                lista_busqueda = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                
                # Normalizar columna de orden seleccionada
                df_master[sel_orden] = df_master[sel_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                
                # Filtrar
                df_res = df_master[df_master[sel_orden].isin(lista_busqueda)].copy()

                if not df_res.empty:
                    # Generar código PL basado en la columna de modelo seleccionada
                    df_res['CODIGO_PL'] = df_res[sel_modelo].apply(limpiar_modelo)

                    # Crear el DataFrame final con el orden exacto solicitado
                    # 1.ORDEN 2.SERIE 3.MODELO 4.TALLER 5.REPUESTO 6.CODIGO
                    df_final = pd.DataFrame({
                        'ORDEN': df_res[sel_orden],
                        'SERIE': df_res[sel_serie],
                        'MODELO': df_res[sel_modelo],
                        'TALLER DE PROCEDENCIA': df_res[sel_taller],
                        'REPUESTO': df_res[sel_repuesto],
                        'CODIGO': df_res['CODIGO_PL']
                    })

                    st.subheader("3. Vista Previa del Pedido")
                    st.dataframe(df_final, use_container_width=True)

                    # Generar Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Pedido')
                    
                    st.download_button(
                        label="📥 Descargar Pedido Formateado",
                        data=output.getvalue(),
                        file_name=f"Pedido_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ No se encontraron coincidencias para esas órdenes.")
            else:
                st.warning("⚠️ Pega los números de orden antes de procesar.")

    except Exception as e:
        st.error(f"Error al procesar: {e}")
else:
    st.info("Cargue el archivo Excel para configurar las columnas.")
