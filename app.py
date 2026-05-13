import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Pedidos TCL", layout="wide")

def limpiar_modelo(nombre_producto):
    """Extrae PL + 5 caracteres del modelo técnico"""
    if pd.isna(nombre_producto): return "S/N"
    temp = re.sub(r'TELEVISOR|TELEVISION', '', str(nombre_producto), flags=re.IGNORECASE).strip()
    match = re.search(r'([A-Z0-9]+)', temp)
    if match:
        modelo_base = match.group(1)
        return f"PL-{modelo_base[:5]}"
    return "OTROS"

# --- INTERFAZ ---
st.title("📊 Generador de Pedidos - Selección Automática")

st.subheader("1. Cargar Base de Datos de Excel")
uploaded_file = st.file_uploader("Sube el archivo de control (.xlsx o .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Leer Excel y limpiar nombres de columnas
        df_master = pd.read_excel(uploaded_file)
        columnas_reales = [str(c).strip() for c in df_master.columns]
        df_master.columns = columnas_reales
        
        # LÓGICA DE SELECCIÓN AUTOMÁTICA (Mapeo por palabras clave)
        def detectar(keywords, default_idx=0):
            for i, col in enumerate(columnas_reales):
                if any(k.upper() in col.upper() for k in keywords):
                    return i
            return default_idx

        st.info("Verifique la selección automática de columnas:")
        
        c1, c2 = st.columns(2)
        with c1:
            # Busca columnas que contengan 'ORDEN', 'SERIE', 'PRODUCTO'
            sel_orden = st.selectbox("1. Columna de ORDEN:", columnas_reales, index=detectar(['#ORDEN', 'ORDEN']))
            sel_serie = st.selectbox("2. Columna de SERIE:", columnas_reales, index=detectar(['SERIE']))
            sel_modelo = st.selectbox("3. Columna de MODELO (Producto):", columnas_reales, index=detectar(['PRODUCTO', 'MODELO']))
        with c2:
            # Diferencia Taller de Procedencia y Taller (usando lógica de exclusión o posición)
            sel_procedencia = st.selectbox("4. Columna TALLER DE PROCEDENCIA:", columnas_reales, index=detectar(['PROCEDENCIA', 'TALLER']))
            sel_taller = st.selectbox("5. Columna TALLER (Técnico):", columnas_reales, index=detectar(['TECNICO', 'TALLER']))
            sel_repuesto = st.selectbox("6. Columna REPUESTO:", columnas_reales, index=detectar(['REPUESTO']))

        st.divider()
        st.subheader("2. Selección de Órdenes")
        input_ordenes = st.text_area("Pega aquí los números de orden (uno por línea):", height=200)

        # Botón Procesar
        if st.button("🚀 Procesar Pedido", type="primary"):
            if input_ordenes:
                lista_busqueda = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                
                # Normalizar columna de orden (quitar .0)
                df_master[sel_orden] = df_master[sel_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                
                # Filtrar órdenes
                df_res = df_master[df_master[sel_orden].isin(lista_busqueda)].copy()

                if not df_res.empty:
                    # Generar código PL (Columna 7)
                    df_res['CODIGO_GENERADO'] = df_res[sel_modelo].apply(limpiar_modelo)

                    # CONSTRUCCIÓN DEL EXCEL CON LAS 7 COLUMNAS EN EL ORDEN SOLICITADO
                    # 1.ORDEN, 2.SERIE, 3.MODELO, 4.TALLER DE PROCEDENCIA, 5.TALLER, 6.REPUESTO, 7.CODIGO
                    df_final = pd.DataFrame({
                        'ORDEN': df_res[sel_orden],
                        'SERIE': df_res[sel_serie],
                        'MODELO': df_res[sel_modelo],
                        'TALLER DE PROCEDENCIA': df_res[sel_procedencia],
                        'TALLER': df_res[sel_taller],
                        'REPUESTO': df_res[sel_repuesto],
                        'CODIGO': df_res['CODIGO_GENERADO']
                    })

                    st.subheader("3. Vista Previa del Pedido")
                    st.dataframe(df_final, use_container_width=True)

                    # Generar Excel para descarga
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='PEDIDO')
                    
                    st.download_button(
                        label="📥 Descargar Excel (7 Columnas)",
                        data=output.getvalue(),
                        file_name=f"Pedido_TCL_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ No se encontraron las órdenes en el archivo cargado.")
            else:
                st.warning("⚠️ Pegue los números de orden antes de procesar.")

    except Exception as e:
        st.error(f"Error al procesar: {e}")
else:
    st.info("Cargue el archivo Excel para que el sistema asigne las columnas automáticamente.")
