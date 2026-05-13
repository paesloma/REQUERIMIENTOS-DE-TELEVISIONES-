import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Generador de Pedidos TCL", layout="wide")

def limpiar_modelo(nombre_producto):
    if pd.isna(nombre_producto):
        return "S/N"
    # Quitar TELEVISOR, extraer modelo y añadir PL-
    temp = re.sub(r'TELEVISOR|TELEVISION', '', str(nombre_producto), flags=re.IGNORECASE).strip()
    match = re.search(r'([A-Z0-9]+)', temp)
    if match:
        modelo_base = match.group(1)
        return f"PL-{modelo_base[:5]}"
    return "OTROS"

# --- INTERFAZ ---
st.title("📊 Generador de Pedidos Pro")

# 1. CARGA DE ARCHIVO
st.subheader("1. Cargar Base de Datos de Excel")
uploaded_file = st.file_uploader("Sube el archivo de control (.xlsx o .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Leer Excel y limpiar nombres de columnas de inmediato
        df_master = pd.read_excel(uploaded_file)
        df_master.columns = [str(c).strip() for c in df_master.columns]
        
        # BUSCADOR FLEXIBLE DE COLUMNAS (Mapeo)
        # Buscamos la columna de Orden, Serie, Producto, Técnico y Repuesto
        def encontrar_columna(lista_posibles):
            return next((c for c in lista_posibles if c in df_master.columns), None)

        col_orden = encontrar_columna(['#Orden', 'ORDEN', 'Orden', '# ORDEN', 'CODIGO'])
        col_serie = encontrar_columna(['Serie', 'SERIE', 'Serie/Chasis'])
        col_producto = encontrar_columna(['Producto', 'MODELO', 'PRODUCTO', 'Modelo'])
        col_tecnico = encontrar_columna(['Técnico', 'TECNICO', 'Taller', 'TALLER'])
        col_repuesto = encontrar_columna(['Repuestos', 'REPUESTOS', 'Repuesto', 'REPUESTO'])

        if col_orden and col_producto:
            st.success(f"✅ Archivo listo. Se detectó la columna de orden como: '{col_orden}'")
            
            # 2. ENTRADA DE DATOS
            st.divider()
            st.subheader("2. Selección de Órdenes")
            input_ordenes = st.text_area(
                "Pegue aquí los números de orden (uno por línea):",
                placeholder="28712\n28711",
                height=200
            )

            # --- BOTÓN PROCESAR ---
            if st.button("🚀 Procesar Pedido", type="primary"):
                if input_ordenes:
                    # Lista de entrada limpia
                    lista_busqueda = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                    
                    # Normalizar columna de orden en el DataFrame a texto limpio
                    df_master[col_orden] = df_master[col_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                    
                    # Filtrar coincidencias
                    df_res = df_master[df_master[col_orden].isin(lista_busqueda)].copy()

                    if not df_res.empty:
                        # Crear columnas faltantes con N/A para no romper el orden
                        if not col_serie: df_res['SERIE_TMP'] = "N/A"; col_serie = 'SERIE_TMP'
                        if not col_tecnico: df_res['TECNICO_TMP'] = "N/A"; col_tecnico = 'TECNICO_TMP'
                        if not col_repuesto: df_res['REPUESTO_TMP'] = "N/A"; col_repuesto = 'REPUESTO_TMP'
                        
                        # Generar el código PL
                        df_res['CODIGO_PL'] = df_res[col_producto].apply(limpiar_modelo)

                        # ARMAR EL EXCEL CON EL ORDEN SOLICITADO:
                        # 1.ORDEN 2.SERIE 3.MODELO 4.TALLER 5.REPUESTO 6.CODIGO
                        df_final = df_res.rename(columns={
                            col_orden: 'ORDEN',
                            col_serie: 'SERIE',
                            col_producto: 'MODELO',
                            col_tecnico: 'TALLER DE PROCEDENCIA',
                            col_repuesto: 'REPUESTO',
                            'CODIGO_PL': 'CODIGO'
                        })[['ORDEN', 'SERIE', 'MODELO', 'TALLER DE PROCEDENCIA', 'REPUESTO', 'CODIGO']]

                        st.subheader("3. Vista Previa del Pedido")
                        st.dataframe(df_final, use_container_width=True)

                        # GENERAR ARCHIVO PARA DESCARGA
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_final.to_excel(writer, index=False, sheet_name='Pedido')
                        
                        st.download_button(
                            label="📥 Descargar Excel de Pedido",
                            data=output.getvalue(),
                            file_name=f"Pedido_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("❌ No se encontró ninguna de esas órdenes en el Excel subido.")
                else:
                    st.warning("⚠️ Pegue los números de orden antes de procesar.")
        else:
            st.error("❌ El Excel no tiene una columna llamada '#Orden' o 'Producto'.")
            st.info(f"Columnas detectadas: {list(df_master.columns)}")

    except Exception as e:
        st.error(f"Error inesperado: {e}")

else:
    st.info("Sube el archivo Excel para comenzar.")
