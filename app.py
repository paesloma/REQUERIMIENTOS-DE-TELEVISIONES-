import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Generador de Pedidos Pro", layout="wide")

def limpiar_modelo(nombre_producto):
    if pd.isna(nombre_producto): return "S/N"
    temp = re.sub(r'TELEVISOR|TELEVISION', '', str(nombre_producto), flags=re.IGNORECASE).strip()
    match = re.search(r'([A-Z0-9]+)', temp)
    if match:
        modelo_base = match.group(1)
        return f"PL-{modelo_base[:5]}"
    return "OTROS"

# --- INTERFAZ ---
st.title("📊 Generador de Pedidos Automático")

st.subheader("1. Cargar Base de Datos de Excel")
uploaded_file = st.file_uploader("Sube el archivo de control (.xlsx o .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Leer Excel y limpiar nombres de columnas
        df_master = pd.read_excel(uploaded_file)
        df_master.columns = [str(c).strip() for c in df_master.columns]
        
        # LÓGICA DE DETECCIÓN AUTOMÁTICA DE COLUMNAS (Por contenido)
        col_orden = None
        col_producto = None
        col_serie = None
        col_tecnico = None
        col_repuesto = None

        # Buscador por nombres comunes
        for c in df_master.columns:
            c_upper = c.upper()
            if any(x in c_upper for x in ['ORDEN', 'CODIGO', '#']): col_orden = c
            if any(x in c_upper for x in ['PRODUCTO', 'MODELO']): col_producto = c
            if any(x in c_upper for x in ['SERIE', 'CHASIS']): col_serie = c
            if any(x in c_upper for x in ['TECNICO', 'TALLER']): col_tecnico = c
            if any(x in c_upper for x in ['REPUESTO', 'COMPONENT']): col_repuesto = c

        if col_orden and col_producto:
            st.success(f"✅ Columnas detectadas: Orden -> '{col_orden}' | Modelo -> '{col_producto}'")
            
            st.divider()
            st.subheader("2. Selección de Órdenes")
            input_ordenes = st.text_area("Pega aquí los números de orden (uno por línea):", height=150)

            if st.button("🚀 Procesar Pedido", type="primary"):
                if input_ordenes:
                    lista_busqueda = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                    
                    # Normalizar columna de orden a texto limpio
                    df_master[col_orden] = df_master[col_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                    
                    # Filtrar
                    df_res = df_master[df_master[col_orden].isin(lista_busqueda)].copy()

                    if not df_res.empty:
                        # Rellenar faltantes
                        if not col_serie: df_res['SERIE'] = "N/A"; col_serie = 'SERIE'
                        if not col_tecnico: df_res['TALLER'] = "N/A"; col_tecnico = 'TALLER'
                        if not col_repuesto: df_res['REPUESTO'] = "N/A"; col_repuesto = 'REPUESTO'
                        
                        # Generar código PL
                        df_res['CODIGO_PL'] = df_res[col_producto].apply(limpiar_modelo)

                        # Reordenar según pedido: 1.ORDEN 2.SERIE 3.MODELO 4.TALLER 5.REPUESTO 6.CODIGO
                        df_final = df_res.rename(columns={
                            col_orden: 'ORDEN',
                            col_serie: 'SERIE',
                            col_producto: 'MODELO',
                            col_tecnico: 'TALLER DE PROCEDENCIA',
                            col_repuesto: 'REPUESTO',
                            'CODIGO_PL': 'CODIGO'
                        })[['ORDEN', 'SERIE', 'MODELO', 'TALLER DE PROCEDENCIA', 'REPUESTO', 'CODIGO']]

                        st.subheader("3. Vista Previa")
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
                        st.error("❌ No se encontró ninguna de esas órdenes en el archivo.")
                else:
                    st.warning("⚠️ Pega los números antes de procesar.")
        else:
            st.error("❌ No se pudo identificar la columna de 'Orden' o 'Producto'.")
            st.info(f"Columnas encontradas: {list(df_master.columns)}")

    except Exception as e:
        st.error(f"Error: {e}")
