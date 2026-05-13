import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Pedidos TCL", layout="wide")

def limpiar_modelo_7(nombre_producto):
    """Extrae PL + 7 caracteres del modelo técnico"""
    if pd.isna(nombre_producto): return "S/N"
    # Eliminar TELEVISOR/TELEVISION
    temp = re.sub(r'TELEVISOR|TELEVISION', '', str(nombre_producto), flags=re.IGNORECASE).strip()
    # Buscar el bloque alfanumérico del modelo
    match = re.search(r'([A-Z0-9]+)', temp)
    if match:
        modelo_base = match.group(1)
        # Ajustado a 7 caracteres según tu solicitud
        return f"PL-{modelo_base[:7]}"
    return "OTROS"

# --- INTERFAZ ---
st.title("📊 Generador de Pedidos - Formato 7 Dígitos")

st.subheader("1. Cargar Base de Datos de Excel")
uploaded_file = st.file_uploader("Sube el archivo de control (.xlsx o .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Leer Excel y limpiar nombres de columnas de espacios
        df_master = pd.read_excel(uploaded_file)
        columnas_reales = [str(c).strip() for c in df_master.columns]
        df_master.columns = columnas_reales
        
        st.info("Seleccione las columnas correspondientes de su archivo:")
        
        # Selección manual para asegurar que no falte ninguna columna (incluyendo Taller)
        c1, c2, c3 = st.columns(3)
        with c1:
            sel_orden = st.selectbox("Columna de ORDEN:", columnas_reales)
            sel_serie = st.selectbox("Columna de SERIE:", columnas_reales)
        with c2:
            sel_modelo = st.selectbox("Columna de MODELO (Producto):", columnas_reales)
            sel_taller = st.selectbox("Columna de TALLER (Procedencia):", columnas_reales)
        with c3:
            sel_repuesto = st.selectbox("Columna de REPUESTO:", columnas_reales)

        st.divider()
        st.subheader("2. Selección de Órdenes")
        input_ordenes = st.text_area("Pega aquí los números de orden (uno por línea):", height=150)

        # Botón Procesar
        if st.button("🚀 Procesar Pedido", type="primary"):
            if input_ordenes:
                lista_busqueda = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                
                # Normalizar columna de orden para evitar errores con .0
                df_master[sel_orden] = df_master[sel_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                
                # Filtrar órdenes
                df_res = df_master[df_master[sel_orden].isin(lista_busqueda)].copy()

                if not df_res.empty:
                    # Generar código PL con 7 caracteres
                    df_res['CODIGO_7'] = df_res[sel_modelo].apply(limpiar_modelo_7)

                    # Crear DataFrame final con el orden exacto solicitado:
                    # 1.ORDEN, 2.SERIE, 3.MODELO, 4.TALLER, 5.REPUESTO, 6.CODIGO
                    df_final = pd.DataFrame({
                        'ORDEN': df_res[sel_orden],
                        'SERIE': df_res[sel_serie],
                        'MODELO': df_res[sel_modelo],
                        'TALLER DE PROCEDENCIA': df_res[sel_taller],
                        'REPUESTO': df_res[sel_repuesto],
                        'CODIGO': df_res['CODIGO_7']
                    })

                    st.subheader("3. Vista Previa (Código de 7 dígitos)")
                    st.dataframe(df_final, use_container_width=True)

                    # Generar Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Pedido')
                    
                    st.download_button(
                        label="📥 Descargar Excel (7 Dígitos)",
                        data=output.getvalue(),
                        file_name=f"Pedido_7D_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ No se encontraron las órdenes en el archivo cargado.")
            else:
                st.warning("⚠️ Pega los números de orden antes de procesar.")

    except Exception as e:
        st.error(f"Error al procesar: {e}")
else:
    st.info("Por favor, cargue el archivo Excel para configurar las columnas.")
