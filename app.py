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
    # Eliminar "TELEVISOR" o "TELEVISION"
    temp = re.sub(r'TELEVISOR|TELEVISION', '', str(nombre_producto), flags=re.IGNORECASE).strip()
    # Buscar bloque alfanumérico (ej: 65P755)
    match = re.search(r'([A-Z0-9]+)', temp)
    if match:
        modelo_base = match.group(1)
        return f"PL-{modelo_base[:5]}"
    return "OTROS"

# --- INTERFAZ DE USUARIO ---
st.title("📊 Generador de Pedidos - Procesamiento por Botón")

# 1. CARGA DEL ARCHIVO EXCEL
st.subheader("1. Cargar Base de Datos (.xlsx o .xls)")
uploaded_file = st.file_uploader("Sube el archivo de control de órdenes", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Leer el archivo y limpiar nombres de columnas de espacios laterales
        df_master = pd.read_excel(uploaded_file)
        df_master.columns = [str(c).strip() for c in df_master.columns]
        
        # BUSCAR COLUMNA DE ORDEN (Flexible a variaciones)
        posibles_nombres = ['#Orden', 'ORDEN', '# ORDEN', 'Nro Orden', 'Orden', 'CODIGO']
        col_orden = next((c for c in posibles_nombres if c in df_master.columns), None)

        if col_orden:
            st.success(f"✅ Archivo cargado. Columna identificada: '{col_orden}'")
            
            st.divider()
            
            # 2. INGRESO MANUAL Y BOTÓN DE PROCESAR
            st.subheader("2. Selección de Órdenes")
            input_ordenes = st.text_area(
                "Pegue aquí los números de orden (uno por línea):",
                placeholder="28712\n28711",
                height=200
            )

            # Botón para ejecutar la búsqueda
            if st.button("🚀 Procesar Pedido", type="secondary"):
                if input_ordenes:
                    # Limpiar lista de entrada
                    lista_solicitada = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                    
                    # Convertir columna de la base a texto para comparar
                    df_master[col_orden] = df_master[col_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                    
                    # Filtrar
                    df_resultado = df_master[df_master[col_orden].isin(lista_solicitada)].copy()

                    if not df_resultado.empty:
                        # Asegurar existencia de columnas para evitar errores de visualización
                        cols_necesarias = {'Serie': 'SERIE', 'Producto': 'MODELO', 'Técnico': 'TALLER DE PROCEDENCIA', 'Repuestos': 'REPUESTO'}
                        for original, nuevo in cols_necesarias.items():
                            if original not in df_resultado.columns:
                                df_resultado[original] = "N/A"
                        
                        # Generar el Código PL
                        df_resultado['CODIGO_PL'] = df_resultado['Producto'].apply(limpiar_modelo)
                        
                        # Reordenar y renombrar según tu requerimiento exacto
                        df_final = df_resultado.rename(columns={
                            col_orden: 'ORDEN',
                            'Serie': 'SERIE',
                            'Producto': 'MODELO',
                            'Técnico': 'TALLER DE PROCEDENCIA',
                            'Repuestos': 'REPUESTO'
                        })[['ORDEN', 'SERIE', 'MODELO', 'TALLER DE PROCEDENCIA', 'REPUESTO', 'CODIGO_PL']]

                        st.subheader("3. Resultado del Procesamiento")
                        st.dataframe(df_final, use_container_width=True)

                        # GENERACIÓN DE EXCEL PARA DESCARGAR
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                            df_final.to_excel(writer, index=False, sheet_name='PEDIDO')
                        buffer.seek(0)

                        st.download_button(
                            label="📥 DESCARGAR EXCEL",
                            data=buffer,
                            file_name=f"Pedido_TCL_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("❌ No se encontraron coincidencias. Revise si los números de orden son correctos.")
                else:
                    st.warning("⚠️ Por favor, pegue al menos un número de orden antes de procesar.")
        else:
            st.error(f"❌ No se encontró la columna de orden. Columnas en su archivo: {list(df_master.columns)}")

    except Exception as e:
        st.error(f"Error técnico: {e}")
else:
    st.info("👋 Bienvenida/o. Por favor, cargue su archivo de Excel para iniciar.")
