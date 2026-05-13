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
        # Prefijo PL + primeros 5 caracteres del modelo
        return f"PL-{modelo_base[:5]}"
    return "OTROS"

# --- INTERFAZ DE USUARIO ---
st.title("📊 Generador de Pedidos - Formato Requerido")

# 1. CARGA DEL ARCHIVO EXCEL ORIGINAL
st.subheader("1. Cargar Base de Datos (.xlsx o .xls)")
uploaded_file = st.file_uploader("Sube el archivo de control de órdenes", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Leer el archivo Excel
        df_master = pd.read_excel(uploaded_file)
        
        # Limpiar columna de orden para asegurar coincidencia
        col_orden = '#Orden'
        if col_orden in df_master.columns:
            df_master[col_orden] = df_master[col_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        
        st.success("✅ Base de datos cargada.")

        # 2. INGRESO MANUAL DE LA LISTA
        st.divider()
        st.subheader("2. Ingrese Lista de Órdenes para el Pedido")
        input_ordenes = st.text_area(
            "Pegue aquí las órdenes (una por línea):",
            placeholder="28712\n28711\n28710",
            height=200
        )

        if input_ordenes:
            lista_solicitada = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
            
            # Filtrar órdenes solicitadas
            df_resultado = df_master[df_master[col_orden].isin(lista_solicitada)].copy()

            if not df_resultado.empty:
                # Generar el Código PL
                df_resultado['CODIGO'] = df_resultado['Producto'].apply(limpiar_modelo)
                
                # --- MAPEO Y RENOMBRADO SEGÚN EL ORDEN SOLICITADO ---
                # 1. ORDEN -> #Orden
                # 2. SERIE -> Serie (Si no existe, se crea vacía)
                # 3. MODELO -> Producto
                # 4. TALLER DE PROCEDENCIA -> Técnico
                # 5. REPUESTO -> Repuestos
                # 6. CODIGO -> El código PL generado
                
                if 'Serie' not in df_resultado.columns:
                    df_resultado['Serie'] = "N/A"
                
                # Renombrar para el reporte
                df_reporte = df_resultado.rename(columns={
                    col_orden: 'ORDEN',
                    'Serie': 'SERIE',
                    'Producto': 'MODELO',
                    'Técnico': 'TALLER DE PROCEDENCIA',
                    'Repuestos': 'REPUESTO',
                    'CODIGO': 'CODIGO_PL'
                })

                # Definir y aplicar el orden estricto de columnas
                columnas_finales = ['ORDEN', 'SERIE', 'MODELO', 'TALLER DE PROCEDENCIA', 'REPUESTO', 'CODIGO_PL']
                df_final = df_reporte[columnas_finales]

                st.subheader("3. Vista Previa del Nuevo Archivo")
                st.dataframe(df_final, use_container_width=True)

                # --- GENERACIÓN DEL EXCEL ---
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False, sheet_name='PEDIDO_FORMATEADO')
                
                buffer.seek(0)

                st.download_button(
                    label="📥 DESCARGAR EXCEL (ORDEN SOLICITADO)",
                    data=buffer,
                    file_name=f"Pedido_Repuestos_Formato_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.warning("⚠️ No se encontraron las órdenes en el archivo cargado.")

    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
else:
    st.info("Cargue el archivo Excel para activar el procesador.")
