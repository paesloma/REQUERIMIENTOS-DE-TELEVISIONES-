import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Pedidos TCL", layout="wide")

def extraer_codigo_desde_numero(nombre_producto):
    """
    Busca el primer dígito numérico en la cadena.
    Extrae ese número y los 4 caracteres siguientes (total 5).
    """
    if pd.isna(nombre_producto): return "S/N"
    
    nombre_str = str(nombre_producto).upper()
    
    # Expresión regular que busca el primer dígito (\d)
    match_numero = re.search(r'\d', nombre_str)
    
    if match_numero:
        # Posición donde empieza el primer número
        inicio = match_numero.start()
        # Extraer 5 caracteres exactos desde esa posición
        bloque = nombre_str[inicio:inicio+5]
        return f"PL-{bloque}"
    
    return "SIN_MODELO"

# --- INTERFAZ DE USUARIO ---
st.title("📊 Generador de Pedidos - Regla de Carácter Numérico")

st.subheader("1. Cargar Base de Datos de Excel")
uploaded_file = st.file_uploader("Sube el archivo de control (.xlsx o .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Leer Excel y limpiar nombres de columnas
        df_master = pd.read_excel(uploaded_file)
        columnas_reales = [str(c).strip() for c in df_master.columns]
        df_master.columns = columnas_reales
        
        # LÓGICA DE DETECCIÓN AUTOMÁTICA DE COLUMNAS
        def detectar(keywords, default_idx=0):
            for i, col in enumerate(columnas_reales):
                if any(k.upper() in col.upper() for k in keywords):
                    return i
            return default_idx

        st.info("Confirme que las columnas se han asignado correctamente:")
        
        c1, c2 = st.columns(2)
        with c1:
            sel_orden = st.selectbox("1. Columna de ORDEN:", columnas_reales, index=detectar(['#ORDEN', 'ORDEN']))
            sel_serie = st.selectbox("2. Columna de SERIE:", columnas_reales, index=detectar(['SERIE']))
            sel_modelo = st.selectbox("3. Columna de MODELO (Producto):", columnas_reales, index=detectar(['PRODUCTO', 'MODELO']))
        with c2:
            sel_procedencia = st.selectbox("4. Columna TALLER DE PROCEDENCIA:", columnas_reales, index=detectar(['PROCEDENCIA']))
            sel_taller = st.selectbox("5. Columna TALLER (Técnico):", columnas_reales, index=detectar(['TECNICO', 'TALLER']))
            sel_repuesto = st.selectbox("6. Columna REPUESTO:", columnas_reales, index=detectar(['REPUESTO']))

        st.divider()
        st.subheader("2. Selección de Órdenes")
        input_ordenes = st.text_area("Pega aquí los números de orden (uno por línea):", height=200)

        # BOTÓN PROCESAR
        if st.button("🚀 Procesar Pedido", type="primary"):
            if input_ordenes:
                lista_busqueda = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                
                # Normalización de la columna de orden (eliminar .0)
                df_master[sel_orden] = df_master[sel_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                
                # Filtrar órdenes solicitadas
                df_res = df_master[df_master[sel_orden].isin(lista_busqueda)].copy()

                if not df_res.empty:
                    # Aplicar la regla: empezar a contar desde el primer número encontrado
                    df_res['CODIGO_GENERADO'] = df_res[sel_modelo].apply(extraer_codigo_desde_numero)

                    # CONSTRUCCIÓN DE LAS 7 COLUMNAS EN EL ORDEN ESTRICTO
                    # 1. ORDEN, 2. SERIE, 3. MODELO, 4. TALLER DE PROCEDENCIA, 5. TALLER, 6. REPUESTO, 7. CODIGO
                    df_final = pd.DataFrame({
                        'ORDEN': df_res[sel_orden],
                        'SERIE': df_res[sel_serie],
                        'MODELO': df_res[sel_modelo],
                        'TALLER DE PROCEDENCIA': df_res[sel_procedencia],
                        'TALLER': df_res[sel_taller],
                        'REPUESTO': df_res[sel_repuesto],
                        'CODIGO': df_res['CODIGO_GENERADO']
                    })

                    st.subheader("3. Vista Previa del Reporte")
                    st.dataframe(df_final, use_container_width=True)

                    # Exportación a Excel
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
    st.info("Suba el archivo de Excel para detectar las columnas automáticamente.")
