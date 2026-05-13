import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Pedidos TCL", layout="wide")

def extraer_codigo_final(nombre_producto):
    """
    1. Limpia la palabra '4K' si está entre espacios.
    2. Busca el primer dígito numérico y empieza a contar desde ahí.
    3. Extrae 5 caracteres y añade prefijo PL-.
    """
    if pd.isna(nombre_producto): return "S/N"
    
    nombre_str = str(nombre_producto).upper()
    
    # Exclusión de '4K' con dos separadores de espacio (palabra independiente)
    nombre_str = re.sub(r'\s+4K\s+', ' ', nombre_str)
    
    # Buscar el primer carácter numérico
    match_numero = re.search(r'\d', nombre_str)
    
    if match_numero:
        inicio = match_numero.start()
        # Extraer 5 caracteres a partir del primer número encontrado
        bloque = nombre_str[inicio:inicio+5]
        return f"PL-{bloque}"
    
    return "SIN_MODELO"

# --- INTERFAZ DE USUARIO ---
st.title("📊 Generador de Pedidos - Lógica de Extracción Avanzada")

st.subheader("1. Cargar Base de Datos de Excel")
uploaded_file = st.file_uploader("Sube el archivo de control (.xlsx o .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df_master = pd.read_excel(uploaded_file)
        columnas_reales = [str(c).strip() for c in df_master.columns]
        df_master.columns = columnas_reales
        
        # DETECCIÓN AUTOMÁTICA DE COLUMNAS
        def detectar(keywords):
            for i, col in enumerate(columnas_reales):
                if any(k.upper() in col.upper() for k in keywords):
                    return i
            return 0

        st.info("Verifique la asignación automática de columnas:")
        
        c1, c2 = st.columns(2)
        with c1:
            sel_orden = st.selectbox("1. Columna de ORDEN:", columnas_reales, index=detectar(['ORDEN', '#']))
            sel_serie = st.selectbox("2. Columna de SERIE:", columnas_reales, index=detectar(['SERIE']))
            sel_modelo = st.selectbox("3. Columna de MODELO (Producto):", columnas_reales, index=detectar(['PRODUCTO', 'MODELO']))
        with c2:
            sel_procedencia = st.selectbox("4. Columna TALLER DE PROCEDENCIA:", columnas_reales, index=detectar(['PROCEDENCIA']))
            sel_taller = st.selectbox("5. Columna TALLER:", columnas_reales, index=detectar(['TALLER', 'TECNICO']))
            sel_repuesto = st.selectbox("6. Columna REPUESTO:", columnas_reales, index=detectar(['REPUESTO']))

        st.divider()
        st.subheader("2. Selección de Órdenes")
        input_ordenes = st.text_area("Pega aquí los números de orden (uno por línea):", height=200)

        # BOTÓN PROCESAR
        if st.button("🚀 Procesar Pedido", type="primary"):
            if input_ordenes:
                lista_busqueda = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                
                # Normalizar columna de orden
                df_master[sel_orden] = df_master[sel_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                
                # Filtrar órdenes
                df_res = df_master[df_master[sel_orden].isin(lista_busqueda)].copy()

                if not df_res.empty:
                    # Generar código PL con las nuevas reglas
                    df_res['CODIGO_GENERADO'] = df_res[sel_modelo].apply(extraer_codigo_final)

                    # CONSTRUCCIÓN DE LAS 7 COLUMNAS EN EL ORDEN SOLICITADO
                    # ORDEN, SERIE, MODELO, TALLER DE PROCEDENCIA, TALLER, REPUESTO, CODIGO
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

                    # Exportación a Excel
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='PEDIDO')
                    
                    st.download_button(
                        label="📥 Descargar Excel (7 Columnas)",
                        data=output.getvalue(),
                        file_name=f"Pedido_TCL_Limpio_{datetime.now().strftime('%d_%m_%y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ No se encontraron las órdenes indicadas en el archivo.")
            else:
                st.warning("⚠️ Ingrese al menos un número de orden.")

    except Exception as e:
        st.error(f"Ocurrió un error: {e}")
else:
    st.info("Cargue el archivo Excel para iniciar el procesamiento.")
