import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Pedidos Pro", layout="wide")

def extraer_codigo_final(nombre_producto):
    if pd.isna(nombre_producto): return "S/N"
    nombre_str = str(nombre_producto).upper()
    nombre_str = re.sub(r'\s+4K\s+', ' ', nombre_str)
    match_numero = re.search(r'\d', nombre_str)
    if match_numero:
        inicio = match_numero.start()
        bloque = nombre_str[inicio:inicio+5]
        return f"PL{bloque}"
    return "SIN_MODELO"

st.title("📊 Generador de Pedidos Profesional")

uploaded_file = st.file_uploader("Sube el archivo de Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df_master = pd.read_excel(uploaded_file)
        # Limpiar nombres de columnas
        df_master.columns = [str(c).strip() for c in df_master.columns]
        columnas_reales = list(df_master.columns)
        
        def detectar(keywords):
            for i, col in enumerate(columnas_reales):
                if any(k.upper() in col.upper() for k in keywords):
                    return i
            return 0

        st.info("Asigne las columnas para el reporte:")
        c1, c2 = st.columns(2)
        with c1:
            sel_orden = st.selectbox("1. ORDEN:", columnas_reales, index=detectar(['ORDEN', '#']))
            sel_serie = st.selectbox("2. SERIE:", columnas_reales, index=detectar(['SERIE']))
            sel_modelo = st.selectbox("3. MODELO:", columnas_reales, index=detectar(['PRODUCTO', 'MODELO']))
        with c2:
            sel_procedencia = st.selectbox("4. TALLER DE PROCEDENCIA:", columnas_reales, index=detectar(['PROCEDENCIA']))
            sel_taller = st.selectbox("5. TALLER:", columnas_reales, index=detectar(['TECNICO', 'TALLER']))
            sel_repuesto = st.selectbox("6. REPUESTO:", columnas_reales, index=detectar(['REPUESTO']))

        st.divider()
        input_ordenes = st.text_area("Pegue las órdenes aquí (una por línea):", height=150)

        if st.button("🚀 Procesar y Formatear Excel", type="primary"):
            if input_ordenes:
                lista_busqueda = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                # Normalizar columna de orden
                df_master[sel_orden] = df_master[sel_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_res = df_master[df_master[sel_orden].isin(lista_busqueda)].copy()

                if not df_res.empty:
                    df_res['CODIGO_PL'] = df_res[sel_modelo].apply(extraer_codigo_final)
                    
                    # Crear DataFrame final
                    df_final = pd.DataFrame({
                        'ORDEN': df_res[sel_orden],
                        'SERIE': df_res[sel_serie],
                        'MODELO': df_res[sel_modelo],
                        'TALLER DE PROCEDENCIA': df_res[sel_procedencia],
                        'TALLER': df_res[sel_taller],
                        'REPUESTO': df_res[sel_repuesto],
                        'CODIGO': df_res['CODIGO_PL']
                    })

                    # SOLUCIÓN AL ERROR: Rellenar vacíos con "REVISAR" para evitar errores de float
                    df_final = df_final.fillna("REVISAR")

                    st.dataframe(df_final, use_container_width=True)

                    # --- CREACIÓN DEL EXCEL ESTILIZADO ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='PEDIDO')
                        ws = writer.sheets['PEDIDO']
                        
                        # Estilos
                        fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                        font_header = Font(color="FFFFFF", bold=True)
                        align_center = Alignment(horizontal="center", vertical="center")
                        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                        top=Side(style='thin'), bottom=Side(style='thin'))

                        # Cabeceras y Ancho
                        for col_num, column in enumerate(df_final.columns, 1):
                            cell = ws.cell(row=1, column=col_num)
                            cell.fill = fill
                            cell.font = font_header
                            cell.alignment = align_center
                            cell.border = border
                            
                            # Ajustar ancho (convertimos a str para evitar el error de len())
                            max_len = max(df_final[column].astype(str).map(len).max(), len(column)) + 5
                            ws.column_dimensions[ws.cell(row=1, column=col_num).column_letter].width = max_len

                        # Bordes a los datos
                        for row in ws.iter_rows(min_row=2, max_row=len(df_final)+1, max_col=7):
                            for cell in row:
                                cell.border = border
                    
                    data_excel = output.getvalue()

                    st.download_button(
                        label="📥 Descargar Excel con Formato",
                        data=data_excel,
                        file_name=f"Pedido_TCL_{datetime.now().strftime('%d_%m_%y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No se encontraron coincidencias en la base de datos.")
    except Exception as e:
        st.error(f"Error técnico: {e}")
