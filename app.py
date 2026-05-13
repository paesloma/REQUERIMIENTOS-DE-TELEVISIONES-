import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Pedidos TCL", layout="wide")

def extraer_codigo_final(nombre_producto):
    """Limpia '4K', busca el primer número y extrae 5 caracteres (Formato PLXXXXX)"""
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

st.subheader("1. Cargar Base de Datos de Excel")
uploaded_file = st.file_uploader("Sube el archivo de control (.xlsx o .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df_master = pd.read_excel(uploaded_file)
        columnas_reales = [str(c).strip() for c in df_master.columns]
        df_master.columns = columnas_reales
        
        def detectar(keywords):
            for i, col in enumerate(columnas_reales):
                if any(k.upper() in col.upper() for k in keywords):
                    return i
            return 0

        st.info("Selección automática de columnas completada:")
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
        st.subheader("2. Selección de Órdenes")
        input_ordenes = st.text_area("Pega aquí los números de orden:", height=150)

        if st.button("🚀 Generar Pedido con Formato Profesional", type="primary"):
            if input_ordenes:
                lista_busqueda = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                df_master[sel_orden] = df_master[sel_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_res = df_master[df_master[sel_orden].isin(lista_busqueda)].copy()

                if not df_res.empty:
                    df_res['CODIGO_PL'] = df_res[sel_modelo].apply(extraer_codigo_final)
                    
                    # Estructura final de 7 columnas
                    df_final = pd.DataFrame({
                        'ORDEN': df_res[sel_orden],
                        'SERIE': df_res[sel_serie],
                        'MODELO': df_res[sel_modelo],
                        'TALLER DE PROCEDENCIA': df_res[sel_procedencia],
                        'TALLER': df_res[sel_taller],
                        'REPUESTO': df_res[sel_repuesto],
                        'CODIGO': df_res['CODIGO_PL']
                    })

                    st.subheader("3. Vista Previa")
                    st.dataframe(df_final, use_container_width=True)

                    # --- EXPORTACIÓN CON FORMATO ESTÉTICO ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='PEDIDO')
                        
                        ws = writer.sheets['PEDIDO']
                        
                        # Definir estilos
                        blue_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                        white_font = Font(color="FFFFFF", bold=True)
                        alignment = Alignment(horizontal="center", vertical="center")
                        border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                        top=Side(style='thin'), bottom=Side(style='thin'))

                        # Aplicar estilos a cabeceras y auto-ajustar ancho
                        for col_num, column in enumerate(df_final.columns, 1):
                            cell = ws.cell(row=1, column=col_num)
                            cell.fill = blue_fill
                            cell.font = white_font
                            cell.alignment = alignment
                            cell.border = border
                            
                            # Lógica de auto-ajuste de columna
                            max_len = max(df_final[column].astype(str).map(len).max(), len(column)) + 4
                            ws.column_dimensions[ws.cell(row=1, column=col_num).column_letter].width = max_len

                        # Aplicar bordes a los datos
                        for row in ws.iter_rows(min_row=2, max_row=len(df_final)+1, max_col=7):
                            for cell in row:
                                cell.border = border

                    st.download_button(
                        label="📥 Descargar Excel Estilizado",
                        data=output.getvalue(),
                        file_name=f"Pedido_TCL_Pro_{datetime.now().strftime('%H%M_%d%m%y')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ No se encontraron las órdenes.")
    except Exception as e:
        st.error(f"Error técnico: {e}")
