import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Generador de Pedidos Pro", layout="wide")

def extraer_codigo_final(nombre_producto):
    """
    1. Excluye '4K' si está entre espacios.
    2. Busca el primer dígito numérico.
    3. Extrae 5 caracteres y devuelve el código SIN guion (PLXXXXX).
    """
    if pd.isna(nombre_producto): return "S/N"
    
    nombre_str = str(nombre_producto).upper()
    
    # Exclusión de '4K' con separadores de espacio
    nombre_str = re.sub(r'\s+4K\s+', ' ', nombre_str)
    
    # Buscar el primer carácter numérico
    match_numero = re.search(r'\d', nombre_str)
    
    if match_numero:
        inicio = match_numero.start()
        # Extraer 5 caracteres a partir del primer número
        bloque = nombre_str[inicio:inicio+5]
        # Retorna el código completo sin guiones
        return f"PL{bloque}"
    
    return "SIN_MODELO"

st.title("📊 Generador de Pedidos - Formato MOTSUR (Sin Guiones)")

uploaded_file = st.file_uploader("Sube el archivo de Excel de control", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df_master = pd.read_excel(uploaded_file)
        # Limpiar espacios en los nombres de las columnas
        df_master.columns = [str(c).strip() for c in df_master.columns]
        columnas_reales = list(df_master.columns)
        
        def detectar(keywords):
            for i, col in enumerate(columnas_reales):
                if any(k.upper() in col.upper() for k in keywords):
                    return i
            return 0

        st.info("Verifique la detección de columnas:")
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

        if st.button("🚀 Procesar Pedido Profesional", type="primary"):
            if input_ordenes:
                lista_busqueda = [o.strip() for o in input_ordenes.split('\n') if o.strip()]
                # Asegurar que la columna de orden sea texto y no tenga decimales .0
                df_master[sel_orden] = df_master[sel_orden].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                df_res = df_master[df_master[sel_orden].isin(lista_busqueda)].copy()

                if not df_res.empty:
                    df_res['CODIGO_PL'] = df_res[sel_modelo].apply(extraer_codigo_final)
                    
                    df_final = pd.DataFrame({
                        'ORDEN': df_res[sel_orden],
                        'SERIE': df_res[sel_serie],
                        'MODELO': df_res[sel_modelo],
                        'TALLER DE PROCEDENCIA': df_res[sel_procedencia],
                        'TALLER': df_res[sel_taller],
                        'REPUESTO': df_res[sel_repuesto],
                        'CODIGO': df_res['CODIGO_PL']
                    })

                    # Rellenar celdas vacías con "REVISAR"
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

                        # Formato y Auto-Ajuste
                        for col_num, column in enumerate(df_final.columns, 1):
                            cell = ws.cell(row=1, column=col_num)
                            cell.fill = fill
                            cell.font = font_header
                            cell.alignment = align_center
                            cell.border = border
                            
                            max_len = max(df_final[column].astype(str).map(len).max(), len(column)) + 5
                            ws.column_dimensions[ws.cell(row=1, column=col_num).column_letter].width = max_len

                        for row in ws.iter_rows(min_row=2, max_row=len(df_final)+1, max_col=7):
                            for cell in row:
                                cell.border = border
                    
                    data_excel = output.getvalue()
                    
                    # NOMBRE DE ARCHIVO SIN GUIONES (Espacios en su lugar)
                    fecha_hoy = datetime.now().strftime('%d %m %Y')
                    nombre_archivo = f"PEDIDO BDG 64 MOTSUR TVS {fecha_hoy}.xlsx"

                    st.download_button(
                        label="📥 Descargar Excel MOTSUR TVS",
                        data=data_excel,
                        file_name=nombre_archivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("No se encontraron coincidencias.")
    except Exception as e:
        st.error(f"Error técnico: {e}")
