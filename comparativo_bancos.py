import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO

# Título
st.title("Conciliación Bancaria")

# Subir archivo Excel
uploaded_file = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if uploaded_file:
    # Cargar Excel
    xls = pd.ExcelFile(uploaded_file)
    
    # Revisar si existen las hojas
    if "Bancos" in xls.sheet_names and "Mayor" in xls.sheet_names:
        df_bancos = pd.read_excel(xls, sheet_name="Bancos")
        df_mayor = pd.read_excel(xls, sheet_name="Mayor")
        
        # Normalizar nombres de columnas: eliminar espacios y uniformizar mayúsculas
        df_bancos.columns = df_bancos.columns.str.strip().str.capitalize()
        df_mayor.columns = df_mayor.columns.str.strip().str.capitalize()
        
        # Crear lista para el resumen
        resumen = []
        
        for index, row in df_bancos.iterrows():
            tipo = str(row['Tipo']).strip().upper()
            valor = row['Valor']
            
            if tipo == 'C':  # Debe
                if df_mayor[df_mayor['Debe'] == valor].empty:
                    resumen.append({'Fila': index+2, 'Tipo': 'Debe', 'Valor': valor})
            elif tipo == 'D':  # Haber
                if df_mayor[df_mayor['Haber'] == valor].empty:
                    resumen.append({'Fila': index+2, 'Tipo': 'Haber', 'Valor': valor})
        
        # Guardar hojas originales en un Excel temporal
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_bancos.to_excel(writer, index=False, sheet_name='Bancos')
            df_mayor.to_excel(writer, index=False, sheet_name='Mayor')
            df_resumen = pd.DataFrame(resumen)
            df_resumen.to_excel(writer, index=False, sheet_name='Resumen')
            writer.save()
        output.seek(0)
        
        # Abrir con openpyxl para sombrear los valores
        wb = load_workbook(output)
        if 'Resumen' in wb.sheetnames:
            ws_resumen = wb['Resumen']
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            
            for row in range(2, ws_resumen.max_row + 1):
                ws_resumen[f'C{row}'].fill = yellow_fill  # Columna Valor
        
        # Guardar cambios finales
        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)
        
        # Botón de descarga
        st.download_button(
            label="Descargar Excel Conciliado",
            data=final_output,
            file_name="Conciliacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("El archivo debe contener las hojas 'Bancos' y 'Mayor'.")

