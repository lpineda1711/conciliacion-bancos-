import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.title("Conciliación Bancaria")

uploaded_file = st.file_uploader("Sube tu archivo Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)

    if "Bancos" in xls.sheet_names and "Mayor" in xls.sheet_names:
        # Leer hojas sin cabecera y asignar nombres exactos de tu plantilla
        df_bancos = pd.read_excel(xls, sheet_name="Bancos", header=None)
        df_bancos.columns = ['Bancos', 'Valor', 'Col3','Col4','Col5','Col6','Col7','Col8']  # ajusta según tu plantilla

        df_mayor = pd.read_excel(xls, sheet_name="Mayor", header=None)
        df_mayor.columns = ['Debe', 'Haber', 'Col3','Col4','Col5','Col6','Col7','Col8','Col9','Col10','Col11','Col12']  # ajusta según tu plantilla

        # Crear resumen de valores que no coinciden
        resumen = []
        for index, row in df_bancos.iterrows():
            tipo = str(row['Bancos']).strip().upper()
            valor = row['Valor']

            if tipo == 'C':  # Debe
                if df_mayor[df_mayor['Debe'] == valor].empty:
                    resumen.append({'Fila': index+2, 'Tipo': 'Debe', 'Valor': valor})
            elif tipo == 'D':  # Haber
                if df_mayor[df_mayor['Haber'] == valor].empty:
                    resumen.append({'Fila': index+2, 'Tipo': 'Haber', 'Valor': valor})

        # Guardar Excel con hojas originales y resumen
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_bancos.to_excel(writer, index=False, sheet_name='Bancos')
            df_mayor.to_excel(writer, index=False, sheet_name='Mayor')
            df_resumen = pd.DataFrame(resumen)
            df_resumen.to_excel(writer, index=False, sheet_name='Resumen')
        output.seek(0)

        # Sombrear valores del resumen
        wb = load_workbook(output)
        if 'Resumen' in wb.sheetnames:
            ws_resumen = wb['Resumen']
            yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
            for row in range(2, ws_resumen.max_row + 1):
                ws_resumen[f'C{row}'].fill = yellow_fill

        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        st.download_button(
            label="Descargar Excel Conciliado",
            data=final_output,
            file_name="Conciliacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error("El archivo debe contener las hojas 'Bancos' y 'Mayor'.")






