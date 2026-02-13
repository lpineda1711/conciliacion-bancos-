import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.title("Conciliación Bancaria")

# Subir archivo
archivo = st.file_uploader("Sube tu archivo Excel", type=["xlsx"])

if archivo is not None:
    # Cargar hojas originales
    bancos = pd.read_excel(archivo, sheet_name="Bancos")
    mayor = pd.read_excel(archivo, sheet_name="Mayor")

    # Normalizar valores
    bancos["Valor"] = bancos["Valor"].round(2)
    mayor["Debe"] = mayor["Debe"].fillna(0).round(2)
    mayor["Haber"] = mayor["Haber"].fillna(0).round(2)

    # Crear columnas esperadas según tipo
    def obtener_valor_esperado(row):
        if row["Tipo"] == "C":
            return row["Valor"], 0  # Debe
        elif row["Tipo"] == "D":
            return 0, row["Valor"]  # Haber
        else:
            return 0, 0

    bancos[["Debe_esperado", "Haber_esperado"]] = bancos.apply(
        lambda row: pd.Series(obtener_valor_esperado(row)), axis=1
    )

    # Crear resumen de conciliación
    resumen = []
    for index, row in bancos.iterrows():
        encontrado = mayor[
            (mayor["Debe"] == row["Debe_esperado"]) &
            (mayor["Haber"] == row["Haber_esperado"])
        ]
        estado = "OK" if not encontrado.empty else "NO COINCIDE"
        resumen.append({
            "Fecha": row["Fecha"],
            "Descripción": row["Descripción"],
            "Tipo": row["Tipo"],
            "Valor": row["Valor"],
            "Debe esperado": row["Debe_esperado"],
            "Haber esperado": row["Haber_esperado"],
            "Estado": estado
        })

    df_resumen = pd.DataFrame(resumen)

    # Guardar todo en un archivo Excel en memoria
    salida = BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        bancos.to_excel(writer, sheet_name="Bancos", index=False)
        mayor.to_excel(writer, sheet_name="Mayor", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

    # Abrir archivo para pintar amarillo donde no coincide
    salida.seek(0)
    wb = load_workbook(salida)
    ws = wb["Resumen"]

    amarillo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        if row[6].value == "NO COINCIDE":  # Columna Estado
            for cell in row:
                cell.fill = amarillo

    # Guardar nuevamente en memoria
    salida_final = BytesIO()
    wb.save(salida_final)
    salida_final.seek(0)

    # Botón de descarga
    st.download_button(
        label="Descargar Excel Conciliado",
        data=salida_final,
        file_name="conciliacion_resultado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Conciliación lista ✅. Descarga tu archivo.")
