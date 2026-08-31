import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Reporte Ocupación Cruz del Sur",
    page_icon="📊",
    layout="wide"
)


def leer_archivo(archivo):
    """
    Lee el archivo Excel exportado por Cruz del Sur.
    """

    df = pd.read_excel(
        archivo,
        header=1
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    return df


def validar_columnas(df):

    columnas_requeridas = [
        "Fecha salida",
        "Hora salida",
        "Ruta",
        "Ocupación"
    ]

    faltantes = [
        col
        for col in columnas_requeridas
        if col not in df.columns
    ]

    if faltantes:

        raise Exception(
            f"Faltan columnas requeridas: {', '.join(faltantes)}"
        )


def limpiar_ocupacion(valor):

    if pd.isna(valor):
        return 0

    valor = str(valor).strip()

#    valor = valor.replace("%", "")
    valor = valor.replace(",", ".")

    try:
        return float(valor)
    except:
        return 0


def generar_tabla(df):

    df = df.copy()

    df["Ocupación"] = (
        df["Ocupación"]
        .apply(limpiar_ocupacion)
    )

    df["Fecha salida"] = pd.to_datetime(
        df["Fecha salida"],
        dayfirst=True,
        errors="coerce"
    )

    df["Dia"] = (
        df["Fecha salida"]
        .dt.day
    )

    tabla = pd.pivot_table(
        df,
        index=["Ruta", "Hora salida"],
        columns="Dia",
        values="Ocupación",
        aggfunc="max",
        fill_value=0
    )

    for dia in range(1, 32):

        if dia not in tabla.columns:
            tabla[dia] = 0

    tabla = tabla[
        sorted(tabla.columns)
    ]

    tabla.reset_index(
        inplace=True
    )

    tabla.rename(
        columns={
            "Ruta": "Etiquetas de fila",
            "Hora salida": "Hora Salida"
        },
        inplace=True
    )

    columnas_finales = (
        ["Etiquetas de fila", "Hora Salida"]
        + list(range(1, 32))
    )

    tabla = tabla[columnas_finales]

    return tabla


def generar_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Hoja1",
            index=False
        )

        workbook = writer.book
        worksheet = writer.sheets["Hoja1"]

        formato_header = workbook.add_format({
            "bold": True,
            "bg_color": "#D9D9D9",
            "border": 1,
            "align": "center",
            "valign": "vcenter"
        })

        formato_porcentaje = workbook.add_format({
            "num_format": "0.00%"
        })

        formato_hora = workbook.add_format({
            "align": "center"
        })

        for col_num, valor in enumerate(df.columns):

            worksheet.write(
                0,
                col_num,
                valor,
                formato_header
            )

        worksheet.freeze_panes(1, 2)

        worksheet.set_column(
            0,
            0,
            55
        )

        worksheet.set_column(
            1,
            1,
            12,
            formato_hora
        )

        for col in range(
            2,
            len(df.columns)
        ):
            worksheet.set_column(
                col,
                col,
                10,
                formato_porcentaje
            )

    output.seek(0)

    return output


# =====================================
# INTERFAZ STREAMLIT
# =====================================

st.title("📊 Reporte de Ocupación Cruz del Sur")

st.markdown("""
Carga el archivo exportado desde Cruz del Sur para generar
automáticamente la tabla consolidada de ocupación por:

- Ruta
- Hora de salida
- Día del mes
""")

archivo = st.file_uploader(
    "Seleccione archivo Excel",
    type=["xlsx"]
)

if archivo:

    try:

        df = leer_archivo(archivo)

        with st.expander("Columnas detectadas"):
            st.write(df.columns.tolist())

        validar_columnas(df)

        resultado = generar_tabla(df)

        st.success(
            f"Proceso completado correctamente. "
            f"Filas generadas: {len(resultado)}"
        )

        st.subheader("Vista previa")

        st.dataframe(
            resultado.head(20),
            use_container_width=True
        )

        st.info("""
Los porcentajes mostrados corresponden a la columna
'Ocupación' del archivo original.

Ejemplos:

• 73,19 % → 73,19 %
• 5,93 % → 5,93 %
• 1,70 % → 1,70 %

El archivo Excel se exporta con formato de porcentaje.
""")

        excel = generar_excel(resultado)

        st.download_button(
            label="📥 Descargar Excel",
            data=excel,
            file_name="cruzdelsur_tabla_resultado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(
            f"Error al procesar archivo: {str(e)}"
        )

        st.exception(e)
