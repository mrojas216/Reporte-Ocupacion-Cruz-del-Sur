import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Reporte Ocupación Cruz del Sur",
    page_icon="📊",
    layout="wide"
)

COLUMNAS_REQUERIDAS = [
    "Fecha salida",
    "Hora salida",
    "Día de semana",
    "Origen (ciudad)",
    "Destino (ciudad)",
    "Ocupación",
    "Folio de viaje",
    "Ruta"
]


def validar_columnas(df):
    faltantes = [
        col for col in COLUMNAS_REQUERIDAS
        if col not in df.columns
    ]

    if faltantes:
        raise ValueError(
            f"Faltan las columnas: {', '.join(faltantes)}"
        )


def procesar_datos(df):
    df = df.copy()

    df["Ocupación"] = (
        df["Ocupación"]
        .astype(str)
        .str.replace("%", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    df["Ocupación"] = pd.to_numeric(
        df["Ocupación"],
        errors="coerce"
    )

    return df


def generar_excel(df):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="xlsxwriter"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="Reporte",
            index=False
        )

        workbook = writer.book
        worksheet = writer.sheets["Reporte"]

        formato_header = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAD3",
            "border": 1
        })

        formato_numero = workbook.add_format({
            "num_format": "0.00"
        })

        for col_num, value in enumerate(df.columns):
            worksheet.write(0, col_num, value, formato_header)

        try:
            idx = df.columns.get_loc("Ocupación")
            worksheet.set_column(idx, idx, 12, formato_numero)
        except:
            pass

        for i, col in enumerate(df.columns):
            ancho = max(
                len(str(col)),
                min(
                    df[col].astype(str).str.len().max(),
                    50
                ) if len(df) else 15
            )

            worksheet.set_column(i, i, ancho + 2)

    output.seek(0)

    return output


st.title("📊 Generador Reporte Ocupación")

st.markdown(
    """
    Cargue el archivo Excel exportado desde el sistema.
    """
)

archivo = st.file_uploader(
    "Seleccione archivo Excel",
    type=["xlsx"]
)

if archivo:

    try:

        df = pd.read_excel(archivo)

        validar_columnas(df)

        df_resultado = procesar_datos(df)

        st.success(
            f"Archivo procesado correctamente. "
            f"Registros: {len(df_resultado):,}"
        )

        st.dataframe(
            df_resultado.head(20),
            use_container_width=True
        )

        excel = generar_excel(df_resultado)

        st.download_button(
            label="📥 Descargar Reporte",
            data=excel,
            file_name="reporte_ocupacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(
            f"Error al procesar archivo: {str(e)}"
        )
