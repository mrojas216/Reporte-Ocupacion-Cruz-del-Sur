import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Tabla Objetivo Ocupación",
    page_icon="📊",
    layout="wide"
)

COLUMNAS_REQUERIDAS = ["Ruta"]


def leer_archivo_excel(archivo):
    """
    El archivo tiene dos filas informativas:
    Reporte de ocupación
    Reporte de ocupación, xx/xx/xxxx - xx/xx/xxxx

    Los encabezados reales comienzan en la fila 3.
    """

    df = pd.read_excel(
        archivo,
        header=2
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    return df


def validar_columnas(df):

    faltantes = [
        col
        for col in COLUMNAS_REQUERIDAS
        if col not in df.columns
    ]

    if faltantes:
        raise ValueError(
            f"Faltan columnas requeridas: {', '.join(faltantes)}"
        )


def construir_tabla(df):

    rutas = (
        df["Ruta"]
        .dropna()
        .astype(str)
        .str.strip()
        .sort_values()
        .unique()
    )

    tabla = pd.DataFrame({
        "Etiquetas de fila": rutas
    })

    patron = [0.0, 4.5, 8.0]

    for dia in range(1, 32):

        if dia == 28:
            valor = 50.0

        elif dia == 29:
            valor = 100.0

        elif dia == 30:
            valor = 80.0

        elif dia == 31:
            valor = 90.0

        else:
            valor = patron[(dia - 1) % 3]

        tabla[dia] = valor

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
            "align": "center"
        })

        formato_numero = workbook.add_format({
            "num_format": "0.0"
        })

        for col_num, valor in enumerate(df.columns):
            worksheet.write(
                0,
                col_num,
                valor,
                formato_header
            )

        worksheet.freeze_panes(1, 1)

        worksheet.set_column(0, 0, 50)

        for col in range(1, 32):
            worksheet.set_column(
                col,
                col,
                8,
                formato_numero
            )

        for fila in range(1, len(df) + 1):
            for col in range(1, 32):

                worksheet.write_number(
                    fila,
                    col,
                    float(df.iloc[fila - 1, col]),
                    formato_numero
                )

    output.seek(0)

    return output


# ----------------------------
# INTERFAZ STREAMLIT
# ----------------------------

st.title("📊 Generador Tabla Objetivo de Ocupación")

st.markdown(
    """
    Cargue el archivo exportado desde Cruz del Sur.

    El programa extraerá automáticamente las rutas y generará
    la matriz objetivo por día del mes.
    """
)

archivo = st.file_uploader(
    "Seleccione archivo Excel",
    type=["xlsx"]
)

if archivo is not None:

    try:

        df = leer_archivo_excel(archivo)

        validar_columnas(df)

        resultado = construir_tabla(df)

        st.success(
            f"Proceso completado correctamente. "
            f"Rutas encontradas: {len(resultado)}"
        )

        with st.expander("Ver columnas detectadas"):
            st.write(df.columns.tolist())

        st.subheader("Vista previa")

        st.dataframe(
            resultado.head(20),
            use_container_width=True
        )

        st.subheader("Interpretación de los valores")

        ejemplo = pd.DataFrame({
            "Valor numérico": [
                0,
                4.5,
                8,
                50,
                80,
                90,
                100
            ],
            "Equivale a": [
                "0%",
                "4,5%",
                "8%",
                "50%",
                "80%",
                "90%",
                "100%"
            ]
        })

        st.table(ejemplo)

        st.info(
            """
            Los valores se guardan como números entre 0 y 100.

            Ejemplo:
            • 4.5 significa 4,5%
            • 50 significa 50%
            • 100 significa 100%

            No se almacenan como porcentaje Excel (0.045, 0.50, 1.00).
            """
        )

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
