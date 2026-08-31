import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Tabla Objetivo Ocupación",
    page_icon="📊",
    layout="wide"
)

COLUMNAS_REQUERIDAS = ["Ruta"]


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
        .astype(str)
        .str.strip()
        .dropna()
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


st.title("📊 Generador Tabla Objetivo de Ocupación")

st.markdown(
    """
    Cargue el archivo exportado desde el sistema.
    
    La aplicación generará automáticamente la tabla de objetivos por ruta y día.
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

        resultado = construir_tabla(df)

        st.success(
            f"Proceso completado. Se encontraron {len(resultado)} rutas."
        )

        st.subheader("Vista previa")

        st.info(
            """
            Los valores se almacenan como números entre 0 y 100.

            Ejemplos:

            • 0 = 0%  
            • 4.5 = 4,5%  
            • 8 = 8%  
            • 50 = 50%  
            • 80 = 80%  
            • 90 = 90%  
            • 100 = 100%
            """
        )

        st.dataframe(
            resultado.head(10),
            use_container_width=True
        )

        st.subheader("Ejemplo de interpretación")

        ejemplo = pd.DataFrame({
            "Valor almacenado": [
                0,
                4.5,
                8,
                50,
                80,
                90,
                100
            ],
            "Interpretación": [
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
