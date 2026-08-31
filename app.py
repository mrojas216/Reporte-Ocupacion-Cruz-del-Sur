import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Reporte Ocupación Cruz del Sur",
    page_icon="📊",
    layout="wide"
)


def leer_archivo(archivo):

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

    requeridas = [
        "Fecha salida",
        "Ruta",
        "Ocupación"
    ]

    faltantes = [
        c
        for c in requeridas
        if c not in df.columns
    ]

    if faltantes:

        raise Exception(
            f"Faltan columnas requeridas: {', '.join(faltantes)}"
        )


def limpiar_ocupacion(valor):

    if pd.isna(valor):
        return 0

    valor = str(valor)

    valor = valor.replace("%", "")
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

    df["Dia"] = df["Fecha salida"].dt.day

    tabla = pd.pivot_table(
        df,
        index="Ruta",
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
            "Ruta": "Etiquetas de fila"
        },
        inplace=True
    )

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

        header = workbook.add_format({
            "bold": True,
            "border": 1,
            "bg_color": "#D9D9D9",
            "align": "center"
        })

        numero = workbook.add_format({
            "num_format": "0.00"
        })

        for col_num, value in enumerate(df.columns):

            worksheet.write(
                0,
                col_num,
                value,
                header
            )

        worksheet.freeze_panes(
            1,
            1
        )

        worksheet.set_column(
            0,
            0,
            50
        )

        for col in range(
            1,
            len(df.columns)
        ):
            worksheet.set_column(
                col,
                col,
                10,
                numero
            )

    output.seek(0)

    return output


st.title(
    "📊 Ocupación Cruz del Sur"
)

archivo = st.file_uploader(
    "Seleccione archivo Excel",
    type=["xlsx"]
)

if archivo:

    try:

        df = leer_archivo(archivo)

        st.expander(
            "Columnas detectadas"
        ).write(
            df.columns.tolist()
        )

        validar_columnas(df)

        resultado = generar_tabla(df)

        st.success(
            f"Rutas detectadas: {len(resultado)}"
        )

        st.subheader(
            "Vista previa"
        )

        st.dataframe(
            resultado.head(20),
            use_container_width=True
        )

        st.info(
            """
            Los valores corresponden a la ocupación real
            obtenida desde la columna "Ocupación".

            Ejemplos:
            73,19% → 73.19
            5,93% → 5.93
            1,70% → 1.70
            """
        )

        excel = generar_excel(
            resultado
        )

        st.download_button(
            "📥 Descargar resultado",
            data=excel,
            file_name="cruzdelsur_tabla_resultado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(str(e))

