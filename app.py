import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="Tabla Objetivo Ocupación",
    page_icon="📊",
    layout="wide"
)


# -----------------------------
# LOCALIZAR ENCABEZADOS
# -----------------------------
def detectar_fila_encabezado(df_raw):

    for i in range(min(30, len(df_raw))):

        fila = (
            df_raw.iloc[i]
            .fillna("")
            .astype(str)
            .str.strip()
            .tolist()
        )

        if "Ruta" in fila:
            return i

    return None


# -----------------------------
# LEER ARCHIVO
# -----------------------------
def leer_archivo(archivo):

    bruto = pd.read_excel(
        archivo,
        header=None,
        dtype=str
    )

    fila_header = detectar_fila_encabezado(bruto)

    if fila_header is None:
        raise Exception(
            "No se encontró la fila de encabezados."
        )

    encabezados = (
        bruto.iloc[fila_header]
        .fillna("")
        .astype(str)
        .str.strip()
        .tolist()
    )

    datos = bruto.iloc[fila_header + 1:].copy()

    datos.columns = encabezados

    datos = datos.reset_index(drop=True)

    datos.columns = (
        pd.Index(datos.columns)
        .astype(str)
        .str.strip()
    )

    return datos


# -----------------------------
# BUSCAR COLUMNA RUTA
# -----------------------------
def localizar_columna_ruta(df):

    for col in df.columns:

        nombre = str(col).strip().lower()

        if nombre == "ruta":
            return col

    for col in df.columns:

        nombre = str(col).strip().lower()

        if "ruta" in nombre:
            return col

    raise Exception(
        f"No se encontró columna Ruta. Columnas detectadas: {list(df.columns)}"
    )


# -----------------------------
# CONSTRUIR RESULTADO
# -----------------------------
def construir_tabla(rutas):

    resultado = pd.DataFrame({
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

        resultado[dia] = valor

    return resultado


# -----------------------------
# EXPORTAR EXCEL
# -----------------------------
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

        formato_num = workbook.add_format({
            "num_format": "0.0"
        })

        for c, nombre in enumerate(df.columns):

            worksheet.write(
                0,
                c,
                nombre,
                formato_header
            )

        worksheet.freeze_panes(1, 1)

        worksheet.set_column(
            0,
            0,
            55
        )

        for col in range(1, 32):

            worksheet.set_column(
                col,
                col,
                8,
                formato_num
            )

        for fila in range(1, len(df) + 1):

            for col in range(1, 32):

                worksheet.write_number(
                    fila,
                    col,
                    float(df.iloc[fila - 1, col]),
                    formato_num
                )

    output.seek(0)

    return output


# -----------------------------
# STREAMLIT
# -----------------------------
st.title("📊 Tabla Objetivo Cruz del Sur")

archivo = st.file_uploader(
    "Seleccione archivo de ocupación",
    type=["xlsx"]
)

if archivo is not None:

    try:

        df = leer_archivo(archivo)

        st.subheader("Columnas detectadas")

        st.write(list(df.columns))

        columna_ruta = localizar_columna_ruta(df)

        rutas = (
            df[columna_ruta]
            .dropna()
            .astype(str)
            .str.strip()
        )

        rutas = rutas[
            rutas.str.len() > 0
        ]

        rutas = sorted(
            rutas.unique().tolist()
        )

        resultado = construir_tabla(rutas)

        st.success(
            f"Rutas detectadas: {len(rutas)}"
        )

        st.subheader("Vista previa")

        st.dataframe(
            resultado.head(20),
            use_container_width=True
        )

        st.subheader(
            "Interpretación valores"
        )

        ejemplo = pd.DataFrame({
            "Valor": [
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

        excel = generar_excel(
            resultado
        )

        st.download_button(
            label="📥 Descargar Excel",
            data=excel,
            file_name="cruzdelsur_tabla_resultado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:

        st.error(
            f"Error: {str(e)}"
        )

        st.exception(e)
