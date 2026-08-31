import streamlit as st
import pandas as pd
from io import BytesIO

# =====================================================
# CONFIGURACIÓN
# =====================================================

st.set_page_config(
    page_title="Reporte Ocupación Cruz del Sur",
    page_icon="📊",
    layout="wide"
)

# =====================================================
# LECTURA ARCHIVO
# =====================================================

def leer_archivo(archivo):

    df = pd.read_excel(
        archivo,
        engine="openpyxl",
        header=1
    )

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    return df

# =====================================================
# VALIDACIONES
# =====================================================

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

# =====================================================
# LIMPIEZA OCUPACIÓN
# =====================================================

def limpiar_ocupacion(valor):

    if pd.isna(valor):
        return 0

    valor = str(valor).strip()
    valor = valor.replace("%", "")
    valor = valor.replace(",", ".")

    try:
        return float(valor) / 100
    except:
        return 0

# =====================================================
# GENERAR TABLA
# =====================================================

def generar_tabla(df):

    df = df.copy()

    df["Ocupación"] = df["Ocupación"].apply(
        limpiar_ocupacion
    )

    df["Fecha salida"] = pd.to_datetime(
        df["Fecha salida"],
        dayfirst=True,
        errors="coerce"
    )

    df["Dia"] = df["Fecha salida"].dt.day

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

    tabla.reset_index(inplace=True)

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

# =====================================================
# GENERAR EXCEL
# =====================================================

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

        formato_hora = workbook.add_format({
            "align": "center"
        })

        formato_porcentaje = workbook.add_format({
            "num_format": "0.00%"
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
            60
        )

        worksheet.set_column(
            1,
            1,
            12,
            formato_hora
        )

        for fila in range(len(df)):

            for col in range(2, len(df.columns)):

                worksheet.write_number(
                    fila + 1,
                    col,
                    float(df.iloc[fila, col]),
                    formato_porcentaje
                )

        for col in range(2, len(df.columns)):

            worksheet.set_column(
                col,
                col,
                12
            )

    output.seek(0)

    return output

# =====================================================
# STREAMLIT
# =====================================================

st.title("📊 Reporte de Ocupación Cruz del Sur")

st.markdown("""
Carga el archivo de ocupación exportado desde Cruz del Sur.

El sistema generará automáticamente una tabla consolidada con:

- Ruta
- Hora Salida
- Días del mes (1-31)
- Ocupación real en porcentaje
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

        vista = resultado.copy()

        for col in vista.columns[2:]:

            vista[col] = (
                vista[col]
                .fillna(0)
                .apply(
                    lambda x: f"{x:.2%}"
                )
            )

        opciones = [20, 50, 100, 200, 500, 1000]

        if len(vista) not in opciones:
            opciones.append(len(vista))

        cantidad_registros = st.selectbox(
            "Cantidad de registros a visualizar",
            options=opciones,
            index=len(opciones) - 1,
            format_func=lambda x: "Todos" if x == len(vista) else str(x)
        )

        st.dataframe(
            vista.head(cantidad_registros),
            use_container_width=True
        )

        st.info("""
Los porcentajes provienen directamente de la columna
'Ocupación' del archivo original.

Ejemplos:

73,19% → 73,19%
5,93% → 5,93%
1,70% → 1,70%

La vista previa y el Excel mostrarán el símbolo %.
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
