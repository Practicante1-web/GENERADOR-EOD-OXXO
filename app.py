import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED

# =========================================================
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Generador EOD OXXO",
    page_icon="🟢",
    layout="wide"
)

st.title("🟢 Generador de Reporte EOD OXXO")
st.write(
    "Carga el CSV de encuestas y la plantilla Word para "
    "generar el informe conservando su diseño."
)

# =========================================================
# FUNCIONES
# =========================================================

def buscar_columna(df, nombres):
    columnas = {
        str(c).strip().lower().replace(" ", "").replace("_", ""): c
        for c in df.columns
    }

    for nombre in nombres:
        clave = (
            nombre.strip()
            .lower()
            .replace(" ", "")
            .replace("_", "")
        )

        if clave in columnas:
            return columnas[clave]

    return None


def porcentaje(cantidad, total):
    if total == 0:
        return 0
    return round((cantidad / total) * 100)


def tabla_frecuencia(df, columna):
    if columna is None or columna not in df.columns:
        return pd.DataFrame(
            columns=["Respuesta", "Cantidad", "%"]
        )

    datos = (
        df[columna]
        .dropna()
        .astype(str)
        .str.strip()
    )

    datos = datos[datos != ""]

    if len(datos) == 0:
        return pd.DataFrame(
            columns=["Respuesta", "Cantidad", "%"]
        )

    tabla = datos.value_counts().reset_index()
    tabla.columns = ["Respuesta", "Cantidad"]

    tabla["%"] = tabla["Cantidad"].apply(
        lambda x: porcentaje(x, len(datos))
    )

    return tabla


def formato_fecha(fecha):
    if fecha is None or pd.isna(fecha):
        return ""

    return pd.to_datetime(fecha).strftime("%d/%m/%Y")


def formato_fecha_hora(fecha):
    if fecha is None or pd.isna(fecha):
        return ""

    return pd.to_datetime(fecha).strftime("%d/%m/%Y %H:%M")


def limpiar_texto(valor):
    if valor is None:
        return ""

    return str(valor).strip()


# =========================================================
# REEMPLAZO DE MARCADORES EN WORD
# =========================================================
#
# IMPORTANTE:
# La plantilla debe usar marcadores ÚNICOS.
#
# Ejemplos:
# {{TIENDA}}
# {{ENCUESTAS}}
# {{ORIGEN_1}}
# {{ORIGEN_1_CANT}}
# {{ORIGEN_1_PCT}}
#
# NO usar [CANTIDAD], [PORCENTAJE], etc. repetidos.
# =========================================================

def reemplazar_marcadores_en_docx(plantilla_bytes, reemplazos):

    entrada = BytesIO(plantilla_bytes)
    salida = BytesIO()

    archivos_xml = [
        "word/document.xml",
        "word/header1.xml",
        "word/header2.xml",
        "word/header3.xml",
        "word/footer1.xml",
        "word/footer2.xml",
        "word/footer3.xml",
    ]

    with ZipFile(entrada, "r") as zin, ZipFile(
        salida,
        "w",
        ZIP_DEFLATED
    ) as zout:

        for item in zin.infolist():

            data = zin.read(item.filename)

            if item.filename in archivos_xml:
                try:
                    texto = data.decode("utf-8")

                    for marcador, valor in reemplazos.items():
                        texto = texto.replace(
                            marcador,
                            limpiar_texto(valor)
                        )

                    data = texto.encode("utf-8")

                except Exception:
                    pass

            zout.writestr(item, data)

    salida.seek(0)
    return salida


# =========================================================
# CARGAR ARCHIVOS
# =========================================================

st.header("1. Archivos")

archivo = st.file_uploader(
    "Selecciona el archivo CSV de encuestas",
    type=["csv"]
)

plantilla = st.file_uploader(
    "Selecciona la plantilla EOD en Word",
    type=["docx"]
)

if archivo is None or plantilla is None:
    st.info(
        "Carga el CSV y la plantilla Word para continuar."
    )
    st.stop()


# =========================================================
# LEER CSV
# =========================================================

try:
    archivo.seek(0)

    df = pd.read_csv(
        archivo,
        encoding="utf-8"
    )

except Exception:

    archivo.seek(0)

    df = pd.read_csv(
        archivo,
        encoding="latin-1"
    )


df.columns = [
    str(c).strip()
    for c in df.columns
]

st.success(
    f"Archivo cargado correctamente: {len(df)} encuestas."
)


# =========================================================
# IDENTIFICAR COLUMNAS
# =========================================================

col_fecha = buscar_columna(
    df,
    [
        "CreationDate",
        "Creation Date"
    ]
)

col_tienda = buscar_columna(
    df,
    [
        "Nombre tienda estudiada"
    ]
)

col_edad = buscar_columna(
    df,
    [
        "Edad"
    ]
)

col_genero = buscar_columna(
    df,
    [
        "Género",
        "Genero"
    ]
)

col_estrato = buscar_columna(
    df,
    [
        "Estrato"
    ]
)

col_ocupacion = buscar_columna(
    df,
    [
        "Ocupación actual",
        "Ocupacion actual"
    ]
)

col_motivo = buscar_columna(
    df,
    [
        "Por qué compraría ahí?"
    ]
)

col_transporte = buscar_columna(
    df,
    [
        "Medio de transporte usado para llegar a OXXO"
    ]
)

col_origen = buscar_columna(
    df,
    [
        "De dónde viene?"
    ]
)

col_destino = buscar_columna(
    df,
    [
        "Hacia dónde se dirige?"
    ]
)

col_alternativa = buscar_columna(
    df,
    [
        "Dónde compraría sino es en OXXO?"
    ]
)

# =========================================================
# VALIDACIÓN
# =========================================================

if col_tienda is None:

    st.error(
        "No encontré la columna 'Nombre tienda estudiada'."
    )

    st.write("Columnas encontradas:")
    st.write(list(df.columns))

    st.stop()


# =========================================================
# SELECCIONAR TIENDA
# =========================================================

st.header("2. Seleccionar tienda")

tiendas = (
    df[col_tienda]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
)

tiendas = sorted(tiendas)

tienda_seleccionada = st.selectbox(
    "Selecciona el CR / tienda",
    tiendas
)


datos_tienda = df[
    df[col_tienda]
    .astype(str)
    .str.strip()
    == tienda_seleccionada
].copy()


# =========================================================
# FECHAS
# =========================================================

fecha_inicio = None
fecha_final = None
dia_mas_encuestas = ""
cantidad_dia_mas = 0
dias_sin_encuestas = []

if col_fecha is not None:

    datos_tienda[col_fecha] = pd.to_datetime(
        datos_tienda[col_fecha],
        errors="coerce"
    )

    fechas = datos_tienda[col_fecha].dropna()

    if len(fechas) > 0:

        fecha_inicio = fechas.min()
        fecha_final = fechas.max()

        conteo_dias = (
            fechas.dt.date
            .value_counts()
            .sort_index()
        )

        if len(conteo_dias) > 0:

            dia_mas_encuestas = conteo_dias.idxmax()

            cantidad_dia_mas = int(
                conteo_dias.max()
            )

            todos_los_dias = pd.date_range(
                start=fecha_inicio.date(),
                end=fecha_final.date(),
                freq="D"
            ).date

            dias_sin_encuestas = [
                dia.strftime("%d/%m/%Y")
                for dia in todos_los_dias
                if dia not in conteo_dias.index
            ]


# =========================================================
# INFORMACIÓN AUTOMÁTICA
# =========================================================

st.header("3. Información automática")

total_encuestas = len(datos_tienda)

if (
    col_fecha is not None
    and fecha_inicio is not None
    and fecha_final is not None
):

    if fecha_inicio.date() == fecha_final.date():

        periodo = formato_fecha(fecha_inicio)

    else:

        periodo = (
            f"{formato_fecha(fecha_inicio)} - "
            f"{formato_fecha(fecha_final)}"
        )

else:

    periodo = ""


col1, col2, col3 = st.columns(3)

with col1:
    st.metric(
        "Encuestas",
        total_encuestas
    )

with col2:
    st.metric(
        "Inicio",
        formato_fecha_hora(fecha_inicio)
    )

with col3:
    st.metric(
        "Finalización",
        formato_fecha_hora(fecha_final)
    )


col4, col5, col6 = st.columns(3)

with col4:
    st.metric(
        "Día con más encuestas",
        str(dia_mas_encuestas)
    )

with col5:
    st.metric(
        "Cantidad ese día",
        cantidad_dia_mas
    )

with col6:
    st.write("**Periodo**")
    st.write(periodo)


if dias_sin_encuestas:

    st.warning(
        "Días sin encuestas: "
        + ", ".join(dias_sin_encuestas)
    )

else:

    st.success(
        "No se encontraron días sin encuestas."
    )


# =========================================================
# PERFIL DEL CLIENTE
# =========================================================

st.header("4. Perfil del cliente")

edad_promedio = ""

if col_edad is not None:

    edades = pd.to_numeric(
        datos_tienda[col_edad],
        errors="coerce"
    )

    if edades.notna().any():

        edad_promedio = round(
            edades.mean(),
            1
        )


hombres = 0
mujeres = 0

if col_genero is not None:

    genero = (
        datos_tienda[col_genero]
        .astype(str)
        .str.strip()
        .str.lower()
    )

    hombres = int(
        genero.str.contains(
            "hombre",
            na=False
        ).sum()
    )

    mujeres = int(
        genero.str.contains(
            "mujer",
            na=False
        ).sum()
    )


estrato_tabla = tabla_frecuencia(
    datos_tienda,
    col_estrato
)

ocupacion_tabla = tabla_frecuencia(
    datos_tienda,
    col_ocupacion
)


p1, p2, p3, p4 = st.columns(4)

with p1:
    st.metric(
        "Edad promedio",
        (
            f"{edad_promedio} años"
            if edad_promedio != ""
            else "N/D"
        )
    )

with p2:
    st.metric(
        "Hombres",
        hombres
    )

with p3:
    st.metric(
        "Mujeres",
        mujeres
    )

with p4:

    if not estrato_tabla.empty:

        estrato_principal = (
            estrato_tabla.iloc[0]["Respuesta"]
        )

        estrato_porcentaje = (
            estrato_tabla.iloc[0]["%"]
        )

        st.metric(
            "Estrato principal",
            (
                f"{estrato_principal} "
                f"({estrato_porcentaje}%)"
            )
        )

    else:

        st.metric(
            "Estrato principal",
            "N/D"
        )


if not ocupacion_tabla.empty:

    st.write("**Ocupación principal**")

    st.dataframe(
        ocupacion_tabla,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# TABLAS DEL ESTUDIO
# =========================================================

st.header("5. Información del estudio")

tablas = {

    "motivo": tabla_frecuencia(
        datos_tienda,
        col_motivo
    ),

    "transporte": tabla_frecuencia(
        datos_tienda,
        col_transporte
    ),

    "origen": tabla_frecuencia(
        datos_tienda,
        col_origen
    ),

    "destino": tabla_frecuencia(
        datos_tienda,
        col_destino
    ),

    "alternativa": tabla_frecuencia(
        datos_tienda,
        col_alternativa
    )
}


t1, t2 = st.columns(2)

with t1:

    st.subheader(
        "Principal motivo de compra"
    )

    st.dataframe(
        tablas["motivo"],
        use_container_width=True,
        hide_index=True
    )


with t2:

    st.subheader(
        "Medio de llegada"
    )

    st.dataframe(
        tablas["transporte"],
        use_container_width=True,
        hide_index=True
    )


t3, t4 = st.columns(2)

with t3:

    st.subheader(
        "De dónde viene"
    )

    st.dataframe(
        tablas["origen"],
        use_container_width=True,
        hide_index=True
    )


with t4:

    st.subheader(
        "A dónde se dirige"
    )

    st.dataframe(
        tablas["destino"],
        use_container_width=True,
        hide_index=True
    )


st.subheader(
    "Alternativa de compra"
)

st.dataframe(
    tablas["alternativa"],
    use_container_width=True,
    hide_index=True
)


# =========================================================
# INFORMACIÓN MANUAL
# =========================================================

st.header("6. Información manual")

st.info(
    "Estos campos los diligencias tú porque corresponden "
    "al análisis manual del estudio."
)

percepcion = st.text_area(
    "Percepción del servicio"
)

radio100 = st.text_input(
    "Radio 100 m"
)

radio200 = st.text_input(
    "Radio 200 m"
)

radio300 = st.text_input(
    "Radio 300 m"
)

radio300mas = st.text_input(
    "Radio +300 m"
)

influencia = st.text_area(
    "¿Le hizo influencia?"
)

observaciones = st.text_area(
    "Observaciones / información adicional"
)

st.subheader("Insights clave")

insight1 = st.text_area(
    "Insight 1"
)

insight2 = st.text_area(
    "Insight 2"
)

insight3 = st.text_area(
    "Insight 3"
)


# =========================================================
# FOTOGRAFÍAS
# =========================================================

st.subheader(
    "Registro fotográfico"
)

fotos = st.file_uploader(
    "Sube las fotografías",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)


# =========================================================
# CONSTRUIR REEMPLAZOS
# =========================================================
#
# Cada campo tiene su propio marcador.
# Así evitamos el problema anterior de que todos los
# [CANTIDAD] y [%] se llenaran con el mismo dato.
# =========================================================

reemplazos = {

    "{{TIENDA}}":
        tienda_seleccionada,

    "{{PERIODO}}":
        periodo,

    "{{ENCUESTAS}}":
        total_encuestas,

    "{{INICIO}}":
        formato_fecha_hora(fecha_inicio),

    "{{FINALIZACION}}":
        formato_fecha_hora(fecha_final),

    "{{DIA_MAS_ENCUESTAS}}":
        dia_mas_encuestas,

    "{{CANTIDAD_DIA_MAS}}":
        cantidad_dia_mas,

    "{{DIAS_SIN_ENCUESTAS}}":
        (
            ", ".join(dias_sin_encuestas)
            if dias_sin_encuestas
            else "Ninguno"
        ),

    "{{EDAD_PROMEDIO}}":
        edad_promedio,

    "{{HOMBRES}}":
        hombres,

    "{{MUJERES}}":
        mujeres,

    "{{ESTRATO}}":
        (
            estrato_tabla.iloc[0]["Respuesta"]
            if not estrato_tabla.empty
            else "N/D"
        ),

    "{{ESTRATO_PCT}}":
        (
            f"{estrato_tabla.iloc[0]['%']}%"
            if not estrato_tabla.empty
            else "N/D"
        ),

    "{{OCUPACION}}":
        (
            ocupacion_tabla.iloc[0]["Respuesta"]
            if not ocupacion_tabla.empty
            else "N/D"
        ),

    "{{OCUPACION_PCT}}":
        (
            f"{ocupacion_tabla.iloc[0]['%']}%"
            if not ocupacion_tabla.empty
            else "N/D"
        ),

    "{{PERCEPCION}}":
        percepcion,

    "{{RADIO_100}}":
        radio100,

    "{{RADIO_200}}":
        radio200,

    "{{RADIO_300}}":
        radio300,

    "{{RADIO_MAS_300}}":
        radio300mas,

    "{{INFLUENCIA}}":
        influencia,

    "{{OBSERVACIONES}}":
        observaciones,

    "{{INSIGHT_1}}":
        insight1,

    "{{INSIGHT_2}}":
        insight2,

    "{{INSIGHT_3}}":
        insight3,
}


# =========================================================
# AGREGAR DATOS DE TABLAS
# =========================================================

def agregar_filas(
    reemplazos,
    tabla,
    prefijo,
    max_filas=4
):

    for i in range(max_filas):

        numero = i + 1

        if i < len(tabla):

            fila = tabla.iloc[i]

            reemplazos[
                f"{{{{{prefijo}_{numero}}}}}"
            ] = fila["Respuesta"]

            reemplazos[
                f"{{{{{prefijo}_{numero}_CANT}}}}"
            ] = fila["Cantidad"]

            reemplazos[
                f"{{{{{prefijo}_{numero}_PCT}}}}"
            ] = f"{fila['%']}%"

        else:

            reemplazos[
                f"{{{{{prefijo}_{numero}}}}}"
            ] = ""

            reemplazos[
                f"{{{{{prefijo}_{numero}_CANT}}}}"
            ] = ""

            reemplazos[
                f"{{{{{prefijo}_{numero}_PCT}}}}"
            ] = ""


agregar_filas(
    reemplazos,
    tablas["motivo"],
    "MOTIVO",
    3
)

agregar_filas(
    reemplazos,
    tablas["transporte"],
    "TRANSPORTE",
    4
)

agregar_filas(
    reemplazos,
    tablas["origen"],
    "ORIGEN",
    4
)

agregar_filas(
    reemplazos,
    tablas["destino"],
    "DESTINO",
    4
)

agregar_filas(
    reemplazos,
    tablas["alternativa"],
    "ALTERNATIVA",
    5
)


# =========================================================
# GENERAR REPORTE
# =========================================================

st.header("7. Generar reporte")

if st.button(
    "🟢 GENERAR REPORTE",
    type="primary"
):

    plantilla.seek(0)

    plantilla_bytes = plantilla.read()

    documento = reemplazar_marcadores_en_docx(
        plantilla_bytes,
        reemplazos
    )

    nombre_archivo = (
        f"Reporte_EOD_{tienda_seleccionada}.docx"
    )

    st.success(
        "¡Reporte generado correctamente! 🎉"
    )

    st.download_button(
        label="📥 Descargar informe EOD",
        data=documento,
        file_name=nombre_archivo,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )
