import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
from docx import Document
from docx.shared import Inches


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
    "Carga el archivo de encuestas y la Plantilla EOD para generar el reporte."
)


# =========================================================
# FUNCIONES GENERALES
# =========================================================

def normalizar(texto):
    """
    Normaliza textos para poder encontrar columnas
    aunque tengan mayúsculas, tildes, espacios o símbolos.
    """
    texto = str(texto).lower().strip()

    reemplazos = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ñ": "n"
    }

    for a, b in reemplazos.items():
        texto = texto.replace(a, b)

    texto = re.sub(r"[^a-z0-9]", "", texto)

    return texto


def buscar_columna(df, nombres):
    """
    Busca una columna por coincidencia exacta normalizada.
    """
    columnas = {
        normalizar(c): c
        for c in df.columns
    }

    for nombre in nombres:

        clave = normalizar(nombre)

        if clave in columnas:
            return columnas[clave]

    return None


def buscar_columna_parcial(df, palabras):
    """
    Busca una columna que contenga todas las palabras indicadas.
    """
    for columna in df.columns:

        texto = normalizar(columna)

        if all(normalizar(palabra) in texto for palabra in palabras):
            return columna

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

    tabla = (
        datos.value_counts()
        .reset_index()
    )

    tabla.columns = [
        "Respuesta",
        "Cantidad"
    ]

    tabla["%"] = tabla["Cantidad"].apply(
        lambda x: porcentaje(x, len(datos))
    )

    return tabla


def formato_fecha(fecha):

    if fecha is None or pd.isna(fecha):
        return ""

    return pd.to_datetime(fecha).strftime(
        "%d/%m/%Y"
    )


def formato_fecha_hora(fecha):

    if fecha is None or pd.isna(fecha):
        return ""

    return pd.to_datetime(fecha).strftime(
        "%d/%m/%Y %H:%M"
    )


# =========================================================
# TRÁFICO / TIEMPO DE LLEGADA
# =========================================================

def encontrar_columna_trafico(df):

    candidatos = []

    for columna in df.columns:

        texto = normalizar(columna)

        palabras_trafico = [
            "trafico",
            "tiempo",
            "llegada",
            "traslado",
            "demora",
            "minutos",
            "tarda"
        ]

        if any(
            palabra in texto
            for palabra in palabras_trafico
        ):

            candidatos.append(columna)

    # Primero buscamos columnas claramente relacionadas
    # con tiempo de llegada.

    prioridades = [
        ["tiempo", "llegada"],
        ["tiempo", "traslado"],
        ["tiempo"],
        ["llegada"],
        ["minutos"],
        ["demora"],
        ["tarda"],
        ["trafico"]
    ]

    for palabras in prioridades:

        for columna in candidatos:

            texto = normalizar(columna)

            if all(
                palabra in texto
                for palabra in palabras
            ):

                return columna

    return None


def calcular_trafico_promedio(df, columna):

    if columna is None:
        return None

    serie = df[columna].dropna()

    if len(serie) == 0:
        return None

    valores = []

    for valor in serie:

        texto = str(valor).lower().strip()

        # Buscar número
        numeros = re.findall(
            r"\d+(?:[.,]\d+)?",
            texto
        )

        if not numeros:
            continue

        numero = numeros[0].replace(",", ".")

        try:
            numero = float(numero)
        except:
            continue

        # Si la respuesta contiene horas,
        # convertir a minutos.
        if "hora" in texto:

            numero = numero * 60

        valores.append(numero)

    if not valores:
        return None

    promedio = sum(valores) / len(valores)

    return round(promedio, 1)


# =========================================================
# LEER CSV
# =========================================================

st.header("1. Archivos")


archivo_csv = st.file_uploader(
    "Selecciona el archivo CSV de encuestas",
    type=["csv"]
)


plantilla = st.file_uploader(
    "Selecciona la Plantilla EOD",
    type=["docx"]
)


if archivo_csv is None or plantilla is None:

    st.info(
        "Carga los dos archivos para continuar."
    )

    st.stop()


# =========================================================
# CARGAR CSV
# =========================================================

try:

    archivo_csv.seek(0)

    df = pd.read_csv(
        archivo_csv,
        encoding="utf-8"
    )

except:

    archivo_csv.seek(0)

    df = pd.read_csv(
        archivo_csv,
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
        "Creation Date",
        "Fecha",
        "Fecha de creación"
    ]
)


col_tienda = buscar_columna(
    df,
    [
        "Nombre tienda estudiada",
        "Nombre de tienda estudiada"
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
        "Por qué compraría ahí?",
        "Por que compraria ahi?"
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
        "De dónde viene?",
        "De donde viene?"
    ]
)


col_destino = buscar_columna(
    df,
    [
        "Hacia dónde se dirige?",
        "Hacia donde se dirige?"
    ]
)


col_alternativa = buscar_columna(
    df,
    [
        "Dónde compraría sino es en OXXO?",
        "Donde compraria sino es en OXXO?"
    ]
)


# =========================================================
# BUSCAR TRÁFICO AUTOMÁTICAMENTE
# =========================================================

col_trafico = encontrar_columna_trafico(df)


# =========================================================
# VALIDAR TIENDA
# =========================================================

if col_tienda is None:

    st.error(
        "No encontré la columna 'Nombre tienda estudiada'."
    )

    st.write("Columnas encontradas:")

    st.write(
        list(df.columns)
    )

    st.stop()


# =========================================================
# SELECCIONAR TIENDA / CR
# =========================================================

st.header("2. Seleccionar CR / tienda")


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


# =========================================================
# FILTRAR TIENDA
# =========================================================

datos_tienda = df[
    df[col_tienda]
    .astype(str)
    .str.strip()
    ==
    tienda_seleccionada
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

    fechas = (
        datos_tienda[col_fecha]
        .dropna()
    )

    if len(fechas) > 0:

        fecha_inicio = fechas.min()

        fecha_final = fechas.max()

        conteo_dias = (
            fechas.dt.date
            .value_counts()
            .sort_index()
        )

        if len(conteo_dias) > 0:

            dia_mas_encuestas = (
                conteo_dias.idxmax()
            )

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


total_encuestas = len(
    datos_tienda
)


if (
    fecha_inicio is not None
    and fecha_final is not None
):

    if fecha_inicio.date() == fecha_final.date():

        periodo = formato_fecha(
            fecha_inicio
        )

    else:

        periodo = (
            f"{formato_fecha(fecha_inicio)} - "
            f"{formato_fecha(fecha_final)}"
        )

else:

    periodo = ""


trafico_promedio = calcular_trafico_promedio(
    datos_tienda,
    col_trafico
)


c1, c2, c3, c4 = st.columns(4)


with c1:

    st.metric(
        "Encuestas",
        total_encuestas
    )


with c2:

    st.metric(
        "Inicio",
        formato_fecha_hora(
            fecha_inicio
        )
    )


with c3:

    st.metric(
        "Finalización",
        formato_fecha_hora(
            fecha_final
        )
    )


with c4:

    if trafico_promedio is not None:

        st.metric(
            "Tráfico promedio",
            f"{trafico_promedio} min"
        )

    else:

        st.metric(
            "Tráfico promedio",
            "N/D"
        )


c5, c6 = st.columns(2)


with c5:

    st.metric(
        "Día con más encuestas",
        (
            dia_mas_encuestas.strftime(
                "%d/%m/%Y"
            )
            if hasattr(
                dia_mas_encuestas,
                "strftime"
            )
            else str(
                dia_mas_encuestas
            )
        )
    )


with c6:

    st.metric(
        "Cantidad ese día",
        cantidad_dia_mas
    )


st.write(
    f"**Periodo:** {periodo}"
)


if dias_sin_encuestas:

    st.warning(
        "Días sin encuestas: "
        +
        ", ".join(
            dias_sin_encuestas
        )
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
            f"{estrato_principal} "
            f"({estrato_porcentaje}%)"
        )

    else:

        st.metric(
            "Estrato principal",
            "N/D"
        )


if not ocupacion_tabla.empty:

    st.write(
        "**Ocupación principal**"
    )

    st.dataframe(
        ocupacion_tabla,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# TABLAS DEL ESTUDIO
# =========================================================

st.header(
    "5. Información del estudio"
)


tablas = {

    "motivo":
        tabla_frecuencia(
            datos_tienda,
            col_motivo
        ),

    "transporte":
        tabla_frecuencia(
            datos_tienda,
            col_transporte
        ),

    "origen":
        tabla_frecuencia(
            datos_tienda,
            col_origen
        ),

    "destino":
        tabla_frecuencia(
            datos_tienda,
            col_destino
        ),

    "alternativa":
        tabla_frecuencia(
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
        "Origen"
    )

    st.dataframe(
        tablas["origen"],
        use_container_width=True,
        hide_index=True
    )


with t4:

    st.subheader(
        "Destino"
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

st.header(
    "6. Información manual"
)


st.info(
    "Estos campos los diligencias manualmente."
)


percepcion = st.text_area(
    "Percepción del servicio",
    placeholder=(
        "Escribe el comentario de percepción..."
    )
)


radio_100 = st.text_input(
    "Radio 100m",
    placeholder="Ejemplo: 47"
)


radio_200 = st.text_input(
    "Radio 200m",
    placeholder="Ejemplo: 48"
)


radio_300 = st.text_input(
    "Radio 300m",
    placeholder="Ejemplo: 74"
)


radio_300_mas = st.text_input(
    "Radio +300m",
    placeholder="Ejemplo: 67"
)


st.subheader(
    "Imagen del radio de influencia"
)


imagen_radio = st.file_uploader(
    "Sube la imagen del radio",
    type=[
        "jpg",
        "jpeg",
        "png"
    ],
    key="imagen_radio"
)


st.subheader(
    "Imagen de la isócrona"
)


imagen_isocrona = st.file_uploader(
    "Sube la imagen de la isócrona",
    type=[
        "jpg",
        "jpeg",
        "png"
    ],
    key="imagen_isocrona"
)


st.subheader(
    "Insights clave"
)


insight1 = st.text_area(
    "Insight 1"
)


insight2 = st.text_area(
    "Insight 2"
)


insight3 = st.text_area(
    "Insight 3"
)


st.subheader(
    "Registro fotográfico"
)


fotos = st.file_uploader(
    "Sube las fotografías",
    type=[
        "jpg",
        "jpeg",
        "png"
    ],
    accept_multiple_files=True
)


# =========================================================
# FACHADA
# =========================================================

st.subheader(
    "Foto de fachada"
)


foto_fachada = st.file_uploader(
    "Sube la foto de fachada",
    type=[
        "jpg",
        "jpeg",
        "png"
    ],
    key="foto_fachada"
)


# =========================================================
# PERCEPCIÓN AUTOMÁTICA
# =========================================================

st.subheader(
    "Percepción del servicio"
)


st.write(
    "La cantidad de estrellas y porcentaje "
    "se pueden diligenciar manualmente mientras "
    "definimos la columna exacta del dashboard."
)


estrellas = st.number_input(
    "Cantidad de estrellas",
    min_value=1,
    max_value=5,
    value=5
)


porcentaje_percepcion = st.number_input(
    "Porcentaje de percepción",
    min_value=0,
    max_value=100,
    value=0
)


# =========================================================
# DATOS PARA REPORTE
# =========================================================

datos = {

    "tienda":
        tienda_seleccionada,

    "encuestas":
        total_encuestas,

    "periodo":
        periodo,

    "inicio":
        formato_fecha_hora(
            fecha_inicio
        ),

    "finalizacion":
        formato_fecha_hora(
            fecha_final
        ),

    "dia_mas_encuestas":
        str(
            dia_mas_encuestas
        ),

    "cantidad_dia_mas":
        cantidad_dia_mas,

    "dias_sin_encuestas":
        dias_sin_encuestas,

    "trafico_promedio":
        (
            f"{trafico_promedio} min"
            if trafico_promedio is not None
            else "N/D"
        ),

    "edad_promedio":
        edad_promedio,

    "hombres":
        hombres,

    "mujeres":
        mujeres,

    "estrato":
        (
            estrato_tabla.iloc[0]["Respuesta"]
            if not estrato_tabla.empty
            else "N/D"
        ),

    "estrato_porcentaje":
        (
            estrato_tabla.iloc[0]["%"]
            if not estrato_tabla.empty
            else "N/D"
        ),

    "ocupacion":
        (
            ocupacion_tabla.iloc[0]["Respuesta"]
            if not ocupacion_tabla.empty
            else "N/D"
        ),

    "ocupacion_porcentaje":
        (
            ocupacion_tabla.iloc[0]["%"]
            if not ocupacion_tabla.empty
            else "N/D"
        ),

    "radio_100":
        radio_100,

    "radio_200":
        radio_200,

    "radio_300":
        radio_300,

    "radio_300_mas":
        radio_300_mas,

    "estrellas":
        estrellas,

    "porcentaje_percepcion":
        porcentaje_percepcion
}


# =========================================================
# REEMPLAZAR TEXTO EN WORD
# =========================================================

def reemplazar_texto(documento, reemplazos):

    for parrafo in documento.paragraphs:

        texto = parrafo.text

        for clave, valor in reemplazos.items():

            texto = texto.replace(
                clave,
                str(valor)
            )

        if texto != parrafo.text:

            parrafo.text = texto


    for tabla in documento.tables:

        for fila in tabla.rows:

            for celda in fila.cells:

                for parrafo in celda.paragraphs:

                    texto = parrafo.text

                    for clave, valor in reemplazos.items():

                        texto = texto.replace(
                            clave,
                            str(valor)
                        )

                    if texto != parrafo.text:

                        parrafo.text = texto


# =========================================================
# GENERAR WORD
# =========================================================

def generar_reporte(
    plantilla,
    datos,
    tablas
):

    plantilla.seek(0)

    documento = Document(
        plantilla
    )


    reemplazos = {

        "[NOMBRE TIENDA / CR]":
            datos["tienda"],

        "[NOMBRE TIENDA]":
            datos["tienda"],

        "[CR]":
            datos["tienda"],

        "[PERIODO]":
            datos["periodo"],

        "[ENCUESTAS]":
            datos["encuestas"],

        "[AUTOMATICO]":
            datos["encuestas"],

        "[INICIO]":
            datos["inicio"],

        "[FINALIZACION]":
            datos["finalizacion"],

        "[TRAFICO PROMEDIO]":
            datos["trafico_promedio"],

        "[EDAD]":
            datos["edad_promedio"],

        "[CANTIDAD HOMBRE]":
            datos["hombres"],

        "[HOMBRE]":
            datos["hombres"],

        "[CANTIDAD MUJER]":
            datos["mujeres"],

        "[MUJER]":
            datos["mujeres"],

        "[ESTRATO]":
            datos["estrato"],

        "[PORCENTAJE]":
            datos["estrato_porcentaje"],

        "[OCUPACION]":
            datos["ocupacion"],

        "[PORCENTAJE OCUPACION]":
            datos["ocupacion_porcentaje"],

        "[RADIO 100]":
            datos["radio_100"],

        "[RADIO 200]":
            datos["radio_200"],

        "[RADIO 300]":
            datos["radio_300"],

        "[RADIO +300]":
            datos["radio_300_mas"],

        "[ESTRELLAS]":
            datos["estrellas"],

        "[PORCENTAJE PERCEPCION]":
            datos["porcentaje_percepcion"]
    }


    reemplazar_texto(
        documento,
        reemplazos
    )


    # =====================================================
    # GUARDAR
    # =====================================================

    buffer = BytesIO()

    documento.save(
        buffer
    )

    buffer.seek(0)

    return buffer


# =========================================================
# GENERAR REPORTE
# =========================================================

st.header(
    "7. Generar reporte"
)


if st.button(
    "🟢 GENERAR REPORTE",
    type="primary"
):

    documento = generar_reporte(
        plantilla,
        datos,
        tablas
    )


    nombre_archivo = (
        f"Reporte_EOD_"
        f"{tienda_seleccionada}.docx"
    )


    st.success(
        "¡Reporte generado correctamente! 🎉"
    )


    st.download_button(
        label="📥 Descargar reporte",
        data=documento,
        file_name=nombre_archivo,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )
