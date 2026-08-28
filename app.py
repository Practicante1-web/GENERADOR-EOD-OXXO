import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata
import re

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT


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
    "Carga el archivo de encuestas y selecciona la tienda "
    "para generar la información del estudio."
)


# =========================================================
# FUNCIONES PARA COLUMNAS
# =========================================================

def normalizar_texto(texto):
    """
    Normaliza un texto para poder comparar nombres de columnas
    aunque tengan diferencias de:
    - mayúsculas
    - minúsculas
    - tildes
    - signos
    - espacios
    """

    texto = str(texto).strip().lower()

    # Quitar tildes
    texto = unicodedata.normalize("NFD", texto)

    texto = "".join(
        c for c in texto
        if unicodedata.category(c) != "Mn"
    )

    # Convertir signos y caracteres especiales en espacios
    texto = re.sub(
        r"[^a-z0-9]+",
        " ",
        texto
    )

    # Quitar espacios repetidos
    texto = re.sub(
        r"\s+",
        " ",
        texto
    ).strip()

    return texto


def buscar_columna(df, nombres):
    """
    Busca una columna tolerando diferencias de:
    - mayúsculas
    - tildes
    - signos de interrogación
    - espacios
    - pequeños cambios de escritura
    """

    columnas_originales = list(df.columns)

    columnas_normalizadas = {
        normalizar_texto(c): c
        for c in columnas_originales
    }

    # -----------------------------------------------------
    # 1. Coincidencia exacta normalizada
    # -----------------------------------------------------

    for nombre in nombres:

        clave = normalizar_texto(nombre)

        if clave in columnas_normalizadas:

            return columnas_normalizadas[clave]

    # -----------------------------------------------------
    # 2. Coincidencia parcial
    # -----------------------------------------------------

    for nombre in nombres:

        clave = normalizar_texto(nombre)

        for columna_normalizada, columna_original in columnas_normalizadas.items():

            if (
                clave in columna_normalizada
                or columna_normalizada in clave
            ):

                return columna_original

    return None


# =========================================================
# FUNCIONES GENERALES
# =========================================================

def porcentaje(cantidad, total):

    if total == 0:
        return 0

    return round(
        (cantidad / total) * 100
    )


def tabla_frecuencia(df, columna):

    if columna is None:
        return pd.DataFrame()

    if columna not in df.columns:
        return pd.DataFrame()

    datos = (
        df[columna]
        .dropna()
        .astype(str)
        .str.strip()
    )

    datos = datos[
        datos != ""
    ]

    if len(datos) == 0:
        return pd.DataFrame()

    tabla = (
        datos
        .value_counts()
        .reset_index()
    )

    tabla.columns = [
        "Respuesta",
        "Cantidad"
    ]

    tabla["%"] = tabla[
        "Cantidad"
    ].apply(
        lambda x: porcentaje(
            x,
            len(datos)
        )
    )

    return tabla


def formato_fecha(fecha):

    if fecha is None:
        return ""

    try:

        if pd.isna(fecha):
            return ""

    except:
        pass

    return pd.to_datetime(
        fecha
    ).strftime(
        "%d/%m/%Y"
    )


def formato_fecha_hora(fecha):

    if fecha is None:
        return ""

    try:

        if pd.isna(fecha):
            return ""

    except:
        pass

    return pd.to_datetime(
        fecha
    ).strftime(
        "%d/%m/%Y %H:%M"
    )


# =========================================================
# FUNCIONES PARA WORD
# =========================================================

def agregar_titulo(doc, texto):

    p = doc.add_paragraph()

    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    run = p.add_run(texto)

    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(
        220,
        0,
        0
    )

    return p


def agregar_tabla_word(
    doc,
    tabla_datos,
    encabezados
):

    if tabla_datos.empty:
        return

    tabla = doc.add_table(
        rows=1,
        cols=len(encabezados)
    )

    tabla.style = "Table Grid"

    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER

    # -----------------------------------------------------
    # Encabezados
    # -----------------------------------------------------

    for i, encabezado in enumerate(encabezados):

        celda = tabla.rows[0].cells[i]

        celda.text = str(
            encabezado
        )

        celda.vertical_alignment = (
            WD_CELL_VERTICAL_ALIGNMENT.CENTER
        )

        for parrafo in celda.paragraphs:

            for run in parrafo.runs:

                run.bold = True

    # -----------------------------------------------------
    # Filas
    # -----------------------------------------------------

    for _, fila in tabla_datos.iterrows():

        celdas = tabla.add_row().cells

        for i, encabezado in enumerate(encabezados):

            if encabezado == "%":

                valor = f"{fila['%']}%"

            elif encabezado == "Cantidad":

                valor = str(
                    fila["Cantidad"]
                )

            elif encabezado == "Competidor":

                valor = str(
                    fila["Respuesta"]
                )

            elif encabezado == "Motivo":

                valor = str(
                    fila["Respuesta"]
                )

            else:

                valor = str(
                    fila.get(
                        "Respuesta",
                        ""
                    )
                )

            celdas[i].text = valor

    return tabla


def agregar_imagen_word(
    doc,
    archivo,
    ancho=5
):

    if archivo is None:
        return

    try:

        datos_imagen = archivo.getvalue()

        if not datos_imagen:
            return

        imagen = BytesIO(
            datos_imagen
        )

        doc.add_picture(
            imagen,
            width=Inches(ancho)
        )

        ultimo_parrafo = doc.paragraphs[-1]

        ultimo_parrafo.alignment = (
            WD_ALIGN_PARAGRAPH.CENTER
        )

    except Exception as e:

        doc.add_paragraph(
            f"No fue posible insertar la imagen: {e}"
        )


# =========================================================
# CREAR WORD
# =========================================================

def crear_word(
    datos,
    manuales,
    tablas
):

    doc = Document()

    # =====================================================
    # TÍTULO
    # =====================================================

    titulo = doc.add_paragraph()

    titulo.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    run = titulo.add_run(
        "OXXO | REPORTE EOD"
    )

    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(
        220,
        0,
        0
    )

    subtitulo = doc.add_paragraph()

    subtitulo.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    run = subtitulo.add_run(
        "ESTUDIO ORIGEN – DESTINO"
    )

    run.bold = True
    run.font.size = Pt(11)

    tienda = doc.add_paragraph()

    tienda.alignment = (
        WD_ALIGN_PARAGRAPH.CENTER
    )

    run = tienda.add_run(
        str(datos["tienda"])
    )

    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(
        220,
        0,
        0
    )

    # =====================================================
    # FOTO DE FACHADA
    # =====================================================

    if manuales.get("fachada"):

        agregar_titulo(
            doc,
            "FOTO DE FACHADA"
        )

        agregar_imagen_word(
            doc,
            manuales["fachada"],
            ancho=5
        )

    # =====================================================
    # 1. INFORMACIÓN GENERAL
    # =====================================================

    agregar_titulo(
        doc,
        "1. INFORMACIÓN GENERAL"
    )

    doc.add_paragraph(
        f"Periodo del estudio: {datos['periodo']}"
    )

    doc.add_paragraph(
        f"Inicio: {datos['inicio']}"
    )

    doc.add_paragraph(
        f"Finalización: {datos['finalizacion']}"
    )

    doc.add_paragraph(
        f"N.º de encuestas: {datos['encuestas']}"
    )

    doc.add_paragraph(
        f"Día con más encuestas: "
        f"{datos['dia_mas_encuestas']}"
    )

    doc.add_paragraph(
        f"Cantidad ese día: "
        f"{datos['cantidad_dia_mas']}"
    )

    # =====================================================
    # 2. PERFIL DEL CLIENTE
    # =====================================================

    agregar_titulo(
        doc,
        "2. PERFIL DEL CLIENTE"
    )

    doc.add_paragraph(
        f"Edad promedio: "
        f"{datos['edad_promedio']} años"
    )

    doc.add_paragraph(
        f"Hombres: "
        f"{datos['hombres']}"
    )

    doc.add_paragraph(
        f"Mujeres: "
        f"{datos['mujeres']}"
    )

    doc.add_paragraph(
        f"Estrato principal: "
        f"{datos['estrato']}"
    )

    doc.add_paragraph(
        f"Ocupación principal: "
        f"{datos['ocupacion']}"
    )

    # =====================================================
    # 3. PERCEPCIÓN DEL SERVICIO
    # =====================================================

    agregar_titulo(
        doc,
        "3. PERCEPCIÓN DEL SERVICIO"
    )

    doc.add_paragraph(
        manuales.get(
            "percepcion",
            ""
        )
    )

    # =====================================================
    # 4. PRINCIPAL MOTIVO DE COMPRA
    # =====================================================

    agregar_titulo(
        doc,
        "4. PRINCIPAL MOTIVO DE COMPRA"
    )

    agregar_tabla_word(
        doc,
        tablas["motivo"],
        [
            "Motivo",
            "Cantidad",
            "%"
        ]
    )

    # =====================================================
    # 5. MEDIO DE LLEGADA
    # =====================================================

    agregar_titulo(
        doc,
        "5. MEDIO DE LLEGADA"
    )

    agregar_tabla_word(
        doc,
        tablas["transporte"],
        [
            "Respuesta",
            "Cantidad",
            "%"
        ]
    )

    # =====================================================
    # 6. ORIGEN – DESTINO
    # =====================================================

    agregar_titulo(
        doc,
        "6. ORIGEN – DESTINO"
    )

    doc.add_paragraph(
        "De dónde viene:"
    )

    agregar_tabla_word(
        doc,
        tablas["origen"],
        [
            "Respuesta",
            "Cantidad",
            "%"
        ]
    )

    doc.add_paragraph(
        "Hacia dónde se dirige:"
    )

    agregar_tabla_word(
        doc,
        tablas["destino"],
        [
            "Respuesta",
            "Cantidad",
            "%"
        ]
    )

    # =====================================================
    # 7. RADIO DE INFLUENCIA
    # =====================================================

    agregar_titulo(
        doc,
        "7. RADIO DE INFLUENCIA"
    )

    doc.add_paragraph(
        manuales.get(
            "radio",
            ""
        )
    )

    # Imagen del radio
    if manuales.get("imagen_radio"):

        doc.add_paragraph(
            "Imagen del radio de influencia:"
        )

        agregar_imagen_word(
            doc,
            manuales["imagen_radio"],
            ancho=6
        )

    # =====================================================
    # 8. ALTERNATIVA DE COMPRA
    # =====================================================

    agregar_titulo(
        doc,
        "8. ALTERNATIVA DE COMPRA"
    )

    if not tablas["alternativa"].empty:

        agregar_tabla_word(
            doc,
            tablas["alternativa"],
            [
                "Competidor",
                "Cantidad",
                "%"
            ]
        )

    else:

        doc.add_paragraph(
            "No se encontraron respuestas "
            "para esta pregunta."
        )

    # Imagen de isócrona
    if manuales.get("imagen_isocrona"):

        doc.add_paragraph(
            "Isócrona:"
        )

        agregar_imagen_word(
            doc,
            manuales["imagen_isocrona"],
            ancho=6
        )

    # =====================================================
    # 9. INSIGHTS CLAVE
    # =====================================================

    agregar_titulo(
        doc,
        "9. INSIGHTS CLAVE"
    )

    if manuales.get("insight1"):

        doc.add_paragraph(
            f"1. {manuales['insight1']}"
        )

    if manuales.get("insight2"):

        doc.add_paragraph(
            f"2. {manuales['insight2']}"
        )

    if manuales.get("insight3"):

        doc.add_paragraph(
            f"3. {manuales['insight3']}"
        )

    # =====================================================
    # 10. REGISTRO FOTOGRÁFICO
    # =====================================================

    agregar_titulo(
        doc,
        "10. REGISTRO FOTOGRÁFICO"
    )

    fotos = manuales.get(
        "fotos"
    )

    if fotos:

        for i, foto in enumerate(
            fotos,
            start=1
        ):

            doc.add_paragraph(
                f"Foto {i}"
            )

            agregar_imagen_word(
                doc,
                foto,
                ancho=5
            )

    else:

        doc.add_paragraph(
            "No se cargaron fotografías."
        )

    # =====================================================
    # GUARDAR
    # =====================================================

    buffer = BytesIO()

    doc.save(
        buffer
    )

    buffer.seek(0)

    return buffer


# =========================================================
# 1. CARGAR ARCHIVO CSV
# =========================================================

st.header(
    "1. Cargar archivo de encuestas"
)

archivo = st.file_uploader(
    "Selecciona el archivo CSV",
    type=["csv"]
)

if archivo is None:

    st.info(
        "Primero carga el archivo EOD_OXXO."
    )

    st.stop()


# =========================================================
# LEER CSV
# =========================================================

try:

    df = pd.read_csv(
        archivo,
        encoding="utf-8"
    )

except:

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
    f"Archivo cargado correctamente: "
    f"{len(df)} encuestas."
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

# =========================================================
# ALTERNATIVA DE COMPRA
# =========================================================

col_alternativa = buscar_columna(
    df,
    [
        "Dónde compraría sino es en OXXO?",
        "Dónde compraría si no es en OXXO?",
        "Donde compraria sino es en OXXO?",
        "Donde compraria si no es en OXXO?",
        "Dónde compraría si no fuera en OXXO?",
        "Donde compraria si no fuera en OXXO?"
    ]
)


# =========================================================
# VALIDACIÓN TIENDA
# =========================================================

if col_tienda is None:

    st.error(
        "No encontré la columna "
        "'Nombre tienda estudiada'."
    )

    st.write(
        "Columnas encontradas:"
    )

    st.write(
        list(df.columns)
    )

    st.stop()


# =========================================================
# AVISO ALTERNATIVA
# =========================================================

if col_alternativa is not None:

    st.success(
        f"✓ Columna de alternativa detectada: "
        f"{col_alternativa}"
    )

else:

    st.warning(
        "No se encontró automáticamente la columna "
        "de alternativa de compra."
    )


# =========================================================
# 2. SELECCIONAR TIENDA
# =========================================================

st.header(
    "2. Seleccionar tienda"
)

tiendas = (
    df[col_tienda]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
)

tiendas = sorted(
    tiendas
)


tienda_seleccionada = st.selectbox(
    "Selecciona el CR / tienda",
    tiendas
)


# =========================================================
# FOTO DE FACHADA
# =========================================================

fachada = st.file_uploader(
    "Foto de fachada de la tienda",
    type=[
        "jpg",
        "jpeg",
        "png"
    ],
    key="fachada"
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

                dia.strftime(
                    "%d/%m/%Y"
                )

                for dia in todos_los_dias

                if dia not in conteo_dias.index

            ]


# =========================================================
# 3. INFORMACIÓN AUTOMÁTICA
# =========================================================

st.header(
    "3. Información automática"
)

total_encuestas = len(
    datos_tienda
)


if (
    col_fecha is not None
    and fecha_inicio is not None
):

    if (
        fecha_inicio.date()
        ==
        fecha_final.date()
    ):

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


col1, col2, col3 = st.columns(3)


with col1:

    st.metric(
        "Encuestas",
        total_encuestas
    )


with col2:

    st.metric(
        "Inicio",
        formato_fecha_hora(
            fecha_inicio
        )
    )


with col3:

    st.metric(
        "Finalización",
        formato_fecha_hora(
            fecha_final
        )
    )


col4, col5, col6 = st.columns(3)


with col4:

    st.metric(
        "Día con más encuestas",
        str(
            dia_mas_encuestas
        )
    )


with col5:

    st.metric(
        "Cantidad ese día",
        cantidad_dia_mas
    )


with col6:

    st.write(
        "**Periodo**"
    )

    st.write(
        periodo
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
# 4. PERFIL DEL CLIENTE
# =========================================================

st.header(
    "4. Perfil del cliente"
)

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
            estrato_tabla.iloc[0][
                "Respuesta"
            ]
        )

        estrato_porcentaje = (
            estrato_tabla.iloc[0][
                "%"
            ]
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
# 5. INFORMACIÓN DEL ESTUDIO
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


# =========================================================
# ALTERNATIVA DE COMPRA
# =========================================================

st.subheader(
    "Alternativa de compra"
)

st.dataframe(
    tablas["alternativa"],
    use_container_width=True,
    hide_index=True
)


# =========================================================
# 6. INFORMACIÓN MANUAL
# =========================================================

st.header(
    "6. Información manual"
)

st.info(
    "Esta información la diligencias tú "
    "porque forma parte del análisis."
)


percepcion = st.text_area(
    "Percepción del servicio",
    placeholder=(
        "Escribe aquí la percepción del servicio..."
    )
)


radio = st.text_input(
    "Radio de influencia",
    placeholder=(
        "Ejemplo: 100m: 47 | "
        "200m: 48 | "
        "300m: 74 | "
        "+300m: 67"
    )
)


# =========================================================
# IMAGEN RADIO
# =========================================================

imagen_radio = st.file_uploader(
    "Pantallazo del radio de influencia",
    type=[
        "jpg",
        "jpeg",
        "png"
    ],
    key="imagen_radio"
)


# =========================================================
# ISÓCRONA
# =========================================================

imagen_isocrona = st.file_uploader(
    "Pantallazo de la isócrona",
    type=[
        "jpg",
        "jpeg",
        "png"
    ],
    key="imagen_isocrona"
)


# =========================================================
# INSIGHTS
# =========================================================

st.subheader(
    "Insights clave"
)


insight1 = st.text_area(
    "Insight 1",
    placeholder=(
        "Escribe el primer insight..."
    )
)


insight2 = st.text_area(
    "Insight 2",
    placeholder=(
        "Escribe el segundo insight..."
    )
)


insight3 = st.text_area(
    "Insight 3",
    placeholder=(
        "Escribe el tercer insight..."
    )
)


# =========================================================
# 7. REGISTRO FOTOGRÁFICO
# =========================================================

st.header(
    "7. Registro fotográfico"
)


fotos = st.file_uploader(
    "Sube las fotografías de la tienda",
    type=[
        "jpg",
        "jpeg",
        "png"
    ],
    accept_multiple_files=True,
    key="registro_fotografico"
)


# =========================================================
# PREPARAR DATOS
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

    "edad_promedio":
        edad_promedio,

    "hombres":
        hombres,

    "mujeres":
        mujeres,

    "estrato":
        (
            estrato_tabla.iloc[0][
                "Respuesta"
            ]
            if not estrato_tabla.empty
            else "N/D"
        ),

    "ocupacion":
        (
            ocupacion_tabla.iloc[0][
                "Respuesta"
            ]
            if not ocupacion_tabla.empty
            else "N/D"
        )
}


manuales = {

    "percepcion":
        percepcion,

    "radio":
        radio,

    "imagen_radio":
        imagen_radio,

    "imagen_isocrona":
        imagen_isocrona,

    "insight1":
        insight1,

    "insight2":
        insight2,

    "insight3":
        insight3,

    "fachada":
        fachada,

    "fotos":
        fotos
}


# =========================================================
# 8. GENERAR REPORTE
# =========================================================

st.header(
    "8. Generar reporte"
)


if st.button(
    "🟢 GENERAR REPORTE",
    type="primary"
):

    try:

        documento = crear_word(
            datos,
            manuales,
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
            label="📥 Descargar reporte Word",
            data=documento,
            file_name=nombre_archivo,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            )
        )

    except Exception as e:

        st.error(
            "Ocurrió un error al generar el reporte."
        )

        st.exception(e)
