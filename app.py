import streamlit as st
import pandas as pd
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
    "Carga el archivo de encuestas y selecciona la tienda "
    "para generar la información del estudio."
)

# =========================================================
# FUNCIONES
# =========================================================

def buscar_columna(df, nombres):
    """
    Busca una columna aunque tenga diferencias de mayúsculas,
    espacios o caracteres.
    """
    columnas = {
        str(c).strip().lower().replace(" ", "").replace("_", ""): c
        for c in df.columns
    }

    for nombre in nombres:
        clave = nombre.strip().lower().replace(" ", "").replace("_", "")
        if clave in columnas:
            return columnas[clave]

    return None


def porcentaje(cantidad, total):
    if total == 0:
        return 0
    return round((cantidad / total) * 100)


def tabla_frecuencia(df, columna):
    if columna is None or columna not in df.columns:
        return pd.DataFrame()

    datos = (
        df[columna]
        .dropna()
        .astype(str)
        .str.strip()
    )

    datos = datos[datos != ""]

    if len(datos) == 0:
        return pd.DataFrame()

    tabla = datos.value_counts().reset_index()
    tabla.columns = ["Respuesta", "Cantidad"]

    tabla["%"] = tabla["Cantidad"].apply(
        lambda x: porcentaje(x, len(datos))
    )

    return tabla


def formato_fecha(fecha):
    if pd.isna(fecha):
        return ""

    return pd.to_datetime(fecha).strftime("%d/%m/%Y")


def formato_fecha_hora(fecha):
    if pd.isna(fecha):
        return ""

    return pd.to_datetime(fecha).strftime("%d/%m/%Y %H:%M")


def crear_word(datos, manuales, tablas):
    """
    Genera el Word con la información del estudio.
    Incluye imágenes cargadas manualmente para:
    - Radio de influencia
    - Isócrona
    """

    doc = Document()

    # =====================================================
    # ENCABEZADO
    # =====================================================

    doc.add_heading("REPORTE EOD", level=1)

    doc.add_paragraph(
        f"Tienda: {datos['tienda']}"
    )

    doc.add_paragraph(
        "Estudio Origen – Destino"
    )

    doc.add_paragraph(
        f"Periodo: {datos['periodo']}"
    )

    doc.add_paragraph(
        f"N. de encuestas: {datos['encuestas']}"
    )

    doc.add_paragraph(
        f"Inicio: {datos['inicio']}"
    )

    doc.add_paragraph(
        f"Finalización: {datos['finalizacion']}"
    )

    doc.add_paragraph(
        f"Día con más encuestas: {datos['dia_mas_encuestas']}"
    )

    doc.add_paragraph(
        f"Cantidad ese día: {datos['cantidad_dia_mas']}"
    )

    # =====================================================
    # PERFIL DEL CLIENTE
    # =====================================================

    doc.add_heading(
        "PERFIL DEL CLIENTE",
        level=2
    )

    doc.add_paragraph(
        f"Edad promedio: {datos['edad_promedio']} años"
    )

    doc.add_paragraph(
        f"Hombre: {datos['hombres']}"
    )

    doc.add_paragraph(
        f"Mujer: {datos['mujeres']}"
    )

    doc.add_paragraph(
        f"Estrato principal: {datos['estrato']}"
    )

    doc.add_paragraph(
        f"Ocupación principal: {datos['ocupacion']}"
    )

    # =====================================================
    # PRINCIPAL MOTIVO DE COMPRA
    # =====================================================

    doc.add_heading(
        "PRINCIPAL MOTIVO DE COMPRA",
        level=2
    )

    if not tablas["motivo"].empty:

        tabla = doc.add_table(
            rows=1,
            cols=3
        )

        tabla.style = "Table Grid"

        tabla.rows[0].cells[0].text = "Motivo"
        tabla.rows[0].cells[1].text = "Cantidad"
        tabla.rows[0].cells[2].text = "%"

        for _, fila in tablas["motivo"].iterrows():

            celdas = tabla.add_row().cells

            celdas[0].text = str(fila["Respuesta"])
            celdas[1].text = str(fila["Cantidad"])
            celdas[2].text = f"{fila['%']}%"

    # =====================================================
    # MEDIO DE LLEGADA
    # =====================================================

    doc.add_heading(
        "MEDIO DE LLEGADA",
        level=2
    )

    if not tablas["transporte"].empty:

        for _, fila in tablas["transporte"].iterrows():

            doc.add_paragraph(
                f"{fila['Respuesta']}: "
                f"{fila['Cantidad']} "
                f"({fila['%']}%)"
            )

    # =====================================================
    # PERCEPCIÓN MANUAL
    # =====================================================

    doc.add_heading(
        "PERCEPCIÓN DEL SERVICIO",
        level=2
    )

    doc.add_paragraph(
        manuales["percepcion"]
    )

    # =====================================================
    # ORIGEN - DESTINO
    # =====================================================

    doc.add_heading(
        "ORIGEN – DESTINO",
        level=2
    )

    if not tablas["origen"].empty:

        doc.add_paragraph(
            "De dónde viene:"
        )

        for _, fila in tablas["origen"].iterrows():

            doc.add_paragraph(
                f"{fila['Respuesta']}: "
                f"{fila['Cantidad']} "
                f"({fila['%']}%)"
            )

    if not tablas["destino"].empty:

        doc.add_paragraph(
            "A dónde se dirige:"
        )

        for _, fila in tablas["destino"].iterrows():

            doc.add_paragraph(
                f"{fila['Respuesta']}: "
                f"{fila['Cantidad']} "
                f"({fila['%']}%)"
            )

    # =====================================================
    # RADIO DE INFLUENCIA
    # =====================================================

    doc.add_heading(
        "RADIO DE INFLUENCIA",
        level=2
    )

    doc.add_paragraph(
        manuales["radio"]
    )

    # Imagen del radio
    if manuales["imagen_radio"]:

        doc.add_picture(
            manuales["imagen_radio"],
            width=Inches(6)
        )

    # =====================================================
    # ALTERNATIVA DE COMPRA
    # =====================================================

    doc.add_heading(
        "ALTERNATIVA DE COMPRA",
        level=2
    )

    if not tablas["alternativa"].empty:

        tabla = doc.add_table(
            rows=1,
            cols=3
        )

        tabla.style = "Table Grid"

        tabla.rows[0].cells[0].text = "Competidor"
        tabla.rows[0].cells[1].text = "Respuestas"
        tabla.rows[0].cells[2].text = "%"

        for _, fila in tablas["alternativa"].iterrows():

            celdas = tabla.add_row().cells

            celdas[0].text = str(fila["Respuesta"])
            celdas[1].text = str(fila["Cantidad"])
            celdas[2].text = f"{fila['%']}%"

    # =====================================================
    # ISÓCRONA
    # =====================================================

    doc.add_heading(
        "ISÓCRONA DE INFLUENCIA",
        level=2
    )

    if manuales["imagen_isocrona"]:

        doc.add_picture(
            manuales["imagen_isocrona"],
            width=Inches(6)
        )

    # =====================================================
    # INSIGHTS CLAVE
    # =====================================================

    doc.add_heading(
        "INSIGHTS CLAVE",
        level=2
    )

    doc.add_paragraph(
        f"1. {manuales['insight1']}"
    )

    doc.add_paragraph(
        f"2. {manuales['insight2']}"
    )

    doc.add_paragraph(
        f"3. {manuales['insight3']}"
    )

    buffer = BytesIO()

    doc.save(buffer)

    buffer.seek(0)

    return buffer


# =========================================================
# CARGAR ARCHIVO
# =========================================================

st.header("1. Cargar archivo de encuestas")

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
# LEER ARCHIVO
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

col_frecuencia = buscar_columna(
    df,
    [
        "Cada cuanto visita la tienda a la semana?"
    ]
)


# =========================================================
# VALIDACIÓN
# =========================================================

if col_tienda is None:

    st.error(
        "No encontré la columna 'Nombre tienda estudiada'."
    )

    st.write(
        "Columnas encontradas:"
    )

    st.write(
        list(df.columns)
    )

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


# =========================================================
# FILTRAR
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
        str(dia_mas_encuestas)
    )


with col5:

    st.metric(
        "Cantidad ese día",
        cantidad_dia_mas
    )


with col6:

    st.write("**Periodo**")

    st.write(
        periodo
    )


if dias_sin_encuestas:

    st.warning(
        "Días sin encuestas: "
        +
        ", ".join(dias_sin_encuestas)
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
        f"{edad_promedio} años"
        if edad_promedio != ""
        else "N/D"
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
# TABLAS
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
    "Esta información la diligencias tú "
    "porque forma parte del análisis."
)


percepcion = st.text_area(
    "Percepción del servicio",
    placeholder=(
        "Escribe aquí la percepción "
        "del servicio..."
    )
)


radio = st.text_input(
    "Radio de influencia",
    placeholder=(
        "Ejemplo: "
        "100m: 47 | "
        "200m: 48 | "
        "300m: 74 | "
        "+300m: 67"
    )
)


# =========================================================
# IMAGEN DEL RADIO
# =========================================================

imagen_radio = st.file_uploader(
    "Sube el pantallazo del radio de influencia",
    type=["jpg", "jpeg", "png"],
    key="imagen_radio"
)


# =========================================================
# IMAGEN DE ISÓCRONA
# =========================================================

st.subheader(
    "Isócrona de influencia"
)

imagen_isocrona = st.file_uploader(
    "Sube el pantallazo de la isócrona",
    type=["jpg", "jpeg", "png"],
    key="imagen_isocrona"
)


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
# FOTOGRAFÍAS
# =========================================================

st.subheader(
    "Registro fotográfico"
)


fotos = st.file_uploader(
    "Sube las fotografías de la tienda",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
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
        str(dia_mas_encuestas),

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
            estrato_tabla.iloc[0]["Respuesta"]
            if not estrato_tabla.empty
            else "N/D"
        ),

    "ocupacion":
        (
            ocupacion_tabla.iloc[0]["Respuesta"]
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

    "fotos":
        fotos

}


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
