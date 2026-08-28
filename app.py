import streamlit as st
import pandas as pd
import re
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED

st.set_page_config(
    page_title="Generador EOD OXXO",
    page_icon="🟢",
    layout="wide"
)

st.title("🟢 Generador de Reporte EOD OXXO")
st.write("Carga el CSV y la plantilla Word para generar el informe.")

# =========================================================
# FUNCIONES
# =========================================================

def normalizar(texto):
    texto = str(texto).lower().strip()
    reemplazos = {
        "á": "a", "é": "e", "í": "i",
        "ó": "o", "ú": "u", "ñ": "n"
    }
    for a, b in reemplazos.items():
        texto = texto.replace(a, b)
    return re.sub(r"[^a-z0-9]", "", texto)


def buscar_columna(df, nombres):
    columnas = {normalizar(c): c for c in df.columns}
    for nombre in nombres:
        clave = normalizar(nombre)
        if clave in columnas:
            return columnas[clave]
    return None


def buscar_parcial(df, palabras):
    for columna in df.columns:
        texto = normalizar(columna)
        if all(normalizar(p) in texto for p in palabras):
            return columna
    return None


def frecuencia(df, columna):
    if not columna:
        return pd.DataFrame(columns=["Respuesta", "Cantidad", "%"])

    serie = df[columna].dropna().astype(str).str.strip()
    serie = serie[serie != ""]

    if serie.empty:
        return pd.DataFrame(columns=["Respuesta", "Cantidad", "%"])

    tabla = serie.value_counts().reset_index()
    tabla.columns = ["Respuesta", "Cantidad"]
    tabla["%"] = (tabla["Cantidad"] / len(serie) * 100).round().astype(int)
    return tabla


def convertir_numero(valor):
    if pd.isna(valor):
        return None

    numeros = re.findall(r"\d+(?:[.,]\d+)?", str(valor))
    if not numeros:
        return None

    try:
        return float(numeros[0].replace(",", "."))
    except:
        return None


def encontrar_trafico(df):
    prioridades = [
        ["flujo", "personas"],
        ["flujo"],
        ["trafico"],
        ["tráfico"],
        ["personas"],
        ["conteo"],
        ["aforo"],
    ]

    for grupo in prioridades:
        col = buscar_parcial(df, grupo)
        if col:
            return col

    return None


def calcular_trafico(df, columna):
    if not columna:
        return None

    valores = df[columna].apply(convertir_numero).dropna()

    if valores.empty:
        return None

    return round(valores.mean(), 1)


def reemplazar_xml(data, reemplazos):
    texto = data.decode("utf-8")
    for marcador, valor in reemplazos.items():
        texto = texto.replace(marcador, str(valor))
    return texto.encode("utf-8")


def generar_word(plantilla_bytes, reemplazos):
    entrada = BytesIO(plantilla_bytes)
    salida = BytesIO()

    with ZipFile(entrada, "r") as zin, ZipFile(
        salida, "w", ZIP_DEFLATED
    ) as zout:

        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename.endswith(".xml"):
                try:
                    data = reemplazar_xml(data, reemplazos)
                except Exception:
                    pass

            zout.writestr(item, data)

    salida.seek(0)
    return salida


# =========================================================
# ARCHIVOS
# =========================================================

st.header("1. Archivos")

csv_file = st.file_uploader(
    "Archivo CSV de encuestas",
    type=["csv"]
)

plantilla_file = st.file_uploader(
    "Plantilla EOD OXXO",
    type=["docx"]
)

if csv_file is None or plantilla_file is None:
    st.info("Carga ambos archivos para continuar.")
    st.stop()


# =========================================================
# LEER CSV
# =========================================================

try:
    csv_file.seek(0)
    df = pd.read_csv(csv_file, encoding="utf-8")
except:
    csv_file.seek(0)
    df = pd.read_csv(csv_file, encoding="latin-1")

df.columns = [str(c).strip() for c in df.columns]

st.success(f"CSV cargado: {len(df)} encuestas.")


# =========================================================
# COLUMNAS
# =========================================================

col_fecha = buscar_columna(
    df,
    ["CreationDate", "Creation Date", "Fecha", "Fecha de creación"]
)

col_tienda = buscar_columna(
    df,
    ["Nombre tienda estudiada", "Nombre de tienda estudiada"]
)

col_edad = buscar_columna(df, ["Edad"])

col_genero = buscar_columna(
    df,
    ["Género", "Genero"]
)

col_estrato = buscar_columna(df, ["Estrato"])

col_ocupacion = buscar_columna(
    df,
    ["Ocupación actual", "Ocupacion actual"]
)

col_motivo = buscar_columna(
    df,
    ["Por qué compraría ahí?", "Por que compraria ahi?"]
)

col_transporte = buscar_columna(
    df,
    ["Medio de transporte usado para llegar a OXXO"]
)

col_origen = buscar_columna(
    df,
    ["De dónde viene?", "De donde viene?"]
)

col_destino = buscar_columna(
    df,
    ["Hacia dónde se dirige?", "Hacia donde se dirige?"]
)

col_alternativa = buscar_columna(
    df,
    [
        "Dónde compraría sino es en OXXO?",
        "Donde compraria sino es en OXXO?"
    ]
)

col_trafico = encontrar_trafico(df)


# =========================================================
# SELECCIÓN TIENDA
# =========================================================

if col_tienda is None:
    st.error("No encontré 'Nombre tienda estudiada'.")
    st.write(list(df.columns))
    st.stop()

st.header("2. Seleccionar CR / tienda")

tiendas = sorted(
    df[col_tienda]
    .dropna()
    .astype(str)
    .str.strip()
    .unique()
)

tienda = st.selectbox(
    "Selecciona la tienda",
    tiendas
)

datos = df[
    df[col_tienda].astype(str).str.strip() == tienda
].copy()


# =========================================================
# FECHAS
# =========================================================

fecha_inicio = None
fecha_final = None
dia_mas = ""
cantidad_dia_mas = 0
dias_sin = []

if col_fecha:

    datos[col_fecha] = pd.to_datetime(
        datos[col_fecha],
        errors="coerce"
    )

    fechas = datos[col_fecha].dropna()

    if not fechas.empty:

        fecha_inicio = fechas.min()
        fecha_final = fechas.max()

        conteo = (
            fechas.dt.date
            .value_counts()
            .sort_index()
        )

        if not conteo.empty:

            dia_mas = conteo.idxmax()
            cantidad_dia_mas = int(conteo.max())

            rango = pd.date_range(
                fecha_inicio.date(),
                fecha_final.date(),
                freq="D"
            ).date

            dias_sin = [
                d.strftime("%d/%m/%Y")
                for d in rango
                if d not in conteo.index
            ]


def fecha_hora(valor):
    if valor is None or pd.isna(valor):
        return ""
    return pd.to_datetime(valor).strftime("%d/%m/%Y %H:%M")


def solo_fecha(valor):
    if valor is None or pd.isna(valor):
        return ""
    return pd.to_datetime(valor).strftime("%d/%m/%Y")


if fecha_inicio is not None and fecha_final is not None:

    if fecha_inicio.date() == fecha_final.date():
        periodo = solo_fecha(fecha_inicio)
    else:
        periodo = (
            f"{solo_fecha(fecha_inicio)} - "
            f"{solo_fecha(fecha_final)}"
        )
else:
    periodo = ""


# =========================================================
# AUTOMÁTICOS
# =========================================================

total_encuestas = len(datos)

trafico_promedio = calcular_trafico(
    datos,
    col_trafico
)

st.header("3. Información automática")

a, b, c, d = st.columns(4)

a.metric("Encuestas", total_encuestas)

b.metric(
    "Tráfico promedio",
    (
        f"{trafico_promedio} personas"
        if trafico_promedio is not None
        else "N/D"
    )
)

c.metric("Inicio", fecha_hora(fecha_inicio))
d.metric("Finalización", fecha_hora(fecha_final))

st.write(f"**Periodo:** {periodo}")

e, f = st.columns(2)

e.metric("Día con más encuestas", str(dia_mas))
f.metric("Cantidad ese día", cantidad_dia_mas)

if dias_sin:
    st.warning(
        "Días sin encuestas: " +
        ", ".join(dias_sin)
    )


# =========================================================
# PERFIL
# =========================================================

st.header("4. Perfil del cliente")

edad_promedio = ""

if col_edad:
    edades = pd.to_numeric(
        datos[col_edad],
        errors="coerce"
    ).dropna()

    if not edades.empty:
        edad_promedio = round(edades.mean(), 1)

hombres = 0
mujeres = 0

if col_genero:

    genero = (
        datos[col_genero]
        .astype(str)
        .str.lower()
        .str.strip()
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

estrato_tabla = frecuencia(
    datos,
    col_estrato
)

ocupacion_tabla = frecuencia(
    datos,
    col_ocupacion
)

p1, p2, p3, p4 = st.columns(4)

p1.metric(
    "Edad promedio",
    f"{edad_promedio} años"
    if edad_promedio != ""
    else "N/D"
)

p2.metric("Hombres", hombres)
p3.metric("Mujeres", mujeres)

if not estrato_tabla.empty:

    p4.metric(
        "Estrato principal",
        (
            f"{estrato_tabla.iloc[0]['Respuesta']} "
            f"({estrato_tabla.iloc[0]['%']}%)"
        )
    )
else:
    p4.metric("Estrato principal", "N/D")


# =========================================================
# TABLAS
# =========================================================

st.header("5. Información del estudio")

motivo = frecuencia(datos, col_motivo)
transporte = frecuencia(datos, col_transporte)
origen = frecuencia(datos, col_origen)
destino = frecuencia(datos, col_destino)
alternativa = frecuencia(datos, col_alternativa)

x1, x2 = st.columns(2)

with x1:
    st.subheader("Principal motivo de compra")
    st.dataframe(
        motivo,
        use_container_width=True,
        hide_index=True
    )

with x2:
    st.subheader("Medio de llegada")
    st.dataframe(
        transporte,
        use_container_width=True,
        hide_index=True
    )

x3, x4 = st.columns(2)

with x3:
    st.subheader("Origen")
    st.dataframe(
        origen,
        use_container_width=True,
        hide_index=True
    )

with x4:
    st.subheader("Destino")
    st.dataframe(
        destino,
        use_container_width=True,
        hide_index=True
    )

st.subheader("Alternativa de compra")

st.dataframe(
    alternativa,
    use_container_width=True,
    hide_index=True
)


# =========================================================
# MANUALES
# =========================================================

st.header("6. Información manual")

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

insight1 = st.text_area("Insight 1")
insight2 = st.text_area("Insight 2")
insight3 = st.text_area("Insight 3")

observaciones = st.text_area(
    "Observaciones"
)

st.info(
    "Las imágenes del radio, isócrona, fachada y registro "
    "fotográfico se agregan manualmente en Word."
)


# =========================================================
# REEMPLAZOS DE LA PLANTILLA
# =========================================================

estrato = (
    estrato_tabla.iloc[0]["Respuesta"]
    if not estrato_tabla.empty
    else "N/D"
)

estrato_pct = (
    estrato_tabla.iloc[0]["%"]
    if not estrato_tabla.empty
    else "N/D"
)

ocupacion = (
    ocupacion_tabla.iloc[0]["Respuesta"]
    if not ocupacion_tabla.empty
    else "N/D"
)

ocupacion_pct = (
    ocupacion_tabla.iloc[0]["%"]
    if not ocupacion_tabla.empty
    else "N/D"
)

dia_mas_texto = (
    dia_mas.strftime("%d/%m/%Y")
    if hasattr(dia_mas, "strftime")
    else str(dia_mas)
)

reemplazos = {

    "[NOMBRE TIENDA / CR]":
        tienda,

    "[NOMBRE TIENDA]":
        tienda,

    "[CR]":
        tienda,

    "[PERIODO]":
        periodo,

    "[AUTOMÁTICO]":
        total_encuestas,

    "[ENCUESTAS]":
        total_encuestas,

    "[TRAFICO PROMEDIO]":
        (
            f"{trafico_promedio} personas"
            if trafico_promedio is not None
            else "N/D"
        ),

    "[INICIO]":
        fecha_hora(fecha_inicio),

    "[FINALIZACION]":
        fecha_hora(fecha_final),

    "[FECHA / HORA]":
        fecha_hora(fecha_inicio),

    "[EDAD]":
        edad_promedio,

    "[CANTIDAD HOMBRES]":
        hombres,

    "[CANTIDAD MUJERES]":
        mujeres,

    "[ESTRATO]":
        estrato,

    "[PORCENTAJE]":
        estrato_pct,

    "[OCUPACIÓN]":
        ocupacion,

    "[PORCENTAJE OCUPACION]":
        ocupacion_pct,

    "[FECHA]":
        dia_mas_texto,

    "[CANTIDAD]":
        cantidad_dia_mas,

    "[DIAS_SIN_ENCUESTAS]":
        (
            ", ".join(dias_sin)
            if dias_sin
            else "Ninguno"
        ),

    "[RADIO 100]":
        radio100,

    "[RADIO 200]":
        radio200,

    "[RADIO 300]":
        radio300,

    "[RADIO +300]":
        radio300mas,

    "[PERCEPCIÓN]":
        percepcion,

    "[PERCEPCION]":
        percepcion,

    "[INSIGHT 1]":
        insight1,

    "[INSIGHT 2]":
        insight2,

    "[INSIGHT 3]":
        insight3,

    "[OBSERVACIONES / INFORMACIÓN ADICIONAL]":
        observaciones
}


# =========================================================
# GENERAR REPORTE
# =========================================================

st.header("7. Generar reporte")

if st.button(
    "🟢 GENERAR REPORTE",
    type="primary"
):

    plantilla_file.seek(0)

    plantilla_bytes = plantilla_file.read()

    reporte = generar_word(
        plantilla_bytes,
        reemplazos
    )

    nombre = (
        f"Reporte_EOD_{tienda}.docx"
    )

    st.success(
        "¡Reporte generado correctamente! 🎉"
    )

    st.download_button(
        label="📥 Descargar informe EOD",
        data=reporte,
        file_name=nombre,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        )
    )
