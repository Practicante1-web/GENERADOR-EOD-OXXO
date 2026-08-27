import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn
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
st.write("Carga el CSV de encuestas y la Plantilla EOD para generar el reporte.")

# =========================================================
# FUNCIONES
# =========================================================
def buscar_columna(df, nombres):
    columnas = {
        str(c).strip().lower().replace(" ", "").replace("_", ""): c
        for c in df.columns
    }
    for nombre in nombres:
        clave = nombre.strip().lower().replace(" ", "").replace("_", "")
        if clave in columnas:
            return columnas[clave]
    return None


def buscar_columna_parcial(df, palabras):
    for c in df.columns:
        texto = str(c).strip().lower()
        if all(p.lower() in texto for p in palabras):
            return c
    return None


def porcentaje(cantidad, total):
    return round((cantidad / total) * 100) if total else 0


def tabla_frecuencia(df, columna):
    if columna is None or columna not in df.columns:
        return pd.DataFrame(columns=["Respuesta", "Cantidad", "%"])

    datos = df[columna].dropna().astype(str).str.strip()
    datos = datos[datos != ""]
    if len(datos) == 0:
        return pd.DataFrame(columns=["Respuesta", "Cantidad", "%"])

    tabla = datos.value_counts().reset_index()
    tabla.columns = ["Respuesta", "Cantidad"]
    tabla["%"] = tabla["Cantidad"].apply(lambda x: porcentaje(x, len(datos)))
    return tabla


def formato_fecha(fecha):
    if pd.isna(fecha) or fecha is None:
        return ""
    return pd.to_datetime(fecha).strftime("%d/%m/%Y")


def formato_fecha_hora(fecha):
    if pd.isna(fecha) or fecha is None:
        return ""
    return pd.to_datetime(fecha).strftime("%d/%m/%Y %H:%M")


def parse_minutos(valor):
    if pd.isna(valor):
        return None
    s = str(valor).strip().lower().replace(",", ".")
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    return float(m.group(1)) if m else None


def calcular_trafico_promedio(df, columna):
    if columna is None:
        return ""
    valores = df[columna].apply(parse_minutos).dropna()
    if len(valores) == 0:
        return ""
    return f"{round(valores.mean())} min"


def reemplazar_en_parrafo(parrafo, reemplazos):
    texto = "".join(run.text for run in parrafo.runs)
    if not texto:
        return

    nuevo = texto
    for clave, valor in reemplazos.items():
        nuevo = nuevo.replace(clave, str(valor))

    if nuevo == texto:
        return

    if parrafo.runs:
        parrafo.runs[0].text = nuevo
        for run in parrafo.runs[1:]:
            run.text = ""


def reemplazar_en_tabla(tabla, reemplazos):
    for fila in tabla.rows:
        for celda in fila.cells:
            for parrafo in celda.paragraphs:
                reemplazar_en_parrafo(parrafo, reemplazos)


def reemplazar_documento(doc, reemplazos):
    for parrafo in doc.paragraphs:
        reemplazar_en_parrafo(parrafo, reemplazos)

    for tabla in doc.tables:
        reemplazar_en_tabla(tabla, reemplazos)

    for seccion in doc.sections:
        for header in [seccion.header, seccion.first_page_header]:
            for p in header.paragraphs:
                reemplazar_en_parrafo(p, reemplazos)
            for t in header.tables:
                reemplazar_en_tabla(t, reemplazos)
        for footer in [seccion.footer, seccion.first_page_footer]:
            for p in footer.paragraphs:
                reemplazar_en_parrafo(p, reemplazos)
            for t in footer.tables:
                reemplazar_en_tabla(t, reemplazos)


def reemplazar_celda(celda, texto):
    # Mantiene el estilo del primer run de la celda.
    p = celda.paragraphs[0]
    if p.runs:
        p.runs[0].text = texto
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(texto)

    # Limpia párrafos adicionales.
    for p_extra in celda.paragraphs[1:]:
        for r in p_extra.runs:
            r.text = ""


def llenar_tabla_frecuencia(tabla, datos, columnas=(0, 1, 2), max_filas=None):
    filas = tabla.rows
    inicio = 1  # primera fila: encabezados

    cantidad_filas_datos = len(filas) - inicio
    limite = cantidad_filas_datos if max_filas is None else min(cantidad_filas_datos, max_filas)

    for i in range(limite):
        fila = filas[inicio + i]
        if i < len(datos):
            item = datos.iloc[i]
            reemplazar_celda(fila.cells[columnas[0]], str(item["Respuesta"]))
            reemplazar_celda(fila.cells[columnas[1]], str(item["Cantidad"]))
            reemplazar_celda(fila.cells[columnas[2]], f'{item["%"]}%')
        else:
            for c in columnas:
                reemplazar_celda(fila.cells[c], "")


def encontrar_tabla_por_texto(doc, texto):
    texto = texto.lower()
    for tabla in doc.tables:
        contenido = " ".join(
            celda.text.lower()
            for fila in tabla.rows
            for celda in fila.cells
        )
        if texto in contenido:
            return tabla
    return None


def reemplazar_imagen_despues_de_titulo(doc, titulo, imagen_bytes):
    """Reemplaza la primera imagen del párrafo inmediatamente posterior al título."""
    if not imagen_bytes:
        return False

    body = doc.element.body
    elementos = list(body.iterchildren())

    for i, elem in enumerate(elementos):
        if elem.tag == qn("w:p"):
            p = Paragraph(elem, doc)
            if titulo.lower() in p.text.lower():
                for siguiente in elementos[i + 1:]:
                    if siguiente.tag == qn("w:p"):
                        p2 = Paragraph(siguiente, doc)
                        blips = p2._p.xpath(".//a:blip")
                        if blips:
                            rid = blips[0].get(qn("r:embed"))
                            part = doc.part.related_parts.get(rid)
                            if part is not None and hasattr(part, "_blob"):
                                part._blob = image_bytes
                                return True
                        # Si encontramos texto antes de una imagen, seguimos buscando.
                    elif siguiente.tag == qn("w:tbl"):
                        continue
    return False


def insertar_imagen_en_lugar_de_parrafo(doc, marcador, imagen_bytes, ancho=2.2):
    if not imagen_bytes:
        return False

    encontrado = False
    for parrafo in doc.paragraphs:
        if marcador.lower() in parrafo.text.lower():
            parrafo.text = ""
            run = parrafo.add_run()
            run.add_picture(BytesIO(imagen_bytes), width=Inches(ancho))
            encontrado = True

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    if marcador.lower() in parrafo.text.lower():
                        parrafo.text = ""
                        run = parrafo.add_run()
                        run.add_picture(BytesIO(imagen_bytes), width=Inches(ancho))
                        encontrado = True
    return encontrado


def agregar_foto_en_celda(celda, imagen_bytes, ancho=1.8):
    if not imagen_bytes:
        return
    p = celda.paragraphs[0]
    p.text = ""
    run = p.add_run()
    run.add_picture(BytesIO(imagen_bytes), width=Inches(ancho))


def llenar_fotos_registro(doc, fotos):
    if not fotos:
        return

    # Busca la sección REGISTRO FOTOGRÁFICO y utiliza los
    # cuadros/espacios de las tablas cercanas cuando existen.
    for tabla in doc.tables:
        texto = " ".join(c.text.lower() for row in tabla.rows for c in row.cells)
        if "registro fotográfico" in texto:
            # Si la tabla tiene celdas vacías, las utiliza.
            celdas = [c for row in tabla.rows for c in row.cells]
            vacias = [c for c in celdas if not c.text.strip()]
            for i, foto in enumerate(fotos[:3]):
                if i < len(vacias):
                    agregar_foto_en_celda(vacias[i], foto, 1.7)
            return

    # Alternativa: insertar después del texto del título.
    for p in doc.paragraphs:
        if p.text.strip().lower() == "registro fotográfico":
            for foto in fotos[:3]:
                run = p.add_run()
                run.add_picture(BytesIO(foto), width=Inches(1.7))
                p.add_run("   ")
            return


def crear_reporte(template_bytes, datos, tablas, manuales, imagenes):
    doc = Document(BytesIO(template_bytes))

    reemplazos = {
        "[NOMBRE TIENDA / CR]": datos["tienda"],
        "[PERIODO]": datos["periodo"],
        "[NUM_ENCUESTAS]": datos["encuestas"],
        "[AUTOMATICO]": datos["encuestas"],
        "[FECHA_INICIO]": datos["fecha_inicio"],
        "[HORA_INICIO]": datos["hora_inicio"],
        "[FECHA_FINALIZACION]": datos["fecha_finalizacion"],
        "[HORA_FINALIZACION]": datos["hora_finalizacion"],
        "[DIA_MAS_ENCUESTAS]": datos["dia_mas_encuestas"],
        "[DIAS_SIN_ENCUESTAS]": datos["dias_sin_encuestas"],

        "[EDAD]": datos["edad_promedio"],
        "[CANTIDAD HOMBRE]": datos["hombres"],
        "[CANTIDAD HOMBRE]": datos["hombres"],
        "[CANTIDAD MUJER]": datos["mujeres"],
        "[ESTRATO]": datos["estrato"],
        "[ESTRATO_PORCENTAJE]": datos["estrato_porcentaje"],
        "[OCUPACIÓN]": datos["ocupacion"],
        "[OCUPACION_PORCENTAJE]": datos["ocupacion_porcentaje"],
        "[PERCEPCIÓN]": manuales["percepcion"],
        "[COMENTARIO PERCEPCIÓN]": manuales["percepcion_comentario"],
        "[TRAFICO_PROMEDIO]": datos["trafico_promedio"],
        "[TRÁFICO PROMEDIO]": datos["trafico_promedio"],
    }

    # Datos de motivo de compra.
    for i in range(4):
        if i < len(tablas["motivo"]):
            r = tablas["motivo"].iloc[i]
            reemplazos[f"[MOTIVO {i+1}]"] = r["Respuesta"]
            reemplazos[f"[CANTIDAD {i+1}]"] = r["Cantidad"]
            reemplazos[f"[PORCENTAJE {i+1}]"] = f'{r["%"]}'
        else:
            reemplazos[f"[MOTIVO {i+1}]"] = ""
            reemplazos[f"[CANTIDAD {i+1}]"] = ""
            reemplazos[f"[PORCENTAJE {i+1}]"] = ""

    # Medio de llegada.
    nombres_transporte = ["A pie", "Moto", "Automóvil", "Otro"]
    for i, nombre in enumerate(nombres_transporte, start=1):
        coincidencia = tablas["transporte"][
            tablas["transporte"]["Respuesta"].str.lower().str.contains(
                nombre.lower().replace("ó", "o"), regex=False
            )
        ] if not tablas["transporte"].empty else pd.DataFrame()

        if not coincidencia.empty:
            r = coincidencia.iloc[0]
            valor = f'{r["Cantidad"]}'
        else:
            valor = ""
        reemplazos[f"[{nombre.upper()}]"] = valor

    # Origen y destino se rellenan directamente en su tabla para evitar
    # conflictos entre [CANTIDAD] y [%] de origen y destino.

    # Radio: solo los números son manuales; 100m/200m/300m/+300m permanecen en plantilla.
    reemplazos["[RADIO_100]"] = manuales["radio_100"]
    reemplazos["[RADIO_200]"] = manuales["radio_200"]
    reemplazos["[RADIO_300]"] = manuales["radio_300"]
    reemplazos["[RADIO_MAS_300]"] = manuales["radio_mas_300"]

    reemplazos["[INSIGHTS 1]"] = manuales["insight1"]
    reemplazos["[INSIGHTS 2]"] = manuales["insight2"]
    reemplazos["[INSIGHTS 3]"] = manuales["insight3"]

    reemplazar_documento(doc, reemplazos)

    # Tabla de alternativa de compra.
    tabla_alt = encontrar_tabla_por_texto(doc, "competidor")
    if tabla_alt is not None and not tablas["alternativa"].empty:
        llenar_tabla_frecuencia(tabla_alt, tablas["alternativa"], max_filas=5)

    # Tabla origen-destino: sustituye las filas de datos.
    tabla_od = encontrar_tabla_por_texto(doc, "de dónde viene")
    if tabla_od is not None:
        origen = tablas["origen"].head(4)
        destino = tablas["destino"].head(4)
        for i in range(4):
            if i + 1 >= len(tabla_od.rows):
                break
            fila = tabla_od.rows[i + 1]
            if i < len(origen):
                r = origen.iloc[i]
                reemplazar_celda(fila.cells[0], f'{r["Respuesta"]}: {r["Cantidad"]} ({r["%"]}%)')
            else:
                reemplazar_celda(fila.cells[0], "")
            if i < len(destino):
                r = destino.iloc[i]
                reemplazar_celda(fila.cells[1], f'{r["Respuesta"]}: {r["Cantidad"]} ({r["%"]}%)')
            else:
                reemplazar_celda(fila.cells[1], "")

    # Fotografías: fachada y mapas.
    if imagenes.get("fachada"):
        insertar_imagen_en_lugar_de_parrafo(doc, "Foto fachada", imagenes["fachada"], 2.8)

    if imagenes.get("radio"):
        reemplazar_imagen_despues_de_titulo(doc, "RADIO DE INFLUENCIA", imagenes["radio"])

    if imagenes.get("isocrona"):
        reemplazar_imagen_despues_de_titulo(doc, "ISÓCRONA DE INFLUENCIA", imagenes["isocrona"])

    llenar_fotos_registro(doc, imagenes.get("fotos", []))

    # Elimina el marcador de comentario de percepción si no fue usado.
    reemplazar_documento(doc, {
        "[PERCEPCION_COMENTARIO]": manuales["percepcion_comentario"],
        "[COMENTARIO PERCEPCIÓN]": manuales["percepcion_comentario"],
    })

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# =========================================================
# CARGA DE ARCHIVOS
# =========================================================
st.header("1. Archivos")

archivo_csv = st.file_uploader("Selecciona el archivo CSV de encuestas", type=["csv"])
archivo_plantilla = st.file_uploader(
    "Selecciona la Plantilla EOD (la que preparaste en Word)",
    type=["docx"]
)

if archivo_csv is None or archivo_plantilla is None:
    st.info("Carga los dos archivos para continuar.")
    st.stop()

try:
    df = pd.read_csv(archivo_csv, encoding="utf-8")
except Exception:
    archivo_csv.seek(0)
    df = pd.read_csv(archivo_csv, encoding="latin-1")

df.columns = [str(c).strip() for c in df.columns]

st.success(f"CSV cargado: {len(df)} registros.")

# =========================================================
# COLUMNAS
# =========================================================
col_fecha = buscar_columna(df, ["CreationDate", "Creation Date", "Fecha"])
col_tienda = buscar_columna(df, ["Nombre tienda estudiada"])
col_edad = buscar_columna(df, ["Edad"])
col_genero = buscar_columna(df, ["Género", "Genero"])
col_estrato = buscar_columna(df, ["Estrato"])
col_ocupacion = buscar_columna(df, ["Ocupación actual", "Ocupacion actual"])
col_motivo = buscar_columna(df, ["Por qué compraría ahí?"])
col_transporte = buscar_columna(df, ["Medio de transporte usado para llegar a OXXO"])
col_origen = buscar_columna(df, ["De dónde viene?"])
col_destino = buscar_columna(df, ["Hacia dónde se dirige?"])
col_alternativa = buscar_columna(df, ["Dónde compraría sino es en OXXO?"])

# Busca una columna de tráfico/tiempo de llegada de forma flexible.
col_trafico = (
    buscar_columna_parcial(df, ["trafico", "promedio"])
    or buscar_columna_parcial(df, ["tiempo", "llegada"])
    or buscar_columna_parcial(df, ["tiempo", "traslado"])
)

if col_tienda is None:
    st.error("No encontré la columna 'Nombre tienda estudiada'.")
    st.write(list(df.columns))
    st.stop()

# =========================================================
# TIENDA
# =========================================================
st.header("2. Seleccionar tienda")

tiendas = sorted(
    df[col_tienda].dropna().astype(str).str.strip().unique()
)
tienda = st.selectbox("Selecciona el CR / tienda", tiendas)

datos_tienda = df[
    df[col_tienda].astype(str).str.strip() == tienda
].copy()

# =========================================================
# FECHAS
# =========================================================
fecha_inicio = fecha_final = None
dia_mas_encuestas = ""
cantidad_dia_mas = 0
dias_sin_encuestas = []

if col_fecha:
    datos_tienda[col_fecha] = pd.to_datetime(
        datos_tienda[col_fecha], errors="coerce"
    )
    fechas = datos_tienda[col_fecha].dropna()

    if len(fechas):
        fecha_inicio = fechas.min()
        fecha_final = fechas.max()

        conteo = fechas.dt.date.value_counts().sort_index()
        dia_mas_encuestas = str(conteo.idxmax())
        cantidad_dia_mas = int(conteo.max())

        todos = pd.date_range(
            fecha_inicio.date(), fecha_final.date(), freq="D"
        ).date
        dias_sin_encuestas = [
            d.strftime("%d/%m/%Y")
            for d in todos if d not in conteo.index
        ]

periodo = (
    formato_fecha(fecha_inicio)
    if fecha_inicio is not None and fecha_inicio.date() == fecha_final.date()
    else f"{formato_fecha(fecha_inicio)} - {formato_fecha(fecha_final)}"
)

# =========================================================
# AUTOMÁTICOS
# =========================================================
total = len(datos_tienda)

edad_promedio = ""
if col_edad:
    edades = pd.to_numeric(datos_tienda[col_edad], errors="coerce").dropna()
    if len(edades):
        edad_promedio = round(edades.mean(), 1)

hombres = mujeres = 0
if col_genero:
    g = datos_tienda[col_genero].astype(str).str.strip().str.lower()
    hombres = int(g.str.contains("hombre", na=False).sum())
    mujeres = int(g.str.contains("mujer", na=False).sum())

estrato_tabla = tabla_frecuencia(datos_tienda, col_estrato)
ocupacion_tabla = tabla_frecuencia(datos_tienda, col_ocupacion)

estrato = str(estrato_tabla.iloc[0]["Respuesta"]) if not estrato_tabla.empty else ""
estrato_pct = str(estrato_tabla.iloc[0]["%"]) if not estrato_tabla.empty else ""

ocupacion = str(ocupacion_tabla.iloc[0]["Respuesta"]) if not ocupacion_tabla.empty else ""
ocupacion_pct = str(ocupacion_tabla.iloc[0]["%"]) if not ocupacion_tabla.empty else ""

trafico_promedio = calcular_trafico_promedio(datos_tienda, col_trafico)

tablas = {
    "motivo": tabla_frecuencia(datos_tienda, col_motivo),
    "transporte": tabla_frecuencia(datos_tienda, col_transporte),
    "origen": tabla_frecuencia(datos_tienda, col_origen),
    "destino": tabla_frecuencia(datos_tienda, col_destino),
    "alternativa": tabla_frecuencia(datos_tienda, col_alternativa),
}

# =========================================================
# DATOS MANUALES
# =========================================================
st.header("3. Información manual")

c1, c2, c3, c4 = st.columns(4)
with c1:
    radio_100 = st.text_input("100 m", placeholder="Ej. 47")
with c2:
    radio_200 = st.text_input("200 m", placeholder="Ej. 48")
with c3:
    radio_300 = st.text_input("300 m", placeholder="Ej. 74")
with c4:
    radio_mas_300 = st.text_input("+300 m", placeholder="Ej. 67")

percepcion_comentario = st.text_area(
    "Comentario de percepción del servicio",
    placeholder="Escribe el motivo/explicación de la percepción."
)

insight1 = st.text_area("Insight 1")
insight2 = st.text_area("Insight 2")
insight3 = st.text_area("Insight 3")

st.caption("La imagen de radios, isócrona y registro fotográfico se cargan manualmente.")

fachada = st.file_uploader("Foto fachada", type=["jpg", "jpeg", "png"], key="fachada")
imagen_radio = st.file_uploader("Imagen del radio de influencia", type=["jpg", "jpeg", "png"], key="radio")
imagen_isocrona = st.file_uploader("Imagen de la isócrona", type=["jpg", "jpeg", "png"], key="isocrona")
fotos = st.file_uploader(
    "Registro fotográfico (hasta 3 fotos)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="fotos"
)

# =========================================================
# RESUMEN
# =========================================================
st.header("4. Resumen automático")

a, b, c, d = st.columns(4)
a.metric("Encuestas", total)
b.metric("Inicio", formato_fecha_hora(fecha_inicio))
c.metric("Finalización", formato_fecha_hora(fecha_final))
d.metric("Día con más encuestas", dia_mas_encuestas)

if dias_sin_encuestas:
    st.warning("Días sin encuestas: " + ", ".join(dias_sin_encuestas))

if not trafico_promedio:
    st.info(
        "No encontré una columna identificable para calcular automáticamente "
        "el tráfico promedio. El resto del reporte sí puede generarse."
    )

# =========================================================
# GENERAR
# =========================================================
st.header("5. Generar reporte")

if st.button("🟢 GENERAR REPORTE", type="primary"):
    datos = {
        "tienda": tienda,
        "periodo": periodo,
        "encuestas": total,
        "fecha_inicio": formato_fecha(fecha_inicio),
        "hora_inicio": pd.to_datetime(fecha_inicio).strftime("%H:%M") if fecha_inicio is not None else "",
        "fecha_finalizacion": formato_fecha(fecha_final),
        "hora_finalizacion": pd.to_datetime(fecha_final).strftime("%H:%M") if fecha_final is not None else "",
        "dia_mas_encuestas": dia_mas_encuestas,
        "dias_sin_encuestas": ", ".join(dias_sin_encuestas) if dias_sin_encuestas else "Ninguno",
        "edad_promedio": edad_promedio,
        "hombres": hombres,
        "mujeres": mujeres,
        "estrato": estrato,
        "estrato_porcentaje": estrato_pct,
        "ocupacion": ocupacion,
        "ocupacion_porcentaje": ocupacion_pct,
        "trafico_promedio": trafico_promedio,
    }

    manuales = {
        "radio_100": radio_100,
        "radio_200": radio_200,
        "radio_300": radio_300,
        "radio_mas_300": radio_mas_300,
        "percepcion": percepcion_comentario,
        "percepcion_comentario": percepcion_comentario,
        "insight1": insight1,
        "insight2": insight2,
        "insight3": insight3,
    }

    imagenes = {
        "fachada": fachada.getvalue() if fachada else None,
        "radio": imagen_radio.getvalue() if imagen_radio else None,
        "isocrona": imagen_isocrona.getvalue() if imagen_isocrona else None,
        "fotos": [f.getvalue() for f in (fotos or [])],
    }

    documento = crear_reporte(
        archivo_plantilla.getvalue(),
        datos,
        tablas,
        manuales,
        imagenes
    )

    nombre = f"Reporte_EOD_{tienda.replace('/', '-')}.docx"

    st.success("¡Reporte generado correctamente! 🎉")
    st.download_button(
        "📥 Descargar reporte EOD",
        data=documento,
        file_name=nombre,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
