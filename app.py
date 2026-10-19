import unicodedata
from io import BytesIO
from datetime import datetime
from xml.sax.saxutils import escape

import altair as alt
import pandas as pd
import streamlit as st
from PIL import Image as PILImage, ImageOps

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import (
    BaseDocTemplate,
    Flowable,
    Frame,
    Image as RLImage,
    KeepTogether,
    NextPageTemplate,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


# =========================================================
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Generador EOD OXXO",
    page_icon="🟥",
    layout="wide"
)

# =========================================================
# COLORES (inspirados en OXXO)
# =========================================================

HEX_ROJO = "#DA291C"
HEX_AMARILLO = "#FFC72C"

ROJO = colors.HexColor(HEX_ROJO)
AMARILLO = colors.HexColor(HEX_AMARILLO)
AMARILLO_SUAVE = colors.HexColor("#FFF4D6")
TEXTO = colors.HexColor("#2B2B2B")
GRIS = colors.HexColor("#6B6B6B")
GRIS_SUAVE = colors.HexColor("#F5F5F5")
GRIS_LINEA = colors.HexColor("#E3E3E3")


# =========================================================
# FUNCIONES DE DATOS
# =========================================================

def normalizar_texto(texto):
    """Quita mayúsculas, espacios y signos para comparar nombres de columnas."""
    if texto is None:
        return ""
    texto = str(texto).strip().lower()
    for caracter in [" ", "_", "-", "?", "¿", "!", "¡", ".", ",",
                     ";", ":", "(", ")", "[", "]"]:
        texto = texto.replace(caracter, "")
    return texto


def buscar_columna(df, nombres):
    columnas = {normalizar_texto(c): c for c in df.columns}
    for nombre in nombres:
        clave = normalizar_texto(nombre)
        if clave in columnas:
            return columnas[clave]
    return None


def porcentaje(cantidad, total):
    if total == 0:
        return 0
    return round((cantidad / total) * 100)


def bonito(valor):
    """'TRANSPORTE PÚBLICO' -> 'Transporte público'."""
    valor = str(valor).strip()
    if valor.isupper():
        return valor.capitalize()
    return valor


def limpiar_serie(serie):
    datos = serie.dropna().astype(str).str.strip()
    return datos[(datos != "") & (datos.str.lower() != "nan")]


def tabla_desde_conteo(conteo):
    total = conteo.sum()
    tabla = conteo.reset_index()
    tabla.columns = ["Respuesta", "Cantidad"]
    tabla["Respuesta"] = tabla["Respuesta"].astype(str)
    tabla["%"] = tabla["Cantidad"].apply(lambda x: porcentaje(x, total))
    return tabla


def tabla_vacia():
    return pd.DataFrame(columns=["Respuesta", "Cantidad", "%"])


def tabla_frecuencia(df, columna, orden=None):
    """
    Cuenta las respuestas de una columna.
    Si se pasa 'orden', respeta ese orden (útil para escalas como 1..5).
    """
    if columna is None or columna not in df.columns:
        return tabla_vacia()

    datos = limpiar_serie(df[columna])
    if len(datos) == 0:
        return tabla_vacia()

    conteo = datos.value_counts()

    if orden is not None:
        conteo = conteo.reindex(
            [o for o in orden if o in conteo.index]
            + [i for i in conteo.index if i not in orden]
        )

    conteo.index = [bonito(i) for i in conteo.index]
    conteo = conteo.groupby(level=0, sort=False).sum()

    return tabla_desde_conteo(conteo)


# ---------------------------------------------------------
# Competidores: unifica nombres escritos de forma distinta
# (D1 / d1 / "D1 ", Carrulla / Carulla, Tostao / Tostado...)
# ---------------------------------------------------------

ALIAS_COMPETIDORES = {
    "carrulla": "carulla",
    "tostao": "tostado",
    "farmato": "farmatodo",
    "fsrmatodo": "farmatodo",
    "ambulante": "ambulantes",
    "tiendita de barrio": "tienda de barrio",
    "tienda barrio": "tienda de barrio",
    "olimpica": "olimpica",
}


def clave_competidor(texto):
    t = unicodedata.normalize("NFKD", str(texto))
    t = t.encode("ascii", "ignore").decode().lower()
    t = t.strip(" ´'`.\"-")
    t = " ".join(t.split())
    return ALIAS_COMPETIDORES.get(t, t)


def tabla_competidores(df, columna):
    if columna is None or columna not in df.columns:
        return tabla_vacia()

    datos = limpiar_serie(df[columna]).str.strip(" ´'`")
    if len(datos) == 0:
        return tabla_vacia()

    claves = datos.apply(clave_competidor)

    # Para cada grupo se muestra la forma de escribirlo más frecuente
    nombres = (
        pd.DataFrame({"clave": claves, "original": datos})
        .groupby("clave")["original"]
        .agg(lambda s: s.value_counts().index[0])
    )

    conteo = claves.value_counts()
    conteo.index = [bonito(nombres[c]) for c in conteo.index]

    return tabla_desde_conteo(conteo)


def tabla_rangos_edad(edades):
    edades = edades.dropna()
    if len(edades) == 0:
        return tabla_vacia()

    rangos = pd.cut(
        edades,
        bins=[0, 17, 24, 34, 44, 54, 200],
        labels=["<18", "18-24", "25-34", "35-44", "45-54", "55+"],
    )
    conteo = rangos.value_counts().sort_index()
    conteo = conteo[conteo > 0]
    conteo.index = conteo.index.astype(str)

    return tabla_desde_conteo(conteo)


def formato_fecha(fecha):
    if fecha is None or pd.isna(fecha):
        return ""
    return pd.to_datetime(fecha).strftime("%d/%m/%Y")


def formato_fecha_hora(fecha):
    if fecha is None or pd.isna(fecha):
        return ""
    return pd.to_datetime(fecha).strftime("%d/%m/%Y %H:%M")


def leer_fechas(serie, convertir_utc):
    try:
        fechas = pd.to_datetime(serie, format="%m/%d/%Y %I:%M:%S %p", errors="coerce")
        if fechas.notna().sum() == 0:
            raise ValueError
    except Exception:
        fechas = pd.to_datetime(serie, errors="coerce")

    if convertir_utc:
        fechas = (
            fechas.dt.tz_localize("UTC")
            .dt.tz_convert("America/Bogota")
            .dt.tz_localize(None)
        )
    return fechas


def agrupar_otros(tabla, max_filas=8):
    """Deja las respuestas principales y agrupa el resto en 'Otros'."""
    if len(tabla) <= max_filas:
        return tabla
    principales = tabla.head(max_filas - 1)
    resto = tabla.iloc[max_filas - 1:]
    otros = pd.DataFrame([{
        "Respuesta": f"Otros ({len(resto)})",
        "Cantidad": int(resto["Cantidad"].sum()),
        "%": int(resto["%"].sum()),
    }])
    return pd.concat([principales, otros], ignore_index=True)


# =========================================================
# GRÁFICOS PARA LA APP (Altair)
# =========================================================

def grafico_barras_app(tabla, horizontal=True, max_filas=8, altura=None):
    if tabla is None or tabla.empty:
        st.caption("Sin respuestas registradas.")
        return

    d = agrupar_otros(tabla, max_filas) if horizontal else tabla.copy()
    d = d.reset_index(drop=True)
    d["Destacado"] = d["Cantidad"] == d["Cantidad"].max()
    d["Etiqueta"] = d["%"].astype(int).astype(str) + "% (" + d["Cantidad"].astype(int).astype(str) + ")"
    orden = list(d["Respuesta"])

    color = alt.condition(
        "datum.Destacado",
        alt.value(HEX_ROJO),
        alt.value(HEX_AMARILLO),
    )
    tooltip = ["Respuesta", "Cantidad", alt.Tooltip("%", title="%")]

    if horizontal:
        base = alt.Chart(d).encode(
            y=alt.Y("Respuesta:N", sort=orden, title=None),
            x=alt.X("Cantidad:Q", title=None, axis=None),
        )
        barras = base.mark_bar(cornerRadiusEnd=4).encode(color=color, tooltip=tooltip)
        textos = base.mark_text(align="left", dx=4, fontSize=12).encode(text="Etiqueta:N")
        alto = altura or max(110, 30 * len(d))
    else:
        base = alt.Chart(d).encode(
            x=alt.X("Respuesta:N", sort=orden, title=None, axis=alt.Axis(labelAngle=0)),
            y=alt.Y("Cantidad:Q", title=None, axis=None),
        )
        barras = base.mark_bar(cornerRadiusEnd=4).encode(color=color, tooltip=tooltip)
        textos = base.mark_text(dy=-8, fontSize=11).encode(text="Cantidad:Q")
        alto = altura or 220

    st.altair_chart(
        (barras + textos).properties(height=alto),
        use_container_width=True,
    )


def grafico_dona_app(tabla):
    if tabla is None or tabla.empty:
        st.caption("Sin respuestas registradas.")
        return
    d = tabla.copy()
    paleta = [HEX_ROJO, HEX_AMARILLO, "#9B9B9B", "#5A5A5A"]
    grafico = alt.Chart(d).mark_arc(innerRadius=55).encode(
        theta="Cantidad:Q",
        color=alt.Color(
            "Respuesta:N",
            scale=alt.Scale(range=paleta[:len(d)]),
            legend=alt.Legend(title=None, orient="right"),
        ),
        tooltip=["Respuesta", "Cantidad", "%"],
    ).properties(height=220)
    st.altair_chart(grafico, use_container_width=True)


# =========================================================
# ESTILOS DEL PDF
# =========================================================

ANCHO_PAGINA, ALTO_PAGINA = letter
MARGEN = 1.8 * cm
ANCHO_UTIL = ANCHO_PAGINA - 2 * MARGEN
SEPARACION_COL = 0.8 * cm
ANCHO_COL = (ANCHO_UTIL - SEPARACION_COL) / 2


def estilo(nombre, **kwargs):
    base = dict(fontName="Helvetica", fontSize=9, leading=12, textColor=TEXTO)
    base.update(kwargs)
    return ParagraphStyle(nombre, **base)


E_TEXTO = estilo("texto", fontSize=9.5, leading=14)
E_SUBTITULO = estilo("sub", fontName="Helvetica-Bold", fontSize=9.5, leading=12)
E_CELDA = estilo("celda", fontSize=8.5, leading=10.5)
E_CELDA_C = estilo("celda_c", fontSize=8.5, leading=10.5, alignment=TA_CENTER)
E_ENCABEZADO = estilo("enc", fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=colors.white)
E_ENCABEZADO_C = estilo("enc_c", fontName="Helvetica-Bold", fontSize=8, leading=10,
                        textColor=colors.white, alignment=TA_CENTER)
E_NOTA = estilo("nota", fontSize=7, leading=9, textColor=GRIS)
E_KPI_VALOR = estilo("kpi_v", fontName="Helvetica-Bold", fontSize=20, leading=23, textColor=ROJO)
E_KPI_LABEL = estilo("kpi_l", fontSize=7.5, leading=9.5, textColor=GRIS)
E_DATO_LABEL = estilo("dato_l", fontSize=7, leading=9, textColor=GRIS)
E_DATO_VALOR = estilo("dato_v", fontName="Helvetica-Bold", fontSize=9.5, leading=12)
E_PIE_FOTO = estilo("pie", fontSize=7.5, leading=9, textColor=GRIS, alignment=TA_CENTER)
E_VACIO = estilo("vacio", fontName="Helvetica-Oblique", fontSize=8.5, textColor=GRIS)
E_INSIGHT = estilo("insight", fontSize=9.5, leading=13.5)
E_NUMERO = estilo("num", fontName="Helvetica-Bold", fontSize=12, leading=14,
                  textColor=colors.white, alignment=TA_CENTER)
E_CHIP_VALOR = estilo("chip_v", fontName="Helvetica-Bold", fontSize=14, leading=16,
                      textColor=ROJO, alignment=TA_CENTER)
E_CHIP_LABEL = estilo("chip_l", fontSize=7.5, leading=9, textColor=GRIS, alignment=TA_CENTER)


def texto_seguro(valor):
    if valor is None:
        return ""
    return escape(str(valor)).replace("\n", "<br/>")


def recortar(texto, fuente, tamano, ancho_max):
    texto = str(texto)
    if stringWidth(texto, fuente, tamano) <= ancho_max:
        return texto
    while texto and stringWidth(texto + "…", fuente, tamano) > ancho_max:
        texto = texto[:-1]
    return texto + "…"


# =========================================================
# GRÁFICOS DEL PDF
# =========================================================

class TituloSeccion(Flowable):
    """Título en rojo con una pequeña franja amarilla."""

    def __init__(self, texto, ancho=ANCHO_UTIL):
        super().__init__()
        self.texto = texto.upper()
        self.ancho = ancho
        self.alto = 0.95 * cm

    def wrap(self, *args):
        return self.ancho, self.alto

    def draw(self):
        c = self.canv
        c.setFillColor(ROJO)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(0, 0.38 * cm, self.texto)
        c.setFillColor(AMARILLO)
        c.rect(0, 0.12 * cm, 1.4 * cm, 0.12 * cm, stroke=0, fill=1)
        c.setStrokeColor(GRIS_LINEA)
        c.setLineWidth(0.6)
        c.line(1.55 * cm, 0.18 * cm, self.ancho, 0.18 * cm)


class BarrasVerticales(Flowable):
    """Columnas para escalas o tiempos (días, horas, edades, estratos...)."""

    def __init__(self, etiquetas, valores, ancho, alto=4.0 * cm,
                 titulo_eje="", mostrar_todos_valores=True):
        super().__init__()
        self.etiquetas = [str(e) for e in etiquetas]
        self.valores = [float(v) for v in valores]
        self.ancho = ancho
        self.alto = alto
        self.titulo_eje = titulo_eje
        self.mostrar_todos = mostrar_todos_valores

    def wrap(self, *args):
        return self.ancho, self.alto

    def draw(self):
        c = self.canv
        n = len(self.valores)
        if n == 0:
            return

        base_y = 14 if not self.titulo_eje else 22
        alto_util = self.alto - base_y - 12
        maximo = max(self.valores) or 1
        indice_max = self.valores.index(maximo)

        espacio = self.ancho / n
        ancho_barra = min(espacio * 0.64, 1.2 * cm)

        # Mostrar etiquetas intercaladas si hay muchas columnas
        cada = 1
        while n / cada > 16:
            cada += 1

        c.setStrokeColor(GRIS_LINEA)
        c.setLineWidth(0.6)
        c.line(0, base_y, self.ancho, base_y)

        for i, (etq, val) in enumerate(zip(self.etiquetas, self.valores)):
            x_centro = espacio * i + espacio / 2
            h = alto_util * val / maximo
            c.setFillColor(ROJO if i == indice_max else AMARILLO)
            c.rect(x_centro - ancho_barra / 2, base_y, ancho_barra, max(h, 1), stroke=0, fill=1)

            if self.mostrar_todos or i == indice_max:
                c.setFillColor(TEXTO)
                c.setFont("Helvetica-Bold", 7 if n > 10 else 7.5)
                c.drawCentredString(x_centro, base_y + h + 3, f"{int(val)}")

            if i % cada == 0 or i == indice_max:
                c.setFillColor(GRIS)
                c.setFont("Helvetica", 6.5 if n > 10 else 7.5)
                c.drawCentredString(
                    x_centro, base_y - 9,
                    recortar(etq, "Helvetica", 7, espacio * cada - 2)
                )

        if self.titulo_eje:
            c.setFillColor(GRIS)
            c.setFont("Helvetica-Oblique", 6.5)
            c.drawCentredString(self.ancho / 2, 2, self.titulo_eje)


class Dona(Flowable):
    """Gráfico de dona con el dato principal en el centro y leyenda."""

    PALETA = [ROJO, AMARILLO, colors.HexColor("#9B9B9B"), colors.HexColor("#5A5A5A")]

    def __init__(self, tabla, ancho, alto=3.5 * cm, centro="", sub=""):
        super().__init__()
        self.partes = [
            (str(r["Respuesta"]), int(r["Cantidad"]), int(r["%"]))
            for _, r in tabla.head(4).iterrows()
        ]
        self.ancho = ancho
        self.alto = alto
        self.centro = centro
        self.sub = sub

    def wrap(self, *args):
        return self.ancho, self.alto

    def draw(self):
        c = self.canv
        total = sum(p[1] for p in self.partes)
        if total == 0:
            return

        r = self.alto / 2 - 2
        cx, cy = r + 2, self.alto / 2

        if len(self.partes) == 1:
            c.setFillColor(self.PALETA[0])
            c.circle(cx, cy, r, stroke=0, fill=1)
        else:
            inicio = 90
            for i, (_, cant, _) in enumerate(self.partes):
                extension = -360 * cant / total
                c.setFillColor(self.PALETA[i % len(self.PALETA)])
                c.wedge(cx - r, cy - r, cx + r, cy + r, inicio, extension, stroke=0, fill=1)
                inicio += extension

        c.setFillColor(colors.white)
        c.circle(cx, cy, r * 0.6, stroke=0, fill=1)

        c.setFillColor(ROJO)
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(cx, cy - 2, self.centro)
        if self.sub:
            c.setFillColor(GRIS)
            c.setFont("Helvetica", 6.5)
            c.drawCentredString(cx, cy - 11, self.sub)

        x_ley = cx + r + 0.6 * cm
        y_ley = cy + (len(self.partes) * 20) / 2 - 14
        for i, (resp, cant, pct) in enumerate(self.partes):
            y = y_ley - i * 20
            c.setFillColor(self.PALETA[i % len(self.PALETA)])
            c.roundRect(x_ley, y, 8, 8, 1.5, stroke=0, fill=1)
            c.setFillColor(TEXTO)
            c.setFont("Helvetica-Bold", 9)
            c.drawString(x_ley + 13, y, f"{pct}%")
            c.setFont("Helvetica", 8)
            c.setFillColor(GRIS)
            c.drawString(
                x_ley + 13 + stringWidth(f"{pct}%", "Helvetica-Bold", 9) + 4, y,
                recortar(f"{resp} ({cant})", "Helvetica", 8,
                         self.ancho - x_ley - 45)
            )


# =========================================================
# BLOQUES DEL PDF
# =========================================================

def bloque(titulo, contenido, nota=None):
    """Subtítulo + gráfico (+ nota opcional)."""
    elementos = [Paragraph(texto_seguro(titulo), E_SUBTITULO), Spacer(1, 6)]
    if contenido is None:
        elementos.append(Paragraph("Sin respuestas registradas.", E_VACIO))
    else:
        elementos.append(contenido)
    if nota:
        elementos += [Spacer(1, 3), Paragraph(texto_seguro(nota), E_NOTA)]
    return elementos


def tabla_pdf(tabla, encabezado="RESPUESTA", ancho=ANCHO_COL, agrupar=True, max_filas=8):
    """
    Tabla de frecuencias: encabezado rojo, filas alternadas
    y una barrita amarilla con el porcentaje.
    """
    if tabla is None or tabla.empty:
        return None

    d = agrupar_otros(tabla, max_filas) if agrupar else tabla

    ancho_cant = 1.3 * cm
    ancho_pct = 1.1 * cm
    ancho_barra = min(3.0 * cm, ancho * 0.28)
    ancho_resp = ancho - ancho_cant - ancho_pct - ancho_barra

    filas = [[
        Paragraph(encabezado, E_ENCABEZADO),
        Paragraph("CANT.", E_ENCABEZADO_C),
        Paragraph("%", E_ENCABEZADO_C),
        "",
    ]]
    for _, f in d.iterrows():
        filas.append([
            Paragraph(texto_seguro(f["Respuesta"]), E_CELDA),
            Paragraph(str(int(f["Cantidad"])), E_CELDA_C),
            Paragraph(f"<b>{int(f['%'])}%</b>", E_CELDA_C),
            BarraPorcentaje(f["%"], ancho_barra - 0.4 * cm),
        ])

    t = Table(filas, colWidths=[ancho_resp, ancho_cant, ancho_pct, ancho_barra])
    estilos = [
        ("BACKGROUND", (0, 0), (-1, 0), ROJO),
        ("SPAN", (2, 0), (3, 0)),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, GRIS_LINEA),
    ]
    for i in range(2, len(filas), 2):
        estilos.append(("BACKGROUND", (0, i), (-1, i), GRIS_SUAVE))
    t.setStyle(TableStyle(estilos))
    return t


class BarraPorcentaje(Flowable):
    """Barrita amarilla del % dentro de las tablas."""

    def __init__(self, pct, ancho, alto=5):
        super().__init__()
        self.pct = max(0, min(100, float(pct or 0)))
        self.ancho = ancho
        self.alto = alto

    def wrap(self, *args):
        return self.ancho, self.alto

    def draw(self):
        c = self.canv
        r = self.alto / 2
        c.setFillColor(GRIS_LINEA)
        c.roundRect(0, 0, self.ancho, self.alto, r, stroke=0, fill=1)
        if self.pct > 0:
            c.setFillColor(AMARILLO)
            c.roundRect(0, 0, max(self.alto, self.ancho * self.pct / 100), self.alto, r, stroke=0, fill=1)


def barras_v(tabla, ancho=ANCHO_COL, **kw):
    if tabla is None or tabla.empty:
        return None
    return BarrasVerticales(list(tabla["Respuesta"]), list(tabla["Cantidad"]), ancho, **kw)


def dona(tabla, ancho=ANCHO_COL, **kw):
    if tabla is None or tabla.empty:
        return None
    return Dona(tabla, ancho, **kw)


def dos_columnas(izquierda, derecha):
    t = Table([[izquierda, "", derecha]], colWidths=[ANCHO_COL, SEPARACION_COL, ANCHO_COL])
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return t


def seccion(titulo, *filas):
    """Título de sección + filas de contenido, sin partirse entre páginas."""
    elementos = [TituloSeccion(titulo), Spacer(1, 8)]
    for i, fila in enumerate(filas):
        if i > 0:
            elementos.append(Spacer(1, 18))
        if isinstance(fila, list):
            elementos += fila
        else:
            elementos.append(fila)
    return [KeepTogether(elementos), Spacer(1, 24)]


def tarjetas_kpi(kpis):
    n = len(kpis)
    separacion = 0.35 * cm
    ancho = (ANCHO_UTIL - separacion * (n - 1)) / n
    celdas, anchos = [], []

    for i, (valor, etiqueta) in enumerate(kpis):
        tarjeta = Table(
            [[Paragraph(texto_seguro(valor), E_KPI_VALOR)],
             [Paragraph(texto_seguro(etiqueta), E_KPI_LABEL)]],
            colWidths=[ancho],
        )
        tarjeta.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), GRIS_SUAVE),
            ("LINEABOVE", (0, 0), (-1, 0), 3, AMARILLO),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("TOPPADDING", (0, 0), (-1, 0), 9),
            ("BOTTOMPADDING", (0, -1), (-1, -1), 9),
            ("TOPPADDING", (0, 1), (-1, 1), 0),
        ]))
        celdas.append(tarjeta)
        anchos.append(ancho)
        if i < n - 1:
            celdas.append("")
            anchos.append(separacion)

    fila = Table([celdas], colWidths=anchos)
    fila.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return fila


def fila_datos(datos_fila):
    n = len(datos_fila)
    celdas = [[
        [Paragraph(texto_seguro(etq).upper(), E_DATO_LABEL),
         Paragraph(texto_seguro(val) or "—", E_DATO_VALOR)]
        for etq, val in datos_fila
    ]]
    t = Table(celdas, colWidths=[ANCHO_UTIL / n] * n)
    t.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 0.6, GRIS_LINEA),
        ("LINEBEFORE", (1, 0), (-1, -1), 0.6, GRIS_LINEA),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return t


def caja_texto(texto, ancho=ANCHO_UTIL, fondo=AMARILLO_SUAVE, borde=AMARILLO):
    contenido = (
        Paragraph(texto_seguro(texto), E_TEXTO)
        if texto and str(texto).strip()
        else Paragraph("Sin información diligenciada.", E_VACIO)
    )
    t = Table([[contenido]], colWidths=[ancho])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), fondo),
        ("LINEBEFORE", (0, 0), (0, -1), 3, borde),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    return t


def chips_radio(texto_radio):
    """'100m: 47 | 200m: 48 | 300m: 74' -> fila de indicadores."""
    if not texto_radio or not str(texto_radio).strip():
        return None

    partes = [p.strip() for p in str(texto_radio).split("|") if p.strip()]
    pares = []
    for parte in partes:
        if ":" not in parte:
            return caja_texto(texto_radio)
        etq, val = parte.split(":", 1)
        pares.append((etq.strip(), val.strip()))

    if not pares or len(pares) > 8:
        return caja_texto(texto_radio)

    n = len(pares)
    separacion = 0.25 * cm
    ancho = (ANCHO_UTIL - separacion * (n - 1)) / n
    celdas, anchos = [], []

    for i, (etq, val) in enumerate(pares):
        chip = Table(
            [[Paragraph(texto_seguro(val), E_CHIP_VALOR)],
             [Paragraph(texto_seguro(etq), E_CHIP_LABEL)]],
            colWidths=[ancho],
        )
        chip.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.8, GRIS_LINEA),
            ("LINEBELOW", (0, -1), (-1, -1), 2.5, AMARILLO),
            ("TOPPADDING", (0, 0), (-1, 0), 7),
            ("BOTTOMPADDING", (0, -1), (-1, -1), 7),
            ("TOPPADDING", (0, 1), (-1, 1), 0),
        ]))
        celdas.append(chip)
        anchos.append(ancho)
        if i < n - 1:
            celdas.append("")
            anchos.append(separacion)

    fila = Table([celdas], colWidths=anchos)
    fila.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return fila


def preparar_imagen(archivo, max_ancho, max_alto):
    """Corrige la rotación de fotos de celular y ajusta el tamaño."""
    if archivo is None:
        return None
    try:
        archivo.seek(0)
        img = PILImage.open(archivo)
        img = ImageOps.exif_transpose(img).convert("RGB")
        img.thumbnail((1800, 1800))
        buffer = BytesIO()
        img.save(buffer, format="JPEG", quality=85)
        buffer.seek(0)
        w, h = img.size
        escala = min(max_ancho / w, max_alto / h)
        return RLImage(buffer, width=w * escala, height=h * escala)
    except Exception:
        return Paragraph("No se pudo cargar la imagen.", E_VACIO)


def imagen_con_pie(archivo, pie, max_ancho, max_alto):
    img = preparar_imagen(archivo, max_ancho, max_alto)
    if img is None:
        return None
    t = Table([[img], [Paragraph(texto_seguro(pie), E_PIE_FOTO)]], colWidths=[max_ancho])
    t.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 1), (-1, 1), 4),
    ]))
    return t


def bloque_insights(insights):
    filas = []
    for texto in insights:
        if texto and str(texto).strip():
            filas.append([
                Paragraph(str(len(filas) + 1), E_NUMERO),
                Paragraph(texto_seguro(texto), E_INSIGHT),
            ])
    if not filas:
        return None

    t = Table(filas, colWidths=[0.9 * cm, ANCHO_UTIL - 0.9 * cm])
    estilos = [
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (1, 0), (1, -1), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
        ("LINEBELOW", (0, 0), (-1, -2), 0.5, GRIS_LINEA),
    ]
    for i in range(len(filas)):
        estilos.append(("BACKGROUND", (0, i), (0, i), ROJO))
    t.setStyle(TableStyle(estilos))
    return t


def grilla_fotos(fotos):
    if not fotos:
        return None
    separacion = 0.5 * cm
    ancho = (ANCHO_UTIL - separacion) / 2
    alto = 6.6 * cm

    imagenes = [preparar_imagen(f, ancho, alto) for f in fotos[:4]]
    if len(imagenes) % 2:
        imagenes.append("")

    filas = [[imagenes[i], "", imagenes[i + 1]] for i in range(0, len(imagenes), 2)]
    t = Table(filas, colWidths=[ancho, separacion, ancho],
              rowHeights=[alto + 0.4 * cm] * len(filas))
    estilos = [
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]
    for f, fila in enumerate(filas):
        for col in (0, 2):
            if fila[col] != "":
                estilos.append(("BACKGROUND", (col, f), (col, f), GRIS_SUAVE))
    t.setStyle(TableStyle(estilos))
    return t


# =========================================================
# ENCABEZADOS Y PIE DE PÁGINA
# =========================================================

ALTO_BANDA_PORTADA = 3.5 * cm
ALTO_BANDA_INTERNA = 1.15 * cm


def dibujar_pie(c, doc):
    c.setStrokeColor(GRIS_LINEA)
    c.setLineWidth(0.6)
    c.line(MARGEN, 1.35 * cm, ANCHO_PAGINA - MARGEN, 1.35 * cm)
    c.setFont("Helvetica", 7.5)
    c.setFillColor(GRIS)
    c.drawString(MARGEN, 0.9 * cm, f"Estudio Origen-Destino · {doc.info_titulo_corto}")
    c.drawRightString(
        ANCHO_PAGINA - MARGEN, 0.9 * cm,
        f"Generado el {doc.info_generado}   ·   Página {doc.page}"
    )


def pagina_portada(c, doc):
    c.saveState()

    c.setFillColor(ROJO)
    c.rect(0, ALTO_PAGINA - ALTO_BANDA_PORTADA, ANCHO_PAGINA, ALTO_BANDA_PORTADA, stroke=0, fill=1)
    c.setFillColor(AMARILLO)
    c.rect(0, ALTO_PAGINA - ALTO_BANDA_PORTADA - 0.22 * cm, ANCHO_PAGINA, 0.22 * cm, stroke=0, fill=1)

    # Etiqueta superior
    c.setFillColor(AMARILLO)
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(MARGEN, ALTO_PAGINA - 1.0 * cm, "REPORTE EOD  ·  ESTUDIO ORIGEN-DESTINO")

    # Píldora con el CR
    texto_cr = f"CR  {doc.info_cr}"
    ancho_pill = stringWidth(texto_cr, "Helvetica-Bold", 10) + 0.7 * cm
    y_pill = ALTO_PAGINA - 1.85 * cm
    c.setFillColor(AMARILLO)
    c.roundRect(MARGEN, y_pill, ancho_pill, 0.62 * cm, 0.31 * cm, stroke=0, fill=1)
    c.setFillColor(ROJO)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(MARGEN + 0.35 * cm, y_pill + 0.2 * cm, texto_cr)

    # Nombre de la tienda
    nombre = doc.info_nombre or "Tienda estudiada"
    estilo_titulo = estilo("titulo", fontName="Helvetica-Bold", fontSize=21,
                           leading=24, textColor=colors.white)
    if len(nombre) > 34:
        estilo_titulo.fontSize = 16
        estilo_titulo.leading = 19
    p = Paragraph(texto_seguro(nombre), estilo_titulo)
    _, h = p.wrap(ANCHO_UTIL * 0.68, 2 * cm)
    p.drawOn(c, MARGEN, y_pill - 0.25 * cm - h)

    # Periodo
    if doc.info_periodo:
        c.setFillColor(colors.white)
        c.setFont("Helvetica", 7.5)
        c.drawRightString(ANCHO_PAGINA - MARGEN, ALTO_PAGINA - 1.75 * cm, "PERIODO DE MEDICIÓN")
        c.setFont("Helvetica-Bold", 10.5)
        c.drawRightString(ANCHO_PAGINA - MARGEN, ALTO_PAGINA - 2.25 * cm, doc.info_periodo)

    dibujar_pie(c, doc)
    c.restoreState()


def pagina_interna(c, doc):
    c.saveState()
    c.setFillColor(ROJO)
    c.rect(0, ALTO_PAGINA - ALTO_BANDA_INTERNA, ANCHO_PAGINA, ALTO_BANDA_INTERNA, stroke=0, fill=1)
    c.setFillColor(AMARILLO)
    c.rect(0, ALTO_PAGINA - ALTO_BANDA_INTERNA - 0.12 * cm, ANCHO_PAGINA, 0.12 * cm, stroke=0, fill=1)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(
        MARGEN, ALTO_PAGINA - 0.72 * cm,
        recortar(doc.info_titulo_corto, "Helvetica-Bold", 9.5, ANCHO_UTIL * 0.75)
    )
    c.setFont("Helvetica", 8)
    c.drawRightString(ANCHO_PAGINA - MARGEN, ALTO_PAGINA - 0.72 * cm, "REPORTE EOD")

    dibujar_pie(c, doc)
    c.restoreState()


# =========================================================
# CREAR PDF
# =========================================================

def crear_pdf(datos, manuales, tablas):

    buffer = BytesIO()

    cr = str(datos["cr"])
    nombre = str(datos["nombre"] or "").strip()

    doc = BaseDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=MARGEN, rightMargin=MARGEN,
        topMargin=MARGEN, bottomMargin=1.9 * cm,
        title=f"Reporte EOD - {cr} {nombre}".strip(),
        author="Generador EOD OXXO",
    )
    doc.info_cr = cr
    doc.info_nombre = nombre
    doc.info_titulo_corto = f"CR {cr} · {nombre}" if nombre else f"CR {cr}"
    doc.info_periodo = str(datos["periodo"] or "")
    doc.info_generado = datetime.now().strftime("%d/%m/%Y")

    pie = 1.9 * cm
    frame_portada = Frame(
        MARGEN, pie, ANCHO_UTIL,
        ALTO_PAGINA - ALTO_BANDA_PORTADA - 0.9 * cm - pie,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )
    frame_interna = Frame(
        MARGEN, pie, ANCHO_UTIL,
        ALTO_PAGINA - ALTO_BANDA_INTERNA - 0.8 * cm - pie,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
    )
    doc.addPageTemplates([
        PageTemplate(id="portada", frames=[frame_portada], onPage=pagina_portada),
        PageTemplate(id="interna", frames=[frame_interna], onPage=pagina_interna),
    ])

    h, m = datos["hombres"], datos["mujeres"]

    historia = [NextPageTemplate("interna")]

    # ---------------- Resumen ----------------
    historia.append(tarjetas_kpi([
        (datos["encuestas"], "Encuestas realizadas"),
        (f"{datos['edad_promedio']} años" if datos["edad_promedio"] != "" else "N/D",
         "Edad promedio"),
        (f"{datos['servicio_promedio']} / 5" if datos["servicio_promedio"] != "" else "N/D",
         "Calificación del servicio"),
        (f"{datos['estrato']}" if datos["estrato"] not in ("", "N/D") else "N/D",
         f"Estrato principal · {datos['estrato_porcentaje']}%"),
    ]))
    historia.append(Spacer(1, 8))

    dia_pico = datos["dia_mas_encuestas"]
    if datos["cantidad_dia_mas"]:
        dia_pico = f"{dia_pico} ({datos['cantidad_dia_mas']})"

    historia.append(fila_datos([
        ("Inicio", datos["inicio"]),
        ("Finalización", datos["finalizacion"]),
        ("Día con más encuestas", dia_pico),
        ("Hora pico", datos["hora_pico"]),
    ]))
    historia.append(Spacer(1, 16))

    # ---------------- Fachada ----------------
    fachada = imagen_con_pie(manuales["foto_fachada"], "Fachada de la tienda", ANCHO_UTIL, 7.2 * cm)
    if fachada is not None:
        historia += seccion("Fachada", fachada)

    # ---------------- Ritmo de encuestas ----------------
    historia += seccion(
        "Ritmo de las encuestas",
        dos_columnas(
            bloque("Encuestas por día", barras_v(tablas["dias"], alto=3.8 * cm)),
            bloque("Franja horaria (hora de Colombia)" if datos["hora_local"] else "Franja horaria",
                   barras_v(tablas["horas"], alto=3.8 * cm)),
        ),
    )

    # ---------------- Perfil ----------------
    total_genero = h + m
    historia += seccion(
        "Perfil del cliente",
        dos_columnas(
            bloque("Género", dona(tablas["genero"], centro=f"{total_genero}", sub="personas")),
            bloque("Rango de edad", barras_v(tablas["edades"], alto=3.4 * cm)),
        ),
        dos_columnas(
            bloque("Ocupación", tabla_pdf(tablas["ocupacion"], "OCUPACIÓN")),
            bloque("Estrato", tabla_pdf(tablas["estrato"], "ESTRATO", agrupar=False)),
        ),
    )

    # ---------------- Hábitos ----------------
    planeaba = tablas["planeaba"]
    pct_si = ""
    if not planeaba.empty:
        fila_si = planeaba[planeaba["Respuesta"].str.lower().isin(["si", "sí"])]
        pct_si = f"{int(fila_si['%'].iloc[0])}%" if not fila_si.empty else "0%"

    historia += seccion(
        "Hábitos de compra",
        dos_columnas(
            bloque("Principal motivo de compra", tabla_pdf(tablas["motivo"], "MOTIVO")),
            bloque("Medio de llegada", tabla_pdf(tablas["transporte"], "MEDIO")),
        ),
        dos_columnas(
            bloque("Visitas a la tienda por semana",
                   barras_v(tablas["frecuencia"], alto=3.4 * cm, titulo_eje="veces por semana")),
            bloque("¿Planeaba venir al OXXO?",
                   dona(planeaba, centro=pct_si, sub="sí planeaba"),
                   nota=(f"Pregunta respondida por {int(planeaba['Cantidad'].sum())} personas."
                         if not planeaba.empty else None)),
        ),
    )

    # ---------------- Origen y destino ----------------
    historia += seccion(
        "Origen y destino",
        dos_columnas(
            bloque("¿De dónde viene?", tabla_pdf(tablas["origen"], "LUGAR")),
            bloque("¿Hacia dónde se dirige?", tabla_pdf(tablas["destino"], "LUGAR")),
        ),
    )

    # ---------------- Competencia ----------------
    no_visito = tablas["no_visito"]
    historia += seccion(
        "Competencia",
        dos_columnas(
            bloque("¿Dónde compraría si no es en OXXO?",
                   tabla_pdf(tablas["alternativa"], "COMPETIDOR")),
            bloque("Si hay un OXXO más cerca, ¿por qué no lo visitó?",
                   tabla_pdf(no_visito, "MOTIVO"),
                   nota=(f"Pregunta respondida por {int(no_visito['Cantidad'].sum())} personas."
                         if not no_visito.empty else None)),
        ),
    )

    # ---------------- Servicio ----------------
    historia += seccion(
        "Servicio",
        dos_columnas(
            bloque("Calificación del servicio (1 a 5)",
                   tabla_pdf(tablas["servicio"], "CALIFICACIÓN", agrupar=False)),
            bloque("Percepción del servicio",
                   caja_texto(manuales["percepcion"], ancho=ANCHO_COL,
                              fondo=GRIS_SUAVE, borde=ROJO)),
        ),
    )

    # ---------------- Zona de influencia ----------------
    zona = []
    chips = chips_radio(manuales["radio"])
    if chips is not None:
        zona.append(bloque("Radio de influencia (encuestados por distancia)", chips))

    img_radio = imagen_con_pie(manuales["imagen_radio"], "Radio de influencia", ANCHO_COL, 7.5 * cm)
    img_iso = imagen_con_pie(manuales["imagen_isocrona"], "Isócrona", ANCHO_COL, 7.5 * cm)

    if img_radio is not None and img_iso is not None:
        zona.append(dos_columnas(img_radio, img_iso))
    elif img_radio is not None or img_iso is not None:
        zona.append(imagen_con_pie(
            manuales["imagen_radio"] or manuales["imagen_isocrona"],
            "Radio de influencia" if manuales["imagen_radio"] else "Isócrona",
            ANCHO_UTIL, 9 * cm,
        ))

    if zona:
        historia += seccion("Zona de influencia", *zona)

    # ---------------- Insights ----------------
    insights = bloque_insights([manuales["insight1"], manuales["insight2"], manuales["insight3"]])
    if insights is not None:
        historia += seccion("Insights clave", insights)

    # ---------------- Registro fotográfico ----------------
    grilla = grilla_fotos(manuales["fotos"])
    if grilla is not None:
        historia += seccion("Registro fotográfico", grilla)

    doc.build(historia)
    buffer.seek(0)
    return buffer


# =========================================================
# INTERFAZ
# =========================================================

st.title("Generador de Reporte EOD OXXO")
st.write(
    "Carga el CSV de encuestas, elige el **CR** y completa la información del análisis. "
    "El reporte se descarga en **PDF**."
)

with st.sidebar:
    st.header("Opciones")
    convertir_utc = st.checkbox(
        "Las horas del CSV vienen en UTC (Survey123): convertir a hora Colombia",
        value=True,
        help="Survey123 guarda la fecha de creación en hora UTC (5 horas más). "
             "Desmárcalo si tus horas ya están en hora local.",
    )


# =========================================================
# 1. CARGAR CSV
# =========================================================

st.header("1. Cargar archivo de encuestas")

archivo = st.file_uploader("Selecciona el CSV de encuestas", type=["csv"])

if archivo is None:
    st.info("Primero carga el archivo EOD_OXXO.")
    st.stop()

try:
    df = pd.read_csv(archivo, encoding="utf-8")
except Exception:
    archivo.seek(0)
    df = pd.read_csv(archivo, encoding="latin-1")

df.columns = [str(c).strip() for c in df.columns]


# =========================================================
# IDENTIFICAR COLUMNAS
# =========================================================

col_fecha = buscar_columna(df, ["CreationDate", "Creation Date"])
col_cr = buscar_columna(df, ["Nombre tienda estudiada", "CR", "Tienda"])
col_edad = buscar_columna(df, ["Edad"])
col_genero = buscar_columna(df, ["Género", "Genero"])
col_estrato = buscar_columna(df, ["Estrato"])
col_ocupacion = buscar_columna(df, ["Ocupación actual", "Ocupacion actual"])
col_motivo = buscar_columna(df, ["Por qué compraría ahí?", "Por que compraria ahi?"])
col_transporte = buscar_columna(df, ["Medio de transporte usado para llegar a OXXO"])
col_origen = buscar_columna(df, ["De dónde viene?", "De donde viene?"])
col_destino = buscar_columna(df, ["Hacia dónde se dirige?", "Hacia donde se dirige?"])
col_frecuencia = buscar_columna(df, ["Cada cuanto visita la tienda a la semana?"])
col_servicio = buscar_columna(df, ["Servicio"])
col_planeaba = buscar_columna(df, ["Planeaba venir al OXXO", "Tenía planeado venir al OXXO"])
col_no_visito = buscar_columna(df, ["Si existe una tienda OXXO más cerca, Por qué no la visitó?"])
col_x = buscar_columna(df, ["x", "lon", "longitud"])
col_y = buscar_columna(df, ["y", "lat", "latitud"])
col_alternativa = buscar_columna(df, [
    "Dónde compraría sino es en OXXO?",
    "Dónde compraría si no es en OXXO?",
    "Donde compraria sino es en OXXO?",
    "Donde compraria si no es en OXXO?",
    "Dónde compraría si no fuera en OXXO?",
    "Alternativa de compra",
    "Competidor",
])

if col_cr is None:
    st.error("No encontré la columna 'Nombre tienda estudiada' (CR).")
    st.write("Columnas encontradas:", list(df.columns))
    st.stop()

df[col_cr] = df[col_cr].astype(str).str.strip()

st.success(
    f"Archivo cargado: **{len(df):,} encuestas** de "
    f"**{df[col_cr].nunique()} tiendas**.".replace(",", ".")
)


# =========================================================
# 2. SELECCIONAR TIENDA
# =========================================================

st.header("2. Seleccionar tienda")

conteo_cr = df[col_cr].value_counts()
crs = sorted(conteo_cr.index)

c1, c2 = st.columns([1, 2])
with c1:
    cr_seleccionado = st.selectbox(
        "CR",
        crs,
        format_func=lambda cr: f"{cr}  ({conteo_cr[cr]} encuestas)",
    )
with c2:
    nombre_tienda = st.text_input(
        "Nombre de la tienda estudiada",
        key=f"nombre_{cr_seleccionado}",
        placeholder="Escribe el nombre de la tienda. Ej: OXXO Cedritos Calle 140",
    )

datos_tienda = df[df[col_cr] == cr_seleccionado].copy()


# ---------------------------------------------------------
# Rango de fechas: detecta encuestas sueltas de otras fechas
# (por ejemplo, 10 encuestas meses después del estudio) y las
# deja por fuera, pero se puede ajustar a mano.
# ---------------------------------------------------------

if col_fecha is not None:
    datos_tienda["_fecha"] = leer_fechas(datos_tienda[col_fecha], convertir_utc)
    dias_con_datos = sorted(datos_tienda["_fecha"].dropna().dt.date.unique())

    if dias_con_datos:
        grupos = [[dias_con_datos[0]]]
        for dia in dias_con_datos[1:]:
            if (dia - grupos[-1][-1]).days > 14:
                grupos.append([dia])
            else:
                grupos[-1].append(dia)

        conteo_por_dia = datos_tienda["_fecha"].dt.date.value_counts()
        principal = max(grupos, key=lambda g: sum(conteo_por_dia[d] for d in g))

        rango = st.date_input(
            "Rango de fechas del estudio",
            value=(principal[0], principal[-1]),
            min_value=dias_con_datos[0],
            max_value=dias_con_datos[-1],
            format="DD/MM/YYYY",
            key=f"rango_{cr_seleccionado}",
        )

        if isinstance(rango, (list, tuple)):
            inicio_rango = rango[0]
            fin_rango = rango[1] if len(rango) > 1 else rango[0]
        else:
            inicio_rango = fin_rango = rango

        solo_fecha = datos_tienda["_fecha"].dt.date
        dentro = datos_tienda["_fecha"].isna() | (
            (solo_fecha >= inicio_rango) & (solo_fecha <= fin_rango)
        )
        excluidas = int((~dentro).sum())

        if excluidas:
            fechas_fuera = sorted({d.strftime("%d/%m/%Y") for d in solo_fecha[~dentro].dropna()})
            st.warning(
                f"Se dejaron por fuera **{excluidas} encuestas** que están fuera del "
                f"periodo del estudio ({', '.join(fechas_fuera[:6])}"
                f"{'…' if len(fechas_fuera) > 6 else ''}). "
                "Si sí deben contar, amplía el rango de fechas."
            )

        datos_tienda = datos_tienda[dentro].copy()


# =========================================================
# CÁLCULOS
# =========================================================

fecha_inicio = fecha_final = None
dia_mas_encuestas, cantidad_dia_mas, hora_pico = "", 0, ""
dias_sin_encuestas = []
tabla_dias = tabla_vacia()
tabla_horas = tabla_vacia()

if col_fecha is not None and "_fecha" in datos_tienda.columns:
    fechas = datos_tienda["_fecha"].dropna()

    if len(fechas) > 0:
        fecha_inicio, fecha_final = fechas.min(), fechas.max()

        conteo_dias = fechas.dt.date.value_counts().sort_index()
        dia_max = conteo_dias.idxmax()
        dia_mas_encuestas = dia_max.strftime("%d/%m/%Y")
        cantidad_dia_mas = int(conteo_dias.max())

        todos = pd.date_range(fecha_inicio.date(), fecha_final.date(), freq="D").date
        dias_sin_encuestas = [d.strftime("%d/%m/%Y") for d in todos if d not in conteo_dias.index]

        conteo_dias.index = [d.strftime("%d/%m") for d in conteo_dias.index]
        tabla_dias = tabla_desde_conteo(conteo_dias)

        conteo_horas = fechas.dt.hour.value_counts().sort_index()
        conteo_horas = conteo_horas.reindex(
            range(conteo_horas.index.min(), conteo_horas.index.max() + 1), fill_value=0
        )
        hora_max = int(conteo_horas.idxmax())
        hora_pico = f"{hora_max}:00 - {hora_max + 1}:00"
        conteo_horas.index = [f"{h}h" for h in conteo_horas.index]
        tabla_horas = tabla_desde_conteo(conteo_horas)

if fecha_inicio is not None:
    periodo = (
        formato_fecha(fecha_inicio)
        if fecha_inicio.date() == fecha_final.date()
        else f"{formato_fecha(fecha_inicio)} - {formato_fecha(fecha_final)}"
    )
else:
    periodo = ""

total_encuestas = len(datos_tienda)

edad_promedio = ""
tabla_edades = tabla_vacia()
edades_invalidas = 0
if col_edad is not None:
    edades = pd.to_numeric(datos_tienda[col_edad], errors="coerce")
    validas = edades[(edades >= 12) & (edades <= 99)]
    edades_invalidas = int(edades.notna().sum() - len(validas))
    if len(validas):
        edad_promedio = round(validas.mean(), 1)
        tabla_edades = tabla_rangos_edad(validas)

hombres = mujeres = 0
tabla_genero = tabla_frecuencia(datos_tienda, col_genero)
if col_genero is not None:
    genero = datos_tienda[col_genero].astype(str).str.strip().str.lower()
    hombres = int(genero.str.contains("hombre", na=False).sum())
    mujeres = int(genero.str.contains("mujer", na=False).sum())

tabla_estrato = tabla_frecuencia(datos_tienda, col_estrato)
if not tabla_estrato.empty:
    estrato_principal = tabla_estrato.iloc[0]["Respuesta"]
    estrato_porcentaje = int(tabla_estrato.iloc[0]["%"])
    tabla_estrato_orden = tabla_estrato.sort_values(
        "Respuesta", key=lambda s: pd.to_numeric(s, errors="coerce")
    ).reset_index(drop=True)
else:
    estrato_principal, estrato_porcentaje = "N/D", 0
    tabla_estrato_orden = tabla_estrato

servicio_promedio = ""
tabla_servicio = tabla_vacia()
if col_servicio is not None:
    notas = pd.to_numeric(datos_tienda[col_servicio], errors="coerce").dropna()
    if len(notas):
        servicio_promedio = round(notas.mean(), 1)
        nombres_nota = {5: "5 · Excelente", 4: "4 · Bueno", 3: "3 · Regular", 2: "2 · Malo", 1: "1 · Muy malo"}
        conteo = notas.round().astype(int).value_counts().reindex([5, 4, 3, 2, 1], fill_value=0)
        conteo.index = [nombres_nota[i] for i in conteo.index]
        tabla_servicio = tabla_desde_conteo(conteo)

tablas = {
    "dias": tabla_dias,
    "horas": tabla_horas,
    "genero": tabla_genero,
    "edades": tabla_edades,
    "estrato": tabla_estrato_orden,
    "ocupacion": tabla_frecuencia(datos_tienda, col_ocupacion),
    "motivo": tabla_frecuencia(datos_tienda, col_motivo),
    "transporte": tabla_frecuencia(datos_tienda, col_transporte),
    "frecuencia": tabla_frecuencia(datos_tienda, col_frecuencia, orden=["1", "2", "3", "4", "5 o más"]),
    "planeaba": tabla_frecuencia(datos_tienda, col_planeaba, orden=["SI", "NO"]),
    "origen": tabla_frecuencia(datos_tienda, col_origen),
    "destino": tabla_frecuencia(datos_tienda, col_destino),
    "alternativa": tabla_competidores(datos_tienda, col_alternativa),
    "no_visito": tabla_frecuencia(datos_tienda, col_no_visito),
    "servicio": tabla_servicio,
}


# =========================================================
# 3. RESUMEN
# =========================================================

st.header("3. Resumen de la tienda")

k1, k2, k3, k4 = st.columns(4)
k1.metric("Encuestas", total_encuestas)
k2.metric("Edad promedio", f"{edad_promedio} años" if edad_promedio != "" else "N/D")
k3.metric("Servicio", f"{servicio_promedio} / 5" if servicio_promedio != "" else "N/D")
k4.metric("Estrato principal", f"{estrato_principal} ({estrato_porcentaje}%)")

k5, k6, k7, k8 = st.columns(4)
k5.metric("Inicio", formato_fecha_hora(fecha_inicio))
k6.metric("Finalización", formato_fecha_hora(fecha_final))
k7.metric("Día con más encuestas", f"{dia_mas_encuestas} ({cantidad_dia_mas})")
k8.metric("Hora pico", hora_pico or "N/D")

if dias_sin_encuestas:
    if len(dias_sin_encuestas) > 10:
        st.warning(
            f"Hay {len(dias_sin_encuestas)} días sin encuestas entre la primera y la última. "
            "Revisa si hay encuestas sueltas de otras fechas."
        )
    else:
        st.warning("Días sin encuestas: " + ", ".join(dias_sin_encuestas))

if edades_invalidas:
    st.caption(f"Se ignoraron {edades_invalidas} edades fuera de rango (menores de 12 o mayores de 99).")

g1, g2 = st.columns(2)
with g1:
    st.subheader("Encuestas por día")
    grafico_barras_app(tablas["dias"], horizontal=False)
with g2:
    st.subheader("Franja horaria")
    grafico_barras_app(tablas["horas"], horizontal=False)

if col_x is not None and col_y is not None:
    puntos = pd.DataFrame({
        "lat": pd.to_numeric(datos_tienda[col_y], errors="coerce"),
        "lon": pd.to_numeric(datos_tienda[col_x], errors="coerce"),
    }).dropna()
    if len(puntos):
        with st.expander("🗺️ Ver en el mapa dónde se hicieron las encuestas"):
            st.map(puntos, color=HEX_ROJO, size=6)


# =========================================================
# 4. PERFIL DEL CLIENTE
# =========================================================

st.header("4. Perfil del cliente")

p1, p2 = st.columns(2)
with p1:
    st.subheader("Género")
    grafico_dona_app(tablas["genero"])
with p2:
    st.subheader("Rango de edad")
    grafico_barras_app(tablas["edades"], horizontal=False)

p3, p4 = st.columns(2)
with p3:
    st.subheader("Estrato")
    grafico_barras_app(tablas["estrato"], horizontal=False)
with p4:
    st.subheader("Ocupación")
    grafico_barras_app(tablas["ocupacion"])


# =========================================================
# 5. HÁBITOS, MOVILIDAD Y COMPETENCIA
# =========================================================

st.header("5. Hábitos de compra")

h1, h2 = st.columns(2)
with h1:
    st.subheader("Principal motivo de compra")
    grafico_barras_app(tablas["motivo"])
with h2:
    st.subheader("Medio de llegada")
    grafico_barras_app(tablas["transporte"])

h3, h4 = st.columns(2)
with h3:
    st.subheader("Visitas por semana")
    grafico_barras_app(tablas["frecuencia"], horizontal=False)
with h4:
    st.subheader("¿Planeaba venir al OXXO?")
    grafico_dona_app(tablas["planeaba"])

o1, o2 = st.columns(2)
with o1:
    st.subheader("¿De dónde viene?")
    grafico_barras_app(tablas["origen"])
with o2:
    st.subheader("¿Hacia dónde se dirige?")
    grafico_barras_app(tablas["destino"])

st.header("6. Competencia y servicio")

c1, c2 = st.columns(2)
with c1:
    st.subheader("¿Dónde compraría si no es en OXXO?")
    grafico_barras_app(tablas["alternativa"])
    st.caption("Los nombres escritos distinto (D1 / d1, Carrulla / Carulla…) se agrupan solos.")
with c2:
    st.subheader("¿Por qué no visitó el OXXO más cercano?")
    grafico_barras_app(tablas["no_visito"])

st.subheader("Calificación del servicio")
grafico_barras_app(tablas["servicio"], max_filas=10)

with st.expander("📋 Ver todas las tablas en detalle"):
    for clave, titulo in [
        ("ocupacion", "Ocupación"), ("motivo", "Motivo de compra"),
        ("transporte", "Medio de llegada"), ("origen", "De dónde viene"),
        ("destino", "Hacia dónde se dirige"), ("alternativa", "Alternativa de compra"),
        ("no_visito", "Por qué no visitó el OXXO más cercano"),
    ]:
        st.write(f"**{titulo}**")
        st.dataframe(tablas[clave], use_container_width=True, hide_index=True)


# =========================================================
# 7. INFORMACIÓN MANUAL
# =========================================================

st.header("7. Información manual")
st.info("Lo que dejes vacío no aparece en el PDF.")

foto_fachada = st.file_uploader("Foto de la fachada", type=["jpg", "jpeg", "png"], key="foto_fachada")

percepcion = st.text_area("Percepción del servicio", placeholder="Escribe aquí la percepción del servicio...")

radio = st.text_input(
    "Radio de influencia",
    placeholder="Ejemplo: 100m: 47 | 200m: 48 | 300m: 74 | +300m: 67",
)

i1, i2 = st.columns(2)
with i1:
    imagen_radio = st.file_uploader("Pantallazo del radio de influencia",
                                    type=["jpg", "jpeg", "png"], key="imagen_radio")
with i2:
    imagen_isocrona = st.file_uploader("Pantallazo de la isócrona",
                                       type=["jpg", "jpeg", "png"], key="imagen_isocrona")

st.subheader("Insights clave")
insight1 = st.text_area("Insight 1", placeholder="Escribe el primer insight...")
insight2 = st.text_area("Insight 2", placeholder="Escribe el segundo insight...")
insight3 = st.text_area("Insight 3", placeholder="Escribe el tercer insight...")

st.subheader("Registro fotográfico")
fotos = st.file_uploader(
    "Fotografías de la tienda (máximo 4)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="fotos",
)
if len(fotos) > 4:
    st.warning("Solo se usarán las primeras 4 fotografías.")
    fotos = fotos[:4]


# =========================================================
# 8. GENERAR REPORTE
# =========================================================

datos = {
    "cr": cr_seleccionado,
    "nombre": nombre_tienda,
    "encuestas": total_encuestas,
    "periodo": periodo,
    "inicio": formato_fecha_hora(fecha_inicio),
    "finalizacion": formato_fecha_hora(fecha_final),
    "dia_mas_encuestas": dia_mas_encuestas,
    "cantidad_dia_mas": cantidad_dia_mas,
    "hora_pico": hora_pico,
    "hora_local": convertir_utc,
    "edad_promedio": edad_promedio,
    "hombres": hombres,
    "mujeres": mujeres,
    "estrato": estrato_principal,
    "estrato_porcentaje": estrato_porcentaje,
    "servicio_promedio": servicio_promedio,
}

manuales = {
    "percepcion": percepcion,
    "radio": radio,
    "foto_fachada": foto_fachada,
    "imagen_radio": imagen_radio,
    "imagen_isocrona": imagen_isocrona,
    "insight1": insight1,
    "insight2": insight2,
    "insight3": insight3,
    "fotos": fotos,
}

st.header("8. Generar reporte")

if not nombre_tienda.strip():
    st.caption("Tip: escribe el nombre de la tienda para que aparezca junto al CR en el PDF.")

if st.button("GENERAR REPORTE PDF", type="primary"):
    try:
        documento = crear_pdf(datos, manuales, tablas)
        nombre_archivo = f"Reporte_EOD_{cr_seleccionado}"
        if nombre_tienda.strip():
            nombre_archivo += "_" + "_".join(nombre_tienda.split())
        st.success("¡Reporte generado correctamente! 🎉")
        st.download_button(
            label="📥 Descargar reporte PDF",
            data=documento,
            file_name=f"{nombre_archivo}.pdf",
            mime="application/pdf",
        )
    except Exception as e:
        st.error(f"No se pudo generar el reporte: {e}")
