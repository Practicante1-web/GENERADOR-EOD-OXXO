def crear_word(datos, manuales, tablas):
    """
    Genera el Word con la información del estudio.
    Incluye:
    - Imagen del radio de influencia
    - Imagen de la isócrona
    - Registro fotográfico de la tienda
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
        f"Día con más encuestas: "
        f"{datos['dia_mas_encuestas']}"
    )

    doc.add_paragraph(
        f"Cantidad ese día: "
        f"{datos['cantidad_dia_mas']}"
    )

    # =====================================================
    # PERFIL DEL CLIENTE
    # =====================================================

    doc.add_heading(
        "PERFIL DEL CLIENTE",
        level=2
    )

    doc.add_paragraph(
        f"Edad promedio: "
        f"{datos['edad_promedio']} años"
    )

    doc.add_paragraph(
        f"Hombre: {datos['hombres']}"
    )

    doc.add_paragraph(
        f"Mujer: {datos['mujeres']}"
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
    # PERCEPCIÓN
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
    if manuales.get("imagen_radio") is not None:

        imagen_radio = manuales["imagen_radio"]

        imagen_radio.seek(0)

        doc.add_picture(
            imagen_radio,
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

    if manuales.get("imagen_isocrona") is not None:

        imagen_isocrona = manuales["imagen_isocrona"]

        imagen_isocrona.seek(0)

        doc.add_picture(
            imagen_isocrona,
            width=Inches(6)
        )

    # =====================================================
    # INSIGHTS
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

    # =====================================================
    # REGISTRO FOTOGRÁFICO
    # =====================================================

    if manuales.get("fotos"):

        doc.add_heading(
            "REGISTRO FOTOGRÁFICO",
            level=2
        )

        for foto in manuales["fotos"]:

            # Reiniciar el archivo antes de leerlo
            foto.seek(0)

            doc.add_picture(
                foto,
                width=Inches(5)
            )

            # Espacio entre fotografías
            doc.add_paragraph("")

    # =====================================================
    # GUARDAR
    # =====================================================

    buffer = BytesIO()

    doc.save(buffer)

    buffer.seek(0)

    return buffer
