def aplicar_heading_word(parrafo, nivel):
    estilos_por_nivel = {
        1: ["Título 1", "Heading 1"],
        2: ["Título 2", "Heading 2"],
        3: ["Título 3", "Heading 3"],
        4: ["Título 4", "Heading 4"],
    }

    if nivel not in estilos_por_nivel:
        raise ValueError(
            f"Nivel de título inválido: {nivel}. "
            "Solo se permiten niveles 1, 2, 3 o 4."
        )

    for estilo in estilos_por_nivel[nivel]:
        try:
            parrafo.style = estilo
            return True
        except Exception:
            pass

    return False


def aplicar_heading_a_texto(doc, texto_objetivo, nivel):
    if not texto_objetivo:
        return False

    aplicado = False

    for parrafo in doc.paragraphs:
        if texto_objetivo in parrafo.text:
            if aplicar_heading_word(parrafo, nivel):
                aplicado = True

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for parrafo in celda.paragraphs:
                    if texto_objetivo in parrafo.text:
                        if aplicar_heading_word(parrafo, nivel):
                            aplicado = True

    return aplicado