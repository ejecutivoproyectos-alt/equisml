from datetime import date, datetime, timedelta


def convertir_str_a_date(fecha_str: str) -> date:
    return datetime.strptime(fecha_str, "%d/%m/%Y").date()


def primer_lunes(anio: int, mes: int) -> date:
    d = date(anio, mes, 1)

    while d.weekday() != 0:
        d += timedelta(days=1)

    return d


def tercer_lunes(anio: int, mes: int) -> date:
    return primer_lunes(anio, mes) + timedelta(days=14)


def obtener_dias_inhabiles_oficiales_mexico(anio: int) -> set[date]:
    festivos = set()

    festivos.add(date(anio, 1, 1))
    festivos.add(primer_lunes(anio, 2))
    festivos.add(tercer_lunes(anio, 3))
    festivos.add(date(anio, 5, 1))
    festivos.add(date(anio, 9, 16))
    festivos.add(tercer_lunes(anio, 11))
    festivos.add(date(anio, 12, 25))

    # Cambio de Poder Ejecutivo Federal cada 6 años
    if anio in [2024, 2030, 2036, 2042, 2048]:
        festivos.add(date(anio, 10, 1))

    return festivos


def es_dia_habil_mexico(fecha: date) -> bool:
    # Lunes=0, martes=1, ..., sábado=5, domingo=6
    if fecha.weekday() == 6:
        return False

    return fecha not in obtener_dias_inhabiles_oficiales_mexico(fecha.year)


def restar_dias_habiles_mexico(fecha_base: date, dias: int) -> date:
    fecha_actual = fecha_base
    contador = 0

    while contador < dias:
        fecha_actual -= timedelta(days=1)

        if es_dia_habil_mexico(fecha_actual):
            contador += 1

    return fecha_actual


def sumar_dias_habiles_mexico(fecha_base: date, dias: int) -> date:
    fecha_actual = fecha_base
    contador = 0

    while contador < dias:
        fecha_actual += timedelta(days=1)

        if es_dia_habil_mexico(fecha_actual):
            contador += 1

    return fecha_actual


def calcular_fecha_habil_mexico(
    fecha_str: str,
    dias_restar: int = 15,
    dias_sumar: int = 0
) -> date:
    fecha_base = convertir_str_a_date(fecha_str)

    fecha_calculada = restar_dias_habiles_mexico(
        fecha_base,
        dias_restar
    )

    fecha_calculada = sumar_dias_habiles_mexico(
        fecha_calculada,
        dias_sumar
    )

    return fecha_calculada


def formatear_fecha_larga_espanol(fecha: date) -> str:
    meses = {
        1: "enero",
        2: "febrero",
        3: "marzo",
        4: "abril",
        5: "mayo",
        6: "junio",
        7: "julio",
        8: "agosto",
        9: "septiembre",
        10: "octubre",
        11: "noviembre",
        12: "diciembre",
    }

    return f"{fecha.day} de {meses[fecha.month]} del {fecha.year}"


def calcular_fecha_larga_habil_mexico(
    fecha_str: str,
    dias_restar: int = 15,
    dias_sumar: int = 0
) -> str:
    fecha = calcular_fecha_habil_mexico(
        fecha_str=fecha_str,
        dias_restar=dias_restar,
        dias_sumar=dias_sumar
    )

    return formatear_fecha_larga_espanol(fecha)