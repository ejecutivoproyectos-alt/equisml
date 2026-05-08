from decimal import Decimal

def formatear_moneda(valor):
    return f"${valor:,.2f}"


def convertir_numero_a_letras_mxn(total):
    unidades = [
        "", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE",
        "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS",
        "DIECISIETE", "DIECIOCHO", "DIECINUEVE"
    ]

    decenas = [
        "", "", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA",
        "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"
    ]

    centenas = [
        "", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS",
        "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"
    ]


    def convertir_menor_mil(n):
        if n == 0:
            return ""
        if n == 100:
            return "CIEN"
        if n < 20:
            return unidades[n]
        if n < 30:
            if n == 20:
                return "VEINTE"
            return "VEINTI" + unidades[n - 20].lower().upper()
        if n < 100:
            d = n // 10
            u = n % 10
            if u == 0:
                return decenas[d]
            return f"{decenas[d]} Y {unidades[u]}"
        c = n // 100
        resto = n % 100
        if resto == 0:
            return centenas[c]
        return f"{centenas[c]} {convertir_menor_mil(resto)}"

    def convertir_entero(n):
        if n == 0:
            return "CERO"

        millones = n // 1_000_000
        resto_millones = n % 1_000_000
        miles = resto_millones // 1000
        resto = resto_millones % 1000

        partes = []

        if millones > 0:
            if millones == 1:
                partes.append("UN MILLÓN")
            else:
                partes.append(f"{convertir_entero(millones)} MILLONES")

        if miles > 0:
            if miles == 1:
                partes.append("MIL")
            else:
                partes.append(f"{convertir_menor_mil(miles)} MIL")

        if resto > 0:
            partes.append(convertir_menor_mil(resto))

        return " ".join(partes)

    total_redondeado = round(float(total), 2)

    entero = int(total_redondeado)
    centavos = int(round((total_redondeado - entero) * 100))

    if centavos == 100:
        entero += 1
        centavos = 0

    letras_entero = convertir_entero(entero)

    if centavos == 0:
        return f"SON: {letras_entero} PESOS 00/100 M.N."

    return f"SON: {letras_entero} PESOS CON {centavos:02d} CENTAVOS {centavos:02d}/100 M.N."