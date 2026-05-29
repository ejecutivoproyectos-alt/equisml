import app.db.base_metadata

from app.db.database import SessionLocal

from app.models.empresa_plantilla_word import EmpresaPlantillaWord
from app.models.empresa_estilo_word import EmpresaEstiloWord


ESTILOS = [
    # WAVELENS
    {
        "nombre_disenio": "wavelens",
        "clave_estilo": "titulo_1",
        "tipografia": "Californian FB",
        "tamanio_letra": 16,
        "color_letra": "#833C0B",
        "negrita": True,
        "cursiva": False,
        "alineacion": "center",
    },
    {
        "nombre_disenio": "wavelens",
        "clave_estilo": "titulo_2",
        "tipografia": "Californian FB",
        "tamanio_letra": 14,
        "color_letra": "#BF8F00",
        "negrita": True,
        "cursiva": False,
        "alineacion": "left",
    },
    {
        "nombre_disenio": "wavelens",
        "clave_estilo": "titulo_3",
        "tipografia": "Calibri Light",
        "tamanio_letra": 12,
        "color_letra": "#1F4D78",
        "negrita": False,
        "cursiva": False,
        "alineacion": "left",
    },

    # LITZA
    {
        "nombre_disenio": "litza",
        "clave_estilo": "titulo_1",
        "tipografia": "Cambria",
        "tamanio_letra": 16,
        "color_letra": "#BF8F00",
        "negrita": True,
        "cursiva": False,
        "alineacion": "center",
    },
    {
        "nombre_disenio": "litza",
        "clave_estilo": "titulo_2",
        "tipografia": "Aptos Display",
        "tamanio_letra": 14,
        "color_letra": "#806000",
        "negrita": True,
        "cursiva": False,
        "alineacion": "left",
    },
    {
        "nombre_disenio": "litza",
        "clave_estilo": "titulo_3",
        "tipografia": "Cambria",
        "tamanio_letra": 12,
        "color_letra": "#806000",
        "negrita": True,
        "cursiva": False,
        "alineacion": "left",
    },
]


def run():
    db = SessionLocal()

    try:
        creados = 0
        existentes = 0

        for estilo_data in ESTILOS:

            plantilla = db.query(EmpresaPlantillaWord).filter(
                EmpresaPlantillaWord.nombre_disenio ==
                estilo_data["nombre_disenio"]
            ).first()

            if not plantilla:
                print(
                    f"No existe plantilla: "
                    f"{estilo_data['nombre_disenio']}"
                )
                continue

            estilo_existente = db.query(EmpresaEstiloWord).filter(
                EmpresaEstiloWord.plantilla_id == plantilla.id,
                EmpresaEstiloWord.clave_estilo ==
                estilo_data["clave_estilo"]
            ).first()

            if estilo_existente:
                print(
                    f"El estilo "
                    f"{estilo_data['clave_estilo']} "
                    f"ya existe."
                )

                existentes += 1
                continue

            nuevo_estilo = EmpresaEstiloWord(
                plantilla_id=plantilla.id,
                clave_estilo=str(estilo_data["clave_estilo"]),
                tipografia=str(estilo_data["tipografia"]),
                tamanio_letra=int(estilo_data["tamanio_letra"]),
                color_letra=str(estilo_data["color_letra"]),
                negrita=bool(estilo_data["negrita"]),
                cursiva=bool(estilo_data["cursiva"]),
                alineacion=str(estilo_data["alineacion"]),
            )

            db.add(nuevo_estilo)

            creados += 1

        db.commit()

        print("Seed empresa_estilo_word ejecutado correctamente")
        print(f"Estilos creados: {creados}")
        print(f"Estilos existentes: {existentes}")

    except Exception as e:
        db.rollback()
        print(f"Error en seed empresa_estilo_word: {e}")
        raise

    finally:
        db.close()


if __name__ == "__main__":
    run()