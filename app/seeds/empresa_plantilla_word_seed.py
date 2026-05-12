import app.db.base_metadata

from datetime import datetime
from sqlalchemy import func

from app.db.database import SessionLocal
from app.models.empresa import Empresa
from app.models.empresa_plantilla_word import EmpresaPlantillaWord


PLANTILLAS = [
    {
        "empresa_nombre": "wavelens",
        "nombre_disenio": "wavelens",
        "tipografia_base": "Arial",
        "tamanio_base": 12,
        "color_texto_base": "#000000",
        "color_primario": "#C45911",
        "color_secundario": "#FBE4D5",
        "color_acento": "#F5E6CC",
        "membrete_path": "membretes/wavelens/7.ACUSE.docx",
    },
    {
        "empresa_nombre": "litza",
        "nombre_disenio": "litza",
        "tipografia_base": "Arial",
        "tamanio_base": 12,
        "color_texto_base": "#000000",
        "color_primario": "#CAA006",
        "color_secundario": "#FFD966",
        "color_acento": "#F5E6CC",
        "membrete_path": "membretes/litza/7.ACUSE.docx",
    },
]


def run():
    db = SessionLocal()

    try:
        creadas = 0
        existentes = 0

        for plantilla_data in PLANTILLAS:
            empresa_nombre = str(plantilla_data["empresa_nombre"])

            empresa = db.query(Empresa).filter(
                func.lower(Empresa.nombre) == empresa_nombre.lower()
            ).first()

            if not empresa:
                print(f"La empresa {empresa_nombre} no existe.")
                continue

            plantilla_existente = db.query(EmpresaPlantillaWord).filter(
                EmpresaPlantillaWord.empresa_id == empresa.id
            ).first()

            if plantilla_existente:
                print(f"La plantilla de {empresa.nombre} ya existe.")
                existentes += 1
                continue

            nueva_plantilla = EmpresaPlantillaWord(
                empresa_id=empresa.id,
                nombre_disenio=str(plantilla_data["nombre_disenio"]),
                tipografia_base=str(plantilla_data["tipografia_base"]),
                tamanio_base=int(plantilla_data["tamanio_base"]),
                color_texto_base=str(plantilla_data["color_texto_base"]),
                color_primario=str(plantilla_data["color_primario"]),
                color_secundario=str(plantilla_data["color_secundario"]),
                color_acento=str(plantilla_data["color_acento"]),
                membrete_path=str(plantilla_data["membrete_path"]),
                creado_en=datetime.now(),
                actualizado_en=datetime.now()
            )

            db.add(nueva_plantilla)
            creadas += 1

        db.commit()

        print("Seed empresa_plantilla_word ejecutado correctamente")
        print(f"Plantillas creadas: {creadas}")
        print(f"Plantillas existentes: {existentes}")

    except Exception as e:
        db.rollback()
        print(f"Error en seed empresa_plantilla_word: {e}")
        raise

    finally:
        db.close()


if __name__ == "__main__":
    run()