import app.db.base_metadata
from app.db.database import SessionLocal
from app.models.empresa import Empresa


EMPRESAS = [
    {
        "nombre": "WAVELENS",
        "razon_social": "WAVELENS SA DE CV",
    },
    {
        "nombre": "LITZA",
        "razon_social": "LITZA SA DE CV",
    },
]


def run():
    db = SessionLocal()

    try:
        empresas_creadas = 0
        empresas_existentes = 0

        for empresa_data in EMPRESAS:
            empresa_existente = db.query(Empresa).filter(
                Empresa.nombre == empresa_data["nombre"]
            ).first()

            if empresa_existente:
                empresas_existentes += 1
                continue

            nueva_empresa = Empresa(
                nombre=empresa_data["nombre"],
                razon_social=empresa_data["razon_social"],
            )

            db.add(nueva_empresa)
            empresas_creadas += 1

        db.commit()

        print(f"Seed empresas ejecutado correctamente.")
        print(f"Empresas creadas: {empresas_creadas}")
        print(f"Empresas ya existentes: {empresas_existentes}")


    except Exception as e:
        db.rollback()
        print(f"Error al ejecutar seed de empresas: {e}")
        raise

    finally:
        db.close()


if __name__ == "__main__":
    run()