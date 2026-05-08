from app.db.database import SessionLocal
from app.models.empresa import Empresa


def run():
    db = SessionLocal()

    try:
        empresas = [
            Empresa(nombre="WAVELENS"),
            Empresa(nombre="LITZA")
        ]

        db.add_all(empresas)
        db.commit()

        print("Empresas insertadas correctamente")

    except Exception as e:
        db.rollback()
        print(f"Error: {e}")

    finally:
        db.close()


if __name__ == "__main__":
    run()