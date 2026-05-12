import app.db.base_metadata
from app.seeds.empresa_seed import run as empresa_seed
from app.seeds.empresa_plantilla_word_seed import run as plantilla_seed
from app.seeds.empresa_estilo_word_seed import run as estilo_seed

def ejecutar_seed(nombre, funcion):
    try:
        print(f"\nEjecutando seed: {nombre}")

        funcion()

        print(f"Seed {nombre} ejecutado correctamente")

    except Exception as e:
        print(f"Error en seed {nombre}: {e}")


def run():
    print("=== INICIANDO SEEDS ===")

    ejecutar_seed("empresas", empresa_seed)
    ejecutar_seed("empresa_plantilla_word", plantilla_seed)
    ejecutar_seed("empresa_estilo_word", estilo_seed)

    print("\n=== SEEDS FINALIZADOS ===")


if __name__ == "__main__":
    run()