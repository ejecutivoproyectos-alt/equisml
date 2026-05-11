from sqlalchemy.orm import Session
from app.repositories.empresa_repository import obtener_empresas


def listar_empresas(db: Session):
    return obtener_empresas(db)