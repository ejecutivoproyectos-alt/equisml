from sqlalchemy.orm import Session
from app.models.empresa import Empresa


def obtener_empresas(db: Session):
    return db.query(Empresa).order_by(Empresa.nombre.asc()).all()


def obtener_empresa_por_id(db: Session, empresa_id: int):
    return db.query(Empresa).filter(Empresa.id == empresa_id).first()