from sqlalchemy import Column, Integer, String, DateTime
from datetime import datetime
from sqlalchemy.orm import relationship
from app.db.base import Base


class Empresa(Base):
    __tablename__ = "empresas"

    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String(255), nullable=False)
    razon_social = Column(String(255), nullable=True)
    color_primario = Column(String(7), nullable=False, default="#000000")
    color_secundario = Column(String(7), nullable=True)
    color_acento = Column(String(7), nullable=True)
    fecha_creacion = Column(DateTime, default=datetime.now)

    plantilla_asignacion = relationship(
        "EmpresaPlantillaAsignacion",
        back_populates="empresa",
        cascade="all, delete-orphan",
        uselist=False
    )

    estilos_word = relationship(
        "EmpresaEstiloWord",
        back_populates="empresa",
        cascade="all, delete-orphan"
    )