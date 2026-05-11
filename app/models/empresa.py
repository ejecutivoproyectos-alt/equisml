from sqlalchemy import Column, Integer, String, DateTime
from datetime import datetime
from sqlalchemy.orm import relationship
from app.db.base import Base

class Empresa(Base):
    __tablename__ = "empresas"

    id = Column(Integer, primary_key=True, index=True)
    nombre = Column(String(255), nullable=False)
    razon_social = Column(String(255), nullable=True)
    fecha_creacion = Column(DateTime, default=datetime.now)

    plantilla_word = relationship(
        "EmpresaPlantillaWord",
        back_populates="empresa",
        cascade="all, delete-orphan"
    )