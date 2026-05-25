from sqlalchemy import Column, Integer, String, ForeignKey, DateTime, Boolean, UniqueConstraint
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func

from app.db.base import Base


class EmpresaPlantillaAsignacion(Base):
    __tablename__ = "empresa_plantilla_asignacion"

    id = Column(Integer, primary_key=True, index=True)

    empresa_externa_id = Column(
        Integer,
        ForeignKey("empresas.id", ondelete="CASCADE"),
        nullable=False
    )

    plantilla_id = Column(
        Integer,
        ForeignKey("empresa_plantilla_word.id", ondelete="CASCADE"),
        nullable=True
    )

    membrete_path = Column(String(500), nullable=True)

    activo = Column(Boolean, default=True, nullable=False)

    creado_en = Column(
        DateTime(timezone=True),
        server_default=func.now(),
        nullable=False
    )

    actualizado_en = Column(
        DateTime(timezone=True),
        server_default=func.now(),
        onupdate=func.now(),
        nullable=False
    )

    empresa = relationship(
        "Empresa",
        back_populates="plantilla_asignacion"
    )

    plantilla = relationship(
        "EmpresaPlantillaWord",
        back_populates="asignaciones"
    )

    __table_args__ = (
        UniqueConstraint(
            "empresa_externa_id",
            name="uq_empresa_plantilla_asignacion_empresa"
        ),
    )