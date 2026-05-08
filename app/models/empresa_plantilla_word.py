from sqlalchemy import (
    Column,
    Integer,
    String,
    ForeignKey,
    DateTime,
)
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func

from app.db.base import Base


class EmpresaPlantillaWord(Base):
    __tablename__ = "empresa_plantilla_word"

    id = Column(Integer, primary_key=True, index=True)

    empresa_id = Column(
        Integer,
        ForeignKey("empresas.id", ondelete="CASCADE"),
        nullable=False,
        unique=True
    )

    nombre_disenio = Column(
        String(100),
        nullable=False
    )

    tipografia_base = Column(
        String(100),
        nullable=False
    )

    tamanio_base = Column(
        Integer,
        nullable=False
    )

    color_texto_base = Column(
        String(7),
        nullable=False
    )

    color_primario = Column(
        String(7),
        nullable=False
    )

    color_secundario = Column(
        String(7),
        nullable=True
    )

    color_acento = Column(
        String(7),
        nullable=True
    )

    membrete_path = Column(
        String(500),
        nullable=True
    )

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

    # RELACIONES

    empresa = relationship(
        "Empresa",
        back_populates="plantilla_word"
    )

    estilos = relationship(
        "EmpresaEstiloWord",
        back_populates="plantilla",
        cascade="all, delete-orphan"
    )