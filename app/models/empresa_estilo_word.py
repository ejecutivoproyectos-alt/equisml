from sqlalchemy import Column, Integer, String, Boolean, ForeignKey
from sqlalchemy.orm import relationship

from app.db.base import Base


class EmpresaEstiloWord(Base):
    __tablename__ = "empresa_estilo_word"

    id = Column(Integer, primary_key=True, index=True)

    plantilla_id = Column(
        Integer,
        ForeignKey("empresa_plantilla_word.id", ondelete="CASCADE"),
        nullable=False
    )

    clave_estilo = Column(String(100), nullable=False)

    tipografia = Column(String(100), nullable=False)
    tamanio_letra = Column(Integer, nullable=False)

    color_letra = Column(String(7), nullable=False)
    color_fondo = Column(String(7), nullable=True)

    negrita = Column(Boolean, default=False, nullable=False)
    cursiva = Column(Boolean, default=False, nullable=False)

    alineacion = Column(String(50), nullable=True)

    plantilla = relationship(
        "EmpresaPlantillaWord",
        back_populates="estilos"
    )