"""agregar tipos empresa y plantilla

Revision ID: b7a8f2c9d4e1
Revises: 444aec754120
Create Date: 2026-06-30 00:00:00.000000

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = "b7a8f2c9d4e1"
down_revision: Union[str, Sequence[str], None] = "444aec754120"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "empresas",
        sa.Column(
            "tipo_empresa",
            sa.String(length=20),
            nullable=False,
            server_default="DOBLE_AA",
        ),
    )
    op.add_column(
        "empresa_plantilla_word",
        sa.Column(
            "tipo_plantilla",
            sa.String(length=20),
            nullable=False,
            server_default="DOBLE_AA",
        ),
    )
    op.execute("UPDATE empresas SET tipo_empresa = 'DOBLE_AA' WHERE tipo_empresa IS NULL")
    op.execute("UPDATE empresa_plantilla_word SET tipo_plantilla = 'DOBLE_AA' WHERE tipo_plantilla IS NULL")
    op.execute("DROP TABLE IF EXISTS plantillas_triple_a")


def downgrade() -> None:
    op.drop_column("empresa_plantilla_word", "tipo_plantilla")
    op.drop_column("empresas", "tipo_empresa")
