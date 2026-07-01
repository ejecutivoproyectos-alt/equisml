"""normalizar rutas membretes

Revision ID: d3e5f7a9b102
Revises: c2d4e6f8a901
Create Date: 2026-07-01 00:00:00.000000

"""
from typing import Sequence, Union

from alembic import op


# revision identifiers, used by Alembic.
revision: str = "d3e5f7a9b102"
down_revision: Union[str, Sequence[str], None] = "c2d4e6f8a901"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute("""
        UPDATE empresa_plantilla_asignacion asignacion
        SET membrete_path = 'membretes/doble_a/' || regexp_replace(
            asignacion.membrete_path,
            '^membretes/',
            ''
        )
        FROM empresas empresa
        WHERE empresa.id = asignacion.empresa_externa_id
          AND empresa.tipo_empresa = 'DOBLE_AA'
          AND asignacion.membrete_path LIKE 'membretes/%'
          AND asignacion.membrete_path NOT LIKE 'membretes/doble_a/%'
          AND asignacion.membrete_path NOT LIKE 'membretes/triple_a/%'
    """)

    op.execute("""
        UPDATE empresa_plantilla_asignacion asignacion
        SET membrete_path = 'membretes/triple_a/' || regexp_replace(
            asignacion.membrete_path,
            '^membretes/',
            ''
        )
        FROM empresas empresa
        WHERE empresa.id = asignacion.empresa_externa_id
          AND empresa.tipo_empresa = 'TRIPLE_AAA'
          AND asignacion.membrete_path LIKE 'membretes/%'
          AND asignacion.membrete_path NOT LIKE 'membretes/doble_a/%'
          AND asignacion.membrete_path NOT LIKE 'membretes/triple_a/%'
    """)


def downgrade() -> None:
    op.execute("""
        UPDATE empresa_plantilla_asignacion
        SET membrete_path = 'membretes/' || regexp_replace(
            membrete_path,
            '^membretes/doble_a/',
            ''
        )
        WHERE membrete_path LIKE 'membretes/doble_a/%'
    """)

    op.execute("""
        UPDATE empresa_plantilla_asignacion
        SET membrete_path = 'membretes/' || regexp_replace(
            membrete_path,
            '^membretes/triple_a/',
            ''
        )
        WHERE membrete_path LIKE 'membretes/triple_a/%'
    """)
