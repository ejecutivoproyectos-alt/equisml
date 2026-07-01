"""normalizar rutas plantillas

Revision ID: c2d4e6f8a901
Revises: b7a8f2c9d4e1
Create Date: 2026-07-01 00:00:00.000000

"""
from typing import Sequence, Union

from alembic import op


# revision identifiers, used by Alembic.
revision: str = "c2d4e6f8a901"
down_revision: Union[str, Sequence[str], None] = "b7a8f2c9d4e1"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.execute("""
        UPDATE empresa_plantilla_word
        SET plantilla_path = 'plantillas/triple_a/' || regexp_replace(
            plantilla_path,
            '^plantillas_triple_a/',
            ''
        )
        WHERE plantilla_path LIKE 'plantillas_triple_a/%'
    """)

    op.execute("""
        UPDATE empresa_plantilla_word
        SET plantilla_path = 'plantillas/doble_a/' || regexp_replace(
            plantilla_path,
            '^plantillas/',
            ''
        )
        WHERE tipo_plantilla = 'DOBLE_AA'
          AND plantilla_path LIKE 'plantillas/%'
          AND plantilla_path NOT LIKE 'plantillas/doble_a/%'
          AND plantilla_path NOT LIKE 'plantillas/triple_a/%'
    """)

    op.execute("""
        UPDATE empresa_plantilla_word
        SET plantilla_path = 'plantillas/triple_a/' || regexp_replace(
            plantilla_path,
            '^plantillas/',
            ''
        )
        WHERE tipo_plantilla = 'TRIPLE_AAA'
          AND plantilla_path LIKE 'plantillas/%'
          AND plantilla_path NOT LIKE 'plantillas/doble_a/%'
          AND plantilla_path NOT LIKE 'plantillas/triple_a/%'
    """)


def downgrade() -> None:
    op.execute("""
        UPDATE empresa_plantilla_word
        SET plantilla_path = 'plantillas/' || regexp_replace(
            plantilla_path,
            '^plantillas/doble_a/',
            ''
        )
        WHERE plantilla_path LIKE 'plantillas/doble_a/%'
    """)

    op.execute("""
        UPDATE empresa_plantilla_word
        SET plantilla_path = 'plantillas_triple_a/' || regexp_replace(
            plantilla_path,
            '^plantillas/triple_a/',
            ''
        )
        WHERE plantilla_path LIKE 'plantillas/triple_a/%'
    """)
