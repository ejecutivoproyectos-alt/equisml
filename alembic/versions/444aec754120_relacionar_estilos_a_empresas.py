"""relacionar estilos a empresas

Revision ID: 444aec754120
Revises: 9e840497c702
Create Date: 2026-06-05 13:23:59.092032

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = '444aec754120'
down_revision: Union[str, Sequence[str], None] = '9e840497c702'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "empresa_estilo_word",
        sa.Column("empresa_id", sa.Integer(), nullable=True)
    )

    op.execute("""
        UPDATE empresa_estilo_word ew
        SET empresa_id = a.empresa_externa_id
        FROM empresa_plantilla_asignacion a
        WHERE ew.plantilla_id = a.plantilla_id
          AND a.activo = true
    """)

    op.execute("""
        DELETE FROM empresa_estilo_word
        WHERE empresa_id IS NULL
    """)

    op.alter_column(
        "empresa_estilo_word",
        "empresa_id",
        nullable=False
    )

    op.drop_constraint(
        op.f("empresa_estilo_word_plantilla_id_fkey"),
        "empresa_estilo_word",
        type_="foreignkey"
    )

    op.create_foreign_key(
        "empresa_estilo_word_empresa_id_fkey",
        "empresa_estilo_word",
        "empresas",
        ["empresa_id"],
        ["id"],
        ondelete="CASCADE"
    )

    op.drop_column("empresa_estilo_word", "plantilla_id")

def downgrade() -> None:
    op.add_column(
        'empresa_estilo_word',
        sa.Column('plantilla_id', sa.INTEGER(), nullable=True)
    )

    op.execute("""
        UPDATE empresa_estilo_word ew
        SET plantilla_id = a.plantilla_id
        FROM empresa_plantilla_asignacion a
        WHERE ew.empresa_id = a.empresa_externa_id
        AND a.activo = true
    """)

    op.alter_column(
        'empresa_estilo_word',
        'plantilla_id',
        nullable=False
    )

    op.drop_constraint(
        'empresa_estilo_word_empresa_id_fkey',
        'empresa_estilo_word',
        type_='foreignkey'
    )

    op.create_foreign_key(
        op.f('empresa_estilo_word_plantilla_id_fkey'),
        'empresa_estilo_word',
        'empresa_plantilla_word',
        ['plantilla_id'],
        ['id'],
        ondelete='CASCADE'
    )

    op.drop_column('empresa_estilo_word', 'empresa_id')
