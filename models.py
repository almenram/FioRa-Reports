from __future__ import annotations

from datetime import date
from sqlalchemy import String, Integer, Numeric, Date, ForeignKey

from sqlalchemy.orm import Mapped, mapped_column

from database import Base


class Material(Base):
    __tablename__ = "materiales"

    codigo: Mapped[str] = mapped_column(
        String,
        primary_key=True
    )

    descripcion: Mapped[str] = mapped_column(
        String,
        nullable=False
    )

    categoria: Mapped[str] = mapped_column(
        String,
        nullable=False
    )


class RelacionLiderSolicitante(Base):
    __tablename__ = "lider_solicitante"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True
    )

    solicitante: Mapped[str] = mapped_column(
        String,
        nullable=False,
        unique=True
    )

    lider: Mapped[str] = mapped_column(
        String,
        nullable=False
    )


class Registro(Base):
    __tablename__ = "registros"

    id: Mapped[int] = mapped_column(
        Integer,
        primary_key=True
    )

    codigo: Mapped[str] = mapped_column(
        String,
        ForeignKey("materiales.codigo"),
        nullable=False
    )

    descripcion: Mapped[str] = mapped_column(
        String,
        nullable=False
    )

    cantidad: Mapped[float] = mapped_column(
        Numeric(10, 2),
        nullable=False
    )

    solicitante: Mapped[str] = mapped_column(
    String,
    nullable=True
    )

    lider: Mapped[str] = mapped_column(
    String,
    nullable=True
    )

    fabricacion: Mapped[str] = mapped_column(
        String,
        nullable=False
    )

    fecha: Mapped[date] = mapped_column(
        Date,
        nullable=False
    )