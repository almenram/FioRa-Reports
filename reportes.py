from datetime import date

from database import SessionLocal
from models import Material, RelacionLiderSolicitante, Registro


# =========================
# MATERIALES
# =========================

def agregar_material(codigo, descripcion, categoria):
    db = SessionLocal()

    try:
        material = Material(
            codigo=codigo,
            descripcion=descripcion,
            categoria=categoria
        )

        db.add(material)
        db.commit()

    finally:
        db.close()


def obtener_material(codigo):
    db = SessionLocal()

    try:
        return (
            db.query(Material)
            .filter_by(codigo=codigo)
            .first()
        )
    finally:
        db.close()


def obtener_materiales():
    db = SessionLocal()

    try:
        return (
            db.query(Material)
            .order_by(Material.descripcion)
            .all()
        )
    finally:
        db.close()
def obtener_categorias():
    db = SessionLocal()

    try:
        resultados = (
            db.query(Material.categoria)
            .distinct()
            .order_by(Material.categoria)
            .all()
        )

        categorias = [resultado[0] for resultado in resultados if resultado[0]]

        # Categorías iniciales del sistema. Las nuevas categorías
        # se incorporan automáticamente al agregarse un material.
        for categoria in ("EPP", "DISCOS", "SOLDADURA"):
            if categoria not in categorias:
                categorias.append(categoria)

        return sorted(categorias)

    finally:
        db.close()


def buscar_materiales(texto):
    db = SessionLocal()

    try:
        resultados = (
            db.query(Material)
            .filter(
                Material.descripcion.ilike(f"%{texto}%")
            )
            .order_by(Material.descripcion)
            .all()
        )

        return resultados

    finally:
        db.close()


# =========================
# SOLICITANTES / LÍDERES
# =========================

def agregar_relacion(solicitante, lider):
    db = SessionLocal()

    try:
        relacion = RelacionLiderSolicitante(
            solicitante=solicitante,
            lider=lider
        )

        db.add(relacion)
        db.commit()

    finally:
        db.close()


def obtener_lider(solicitante):
    db = SessionLocal()

    try:
        relacion = (
            db.query(RelacionLiderSolicitante)
            .filter_by(solicitante=solicitante)
            .first()
        )

        if relacion:
            return relacion.lider

        return None

    finally:
        db.close()


def obtener_relaciones():
    db = SessionLocal()

    try:
        return (
            db.query(RelacionLiderSolicitante)
            .order_by(
                RelacionLiderSolicitante.lider,
                RelacionLiderSolicitante.solicitante
            )
            .all()
        )
    finally:
        db.close()

def obtener_relacion(id_relacion):
    db = SessionLocal()

    try:
        return (
            db.query(RelacionLiderSolicitante)
            .filter_by(id=id_relacion)
            .first()
        )
    finally:
        db.close()

def actualizar_relacion(id_relacion, solicitante, lider):
    db = SessionLocal()

    try:
        relacion = (
            db.query(RelacionLiderSolicitante)
            .filter_by(id=id_relacion)
            .first()
        )

        if not relacion:
            raise ValueError("La relación no existe")

        relacion.solicitante = solicitante
        relacion.lider = lider

        db.commit()

    finally:
        db.close()


def obtener_lideres():
    db = SessionLocal()

    try:
        resultados = (
            db.query(RelacionLiderSolicitante.lider)
            .distinct()
            .order_by(RelacionLiderSolicitante.lider)
            .all()
        )

        return [resultado[0] for resultado in resultados]

    finally:
        db.close()


def obtener_solicitantes_por_lider(lider):
    db = SessionLocal()

    try:
        resultados = (
            db.query(RelacionLiderSolicitante.solicitante)
            .filter_by(lider=lider)
            .order_by(
                RelacionLiderSolicitante.solicitante
            )
            .all()
        )

        return [resultado[0] for resultado in resultados]

    finally:
        db.close()


# =========================
# REGISTROS
# =========================

def agregar_registro(codigo, cantidad, solicitante, lider, fabricacion):
    db = SessionLocal()

    try:
        material = (
            db.query(Material)
            .filter_by(codigo=codigo)
            .first()
        )

        if not material:
            raise ValueError(
                "El código de material no existe"
            )

        # Todo reporte requiere líder
        if not lider:
            raise ValueError(
                "Todos los reportes requieren un líder"
            )

        # EPP requiere solicitante y que su líder corresponda
        if material.categoria == "EPP":

            if not solicitante:
                raise ValueError(
                    "Los materiales EPP requieren un solicitante"
                )

            relacion = (
                db.query(RelacionLiderSolicitante)
                .filter_by(solicitante=solicitante)
                .first()
            )

            if not relacion:
                raise ValueError(
                    "El solicitante no tiene un líder asignado"
                )

            if relacion.lider != lider:
                raise ValueError(
                    "El líder seleccionado no corresponde al solicitante"
                )

        else:
            # Los materiales que no son EPP no llevan solicitante
            solicitante = None

        registro = Registro(
            codigo=codigo,
            descripcion=material.descripcion,
            cantidad=cantidad,
            solicitante=solicitante,
            lider=lider,
            fabricacion=fabricacion,
            fecha=date.today()
        )

        db.add(registro)
        db.commit()
        db.refresh(registro)

        return registro

    finally:
        db.close()


def obtener_registros(
    id_registro=None,
    codigo=None,
    solicitante=None,
    lider=None,
    fabricacion=None,
    categoria=None,
    fecha_inicio=None,
    fecha_fin=None
):
    db = SessionLocal()

    try:

        consulta = db.query(Registro)

        if id_registro:
            consulta = consulta.filter(
                Registro.id == id_registro
            )

        if codigo:
            consulta = consulta.filter(
                Registro.codigo == codigo
            )

        if solicitante:
            consulta = consulta.filter(
                Registro.solicitante == solicitante
            )

        if lider:
            consulta = consulta.filter(
                Registro.lider == lider
            )

        if fabricacion:
            consulta = consulta.filter(
                Registro.fabricacion.ilike(f"%{fabricacion}%")
            )

        if categoria:
            consulta = (
                consulta
                .join(Material, Registro.codigo == Material.codigo)
                .filter(Material.categoria == categoria)
            )

        if fecha_inicio:
            consulta = consulta.filter(
            Registro.fecha >= fecha_inicio
        )

        if fecha_fin:
            consulta = consulta.filter(
            Registro.fecha <= fecha_fin
        )

        return (
            consulta
            .order_by(Registro.fecha)
            .all()
        )

    finally:
        db.close()

# =========================
# ACTUALIZACIONES
# =========================

def actualizar_material(codigo, descripcion, categoria):
    db = SessionLocal()

    try:
        material = (
            db.query(Material)
            .filter_by(codigo=codigo)
            .first()
        )

        if not material:
            raise ValueError("El material no existe")

        material.descripcion = descripcion
        material.categoria = categoria
        db.commit()
        db.refresh(material)

        return material

    finally:
        db.close()


def obtener_registro(id_registro):
    db = SessionLocal()

    try:
        return (
            db.query(Registro)
            .filter_by(id=id_registro)
            .first()
        )
    finally:
        db.close()


def actualizar_registro(
    id_registro,
    codigo,
    cantidad,
    solicitante,
    lider,
    fabricacion
):
    db = SessionLocal()

    try:
        registro = (
            db.query(Registro)
            .filter_by(id=id_registro)
            .first()
        )

        if not registro:
            raise ValueError("El reporte no existe")

        material = (
            db.query(Material)
            .filter_by(codigo=codigo)
            .first()
        )

        if not material:
            raise ValueError("El código de material no existe")

        if not lider:
            raise ValueError(
                "Todos los reportes requieren un líder"
            )

        if material.categoria == "EPP":
            if not solicitante:
                raise ValueError(
                    "Los materiales EPP requieren un solicitante"
                )

            relacion = (
                db.query(RelacionLiderSolicitante)
                .filter_by(solicitante=solicitante)
                .first()
            )

            if not relacion:
                raise ValueError(
                    "El solicitante no tiene un líder asignado"
                )

            if relacion.lider != lider:
                raise ValueError(
                    "El líder seleccionado no corresponde al solicitante"
                )

        else:
            solicitante = None

        registro.codigo = codigo
        registro.descripcion = material.descripcion
        registro.cantidad = cantidad
        registro.solicitante = solicitante
        registro.lider = lider
        registro.fabricacion = fabricacion

        db.commit()
        db.refresh(registro)

        return registro

    finally:
        db.close()
