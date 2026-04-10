from sqlalchemy import Column, Integer, Text, Date, DateTime, Float, Boolean, JSON, ForeignKey
from database import Base
from datetime import datetime


class Usuario(Base):
    """Modelo para usuarios del sistema de conciliación."""

    __tablename__ = "usuarios"

    id = Column(Integer, primary_key=True)
    username = Column(Text, unique=True, index=True, nullable=False)
    nombre_completo = Column(Text, nullable=False)
    password_hash = Column(Text, nullable=False)
    rol = Column(Text, nullable=False)   # ADMIN | FINANZAS | LOGISTICA | VENTAS
    activo = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    last_login = Column(DateTime, nullable=True)

    def to_dict(self):
        return {
            "id": self.id,
            "username": self.username,
            "nombreCompleto": self.nombre_completo,
            "rol": self.rol,
            "activo": self.activo,
            "createdAt": self.created_at.isoformat() if self.created_at else None,
            "lastLogin": self.last_login.isoformat() if self.last_login else None,
        }


def safe_float(value):
    """Convierte un valor a float de manera segura."""
    if value is None:
        return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


class PedidoMeli(Base):
    """Modelo para la tabla pedidos_meli."""

    __tablename__ = "pedidos_meli"

    id = Column(Integer, primary_key=True, index=True)
    empresa = Column(Text)
    mov = Column(Text)
    movid = Column(Text, index=True)
    fechaemision = Column(Date)
    ultimocambio = Column(DateTime)
    usuario = Column(Text)
    referencia = Column(Text, index=True)
    observaciones = Column(Text)
    estatus = Column(Text, index=True)
    fechaoriginal = Column(Date)
    cliente = Column(Text)
    almacen = Column(Text)
    fecharequerida = Column(Date)
    vencimiento = Column(Date)
    importe = Column(Text)
    impuestos = Column(Text)
    saldo = Column(Text)
    descuentolineal = Column(Text)
    comisiontotal = Column(Text)
    costototal = Column(Text)
    origen = Column(Text)
    origenid = Column(Text)
    poliza = Column(Text)
    polizaid = Column(Text)
    generarpoliza = Column(Text)
    ejercicio = Column(Integer)
    periodo = Column(Integer)
    fecharegistro = Column(DateTime)
    fechaconclusion = Column(DateTime)
    fechacancelacion = Column(DateTime)
    fechaentrega = Column(DateTime)
    embarqueestado = Column(Text)
    atencion = Column(Text, index=True)
    listapreciosesp = Column(Text)
    sucursal = Column(Text)
    ivafiscal = Column(Text)
    formapagotipo = Column(Text)
    adicional1 = Column(Text)
    adicional2 = Column(Text)
    adicional3 = Column(Text)
    adicional4 = Column(Text)
    adicional5 = Column(Text, index=True)
    adicional6 = Column(Text)
    adicional7 = Column(Text)
    adicional8 = Column(Text)
    guia = Column(Text)
    proveedorvta = Column(Text)
    zonavta = Column(Text)
    adicional9 = Column(Text)
    adicional10 = Column(Text)
    adicional11 = Column(Text)
    adicional12 = Column(Text)
    gastocomision = Column(Text)
    gastoenvio = Column(Text)
    gastoempaque = Column(Text)
    gastologistico = Column(Text)
    gastoinversa = Column(Text)
    motivocanp = Column(Text)
    motivocanf = Column(Text)
    region = Column(Text)
    dataintelisis = Column(Text)
    urltracking = Column(Text)
    currierintelisis = Column(Text)
    embarque = Column(Text)
    embarque_observaciones = Column(Text)
    embarque_agente = Column(Text)
    embarque_estatus = Column(Text)

    def to_dict(self):
        """Convierte el modelo a diccionario."""
        importe = safe_float(self.importe)
        impuestos = safe_float(self.impuestos)
        return {
            "Mov": self.mov,
            "MovID": self.movid,
            "Estatus": self.estatus,
            "Importe": importe,
            "Impuestos": impuestos,
            "Saldo": safe_float(self.saldo),
            "PrecioTotal": importe + impuestos,
            "FechaEmision": self.fechaemision.isoformat() if self.fechaemision else None,
            "Cliente": self.cliente,
            "Referencia": self.referencia,
            "Atencion": self.atencion,
            "Embarque": self.embarque,
            "EmbarqueObservaciones": self.embarque_observaciones,
            "EmbarqueAgente": self.embarque_agente,
            "EmbarqueEstatus": self.embarque_estatus,
        }


class CasoSeguimiento(Base):
    """Modelo para casos de seguimiento del Portal de Conciliación."""

    __tablename__ = "casos_seguimiento"

    id = Column(Integer, primary_key=True, index=True)
    referencia = Column(Text, unique=True, index=True, nullable=False)
    tipo_problema = Column(Text, index=True)  # DEVOLUCION, EN_DISPUTA, CONTRACARGO, etc.
    estado_seguimiento = Column(Text, default='pendiente', index=True)  # pendiente, revision, resuelto
    monto_neto = Column(Float, default=0)
    monto_bruto = Column(Float, default=0)
    costos = Column(Float, default=0)
    notas = Column(Text, default='')
    fecha_deteccion = Column(DateTime, default=datetime.utcnow)
    fecha_ultima_actualizacion = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    en_ultimo_proceso = Column(Boolean, default=True)
    datos_orden = Column(JSON, default={})  # sourceIds, metodoPago, totalMovimientos, fechaLiberacion
    historial = Column(JSON, default=[])  # Lista de cambios

    def to_dict(self):
        """Convierte el modelo a diccionario."""
        return {
            "id": self.id,
            "referencia": self.referencia,
            "tipoProblema": self.tipo_problema,
            "estadoSeguimiento": self.estado_seguimiento,
            "montoNeto": self.monto_neto,
            "montoBruto": self.monto_bruto,
            "costos": self.costos,
            "notas": self.notas or '',
            "fechaDeteccion": self.fecha_deteccion.isoformat() if self.fecha_deteccion else None,
            "fechaUltimaActualizacion": self.fecha_ultima_actualizacion.isoformat() if self.fecha_ultima_actualizacion else None,
            "enUltimoProceso": self.en_ultimo_proceso,
            "datosOrden": self.datos_orden or {},
            "historial": self.historial or []
        }


class HistoricoConciliacion(Base):
    """Resumen/cabecera de cada conciliación ejecutada."""

    __tablename__ = "historico_conciliaciones"

    id = Column(Integer, primary_key=True, index=True)
    fecha = Column(DateTime, default=datetime.utcnow, index=True)
    total_ordenes = Column(Integer, default=0)
    coincidencias = Column(Integer, default=0)
    casos_creados = Column(Integer, default=0)
    ordenes_validadas = Column(Integer, default=0)
    monto_estado_cuenta = Column(Float, default=0)
    total_ingresos_bruto = Column(Float, default=0)
    total_costos = Column(Float, default=0)
    total_neto_real = Column(Float, default=0)
    resumen_por_estatus = Column(JSON, default={})
    notas = Column(Text, default='')

    def to_dict(self):
        return {
            "id": self.id,
            "fecha": self.fecha.isoformat() if self.fecha else None,
            "totalOrdenes": self.total_ordenes,
            "coincidencias": self.coincidencias,
            "casosCreados": self.casos_creados,
            "ordenesValidadas": self.ordenes_validadas,
            "montoEstadoCuenta": self.monto_estado_cuenta,
            "totalIngresosBruto": self.total_ingresos_bruto,
            "totalCostos": self.total_costos,
            "totalNetoReal": self.total_neto_real,
            "resumenPorEstatus": self.resumen_por_estatus or {},
            "notas": self.notas or ''
        }


class HistoricoOrden(Base):
    """Detalle de cada orden por conciliación."""

    __tablename__ = "historico_ordenes"

    id = Column(Integer, primary_key=True, index=True)
    conciliacion_id = Column(Integer, ForeignKey("historico_conciliaciones.id", ondelete="CASCADE"), index=True)
    referencia_erp = Column(Text, index=True)
    estatus = Column(Text, index=True)
    datos_orden = Column(JSON, default={})

    def to_dict(self):
        return {
            "id": self.id,
            "conciliacionId": self.conciliacion_id,
            "referenciaERP": self.referencia_erp,
            "estatus": self.estatus,
            "datosOrden": self.datos_orden or {}
        }
