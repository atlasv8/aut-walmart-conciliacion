import pyodbc
import pandas as pd
import os
from sqlalchemy import create_engine, text
from datetime import datetime

# ---- CONFIGURACIÓN SQL SERVER (Origen) ----
server = os.getenv('ERP_SERVER')
database = os.getenv('ERP_DATABASE')
username = os.getenv('ERP_USERNAME')
password = os.getenv('ERP_PASSWORD')

# ---- CONFIGURACIÓN POSTGRESQL (Destino) ----
pg_host = os.getenv('PG_HOST', 'localhost')
pg_port = os.getenv('PG_PORT', '5432')
pg_database = os.getenv('PG_DATABASE', 'postgres')
pg_username = os.getenv('PG_USERNAME', 'postgres')
pg_password = os.getenv('PG_PASSWORD', 'Gcatlas03')

# ---- CONEXIÓN SQL SERVER ----
conn_sqlserver = pyodbc.connect(
    f"DRIVER={{SQL Server}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"UID={username};"
    f"PWD={password}"
)

# ---- CONEXIÓN POSTGRESQL ----
pg_engine = create_engine(
    f'postgresql://{pg_username}:{pg_password}@{pg_host}:{pg_port}/{pg_database}'
)

query = """
SELECT  
    -- Identificadores
    V.ID,
    V.Empresa,
    V.Mov,
    V.MovID,

    -- Fechas
    V.FechaEmision,
    V.UltimoCambio,
    V.FechaOriginal,
    V.FechaRequerida,
    V.Vencimiento,
    V.FechaRegistro,
    V.FechaConclusion,
    V.FechaCancelacion,
    V.FechaEntrega,

    -- Usuario / Cliente
    V.Usuario,
    V.Cliente,
    V.Atencion,

    -- Referencias y estatus
    V.Referencia,
    V.Observaciones,
    V.Estatus,
    V.EmbarqueEstado,

    -- Importes encabezado
    V.Importe,
    V.Impuestos,
    V.Saldo,
    V.DescuentoLineal,
    V.ComisionTotal,
    V.CostoTotal,

    -- Detalle de venta (VentaDCalc)
    VD.Articulo,
    VD.Cantidad,
    VD.Importe        AS ImporteDetalle,
    VD.PrecioTotal    AS PrecioTotalDetalle,

    -- Origen / pólizas
    V.Origen,
    V.OrigenID,
    V.Poliza,
    V.PolizaID,
    V.GenerarPoliza,

    -- Ejercicio / periodo
    V.Ejercicio,
    V.Periodo,

    -- Almacén / sucursal
    V.Almacen,
    V.Sucursal,
    S.Region,

    -- Fiscales / pago
    V.IVAFiscal,
    V.FormaPagoTipo,
    V.ListaPreciosEsp,

    -- Adicionales
    V.Adicional1,
    V.Adicional2,
    V.Adicional3,
    V.Adicional4,
    V.Adicional5,
    V.Adicional6,
    V.Adicional7,
    V.Adicional8,
    V.Adicional9,
    V.Adicional10,
    V.Adicional11,
    V.Adicional12,

    -- Logística / guías
    V.Guia,
    V.ProveedorVta,
    V.ZonaVta,

    -- Gastos
    V.GastoComision,
    V.GastoEnvio,
    V.GastoEmpaque,
    V.GastoLogistico,
    V.GastoInversa,

    -- Cancelaciones
    V.MotivoCanP,
    V.MotivoCanF,

    -- Intelisis / tracking
    V.DataIntelisis,
    V.UrlTracking,
    V.CurrierIntelisis,

    -- Embarque más reciente (solo aplica para V.Mov = 'Pedido')
    EM.embarque,
    EM.embarque_observaciones,
    EM.embarque_estatus,
    EM.embarque_agente

FROM Venta V

LEFT JOIN Sucursal S 
    ON V.Sucursal = S.Sucursal

LEFT JOIN VentaDCalc VD 
    ON V.ID = VD.ID

-- Subquery: trae el embarque más reciente por MovID (via EmbarqueMov -> Embarque)
LEFT JOIN (
    SELECT 
        ed.MovID,
        e.MovID AS embarque,
        e.ID             AS embarqueid,
        e.Observaciones  AS embarque_observaciones,
        e.Estatus        AS embarque_estatus,
        e.Agente         AS embarque_agente
    FROM EmbarqueMov ed
    INNER JOIN (
        SELECT MovID, MAX(ID) AS MaxID
        FROM EmbarqueMov
        GROUP BY MovID
    ) ult ON ed.MovID = ult.MovID AND ed.ID = ult.MaxID
    INNER JOIN Embarque e ON e.ID = ed.AsignadoID
) EM ON V.MovID = EM.MovID AND V.Mov = 'Pedido'

WHERE 
    V.FechaEmision BETWEEN DATEADD(DAY, -200, GETDATE()) AND GETDATE()
    AND S.Region IN ('INTERNET', '.COM')
    AND V.Estatus IN ('CONCLUIDO', 'PENDIENTE', 'CANCELADO');
"""

print("Ejecutando consulta en SQL Server...")
df = pd.read_sql(query, conn_sqlserver)

print(f"Registros obtenidos: {len(df)}")

# ---- LIMPIEZA DE DATOS (opcional pero recomendado) ----
# Convertir nombres de columnas a minúsculas para PostgreSQL
df.columns = df.columns.str.lower()

# Reemplazar NaN con None para PostgreSQL
df = df.where(pd.notnull(df), None)

# ---- INSERTAR EN POSTGRESQL ----
try:
    # Nombre de la tabla en PostgreSQL
    tabla_destino = 'pedidos_meli'

    print(f"Preparando inserción en PostgreSQL (tabla: {tabla_destino})...")

    # TRUNCATE TABLE antes de insertar
    with pg_engine.connect() as conn:
        print(f"Ejecutando TRUNCATE TABLE {tabla_destino}...")
        conn.execute(text(f"TRUNCATE TABLE {tabla_destino}"))
        conn.commit()
        print(f"✓ Tabla {tabla_destino} limpiada exitosamente")

    print(f"Insertando datos en PostgreSQL...")

    df.to_sql(
        tabla_destino,
        pg_engine,
        if_exists='append',  # Cambiado a 'append' ya que ahora hacemos TRUNCATE manual
        index=False,
        chunksize=1000
    )

    print(f"✓ Datos insertados exitosamente en PostgreSQL")
    print(f"  Tabla: {tabla_destino}")
    print(f"  Registros: {len(df)}")

    # ---- REGISTRAR LOG DE EJECUCIÓN (opcional) ----
    with pg_engine.connect() as conn:
        log_query = text("""
            INSERT INTO logs_etl (tabla, registros, fecha_ejecucion, estatus)
            VALUES (:tabla, :registros, :fecha, :estatus)
        """)

        # Esta tabla debe existir previamente o puedes omitir este paso
        try:
            conn.execute(log_query, {
                'tabla': tabla_destino,
                'registros': len(df),
                'fecha': datetime.now(),
                'estatus': 'SUCCESS'
            })
            conn.commit()
        except Exception as e:
            print(f"Nota: No se pudo registrar log (tabla logs_etl podría no existir): {e}")

except Exception as e:
    print(f"✗ Error al insertar en PostgreSQL: {e}")
    raise

finally:
    # ---- CERRAR CONEXIONES ----
    conn_sqlserver.close()
    pg_engine.dispose()
    print("\nConexiones cerradas")

# ---- OPCIONAL: GUARDAR CSV DE RESPALDO ----
output_file = "pedidos_meli.csv"
df.to_csv(output_file, index=False, encoding='utf-8-sig')
print(f"CSV de respaldo guardado: {output_file}")