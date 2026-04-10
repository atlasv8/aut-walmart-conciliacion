# Conciliador MercadoLibre - Intelisis

Sistema de conciliacion de pagos entre MercadoLibre y el ERP Intelisis. Permite procesar archivos de liberaciones, identificar discrepancias, gestionar casos pendientes y generar reportes de pagos CxC.

## Tabla de Contenidos

- [Requisitos](#requisitos)
- [Instalacion](#instalacion)
- [Configuracion](#configuracion)
- [Estructura del Proyecto](#estructura-del-proyecto)
- [Archivos de Entrada](#archivos-de-entrada)
- [Reglas de Negocio](#reglas-de-negocio)
- [Flujos de Trabajo](#flujos-de-trabajo)
- [API Endpoints](#api-endpoints)
- [Uso del Sistema](#uso-del-sistema)
- [Base de Datos](#base-de-datos)

---

## Requisitos

### Backend (Python)
- Python 3.9+
- PostgreSQL 12+

### Frontend
- Navegador moderno (Chrome, Firefox, Edge)
- Conexion al backend en localhost:8000

---

## Instalacion

### 1. Clonar el repositorio

```bash
git clone <url-del-repositorio>
cd conciliador
```

### 2. Crear entorno virtual e instalar dependencias

```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

pip install -r requirements.txt
```

### 3. Configurar variables de entorno

Crear archivo `.env` en la raiz del proyecto:

```env
DB_HOST=localhost
DB_PORT=5432
DB_NAME=nombre_base_datos
DB_USER=usuario
DB_PASSWORD=contraseña
API_PORT=8000
```

### 4. Iniciar el backend

```bash
python main.py
```

El servidor se iniciara en `http://localhost:8000`

### 5. Abrir el frontend

Abrir `index.html` en un navegador web.

---

## Configuracion

### Archivo `config.py`

| Variable | Descripcion | Default |
|----------|-------------|---------|
| `DB_HOST` | Host de PostgreSQL | localhost |
| `DB_PORT` | Puerto de PostgreSQL | 5432 |
| `DB_NAME` | Nombre de la base de datos | - |
| `DB_USER` | Usuario de base de datos | - |
| `DB_PASSWORD` | Contraseña de base de datos | - |
| `API_PORT` | Puerto de la API | 8000 |

### Configuracion del Frontend (`app.js`)

| Parametro | Valor | Descripcion |
|-----------|-------|-------------|
| `toleranciaNeto` | 0.05 | Tolerancia en pesos para considerar montos iguales |
| `toleranciaSaldo` | 1.00 | Tolerancia para validacion de saldos |
| `tasaIVA` | 0.16 | Tasa de IVA (16%) |

---

## Estructura del Proyecto

```
conciliador/
├── main.py              # API FastAPI principal
├── models.py            # Modelos SQLAlchemy
├── database.py          # Configuracion de base de datos
├── config.py            # Configuracion de la aplicacion
├── index.html           # Frontend principal
├── styles.css           # Estilos CSS
├── app.js               # Logica JavaScript del frontend
├── requirements.txt     # Dependencias Python
├── .env                 # Variables de entorno (crear)
└── README.md            # Este archivo
```

---

## Archivos de Entrada

El sistema requiere **dos archivos Excel** de MercadoLibre:

### Archivo 1: Liberaciones ML

Reporte de liberaciones de pagos con las siguientes columnas:

| Columna | Descripcion |
|---------|-------------|
| `SOURCE_ID` | ID unico de la transaccion en ML |
| `EXTERNAL_REFERENCE` | Referencia externa (ID de venta) |
| `GROSS_AMOUNT` | Monto bruto de la venta |
| `NET_CREDIT_AMOUNT` | Monto neto a recibir |
| `NET_DEBIT_AMOUNT` | Debitos (devoluciones, comisiones) |
| `MP_FEE_AMOUNT` | Comision de MercadoPago |
| `SHIPPING_FEE_AMOUNT` | Costo de envio |
| `FINANCING_FEE_AMOUNT` | Costo de financiamiento |
| `TAXES_AMOUNT` | Impuestos |
| `PAYMENT_METHOD` | Metodo de pago |
| `DESCRIPTION` | Tipo de movimiento (payment, refund, etc.) |
| `RELEASE_DATE` | Fecha de liberacion |
| `APPROVAL_DATE` | Fecha de aprobacion |

### Archivo 2: Estado de Cuenta ML

Reporte del estado de cuenta con las siguientes columnas:

| Columna | Descripcion |
|---------|-------------|
| `REFERENCE_ID` | ID de referencia (SOURCE_ID) |
| `TRANSACTION_TYPE` | Tipo de transaccion |
| `TRANSACTION_NET_AMOUNT` | Monto neto de la transaccion |

---

## Reglas de Negocio

### Estatus de Ordenes

El sistema clasifica automaticamente cada orden segun su estado:

| Estatus | Descripcion | Accion Requerida |
|---------|-------------|------------------|
| **PAGADO** | Orden pagada correctamente | Ninguna - procesar normalmente |
| **DEVOLUCION** | Cliente devolvio el producto | Cancelar factura, registrar NC |
| **EN_DISPUTA** | Disputa abierta con el cliente | Esperar resolucion de ML |
| **CONTRACARGO** | Contracargo del banco | Gestionar con finanzas |
| **MEDIACION_GANADA** | Disputa resuelta a favor | Verificar monto recibido |
| **MEDIACION_PERDIDA** | Disputa resuelta en contra | Cancelar factura |
| **SIN_COBRO** | Sin transaccion de cobro | Verificar en ML |
| **REVISAR** | No clasificable automaticamente | Revision manual |

### Prioridades de Casos

| Prioridad | Criterio | Color |
|-----------|----------|-------|
| **URGENTE** | Monto > $5,000 o contracargo | Rojo |
| **IMPORTANTE** | Monto > $0 y < $5,000 | Amarillo |
| **BAJA** | Revisar informacion | Verde |

### Asignacion de Responsables

| Tipo de Caso | Responsable |
|--------------|-------------|
| Devoluciones | Logistica |
| Discrepancias | Finanzas |
| Contracargos | Finanzas |
| Anticipos CxC | Finanzas |
| Cobros a Factura | Finanzas |
| Pedidos no encontrados | Finanzas |

### Reglas de Pagos CxC

El sistema identifica automaticamente pagos para Cuentas por Cobrar:

**Anticipo CxC** (Solo Pedido):
- Pedido en estado "Pendiente" en Intelisis
- Sin factura generada
- Pago recibido en MercadoLibre

**Cobro a Factura**:
- Factura en estado "Concluida" en Intelisis
- Saldo pendiente en la factura
- Pago recibido en MercadoLibre

---

## Flujos de Trabajo

### DEVOLUCION_PENDIENTE
**Devolucion con factura concluida - requiere cancelacion contable**

SLA: 24 horas | Responsable: Logistica

Pasos:
1. Confirmar recepcion fisica de mercancia en almacen (Obligatorio)
2. Verificar condicion de la mercancia recibida (Obligatorio)
3. Cancelar factura en Intelisis (Obligatorio)
4. Registrar Nota de Credito correspondiente (Obligatorio)
5. Actualizar inventario con mercancia devuelta (Obligatorio)
6. Notificar a finanzas sobre NC generada (Opcional)

Documentos: Factura, Guia de devolucion ML

---

### DEVOLUCION_ORDEN_DEVUELTA
**Orden devuelta sin facturar - solo cancelar pedido**

SLA: 12 horas | Responsable: Logistica

Pasos:
1. Verificar estado de cancelacion en MercadoLibre (Obligatorio)
2. Confirmar que NO se genero factura (Obligatorio)
3. Cancelar pedido en Intelisis (Obligatorio)
4. Documentar razon de la devolucion (Opcional)

Documentos: Comprobante de cancelacion ML

---

### DEVOLUCION_EN_PROCESO
**Cliente menciono devolucion pero orden aun activa**

SLA: 48 horas | Responsable: Logistica

Pasos:
1. Contactar al cliente para confirmar intencion (Obligatorio)
2. Verificar estado actual en MercadoLibre (Obligatorio)
3. Esperar confirmacion formal de devolucion (Obligatorio)
4. Re-evaluar cuando cambie el estado en ML (Obligatorio)

---

### DISCREPANCIA
**Diferencia de montos entre ML e Intelisis**

SLA: 48 horas | Responsable: Finanzas

Pasos:
1. Comparar factura vs orden de MercadoLibre linea por linea (Obligatorio)
2. Verificar ajustes, descuentos o cupones en ML (Obligatorio)
3. Revisar comisiones de MercadoLibre aplicadas (Obligatorio)
4. Si persiste diferencia, contactar soporte ML (Opcional)
5. Ajustar factura en Intelisis o registrar NC/ND segun corresponda (Obligatorio)

Documentos: Factura, Detalle de orden ML, Comprobante de comisiones

---

### ANTICIPO
**Cliente pago pero aun no se factura - registrar anticipo CxC**

SLA: 24 horas | Responsable: Finanzas

Pasos:
1. Verificar forma de cobro en MercadoLibre (Obligatorio)
2. Crear anticipo en modulo CxC de Intelisis (Obligatorio)
3. Vincular anticipo con pedido pendiente (Obligatorio)
4. Programar aplicacion automatica al generar factura (Obligatorio)
5. Verificar que referencia ML quede registrada (Obligatorio)

Documentos: Comprobante de pago ML, Pedido pendiente

---

### COBRO_FACTURA
**Factura concluida y pago recibido - registrar cobro**

SLA: 24 horas | Responsable: Finanzas

Pasos:
1. Verificar factura concluida en Intelisis (Obligatorio)
2. Confirmar pago recibido en MercadoLibre (Obligatorio)
3. Registrar cobro en modulo CxC (Obligatorio)
4. Aplicar cobro a saldo pendiente de la factura (Obligatorio)
5. Verificar que saldo quede en cero (Obligatorio)

Documentos: Factura, Comprobante de pago ML

---

### PENDIENTE_ML
**Orden pagada en ML pero sin pedido en Intelisis**

SLA: 12 horas | Responsable: Finanzas

Pasos:
1. Verificar que orden este realmente pagada en ML (Obligatorio)
2. Buscar si pedido existe con otro numero/formato (Obligatorio)
3. Si no existe: crear pedido en Intelisis (Obligatorio)
4. Vincular orden ML con pedido creado (Obligatorio)
5. Proceder con facturacion normal (Obligatorio)

Documentos: Orden completa de ML

---

## API Endpoints

### Consultas de Ventas

#### `GET /ventas/venta-id/{venta_id}`
Busca pedidos por EXTERNAL_REFERENCE (ID de venta de ML).

**Parametros:**
- `venta_id`: ID de venta de MercadoLibre

**Respuesta:** Lista de pedidos encontrados en Intelisis

---

#### `GET /ventas/atencion/{atencion_id}`
Busca pedidos por campo atencion (EXTERNAL_REFERENCE o SOURCE_ID).

**Parametros:**
- `atencion_id`: Identificador en campo atencion

**Respuesta:** Lista de pedidos encontrados

---

### Portal de Seguimiento

#### `GET /seguimiento/casos`
Lista todos los casos de seguimiento.

**Query params opcionales:**
- `estado`: pendiente | revision | resuelto | todos
- `tipo`: DEVOLUCION | DISCREPANCIA | etc.
- `en_proceso`: true | false

---

#### `GET /seguimiento/casos/{referencia}`
Obtiene un caso especifico por referencia.

---

#### `POST /seguimiento/casos`
Crea un nuevo caso de seguimiento.

**Body:**
```json
{
  "referencia": "12345-001",
  "tipoProblema": "DEVOLUCION",
  "estadoSeguimiento": "pendiente",
  "montoNeto": 1500.00,
  "montoBruto": 1740.00,
  "costos": 240.00,
  "notas": "",
  "enUltimoProceso": true,
  "datosOrden": {},
  "historial": []
}
```

---

#### `PUT /seguimiento/casos/{referencia}`
Actualiza un caso existente.

---

#### `DELETE /seguimiento/casos/{referencia}`
Elimina un caso de seguimiento.

---

#### `POST /seguimiento/sincronizar`
Sincroniza casos desde el frontend.

- Crea casos nuevos
- Actualiza casos existentes (conserva notas y estado)
- Marca casos no presentes como fuera de proceso

---

#### `GET /seguimiento/estadisticas`
Obtiene estadisticas del portal.

**Respuesta:**
```json
{
  "total": 50,
  "pendientes": 30,
  "revision": 15,
  "resueltos": 5
}
```

---

### Envio de Correos

#### `POST /enviar-correo-pagos`
Envia correo con reporte de pagos de MercadoLibre.

**Body:**
```json
{
  "pagos": [...],
  "resumen": {
    "total_ordenes": 10,
    "total_venta": 50000.00,
    "total_neto": 45000.00,
    "fecha": "2024-01-15"
  }
}
```

---

#### `POST /enviar-correo-pagos-cxc`
Envia correo con reporte de pagos CxC (anticipos y cobros a factura).

---

### Health Check

#### `GET /health`
Verifica el estado del servicio y conexion a base de datos.

---

## Uso del Sistema

### 1. Dashboard
Vista principal con:
- Resumen de ultima conciliacion
- Casos activos y su estado
- Tasa de resolucion
- Monto en discrepancias
- Lista de casos urgentes
- Distribucion por responsable

### 2. Nueva Conciliacion
1. Cargar **Archivo 1** (Liberaciones ML)
2. Cargar **Archivo 2** (Estado de Cuenta ML)
3. Click en **Procesar Conciliacion**
4. Revisar resultados y estadisticas
5. Click en **Conciliar con Intelisis** para validar contra el ERP

### 3. Portal de Casos
Panel lateral izquierdo con:
- Contadores por prioridad (Urgentes, Importantes, Revisar)
- Filtros por responsable (Todos, Contador, Finanzas, Logistica)
- Lista de casos pendientes
- Click en un caso para ver detalle

### 4. Detalle de Caso
Modal con:
- Informacion del caso (referencia, montos, estado)
- Informacion del ERP (si se encontro)
- Checklist de pasos a seguir
- Documentos requeridos
- Historial de acciones
- Campo de notas
- Botones: Reasignar, Marcar Resuelto

### 5. Pagos CxC
Despues de conciliar:
- Ver total de pagos identificados
- Descargar Excel con pagos CxC
- Enviar reporte por correo

### 6. Exportar/Importar
- **Exportar**: Descarga casos en formato JSON
- **Importar**: Carga casos desde archivo JSON

---

## Base de Datos

### Tabla: `pedidos_meli`
Almacena los pedidos sincronizados desde Intelisis.

| Campo | Tipo | Descripcion |
|-------|------|-------------|
| id | Integer | Primary key |
| empresa | Text | Codigo de empresa |
| mov | Text | Tipo de movimiento |
| movid | Text | ID del movimiento |
| fechaemision | Date | Fecha de emision |
| referencia | Text | Referencia (EXTERNAL_REF-secuencia) |
| estatus | Text | Estado del documento |
| cliente | Text | Nombre del cliente |
| importe | Text | Importe sin IVA |
| impuestos | Text | Monto de impuestos |
| saldo | Text | Saldo pendiente |
| atencion | Text | Campo atencion (EXTERNAL_REF o SOURCE_ID) |
| adicional5 | Text | SOURCE_IDs asociados |

### Tabla: `casos_seguimiento`
Almacena los casos del portal de seguimiento.

| Campo | Tipo | Descripcion |
|-------|------|-------------|
| id | Integer | Primary key |
| referencia | Text | Referencia unica del caso |
| tipo_problema | Text | DEVOLUCION, DISCREPANCIA, etc. |
| estado_seguimiento | Text | pendiente, revision, resuelto |
| monto_neto | Float | Monto neto del caso |
| monto_bruto | Float | Monto bruto |
| costos | Float | Costos operativos |
| notas | Text | Notas del usuario |
| fecha_deteccion | DateTime | Fecha de creacion |
| fecha_ultima_actualizacion | DateTime | Ultima modificacion |
| en_ultimo_proceso | Boolean | Si aparece en ultima conciliacion |
| datos_orden | JSON | Datos adicionales de la orden |
| historial | JSON | Historial de cambios |

---

## Configuracion de Correo

El sistema envia correos via SMTP de Gmail. Configurar en `main.py`:

```python
EMAIL_CONFIG = {
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': 465,
    'sender': 'correo@gmail.com',
    'password': 'contraseña_de_aplicacion',
    'recipients': ['destinatario1@empresa.com', 'destinatario2@empresa.com']
}
```

**Nota:** Para Gmail, usar una "Contrasena de aplicacion" en lugar de la contrasena normal.

---

## Solucion de Problemas

### Error de conexion a base de datos
- Verificar que PostgreSQL este corriendo
- Verificar credenciales en `.env`
- Verificar que la base de datos exista

### Error al cargar archivos Excel
- Verificar que los archivos tengan las columnas requeridas
- Verificar que no esten corruptos
- Verificar codificacion (UTF-8)

### Casos no se sincronizan
- Verificar conexion al backend (localhost:8000)
- Revisar consola del navegador para errores
- Verificar que el endpoint `/seguimiento/casos` responda

### Correos no se envian
- Verificar configuracion SMTP
- Verificar que la contrasena de aplicacion sea correcta
- Revisar logs del backend

---

## Licencia

Uso interno - GC Atlas
