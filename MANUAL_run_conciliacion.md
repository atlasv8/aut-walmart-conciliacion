# Manual de uso — Conciliación Walmart MX

Proceso mensual de conciliación de pagos **Walmart MX** para **Atlas del Descanso** y **Spring**.
Dos puntos de entrada según el alcance del reporte:

| Script | Uso |
|---|---|
| `run_consolidado.py` | **Consolidado ambas cuentas** (Atlas + Spring) → un solo Excel |
| `run_conciliacion.py` | Reporte **individual** para una cuenta |

---

## Tabla de contenidos

1. [Requisitos previos](#1-requisitos-previos)
2. [Archivos del sistema](#2-archivos-del-sistema)
3. [Configuración inicial — archivo `.env`](#3-configuración-inicial--archivo-env)
4. [Flujo mensual típico](#4-flujo-mensual-típico)
5. [Opciones disponibles](#5-opciones-disponibles)
6. [Modos de ejecución](#6-modos-de-ejecución)
7. [Hojas del Excel — Consolidado](#7-hojas-del-excel--consolidado)
8. [Hojas del Excel — Individual](#8-hojas-del-excel--individual)
9. [Clasificación contable (acciones ERP)](#9-clasificación-contable-acciones-erp)
10. [Solución de errores comunes](#10-solución-de-errores-comunes)

---

## 1. Requisitos previos

```bash
pip install pandas openpyxl requests python-dotenv
```

Python 3.10 o superior.

---

## 2. Archivos del sistema

Todos los archivos deben estar en la misma carpeta (`conciliacion/`):

| Archivo | Descripción |
|---|---|
| `run_consolidado.py` | **Consolidado Atlas + Spring** |
| `run_conciliacion.py` | Reporte individual por cuenta |
| `conciliacion_walmart.py` | Lectura y clasificación del CSV |
| `preparar_erp_walmart.py` | Enriquecimiento con PO# (API Walmart) y datos ERP |
| `.env` | Credenciales Walmart y ERP (no se sube a git) |

---

## 3. Configuración inicial — archivo `.env`

Crea el archivo `.env` en la carpeta `conciliacion/`:

```env
# Cuenta Atlas del Descanso
WALMART_CLIENT_ID=tu_client_id_aqui
WALMART_CLIENT_SECRET=tu_client_secret_aqui

# Cuenta Spring
WALMART_CLIENT_ID_SPRING=otro_client_id
WALMART_CLIENT_SECRET_SPRING=otro_client_secret

# ERP
ERP_BASE_URL=https://...
ERP_USERNAME=usuario
ERP_PASSWORD=contraseña
```

> Las credenciales Walmart se obtienen desde: Portal Marketplace → Configuración → Credenciales API.

---

## 4. Flujo mensual típico

```
1. Descarga el CSV de pagos de cada cuenta desde el portal Walmart Marketplace
2. Copia los CSV a la carpeta conciliacion/
3. Ejecuta run_consolidado.py
4. Revisa el Excel generado
```

### Comando principal

```bash
cd conciliacion
python run_consolidado.py --atlas Atlas_04-2026.csv --spring Spring_04-2026.csv
```

El script automáticamente:
1. Descarga las órdenes vía API Walmart (batch, últimos 120 días) para obtener PO#
2. Lee y procesa ambos CSV
3. Consulta el ERP Atlas del Descanso para cada cuenta
4. Genera el Excel `consolidado_<periodo>.xlsx`

---

## 5. Opciones disponibles

### `run_consolidado.py`

```
python run_consolidado.py --atlas <csv> --spring <csv> [opciones]
```

| Opción | Descripción | Default |
|---|---|---|
| `--atlas` | CSV de pagos cuenta Atlas del Descanso | (requerido) |
| `--spring` | CSV de pagos cuenta Spring | (requerido) |
| `--output` / `-o` | Nombre del Excel de salida | `consolidado_<periodo>.xlsx` |
| `--skip-erp` | Omite la consulta al ERP por completo | desactivado |

### `run_conciliacion.py`

```
python run_conciliacion.py <archivo_csv> [opciones]
```

| Opción | Descripción | Default |
|---|---|---|
| `csv` | Ruta al archivo CSV de pagos Walmart | (requerido) |
| `--output` / `-o` | Nombre del Excel de salida | `<csv>_reporte.xlsx` |
| `--skip-erp` | Omite la consulta al ERP por completo | desactivado |
| `--cuenta` | Cuenta Walmart: `atlas` o `spring` | `atlas` |

---

## 6. Modos de ejecución

### Consolidado (ambas cuentas)

```bash
python run_consolidado.py --atlas Atlas_04-2026.csv --spring Spring_04-2026.csv
```

### Consolidado sin ERP

```bash
python run_consolidado.py --atlas Atlas_04-2026.csv --spring Spring_04-2026.csv --skip-erp
```

### Individual — Atlas

```bash
python run_conciliacion.py Atlas_04-2026.csv
```

### Individual — Spring

```bash
python run_conciliacion.py Spring_04-2026.csv --cuenta spring
```

### Individual sin ERP

```bash
python run_conciliacion.py Atlas_04-2026.csv --skip-erp
```

### Archivo de salida personalizado

```bash
python run_consolidado.py --atlas Atlas_04-2026.csv --spring Spring_04-2026.csv --output Conciliacion_Abril2026.xlsx
```

---

## 7. Hojas del Excel — Consolidado

El archivo `consolidado_<periodo>.xlsx` contiene:

| Hoja | Contenido |
|---|---|
| **📊 CONSOLIDADO** | Resumen de acciones ERP + desglose por concepto, Atlas vs Spring lado a lado |
| **A Pagar Proveedor** | Ventas liquidadas de ambas cuentas (col. **Cuenta**: Atlas / Spring) |
| **Notas de Crédito** | Devoluciones de ambas cuentas (col. **Cuenta**) |
| **Gastos Plataforma** | Comisiones, SEM, WFS, Killer de ambas cuentas (col. **Cuenta**) |
| **Retenciones** | IVA e ISR de ambas cuentas (col. **Cuenta**) |
| **ERP_Conciliacion** | CSV vs ERP por orden ✅/⚠/❌, ambas cuentas (col. **Cuenta**) — *solo si ERP activo* |
| **Pendientes** | Seller sin PO# de ambas cuentas (col. **Cuenta**) — *solo si hay pendientes* |
| **Sin Mapeo** | Movimientos no clasificados — *solo si existen* |

---

## 8. Hojas del Excel — Individual

El archivo `<csv>_reporte.xlsx` contiene:

| Hoja | Contenido |
|---|---|
| **📊 Resumen Ejecutivo** | Resumen de acciones ERP y desglose por concepto |
| **💰 A Pagar Proveedor** | Ventas liquidadas por Walmart |
| **🔄 Notas de Crédito** | Devoluciones de clientes |
| **📉 Gastos Plataforma** | Comisiones, SEM, WFS, Killer |
| **🏛 Retenciones Fiscales** | IVA e ISR retenidos/devueltos (col. Income/Outcome indica dirección) |
| **📋 Detalle Completo** | Todos los registros enriquecidos con ERP_ID y clasificación |
| **🔗 ERP_Conciliacion** | Resumen por orden: CSV vs ERP ✅/⚠/❌ — *solo si ERP activo* |
| **⚠ ERP_Pendientes** | Seller sin PO# — requieren seguimiento manual — *solo si hay pendientes* |
| **🔍 Sin Mapeo** | Movimientos no clasificados — *solo si existen* |

---

## 9. Clasificación contable (acciones ERP)

| Acción | Color | Descripción |
|---|---|---|
| **A pagar proveedor** | Verde | Venta liquidada por Walmart; el marketplace paga al proveedor |
| **Nota de crédito** | Naranja | Devolución de cliente; revierte la venta |
| **Gasto de plataforma** | Amarillo | Comisiones, WFS, SEM, Killer — costos operativos deducibles |
| **Retención fiscal** | Azul | IVA e ISR retenidos; Ingreso = devolución por reembolso, Egreso = retención normal |

Registros sin clasificación aparecen en gris como **Sin clasificar** → revisar si Walmart agregó nuevas categorías y actualizar el mapeo en `conciliacion_walmart.py`.

---

## 10. Solución de errores comunes

### `ERROR: No se encontró el CSV`

Verifica la ruta o cambia al directorio correcto:

```bash
cd C:\Users\sistemasgca\aut_walmart\conciliacion
python run_consolidado.py --atlas Atlas_04-2026.csv --spring Spring_04-2026.csv
```

---

### `ERROR: pip install pandas openpyxl requests python-dotenv`

```bash
pip install pandas openpyxl requests python-dotenv
```

Si usas entorno virtual, actívalo primero:

```bash
.venv\Scripts\activate
pip install pandas openpyxl requests python-dotenv
```

---

### `❌ Credenciales no configuradas`

El archivo `.env` no existe o le faltan variables. Verifica que esté en la carpeta `conciliacion/` con las 4 variables de Walmart y las 3 del ERP.

---

### `❌ Error HTTP 401`

Token de API expirado. El script lo renueva automáticamente; si persiste, verifica que las credenciales en `.env` sean correctas y estén vigentes en el portal Walmart.

---

### `⚠ N registros Seller sin PO#`

La API no encontró el número de orden de compra para algunos pedidos Seller Fulfillment. El script ya consulta los últimos 120 días en batch. Si persiste, localiza el PO# manualmente en el portal Walmart y complétalo en el ERP.

---

### `⚠ Movimientos sin mapeo`

El CSV contiene tipos de transacción no reconocidos. Aparecen en la hoja **Sin Mapeo**. Revisar si Walmart agregó nuevas categorías y actualizar el mapeo en `conciliacion_walmart.py`.
