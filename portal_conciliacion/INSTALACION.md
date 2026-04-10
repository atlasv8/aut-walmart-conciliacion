# Guía Rápida de Instalación

## Paso a Paso

### 1. Instalar Python (si no lo tienes)

Descarga Python 3.8+ desde: https://www.python.org/downloads/

Durante la instalación, marca la casilla "Add Python to PATH"

### 2. Crear entorno virtual

Abre una terminal/cmd en la carpeta del proyecto:

```bash
python -m venv .venv
```

### 3. Activar entorno virtual

**Windows:**
```bash
.venv\Scripts\activate
```

**Linux/Mac:**
```bash
source .venv/bin/activate
```

Deberías ver `(.venv)` al inicio de tu línea de comando.

### 4. Instalar dependencias

```bash
pip install -r requirements.txt
```

Esto instalará:
- FastAPI (framework web)
- Uvicorn (servidor ASGI)
- SQLAlchemy (ORM para base de datos)
- psycopg2 (driver de PostgreSQL)
- python-dotenv (manejo de variables de entorno)
- pydantic (validación de datos)

### 5. Configurar base de datos

1. Copia el archivo de ejemplo:
   ```bash
   copy .env.example .env
   ```

2. Edita `.env` con tus credenciales de PostgreSQL:
   ```env
   DB_HOST=localhost
   DB_PORT=5432
   DB_NAME=tu_base_de_datos
   DB_USER=tu_usuario
   DB_PASSWORD=tu_contraseña
   API_PORT=8000
   ```

### 6. Probar conexión

Verifica que todo está bien configurado:

```bash
python test_connection.py
```

Si ves "✓ CONEXIÓN EXITOSA", todo está listo.

### 7. Iniciar el servidor

**Opción A - Script automático (recomendado):**

Windows:
```bash
start.bat
```

Linux/Mac:
```bash
chmod +x start.sh
./start.sh
```

**Opción B - Manual:**
```bash
python main.py
```

El servidor estará en: http://localhost:8000

### 8. Usar la aplicación

1. Abre `CONCILIADOR.html` en tu navegador
2. Arrastra o selecciona tu archivo Excel de MercadoLibre
3. Haz clic en "Conciliar con BD"
4. Descarga el resultado

## Verificar que funciona

Abre en tu navegador: http://localhost:8000

Deberías ver:
```json
{
  "message": "API Conciliador MercadoLibre",
  "status": "running",
  "version": "1.0.0"
}
```

## Problemas Comunes

### Error: "python no se reconoce como comando"

Python no está en el PATH. Opciones:
1. Reinstala Python marcando "Add to PATH"
2. Usa la ruta completa: `C:\Python39\python.exe`

### Error: "pip no se reconoce como comando"

```bash
python -m pip install -r requirements.txt
```

### Error: "Access is denied"

En Windows, ejecuta como Administrador o:
```bash
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Error: "Could not connect to database"

1. Verifica que PostgreSQL está ejecutándose
2. Revisa usuario/contraseña en `.env`
3. Prueba conexión manual:
   ```bash
   psql -h localhost -U tu_usuario -d tu_base_datos
   ```

### Error: "ModuleNotFoundError"

Asegúrate de:
1. Activar el entorno virtual: `.venv\Scripts\activate`
2. Instalar dependencias: `pip install -r requirements.txt`

## Detener el Servidor

Presiona `Ctrl+C` en la terminal donde está corriendo.

## Desactivar entorno virtual

```bash
deactivate
```

## Actualizar dependencias

Si cambias algo en el código:

```bash
pip install -r requirements.txt --upgrade
```

## Próximos Pasos

Lee el archivo `README.md` para más detalles sobre:
- Estructura de la base de datos
- Endpoints disponibles
- Estados de órdenes
- Solución de problemas avanzados
