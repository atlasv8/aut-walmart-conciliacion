from pydantic_settings import BaseSettings
from functools import lru_cache
import os


class Settings(BaseSettings):
    """Configuración de la aplicación."""

    # Base de datos PostgreSQL
    db_host: str = "localhost"
    db_port: int = 5432
    db_name: str
    db_user: str
    db_password: str

    # API
    api_port: int = 8000

    # JWT
    jwt_secret: str = "change-me-in-production"
    jwt_algorithm: str = "HS256"
    jwt_expire_hours: int = 8

    # Correo Gmail
    email_sender: str = ""
    email_password: str = ""
    email_recipients: str = ""  # CSV separado por comas

    # MercadoLibre token service
    ml_token_api_key: str = ""

    class Config:
        env_file = ".env"
        case_sensitive = False
        env_file_encoding = 'utf-8'

    @property
    def database_url(self) -> str:
        """Construye la URL de conexión a PostgreSQL."""
        return f"postgresql://{self.db_user}:{self.db_password}@{self.db_host}:{self.db_port}/{self.db_name}"


@lru_cache()
def get_settings() -> Settings:
    """
    Retorna la instancia singleton de configuración.
    FUERZA la lectura del archivo .env e ignora variables de entorno del sistema.
    """
    # Guardar variables de entorno originales
    original_env = {}
    db_vars = ['DB_HOST', 'DB_PORT', 'DB_NAME', 'DB_USER', 'DB_PASSWORD', 'API_PORT',
               'JWT_SECRET', 'JWT_ALGORITHM', 'JWT_EXPIRE_HOURS',
               'EMAIL_SENDER', 'EMAIL_PASSWORD', 'EMAIL_RECIPIENTS', 'ML_TOKEN_API_KEY']

    for var in db_vars:
        if var in os.environ:
            original_env[var] = os.environ[var]
            # Temporalmente eliminar del entorno
            del os.environ[var]

    try:
        # Crear settings solo desde el archivo .env
        settings = Settings(_env_file='.env', _env_file_encoding='utf-8')
        return settings
    finally:
        # Restaurar variables de entorno originales
        for var, value in original_env.items():
            os.environ[var] = value
