import configparser
import os
import logging

logger = logging.getLogger(__name__)

class Config:
    """
    Clase para leer y gestionar la configuración desde el archivo config.ini.
    Provee una interfaz similar a un diccionario para acceder a las secciones
    y valores de configuración.
    """
    _instance = None

    def __new__(cls, *args, **kwargs):
        # Implementamos un Singleton para asegurar que la configuración se lea una sola vez.
        if not cls._instance:
            cls._instance = super(Config, cls).__new__(cls)
        return cls._instance

    def __init__(self, filename='config.ini'):
        # El constructor se llamará múltiples veces, pero el `_initialized`
        # previene la reinicialización.
        if hasattr(self, '_initialized') and self._initialized:
            return
            
        self.config = configparser.ConfigParser()
        
        # La ruta del config.ini se resuelve relativa a la ubicación de este archivo.
        # Asumimos que config.ini está en el directorio padre de `src`.
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        config_path = os.path.join(base_dir, filename)

        if not os.path.exists(config_path):
            logger.error(f"¡Archivo de configuración no encontrado en '{config_path}'!")
            raise FileNotFoundError(f"No se encontró el archivo de configuración: {config_path}")
            
        self.config.read(config_path)
        logger.info(f"Configuración cargada desde '{config_path}'. Secciones encontradas: {self.config.sections()}")
        self._initialized = True

    def __getitem__(self, section):
        """
        Permite acceder a una sección de la configuración como si fuera un diccionario.
        Ej: config['database_local']
        """
        return self.config[section]

    def get(self, section, option, fallback=None):
        """
        Obtiene un valor específico de una sección, con un valor por defecto opcional.
        """
        return self.config.get(section, option, fallback=fallback)

    def getint(self, section, option, fallback=None):
        """
        Obtiene un valor específico como un entero.
        """
        return self.config.getint(section, option, fallback=fallback)

    def getboolean(self, section, option, fallback=None):
        """
        Obtiene un valor específico como un booleano.
        """
        return self.config.getboolean(section, option, fallback=fallback)

