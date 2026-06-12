# Estado del Proyecto: Nuevo Servicio de Comunicaciones

## Lo que hemos logrado hasta ahora:

1.  **Arquitectura Limpia (Clean Architecture):** Hemos separado el código en un diseño modular profesional.
    *   `core/`: Contiene la inteligencia del negocio.
        *   `engine.py`: El motor de replicación (Lee binlogs, traduce a SQL). **¡100% Agnóstico!** Tiene un único método `run()` genérico. No sabe si va de ida o de vuelta; todo su comportamiento se inyecta desde afuera.
        *   `buffer.py`: El gestor de transacciones ACID (agrupa eventos entre BEGIN y COMMIT garantizando que no queden datos a medias).
        *   `state.py`: La "pizarra" pública donde se anota el estado del servicio en tiempo real.
    *   `database/connection.py`: Encargado de las conexiones y el "Auto-Descubrimiento" inicial (quién soy, quién es la central, qué tablas ignoro).
    *   `api/routes.py`: Una API liviana en FastAPI que corre en paralelo solo para consultar el estado del servicio.
    *   `main.py`: El orquestador supremo. Coordina las conexiones, define las reglas asimétricas (Callbacks de ida y vuelta) y lanza los trabajadores en paralelo.

2.  **Lógica Avanzada de Replicación:**
    *   **Asincronismo Real:** El lector de MariaDB (`BinLogStreamReader`) corre en un hilo separado (`run_in_executor`) sin congelar la API ni el sistema.
    *   **Integridad Transaccional (ACID):** `buffer.py` agrupa eventos y hace un `ROLLBACK` completo en la base de destino si falla un solo registro del lote.
    *   **Traducción Dinámica:** Los eventos binarios se traducen a SQL dinámico en tiempo real, sin hardcodear nombres de columnas.
    *   **Filtros de Exclusión:** Se ignoran dinámicamente las tablas marcadas en `tablas_exceptuadas_sucursal` sin procesarlas ni un milisegundo de más.

## Siguientes Pasos (Para la próxima sesión):

*   **Gestión Dinámica de Dependencias (Foreign Keys):** Implementar la lógica en el bloque de excepciones (`IntegrityError` 1452) para que el sistema consulte dinámicamente `information_schema.KEY_COLUMN_USAGE` y loguee exactamente qué ID falta en qué tabla madre.
*   **Pruebas End-to-End:** Conectar a bases de datos con el motor configurado (`log-bin`, `sync_binlog=1`, etc.), generar transacciones de prueba y ver cómo fluyen los datos (quitando los comentarios de los `cursor.execute` en el motor).
