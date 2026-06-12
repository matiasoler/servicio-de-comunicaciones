# Decisiones y Pautas Arquitectónicas: Servicio de Comunicaciones V&S

Este documento detalla las decisiones de diseño, patrones y reglas estrictas para el desarrollo del nuevo sistema de replicación de datos (Agentes Python) para Zafiro.

## 1. Visión General y Alcance

*   **Objetivo:** Reemplazar el servicio monolítico en Delphi por una arquitectura descentralizada basada en microservicios Python.
*   **Topología:** Estrella. Cada sucursal tiene un **Agente** autónomo que replica datos hacia la **Base de Datos Central**.
*   **Rol del Panel:** El Panel Central (FastAPI) actúa exclusivamente como torre de control (telemetría, logs, envío de comandos). **NO** actúa como proxy de datos. La inserción de datos es directa Agente -> BD Central.
*   **Fuera de Alcance (Out of Scope):** Integraciones legacy (CloseUp, Zafiro BI, DataView, Recálculo de Costos) quedan excluidas de este nuevo desarrollo.

## 2. Decisiones de Diseño

### 2.1. Concurrencia y Modelo de Ejecución
*   **Decisión:** Uso estricto de `asyncio` (Corrutinas) en lugar de hilos del sistema operativo (Threads).
*   **Motivo:** Evitar la sobrecarga de *Context Switching* y bloqueos de I/O que sufría el sistema Delphi.
*   **Patrón:** Productor-Consumidor usando `asyncio.Queue`.
    *   **Productor:** `lector_binlog` (lee cambios locales).
    *   **Consumidor:** `transmisor` (envía datos a Central).
    *   **Monitor:** `escucha_comandos` (Heartbeat y recepción de órdenes).

### 2.2. Lectura de Cambios (Binlog)
*   **Decisión:** Utilizar la librería `python-mysql-replication` para leer el stream binario del motor MariaDB.
*   **Motivo:** El método anterior (parsear `SHOW BINLOG EVENTS` como texto) es frágil y propenso a errores. El stream binario ofrece eventos estructurados (diccionarios Python) y es nativo del protocolo de replicación.
*   **Encapsulamiento:** Dado que esta librería es bloqueante, debe ejecutarse dentro de un `ThreadPoolExecutor` (`loop.run_in_executor`) para no frenar el Event Loop de asyncio.

### 2.3. Atomicidad Transaccional
*   **Decisión:** El Agente debe procesar y transmitir datos en bloques transaccionales completos (`BEGIN` ... `COMMIT`), no registro por registro.
*   **Motivo:** Garantizar la integridad de documentos complejos (ej. un comprobante de venta con sus líneas y movimientos contables). Si falla una parte, no se debe transmitir nada (Rollback).

### 2.4. Prevención de Bucles de Replicación (Anti-Eco)
*   **Decisión:** Al aplicar cambios recibidos desde Central hacia la BD local, se debe ejecutar `SET SESSION sql_log_bin = 0;`.
*   **Motivo:** Evitar que un cambio originado en la Sucursal A, replicado a Central, vuelva a la Sucursal A y sea re-transmitido infinitamente (Loop infinito).

## 3. Pautas de Implementación

### 3.1. Manejo de Errores y Resiliencia
*   **Exponential Backoff:** El `transmisor` debe implementar reintentos con espera exponencial ante fallos de red o base de datos. No descartar datos ante fallos temporales.
*   **Backpressure (Contrapresión):** No implementar límites de velocidad artificiales (`sleep`). Confiar en el mecanismo nativo de TCP y `asyncio.Queue` para regular el flujo. Si el consumidor es lento, el productor se pausará automáticamente.

### 3.2. Límites de Memoria
*   **MaxSize en Queue:** Definir un `maxsize` explícito en las colas de asyncio (ej. 10,000 o 50,000 eventos) para evitar que el agente consuma toda la RAM de la máquina en caso de cortes prolongados de red.

### 3.3. Calidad de Código
*   **Tipado:** Uso obligatorio de Type Hints en todas las funciones.
*   **Logging:** Prohibido el uso de `print()`. Todo debe registrarse mediante el módulo `logging` con niveles adecuados (INFO, WARNING, ERROR).
*   **Configuración:** Credenciales y parámetros de conexión deben externalizarse (archivos `.env` o JSON de configuración), nunca hardcodeados.

## 4. Seguridad

### 4.1. Red y Comunicación
*   **VPN:** Obligatorio el uso de VPN (WireGuard o Tailscale) para toda la comunicación entre sucursales y central.
*   **Autenticación:** El Panel Central debe exigir autenticación (JWT o API Key) a los agentes, incluso dentro de la VPN (Zero Trust).

### 4.2. Base de Datos
*   **Principio de Menor Privilegio:** El usuario de BD que usa el Agente para leer el binlog local solo debe tener permisos de `REPLICATION SLAVE` y `REPLICATION CLIENT`. Nunca usar usuario `root`.

### 4.3. Panel de Control
*   **Prohibición de SQL Crudo:** El Panel **NUNCA** debe permitir la ejecución de sentencias SQL arbitrarias enviadas por el usuario.
*   **Comandos RPC:** La interacción debe ser mediante comandos predefinidos y validados (ej. `{"cmd": "sync_articulos"}`) que el agente ejecuta localmente con queries parametrizados.
