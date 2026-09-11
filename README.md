# Registro de Producción en Planta

Aplicación de escritorio en Python para que los operarios registren en tiempo real la producción de cada máquina de una planta, con un servicio que consolida todos los registros en una sola base sin riesgo de corromper el archivo.

## El problema

Cada máquina de la planta llevaba su producción en un archivo Excel propio: el operario digitaba a mano la orden de producción, la referencia, el cliente, el evento y los tiempos. Eso traía tres problemas:

- **Errores de digitación:** referencias mal escritas, códigos que no existían, horas inconsistentes.
- **Datos dispersos:** consolidar 16 archivos para analizar la producción era un trabajo manual y lento.
- **Archivos corruptos:** cuando varias máquinas escribían sobre un mismo Excel compartido, el archivo se dañaba y se perdían registros.

## La solución

Una interfaz de escritorio instalada en el PC de cada máquina, más un consolidador que corre en un PC dedicado.

**En la máquina (`app.py`):**
- El operario presiona **Iniciar** y un cronómetro mide el evento en tiempo real.
- Al digitar la **OP**, la app autocompleta tipo de trabajo, referencia y cliente desde la base de órdenes de producción.
- Eventos y operarios se eligen de listas desplegables con búsqueda, alimentadas desde las maestras: no hay texto libre donde hay un código válido.
- El formulario **cambia según el tipo de máquina**: las de impresión piden planchas, las de conversión piden material, bobina, calibre, ancho y corte.
- El turno (mañana, tarde, noche) y la fecha de turno se calculan solos a partir de la hora de inicio.

**En el PC dedicado (`consolidador.py`):**
- Cada 30 segundos revisa la carpeta de entrada, agrega los registros nuevos al Excel consolidado y archiva cada archivo procesado en una subcarpeta mensual.

## La decisión técnica central: un archivo por registro

El problema de fondo era la escritura concurrente: 16 PCs abriendo el mismo Excel terminaban corrompiéndolo.

La solución fue invertir el flujo. **Ninguna máquina escribe el Excel.** Cada registro se guarda como un archivo JSON independiente, con nombre único (`MÁQUINA_CÓDIGO_FECHA_HORA_aleatorio.json`), en una carpeta compartida. Un solo proceso —el consolidador— lee esos archivos y es el único que abre el Excel.

Las consecuencias de ese diseño:

- **Sin colisiones:** dos máquinas nunca escriben el mismo archivo, ni aunque guarden en el mismo segundo.
- **Escritura atómica:** cada archivo se escribe primero como `.tmp` y luego se renombra, así el consolidador jamás lee un archivo a medio escribir.
- **Tolerancia a fallos de red:** si la red se cae, el registro se guarda en el PC local y se sube solo en el siguiente guardado con red. El operario no pierde su trabajo.
- **Tolerancia al Excel ocupado:** si alguien tiene el archivo abierto, el consolidador reintenta y deja los pendientes para el siguiente ciclo. No se pierde nada.
- **El operario ve su registro al instante:** el historial de la app combina lo que ya está en el Excel con lo que sigue pendiente de consolidar.

## Otras decisiones

- **Configuración por PC:** cada equipo guarda en `config.json` qué máquina tiene asignada. Se configura una sola vez, con una ventana inicial.
- **Maestras siempre frescas:** eventos, operarios y órdenes se recargan cada 10 minutos en segundo plano, y también justo antes de guardar cada registro.
- **Errores en lenguaje del operario:** si no hay red, el mensaje explica qué revisar y a quién avisar, no muestra un error técnico.
- **Formularios declarativos:** los campos de cada grupo de máquina están definidos como listas de diccionarios, así que agregar un campo no implica tocar la lógica de la interfaz.

## Stack

Python 3.10+ · CustomTkinter · openpyxl · Pillow

## Estructura

```
registro-produccion-planta/
├── app.py                  # Interfaz del operario (CustomTkinter)
├── data_manager.py         # Maestras, órdenes, validación y guardado
├── consolidador.py         # Servicio que vuelca los JSON al Excel consolidado
├── config.py               # Rutas, parámetros y configuración por PC
├── iniciar.bat             # Abre la app en el PC de la máquina
├── consolidador.bat        # Arranca el consolidador en el PC dedicado
├── instalar.bat            # Instala dependencias
├── diagnostico.bat         # Abre la app mostrando errores en consola
├── requirements.txt
├── config.example.json
├── .env.example
└── .gitignore
```

## Cómo usarlo

1. Instalar dependencias: `pip install -r requirements.txt` (o doble clic en `instalar.bat`).
2. Copiar `.env.example` como `.env` y apuntar `CARPETA_DATOS` a la carpeta compartida.
3. En el PC de cada máquina: ejecutar `iniciar.bat` y seleccionar la máquina en la ventana de configuración inicial.
4. En el PC dedicado: ejecutar `consolidador.bat` y dejarlo corriendo.

## Datos que espera

- **`MAESTRAS.xlsx`:** hojas `MAESTRA MAQUINAS`, `MAESTRA AREAS`, `MAESTRA EVENTOS` y `MAESTRA OPERARIOS`.
- **Base de órdenes de producción:** una hoja con columnas de OP, tipo de trabajo, código de producto, referencia, cliente y documento de identificación.
- **Excel consolidado:** se crea solo con su encabezado la primera vez.

> Por confidencialidad, este repositorio no incluye datos reales de producción, clientes ni operarios. Los códigos de máquina y las rutas son de ejemplo.
