# Automatización de Notas desde Canvas LMS a Google Sheets

## Objetivo

El objetivo de este proyecto es desarrollar un sistema que extraiga las calificaciones desde Canvas LMS y las registre automáticamente en un documento de Google Sheets. Con esta solución se busca garantizar que cada nota se inserte en la celda correspondiente, evitando errores en la asignación de calificaciones.

## Fases Completadas

1. **Identificación de la estructura del archivo Excel**  
   - **Hoja relevante:** `EVALUACIÓN`.
   - **Alumnos:** Listados en la columna **C** a partir de la fila **10** (ordenados alfabéticamente).
   - **Actividades:** Se encuentran en la fila **7**, agrupadas mediante celdas combinadas que contienen el texto  
     **"RESULTADO APRENDIZAJE-CRITERIO DE EVALUACIÓN PRÁCTICOS"**.
   - **Notas de alumnos:** Ubicadas en la intersección entre la fila del alumno y la columna de la actividad.

2. **Creación de la aplicación para identificación de celdas**  
   Se desarrolló una interfaz gráfica con Tkinter que permite:
   - Seleccionar el archivo de Excel.
   - Introducir el nombre del alumno y de la actividad.
   - Determinar la celda exacta donde se debe insertar la nota.
   - Mostrar la referencia de la celda (por ejemplo, `AC10`) junto con su valor actual.

3. **Correcciones y mejoras**  
   - Se implementó el uso de `openpyxl` para manejar archivos **.xlsx**.
   - Se corrigió la conversión de índices de columnas utilizando `get_column_letter()`, permitiendo trabajar con columnas más allá de "Z".
   - Se incluyó el manejo de errores para casos en los que la celda no exista.

## Instalación y Requisitos

- **Python 3.6+**
- **Librerías necesarias:**
  - `pandas`
  - `openpyxl`
  - `tkinter` (incluido en la mayoría de las instalaciones de Python)
  - *(Para futuras integraciones: `google-auth` y `google-api-python-client` para conectar con Google Sheets)*

Para instalar las dependencias, ejecuta:

```bash
pip install -r requirements.txt
```
## Mapeo de Celdas

La aplicación utiliza el contenido textual de los encabezados para delimitar bloques de actividades. Los archivos Excel cuentan con celdas combinadas (generalmente en las filas **7** y **8**) que contienen el texto:

> "RESULTADO APRENDIZAJE-CRITERIO DE EVALUACIÓN PRÁCTICOS"

El número de estas celdas combinadas varía según el curso:

- **Primer curso:** 3 trimestres.
- **Segundo curso:** 2 trimestres.

Dentro de cada bloque, en la fila **9** se encuentran los nombres específicos de las actividades. La función de mapeo recorre cada rango combinado que contenga el encabezado indicado, extrae de la fila **9** el nombre de cada actividad y genera un diccionario en el que se asocia cada actividad con la(s) referencia(s) de celda correspondiente(s).

### Ejemplo del Diccionario de Mapeo

```python
{
  "1 h": ["AC9", "AE9"],
  "1K": ["AD9"],
  "Tarea 1": ["AF9", "AH9"],
  ...
}
```
## Verificación del Mapeo

Se implementó un modo de verificación en la interfaz gráfica para confirmar la correspondencia entre alumno, tarea y nota:

### Búsqueda de la fila del alumno
Se utiliza **pandas** para leer el Excel sin interpretar encabezados, de modo que los índices del DataFrame se correspondan directamente con las filas del Excel (por ejemplo, la fila `C10` se lee correctamente).

### Verificación de la tarea
Con el diccionario de mapeo obtenido, se reconstruye la referencia de celda en la que debe insertarse la nota para un alumno específico y se muestra el valor actual de esa celda. Esto permite comprobar que la lógica de mapeo es correcta y que, en el futuro, la inserción de datos provenientes de Canvas se realizará en la celda adecuada.

## Próximos Pasos

Las fases siguientes del proyecto incluyen:

### Integración con la API de Canvas LMS
Se desarrollará un módulo para extraer las calificaciones directamente desde Canvas, utilizando su API.

### Automatización en Google Sheets
Se implementará la conexión con la API de Google Sheets para escribir las notas en las celdas correspondientes, usando las credenciales configuradas en `credentials.json`.

### Orquestación del flujo completo
Se coordinará la extracción de datos, el mapeo de actividades y la actualización en Google Sheets para lograr una solución integral que automatice el proceso de calificación.

## Estructura del Proyecto
```plaintext
canvas_to_sheets/
├── config/
│   ├── settings.py           # Variables globales y configuración
│   ├── credentials.json      # Credenciales de Google Sheets
├── evaluator/
│   ├── __init__.py           # Permite que sea un módulo
│   ├── cell_locator.py       # Lógica para identificar la celda correcta
│   ├── test_locator.py       # Pruebas del localizador de celdas y mapeo
├── main.py                   # Punto de entrada de la aplicación
├── requirements.txt          # Dependencias necesarias
└── README.md                 # Documentación del proyecto (este archivo)
```
