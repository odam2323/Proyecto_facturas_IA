# Sistema de Auditoría de Facturas con IA (Gemini)

Este proyecto automatiza la extracción, validación y conciliación de información financiera contenida en facturas (PDFs y estructuras semi-estructuradas) utilizando Modelos de Lenguaje Grande (LLMs).

El sistema compara los datos extraídos de los documentos soporte contra un Excel maestro para generar un reporte de auditoría y conciliación.

## Descripción del Problema
El reto se abordó como un desafío de **Procesamiento de Lenguaje Natural (PLN)** y extracción de información. 

Inicialmente se consideró el uso de *LayoutLMv3*. Sin embargo, se optó por **Google Gemini (Multimodal)** debido a la complejidad de calibración que requerían los modelos tradicionales para un dataset pequeño. 

**Ventajas de la arquitectura elegida:**
- **Capacidad Zero-shot:** No requiere entrenamiento previo.
- **Flexibilidad:** Manejo nativo de texto e imágenes (OCR implícito).
- **Implementación Ágil:** Iteración rápida mediante ingeniería de prompts.
- **Mitigación de Alucinaciones:** Implementación de reglas de validación matemática estricta.

---

## Funcionamiento del Sistema

El flujo de ejecución del script principal (`main.py`) se divide en 7 etapas críticas:

### 1. Selección Dinámica del Modelo
El sistema consulta la API de Google para identificar los modelos disponibles con la API-KEY proveída.
* Filtra el catálogo para seleccionar aquellos con capacidades de visión (análisis de imágenes).
* Selecciona automáticamente la versión más avanzada disponible.
* Incluye manejo de errores en caso de fallo en la respuesta de la API.

### 2. Normalización y Formato
Se implementan funciones para estandarizar los datos financieros:
* **`parsear_valor_interno`**: Convierte números de diversos formatos a un estándar flotante para cálculos (elimina separadores visuales).
* **`formato_colombiano_estricto`**: Formatea la salida final con puntuación local (puntos para miles, comas para decimales).
* **Búsqueda inteligente de columnas**: Algoritmo capaz de encontrar columnas en el Excel maestro iterando por palabras clave y orden de importancia, mitigando cambios arbitrarios en los nombres de los encabezados.

### 3. Motor de Extracción (Gemini)
Conecta los PDFs con la IA bajo el rol de *"Experto Contable Internacional"*.
* **Prompt Engineering:** Se solicita una estructura JSON estricta con 16 campos definidos.
* **Limpieza de Entrada:** Normalización de fechas (DD/MM/AAAA) y limpieza de NITs.
* **Autocorrección Matemática:** Protocolo vital donde la IA verifica y corrige el cálculo `Subtotal + Impuestos = Total` antes de entregar el resultado.
* **Temperatura 0.0:** Para eliminar la "creatividad" del modelo y asegurar datos fácticos.

### 4. Procesamiento de Archivos PDF
La función `procesar_todo()` orquesta la lectura:
* Detecta si el PDF es texto seleccionable o imagen escaneada.
* Si es imagen, convierte el archivo a PNG y lo envía a Gemini (Vision).
* Realiza una validación post-extracción para asegurar la integridad matemática de los montos.
* Genera un respaldo intermedio en `.json` y un Excel visual.

### 5. Limpieza del Excel Maestro
Prepara la base de datos de referencia:
* Escaneo de las primeras 25 filas para detectar automáticamente la fila de encabezados (buscando keywords como NIT, INVOICE, TOTAL).
* **Regla de 9 dígitos:** Limpieza de NITs eliminando prefijos (CO, NIT) y el dígito de verificación (si existe), estandarizando a 9 dígitos para el cruce.
* Eliminación de columnas sin datos relevantes.

### 6. Análisis Descriptivo
Antes de la conciliación, se generan métricas sobre los datos:
* **Excel Maestro:** Identificación de proveedores únicos, rangos de fechas y porcentaje de celdas vacías.
* **Facturas (PDFs):** Informe estadístico sobre la frecuencia de campos encontrados y tipos de facturas procesadas.

### 7. Cruce y Conciliación (Reporte Final)
Se genera el `REPORTE_SISTEMA_AUDITORIA_IA.json`.
* **Lógica de Cruce:** Se compara no solo por NIT, sino también por Número de Factura.
* **Manejo de Duplicados:** Se detectó que el Excel maestro contenía múltiples registros para un mismo número de factura. El sistema agrupa estos registros y los compara en conjunto contra el PDF para asegurar la congruencia de los datos.
* **Salida:** Un diccionario unificado que reporta coincidencias, discrepancias y validaciones matemáticas.

---

##  Ejecución

El punto de entrada del programa es el archivo:
`main.py`

El resultado final se guardará automáticamente como:
`REPORTE_SISTEMA_AUDITORIA_IA.json`

##  Requisitos
* Python 3.x
* Librerías: Pandas, Google Generative AI, entre otras (ver `requirements.txt`).
* API Key de Google Gemini.
