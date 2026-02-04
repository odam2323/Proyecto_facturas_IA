
from datetime import datetime 
import os
import fitz  # PyMuPDF
import pandas as pd
import json
import time
import requests
import base64
import logging
import sys
import re 

# para hacer el codigo ejecutable con logging.

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)

# inicializamos información relevante para el ceodigo.

API_KEY = "Aquí va la API"
CARPETA_PDFS = "/Users/andres/Desktop/proyecto_facturas_IA/Datos"
NOMBRE_EXCEL_SALIDA = "Reporte_Facturacion_PDFs.xlsx"
NOMBRE_JSON_SALIDA = "Reporte_Ejecutivo.json"
RUTA_MAESTRO_ORIGINAL = "/Users/andres/Desktop/proyecto_facturas_IA/Datos/YCO01 - AP LISTING 20260104 con nit.xlsx"
RUTA_MAESTRO_LIMPIO = "/Users/andres/Desktop/proyecto_facturas_IA/Datos/Maestro_Limpio_Procesado.xlsx"
NOMBRE_JSON_ANALISIS = "JSON_EXCEL_MAESTRO.json"

# funciones para obtener el modelo disponible en la API de Google Gemini.

def obtener_modelo_disponible():
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={API_KEY}"
    try:
        response = requests.get(url)
        modelos = response.json()
        if 'models' in modelos:
            lista = [m['name'] for m in modelos['models'] if 'generateContent' in m['supportedGenerationMethods']]
            return lista
        return []
    except:
        return []

def procesar_modelo():
    logging.info("Consultando modelos permitidos...")
    modelos = obtener_modelo_disponible()
    if not modelos:
        logging.error("La API no devolvió modelos.")
        return None
    # Preferimos la versión 1.5-flash o 2.0 si existen
    modelo_elegido = next((m for m in modelos if "gemini" in m), modelos[0])
    logging.info(f"Modelo seleccionado: {modelo_elegido}")
    return modelo_elegido

MODELO_DISPONIBLE = procesar_modelo()

if MODELO_DISPONIBLE:
    URL_API = f"https://generativelanguage.googleapis.com/v1beta/{MODELO_DISPONIBLE}:generateContent?key={API_KEY}"
else:
    URL_API = None
    logging.critical("Error crítico: No hay URL de API disponible.")


def parsear_valor_interno(val):
    """
    Convierte valores a float. Maneja el error de 'Punto como miles' (272.000 -> 272000).
    """
    if pd.isna(val) or val == "" or str(val).strip().upper() in ["N/A", "NAN"]:
        return 0.0
    
    if isinstance(val, (int, float)):
        return float(val)
        
    # Limpieza inicial
    val_str = str(val).replace("$", "").replace(" ", "").replace("COP", "").replace("USD", "").replace("'", "").strip()
    
    try:
        # CASO CRÍTICO: Detectar formato "272.000" (Punto miles, sin decimales o decimal ,00)
        # Si hay puntos pero NO comas, y el formato parece miles (ej: 12.000 o 272.000)
        if "." in val_str and "," not in val_str:
            # Si tiene más de un punto (1.500.000) es fijo miles -> Borrar puntos
            if val_str.count(".") > 1:
                val_str = val_str.replace(".", "")
            # Si tiene un solo punto, verificamos los decimales
            else:
                partes = val_str.split('.')
                # Si la parte "decimal" tiene exactamente 3 dígitos (ej: .000, .500)
                # asumimos que es miles, porque en facturación nadie cobra 0.000 de fracción
                if len(partes[1]) == 3:
                    val_str = val_str.replace(".", "")
                # Si son 2 digitos (.50) lo dejamos como decimal
        
        # Lógica estándar (Latam vs USA)
        elif "," in val_str and "." in val_str:
            if val_str.rfind(",") > val_str.rfind("."): 
                val_str = val_str.replace(".", "").replace(",", ".") # 1.500,00 -> 1500.00
            else:
                val_str = val_str.replace(",", "") # 1,500.00 -> 1500.00
        
        elif "," in val_str:
            # Caso 100,50 -> 100.50
            if len(val_str) - val_str.rfind(",") <= 3:
                val_str = val_str.replace(",", ".")
            else:
                val_str = val_str.replace(",", "") # 1,000 -> 1000

        return float(val_str)
    except:
        return 0.0

def formato_colombiano_estricto(valor_float):
    """
    Convierte un número Python a TEXTO visual para Excel.
    Formato: 1'500.000,00 (Miles con punto, Millones con comilla, Decimales con coma).
    """
    if not isinstance(valor_float, (int, float)):
        return "0,00"
    
    # 1. Separar entero y decimales (siempre 2 decimales)
    texto_base = "{:.2f}".format(valor_float)
    entero, decimal = texto_base.split('.')
    
    # 2. Formatear la parte entera
    invertido = entero[::-1]
    nuevo_entero = ""
    
    for i in range(0, len(invertido), 3):
        grupo = invertido[i:i+3]
        if i > 0:
            # Cada 6 dígitos (millones) usamos comilla ('), si no punto (.)
            if i % 6 == 0:
                separador = "'"
            else:
                separador = "."
            nuevo_entero += separador
        nuevo_entero += grupo
        
    entero_final = nuevo_entero[::-1]
    
    return f"{entero_final},{decimal}"

def encontrar_columna(lista_prioridad, columnas_df):
    """Busca una columna en el DataFrame basándose en una lista de palabras clave por prioridad."""
    columnas_upper = [c.upper() for c in columnas_df]
    for keyword in lista_prioridad:
        matches = [c for c, c_up in zip(columnas_df, columnas_upper) if keyword in c_up]
        if matches:
            return matches[0]
    return None

# extracción con Gemini IMPORTANTE. 

def extraer_con_gemini(contenido, es_imagen=False):
    prompt_texto = """
    Eres un experto contable internacional. Analiza esta factura y extrae la información en JSON.
    Si la imagen está rotada o borrosa, haz tu mejor esfuerzo por leerla.

    Campos requeridos en el JSON:
    1. nombre_proveedor (En mayúsculas), 
    2. nit_proveedor (Solo números/letras, sin guiones, sin puntos, sin difgito de verificación, y sin prefijos como 'CO', suele estar cerca de la palabra nit o tax id)
    3. digito_verificacion (Si no existe en facturas internacionales, poner "N/A")
    4. numero_factura (en algunos casos este está bajo el nombre de Documento No, esto no puede comenzar por DEE)
    5. fecha_emision (Formato estricto DD/MM/YYYY. si solo dice fecha esa es la fecha_emision)
    6. fecha_vencimiento (Formato estricto DD/MM/YYYY. si solo dice fecha esa es la fecha_emision, pon N/A entonces en fecha_vencimiento, no asumas ni calcules)
    7. moneda (ISO de 3 letras: COP, USD, EUR, etc.)
    8. subtotal (Valor numérico, en algunos casos hay que calcularlo con la suma de cargo fijo, cargo variable y cargo variable aprovecha, en otras aparece explícito.
    es el valor antes de impuestos (iva u otros) y descuentos. Es necesario que este sea puesto en formato numérico colombiano: punto como separador de miles y coma como separador decimal.
    También suele ser llamado valor bruto. Suele ser un numero menor al total, es el valor antes de impuestos.ignorar símbolos de moneda ($) para evitar confusiones de formato.)
    9. iva_porcentage (Solo el número, ej: 19)
    10. iva_monto (Valor numérico, siempre cerca de la palabra iva o IVA)
    11. otros_impuestos (Suma de otros cargos o impuestos. En algunos casos, tienen nombres especificos como contribucion, contrib, contr o tasa, pueden tener un '%' en su nombre. 
    Si encuentras retefuente, ReteIVA 15% y ReteICA 9.66 x mil Bogotea son subsidios, asi que ponlos en negativo en la tabla de otros_impuestos, pero solo en estos casos)
    12. total (Valor neto a pagar del periodo o Total Factura, en algunos casos esta cerca de la palabra total a pagar o total) 
    13. numero_lineas (Cantidad de ítems/servicios listados, en algunos casos los items son: acueducto, alcantarillado, energia y otras endidades)
    14. orden_compra (Purchase Order / PO number, sino N/A)
    15. cufe (Sino existe, N/A. Si hay Cude (como en el caso de las empresas publicas de medellin), no existe cufe, poner N/A)
    16. tipo_factura (electrónica, equivalente o pos, si no está explícito, poner N/A)

    
    Instrucciones de mapeo para facturas en inglés:
    - nit_proveedor: Extraer de 'Tax ID', 'VAT Number', 'EIN'.
    - fecha_emision/vencimiento: Convertir siempre a DD/MM/YYYY.
    - cufe: Si no existe, "N/A".

    Reglas estrictas:
    - nit_proveedor: el nit termina justo antes del -, no incluyas ningun número despues del guion, es decir, no incluyas el digito de verificación.
    - nit_proveedor: si el nit no está escrito explicitamente, pon "N/A".
    - numero_factura: ningún número_factura comienza con un PO
    - fecha_emision: Si la fecha de vencimiento NO está escrita explícitamente en el documento, pon "N/A". NO calcules ni asumas una fecha basada en la emisión.
    - fecha_vencimiento: Si la fecha de vencimiento NO está escrita explícitamente en el documento, pon "N/A". 
    NO calcules ni asumas una fecha basada en la emisión, no pongas una fecha de vencimiento un dia despues de la fecha_emision.
    - Exactitud Numérica: No redondees valores. Extrae los decimales exactos.
    - Integridad de Datos: Si un campo como 'orden_compra' o 'cufe' no es visible, usa "N/A" obligatoriamente.
    - Responde EXCLUSIVAMENTE con el objeto JSON.
    - Si un dato no es localizable, usa "N/A".
    
    
    *** PROTOCOLO DE AUTOCORRECCIÓN MATEMÁTICA ***
    Antes de generar el JSON final, realiza la siguiente verificación interna paso a paso:
    1. Suma: subtotal + iva_monto + otros_impuestos.
    2. Compara el resultado con el 'total' extraído del documento.
    3. SI NO COINCIDEN (Diferencia mayor a 1.0):
       - Prioridad 1 (La Verdad): El valor explícito impreso como "TOTAL A PAGAR" en la factura es la verdad absoluta. No lo cambies.
       - Acción: Recalcula el 'subtotal' o busca si omitiste algún cargo en 'otros_impuestos' (como Reteica, bolsas, propina, tasas).
       - Ajuste: Si el IVA es incorrecto respecto al subtotal, ajústalo para que cuadre con el Total.
       - Objetivo: El JSON final DEBE cumplir la ecuación: subtotal + iva + otros = total.
    """

    parts = [{"text": prompt_texto}]

    if es_imagen:
        for img in contenido:
            parts.append({"inline_data": {"mime_type": "image/png", "data": img}})
    else:
        parts.append({"text": f"DOCUMENTO:\n{contenido}"})

    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {
            "temperature": 0.0,
            "response_mime_type": "application/json"
        }
    }
    
    try:
        response = requests.post(URL_API, headers={'Content-Type': 'application/json'}, json=payload, timeout=45)
        res_json = response.json()
        
        if 'candidates' in res_json:
            raw = res_json['candidates'][0]['content']['parts'][0]['text'].strip()
            # Limpieza de bloques de código
            raw = raw.replace("```json", "").replace("```", "").strip()
            return json.loads(raw)
        else:
            logging.error(f"Error API: {res_json.get('error')}")
            return None
    except Exception as e:
        logging.error(f"Excepción: {e}")
        return None

# procesamiento principal de los PDFs

def procesar_todo():
    if not os.path.exists(CARPETA_PDFS):
        logging.error(f"No existe la carpeta: {CARPETA_PDFS}")
        return

    archivos = [f for f in os.listdir(CARPETA_PDFS) if f.lower().endswith('.pdf')]
    logging.info(f" Iniciando procesamiento de {len(archivos)} facturas...")
    
    resultados_finales = []

    for i, nombre in enumerate(archivos, 1):
        logging.info(f"[{i}/{len(archivos)}] Procesando: {nombre}")
        ruta = os.path.join(CARPETA_PDFS, nombre)
        
        datos = None
        
        try:
            doc = fitz.open(ruta)
            texto = "".join([p.get_text() for p in doc])
            
            # Decisión: Texto vs OCR (Imagen)
            if texto.strip() and len(texto) > 50:
                datos = extraer_con_gemini(texto[:15000], False)
            else:
                logging.info("Modo OCR (Imagen)")
                imgs = []
                for idx, p in enumerate(doc):
                    if idx >= 2: break 
                    pix = p.get_pixmap(matrix=fitz.Matrix(2,2))
                    imgs.append(base64.b64encode(pix.tobytes("png")).decode('utf-8'))
                datos = extraer_con_gemini(imgs, True)
            doc.close()
        except Exception as e:
            logging.error(f"Error leyendo PDF: {e}")

        # Construcción de la fila
        fila = {}
        if datos:
            # 1. Convertir a float real para poder sumar (Lógica interna)
            f_subtotal = parsear_valor_interno(datos.get('subtotal'))
            f_iva = parsear_valor_interno(datos.get('iva_monto'))
            f_otros = parsear_valor_interno(datos.get('otros_impuestos'))
            f_total = parsear_valor_interno(datos.get('total'))
            
            # 2. VERIFICACIÓN MATEMÁTICA
            # Sumamos y comparamos con margen de error de 1.0
            suma_calculada = f_subtotal + f_iva + f_otros
            es_correcto = abs(suma_calculada - f_total) <= 1.0
            
            # 3. Formatear visualmente para Excel (Estilo Colombiano)
            fila = {
                'archivo_origen': nombre,
                'estado': 'Exitoso',
                'nombre_proveedor': datos.get('nombre_proveedor', 'N/A'),
                'nit_proveedor': datos.get('nit_proveedor', 'N/A'),
                'digito_verificación': datos.get('digito_verificacion', 'N/A'),
                'numero_factura': datos.get('numero_factura', 'N/A'),
                'fecha_emision': datos.get('fecha_emision', 'N/A'),
                'fecha_vencimiento': datos.get('fecha_vencimiento', 'N/A'),
                'moneda': datos.get('moneda', 'N/A'),
                
                # Valores numéricos con formato visual
                'subtotal': formato_colombiano_estricto(f_subtotal),
                'iva_porcentage': datos.get('iva_porcentage', 'N/A'),
                'iva_monto': formato_colombiano_estricto(f_iva),
                'otros_impuestos': formato_colombiano_estricto(f_otros),
                'total': formato_colombiano_estricto(f_total),
                
                'numero_lineas': datos.get('numero_lineas', 'N/A'),
                'orden_compra': datos.get('orden_compra', 'N/A'),
                'cufe': datos.get('cufe', 'N/A'),
                'tipo_factura': datos.get('tipo_factura', 'N/A'),
                
                # LA COLUMNA DE VERIFICACIÓN SOLICITADA
                'verificación': es_correcto 
            }
        else:
            fila = {'archivo_origen': nombre, 'estado': 'Fallido'}
            for k in ['nombre_proveedor','subtotal','total','verificación']: fila[k] = "N/A"

        resultados_finales.append(fila)
        time.sleep(1.5)
        
        #guardar resultados
    
    # Lista ordenada de columnas incluyendo 'verificación' al final
    columnas_ordenadas = [
        'archivo_origen', 'estado', 'nombre_proveedor', 'nit_proveedor', 
        'digito_verificación', 'numero_factura', 'fecha_emision', 
        'fecha_vencimiento', 'moneda', 'subtotal', 'iva_porcentage', 
        'iva_monto', 'otros_impuestos', 'total', 'numero_lineas', 
        'orden_compra', 'cufe', 'tipo_factura', 
        'verificación'  # <--- Aquí está la columna final
    ]

    if resultados_finales:
        df = pd.DataFrame(resultados_finales)
        
        # Filtramos para asegurar que el orden sea exacto
        cols_final = [c for c in columnas_ordenadas if c in df.columns]
        df = df[cols_final]
        
        try:
            # Guardar Excel
            df.to_excel(NOMBRE_EXCEL_SALIDA, index=False)
            logging.info(f"Excel generado: {NOMBRE_EXCEL_SALIDA}")
            
            # Guardar JSON
            with open(NOMBRE_JSON_SALIDA, 'w', encoding='utf-8') as f:
                json.dump(resultados_finales, f, indent=4, ensure_ascii=False, default=str)
            logging.info(f"JSON generado: {NOMBRE_JSON_SALIDA}")
            
        except Exception as e:
            logging.error(f"Error guardando archivos: {e}")
    else:
        logging.warning("No se procesaron datos.")
        

# limpieza del excel maestro 

def generar_maestro_limpio():
    ruta_entrada = "/Users/andres/Desktop/proyecto_facturas_IA/Datos/YCO01 - AP LISTING 20260104 con nit.xlsx"
    ruta_salida = "/Users/andres/Desktop/proyecto_facturas_IA/Datos/Maestro_Limpio_Procesado.xlsx"

    logging.info(f"🧹 Iniciando limpieza y depuración del archivo maestro...")

    if not os.path.exists(ruta_entrada):
        logging.error(" No se encuentra el archivo maestro en la ruta especificada.")
        return

    try:
        # DETECCIÓN INTELIGENTE DE CABECERAS
        df_preview = pd.read_excel(ruta_entrada, header=None, nrows=25)
        fila_cabecera = 0
        keywords_clave = ["NIT", "TAX ID", "INVOICE", "DOCUMENT NO", "DATE", "AMOUNT", "TOTAL", "SUPPLIER"]
        
        max_score = 0
        for idx, row in df_preview.iterrows():
            txt_fila = " ".join([str(x).upper() for x in row.values])
            score = sum(1 for k in keywords_clave if k in txt_fila)
            if score > max_score:
                max_score = score
                fila_cabecera = idx
        
        logging.info(f"Cabeceras detectadas en la fila {fila_cabecera + 1}")
        
        # Cargar el Excel con la cabecera correcta
        df = pd.read_excel(ruta_entrada, header=fila_cabecera)
        
        # Normalizar nombres de columnas (Trim y Upper temporal para búsquedas)
        df.columns = df.columns.astype(str).str.strip()
        
        # eliminación de las columnas innecesarias, esto se hace porque el original tiene muchos huecos en cero innecesarios. 
        
        columnas_a_borrar = [
            "TOTAL A PAGAR", 
            "VAT/WHT3", "VAT/WHT4", "VAT/WHT5", 
            "TAX APPLICABILITY CODE", "TAX RATE", 
            "VAT1", "VAT2","VAT3", "VAT4", "VAT5", "VAT TOTAL", 
            "CODIGO WHT", "WHT1", "WHT2", "WHT 3", "WHT4", "WHT5", "WHT TOTAL"
        ]
        
        cols_eliminadas = []
        for col in df.columns:
            # Comparamos en mayúsculas para asegurar coincidencia
            if col.upper() in [x.upper() for x in columnas_a_borrar]:
                df.drop(columns=[col], inplace=True)
                cols_eliminadas.append(col)
        
        if cols_eliminadas:
            logging.info(f"   Columnas eliminadas ({len(cols_eliminadas)}): {', '.join(cols_eliminadas[:5])}...")


        def limpiar_nit_regla_9_digitos(valor):
            """
            Limpia el NIT y aplica la regla: Si longitud es 10, corta el último dígito.
            """
            if pd.isna(valor) or valor == "": return "N/A"
            s = str(valor).upper().strip()
            
            # Quitar basura común
            s = s.replace("CO", "").replace("NIT", "")
            
            # Si tiene guion, cortar antes del guion
            if "-" in s: s = s.split("-")[0]
            
            # Dejar SOLO números
            limpio = re.sub(r'[^0-9]', '', s)
            
            if not limpio: return "N/A"

            # REGLA DE NEGOCIO: SI TIENE 10 DÍGITOS, ELIMINAR EL ÚLTIMO 
            if len(limpio) == 10:
                limpio = limpio[:-1] # Se asume que el décimo era el DV pegado
            
            return limpio

        def limpiar_fecha(valor):
            """Formato DD/MM/YYYY"""
            if pd.isna(valor) or valor == "": return "N/A"
            try:
                return pd.to_datetime(valor).strftime("%d/%m/%Y")
            except:
                return str(valor)[:10]

        def limpiar_dinero(valor):
            """Formato Colombiano: 1.500.000,00"""
            if pd.isna(valor) or str(valor).strip() == "": return "0,00"
            try:
                num = 0.0
                if isinstance(valor, (int, float)):
                    num = float(valor)
                else:
                    s = str(valor).replace("$", "").replace("USD", "").replace("COP", "").strip().replace(" ", "")
                    # Lógica inteligente de puntos y comas
                    if "," in s and "." in s:
                        if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
                        else: s = s.replace(",", "")
                    elif "," in s:
                        if len(s) - s.rfind(",") == 3: s = s.replace(",", ".")
                        else: s = s.replace(",", "")
                    num = float(s)
                return formato_colombiano_estricto(num)
            except:
                return "0,00"
            
        # aplicación de la limpieza. 
        
        columnas_transformadas = []
        
        for col in df.columns:
            name_up = col.upper()
            
            # A. NIT (Evitamos confundir con montos de impuestos)
            if any(k in name_up for k in ["NIT", "TAX ID", "VAT REG", "TAX NO"]) and "AMOUNT" not in name_up:
                df[col] = df[col].apply(limpiar_nit_regla_9_digitos)
                columnas_transformadas.append(f"{col} (NIT)")

            # B. FECHAS
            elif any(k in name_up for k in ["DATE", "FECHA", "POSTING", "DUE"]):
                df[col] = df[col].apply(limpiar_fecha)

            # C. DINERO (Solo si no es una ID y parece dinero)
            elif (
                any(k in name_up for k in ["AMOUNT", "TOTAL", "SUBTOTAL", "BALANCE"]) or
                (any(k in name_up for k in ["VAT", "WHT"]) and not any(ex in name_up for ex in ["NO", "ID", "REG", "CODE"]))
            ):
                df[col] = df[col].apply(limpiar_dinero)

        # GUARDAR RESULTADO FINAL
        df.to_excel(ruta_salida, index=False)
        logging.info(f"Excel Maestro Limpio generado: {ruta_salida}")

    except Exception as e:
        logging.error(f"Error procesando el maestro: {e}")


# Analisis descriptivos

def generar_analisis_descriptivo_json():
    """Genera un JSON con estadísticas y metadatos del Excel Maestro Limpio"""
    logging.info(f"Generando análisis descriptivo del Maestro...")
    
    if not os.path.exists(RUTA_MAESTRO_LIMPIO):
        logging.error("No existe el Maestro Limpio para analizar.")
        return

    try:
        df = pd.read_excel(RUTA_MAESTRO_LIMPIO)
        columnas = df.columns.tolist()
        print(f"\nCOLUMNAS DETECTADAS: {columnas}") 
        
        total_registros = len(df)
        
        # --- FUNCIÓN HELPER DE BÚSQUEDA ---
        def encontrar_columna(lista_prioridad, columnas_df):
            columnas_upper = [c.upper() for c in columnas_df]
            for keyword in lista_prioridad:
                # Busca si la palabra clave está dentro del nombre de la columna
                matches = [c for c, c_up in zip(columnas_df, columnas_upper) if keyword in c_up]
                if matches: return matches[0]
            return None

        # --- FUNCIÓN HELPER DE FECHA (FORMATO ESTRICTO DD/MM/YYYY) ---
        def analizar_rango_fecha(df, col_nombre):
            if not col_nombre: return "No encontrada"
            try:
                # LEEMOS ESTRICTAMENTE DD/MM/YYYY (Lo que generamos en la limpieza)
                fechas = pd.to_datetime(df[col_nombre], format="%d/%m/%Y", errors='coerce')
                validas = fechas.dropna()
                
                if validas.empty: return "Sin fechas válidas"
                
                min_f = validas.min().strftime("%d/%m/%Y")
                max_f = validas.max().strftime("%d/%m/%Y")
                return f"{min_f} al {max_f}"
            except Exception as e:
                return f"Error leyendo fechas: {str(e)}"

        # 1. ANÁLISIS DE MONEDAS (Restaurado)
        # Prioridad: "CURR" (como pediste), luego "CURRENCY", luego español
        col_moneda = encontrar_columna(["CURR.", "CURRENCY", "DIVISA", "MONEDA"], columnas)
        
        monedas_detectadas = []
        if col_moneda:
            # Obtenemos únicos, convertimos a string y quitamos vacíos
            monedas_detectadas = df[col_moneda].dropna().astype(str).unique().tolist()
        else:
            monedas_detectadas = ["No se encontró columna (CURR)"]
            logging.warning("⚠️ No se detectó columna de Moneda.")

        # 2. ANÁLISIS DE FECHAS (Separado Emisión vs Vencimiento)
        # A. Invoice Date
        col_invoice = encontrar_columna(["INVOICE DATE", "DOC DATE", "DOCUMENT DATE"], columnas)
        rango_emision = analizar_rango_fecha(df, col_invoice)
        
        # B. Due Date
        col_due = encontrar_columna(["DUE DATE", "VENCIMIENTO", "NET DATE"], columnas)
        rango_vencimiento = analizar_rango_fecha(df, col_due)

        # 3. ANÁLISIS DE PROVEEDORES
        col_prov = encontrar_columna(["SUPPLIER/BENEFICIARY", "VENDOR NAME", "TERCERO", "NAME"], columnas)
        
        count_provs = 0
        lista_provs = []
        if col_prov:
            raw = df[col_prov].astype(str).str.strip().str.upper()
            # Filtro de basura
            clean = raw[~raw.isin(["NAN", "N/A", "NONE", "", "0", "NULL", "TOTAL"])]
            lista_provs = sorted(clean.unique().tolist())
            count_provs = len(lista_provs)

        # 4. DATOS FALTANTES
        celdas_vacias = 0
        vacios_por_columna = {}
        for col in df.columns:
            cnt = df[col].apply(lambda x: 1 if pd.isna(x) or str(x).strip() in ["", "N/A", "nan"] else 0).sum()
            vacios_por_columna[col] = int(cnt)
            celdas_vacias += cnt
            
        porcentaje = round(100 - (celdas_vacias / df.size * 100), 2) if df.size > 0 else 0

        # CONSTRUCCIÓN DEL JSON
        analisis_json = {
            "titulo": "DESCRIPCIÓN Y ANÁLISIS DEL ARCHIVO MAESTRO EXCEL",
            "metadata": {
                "fecha_analisis": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "archivo_origen": RUTA_MAESTRO_ORIGINAL
            },
            "contenido_clave": {
                "monedas": {
                    "columna_identificada": str(col_moneda),
                    "valores_encontrados": monedas_detectadas
                },
                "fechas": {
                    "rango_emision_facturas": rango_emision,
                    "columna_emision_usada": str(col_invoice),
                    "rango_vencimientos_pago": rango_vencimiento,
                    "columna_vencimiento_usada": str(col_due)
                },
                "proveedores": {
                    "total_unicos": count_provs,
                    "columna_usada": str(col_prov),
                    "lista_completa": lista_provs
                }
            },
            "calidad_datos": {
                "integridad_global": f"{porcentaje}%",
                "vacios_por_columna": vacios_por_columna
            }
        }

        with open(NOMBRE_JSON_ANALISIS, 'w', encoding='utf-8') as f:
            json.dump(analisis_json, f, indent=4, ensure_ascii=False)
            
        logging.info(f"JSON generado: {NOMBRE_JSON_ANALISIS}")
        
        # Resumen en Consola
        print("\n--- RESUMEN RÁPIDO DEL MAESTRO ---")
        print(f"Registros: {total_registros}")
        print(f"Monedas ({col_moneda}): {monedas_detectadas}")
        print(f"Emisión ({col_invoice}): {rango_emision}")
        print(f"Vencimiento ({col_due}): {rango_vencimiento}")
        print(f"Empresas Únicas: {count_provs}")
        print("----------------------------------\n")

    except Exception as e:
        logging.error(f"Error generando análisis JSON: {e}", exc_info=True)
        
# analisis descriptivo de los PDFs.

def generar_analisis_calidad_pdf():
    """Analiza el JSON de resultados de la IA para describir tipos de facturas y completitud"""
    logging.info(f"Generando análisis de calidad de los PDFs procesados...")
    
    archivo_json_ia = NOMBRE_JSON_SALIDA # "Reporte_Ejecutivo.json"
    archivo_salida_analisis = "Analisis_Calidad_PDFs.json"

    if not os.path.exists(archivo_json_ia):
        logging.error("No existe el reporte de PDFs (Reporte_Ejecutivo.json) para analizar.")
        return

    try:
        # Cargar datos
        with open(archivo_json_ia, 'r', encoding='utf-8') as f:
            datos = json.load(f)
        
        df = pd.DataFrame(datos)
        
        # Filtrar solo exitosos para el análisis de campos
        df_ok = df[df['estado'] == 'Exitoso'].copy()
        
        if df_ok.empty:
            logging.warning(" No hay extracciones exitosas para analizar.")
            return

        # Normalizar el Tipo de Factura (Mayúsculas y quitar espacios)
        df_ok['tipo_factura'] = df_ok['tipo_factura'].astype(str).str.upper().str.strip()
        
        # Tipos encontrados
        tipos_encontrados = df_ok['tipo_factura'].value_counts()
        
        # Campos clave a analizar
        campos_clave = [
            'cufe', 'nit_proveedor', 'orden_compra', 
            'fecha_vencimiento', 'resolucion_dian', 'digito_verificación'
        ]

        analisis_por_tipo = {}

        for tipo in tipos_encontrados.index:
            sub_df = df_ok[df_ok['tipo_factura'] == tipo]
            total_tipo = len(sub_df)
            
            detalles_campos = {}
            
            for campo in campos_clave:
                # Si el campo existe en el DF, contamos cuántos NO son "N/A" ni vacíos
                if campo in sub_df.columns:
                    # Criterio de "Dato Presente": No es N/A, ni None, ni vacío
                    con_dato = sub_df[campo].apply(lambda x: 1 if str(x).strip().upper() not in ["N/A", "NONE", "NAN", ""] else 0).sum()
                    porcentaje = (con_dato / total_tipo) * 100
                    
                    # Generamos una descripción cualitativa
                    if porcentaje == 100: desc = "Siempre presente (100%)"
                    elif porcentaje > 80: desc = "Muy frecuente (>80%)"
                    elif porcentaje > 50: desc = "Común (>50%)"
                    elif porcentaje > 0:  desc = "Ocasional (<50%)"
                    else: desc = "No detectado (0%)"
                    
                    detalles_campos[campo] = {
                        "cantidad": int(con_dato),
                        "porcentaje": round(porcentaje, 1),
                        "descripcion": desc
                    }
            
            # Construimos el perfil del tipo de factura
            analisis_por_tipo[tipo] = {
                "cantidad_documentos": int(total_tipo),
                "representacion_del_total": f"{round((total_tipo / len(df_ok))*100, 1)}%",
                "analisis_campos": detalles_campos,
                "conclusion_ia": f"Este tipo de documento ({tipo}) suele tener {detalles_campos.get('cufe', {}).get('descripcion', 'N/A')} el CUFE y {detalles_campos.get('nit_proveedor', {}).get('descripcion', 'N/A')} el NIT."
            }

        # Construcción del JSON Final
        json_final = {
            "titulo": "ANÁLISIS DE COMPLETITUD Y TIPOLOGÍA DE FACTURAS (FUENTE: PDFs)",
            "metadata": {
                "fecha_analisis": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "total_procesados": len(df),
                "total_exitosos": len(df_ok)
            },
            "distribucion_global": {
                k: int(v) for k, v in tipos_encontrados.items()
            },
            "detalle_por_tipo_documento": analisis_por_tipo
        }

        # Guardar
        with open(archivo_salida_analisis, 'w', encoding='utf-8') as f:
            json.dump(json_final, f, indent=4, ensure_ascii=False)

        logging.info(f"Análisis de Calidad PDF generado: {archivo_salida_analisis}")
        
        # Resumen en Consola
        print("\n--- RESUMEN DE TIPOLOGÍA ENCONTRADA ---")
        for tipo, count in tipos_encontrados.items():
            print(f"{tipo}: {count} documentos")
            # Mostrar un dato clave rápido
            pct_cufe = analisis_por_tipo[tipo]['analisis_campos'].get('cufe', {}).get('porcentaje', 0)
            print(f" Presencia de CUFE: {pct_cufe}%")
        print("---------------------------------------\n")

    except Exception as e:
        logging.error(f"Error en análisis de calidad PDF: {e}", exc_info=True)
        
# creación del JSON definitivo
        
def generar_tabla_unica_proveedores():
    """Cruza los proveedores del Maestro con los de los PDFs y genera un dataset de comparación"""
    logging.info(f" Generando tabla única de proveedores (Cruce Maestro vs PDFs)...")
    
    archivo_maestro = RUTA_MAESTRO_LIMPIO
    archivo_json_ia = NOMBRE_JSON_SALIDA
    NOMBRE_TABLA_PROV = "Tabla_Unica_Proveedores.json"

    if not os.path.exists(archivo_maestro) or not os.path.exists(archivo_json_ia):
        logging.error(" Faltan archivos para generar la tabla única.")
        return

    try:
        # 1. Cargar y preparar datos del Maestro
        df_m = pd.read_excel(archivo_maestro)
        col_nit_m = encontrar_columna(["NIT", "TAX ID"], df_m.columns)
        col_nom_m = encontrar_columna(["SUPPLIER", "BENEFICIARY", "NAME"], df_m.columns)
        
        # Limpieza profunda de NITs del maestro para el cruce
        def limpiar_nit_cruce(val):
            s = str(val).split('.')[0] # Quitar decimales si es float
            return re.sub(r'[^0-9A-Z]', '', s.upper().strip())

        df_m['NIT_KEY'] = df_m[col_nit_m].apply(limpiar_nit_cruce)

        maestro_unicos = df_m.groupby('NIT_KEY').agg({
            col_nom_m: 'first'
        }).reset_index().rename(columns={'NIT_KEY': 'nit', col_nom_m: 'nombre_maestro'})

        # 2. Cargar y preparar datos de los PDFs (IA)
        with open(archivo_json_ia, 'r', encoding='utf-8') as f:
            datos_ia = json.load(f)
        
        df_ia = pd.DataFrame(datos_ia)
        df_ia_exitosos = df_ia[df_ia['estado'] == 'Exitoso'].copy()
        
        # Limpieza de NITs de la IA
        df_ia_exitosos['NIT_KEY'] = df_ia_exitosos['nit_proveedor'].apply(limpiar_nit_cruce)
        
        ia_unicos = df_ia_exitosos.groupby('NIT_KEY').agg({
            'nombre_proveedor': 'first'
        }).reset_index().rename(columns={'NIT_KEY': 'nit', 'nombre_proveedor': 'nombre_pdf'})

        # 3. REALIZAR EL CRUCE (JOIN)
        cruce = pd.merge(maestro_unicos, ia_unicos, on='nit', how='outer', indicator=True)

        tabla_proveedores = []
        for _, row in cruce.iterrows():
            orig = row['_merge']
            clasificacion = ""
            if orig == 'both': clasificacion = "En ambos"
            elif orig == 'left_only': clasificacion = "Solo en Excel Maestro"
            else: clasificacion = "Solo en PDFs"

            tabla_proveedores.append({
                "nit": str(row['nit']),
                "nombre_en_maestro": str(row['nombre_maestro']) if pd.notnull(row['nombre_maestro']) else "N/A",
                "nombre_en_pdf": str(row['nombre_pdf']) if pd.notnull(row['nombre_pdf']) else "N/A",
                "presencia": clasificacion
            })

        with open(NOMBRE_TABLA_PROV, 'w', encoding='utf-8') as f:
            json.dump(tabla_proveedores, f, indent=4, ensure_ascii=False)

        logging.info(f"Tabla única de proveedores generada exitosamente.")

    except Exception as e:
        logging.error(f"Error al generar tabla única de proveedores: {e}")

        
def generar_cruce_y_conciliacion():
    """
    Realiza la conciliación manejando facturas con MÚLTIPLES registros en el Maestro.
    Genera un único objeto JSON por factura, consolidando las filas repetidas dentro de una lista
    y comparando la suma total contra el PDF.
    """
    logging.info(f"Iniciando Conciliación (Modo Agrupación de Duplicados)...")
    
    conciliacion_detalle = []
    discrepancias_monetarias = []

    try:
        # --- 1. CARGA DE DATOS ---
        df_m = pd.read_excel(RUTA_MAESTRO_LIMPIO)
        df_m.columns = [str(c).strip() for c in df_m.columns]
        
        with open(NOMBRE_JSON_SALIDA, 'r', encoding='utf-8') as f:
            datos_ia = json.load(f)
        df_ia = pd.DataFrame(datos_ia)
        df_ia = df_ia[df_ia['estado'] == 'Exitoso'].copy()

        # NOMBRES DE COLUMNAS
        col_nit_m = "NIT"
        col_fact_m = "Invoice ID"
        col_nom_m = "Supplier/Beneficiary Name"
        col_sub_m = "Subtotal"
        col_iva_m = "VAT/WHT1"
        col_fecha_m = "Invoice Date"

        # --- 2. GENERACIÓN DE LLAVES ---
        def limpiar_id(val):
            if pd.isna(val) or str(val).strip().upper() in ["N/A", "NAN", ""]: return ""
            return re.sub(r'[^A-Z0-9]', '', str(val).upper().strip())

        df_m['nit_join'] = df_m[col_nit_m].apply(limpiar_id)
        df_m['fact_join'] = df_m[col_fact_m].apply(limpiar_id)
        
        df_ia['nit_join'] = df_ia['nit_proveedor'].apply(limpiar_id)
        df_ia['fact_join'] = df_ia['numero_factura'].apply(limpiar_id)

        # --- 3. LOGICA PRINCIPAL: ITERAR POR LLAVES ÚNICAS ---
        # Obtenemos el universo de facturas (Llaves únicas)
        llaves_maestro = set(zip(df_m['nit_join'], df_m['fact_join']))
        llaves_ia = set(zip(df_ia['nit_join'], df_ia['fact_join']))
        todas_las_llaves = llaves_maestro.union(llaves_ia)

        for nit, factura in todas_las_llaves:
            if not nit or not factura: continue

            # Obtenemos TODOS los registros asociados a esta llave
            # (Aquí es donde capturamos los duplicados del Maestro)
            registros_m = df_m[(df_m['nit_join'] == nit) & (df_m['fact_join'] == factura)]
            registros_ia = df_ia[(df_ia['nit_join'] == nit) & (df_ia['fact_join'] == factura)]

            # Estado de Clasificación
            existe_m = not registros_m.empty
            existe_ia = not registros_ia.empty

            clasificacion = ""
            if existe_m and existe_ia: clasificacion = "En ambos"
            elif existe_m: clasificacion = "Solo en Maestro"
            else: clasificacion = "Solo en PDFs"

            # Estructura base del Item
            item = {
                "nit": nit,
                "factura": factura,
                "clasificacion": clasificacion,
                "alertas": [],
                "detalle_maestro": [], # Aquí guardaremos las múltiples filas
                "comparacion_global": "N/A"
            }

            # --- PROCESAMIENTO MAESTRO (Puede haber múltiples filas) ---
            suma_subtotal_m = 0.0
            suma_iva_m = 0.0
            nombre_m_ref = "N/A"
            fecha_m_ref = "N/A"

            if existe_m:
                count_rows = 0
                for _, row in registros_m.iterrows():
                    count_rows += 1
                    
                    # Parsear valores de esta fila específica
                    val_sub = parsear_valor_interno(row.get(col_sub_m))
                    val_iva = parsear_valor_interno(row.get(col_iva_m))
                    
                    suma_subtotal_m += val_sub
                    suma_iva_m += val_iva
                    
                    # Guardamos referencias de texto (tomamos la primera válida)
                    if nombre_m_ref == "N/A": nombre_m_ref = str(row.get(col_nom_m, "N/A"))
                    if fecha_m_ref == "N/A": fecha_m_ref = str(row.get(col_fecha_m, "N/A"))

                    # Agregamos al desglose para que lo veas en el JSON
                    item["detalle_maestro"].append({
                        "fila_origen": count_rows,
                        "subtotal_linea": val_sub,
                        "iva_linea": val_iva,
                        "concepto": str(row.get("Line Item Description", "Sin descripción")) # Opcional si existe
                    })

                # ALERTA DE DUPLICADOS
                if count_rows > 1:
                    item["alertas"].append(f"Factura repetida en Maestro ({count_rows} registros)")

            # --- PROCESAMIENTO PDF (Normalmente es 1 fila) ---
            val_sub_ia = 0.0
            val_iva_ia = 0.0
            nombre_ia = "N/A"
            fecha_ia = "N/A"

            if existe_ia:
                # Tomamos el primer registro (asumiendo que el PDF es único por factura)
                row_ia = registros_ia.iloc[0]
                val_sub_ia = parsear_valor_interno(row_ia.get('subtotal'))
                val_iva_ia = parsear_valor_interno(row_ia.get('iva_monto'))
                nombre_ia = str(row_ia.get('nombre_proveedor', "N/A"))
                fecha_ia = str(row_ia.get('fecha_emision', "N/A"))

            # --- COMPARACIÓN (Solo si está en ambos) ---
            if clasificacion == "En ambos":
                diff_sub = abs(suma_subtotal_m - val_sub_ia)
                
                item["comparacion_global"] = {
                    "proveedor": {"maestro": nombre_m_ref, "pdf": nombre_ia},
                    "subtotal": {
                        "maestro_total_calculado": suma_subtotal_m, # La suma de los duplicados
                        "pdf": val_sub_ia,
                        "diferencia": diff_sub
                    },
                    "iva": {
                        "maestro_total_calculado": suma_iva_m,
                        "pdf": val_iva_ia
                    },
                    "fecha": {"maestro": fecha_m_ref, "pdf": fecha_ia}
                }

                # ALERTAS DE DISCREPANCIA
                if nombre_m_ref.strip().upper() != nombre_ia.strip().upper():
                    item["alertas"].append("Diferencia en Nombre")
                
                if diff_sub > 5.0: # Umbral de tolerancia
                    item["alertas"].append(f"Diferencia Subtotal Global: ${diff_sub:,.2f}")
                    discrepancias_monetarias.append({"nit": nit, "diff": diff_sub})
                
                if abs(suma_iva_m - val_iva_ia) > 5.0:
                    item["alertas"].append("Diferencia en IVA Global")

                if not item["alertas"]:
                    item["alertas"].append("Coincidencia Total")

            else:
                item["alertas"].append(f"Documento único en {clasificacion}")

            conciliacion_detalle.append(item)

        # --- 4. RESUMEN EJECUTIVO ---
        discrepancias_monetarias.sort(key=lambda x: x['diff'], reverse=True)
        
        resumen_ejecutivo = {
            "metricas_globales": {
                "total_facturas_unicas_analizadas": len(conciliacion_detalle),
                "coincidencias_completas": sum(1 for x in conciliacion_detalle if x['clasificacion'] == "En ambos"),
                "total_discrepancias_monto": sum(x['diff'] for x in discrepancias_monetarias)
            },
            "top_discrepancias": discrepancias_monetarias[:5],
            "hallazgos": [
                f"Facturas con registros múltiples en Excel: {sum(1 for x in conciliacion_detalle if len(x['detalle_maestro']) > 1)}",
                f"Registros únicos en PDF: {sum(1 for x in conciliacion_detalle if x['clasificacion'] == 'Solo en PDFs')}"
            ]
        }

        return conciliacion_detalle, resumen_ejecutivo

    except Exception as e:
        logging.error(f"Error en conciliación: {e}", exc_info=True)
        return [], {"error": str(e)}    
    

    except Exception as e:
        logging.error(f"Error en conciliación: {e}", exc_info=True)
        return [], {"error": str(e)}
    
def consolidar_reporte_final_unificado():
    logging.info(f"Consolidando Reporte Maestro Final...")
    
    # Obtenemos los nuevos datos de cruce
    detalle_cruce, resumen_ejecutivo = generar_cruce_y_conciliacion()
    
    reporte_final = {
        "titulo_general": "INFORME INTEGRAL DE AUDITORÍA CONTABLE IA",
        "analisis_maestro_excel": {}, # Bloque descriptivo inicial
        "analisis_calidad_tipologia_pdf": {}, # Bloque de calidad
        "EXTRACCIÓN DE DATOS PDF": [], # Bloque de la IA
        "cruce y conciliación": detalle_cruce, # Tu nueva sección
        "resumen ejecutivo": resumen_ejecutivo # El cierre estadístico
    }

    # Carga de archivos previos (Análisis Maestro, Calidad, etc.)
    try:
        if os.path.exists(NOMBRE_JSON_ANALISIS):
            with open(NOMBRE_JSON_ANALISIS, 'r', encoding='utf-8') as f:
                reporte_final["analisis_maestro_excel"] = json.load(f)
        
        if os.path.exists("Analisis_Calidad_PDFs.json"):
            with open("Analisis_Calidad_PDFs.json", 'r', encoding='utf-8') as f:
                reporte_final["analisis_calidad_tipologia_pdf"] = json.load(f)

        if os.path.exists(NOMBRE_JSON_SALIDA):
            with open(NOMBRE_JSON_SALIDA, 'r', encoding='utf-8') as f:
                reporte_final["EXTRACCIÓN DE DATOS PDF"] = json.load(f)

        # Guardar el "Gran Reporte"
        with open("REPORTE_SISTEMA_AUDITORIA_IA.json", 'w', encoding='utf-8') as f:
            json.dump(reporte_final, f, indent=4, ensure_ascii=False)
            
        print("\n ¡SISTEMA COMPLETADO! Revisa 'REPORTE_SISTEMA_AUDITORIA_IA.json'")

    except Exception as e:
        logging.error(f"Error unificando: {e}")

if __name__ == "__main__":
    #procesar_todo() 
    generar_analisis_calidad_pdf()
    generar_maestro_limpio()
    generar_analisis_descriptivo_json()
    consolidar_reporte_final_unificado()

