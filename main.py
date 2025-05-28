from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, NamedStyle, Border, Side, PatternFill, Font
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage
import win32com.client as win32
from openpyxl.utils import get_column_letter
import pandas as pd
import time
import pyautogui
import os
import glob
import random
import xlwings as xw
import pdfkit
import os
import pandas as pd
import glob

# Obtener la ruta base del directorio donde está el script
base_dir = os.path.dirname(os.path.abspath(__file__))

# Definir rutas a las carpetas y archivos
input_folder_excel = os.path.join(base_dir, "data", "input", "Deudas")
output_folder_csv = os.path.join(base_dir, "data", "input", "DeudasCSV")
output_file_csv = os.path.join(base_dir, "data", "Resumen_deudas.csv")
output_file_xlsx = os.path.join(base_dir, "data", "Resumen_deudas.xlsx")

# Leer el archivo Excel
input_excel_clientes = os.path.join(base_dir, "data", "input", "clientes.xlsx")
df = pd.read_excel(input_excel_clientes)

# Suposición de nombres de columnas
cuit_login_list = df['CUIT para ingresar'].tolist()
cuit_represent_list = df['CUIT representado'].tolist()
password_list = df['Contraseña'].tolist()
download_list = df['Ubicacion descarga'].tolist()
posterior_list = df['Posterior'].tolist()
anterior_list = df['Anterior'].tolist()
clientes_list = df['Cliente'].tolist()

# Configuración de opciones de Chrome
options = Options()
options.add_argument("--start-maximized")

# Configurar preferencias de descarga
prefs = {
    "download.prompt_for_download": True,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)

# Inicializar driver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# Crear el archivo de resultados
resultados = []

def human_typing(element, text):
    for char in str(text):
        element.send_keys(char)
        time.sleep(random.uniform(0.05, 0.3))

def actualizar_excel(row_index, mensaje):
    """Actualiza la última columna del archivo Excel con un mensaje de error."""
    df.at[row_index, 'Error'] = mensaje
    df.to_excel(input_excel_clientes, index=False)

def iniciar_sesion(cuit_ingresar, password, row_index):
    """Inicia sesión en el sitio web con el CUIT y contraseña proporcionados."""
    try:
        driver.get('https://auth.afip.gob.ar/contribuyente_/login.xhtml')
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'F1:username')))
        element.clear()
        time.sleep(5)

        human_typing(element, cuit_ingresar)
        driver.find_element(By.ID, 'F1:btnSiguiente').click()
        time.sleep(5)

        # Verificar si el CUIT es incorrecto
        try:
            error_message = driver.find_element(By.ID, 'F1:msg').text
            if error_message == "Número de CUIL/CUIT incorrecto":
                actualizar_excel(row_index, "Número de CUIL/CUIT incorrecto")
                return False
        except:
            pass

        element_pass = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'F1:password')))
        human_typing(element_pass, password)
        time.sleep(15)
        driver.find_element(By.ID, 'F1:btnIngresar').click()
        time.sleep(5)

        # Verificar si la contraseña es incorrecta
        try:
            error_message = driver.find_element(By.ID, 'F1:msg').text
            if error_message == "Clave o usuario incorrecto":
                actualizar_excel(row_index, "Clave incorrecta")
                return False
        except:
            pass

        return True
    except Exception as e:
        print(f"Error al iniciar sesión: {e}")
        actualizar_excel(row_index, "Error al iniciar sesión")
        return False

def ingresar_modulo(cuit_ingresar, password, row_index):
    """Ingresa al módulo específico del sistema de cuentas tributarias."""

    # Verificar si el botón "Ver todos" está presente y hacer clic
    boton_ver_todos = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, "Ver todos")))
    if boton_ver_todos:
        boton_ver_todos.click()
        time.sleep(5)

    # Buscar input del buscador y escribir
    buscador = driver.find_element(By.ID, 'buscadorInput')
    if buscador:
        human_typing(buscador, 'tas tr') 
        time.sleep(5)

    # Seleccionar la opción del menú
    opcion_menu = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'rbt-menu-item-0')))
    if opcion_menu:
        opcion_menu.click()
        time.sleep(10)

    # Manejar modal si aparece
    modales = driver.find_elements(By.CLASS_NAME, 'modal-content')
    if modales and modales[0].is_displayed():
        boton_continuar = driver.find_element(By.XPATH, '//button[text()="Continuar"]')
        if boton_continuar:
            boton_continuar.click()
            time.sleep(5)

    # Cambiar a la última pestaña abierta
    driver.switch_to.window(driver.window_handles[-1])

    # Verificar mensaje de error de autenticación
    error_message_elements = driver.find_elements(By.TAG_NAME, 'pre')
    if error_message_elements and error_message_elements[0].text == "Ha ocurrido un error al autenticar, intente nuevamente.":
        actualizar_excel(row_index, "Error autenticacion")
        driver.refresh()
        time.sleep(5)

    # Verificar si es necesario iniciar sesión nuevamente
    username_input = driver.find_elements(By.ID, 'F1:username')
    if username_input:
        username_input[0].clear()
        time.sleep(5)
        human_typing(username_input[0], cuit_ingresar)
        driver.find_element(By.ID, 'F1:btnSiguiente').click()
        time.sleep(5)

        password_input = driver.find_elements(By.ID, 'F1:password')
        if password_input:
            human_typing(password_input[0], password)
            time.sleep(15)
            driver.find_element(By.ID, 'F1:btnIngresar').click()
            time.sleep(5)
            actualizar_excel(row_index, "Error volver a iniciar sesion")

def seleccionar_cuit_representado(cuit_representado):
    """Selecciona el CUIT representado en el sistema."""
    try:
        select_present = EC.presence_of_element_located((By.NAME, "$PropertySelection"))
        if WebDriverWait(driver, 5).until(select_present):
            current_selection = Select(driver.find_element(By.NAME, "$PropertySelection")).first_selected_option.text
            if current_selection != str(cuit_representado):
                select_element = Select(driver.find_element(By.NAME, "$PropertySelection"))
                select_element.select_by_visible_text(str(cuit_representado))
    except Exception:
        try:
            cuit_element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'span.cuit')))
            cuit_text = cuit_element.text.replace('-', '')
            if cuit_text != str(cuit_representado):
                print(f"El CUIT ingresado no coincide con el CUIT representado: {cuit_representado}")
                return False
        except Exception as e:
            print(f"Error al verificar CUIT: {e}")
            return False
    # Esperar que el popup esté visible y hacer clic en el botón de cerrar por XPATH
    try:
    # Usamos el XPATH para localizar el botón de cerrar
        xpath_popup = "/html/body/div[2]/div[2]/div/div/a"
        element_popup = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath_popup)))
        element_popup.click()
        print("Popup cerrado exitosamente.")
    except Exception as e:
        print(f"Error al intentar cerrar el popup: {e}")
    return True

def exportar_excel(ubicacion_descarga, cuit_representado, cliente):
    """Descarga y guarda el archivo Excel en la ubicación especificada."""
    try:       
        # Exportar XLSX
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, "//*[@id='DataTables_Table_0_wrapper']/div[1]/a[2]/span"))).click()
        time.sleep(5)

        # Guardarlo con nombre y carpeta especifica
        nombre_archivo = f"Reporte - {cliente}.xlsx"
        pyautogui.write(nombre_archivo)
        time.sleep(1)
        pyautogui.hotkey('alt', 'd')
        time.sleep(0.5)
        pyautogui.write(ubicacion_descarga)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.hotkey('alt', 't')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1)
    except Exception as e:
        print(f"Error al exportar el archivo Excel: {e}")

def cerrar_sesion():
    """Cierra la sesión actual."""
    try:
        driver.close()
        window_handles = driver.window_handles
        driver.switch_to.window(window_handles[0])
        driver.find_element(By.ID, "iconoChicoContribuyenteAFIP").click()
        driver.find_element(By.XPATH, '//*[@id="contBtnContribuyente"]/div[6]/button/div/div[2]').click()
        time.sleep(5)
    except Exception as e:
        print(f"Error al cerrar sesión: {e}")

def extraer_datos_nuevo(cuit_ingresar, cuit_representado, password, ubicacion_descarga, posterior, cliente, indice):
    """Extrae datos para un nuevo usuario."""
    try:
        control_sesion = iniciar_sesion(cuit_ingresar, password, indice)
        if control_sesion:
            ingresar_modulo(cuit_ingresar, password, indice)
            # Esperar que el popup esté visible y hacer clic en el botón de cerrar por XPATH
            try:
                xpath_popup = "/html/body/div[2]/div[2]/div/div/a"
                element_popup = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath_popup)))
                element_popup.click()
                print("Popup cerrado exitosamente.")
            except Exception as e:
                print(f"Error al intentar cerrar el popup: {e}")
            if seleccionar_cuit_representado(cuit_representado):
                exportar_excel(ubicacion_descarga, cuit_representado, cliente)
                if posterior == 0:
                    cerrar_sesion()
    except Exception as e:
        print(f"Error al extraer datos para el nuevo usuario: {e}")

def extraer_datos(cuit_representado, ubicacion_descarga, posterior, cliente):
    """Extrae datos para un usuario existente."""
    try:
        if seleccionar_cuit_representado(cuit_representado):
            exportar_excel(ubicacion_descarga, cuit_representado, cliente)
            if posterior == 0:
                cerrar_sesion()
    except Exception as e:
        print(f"Error al extraer datos: {e}")

# Función para convertir Excel a CSV utilizando xlwings
def excel_a_csv(input_folder, output_folder):
    for excel_file in glob.glob(os.path.join(input_folder, "*.xlsx")):
        try:
            app = xw.App(visible=False)
            wb = app.books.open(excel_file)
            sheet = wb.sheets[0]
            df = sheet.used_range.options(pd.DataFrame, header=1, index=False).value

            # Convertir la columna 'FechaVencimiento' a datetime, ajustar según sea necesario
            if 'FechaVencimiento' in df.columns:
                df['FechaVencimiento'] = pd.to_datetime(df['FechaVencimiento'], errors='coerce')

            wb.close()
            app.quit()

            base = os.path.basename(excel_file)
            csv_file = os.path.join(output_folder, base.replace('.xlsx', '.csv'))
            df.to_csv(csv_file, index=False, encoding='utf-8-sig', sep=';')
            print(f"Convertido {excel_file} a {csv_file}")
        except Exception as e:
            print(f"Error al convertir {excel_file} a CSV: {e}")

# Función para obtener el nombre del cliente a partir del nombre del archivo
def obtener_nombre_cliente(filename):
    base = os.path.basename(filename)
    nombre_cliente = base.split('-')[1].strip()
    return nombre_cliente

# Iterar sobre cada cliente
indice = 0
for cuit_ingresar, cuit_representado, password, download, posterior, anterior, cliente in zip(cuit_login_list, cuit_represent_list, password_list, download_list, posterior_list, anterior_list, clientes_list):
    if anterior == 0:
        extraer_datos_nuevo(cuit_ingresar, cuit_representado, password, download, posterior, cliente, indice)
    else:
        extraer_datos(cuit_representado, download, posterior, cliente)
    indice = indice + 1

output_folder_pdf = os.path.join(base_dir, "data", "Reportes")
imagen = os.path.join(base_dir, "data", "imagen.png")

def forzar_guardado_excel(excel_file):
    try:
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(excel_file)
        wb.Save()
        wb.Close(False)
    except Exception as e:
        print(f"Error forzando guardado en {excel_file}: {e}")
    finally:
        excel.Quit()

def ajustar_diseno_excel(ws):
    """
    Ajusta el diseño del archivo Excel para que todo el contenido (imagen y tabla) 
    quepa en una sola página PDF.
    """
    # Configurar ajuste de página para que quepa todo en una página
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.orientation = "landscape"  # Apaisado
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

def procesar_excel(excel_file, output_pdf, imagen):
    # FUNCIÓN AUXILIAR para convertir números argentinos
    def convertir_numero_argentino(valor_str):
        """
        Convierte números en formato argentino a float.
        Ejemplos:
        - "63.133,14" -> 63133.14
        - "151,61" -> 151.61
        - "1.234.567,89" -> 1234567.89
        - "0,00" -> 0.0
        """
        try:
            if not valor_str or valor_str in ['', ' ', None]:
                return 0.0
            
            # Convertir a string y limpiar espacios
            valor_str = str(valor_str).strip()
            
            # Si contiene tanto punto como coma (formato: 63.133,14)
            if '.' in valor_str and ',' in valor_str:
                # Los puntos son separadores de miles, la coma es decimal
                valor_str = valor_str.replace('.', '')  # Eliminar puntos (miles)
                valor_str = valor_str.replace(',', '.')  # Convertir coma a punto (decimal)
            
            # Si solo contiene coma (formato: 151,61)
            elif ',' in valor_str and '.' not in valor_str:
                # La coma es el separador decimal
                valor_str = valor_str.replace(',', '.')
            
            # Si solo contiene punto o ninguno, asumir que ya está en formato correcto
            
            return float(valor_str)
        
        except (ValueError, TypeError, AttributeError):
            return 0.0

    try:
        # Ignorar archivos temporales de Excel que comienzan con ~$
        if os.path.basename(excel_file).startswith('~$'):
            print(f"Omitiendo archivo temporal: {excel_file}")
            return

        # Extraer el nombre del cliente desde el nombre del archivo
        nombre_archivo = os.path.basename(excel_file)
        # Asumiendo que el formato es "Reporte - NOMBRE_CLIENTE.xlsx"
        if " - " in nombre_archivo:
            cliente = nombre_archivo.split(" - ")[1].replace(".xlsx", "")
        else:
            # Fallback si el formato es diferente
            cliente = nombre_archivo.replace(".xlsx", "")
        
        print(f"Procesando cliente: {cliente}")

        # Cargar el archivo Excel con pandas
        df = pd.read_excel(excel_file)

        # Definir la lista de impuestos a incluir en el filtro
        impuestos_incluir = [
            'ganancias sociedades',
            'iva',
            'bp-acciones o participaciones',
            'sicore-impto.a las ganancias',
            'empleador-aportes seg. social',
            'contribuciones seg. social',
            'ret art 79 ley gcias in a,byc'
        ]

        # Filtrar por múltiples tipos de "Impuesto"
        # Creo una condición que sea true si el impuesto contiene cualquiera de los términos buscados
        condicion_impuestos = df['Impuesto'].str.contains('|'.join(impuestos_incluir), case=False, na=False)
        
        df_filtrado = df[condicion_impuestos].copy()  # Crear una copia para evitar SettingWithCopyWarning

        print(f"Impuestos buscados: {impuestos_incluir}")
        print(f"Registros encontrados con estos impuestos: {len(df_filtrado)}")

        # Obtener la fecha actual y el año actual
        fecha_actual = pd.Timestamp.now().date()
        año_actual = fecha_actual.year
        fecha_inicio = pd.Timestamp(year=año_actual - 7, month=1, day=1).date()  # 1 de enero de 8 años atrás
        
        print(f"Rango de fechas: desde {fecha_inicio} hasta {fecha_actual}")

        # Identificar el nombre correcto de la columna de fecha
        columna_fecha_encontrada = None
        posibles_columnas_fecha = ['FechaVencimiento', 'Fecha de Vencimiento', 'Fecha Vencimiento']
        
        for col in posibles_columnas_fecha:
            if col in df_filtrado.columns:
                columna_fecha_encontrada = col
                break
        
        # Si encontramos la columna de fecha, procesarla y filtrar
        if columna_fecha_encontrada:
            # Convertir a datetime con formato específico y dayfirst=True para formato dd/mm/yyyy
            df_filtrado['fecha_procesada'] = pd.to_datetime(
                df_filtrado[columna_fecha_encontrada], 
                errors='coerce',
                dayfirst=True,  # Especificar que el día va primero (formato dd/mm/yyyy)
                format='%d/%m/%Y'  # Especificar el formato explícitamente
            ).dt.date
            
            # Imprimir información de diagnóstico
            print(f"Registros antes de filtrar por fecha: {len(df_filtrado)}")
            
            # Filtrar solo por fecha de vencimiento menor a fecha actual (vencido)
            mascara_fechas_validas = df_filtrado['fecha_procesada'].notna()
            
            # Aplicar filtro solo por fecha
            df_filtrado = df_filtrado[
                mascara_fechas_validas & 
                (df_filtrado['fecha_procesada'] >= fecha_inicio) &
                (df_filtrado['fecha_procesada'] <= fecha_actual)
            ]
            
            print(f"Registros después de filtrar por fecha: {len(df_filtrado)}")
            
            # Eliminar la columna temporal después de filtrar
            df_filtrado = df_filtrado.drop(['fecha_procesada'], axis=1)
        else:
            print(f"Advertencia: No se encontró columna de fecha de vencimiento en {excel_file}")
            print(f"Columnas disponibles: {list(df_filtrado.columns)}")

        # Verificar si la tabla está vacía
        if df_filtrado.shape[0] == 0:
            output_pdf = output_pdf.replace(".pdf", " - vacio.pdf")
            print(f"No se encontraron registros que cumplan con los criterios en {excel_file}")

        # Eliminar la columna de Int. punitorios y concepto / subconcepto
        columnas_a_eliminar = ['Int. punitorios', 'Concepto / Subconcepto']
        for columna in columnas_a_eliminar:
            if columna in df_filtrado.columns:
                df_filtrado = df_filtrado.drop(columna, axis=1)

        # Guardar el DataFrame filtrado en el archivo Excel
        df_filtrado.to_excel(excel_file, index=False)

        # Cargar el archivo para aplicar formato con openpyxl
        wb = load_workbook(excel_file)
        ws = wb.active

        # Insertar filas adicionales para una nueva imagen
        ws.insert_rows(1, amount=7)

        # Agregar una imagen encima del encabezado (A1)
        # Obtener el ancho combinado de la tabla
        ultima_columna = ws.max_column
        ultima_letra_columna = get_column_letter(ultima_columna)

        # Insertar la imagen
        img = ExcelImage(imagen)
        # Ajustar el tamaño de la imagen
        img.width = ws.column_dimensions['A'].width * ultima_columna * 6  # Ajustar al ancho combinado
        img.height = 120  # Altura fija
        # Agregar la imagen a la hoja
        ws.add_image(img, 'A1')

        # Insertar filas adicionales para una nueva imagen
        ws.insert_rows(7, amount=1)

        # Fila donde se agregará el texto
        fila_texto = 8

        # Obtener el número de columnas ocupadas por la tabla
        ultima_columna = ws.max_column
        ultima_letra_columna = get_column_letter(ultima_columna)

        # Combinar celdas en la fila de separación
        ws.merge_cells(f'A{fila_texto}:{ultima_letra_columna}{fila_texto}')

        # Establecer el texto en la celda combinada
        celda_texto = ws[f'A{fila_texto}']
        celda_texto.value = f"Reporte de deudas del SCT - {cliente} "

        # Aplicar formato centrado y en negrita
        celda_texto.alignment = Alignment(horizontal='center', vertical='center')
        celda_texto.font = Font(bold=True, size=20)

        # Cambiar el color del encabezado a lila
        header_fill = PatternFill(start_color="AA0EAA", end_color="AA0EAA", fill_type="solid")
        for cell in ws[9]:
            cell.fill = header_fill

        # Ajustar el ancho de las columnas con control específico para "Impuesto"
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter  # Get the column name
            column_header = ""
            
            # Obtener el nombre del encabezado de la columna
            for cell in col:
                if cell.row == 9 and cell.value:  # Fila 9 es donde están los encabezados
                    column_header = str(cell.value).lower()
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            
            # Ajuste especial para la columna "Impuesto"
            if "impuesto" in column_header:
                # Limitar el ancho máximo de la columna Impuesto a 35 caracteres
                adjusted_width = min(35, (max_length + 2) * 1.2)
            else:
                # Para el resto de columnas, usar el cálculo normal
                adjusted_width = (max_length + 2) * 1.2
            
            ws.column_dimensions[column].width = adjusted_width

        # Encontrar las columnas "Fecha de Vencimiento", "Saldo" e "Int. resarcitorios" para totales y alineación
        fecha_vencimiento_col = None
        saldo_col = None
        int_resarcitorios_col = None
        columnas_derecha = []
        header_row = 9  # Fila donde están los encabezados
        
        for col_num, cell in enumerate(ws[header_row], 1):
            if cell.value and isinstance(cell.value, str):
                cell_value_lower = cell.value.lower()
                
                # Buscar columna Fecha de Vencimiento
                if any(term in cell_value_lower for term in ['fecha de vencimiento', 'fechavencimiento', 'vencimiento']):
                    fecha_vencimiento_col = col_num
                
                # Buscar columna Saldo
                elif 'saldo' in cell_value_lower:
                    saldo_col = col_num
                    columnas_derecha.append(col_num)
                
                # Buscar columna Int. resarcitorios
                elif any(term in cell_value_lower for term in ['int. resarcitorios', 'int.resarcitorios', 'resarcitorios']):
                    int_resarcitorios_col = col_num
                    columnas_derecha.append(col_num)
        
        # Aplicar alineación a todas las celdas
        for row in ws.iter_rows(min_row=9, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for col_num, cell in enumerate(row, 1):
                if col_num in columnas_derecha:
                    # Alinear a la derecha las columnas de Saldo e Int. resarcitorios
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                else:
                    # Centrar el resto de columnas
                    cell.alignment = Alignment(horizontal='center', vertical='center')

        # Encontrar la última fila con datos y agregar fila de totales
        ultima_fila_datos = ws.max_row
        fila_total = ultima_fila_datos + 1
        
        # Agregar "TOTAL" en la columna de fecha de vencimiento
        if fecha_vencimiento_col:
            celda_total = ws.cell(row=fila_total, column=fecha_vencimiento_col)
            celda_total.value = "TOTAL"
            celda_total.font = Font(bold=True)
            celda_total.alignment = Alignment(horizontal='center', vertical='center')
        
        # Calcular y agregar sumatoria de Saldo
        if saldo_col:
            suma_saldo = 0
            for fila in range(10, ultima_fila_datos + 1):  # Desde fila 10 hasta la última
                celda_saldo = ws.cell(row=fila, column=saldo_col)
                if celda_saldo.value:
                    suma_saldo += convertir_numero_argentino(celda_saldo.value)
            
            # Insertar la suma en la fila total
            celda_suma_saldo = ws.cell(row=fila_total, column=saldo_col)
            celda_suma_saldo.value = round(suma_saldo, 2)
            celda_suma_saldo.font = Font(bold=True)
            celda_suma_saldo.alignment = Alignment(horizontal='right', vertical='center')

        # Calcular y agregar sumatoria de Int. resarcitorios
        if int_resarcitorios_col:
            suma_int_resarcitorios = 0
            for fila in range(10, ultima_fila_datos + 1):  # Desde fila 10 hasta la última
                celda_int = ws.cell(row=fila, column=int_resarcitorios_col)
                if celda_int.value:
                    suma_int_resarcitorios += convertir_numero_argentino(celda_int.value)
            
            # Insertar la suma en la fila total
            celda_suma_int = ws.cell(row=fila_total, column=int_resarcitorios_col)
            celda_suma_int.value = round(suma_int_resarcitorios, 2)
            celda_suma_int.font = Font(bold=True)
            celda_suma_int.alignment = Alignment(horizontal='right', vertical='center')

        # Guardar los cambios
        wb.save(excel_file)

        # Convertir el archivo Excel a PDF con pywin32
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(excel_file)

        print("Configurando área de impresión...")
        ws = wb.Worksheets(1)

        # Definir el rango del área de impresión manualmente
        last_row = ws.UsedRange.Rows.Count
        last_col = ws.UsedRange.Columns.Count
        ws.PageSetup.PrintArea = f"A1:{get_column_letter(last_col)}{last_row + 8}"  # Incluir imagen y tabla

        # Ajustar a una página
        ws.PageSetup.Orientation = 2  # Paisaje
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1

        # Configurar centrado en la página
        ws.PageSetup.CenterHorizontally = True
        ws.PageSetup.CenterVertically = False  # Verticalmente opcional, según el diseño

        # Configurar márgenes
        ws.PageSetup.LeftMargin = 0.25
        ws.PageSetup.RightMargin = 0.25
        ws.PageSetup.TopMargin = 0.5
        ws.PageSetup.BottomMargin = 0.5

        print("Guardando como PDF...")
        wb.ExportAsFixedFormat(0, output_pdf)  # 0 indica formato PDF
        wb.Close(False)
        print(f"Archivo convertido a PDF: {output_pdf}")

    except Exception as e:
        print(f"Error al procesar {excel_file}: {e}")
        import traceback
        traceback.print_exc()  # Imprimir el traceback completo para mejor diagnóstico
    finally:
        if 'excel' in locals():
            excel.Quit()

# Recorrer todos los archivos Excel en la carpeta
for excel_file in glob.glob(os.path.join(input_folder_excel, "*.xlsx")):
    try:
        # Forzar guardado para evitar problemas con archivos corruptos o no calculados
        forzar_guardado_excel(excel_file)

        # Obtener el nombre base del archivo para usarlo en el nombre del PDF
        base_name = os.path.splitext(os.path.basename(excel_file))[0]
        output_pdf = os.path.join(output_folder_pdf, f"{base_name}.pdf")
        
        # Llamar a la función para procesar el archivo Excel y generar el PDF
        procesar_excel(excel_file, output_pdf, imagen)
        
        print(f"Archivo {excel_file} procesado y guardado como {output_pdf}")
    
    except Exception as e:
        print(f"Error al procesar {excel_file}: {e}")
