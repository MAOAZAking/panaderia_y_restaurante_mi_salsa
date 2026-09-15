import tkinter as tk
from tkinter import messagebox
import json
import os
import sys
from datetime import datetime
import urllib.request
import base64
import re
import openpyxl

# ==========================================
# CONFIGURACIÓN DE RUTAS Y ARCHIVOS JSON
# ==========================================
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

PROD_JSON = os.path.join(application_path, "productos_y_precios.json")
CLI_JSON = os.path.join(application_path, "clientes.json")
CONFIG_JSON = os.path.join(application_path, "config.json")

# ==========================================
# FUNCIONES AUXILIARES Y DE EXCEL
# ==========================================
def es_bandeja(texto):
    palabras = texto.lower().split()
    for p in palabras:
        if p.startswith("band") or p in ["bdja", "bandj", "bande"]:
            return True
    return False

def es_almuerzo(texto):
    palabras = texto.lower().split()
    for p in palabras:
        if p.startswith("almu") or p.startswith("amuer") or p in ["alm", "almuer"]:
            return True
    return False

def extraer_cliente_y_nit_cc(cliente_input):
    if not cliente_input or not cliente_input.strip():
        return "Nombre", "222222222222"
    
    pattern = r'(?i)\b(nit|cc)\b[:.]?\s*([0-9\-]+)'
    match = re.search(pattern, cliente_input)
    
    if match:
        nit_cc = match.group(2)
        nombre_limpio = re.sub(pattern, '', cliente_input).strip()
        if not nombre_limpio:
            nombre_limpio = "CONSUMIDOR FINAL"
        else:
            nombre_limpio = nombre_limpio.title() if nombre_limpio.lower() != "consumidor final" else "CONSUMIDOR FINAL"
        return nombre_limpio, nit_cc
    else:
        nombre_limpio = cliente_input.strip()
        nombre_limpio = nombre_limpio.title() if nombre_limpio.lower() != "consumidor final" else "CONSUMIDOR FINAL"
        return nombre_limpio, "222222222222"

def formatear_cliente_para_db(texto):
    if not texto or not texto.strip():
        return "CONSUMIDOR FINAL"
    p_match = re.search(r'(?i)\b(nit|cc)\b[:.]?\s*([0-9\-]+)', texto)
    if p_match:
        tipo = p_match.group(1).upper()
        num = p_match.group(2)
        nombre_p = re.sub(r'(?i)\b(nit|cc)\b[:.]?\s*[0-9\-]+', '', texto).strip()
        if nombre_p:
            return f"{nombre_p.title()} {tipo} {num}"
        else:
            return f"{tipo} {num}"
    else:
        return texto.title() if texto.lower() != "consumidor final" else "CONSUMIDOR FINAL"

def registrar_cuenta_por_cobrar(nombre_cliente, total):
    try:
        import openpyxl
    except ImportError:
        messagebox.showerror("Error Excel", "La librería 'openpyxl' no está instalada.")
        return

    ruta_xlsx = os.path.join(application_path, "cuenta_por_cobrar.xlsx")
    ruta_xlxs = os.path.join(application_path, "cuenta_por_cobrar.xlxs")
    ruta_final = ruta_xlxs if os.path.exists(ruta_xlxs) else ruta_xlsx

    try:
        if os.path.exists(ruta_final):
            try:
                wb = openpyxl.load_workbook(ruta_final)
                ws = wb.active
            except Exception:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Cuentas por Cobrar"
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Cuentas por Cobrar"

        val_a1 = str(ws.cell(row=1, column=1).value or "").strip()
        if not val_a1:
            ws.cell(row=1, column=1, value="NOMBRE")
            ws.cell(row=1, column=2, value="TOTAL")

        ws.append([nombre_cliente, total])
        wb.save(ruta_final)
    except PermissionError:
        messagebox.showwarning(
            "Archivo Excel Abierto",
            f"No se pudo actualizar '{os.path.basename(ruta_final)}' porque está abierto en Excel.\nPor favor ciérrelo para registrar la cuenta por cobrar."
        )
    except Exception as e:
        messagebox.showerror("Error Excel", f"Error al guardar en cuenta por cobrar: {str(e)}")

def registrar_factura_excel(nombre_cliente, total, fecha_hora):
    try:
        import openpyxl
    except ImportError:
        messagebox.showerror("Error Excel", "La librería 'openpyxl' no está instalada.")
        return

    meses = {
        1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril', 5: 'mayo', 6: 'junio',
        7: 'julio', 8: 'agosto', 9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'
    }
    nombre_hoja_actual = f"{meses[fecha_hora.month]}-{fecha_hora.year}"
    fecha_fmt = fecha_hora.strftime("%d/%m/%Y %H:%M")

    ruta_xlsx = os.path.join(application_path, "facturas.xlsx")
    ruta_xlxs = os.path.join(application_path, "facturas.xlxs")
    ruta_final = ruta_xlxs if os.path.exists(ruta_xlxs) else ruta_xlsx

    try:
        if os.path.exists(ruta_final):
            try:
                wb = openpyxl.load_workbook(ruta_final)
            except Exception:
                wb = openpyxl.Workbook()
        else:
            wb = openpyxl.Workbook()

        hoja_objetivo = None
        if wb.worksheets:
            ultima_hoja = wb.worksheets[-1]
            mismo_mes = False

            if ultima_hoja.title.lower() == nombre_hoja_actual.lower():
                mismo_mes = True
            else:
                max_r = ultima_hoja.max_row
                if max_r > 1:
                    val_fecha = ultima_hoja.cell(row=max_r, column=3).value
                    if val_fecha:
                        val_str = str(val_fecha).strip()
                        try:
                            if isinstance(val_fecha, datetime):
                                if val_fecha.month == fecha_hora.month and val_fecha.year == fecha_hora.year:
                                    mismo_mes = True
                            elif '/' in val_str:
                                partes = val_str.split('/')
                                if len(partes) >= 2 and int(partes[1]) == fecha_hora.month:
                                    mismo_mes = True
                        except Exception:
                            pass

            if mismo_mes:
                hoja_objetivo = ultima_hoja
            elif nombre_hoja_actual in wb.sheetnames:
                hoja_objetivo = wb[nombre_hoja_actual]
            else:
                if len(wb.worksheets) == 1 and wb.worksheets[0].title in ["Sheet", "Hoja"] and wb.worksheets[0].max_row <= 1 and not wb.worksheets[0].cell(row=1, column=1).value:
                    hoja_objetivo = wb.worksheets[0]
                    hoja_objetivo.title = nombre_hoja_actual
                else:
                    hoja_objetivo = wb.create_sheet(title=nombre_hoja_actual)
        else:
            hoja_objetivo = wb.create_sheet(title=nombre_hoja_actual)

        val_a1 = str(hoja_objetivo.cell(row=1, column=1).value or "").strip()
        if not val_a1:
            hoja_objetivo.cell(row=1, column=1, value="NOMBRE")
            hoja_objetivo.cell(row=1, column=2, value="TOTAL")
            hoja_objetivo.cell(row=1, column=3, value="FECHA")

        hoja_objetivo.append([nombre_cliente, total, fecha_fmt])
        wb.save(ruta_final)
    except PermissionError:
        messagebox.showwarning(
            "Archivo Excel Abierto",
            f"No se pudo actualizar '{os.path.basename(ruta_final)}' porque está abierto en Excel.\nPor favor ciérrelo para registrar la factura."
        )
    except Exception as e:
        messagebox.showerror("Error Excel", f"Error al guardar en facturas.xlsx: {str(e)}")

# ==========================================
# FUNCIONES DE ARRANQUE 
# ==========================================
def descargar_archivos_github():
    if not os.path.exists(CONFIG_JSON):
        messagebox.showwarning(
            "Archivo de configuración faltante", 
            "No se encontró el archivo 'config.json' junto al ejecutable.\n\nEl sistema iniciará de forma local y no se conectará a GitHub."
        )
        return

    try:
        with open(CONFIG_JSON, "r", encoding="utf-8") as f:
            config_data = json.load(f)
            TOKEN = config_data.get("github_token")
    except Exception:
        return

    if not TOKEN or TOKEN == "apidegithub" or TOKEN.strip() == "":
        return 

    USUARIO = "MAOAZAking"
    REPO = "panaderia_y_restaurante_mi_salsa"
    archivos_a_bajar = {
        "productos_y_precios.json": PROD_JSON,
        "clientes.json": CLI_JSON
    }

    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "Python-App"
    }

    errores = []
    for nombre_github, ruta_local in archivos_a_bajar.items():
        url = f"https://api.github.com/repos/{USUARIO}/{REPO}/contents/{nombre_github}"
        try:
            req_get = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req_get, timeout=5) as response:
                data_github = json.loads(response.read().decode("utf-8"))
                if "content" in data_github:
                    contenido_decodificado = base64.b64decode(data_github["content"]).decode("utf-8")
                    with open(ruta_local, "w", encoding="utf-8") as f:
                        f.write(contenido_decodificado)
        except Exception as e:
            errores.append(f"No se pudo descargar {nombre_github}: {str(e)}")
    
    if errores:
        mensaje = "No se pudieron descargar los archivos de GitHub.\nSe usarán los datos locales.\n\nDetalles:\n" + "\n".join(errores)
        messagebox.showwarning("Aviso de Sincronización", mensaje)

def crear_archivos_base_si_no_existen():
    if not os.path.exists(PROD_JSON):
        datos_base_prod = {
    "pan cacho": {
        "precio": 1700
    },
    "pan agridulce": {
        "precio": 700,
        "promocion": {
            "cantidad": 3,
            "precio_promo": 2000
        }
    },
    "pan queso": {
        "precio": 700,
        "promocion": {
            "cantidad": 3,
            "precio_promo": 2000
        }
    },
    "borracho": {
        "precio": 700,
        "promocion": {
            "cantidad": 3,
            "precio_promo": 2000
        }
    },
    "desayuno huvos revueltos + arroz": {
        "precio": 6000
    },
    "desayuno huevos pericos + arroz": {
        "precio": 6500
    },
    "pan 2000": {
        "precio": 2000
    },
    "desayuno completo ranchero": {
        "precio": 10000
    },
    "promo pan queso": {
        "precio": 2000
    },
    "porcion huevo ranchero": {
        "precio": 5500
    },
    "agua saborizada manzana": {
        "precio": 2500
    },
    "agua saborizada h2o limon": {
        "precio": 3500
    },
    "gaseosa quatro 1.5l": {
        "precio": 8000
    },
    "jugo del valle 1.5l": {
        "precio": 8000
    },
    "papa aborrajada": {
        "precio": 1500
    },
    "salchichon": {
        "precio": 1500
    },
    "papa con salchichon": {
        "precio": 3000
    },
    "pan 3000": {
        "precio": 3000
    },
    "pan 5000": {
        "precio": 5000
    },
    "chicharron panaderia": {
        "precio": 2500
    },
    "pan dulce guayaba": {
        "precio": 600
    },
    "pan dulce arequipe": {
        "precio": 600
    },
    "roscon": {
        "precio": 2000
    },
    "pan coco": {
        "precio": 600
    },
    "empanada": {
        "precio": 1500
    },
    "pan hawaiano": {
        "precio": 2500
    },
    "pandebono": {
        "precio": 1000
    },
    "bu\u00f1uelo": {
        "precio": 1000
    },
    "arepita": {
        "precio": 1000
    },
    "pan 1000": {
        "precio": 1000
    },
    "almuerzo sancocho de pescado + pescado frito (sin arroz, pero con mas ensalada)": {
        "precio": 16000
    },
    "galleta de pepitas": {
        "precio": 1000
    },
    "achira": {
        "precio": 500
    },
    "almuerzo albondiga, frijoles, res asada": {
        "precio": 16000
    },
    "croissant": {
        "precio": 1700
    },
    "leche mediana": {
        "precio": 3200
    },
    "leche grande": {
        "precio": 6500
    },
    "carne en bistec": {
        "precio": 13000
    },
    "agua natural mini": {
        "precio": 1000
    },
    "agua natural personal": {
        "precio": 2000
    },
    "agua natural 1l": {
        "precio": 3000
    },
    "agua con gas personal": {
        "precio": 3000
    },
    "glacial cola negra personal": {
        "precio": 2500
    },
    "glacial cola negra 1l": {
        "precio": 4000
    },
    "glacial cola roja personal": {
        "precio": 2500
    },
    "glacial cola roja 1l": {
        "precio": 4000
    },
    "glacial manzana personal": {
        "precio": 2500
    },
    "glacial manzana 1l": {
        "precio": 4000
    },
    "glacial uva personal": {
        "precio": 2500
    },
    "glacial uva 1l": {
        "precio": 4000
    },
    "soda glacial personal": {
        "precio": 2500
    },
    "soda glacial 1l": {
        "precio": 4000
    },
    "papa rellena": {
        "precio": 5000
    },
    "libra azucar": {
        "precio": 3000
    },
    "glacial gran cola 1.7l": {
        "precio": 6000
    },
    "speed max lata": {
        "precio": 4000
    },
    "speed max botella": {
        "precio": 5000
    },
    "dona chocolate": {
        "precio": 3500
    },
    "almojabana": {
        "precio": 1000
    },
    "cocacola 1.5l": {
        "precio": 8000
    },
    "paquete de vasos": {
        "precio": 4000
    },
    "torta fria de media libra": {
        "precio": 35000
    },
    "vasos plastico": {
        "precio": 100
    },
    "cucharas clasticas": {
        "precio": 100
    },
    "cocacola 3l": {
        "precio": 13000
    },
    "gelatina amarilla": {
        "precio": 2000
    },
    "gatorade": {
        "precio": 5000
    },
    "yogurt melocoton": {
        "precio": 3000
    },
    "jugo hit": {
        "precio": 3500
    },
    "pepsi 2.5l": {
        "precio": 8000
    },
    "cocacola personal": {
        "precio": 3500
    },
    "papel higienico": {
        "precio": 2000
    },
    "pan abuela": {
        "precio": 3000
    },
    "jugo hit naranja-pi\u00f1a": {
        "precio": 3500
    },
    "galleta de chocolate": {
        "precio": 1000
    },
    "reina": {
        "precio": 2500
    },
    "reina roja": {
        "precio": 2500
    },
    "reina coco": {
        "precio": 2500
    },
    "dona azucar": {
        "precio": 3000
    },
    "dona azucar redonda": {
        "precio": 3000
    },
    "porcion torta chocolate": {
        "precio": 2500
    },
    "porcion torta mantecada": {
        "precio": 2500
    },
    "porcion torta vainilla": {
        "precio": 2500
    },
    "porcion torta banano": {
        "precio": 2500
    },
    "pan ca\u00f1a": {
        "precio": 2000
    },
    "jugo hit mango": {
        "precio": 3500
    },
    "jugo hit mora": {
        "precio": 3500
    },
    "colombiana personal": {
        "precio": 3500
    },
    "desayuno completo cacerola, cafe en leche": {
        "precio": 9000
    },
    "desayuno completo ranchero, cafe con leche": {
        "precio": 11500
    },
    "cafe en leche": {
        "precio": 2500
    },
    "desayuno completo revuelto, cafe en leche, pan queso": {
        "precio": 9200
    },
    "huevo ranchero, arroz": {
        "precio": 8500
    },
    "porcion bistec": {
        "precio": 7000
    },
    "almuerzo sancocho, carne molida (sin principio)": {
        "precio": 16000
    },
    "almuerzo completo sopa, blanquillos, chuleta pollo": {
        "precio": 16000
    },
    "bandeja chuleta de cerdo (sin arroz, en cambio papa o maduro), sin ensalada": {
        "precio": 14000
    },
    "bandeja chuleta de cerdo, frijo (sin ensalada)": {
        "precio": 14000
    },
    "almuerzo completo sopa de fideos, firjol, chuleta de pollo": {
        "precio": 16000
    },
    "bandeja carne molida y blanquillos": {
        "precio": 14000
    },
    "almuerzo completo": {
        "precio": 16000
    },
    "pepsi 1.75l": {
        "precio": 6000
    },
    "pepsi 3.125l": {
        "precio": 10000
    },
    "roscon arequipe": {
        "precio": 2000
    },
    "roscon guayaba": {
        "precio": 2000
    },
    "carne en bistec, cafe en leche": {
        "precio": 13000
    },
    "tolima": {
        "precio": 500
    },
    "jugo valle personal": {
        "precio": 2500
    },
    "huevo perico, arroz, pan": {
        "precio": 8500
    },
    "fruti\u00f1o limonada": {
        "precio": 1200
    },
    "almuerzo completo sopa, alverja, pollo desmechado con verduras": {
        "precio": 16000
    },
    "cafe negro": {
        "precio": 2000
    },
    "dedito": {
        "precio": 1000
    },
    "huevo revuelto, pan, cafe en leche (sin arroz)": {
        "precio": 7500
    },
    "huevo revuelto, arroz, pan": {
        "precio": 7000
    },
    "almuerzo completo, con arroz con pollo": {
        "precio": 16000
    },
    "huevo perico, con arroz y pancito": {
        "precio": 8000
    },
    "almuerzo completo sopa, lentejas, res asada": {
        "precio": 16000
    },
    "almuerzo completo sopa y arroz con pollo": {
        "precio": 16000
    },
    "bandeja frijol y chuleta cerdo": {
        "precio": 14000
    },
    "bandeja frijol y filete de pollo": {
        "precio": 14000
    },
    "almuerzo completo sopa con arroz con pollo": {
        "precio": 16000
    },
    "almuerzo completo sopa, lenteja y filete cerdo": {
        "precio": 16000
    },
    "almuerzo completo sopa lenteja y chuleta de cerdo": {
        "precio": 16000
    },
    "almuerzo completo sopa, lenteja, chuleta cerdo": {
        "precio": 16000
    },
    "almuerzo completo sopa, lenteja, res asada": {
        "precio": 16000
    },
    "almuerzo completo sopa, firjol, res asada": {
        "precio": 16000
    },
    "porcion frijol": {
        "precio": 6000
    },
    "cocacola 2.5l": {
        "precio": 10000
    },
    "cocacola 2.25l": {
        "precio": 10000
    },
    "mil hojas": {
        "precio": 5500
    },
    "chocorramo": {
        "precio": 3500
    },
    "colombiana 3l": {
        "precio": 10000
    },
    "jugo hit lulo": {
        "precio": 3500
    },
    "porcion huevo revuelto": {
        "precio": 4500
    },
    "pan salchicha": {
        "precio": 1700
    },
    "perico cebolla cabezona, poco arroz, cafe en leche": {
        "precio": 10000
    },
    "porcion huevo perico": {
        "precio": 5500
    },
    "porcion huevo perico, pan": {
        "precio": 6000
    },
    "jugo del valle personal": {
        "precio": 2500
    },
    "jugo hit naranja 1l": {
        "precio": 6000
    },
    "brazo de reina": {
        "precio": 2500
    },
    "porcion de torta chocolate": {
        "precio": 2500
    },
    "desayuno competo ranchero, cafe en leche": {
        "precio": 11500
    },
    "desayuno completo ranchero, cafe negro": {
        "precio": 11000
    },
    "bandeja spaguetis, filete pollo": {
        "precio": 14000
    },
    "milo frio": {
        "precio": 6000
    },
    "milo caliente": {
        "precio": 6000
    },
    "milo tibio": {
        "precio": 6000
    },
    "desayuno completo perico (sin cafe)": {
        "precio": 8500
    },
    "porcion caldo de costilla": {
        "precio": 6000
    },
    "yigurt melocoton baso": {
        "precio": 3000
    },
    "porciones pastel chocolate": {
        "precio": 2500
    },
    "nectar pera 1l": {
        "precio": 7000
    },
    "almuerzo completo frijolada, chorizo": {
        "precio": 16000
    },
    "almuerzo completo frijolada, chicharon": {
        "precio": 16000
    },
    "porcion de arroz": {
        "precio": 2500
    },
    "almuerzo completo consome, frijol, chuleta de pollo": {
        "precio": 16000
    },
    "bandeja frijol, chorizo": {
        "precio": 14000
    },
    "huevo ranchero, arroz, pan": {
        "precio": 9000
    },
    "pan 500": {
        "precio": 500
    },
    "jugo hit tropical personal": {
        "precio": 3500
    },
    "desayuno perico, 2 panes": {
        "precio": 8500
    },
    "desayuno revuelto, croassant": {
        "precio": 7700
    },
    "desayuno revuelto": {
        "precio": 6000
    },
    "desayuno revuelto completo": {
        "precio": 12900
    },
    "jugo en caja personal": {
        "precio": 2500
    },
    "almuerzo complejo ajiaco, arroz con pollo": {
        "precio": 16000
    },
    "bandeja chuleta de pollo, fijoles": {
        "precio": 14000
    },
    "bandeja arroz con pollo": {
        "precio": 14000
    },
    "desayuno revuelto, arroz": {
        "precio": 6500
    },
    "jugo naranja fruttsi": {
        "precio": 6000
    },
    "almuerzo completo sanchocho pescado, fijoles y pescado frito": {
        "precio": 16000
    },
    "bandeja frijol y res asada": {
        "precio": 14000
    },
    "almuerzo completo sancocho de pescado, spaghetti, pescado frito": {
        "precio": 16000
    },
    "colombiana 2.5l": {
        "precio": 8000
    },
    "leche en polvo": {
        "precio": 2000
    },
    "pony malta mini": {
        "precio": 2000
    },
    "yogurt de fresa": {
        "precio": 3000
    },
    "huevo tibio": {
        "precio": 1200
    },
    "avena en bolsa": {
        "precio": 3000
    },
    "gaseosa manzana personal": {
        "precio": 3500
    },
    "bandeja lenteja, frilete de pollo": {
        "precio": 14000
    },
    "almuerzo completo sanchocho, lenteja, filete de pollo": {
        "precio": 16000
    },
    "almuerzo completo sancocho, (sin principio), chuleta de cerdo": {
        "precio": 16000
    },
    "bandeja lenteja, filete cerdo": {
        "precio": 14000
    },
    "almuerzo completo cumplea\u00f1os sancocho, firjol, pollo a la plancha": {
        "precio": 22000
    },
    "almuerzo completo sancocho, lenteja, res asada": {
        "precio": 16000
    },
    "almuerzo completo sancocho, lenteja, carne en bistec": {
        "precio": 16000
    }
}
        with open(PROD_JSON, 'w', encoding='utf-8') as f:
            json.dump(datos_base_prod, f, indent=4)

    if not os.path.exists(CLI_JSON):
        with open(CLI_JSON, 'w', encoding='utf-8') as f:
            json.dump([
    "Hector Tmv",
    "Termovapor",
    "Sebastian Garcia (Prixma)",
    "Juan Otero (Armametal)",
    "Cristian (Tissue)",
    "Yuli Laso (Porton Negro)",
    "Edwin (Acrilan)",
    "Julian Rios (Tissue)",
    "Yoimer (Antiguo Control Papagayo)",
    "Luisa (Arqustik)",
    "Laura (Blockers Klhar)",
    "Julian Jimenes (Grupo Textil)",
    "Angie (Conaldesa)",
    "Patricia (3 Piso)",
    "Sara Grupo Textil",
    "Juan Carlos (Bodega 13)",
    "Manuel (Tmv)",
    "Laminas Y Cortes Industriales Sa Nit 900035068-6",
    "Guillermo (Galvanizado)",
    "Jeimmy Valendia (Contactemos 1)",
    "Harold (Aqustik)",
    "Adelina (Cueva Del Humo)",
    "Maria (Termovapor)",
    "Susana (Bodega12)",
    "Paola (Antiguo Control Papagayo)",
    "Gabriela (Forraje)",
    "Alejandro (Ingal Planta Fibra)",
    "Katherin (Antiguo Control Papagayo)",
    "Camilogutierrez (Motomart)",
    "Viviana (Fundimetal)",
    "Andrea (Tecnoempaques)",
    "Fausto (Bodega 9)",
    "Martha (Armametal)",
    "Walter (Laminas Y Cortes)",
    "David Silva (Tensoactivos)",
    "Vanessa (Tissu)",
    "Andera (Funales)",
    "Coste\u00f1o (Grupo Textil)",
    "Eduardo (Tintuvalle)",
    ".",
    "Felipe (Bodega 6)",
    "Hector (Fabripunto)",
    "Maria (Bronces)",
    "Daniela Cardona (Berna)",
    "Frank (Armametal)",
    "Omar (Ipr)",
    "Diana (Fundimetal)",
    "Sofia (Rycarnes)",
    "Jose (Rycarnes)",
    "Guarda (Tisuu)",
    "Uriel (Italcol - Cerca De Bomba Primax)",
    "Marisol (Ecoindustrial)",
    "Karen Dayana (Bodega 17 Donde El Paisa)",
    "Johana (Sia Logistica)",
    "Silquin Itda. Nit: 890325787",
    "Silquin Itda Nit 890325787",
    "Steven (Acrilan)",
    "Victor Roman (Termovapor)",
    "Wil Duque (Tensoactivos)",
    "Abraham (Termovapor)",
    "Fernanda Vargas (Grupotextil)",
    "Estefania Cortez (Textiles Y Confecciones Del Valle)",
    "Diana (El Paso Casa 169)",
    "Luis Ramos (Holcim)",
    "Yoselin (Intergrafic)",
    "Tito (Megatextiles)",
    "Luis (Tmv)",
    "Mayra (Funales)",
    "Duvan (Ingal Galvanizado)",
    "Carolina",
    "Carlos (Papagayo)",
    "Leidy (Moto Mart)",
    "Alexandra (Tissue)",
    "Angie Suarez (Grupo Textil)",
    "Sofia (Grupo Textil)",
    "Deysi (Intergrafic)",
    "Esteban (Jaramillo Mora)",
    "Macar",
    "Daniela (Ingal Fibra)",
    "Yulieth Cuartas (Contactemos 2)",
    "Marie Eco Equipos",
    "Monica (Jaramillo Mora)",
    "Jhonatan Zapata (Berna)",
    "Jilary (Do\u00f1a Lupe)",
    "Diana (Bodega 12)",
    "Fabio Rodriguez (Remorques Dial)",
    "Yeni (Armametal)",
    "Camilo (Ipr Porteria 2)",
    "Sandra Silba (Tisuue)",
    "Ynmer (Fabrpunto)",
    "Esteban (Grupo Textil)",
    "Monica Posso (Corema)",
    "Juan Carlos Diez (Bodega 13)",
    "Alejandra (Sermac)",
    "Alejandro Martinez",
    "Yohana (Silquin)",
    "Gloria Milena (Industrias Macar)",
    "Sara Hernandez (Intergrafic)",
    "Tensoactivos",
    "Tmv",
    "Radio (Termovapor)",
    "Mishel Logistica (Tissue)",
    "Daniela Olaya (Fadepal)",
    "Contactamos Equipos Sas 805027728",
    "Moffatt Nit 900152835-1",
    "Arturo (Tubolaminas)",
    "Tmi",
    "Moffatt",
    "Silquin Itda",
    "Gabriel (Bodega 12)",
    "Johana (Tintuvalle)",
    "Marisol (Antiguo Control Papagayo)",
    "Jaramillo Mora",
    "Germ\u00e1n (Cueva Del Humo)",
    "Isabel (Intergraphic)",
    "Porteros (Tissue)",
    "Julian (Motomart)",
    "Giovanna (Bloques Klahr)",
    "Ruben Cano (Jaramillo Mora)",
    "Natalia (Remolques Dial)",
    "Alvaro Campo (Diaco)",
    "Andres (Tissue)",
    "Ingal",
    "Rodrigo (Prixma)",
    "Jorge (Nuevo Control Papagayo)",
    "Hector (Tmv)",
    "Jeferson Toro (Armametal Principal)",
    "Andres Amado (Proaceros)",
    "Juan David Preciado (Tintuvalle)",
    "Angela (Integrafic)",
    "Yamileth (Fabripunto)",
    "Alejandra (Berna)",
    "Soexcol",
    "Daniela (Tissue)",
    "Soexco",
    "Oscar (Porton Azul)",
    "Lorena (Fabripunto)",
    "Cecilia (Silquin)",
    "Nestor (Bascula)",
    "Raul Cardenas (Bodega 1)",
    "Empresa Textiles Y Manofacturas Del Valle Sas"
], f, indent=4)

# ==========================================
# CLASE PRINCIPAL DE LA APLICACIÓN
# ==========================================
class AppFacturacion:
    def __init__(self, root):
        self.root = root
        self.root.title("Facturación - Panadería y Restaurante Mi Salsa")
        self.root.geometry("900x700")
        self.root.configure(bg="#f4f4f4")
        
        self.productos_db = self.cargar_json(PROD_JSON)
        self.clientes_db = self.cargar_json(CLI_JSON)
        
        self.factura_items = []
        self.domicilio_eliminado = False
        self.linea_a_item_idx = {} 

        self.construir_interfaz()
        self.actualizar_vista_factura()
        
        self.entry_cant.focus_set()

    def cargar_json(self, ruta):
        with open(ruta, 'r', encoding='utf-8') as f:
            return json.load(f)

    def guardar_json(self, ruta, datos):
        with open(ruta, 'w', encoding='utf-8') as f:
            json.dump(datos, f, indent=4)

    def construir_interfaz(self):
        frame_izq = tk.Frame(self.root, bg="#f4f4f4", padx=20, pady=20)
        frame_izq.pack(side="left", fill="both", expand=True)

        frame_der = tk.Frame(self.root, bg="white", padx=10, pady=10, relief="sunken", borderwidth=2)
        frame_der.pack(side="right", fill="both", expand=True, padx=20, pady=20)

        tk.Label(frame_izq, text="SISTEMA DE FACTURACIÓN", font=("Arial", 14, "bold"), bg="#f4f4f4").pack(pady=(0, 10))

        tk.Label(frame_izq, text="Cantidad (Enter si está vacío para cobrar):", bg="#f4f4f4").pack(anchor="w")
        self.entry_cant = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_cant.pack(fill="x", pady=5)
        self.entry_cant.bind("<Return>", self.on_cant_enter)

        tk.Label(frame_izq, text="Producto:", bg="#f4f4f4").pack(anchor="w", pady=(5,0))
        self.entry_prod = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_prod.pack(fill="x", pady=5)
        
        self.listbox_prod = tk.Listbox(frame_izq, height=4, font=("Arial", 11))
        self.listbox_prod.pack(fill="x")
        self.listbox_prod.pack_forget() 
        
        self.entry_prod.bind("<KeyRelease>", self.filtrar_productos)
        self.entry_prod.bind("<Down>", lambda e: self.listbox_prod.focus_set() if self.listbox_prod.winfo_ismapped() else None)
        self.listbox_prod.bind("<Return>", self.seleccionar_producto)
        self.entry_prod.bind("<Return>", self.on_prod_enter)

        tk.Label(frame_izq, text="Precio Total (Corregir si es necesario):", bg="#f4f4f4").pack(anchor="w", pady=(5,0))
        self.entry_precio = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_precio.pack(fill="x", pady=5)
        self.entry_precio.bind("<Return>", self.agregar_producto_a_factura)

        tk.Frame(frame_izq, height=2, bg="#ccc").pack(fill="x", pady=10)

        tk.Label(frame_izq, text="Forma de Pago (Efectivo/Nequi/Anotar):", bg="#f4f4f4").pack(anchor="w")
        self.entry_pago = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_pago.pack(fill="x", pady=5)
        self.entry_pago.bind("<KeyRelease>", self.toggle_pago)
        self.entry_pago.bind("<Up>", self.toggle_pago_flechas)
        self.entry_pago.bind("<Down>", self.toggle_pago_flechas)
        self.entry_pago.bind("<Return>", self.validar_pago) 

        tk.Label(frame_izq, text="Dinero Recibido (Dejar vacío si es exacto o Nequi):", bg="#f4f4f4").pack(anchor="w", pady=(5,0))
        self.entry_recibido = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_recibido.pack(fill="x", pady=5)
        self.entry_recibido.bind("<Return>", self.on_recibido_enter)
        self.entry_recibido.bind("<KeyRelease>", lambda e: self.actualizar_vista_factura())

        tk.Label(frame_izq, text="Cliente (Para NIT o CC agregar 'nit' o 'cc'):", bg="#f4f4f4").pack(anchor="w", pady=(5,0))
        self.entry_cliente = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_cliente.pack(fill="x", pady=5)
        
        self.listbox_cli = tk.Listbox(frame_izq, height=3, font=("Arial", 11))
        self.listbox_cli.pack(fill="x")
        self.listbox_cli.pack_forget()
        
        self.entry_cliente.bind("<KeyRelease>", self.filtrar_clientes)
        self.entry_cliente.bind("<Down>", lambda e: self.listbox_cli.focus_set() if self.listbox_cli.winfo_ismapped() else None)
        self.listbox_cli.bind("<Return>", self.seleccionar_cliente)
        self.entry_cliente.bind("<Return>", self.finalizar_factura)

        self.txt_factura = tk.Text(frame_der, font=("Courier", 10), state="disabled", bg="white", wrap="word")
        self.txt_factura.pack(fill="both", expand=True)
        self.txt_factura.bind("<Double-Button-1>", self.interactuar_factura_click)

    # --- LÓGICA DE EVENTOS ---
    def on_cant_enter(self, event):
        cant = self.entry_cant.get().strip()
        if cant == "":
            self.entry_pago.focus_set()
        else:
            self.entry_prod.focus_set()
        return "break"

    def filtrar_productos(self, event):
        if event.keysym in ["Down", "Up", "Return"]: return
        busqueda = self.entry_prod.get().lower()
        self.listbox_prod.delete(0, tk.END)
        if busqueda:
            coincidencias = sorted([p for p in self.productos_db.keys() if busqueda in p.lower()])
            if coincidencias:
                self.listbox_prod.pack(fill="x", before=self.entry_precio)
                for c in coincidencias:
                    self.listbox_prod.insert(tk.END, c)
            else:
                self.listbox_prod.pack_forget()
        else:
            self.listbox_prod.pack_forget()

    def seleccionar_producto(self, event):
        seleccion = self.listbox_prod.get(tk.ACTIVE)
        self.entry_prod.delete(0, tk.END)
        self.entry_prod.insert(0, seleccion)
        self.listbox_prod.pack_forget()
        self.on_prod_enter(None)
        return "break"

    def on_prod_enter(self, event):
        self.listbox_prod.pack_forget()
        producto = self.entry_prod.get().strip().lower()
        cant_str = self.entry_cant.get().strip()
        
        if not cant_str.isdigit():
            messagebox.showerror("Error", "La cantidad debe ser un número entero.")
            self.entry_cant.focus_set()
            return "break"
            
        cant = int(cant_str)
        precio_total = 0
        
        if es_bandeja(producto):
            precio_total = cant * 14000
        elif es_almuerzo(producto):
            precio_total = cant * 16000
        elif producto in self.productos_db:
            info = self.productos_db[producto]
            if "promocion" in info:
                promo = info["promocion"]
                cant_promo = promo["cantidad"]
                precio_promo = promo["precio_promo"]
                precio_unidad = info["precio"]
                
                paquetes = cant // cant_promo
                sueltos = cant % cant_promo
                precio_total = (paquetes * precio_promo) + (sueltos * precio_unidad)
            else:
                precio_total = cant * info["precio"]
        
        self.entry_precio.delete(0, tk.END)
        if precio_total > 0:
            self.entry_precio.insert(0, str(precio_total))
        
        self.entry_precio.focus_set()
        self.entry_precio.select_range(0, tk.END)
        return "break"

    def agregar_producto_a_factura(self, event):
        cant = self.entry_cant.get().strip()
        prod = self.entry_prod.get().strip().lower()
        precio_total_str = self.entry_precio.get().strip()
        
        if not cant or not prod or not precio_total_str:
            return "break"
            
        precio_total = int(precio_total_str)
        cant = int(cant)
        
        if prod not in self.productos_db:
            precio_unitario = precio_total // cant
            self.productos_db[prod] = {"precio": precio_unitario}
            self.guardar_json(PROD_JSON, self.productos_db)
            
        self.factura_items = [item for item in self.factura_items if item["prod"] != "domicilio"]
        self.factura_items.append({"cant": cant, "prod": prod, "precio": precio_total})
        
        hay_almuerzo = any(es_almuerzo(item["prod"]) or es_bandeja(item["prod"]) for item in self.factura_items)
        if not hay_almuerzo and not self.domicilio_eliminado and len(self.factura_items) > 0:
            self.factura_items.append({"cant": 1, "prod": "domicilio", "precio": 1000})

        self.entry_cant.delete(0, tk.END)
        self.entry_prod.delete(0, tk.END)
        self.entry_precio.delete(0, tk.END)
        self.actualizar_vista_factura()
        self.entry_cant.focus_set()
        return "break"

    def toggle_pago(self, event):
        if event.keysym in ["Return", "Up", "Down"]: return
        val = self.entry_pago.get().lower()
        if val == "e":
            self.entry_pago.delete(0, tk.END)
            self.entry_pago.insert(0, "Efectivo")
        elif val == "n":
            self.entry_pago.delete(0, tk.END)
            self.entry_pago.insert(0, "Nequi")
        elif val == "a":
            self.entry_pago.delete(0, tk.END)
            self.entry_pago.insert(0, "Anotar")

    def toggle_pago_flechas(self, event):
        actual = self.entry_pago.get().lower()
        self.entry_pago.delete(0, tk.END)
        if "efectivo" in actual:
            self.entry_pago.insert(0, "Nequi")
        elif "nequi" in actual:
            self.entry_pago.insert(0, "Anotar")
        else:
            self.entry_pago.insert(0, "Efectivo")

    def validar_pago(self, event):
        val = self.entry_pago.get().strip().lower()
        if val in ["e", "efectivo"]:
            self.entry_pago.delete(0, tk.END)
            self.entry_pago.insert(0, "Efectivo")
        elif val in ["n", "nequi"]:
            self.entry_pago.delete(0, tk.END)
            self.entry_pago.insert(0, "Nequi")
        elif val in ["a", "anotar"]:
            self.entry_pago.delete(0, tk.END)
            self.entry_pago.insert(0, "Anotar")
        else:
            messagebox.showwarning("Atención", "Escriba 'e' para Efectivo, 'n' para Nequi o 'a' para Anotar.")
            self.entry_pago.focus_set()
            return "break"
        
        self.entry_recibido.focus_set()
        self.actualizar_vista_factura()
        return "break"

    def on_recibido_enter(self, event):
        self.entry_cliente.focus_set()
        self.actualizar_vista_factura()
        return "break"

    def filtrar_clientes(self, event):
        self.actualizar_vista_factura()
        if event.keysym in ["Down", "Up", "Return"]: return
        busqueda = self.entry_cliente.get().lower()
        self.listbox_cli.delete(0, tk.END)
        if busqueda:
            coincidencias = sorted([c for c in self.clientes_db if busqueda in c.lower()])
            if coincidencias:
                self.listbox_cli.pack(fill="x")
                for c in coincidencias:
                    self.listbox_cli.insert(tk.END, c)
            else:
                self.listbox_cli.pack_forget()
        else:
            self.listbox_cli.pack_forget()

    def seleccionar_cliente(self, event):
        seleccion = self.listbox_cli.get(tk.ACTIVE)
        self.entry_cliente.delete(0, tk.END)
        self.entry_cliente.insert(0, seleccion)
        self.listbox_cli.pack_forget()
        self.finalizar_factura(None)
        return "break"

    def actualizar_vista_factura(self, fecha_hora=None):
        self.txt_factura.config(state="normal")
        self.txt_factura.delete("1.0", tk.END)
        self.linea_a_item_idx = {} 

        if fecha_hora is None:
            fecha_hora = datetime.now()

        fecha_str = fecha_hora.strftime("%d/%m/%Y")
        hora_str = fecha_hora.strftime("%H:%M")

        # --- Lógica de Extracción de Cliente y NIT / CC ---
        cliente_input = self.entry_cliente.get().strip()
        nombre_cliente, nit_cc = extraer_cliente_y_nit_cc(cliente_input)

        metodo_pago_raw = self.entry_pago.get().strip() or "Efectivo"
        if metodo_pago_raw.lower() == "anotar":
            metodo_pago = "Efectivo"
        else:
            metodo_pago = metodo_pago_raw.title()

        ancho_total = 29
        def centrar(texto):
            return '\n'.join(linea.strip().center(ancho_total) for linea in texto.split('\n'))

        encabezado_negocio = """PANADERIA Y RESTAURANTE
MI SALSA
Nit: 1130598879
Dir: CALLE 1 # TV. 1-250
Cel: 3023942042"""

        encabezado = centrar(encabezado_negocio) + "\n\n\n"
        encabezado += f"FACTURA ELECTRONICA DE VENTA\n\n\n"
        encabezado += f"Cajero        : Miguel Angel O.\n"
        encabezado += f"Fecha         : {fecha_str} HORA: {hora_str}\n"
        encabezado += f"Forma de pago : {metodo_pago}\n"
        encabezado += f"Cliente       : {nombre_cliente}\n"
        encabezado += f"Nit/CC        : {nit_cc}\n"
        encabezado += "-" * ancho_total + "\n"
        
        # Títulos de las columnas alineados (Cant 5, Produc, Total derecha)
        titulos = "Can. Produc." + " " * (ancho_total - 12 - 5) + "Total"
        encabezado += f"{titulos}\n"

        self.txt_factura.insert(tk.END, encabezado)
        
        # Aplicamos la negrita y centrado visual al encabezado de la tienda
        self.txt_factura.tag_add("bold_center", "1.0", "9.0")
        self.txt_factura.tag_configure("bold_center", font=("Courier", 10, "bold"), justify="center")

        texto_final = encabezado
        suma = 0
        suma_items = 0
        
        # --- Lógica de Productos y Columnas (5 espacios para cantidad) ---
        if not self.factura_items:
            vacio_str = "1    Producto Ejemplo                $0\n"
            self.txt_factura.insert(tk.END, vacio_str)
            texto_final += vacio_str
        else:
            for idx, item in enumerate(self.factura_items):
                suma += item["precio"]
                
                # --- CAMBIO AQUÍ: Solo suma la cantidad si el producto NO es el domicilio ---
                if item["prod"] != "domicilio":
                    suma_items += item["cant"]
                
                linea_inicio = int(self.txt_factura.index("end-1c").split('.')[0])
                
                texto_item = ""
                cant_str = str(item["cant"]).ljust(5) # 5 espacios exactos
                desc_completa = item["prod"].capitalize() # Solo la primera en mayúscula
                precio_str = f"${item['precio']}"
                
                ancho_precio = len(precio_str)
                ancho_prod = ancho_total - 5 - ancho_precio - 1
                if ancho_prod < 10: ancho_prod = 10
                
                palabras = desc_completa.split()
                lineas_prod = []
                linea_act = ""
                
                for p in palabras:
                    if len(linea_act) + len(p) + (1 if linea_act else 0) <= ancho_prod:
                        linea_act += (" " if linea_act else "") + p
                    else:
                        lineas_prod.append(linea_act)
                        linea_act = p
                if linea_act:
                    lineas_prod.append(linea_act)
                if not lineas_prod: lineas_prod = [""]

                espacios_medio = ancho_total - 5 - len(lineas_prod[0]) - ancho_precio
                texto_item += f"{cant_str}{lineas_prod[0]}{' ' * espacios_medio}{precio_str}\n"
                
                for linea in lineas_prod[1:]:
                    texto_item += f"     {linea}\n" # Mismos 5 espacios abajo

                self.txt_factura.insert(tk.END, texto_item)
                texto_final += texto_item
                
                linea_fin = int(self.txt_factura.index("end-1c").split('.')[0])
                for l in range(linea_inicio, linea_fin):
                    self.linea_a_item_idx[l] = idx

        # --- Lógica de Devuelta ---
        recibido_str = self.entry_recibido.get().strip()
        recibido = int(recibido_str) if recibido_str.isdigit() else suma
        if recibido < suma: recibido = suma # Evita devoluciones negativas por error
        devuelta = recibido - suma

        pie = "\n" + "-" * ancho_total + "\n"
        
        def alinear_derecha(etiqueta, valor_str):
            espacios = ancho_total - len(etiqueta) - len(valor_str)
            if espacios < 1: espacios = 1
            return f"{etiqueta}{' ' * espacios}{valor_str}\n"

        pie += alinear_derecha("T O T A L.......", f"${suma}")
        pie += alinear_derecha("TOTAL ITEMS.....", f"{suma_items}\n\n")
        pie += centrar("----[ MEDIOS DE PAGO ]---") + "\n\n"
        
        pie += alinear_derecha(metodo_pago.upper(), f"${recibido}")
        pie += alinear_derecha("CAMBIO:", f"${devuelta}")
        pie += "\n\n"
        pie += centrar("Fabricante del software y") + "\n"
        pie += centrar("proveedor tecnológico:") + "\n\n"
        pie += centrar("- MAOAZA_king -")

        self.txt_factura.insert(tk.END, pie)
        texto_final += pie
        
        self.txt_factura.config(state="disabled")
        return texto_final

    def interactuar_factura_click(self, event):
        posicion = self.txt_factura.index(f"@{event.x},{event.y}")
        linea_clicada = int(posicion.split('.')[0])
        
        if linea_clicada not in self.linea_a_item_idx:
            return
            
        idx = self.linea_a_item_idx[linea_clicada]
        producto_seleccionado = self.factura_items[idx]
        
        if producto_seleccionado["prod"] == "domicilio":
            if messagebox.askyesno("Quitar Domicilio", "¿Deseas eliminar el cobro del domicilio de esta factura?"):
                self.domicilio_eliminado = True
                self.factura_items.pop(idx)
                self.actualizar_vista_factura()
                self.entry_cant.focus_set()
            return 
        
        respuesta = messagebox.askyesno(
            "Modificar Producto", 
            f"¿Deseas cargar '{producto_seleccionado['prod'].capitalize()}' para corregirlo o eliminarlo?"
        )
        
        if respuesta:
            self.entry_cant.delete(0, tk.END)
            self.entry_cant.insert(0, str(producto_seleccionado["cant"]))
            self.entry_prod.delete(0, tk.END)
            self.entry_prod.insert(0, producto_seleccionado["prod"])
            self.entry_precio.delete(0, tk.END)
            self.entry_precio.insert(0, str(producto_seleccionado["precio"]))
            
            self.factura_items.pop(idx)
            self.actualizar_vista_factura()
            self.entry_cant.focus_set()
            self.entry_cant.select_range(0, tk.END)

    def subir_archivos_github(self):
        if not os.path.exists(CONFIG_JSON): return

        try:
            with open(CONFIG_JSON, "r", encoding="utf-8") as f:
                config_data = json.load(f)
                TOKEN = config_data.get("github_token")
        except Exception:
            return

        if not TOKEN or TOKEN == "apidegithub" or TOKEN.strip() == "": return

        USUARIO = "MAOAZAking"
        REPO = "panaderia_y_restaurante_mi_salsa"
        archivos_a_subir = {"productos_y_precios.json": PROD_JSON, "clientes.json": CLI_JSON}

        headers = {
            "Authorization": f"Bearer {TOKEN}",
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "Python-App"
        }

        errores = []
        for nombre_github, ruta_local in archivos_a_subir.items():
            if not os.path.exists(ruta_local): continue
            
            url = f"https://api.github.com/repos/{USUARIO}/{REPO}/contents/{nombre_github}"
            
            try:
                with open(ruta_local, "rb") as f:
                    contenido_base64 = base64.b64encode(f.read()).decode("utf-8")

                req_get = urllib.request.Request(url, headers=headers)
                sha = None
                try:
                    with urllib.request.urlopen(req_get, timeout=5) as response:
                        data_github = json.loads(response.read().decode("utf-8"))
                        sha = data_github.get("sha")
                except urllib.error.HTTPError as e:
                    pass

                payload = {
                    "message": "Actualización automática de POS tras emitir factura",
                    "content": contenido_base64,
                    "branch": "main"
                }
                if sha: payload["sha"] = sha
                
                data_json = json.dumps(payload).encode("utf-8")
                req_put = urllib.request.Request(url, data=data_json, headers=headers, method="PUT")
                
                with urllib.request.urlopen(req_put, timeout=5) as response:
                    pass
            except Exception as e:
                errores.append(f"Fallo al subir {nombre_github}: {str(e)}")
                
        if errores:
            mensaje = "La factura se imprimió, pero falló GitHub.\n\nDetalles:\n" + "\n".join(errores)
            messagebox.showerror("Error de Subida", mensaje)

    def finalizar_factura(self, event):
        cliente_input = self.entry_cliente.get().strip()
        
        # --- Alerta de Nombre Vacío ---
        if not cliente_input:
            respuesta = messagebox.askyesno(
                "Falta Nombre", 
                "No ha ingresado el nombre del cliente.\n\n¿Desea regresar para poner el nombre?\n\n(Sí = Regresar, No = Continuar sin nombre)",
                default=messagebox.YES
            )
            if respuesta:
                self.entry_cliente.focus_set()
                return "break"
            else:
                self.entry_cliente.insert(0, "CONSUMIDOR FINAL")
                cliente_input = "CONSUMIDOR FINAL"

        # Extrae el nombre limpio del cliente (sin NIT/CC) y formatea el nombre completo para la DB
        nombre_limpio_db, nit_cc = extraer_cliente_y_nit_cc(cliente_input)
        cliente_para_db = formatear_cliente_para_db(cliente_input)

        if cliente_para_db and cliente_para_db != "CONSUMIDOR FINAL" and cliente_para_db not in self.clientes_db:
            self.clientes_db.append(cliente_para_db)
            self.guardar_json(CLI_JSON, self.clientes_db)

        # Genera el texto final leyendo la hora exacta
        ahora = datetime.now()
        texto_final = self.actualizar_vista_factura(ahora)

        # Obtener el total acumulado de la factura actual
        suma_total = sum(item["precio"] for item in self.factura_items)

        # Registrar la factura en el Excel general 'facturas.xlsx'
        registrar_factura_excel(nombre_limpio_db, suma_total, ahora)

        # Si el método de pago seleccionado fue 'Anotar', se registra en 'cuenta_por_cobrar.xlsx'
        metodo_pago_ingresado = self.entry_pago.get().strip().lower()
        if metodo_pago_ingresado in ["a", "anotar"]:
            registrar_cuenta_por_cobrar(nombre_limpio_db, suma_total)

        # --- Lógica de Carpetas por Fecha Actual ---
        meses = {1: 'enero', 2: 'febrero', 3: 'marzo', 4: 'abril', 5: 'mayo', 6: 'junio', 7: 'julio', 8: 'agosto', 9: 'septiembre', 10: 'octubre', 11: 'noviembre', 12: 'diciembre'}
        nombre_carpeta = f"{ahora.day}-{meses[ahora.month]}-{ahora.year}"
        ruta_carpeta = os.path.join(application_path, nombre_carpeta)
        
        if not os.path.exists(ruta_carpeta):
            os.makedirs(ruta_carpeta)

        # Prepara el nombre de guardado del archivo
        nombre_archivo_cliente = "CONSUMIDOR_FINAL"
        if nombre_limpio_db:
            nombre_archivo_cliente = nombre_limpio_db.replace(" ", "_")
            
        str_hora = ahora.strftime("%H_%M_%S")
        nombre_archivo = f"{nombre_archivo_cliente}_{str_hora}.txt"
        ruta_archivo = os.path.join(ruta_carpeta, nombre_archivo)

        # Guarda e Imprime
        try:
            with open(ruta_archivo, "w", encoding="utf-8") as f:
                f.write(texto_final)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la factura: {str(e)}")
            return "break"

        try:
            os.startfile(ruta_archivo, "print")
            messagebox.showinfo("Factura Lista", f"¡Comprobante generado y enviado a la impresora!")
        except Exception as e:
            messagebox.showinfo("Atención", f"Factura guardada.\n(No se pudo iniciar la impresora automáticamente)")

        self.subir_archivos_github()

        # Limpiar Todo
        self.factura_items = []
        self.domicilio_eliminado = False
        self.entry_cant.delete(0, tk.END)
        self.entry_prod.delete(0, tk.END)
        self.entry_precio.delete(0, tk.END)
        self.entry_pago.delete(0, tk.END)
        self.entry_recibido.delete(0, tk.END)
        self.entry_cliente.delete(0, tk.END)
        
        self.actualizar_vista_factura()
        self.entry_cant.focus_set()
        
        return "break"

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw() 
    
    descargar_archivos_github()
    crear_archivos_base_si_no_existen()
    
    root.deiconify()
    app = AppFacturacion(root)
    root.mainloop()