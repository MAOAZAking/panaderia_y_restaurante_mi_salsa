import tkinter as tk
from tkinter import messagebox
import json
import os
import sys
from datetime import datetime
import urllib.request
import base64
import re

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
            "pan cacho": {"precio": 1700},
            "pan agridulce": {"precio": 700, "promocion": {"cantidad": 3, "precio_promo": 2000}},
            "almuerzo corriente": {"precio": 15000},
            "bandeja paisa": {"precio": 25000}
        }
        with open(PROD_JSON, 'w', encoding='utf-8') as f:
            json.dump(datos_base_prod, f, indent=4)

    if not os.path.exists(CLI_JSON):
        with open(CLI_JSON, 'w', encoding='utf-8') as f:
            json.dump(["Majo", "Hector TMV"], f, indent=4)

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

        tk.Label(frame_izq, text="Forma de Pago (Efectivo/Nequi):", bg="#f4f4f4").pack(anchor="w")
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

        tk.Label(frame_izq, text="Cliente (Para NIT agregar la palabra 'nit'):", bg="#f4f4f4").pack(anchor="w", pady=(5,0))
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
        
        if producto in self.productos_db:
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
        
        hay_almuerzo = any(any(kw in item["prod"].lower() for kw in ["almuer", "amuer", "bande"]) for item in self.factura_items)
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

    def toggle_pago_flechas(self, event):
        actual = self.entry_pago.get().lower()
        self.entry_pago.delete(0, tk.END)
        if "efectivo" in actual:
            self.entry_pago.insert(0, "Nequi")
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
        else:
            messagebox.showwarning("Atención", "Escriba 'e' para Efectivo o 'n' para Nequi.")
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

        # --- Lógica de Extracción de Cliente y NIT ---
        cliente_input = self.entry_cliente.get().strip()
        nit_cc = "222222222222"
        nombre_cliente = cliente_input

        if cliente_input:
            match = re.search(r'nit\s*([0-9\-]+)', cliente_input.lower())
            if match:
                nit_cc = match.group(1)
                nombre_cliente = re.sub(r'(?i)nit\s*[0-9\-]+', '', cliente_input).strip()
            
            if not nombre_cliente:
                nombre_cliente = "CONSUMIDOR FINAL"
            else:
                nombre_cliente = nombre_cliente.upper() if nombre_cliente == "CONSUMIDOR FINAL" else nombre_cliente.title()
        else:
            nombre_cliente = "Nombre"

        metodo_pago = self.entry_pago.get().strip().title() or "Efectivo"

        ancho_total = 40
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

        pie += alinear_derecha("T O T A L............", f"${suma}")
        pie += alinear_derecha("TOTAL ITEMS..........", f"{suma_items}\n\n")
        pie += centrar("-----------[ MEDIOS DE PAGO ]-----------") + "\n\n"
        
        pie += alinear_derecha(metodo_pago.upper(), f"${recibido}")
        pie += alinear_derecha("CAMBIO:", f"${devuelta}")
        pie += "\n\n"
        pie += centrar("Fabricante del software y proveedor") + "\n"
        pie += centrar("tecnológico:") + "\n\n"
        pie += centrar("- MAOAZA_king -") + "\n\n"

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

        # Guarda el cliente limpio de NIT en la base de datos local
        nombre_limpio_db = cliente_input
        match_db = re.search(r'(?i)nit\s*[0-9\-]+', cliente_input)
        if match_db:
            n = re.sub(r'(?i)nit\s*[0-9\-]+', '', cliente_input).strip()
            if n: nombre_limpio_db = n.title()
        else:
            nombre_limpio_db = cliente_input.title() if cliente_input != "CONSUMIDOR FINAL" else cliente_input

        if nombre_limpio_db and nombre_limpio_db != "CONSUMIDOR FINAL" and nombre_limpio_db not in self.clientes_db:
            self.clientes_db.append(nombre_limpio_db)
            self.guardar_json(CLI_JSON, self.clientes_db)

        # Genera el texto final leyendo la hora exacta
        ahora = datetime.now()
        texto_final = self.actualizar_vista_factura(ahora)

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