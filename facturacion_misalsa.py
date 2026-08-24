import tkinter as tk
from tkinter import messagebox
import json
import os
import sys
from datetime import datetime
import urllib.request
import base64

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

    # Uso de Bearer para compatibilidad total con Fine-grained tokens
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
        mensaje = "No se pudieron descargar los archivos de GitHub.\nSe usarán los datos locales.\n\nDetalles:\n" + "\n".join(errores) + "\n\n(Nota: Si es Error 403, revisa en GitHub que tu Token tenga permisos sobre 'Only select repositories' eligiendo este repositorio exacto)."
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
        self.root.geometry("850x600")
        self.root.configure(bg="#f4f4f4")
        
        self.productos_db = self.cargar_json(PROD_JSON)
        self.clientes_db = self.cargar_json(CLI_JSON)
        
        self.factura_items = []
        self.total_factura = 0
        self.domicilio_eliminado = False
        self.linea_a_item_idx = {} # Mapeo exacto para clics en líneas

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

        tk.Label(frame_izq, text="SISTEMA DE FACTURACIÓN", font=("Arial", 14, "bold"), bg="#f4f4f4").pack(pady=(0, 20))

        tk.Label(frame_izq, text="Cantidad (Enter si está vacío para finalizar):", bg="#f4f4f4").pack(anchor="w")
        self.entry_cant = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_cant.pack(fill="x", pady=5)
        self.entry_cant.bind("<Return>", self.on_cant_enter)

        tk.Label(frame_izq, text="Producto:", bg="#f4f4f4").pack(anchor="w", pady=(10,0))
        self.entry_prod = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_prod.pack(fill="x", pady=5)
        
        self.listbox_prod = tk.Listbox(frame_izq, height=4, font=("Arial", 11))
        self.listbox_prod.pack(fill="x")
        self.listbox_prod.pack_forget() 
        
        self.entry_prod.bind("<KeyRelease>", self.filtrar_productos)
        self.entry_prod.bind("<Down>", lambda e: self.listbox_prod.focus_set() if self.listbox_prod.winfo_ismapped() else None)
        self.listbox_prod.bind("<Return>", self.seleccionar_producto)
        self.entry_prod.bind("<Return>", self.on_prod_enter)

        tk.Label(frame_izq, text="Precio Total (Corregir si es necesario):", bg="#f4f4f4").pack(anchor="w", pady=(10,0))
        self.entry_precio = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_precio.pack(fill="x", pady=5)
        self.entry_precio.bind("<Return>", self.agregar_producto_a_factura)

        tk.Frame(frame_izq, height=2, bg="#ccc").pack(fill="x", pady=20)

        tk.Label(frame_izq, text="Forma de Pago (Efectivo/Nequi):", bg="#f4f4f4").pack(anchor="w", pady=(10,0))
        self.entry_pago = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_pago.pack(fill="x", pady=5)
        self.entry_pago.bind("<KeyRelease>", self.toggle_pago)
        self.entry_pago.bind("<Up>", self.toggle_pago_flechas)
        self.entry_pago.bind("<Down>", self.toggle_pago_flechas)
        self.entry_pago.bind("<Return>", self.validar_pago) 

        tk.Label(frame_izq, text="Cliente:", bg="#f4f4f4").pack(anchor="w", pady=(10,0))
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
        # Enlace general para doble clic gestionado de forma inteligente por líneas
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
            
        # 1. Quitamos temporalmente el domicilio si ya estaba agregado
        self.factura_items = [item for item in self.factura_items if item["prod"] != "domicilio"]
        
        # 2. Agregamos el nuevo producto
        self.factura_items.append({"cant": cant, "prod": prod, "precio": precio_total})
        
        # 3. Evaluamos si hay almuerzo o abreviaturas ("almuer", "amuer", "bande")
        hay_almuerzo = any(any(kw in item["prod"].lower() for kw in ["almuer", "amuer", "bande"]) for item in self.factura_items)
                
        # 4. Si no hay almuerzo y no fue eliminado a mano, agregamos el domicilio AL FINAL
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
            self.entry_cliente.focus_set()
        elif val in ["n", "nequi"]:
            self.entry_pago.delete(0, tk.END)
            self.entry_pago.insert(0, "Nequi")
            self.entry_cliente.focus_set()
        else:
            messagebox.showwarning("Atención", "Escriba 'e' para Efectivo o 'n' para Nequi.")
            self.entry_pago.focus_set()
        return "break"

    def filtrar_clientes(self, event):
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

    def actualizar_vista_factura(self, pago="", cliente="", fecha_hora=None):
        self.txt_factura.config(state="normal")
        self.txt_factura.delete("1.0", tk.END)
        self.linea_a_item_idx = {} # Reiniciamos el diccionario de líneas

        if fecha_hora is None:
            fecha_str = "dia/mes/año"
            hora_str = "horas:minutos"
        else:
            fecha_str = fecha_hora.strftime("%d/%m/%Y")
            hora_str = fecha_hora.strftime("%H:%M")

        encabezado = f"""PANADERIA Y RESTAURANTE
      MI SALSA
-------------------
EDWARD ARROYAVE
NIT:1130598879
FECHA:{fecha_str} HORA: {hora_str}
VENDEDOR: MIGUEL
-------------------\n"""
        
        self.txt_factura.insert(tk.END, encabezado)
        
        suma = 0
        texto_final = encabezado
        
        if not self.factura_items:
            vacio_str = "   Cant. Producto        $precio\n"
            self.txt_factura.insert(tk.END, vacio_str)
            texto_final += vacio_str
        else:
            for idx, item in enumerate(self.factura_items):
                suma += item["precio"]
                
                # Registramos en qué línea física del widget Text comienza este producto
                linea_inicio = int(self.txt_factura.index("end-1c").split('.')[0])
                
                texto_item = ""
                desc_producto = f'{item["cant"]} {item["prod"].title()}'
                precio_str = f"${item['precio']}"
                ancho_max_texto = 28 
                
                palabras = desc_producto.split()
                lineas_producto = []
                linea_actual = ""
                
                for palabra in palabras:
                    if len(linea_actual) + len(palabra) + (1 if linea_actual else 0) <= ancho_max_texto:
                        linea_actual += (" " if linea_actual else "") + palabra
                    else:
                        lineas_producto.append(linea_actual)
                        linea_actual = palabra
                if linea_actual:
                    lineas_producto.append(linea_actual)
                if not lineas_producto: lineas_producto = [desc_producto[:ancho_max_texto]]

                for i, linea in enumerate(lineas_producto):
                    if i == len(lineas_producto) - 1:
                        espacios = 39 - len(linea) - len(precio_str)
                        if espacios < 1: espacios = 1
                        texto_item += f"{linea}{' ' * espacios}{precio_str}\n"
                    else:
                        texto_item += f"{linea}\n"

                self.txt_factura.insert(tk.END, texto_item)
                texto_final += texto_item
                
                # Registramos la línea final del producto
                linea_fin = int(self.txt_factura.index("end-1c").split('.')[0])
                
                # Mapeamos todas las líneas físicas de este bloque al índice exacto del item
                for l in range(linea_inicio, linea_fin):
                    self.linea_a_item_idx[l] = idx

        pie = f"\n-------------------\n"
        suma_str = f'TOTAL:{" "*20}${suma}'
        pie += f'{suma_str}\n'
        pie += "-------------------\n\n"
        pie += f'FORMA PAGO: {pago if pago else ""}\n'
        pie += f'CLIENTE: {cliente if cliente else "Nombre"}\n'

        self.txt_factura.insert(tk.END, pie)
        texto_final += pie
        
        self.txt_factura.config(state="disabled")
        return texto_final

    def interactuar_factura_click(self, event):
        # Obtenemos exactamente la línea física donde se hizo doble clic
        posicion = self.txt_factura.index(f"@{event.x},{event.y}")
        linea_clicada = int(posicion.split('.')[0])
        
        # Verificamos si la línea pertenece a un producto válido
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
            f"¿Deseas cargar '{producto_seleccionado['prod'].title()}' para corregirlo o eliminarlo?"
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
        if not os.path.exists(CONFIG_JSON):
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
        archivos_a_subir = {
            "productos_y_precios.json": PROD_JSON,
            "clientes.json": CLI_JSON
        }

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
                    # Esto imprimirá el error exacto de GitHub en la consola de comandos
                    print(f"URL fallida: {url}")
                    print(f"Código de error HTTP: {e.code} - {e.reason}")
                    print(f"Respuesta de GitHub: {e.read().decode('utf-8')}")
                    errores.append(f"No se pudo descargar {nombre_github}: HTTP Error {e.code}")
            # Esto imprimirá el error exacto de GitHub en la consola de comandos
                    # except urllib.error.HTTPError as e:
                    #         if e.code != 404: raise e

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
            mensaje = "La factura se imprimió correctamente, pero no se pudo actualizar la base de datos en GitHub.\n\nDetalles:\n" + "\n".join(errores)
            messagebox.showerror("Error de Subida", mensaje)

    def finalizar_factura(self, event):
        cliente = self.entry_cliente.get().strip().title()
        pago = self.entry_pago.get().strip().title()
        
        if not cliente:
            messagebox.showerror("Error", "Ingrese el nombre del cliente.")
            return "break"

        if cliente not in self.clientes_db:
            self.clientes_db.append(cliente)
            self.guardar_json(CLI_JSON, self.clientes_db)

        ahora = datetime.now()
        texto_final = self.actualizar_vista_factura(pago, cliente, ahora)

        nombre_cliente_limpio = cliente.replace(" ", "_")
        str_fecha = ahora.strftime("%d_%m_%Y_%H_%M")
        nombre_archivo = f"{nombre_cliente_limpio}_{str_fecha}.txt"
        ruta_archivo = os.path.join(application_path, nombre_archivo)

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

        self.factura_items = []
        self.domicilio_eliminado = False
        self.entry_cant.delete(0, tk.END)
        self.entry_prod.delete(0, tk.END)
        self.entry_precio.delete(0, tk.END)
        self.entry_cliente.delete(0, tk.END)
        self.entry_pago.delete(0, tk.END)
        
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