import tkinter as tk
from tkinter import messagebox
import json
import os
import sys
from datetime import datetime

# ==========================================
# CONFIGURACIÓN DE RUTAS Y ARCHIVOS JSON
# ==========================================
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

PROD_JSON = os.path.join(application_path, "productos_y_precios.json")
CLI_JSON = os.path.join(application_path, "clientes.json")

if not os.path.exists(PROD_JSON):
    datos_base_prod = {
        "pan cacho": {"precio": 1700},
        "pan agridulce": {
            "precio": 700,
            "promocion": {"cantidad": 3, "precio_promo": 2000}
        }
    }
    with open(PROD_JSON, 'w', encoding='utf-8') as f:
        json.dump(datos_base_prod, f, indent=4)

if not os.path.exists(CLI_JSON):
    with open(CLI_JSON, 'w', encoding='utf-8') as f:
        json.dump(["Hector TMV", "Juan Perez"], f, indent=4)

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

        # --- CONTROLES IZQUIERDA ---
        tk.Label(frame_izq, text="SISTEMA DE FACTURACIÓN", font=("Arial", 14, "bold"), bg="#f4f4f4").pack(pady=(0, 20))

        # 1. Cantidad
        tk.Label(frame_izq, text="Cantidad (Enter si está vacío para finalizar):", bg="#f4f4f4").pack(anchor="w")
        self.entry_cant = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_cant.pack(fill="x", pady=5)
        self.entry_cant.bind("<Return>", self.on_cant_enter)

        # 2. Producto
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

        # 3. Precio
        tk.Label(frame_izq, text="Precio Total (Corregir si es necesario):", bg="#f4f4f4").pack(anchor="w", pady=(10,0))
        self.entry_precio = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_precio.pack(fill="x", pady=5)
        self.entry_precio.bind("<Return>", self.agregar_producto_a_factura)

        tk.Frame(frame_izq, height=2, bg="#ccc").pack(fill="x", pady=20)


        # 4. Forma de Pago
        tk.Label(frame_izq, text="Forma de Pago (Efectivo/Nequi):", bg="#f4f4f4").pack(anchor="w", pady=(10,0))
        self.entry_pago = tk.Entry(frame_izq, font=("Arial", 12))
        self.entry_pago.pack(fill="x", pady=5)
        self.entry_pago.bind("<KeyRelease>", self.toggle_pago)
        self.entry_pago.bind("<Up>", self.toggle_pago_flechas)
        self.entry_pago.bind("<Down>", self.toggle_pago_flechas)
        self.entry_pago.bind("<Return>", self.validar_pago) # <-- Validación estricta

        # 5. Cliente
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

        # --- VISTA DERECHA ---
        self.txt_factura = tk.Text(frame_der, font=("Courier", 10), state="disabled", bg="white")
        self.txt_factura.pack(fill="both", expand=True)

    # --- LÓGICA DE EVENTOS ---
    def on_cant_enter(self, event):
        cant = self.entry_cant.get().strip()
        if cant == "":
            self.entry_pago.focus_set()
        else:
            self.entry_prod.focus_set()

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

    def on_prod_enter(self, event):
        self.listbox_prod.pack_forget()
        producto = self.entry_prod.get().strip().lower()
        cant_str = self.entry_cant.get().strip()
        
        if not cant_str.isdigit():
            messagebox.showerror("Error", "La cantidad debe ser un número entero.")
            self.entry_cant.focus_set()
            return
            
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

    def agregar_producto_a_factura(self, event):
        cant = self.entry_cant.get().strip()
        prod = self.entry_prod.get().strip().lower()
        precio_total_str = self.entry_precio.get().strip()
        
        if not cant or not prod or not precio_total_str:
            return
            
        precio_total = int(precio_total_str)
        cant = int(cant)
        
        if prod not in self.productos_db:
            precio_unitario = precio_total // cant
            self.productos_db[prod] = {"precio": precio_unitario}
            self.guardar_json(PROD_JSON, self.productos_db)
            self.sincronizar_con_github()
            
        self.factura_items.append({"cant": cant, "prod": prod, "precio": precio_total})
        
        self.entry_cant.delete(0, tk.END)
        self.entry_prod.delete(0, tk.END)
        self.entry_precio.delete(0, tk.END)
        self.actualizar_vista_factura()
        self.entry_cant.focus_set()

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

    def actualizar_vista_factura(self, pago="", cliente="", fecha_hora=None):
        if fecha_hora is None:
            fecha_str = "dia/mes/año"
            hora_str = "horas:minutos"
        else:
            fecha_str = fecha_hora.strftime("%d/%m/%Y")
            hora_str = fecha_hora.strftime("%H:%M")

        texto = f"""        PANADERIA Y RESTAURANTE
                MI SALSA
---------------------------------------
EDWARD ARROYAVE
NIT:1130598879
FECHA:{fecha_str} HORA: {hora_str}
VENDEDOR: MIGUEL
---------------------------------------\n"""
        
        suma = 0
        hay_almuerzo = False
        
        # --- Lógica de Placeholders y Productos ---
        if not self.factura_items:
            # Aquí aparece el "placeholder" cuando no hay nada
            texto += "   Cant. Producto        $precio\n"
        else:
            # Si hay productos, los lista
            for item in self.factura_items:
                suma += item["precio"]
                # Convertimos a minúsculas y validamos ambas opciones
                if "almuerzo" in item["prod"].lower() or "bandeja" in item["prod"].lower():
                    hay_almuerzo = True
                
                linea_prod = f'{item["cant"]} {item["prod"].title()}'
                if len(linea_prod) > 28: linea_prod = linea_prod[:28]
                
                espacios = 39 - len(linea_prod) - len(str(item["precio"])) - 1
                texto += f'{linea_prod}{" "*espacios}${item["precio"]}\n'

            # --- Lógica automática de Domicilio ---
            # Si hay al menos un producto, y no hay almuerzo, se suma el domicilio
            if not hay_almuerzo and len(self.factura_items) > 0:
                texto += f'1 Domicilio{" "*23}$1000\n'
                suma += 1000

        texto += "\n---------------------------------------\n"
        
        suma_str = f'TOTAL:{" "*20}${suma}'
        texto += f'{suma_str}\n'
        texto += "---------------------------------------\n\n"
        texto += f'FORMA PAGO: {pago if pago else ""}\n'
        texto += f'CLIENTE: {cliente if cliente else "Nombre"}\n'

        self.txt_factura.config(state="normal")
        self.txt_factura.delete("1.0", tk.END)
        self.txt_factura.insert(tk.END, texto)
        self.txt_factura.config(state="disabled")

        return texto

    def finalizar_factura(self, event):
        cliente = self.entry_cliente.get().strip().title()
        pago = self.entry_pago.get().strip().title()
        
        if not cliente:
            messagebox.showerror("Error", "Ingrese el nombre del cliente.")
            return

        if cliente not in self.clientes_db:
            self.clientes_db.append(cliente)
            self.guardar_json(CLI_JSON, self.clientes_db)
            self.sincronizar_con_github()

        ahora = datetime.now()
        texto_final = self.actualizar_vista_factura(pago, cliente, ahora)

        nombre_cliente_limpio = cliente.replace(" ", "_")
        str_fecha = ahora.strftime("%d_%m_%Y_%H_%M")
        nombre_archivo = f"{nombre_cliente_limpio}_{str_fecha}.txt"
        ruta_archivo = os.path.join(application_path, nombre_archivo)

        with open(ruta_archivo, "w", encoding="utf-8") as f:
            f.write(texto_final)

        try:
            os.startfile(ruta_archivo, "print")
            # Mensaje que pausa el flujo, te confirma que ya se envió y deja listo para la próxima
            messagebox.showinfo("Factura Lista", f"¡Comprobante generado y enviado a la impresora!\n\nSe guardó como: {nombre_archivo}\n\nPresione Aceptar para limpiar y continuar.")
        except Exception as e:
            messagebox.showinfo("Atención", f"Factura guardada como {nombre_archivo}\n(No se pudo iniciar la impresora automáticamente)")

        # Limpiar todo para la próxima factura
        self.factura_items = []
        self.entry_cant.delete(0, tk.END)
        self.entry_prod.delete(0, tk.END)
        self.entry_precio.delete(0, tk.END)
        self.entry_cliente.delete(0, tk.END)
        self.entry_pago.delete(0, tk.END)
        
        # AQUÍ ESTÁ LA ORDEN CLAVE:
        self.actualizar_vista_factura()
        self.entry_cant.focus_set()
    
    def sincronizar_con_github(self):
        import urllib.request
        import base64
        import json

        ruta_config = os.path.join(application_path, "config.json")
        if not os.path.exists(ruta_config):
            print("Error: No se encontró el archivo config.json con el token de acceso.")
            return

        try:
            with open(ruta_config, "r", encoding="utf-8") as f:
                config_data = json.load(f)
                TOKEN = config_data.get("github_token")
        except Exception as e:
            print(f"Error al leer config.json: {e}")
            return

        if not TOKEN or TOKEN == "github_pat_AQUÍ_PONES_TU_TOKEN_REAL":
            print("Error: El token en config.json no es válido.")
            return

        USUARIO = "MAOAZAking"
        REPO = "panaderia_y_restaurante_mi_salsa"
        
        archivos_a_subir = {
            "productos_y_precios.json": PROD_JSON,
            "clientes.json": CLI_JSON
        }

        headers = {
            "Authorization": f"token {TOKEN}",
            "Accept": "application/vnd.github.v3+json",
            "User-Agent": "Python-App"
        }

        for nombre_github, ruta_local in archivos_a_subir.items():
            if not os.path.exists(ruta_local):
                continue

            # Corrección aplicada a la URL de la API de GitHub
            url = f"https://api.github.com/repos/{USUARIO}/{REPO}/contents/{nombre_github}"

            try:
                with open(ruta_local, "rb") as f:
                    contenido_base64 = base64.b64encode(f.read()).decode("utf-8")

                req_get = urllib.request.Request(url, headers=headers)
                sha = None
                try:
                    with urllib.request.urlopen(req_get) as response:
                        data_github = json.loads(response.read().decode("utf-8"))
                        sha = data_github.get("sha")
                except urllib.error.HTTPError as e:
                    if e.code != 404:
                        raise e

                payload = {
                    "message": "Actualización automática desde el sistema de facturación",
                    "content": contenido_base64,
                    "branch": "main"
                }
                if sha:
                    payload["sha"] = sha

                data_json = json.dumps(payload).encode("utf-8")
                req_put = urllib.request.Request(url, data=data_json, headers=headers, method="PUT")
                
                with urllib.request.urlopen(req_put) as response:
                    if response.status in [200, 201]:
                        print(f"Sincronizado con éxito: {nombre_github}")

            except Exception as e:
                print(f"Error sincronizando {nombre_github} con la API: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AppFacturacion(root)
    root.mainloop()