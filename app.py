

import sys
import os
import sqlite3
import uuid
import io
import shutil
import pandas as pd
import qrcode
import webview
import threading
import re
from datetime import datetime
from urllib.parse import quote
from functools import wraps
from io import BytesIO

from flask import (
    Flask, render_template, redirect, url_for,
    request, jsonify, send_file, send_from_directory, session, flash
)
from waitress import serve
from werkzeug.utils import secure_filename

from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer,
    Image, Table, TableStyle
)
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib import colors
from tkinter import Tk
from tkinter.filedialog import asksaveasfilename
from tkinter.filedialog import askopenfilename
# =========================================================
#  CONFIGURACIÓN DE RUTAS (UNIFICADA DEV/EXE)
# =========================================================
# =========================================================
#  CONFIGURACIÓN DE RUTAS (CORREGIDA PARA .EXE)
# =========================================================

def generar_inventario():
    print("USANDO DB:", DB_PATH)
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM DATOS_PAPELERIA", conn)
    conn.close()

    # Oculta ventana Tkinter
    Tk().withdraw()

    ruta = asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivos Excel", "*.xlsx")],
        title="Guardar inventario como"
    )

    if ruta:
        df.to_excel(ruta, index=False)
        try:
            os.startfile(ruta)  # Abre Excel automáticamente
        except Exception as e:
            print(f"Error al abrir el archivo: {e}")
        print(f"Inventario guardado en: {ruta}")
    else:
        print("Operación cancelada")

def importar_excel_desktop():
    Tk().withdraw()

    ruta_archivo = askopenfilename(
        title="Selecciona el archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )

    if not ruta_archivo:
        print("Operación cancelada")
        return

    df = pd.read_excel(ruta_archivo)
    conn = get_db_connection()
    cursor = conn.cursor()

    # Pregunta al usuario modo de carga
    modo = input("Modo de carga (sumar/reemplazar): ").strip().lower()
    if modo not in ["sumar", "reemplazar"]:
        modo = "sumar"

    for _, r in df.iterrows():
        codigo = str(r.get("codigo"))
        descripcion = r.get("descripcion")
        precio = r.get("precio")
        cantidad_excel = r.get("inventario")

        if pd.notna(codigo) and pd.notna(cantidad_excel):
            cantidad_int = int(cantidad_excel)
            precio_valor = float(precio) if pd.notna(precio) else 0

            existe = cursor.execute(
                "SELECT inventario FROM DATOS_PAPELERIA WHERE codigo = ?",
                (codigo,)
            ).fetchone()

            if existe:
                if modo == "sumar":
                    cursor.execute("""
                        UPDATE DATOS_PAPELERIA
                        SET inventario = inventario + ?, precio = ?
                        WHERE codigo = ?
                    """, (cantidad_int, precio_valor, codigo))
                else:
                    cursor.execute("""
                        UPDATE DATOS_PAPELERIA
                        SET inventario = ?, precio = ?
                        WHERE codigo = ?
                    """, (cantidad_int, precio_valor, codigo))
            else:
                cursor.execute("""
                    INSERT INTO DATOS_PAPELERIA
                    (codigo, descripcion, inventario, precio, imagen)
                    VALUES (?, ?, ?, ?, ?)
                """, (codigo, descripcion, cantidad_int, precio_valor, "default.jpg"))

    conn.commit()
    conn.close()
    print("Importación completada ")

def obtener_ruta_base():
    """ Determina la carpeta real donde reside el .exe o el script .py """
    if getattr(sys, 'frozen', False):
        # Si es un ejecutable, usamos la carpeta donde está el archivo .exe
        return os.path.dirname(sys.executable)
    # Si es un script .py, usamos la carpeta del archivo
    return os.path.dirname(os.path.abspath(__file__))

def resource_path(relative_path):
    """ Busca archivos que FUERON EMPAQUETADOS dentro del .exe (Templates) """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 1. Definimos la base
BASE_DIR = obtener_ruta_base()

# 2. Carpeta de Base de Datos (Fija en C:)
DB_FOLDER = r"C:\abarrotes_app"
DB_PATH = os.path.join(DB_FOLDER, "DATOS_PAPELERIA.db")

# 3. Carpetas de Archivos Externos (Deben estar junto al .exe para poder escribir en ellas)
# IMPORTANTE: Estas NO deben estar dentro de la carpeta temporal _MEIPASS
PRODUCTOS_FOLDER = os.path.join(BASE_DIR, "static", "uploads")
TICKETS_FOLDER = os.path.join(BASE_DIR, "tickets_generados")

# 4. Archivo por defecto (Este puede ser interno o externo, mejor externo para evitar errores)
RUTA_DEFAULT_JPG = os.path.join(BASE_DIR, "static", "productos", "default.jpg")

# 5. Crear directorios físicamente en el disco
# Usamos PRODUCTOS_FOLDER directamente para asegurar la creación de la ruta completa
for carpeta in [DB_FOLDER, PRODUCTOS_FOLDER, TICKETS_FOLDER]:
    if not os.path.exists(carpeta):
        os.makedirs(carpeta, exist_ok=True)
def cargar_excel_automatico():
    ruta_excel = os.path.join(BASE_DIR, "inventario.xlsx")

    if not os.path.exists(ruta_excel):
        print(" No hay archivo inventario.xlsx")
        return

    print(" Cargando inventario desde Excel...")

    try:
        df = pd.read_excel(ruta_excel)
    except Exception as e:
        print(f" Error leyendo Excel: {e}")
        return

    conn = get_db_connection()
    cursor = conn.cursor()

    for _, r in df.iterrows():
        codigo = str(r.get("codigo"))
        descripcion = r.get("descripcion")
        precio = r.get("precio")
        cantidad = r.get("inventario")

        if pd.notna(codigo) and pd.notna(cantidad):
            cantidad = int(cantidad)
            precio = float(precio) if pd.notna(precio) else 0

            existe = cursor.execute(
                "SELECT inventario FROM DATOS_PAPELERIA WHERE codigo = ?",
                (codigo,)
            ).fetchone()

            if existe:
                cursor.execute("""
                    UPDATE DATOS_PAPELERIA
                    SET inventario = ?, precio = ?
                    WHERE codigo = ?
                """, (cantidad, precio, codigo))
            else:
                cursor.execute("""
                    INSERT INTO DATOS_PAPELERIA
                    (codigo, descripcion, inventario, precio, imagen)
                    VALUES (?, ?, ?, ?, ?)
                """, (codigo, descripcion, cantidad, precio, "default.jpg"))

    conn.commit()
    conn.close()

    print(" Inventario cargado automáticamente")

# =========================================================
#  CONFIGURACIÓN DE FLASK
# =========================================================

app = Flask(__name__,
            # 'static' debe apuntar a la carpeta EXTERNA junto al .exe
            static_folder=os.path.join(BASE_DIR, 'static'), 
            # 'templates' debe apuntar a la carpeta INTERNA (dentro del .exe)
            template_folder=resource_path('templates'))

# Configuramos la ruta de carga para usarla en tus funciones de guardado
app.config['UPLOAD_FOLDER'] = PRODUCTOS_FOLDER

app.secret_key = "clave_secreta_la_PAPELERIA_2026"

# =========================================================
#  GESTIÓN DE BASE DE DATOS
# =========================================================
def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row

    return conn

def inicializar_base_de_datos():
    crear_tablas()
    actualizar_tablas()

    conn = get_db_connection()
    cursor = conn.cursor()

def actualizar_tablas():
    conn = get_db_connection()
    cursor = conn.cursor()

    try:
        cursor.execute("ALTER TABLE venta_detalle ADD COLUMN precio_unit REAL")
        print(" Columna precio_unit agregada")
    except Exception:
        pass

    conn.commit()
    conn.close()

def cargar_inventario_inicial():
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM DATOS_PAPELERIA")
    total = cursor.fetchone()[0]

    archivo_excel = os.path.join(BASE_DIR, "inventario_inicial.xlsx")

    if total == 0 and os.path.exists(archivo_excel):

        df = pd.read_excel(archivo_excel)
        df.columns = df.columns.str.strip().str.lower()

        contador = 0

        for _, r in df.iterrows():

            codigo = str(r["codigo"])
            descripcion = r["descripcion"]
            inventario = int(r["inventario"])
            precio = r["precio"]

            cursor.execute("""
                INSERT INTO DATOS_PAPELERIA
                (codigo, descripcion, inventario, precio)
                VALUES (?, ?, ?, ?)
            """, (codigo, descripcion, inventario, precio))

            contador += 1

        conn.commit()
        print(f" Inventario inicial cargado. Productos: {contador}")

    else:
        print(" No se cargó inventario inicial.")

    conn.close()
def crear_respaldo_db():
    db_path = DB_PATH
    carpeta_respaldos = os.path.join(BASE_DIR, "respaldos")

    if not os.path.exists(carpeta_respaldos):
        os.makedirs(carpeta_respaldos)

    if os.path.exists(db_path):
        fecha = datetime.now().strftime("%Y-%m-%d_%H-%M")
        respaldo = os.path.join(carpeta_respaldos, f"respaldo_{fecha}.db")
        shutil.copy2(db_path, respaldo)
        print(f" Respaldo creado: {respaldo}")

def crear_tablas():
    conn = get_db_connection()
    cursor = conn.cursor()

    # Tabla de Productos
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS DATOS_PAPELERIA (
            codigo TEXT PRIMARY KEY,
            descripcion TEXT,
            inventario INTEGER DEFAULT 0,
            precio REAL,
            imagen TEXT
        )
    """)

    # Configuración
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS configuracion (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre_negocio TEXT,
            direccion TEXT,
            telefono TEXT,
            logo TEXT,
            zonas TEXT
        )
    """)

    cursor.execute("""
        INSERT OR IGNORE INTO configuracion (id, nombre_negocio)
        VALUES (1, 'Mi Papelería')
    """)

    # Usuarios
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS ADMIN_USUARIOS (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE,
            password TEXT
        )
    """)

    cursor.execute("""
        INSERT OR IGNORE INTO ADMIN_USUARIOS (usuario, password)
        VALUES ('admin', '13579rhv')
    """)

    # Ventas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS VENTAS (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            total REAL,
            corte_id INTEGER
        )
    """)

    # Detalle ventas
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS venta_detalle (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            venta_id INTEGER,
            codigo TEXT,
            descripcion TEXT,
            cantidad INTEGER,
            precio_unit REAL,
            subtotal REAL
        )
    """)

    # Cortes
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cortes_caja (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            total REAL,
            tickets INTEGER
        )
    """)

    # Historial
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS HISTORIAL_RECARGAS (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT,
            cantidad INTEGER,
            usuario TEXT,
            motivo TEXT,
            fecha TEXT
        )
    """)

    #  VALIDAR SI EXISTE LA COLUMNA 'destacado'
    def columna_existe(conn, tabla, columna):
        cursor = conn.execute(f"PRAGMA table_info({tabla})")
        columnas = [row[1] for row in cursor.fetchall()]
        return columna in columnas

   
    conn.commit()
    conn.close()


# Llamar a inicialización al cargar el script
inicializar_base_de_datos()


# =========================================================
#  FUNCIONES DE APOYO
# =========================================================

def zona_requerida(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("zona"):
            flash("Por favor, selecciona una zona de entrega para continuar.", "warning")
            return redirect(url_for("zona"))
        return f(*args, **kwargs)
    return decorated_function

def parche_base_datos():
    conn = get_db_connection()
    try:
        conn.execute("ALTER TABLE VENTAS ADD COLUMN corte_id INTEGER")
        conn.commit()
    except sqlite3.OperationalError:
        pass
    finally:
        conn.close()

def crear_indices():
    conn = get_db_connection()
    try:
        conn.execute("CREATE INDEX IF NOT EXISTS idx_codigo_producto ON DATOS_PAPELERIA (codigo)")
        conn.commit()
    except:
        pass
    finally:
        conn.close()

def limpiar_fotos_huerfanas():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT imagen FROM DATOS_PAPELERIA WHERE imagen IS NOT NULL")
    filas = cursor.fetchall()
    
    nombres_en_db = set()
    for fila in filas:
        if fila['imagen']:
            partes = [img.strip() for img in fila['imagen'].split(',')]
            nombres_en_db.update(partes)
    conn.close()

    if not os.path.exists(PRODUCTOS_FOLDER):
        return 0

    archivos = os.listdir(PRODUCTOS_FOLDER)
    borrados = 0
    for archivo in archivos:
        if archivo in ['default.jpg', 'logo.png', 'default.png']:
            continue
        if archivo not in nombres_en_db:
            try:
                os.remove(os.path.join(PRODUCTOS_FOLDER, archivo))
                borrados += 1
            except:
                pass
    return borrados

# A partir de aquí siguen tus @app.route...
# =========================================================
#  DECORADORES
# =========================================================

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("admin_logueado"):
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return decorated_function

def zona_requerida(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("zona"):
            flash("Selecciona una zona", "warning")
            return redirect(url_for("zona"))
        return f(*args, **kwargs)
    return decorated_function
@app.context_processor
def inject_config():
    try:
        conn = get_db_connection()
        #  ESTA LÍNEA ES CLAVE: Convierte la fila en algo que el HTML entiende por nombre
        conn.row_factory = sqlite3.Row 
        
        config = conn.execute("SELECT * FROM configuracion WHERE id = 1").fetchone()
        conn.close()
        
        # Si config existe, lo pasamos tal cual; si no, un diccionario con valores por defecto
        if config:
            return {'config': config}
        else:
            return {'config': {'nombre_negocio': 'papeleria', 'logo': 'logo.png'}}
            
    except Exception as e:
        print(f"Error en inject_config: {e}")
        return {'config': {}}

# =========================================================
#  INICIO
# =========================================================
from flask import send_file
import os
import pandas as pd

@app.route('/subir_imagen', methods=['POST'])
def subir_imagen():
    codigos = request.form.getlist('codigo_imagen[]')
    files = request.files.getlist('imagenes[]')
    
    conn = get_db_connection()
    
    for i, file in enumerate(files):
        if file and file.filename != '':
            filename = secure_filename(file.filename)
            # Guardamos en la ruta del EXE (static/uploads)
            ruta_destino = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(ruta_destino)
            
            if i < len(codigos):
                codigo_p = codigos[i]
                
                # Buscamos qué fotos ya tiene el producto para NO borrarlas
                prod = conn.execute("SELECT imagen FROM DATOS_PAPELERIA WHERE codigo = ?", (codigo_p,)).fetchone()
                
                if prod and prod['imagen']:
                    # Si ya tiene fotos, sumamos la nueva con una coma
                    nuevas_fotos = f"{prod['imagen']},{filename}"
                else:
                    nuevas_fotos = filename
                
                conn.execute("UPDATE DATOS_PAPELERIA SET imagen = ? WHERE codigo = ?", (nuevas_fotos, codigo_p))
    
    conn.commit()
    conn.close()
    return redirect(url_for('admin_productos'))

@app.route("/")
def index():
    conn = get_db_connection()

    tablas = conn.execute("PRAGMA table_info(DATOS_PAPELERIA)").fetchall()
    
    
    conn = get_db_connection()

    productos = conn.execute("""
    SELECT codigo, descripcion, precio
    FROM DATOS_PAPELERIA
    WHERE destacado = 1
    LIMIT 6
    """).fetchall()

    # AGREGAR ESTO
    config = conn.execute("SELECT * FROM configuracion WHERE id=1").fetchone()

    conn.close()

    return render_template("index.html", productos=productos, config=config)


@app.route("/zona", methods=["GET", "POST"])
def zona():

    #  LIMPIAR SIEMPRE LA ZONA
    session.pop("zona", None)

    conn = get_db_connection()
    config = conn.execute("SELECT * FROM configuracion WHERE id = 1").fetchone()
    conn.close()

    #  LISTA LIMPIA DE ZONAS
    zonas = [z.strip() for z in config['zonas'].split(',')] if config and config['zonas'] else []

    #print("ZONAS CARGADAS:", zonas)

    if request.method == "POST":
        session["zona"] = request.form.get("zona")
        return redirect(url_for("catalogo"))
    return render_template("zona.html", zonas=zonas)
@app.route("/catalogo")
def catalogo():
    zona = session.get("zona")
    if not zona: return redirect(url_for("zona"))

    conn = get_db_connection()
    productos = [dict(row) for row in conn.execute("SELECT * FROM DATOS_PAPELERIA").fetchall()]
    conn.close()

    for p in productos:
        if p.get('imagen'):
            # Convertimos el texto "foto1.jpg,foto2.jpg" en una LISTA real
            lista_fotos = [f.strip() for f in p['imagen'].split(',') if f.strip()]
            # Creamos las rutas completas para cada foto
            p['fotos_lista'] = [f"uploads/{f}" for f in lista_fotos]
        else:
            # Si no hay nada, lista con la imagen por defecto
            p['fotos_lista'] = ["productos/default.jpg"]
            
    return render_template("catalogo.html", productos=productos)
@app.route("/buscar_rapido")
def buscar_rapido():
    q = request.args.get("q", "").lower()
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT codigo, descripcion, precio FROM DATOS_PAPELERIA WHERE LOWER(descripcion) LIKE ?", (f"%{q}%",))
    productos = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify([{"codigo": p["codigo"], "descripcion": p["descripcion"], "precio": f"{p['precio']:.2f}"} for p in productos])
@app.route('/agregar/<codigo>')
def agregar(codigo):
    # 1. Conectamos a la base de datos para revisar el inventario
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT inventario, descripcion FROM DATOS_PAPELERIA WHERE codigo = ?", (codigo,))
    producto = cursor.fetchone()
    conn.close()

    if producto:
        # 2. REGLA DE ORO: Si no hay stock, no dejamos agregar
        if producto['inventario'] <= 0:
            return jsonify(ok=False, mensaje=f"Sin existencias de {producto['descripcion']}")

        # 3. Si hay stock, procedemos con tu lógica original del carrito
        carrito = session.get('carrito', {})
        carrito[codigo] = carrito.get(codigo, 0) + 1
        session['carrito'] = carrito
        session.modified = True
        return jsonify(ok=True)
    
    return jsonify(ok=False, mensaje="Producto no encontrado")

# =========================================================
# 🛒 CARRITO
# =========================================================
@app.route('/cambiar_cantidad/<codigo>/<cambio>')
def cambiar_cantidad(codigo, cambio):
    carrito = session.get('carrito', {})

    try:
        cambio = int(cambio)

        if codigo in carrito:

            nueva_cantidad = carrito[codigo] + cambio

            #  SOLO VALIDAMOS SI VA A AUMENTAR
            if cambio > 0:
                conn = get_db_connection()
                cursor = conn.cursor()

                cursor.execute("SELECT inventario FROM DATOS_PAPELERIA WHERE codigo = ?", (codigo,))
                producto = cursor.fetchone()

                conn.close()

                if not producto:
                    return jsonify(ok=False, error="Producto no encontrado")

                inventario_disponible = producto['inventario']

                #  VALIDACIÓN CLAVE
                if nueva_cantidad > inventario_disponible:
                    return jsonify(ok=False, error="Sin inventario suficiente")

            #  SI PASA VALIDACIÓN
            if nueva_cantidad <= 0:
                del carrito[codigo]
            else:
                carrito[codigo] = nueva_cantidad

        session['carrito'] = carrito
        session.modified = True

        return jsonify(ok=True)

    except Exception as e:
        return jsonify(ok=False, error=str(e)), 400
@app.route('/vaciar_carrito')
def vaciar_carrito():
    session['carrito'] = {}
    session.modified = True
    return jsonify(ok=True)
@app.route('/carrito')
def ver_carrito():
    # Es vital que diga 'return' al principio
    return mostrar_resumen_carrito('carrito.html')

@app.route('/api/validar-cp', methods=['GET'])
def validar_cp():
    cp = request.args.get('cp', '').strip()
    if not cp:
        return jsonify({'valido': False, 'mensaje': 'Ingresa un código postal.'}), 400
    
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute('SELECT disponible FROM zonas_entrega WHERE codigo_postal = ?', (cp,))
    resultado = cursor.fetchone()
    conn.close()
    
    if resultado and resultado['disponible']:
        return jsonify({'valido': True, 'mensaje': '✅ Cobertura disponible en esta zona.'})
            
    return jsonify({'valido': False, 'mensaje': '❌ Lo sentimos, aún no realizamos entregas en esta zona.'})


@app.route('/confirmar_venta', methods=["GET", "POST"])
def confirmar_venta():

    if request.method == "POST":
        cp_raw = request.form.get('codigo_postal', '').strip()

        # Extraer únicamente la secuencia de 5 dígitos numéricos (ej. '70461' de 'macuil centro 70461')
        match = re.search(r'\d{5}', cp_raw)
        cp_cliente = match.group(0) if match else cp_raw

        # 1. Validación en el Servidor (Flask)
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT disponible FROM zonas_entrega WHERE TRIM(codigo_postal) = ?', (cp_cliente,))
        resultado = cursor.fetchone()
        conn.close()

        es_valido = False
        if resultado:
            es_valido = bool(resultado['disponible'] if isinstance(resultado, dict) or hasattr(resultado, 'keys') else resultado[0])

        if not es_valido:
            # Si el CP no es válido o no tiene cobertura, recargamos la página con la advertencia
            return mostrar_resumen_carrito('confirmar_venta.html', error_cp="No realizamos entregas en el código postal ingresado.")

        # 2. Guardar datos del cliente en la sesión
        session['nombre_cliente'] = request.form.get('nombre')
        session['telefono_cliente'] = request.form.get('telefono')
        session['direccion_cliente'] = request.form.get('direccion')
        session['codigo_postal_cliente'] = cp_cliente
        session['referencias_cliente'] = request.form.get('referencias')

        # 3. Redirigir a venta_exitosa
        return redirect(url_for('venta_exitosa'))

    # Si alguien entra por GET, mostrar el resumen
    return mostrar_resumen_carrito('confirmar_venta.html')
  
@app.route("/acerca")
def acerca():
    return render_template("acerca.html")

@app.route('/venta_exitosa')
def venta_exitosa():
    from urllib.parse import quote
    import os
    import uuid
    from datetime import datetime

    resumen = obtener_datos_carrito()
    total_venta = resumen['total']
    productos = resumen['productos']
    if not productos:
        flash("No hay productos en la compra", "warning")
        return redirect(url_for("catalogo"))
    # Generamos un ID único para el pedido
    id_pedido = str(uuid.uuid4())[:6].upper()

    # Crear lista de productos para WhatsApp
    lista_texto = ""
    for p in productos:
        lista_texto += f"- {p['cantidad']}x {p['descripcion']} (${p['precio_total_producto']:.2f})\n"

    # Guardar venta y descontar inventario en la base de datos
    conn = get_db_connection()
    cursor = conn.cursor()

    # Guardamos la venta
    fecha_iso = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    cursor.execute(
        "INSERT INTO VENTAS (fecha, total) VALUES (?, ?)",
        (fecha_iso, total_venta)
    )
    venta_id_generada = cursor.lastrowid

    # Guardamos detalle y actualizamos inventario
    for p in productos:
        cursor.execute("""
            INSERT INTO venta_detalle 
            (venta_id, codigo, descripcion, cantidad, precio_unit, subtotal)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            venta_id_generada,
            p["codigo"],
            p["descripcion"],
            p["cantidad"],
            p["precio"],
            p["cantidad"] * p["precio"]
        ))

        cursor.execute(
            "UPDATE DATOS_PAPELERIA SET inventario = inventario - ? WHERE codigo = ?",
            (p["cantidad"], p["codigo"])
        )

    conn.commit()
    conn.close()
 # -----------------------------
    # GENERAR PDF DEL TICKET
    # -----------------------------
    pdf_filename = generar_ticket_pdf(id_pedido, productos, total_venta)

    # -----------------------------
    # DATOS DEL CLIENTE
    # -----------------------------
    nombre = session.get('nombre_cliente', 'No especificado')
    telefono = session.get('telefono_cliente', 'No especificado')
    direccion = session.get('direccion_cliente', 'No especificado')

    # -----------------------------
    # MENSAJE PARA WHATSAPP
    # -----------------------------
    mensaje_completo = (
        f" *NUEVO PEDIDO: #{id_pedido}*\n\n"
        f" *Cliente:* {nombre}\n"
        f" *Teléfono:* {telefono}\n"
        f" *Dirección:* {direccion}\n\n"
        f"{lista_texto}\n"
        f"*TOTAL A PAGAR: ${total_venta:.2f}*\n\n"
        f" Pago al recibir la mercancía."
    )

    whatsapp_url = f"https://wa.me/529513928223?text={quote(mensaje_completo)}"

    # Limpiamos carrito
    session.pop('carrito', None)
    session.pop('zona', None)
    # MOSTRAR VENTA EXITOSA
    return render_template(
        'venta_exitosa.html',
        id_pedido=id_pedido,
        total=total_venta,
        pdf_filename=pdf_filename,
        whatsapp_url=whatsapp_url
    )
@app.route('/nuevo_cliente')
def nuevo_cliente():
    # Borramos la zona y el carrito para que el "candado" se cierre
    session.pop('carrito', None)
    session.pop('zona', None)
    session.pop('nombre_cliente', None)
    
    # Al redireccionar a 'index', el decorador @zona_requerida 
    # detectará que no hay zona y lo mandará directo a elegir una.
    return redirect(url_for("index"))
def obtener_datos_carrito():
    carrito = session.get('carrito', {})
    productos_en_carrito = []
    total = 0.0
    conn = get_db_connection()
    
    for codigo, cantidad in carrito.items():
        row = conn.execute("SELECT codigo, descripcion, precio FROM DATOS_PAPELERIA WHERE codigo = ?", (codigo,)).fetchone()
        if row:
            p = dict(row)
            # Limpiar precio por si tiene símbolos
            precio_val = float(str(p["precio"]).replace("$", "").replace(",", "").strip())
            item_total = precio_val * cantidad
            total += item_total
            productos_en_carrito.append({
                "codigo": p["codigo"], 
                "descripcion": p["descripcion"], 
                "precio": precio_val, 
                "cantidad": cantidad, 
                "precio_total_producto": item_total
            })
    conn.close()
    return {"productos": productos_en_carrito, "total": total}
def mostrar_resumen_carrito(template_name, **kwargs):
    conn = get_db_connection()
    config = conn.execute("SELECT * FROM configuracion LIMIT 1").fetchone()

    carrito_session = session.get("carrito", {})
    productos = []
    total = 0

    if carrito_session:
        claves = [str(k) for k in carrito_session.keys() if str(k).strip()]

        if claves:
            placeholders = ",".join(["?"] * len(claves))
            
            try:
                items = conn.execute(f"SELECT * FROM DATOS_PAPELERIA WHERE codigo IN ({placeholders})", claves).fetchall()
            except Exception:
                items = conn.execute(f"SELECT * FROM DATOS_PAPELERIA WHERE rowid IN ({placeholders})", claves).fetchall()

            for item in items:
                keys = item.keys() if hasattr(item, 'keys') else []
                
                prod_id = item["id"] if "id" in keys else (item["codigo"] if "codigo" in keys else item[0])
                prod_codigo = item["codigo"] if "codigo" in keys else str(prod_id)
                prod_desc = item["descripcion"] if "descripcion" in keys else ""
                prod_precio = item["precio"] if "precio" in keys else 0

                cantidad = carrito_session.get(str(prod_codigo)) or carrito_session.get(str(prod_id)) or 0

                if cantidad > 0:
                    subtotal = prod_precio * cantidad
                    total += subtotal
                    productos.append({
                        "id": prod_id,
                        "codigo": prod_codigo,
                        "descripcion": prod_desc,
                        "nombre": prod_desc,
                        "precio": prod_precio,
                        "cantidad": cantidad,
                        "subtotal": subtotal,
                        "precio_total_producto": subtotal  # 🔥 Clave que requiere confirmar_venta.html
                    })

    conn.close()

    return render_template(
        template_name,
        productos=productos,
        total=total,
        config=config,
        **kwargs
    )
# -----------------------------
# GENERAR TICKET PDF
# -----------------------------
def generar_ticket_pdf(id_pedido, productos, total):
    if not productos:
        return None

    conn = get_db_connection()
    config = conn.execute("SELECT * FROM configuracion WHERE id = 1").fetchone()
    conn.close()

    nombre_negocio = config['nombre_negocio'] if config and 'nombre_negocio' in config.keys() else "Mi Negocio"
    direccion = config['direccion'] if config and 'direccion' in config.keys() else ""
    telefono = config['telefono'] if config and 'telefono' in config.keys() else ""

    if not os.path.exists(TICKETS_FOLDER):
        os.makedirs(TICKETS_FOLDER, exist_ok=True)

    filename_pdf = f"ticket_{id_pedido}.pdf"
    filepath = os.path.join(TICKETS_FOLDER, filename_pdf)

    doc = SimpleDocTemplate(
        filepath,
        pagesize=(80 * mm, 200 * mm),
        rightMargin=3 * mm,
        leftMargin=3 * mm,
        topMargin=4 * mm,
        bottomMargin=4 * mm
    )

    styles = getSampleStyleSheet()
    style_center = styles['Normal'].clone('Center')
    style_center.alignment = 1
    style_center.fontSize = 8
    style_center.leading = 10

    style_header = styles['Heading1'].clone('Header')
    style_header.alignment = 1
    style_header.fontSize = 12
    style_header.leading = 14

    story = []
    story.append(Paragraph(f"<b>{nombre_negocio}</b>", style_header))
    if direccion:
        story.append(Paragraph(direccion, style_center))
    if telefono:
        story.append(Paragraph(f"Tel: {telefono}", style_center))

    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(f"<b>Ticket #:</b> {id_pedido}", style_center))
    story.append(Paragraph(f"<b>Fecha:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}", style_center))
    story.append(Spacer(1, 4 * mm))

    tabla_datos = [["Cant", "Descripción", "P.U.", "Subt."]]
    for p in productos:
        desc = p.get('descripcion', '')
        if len(desc) > 16:
            desc = desc[:14] + ".."
        tabla_datos.append([
            str(p.get('cantidad', 1)),
            desc,
            f"${p.get('precio', 0.0):.2f}",
            f"${p.get('precio_total_producto', p.get('subtotal', 0.0)):.2f}"
        ])

    tabla_datos.append(["", "", "TOTAL:", f"${total:.2f}"])

    tabla = Table(tabla_datos, colWidths=[10 * mm, 32 * mm, 16 * mm, 16 * mm])
    tabla.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('ALIGN', (2, 0), (-1, -1), 'RIGHT'),
        ('LINEBELOW', (0, 0), (-1, 0), 0.5, colors.black),
        ('LINEABOVE', (0, -1), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
    ]))

    story.append(tabla)
    story.append(Spacer(1, 6 * mm))
    story.append(Paragraph("¡Gracias por su compra!", style_center))

    try:
        doc.build(story)
        return filename_pdf
    except Exception as e:
        print(f"Error generando PDF: {e}")
        return None
@app.route('/ver_ticket/<path:filename>')
def ver_ticket(filename):
    import os
    nombre_limpio = os.path.basename(filename)

    return send_from_directory(
        TICKETS_FOLDER,
        nombre_limpio
    )

@app.route('/descargar_ticket/<path:filename>')
def descargar_ticket(filename):
    import os
    nombre_limpio = os.path.basename(filename)

    if not nombre_limpio.lower().endswith('.pdf'):
        return "Archivo no válido", 400

    return send_from_directory(
        TICKETS_FOLDER,
        nombre_limpio,
        as_attachment=True
    )

from flask import request, redirect, url_for, flash, render_template
from werkzeug.utils import secure_filename
import os

CONFIG_FOLDER = os.path.join('static', 'config')
if not os.path.exists(CONFIG_FOLDER):
    os.makedirs(CONFIG_FOLDER)
@app.route('/admin/configuracion', methods=["GET", "POST"])
@login_required
def configuracion_admin():
    conn = get_db_connection()
    cursor = conn.cursor()

    if request.method == "POST":

        # -------------------------
        # DATOS GENERALES
        # -------------------------
        nombre = request.form.get("nombre_negocio")
        direccion = request.form.get("direccion")
        telefono = request.form.get("telefono")
        mensaje = request.form.get("mensaje_ticket")

        cursor.execute("""
            UPDATE configuracion
            SET nombre_negocio=?, direccion=?, telefono=?, mensaje_ticket=?
            WHERE id=1
        """, (nombre, direccion, telefono, mensaje))


        # -------------------------
        # LOGO
        # -------------------------
        logo = request.files.get("logo")
        if logo and logo.filename != "":
            nombre_logo = secure_filename(logo.filename)
            ruta_logo = os.path.join(CONFIG_FOLDER, nombre_logo)
            logo.save(ruta_logo)

            cursor.execute("UPDATE configuracion SET logo=? WHERE id=1", (nombre_logo,))


        # -------------------------
        # BANNER
        # -------------------------
        banner = request.files.get("banner")
        if banner and banner.filename != "":
            nombre_banner = secure_filename(banner.filename)
            ruta_banner = os.path.join(CONFIG_FOLDER, nombre_banner)
            banner.save(ruta_banner)

            cursor.execute("UPDATE configuracion SET banner=? WHERE id=1", (nombre_banner,))


        # -------------------------
        # FONDO
        # -------------------------
        fondo = request.files.get("fondo")
        if fondo and fondo.filename != "":
            nombre_fondo = secure_filename(fondo.filename)
            ruta_fondo = os.path.join(CONFIG_FOLDER, nombre_fondo)
            fondo.save(ruta_fondo)

            cursor.execute("UPDATE configuracion SET fondo=? WHERE id=1", (nombre_fondo,))


        conn.commit()
        conn.close()

        flash(" Configuración actualizada correctamente", "success")
        return redirect(url_for("configuracion_admin"))

    # -------------------------
    # GET: CARGAR DATOS
    # -------------------------
    config = cursor.execute("SELECT * FROM configuracion WHERE id = 1").fetchone()
    conn.close()

    return render_template("configuracion.html", config=config)


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        usuario = request.form.get("usuario")
        password = request.form.get("password")

        conn = get_db_connection()
        admin = conn.execute(
            "SELECT * FROM ADMIN_USUARIOS WHERE usuario = ? AND password = ?",
            (usuario, password)
        ).fetchone()
        conn.close()

        if admin:
            session["admin_logueado"] = True
            session["admin_usuario"] = usuario
            return redirect("/admin")
        else:
            flash("Error de acceso")

    return render_template("admin_login.html")
# -----------------------------
# RUTA   ADMIN
# ----------------------------
# Esta es una función totalmente aparte
@app.route("/admin/productos_mas_vendidos")
@login_required
def productos_mas_vendidos():
    try:
        conn = get_db_connection()
        conn.row_factory = sqlite3.Row
        
        # AQUÍ ESTÁ EL CIERRE QUE FALTA: """,
        query = """
        SELECT 
            descripcion, 
            SUM(cantidad) as total_vendido
        FROM venta_detalle
        GROUP BY descripcion
        ORDER BY total_vendido DESC
        LIMIT 10
        """
        
        productos = conn.execute(query).fetchall()
        conn.close()
        
        return render_template("productos_mas_vendidos.html", productos=productos)
        
    except Exception as e:
        print(f"Error en productos_mas_vendidos: {e}")
        return "Error al cargar el reporte", 500

# ASEGÚRATE DE QUE ESTO SEA LO ÚLTIMO QUE APARECE SI TU ARCHIVO TERMINA AQUÍ
@app.route('/admin/guardar_producto', methods=['POST'])
def guardar_producto():
    codigo = request.form.get('codigo')
    descripcion = request.form.get('descripcion')
    precio = request.form.get('precio')
    
    # Manejo de imágenes
    imagenes_nombres = []
    if 'imagenes[]' in request.files:
        files = request.files.getlist('imagenes[]')
        for file in files:
            if file and file.filename != '':
                # Usamos secure_filename y la ruta que ya configuramos arriba
                filename = secure_filename(file.filename)
                
                # REGLA DE ORO: Usar app.config['UPLOAD_FOLDER']
                ruta_final = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(ruta_final)
                imagenes_nombres.append(filename)

    # Aquí iría tu lógica de SQL para insertar en la DB...
    # imagen_db = ",".join(imagenes_nombres)
    return redirect(url_for('admin_productos'))
@app.route('/admin/subir_mas_imagenes', methods=['POST'])
def subir_mas_imagenes():
    codigos = request.form.getlist('codigo_imagen[]')
    files_list = request.files.getlist('imagenes[]')

    for i, file in enumerate(files_list):
        if file and file.filename != '':
            filename = secure_filename(file.filename)
            # Guardar en la carpeta correcta
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            
            # Lógica para actualizar el string en la base de datos...
            
    return redirect(url_for('admin_productos'))
@app.route("/admin/logout")
def admin_logout():
    session.pop("admin_logueado", None)
    session.pop("admin_usuario", None)
    return redirect(url_for("admin_login"))
@app.route('/guardar_config', methods=['POST'])
@app.route('/guardar_config', methods=['POST'])
def guardar_config():
    if not session.get('admin_logueado'):
        return redirect('/admin/login')

    # Obtenemos los datos del formulario
    nombre = request.form.get('nombre')
    direccion = request.form.get('direccion')
    telefono = request.form.get('telefono')
    zonas = request.form.get('zonas')
    archivo_logo = request.files.get('logo')

    try:
        conn = get_db_connection()
        # Aseguramos que la conexión use Row para poder leer config['logo']
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # 1. Obtener el logo actual por si no se sube uno nuevo
        config_actual = cursor.execute("SELECT logo FROM configuracion WHERE id = 1").fetchone()
        nombre_logo = config_actual['logo'] if config_actual else 'logo.png'

        # 2. Manejo del nuevo logo (si el usuario subió uno)
        if archivo_logo and archivo_logo.filename != '':
            folder_path = os.path.join('static', 'uploads')
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            
            filename = secure_filename(archivo_logo.filename)
            ruta_completa = os.path.join(folder_path, filename)
            archivo_logo.save(ruta_completa)

            # Guardamos la ruta que usará el HTML
            nombre_logo = f"uploads/{filename}"

        # 3. Actualizar la base de datos
        cursor.execute("""
            UPDATE configuracion
            SET nombre_negocio = ?, direccion = ?, telefono = ?, logo = ?, zonas = ?
            WHERE id = 1
        """, (nombre, direccion, telefono, nombre_logo, zonas))

        conn.commit()
        conn.close()
        
        # 4. Mensaje de éxito
        flash(" ¡Configuración actualizada con éxito!", "success")
        
        #  CAMBIO IMPORTANTE: Redirigir a la misma página de configuración
        # para ver los cambios aplicados inmediatamente.
        return redirect(url_for('configuracion_admin'))

    except Exception as e:
        print(f" Error al guardar configuración: {e}")
        flash(f"Error al guardar: {e}", "error")
        return redirect(url_for('configuracion_admin'))
@app.route("/admin")
@login_required
def admin_dashboard():
    conn = get_db_connection()

    ventas = conn.execute("""
        SELECT id, fecha, total
        FROM ventas
        ORDER BY fecha DESC
        LIMIT 10
    """).fetchall()

    ventas_hoy = conn.execute("""
        SELECT SUM(total) as total
        FROM ventas
        WHERE date(fecha) = date('now', 'localtime')
    """).fetchone()["total"] or 0

    ventas_mes = conn.execute("""
        SELECT SUM(total) as total
        FROM ventas
        WHERE strftime('%Y-%m', fecha) = strftime('%Y-%m', 'now', 'localtime')
    """).fetchone()["total"] or 0

    total_productos = conn.execute(
        "SELECT COUNT(*) as total FROM DATOS_PAPELERIA"
    ).fetchone()["total"]

    bajo_stock = conn.execute("""
        SELECT * FROM DATOS_PAPELERIA WHERE inventario <= 3
    """).fetchall()

    mas_vendidos = conn.execute("""
        SELECT descripcion, SUM(cantidad) as total_vendido
        FROM venta_detalle
        GROUP BY descripcion
        ORDER BY total_vendido DESC
        LIMIT 5
    """).fetchall()

    total_ingresos = conn.execute(
        "SELECT SUM(total) as total FROM ventas"
    ).fetchone()["total"] or 0

    #   ESTA LÍNEA FALTABA
    config = conn.execute(
        "SELECT * FROM configuracion WHERE id = 1"
    ).fetchone()

    conn.close()

    return render_template(
        "admin_dashboard.html",
        ventas=ventas,
        ventas_hoy=ventas_hoy,
        ventas_mes=ventas_mes,
        total_productos=total_productos,
        bajo_stock=bajo_stock,
        mas_vendidos=mas_vendidos,
        total_ingresos=total_ingresos,
        config=config  #   IMPORTANTE
    )






# -----------------------------
# EXPORTAR / IMPORTAR EXCEL
# -----------------------------
import io
from flask import send_file

#  IMPORTAR EXCEL (solo UNA vez)

@app.route("/admin/importar_excel", methods=["POST"])
@login_required
def importar_excel():
    archivo = request.files.get("archivo_excel")
    modo = request.form.get("modo_carga", "sumar")

    if not archivo or not archivo.filename:
        flash("Selecciona un archivo Excel.", "warning")
        return redirect(url_for("admin_productos"))

    if modo not in ["sumar", "reemplazar"]:
        flash("Modo de carga no válido.", "danger")
        return redirect(url_for("admin_productos"))

    try:
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.strip().str.lower()

        columnas = {"codigo", "descripcion", "precio", "inventario"}
        if not columnas.issubset(df.columns):
            flash("El Excel no contiene todas las columnas requeridas.", "danger")
            return redirect(url_for("admin_productos"))

        productos = {}

        for _, r in df.iterrows():
            if pd.isna(r["codigo"]) or pd.isna(r["inventario"]):
                continue

            codigo = str(r["codigo"]).strip()
            if codigo.endswith(".0") and codigo[:-2].isdigit():
                codigo = codigo[:-2]

            if not codigo:
                continue

            descripcion = "" if pd.isna(r["descripcion"]) else str(r["descripcion"]).strip()
            precio = 0 if pd.isna(r["precio"]) else float(r["precio"])
            cantidad = int(r["inventario"])

            if cantidad < 0 or precio < 0:
                raise ValueError(f"Cantidad o precio inválido en {codigo}")

            productos[codigo] = (descripcion, precio, cantidad)

        if not productos:
            flash("El Excel no contiene artículos válidos. No se modificó el inventario.", "warning")
            return redirect(url_for("admin_productos"))

        conn = get_db_connection()
        cursor = conn.cursor()

        try:
            # En modo reemplazar, quitar del catálogo
            # los artículos que no aparecen en el Excel.
            if modo == "reemplazar":
                actuales = cursor.execute(
                    "SELECT codigo FROM DATOS_PAPELERIA"
                ).fetchall()

                codigos_excel = set(productos.keys())

                for fila in actuales:
                    if fila["codigo"] not in codigos_excel:
                        cursor.execute(
                            "DELETE FROM DATOS_PAPELERIA WHERE codigo = ?",
                            (fila["codigo"],)
                        )

            # Insertar o actualizar los artículos del Excel.
            for codigo, (descripcion, precio, cantidad) in productos.items():
                existe = cursor.execute(
                    "SELECT codigo FROM DATOS_PAPELERIA WHERE codigo = ?",
                    (codigo,)
                ).fetchone()

                if existe:
                    if modo == "sumar":
                        cursor.execute("""
                            UPDATE DATOS_PAPELERIA
                            SET inventario = inventario + ?,
                                precio = ?,
                                descripcion = ?
                            WHERE codigo = ?
                        """, (cantidad, precio, descripcion, codigo))
                    else:
                        cursor.execute("""
                            UPDATE DATOS_PAPELERIA
                            SET inventario = ?,
                                precio = ?,
                                descripcion = ?
                            WHERE codigo = ?
                        """, (cantidad, precio, descripcion, codigo))
                else:
                    cursor.execute("""
                        INSERT INTO DATOS_PAPELERIA
                        (codigo, descripcion, inventario, precio, imagen)
                        VALUES (?, ?, ?, ?, ?)
                    """, (codigo, descripcion, cantidad, precio, "default.jpg"))

            conn.commit()

        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

        flash("Inventario reemplazado correctamente." if modo == "reemplazar"
              else "Importación completada correctamente.", "success")

    except Exception as e:
        flash(f"Error al importar el Excel: {e}", "danger")

    return redirect(url_for("admin_productos"))

#  EXPORTAR EXCEL (LO QUE TE FALTABA)
import os
import sys
from flask import send_file

@app.route("/admin/exportar_excel")
@login_required
def exportar_excel():
    conn = get_db_connection()

    df = pd.read_sql_query(
        "SELECT codigo, descripcion, inventario, precio FROM DATOS_PAPELERIA",
        conn
    )

    conn.close()

    #  Detectar entorno (normal o .exe)
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    ruta_archivo = os.path.join(base_dir, "inventario.xlsx")

    # Guardar archivo físico
    df.to_excel(ruta_archivo, index=False)

    # Enviar archivo
    return send_file(ruta_archivo, as_attachment=True)

from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename
import os
from functools import wraps

# Carpeta donde se guardarán las imágenes
UPLOAD_FOLDER = app.config['UPLOAD_FOLDER']
@app.route("/admin/productos", methods=["GET", "POST"])
@login_required
def admin_productos():
    conn = get_db_connection()

    if request.method == "POST":
        codigos = request.form.getlist("codigos_productos[]")

        #  1. INVENTARIO + DESTACADOS
        for codigo in codigos:
            cantidad = request.form.get(f"inv_{codigo}", type=float)
            tipo = request.form.get(f"tipo_{codigo}")

            # INVENTARIO
            if cantidad and tipo:
                if tipo == "compra":
                    conn.execute(
                        "UPDATE DATOS_PAPELERIA SET inventario = inventario + ? WHERE codigo=?",
                        (cantidad, codigo)
                    )
                elif tipo == "salida":
                    conn.execute(
                        "UPDATE DATOS_PAPELERIA SET inventario = inventario - ? WHERE codigo=?",
                        (cantidad, codigo)
                    )

            #  DESTACADO (AQUÍ VA BIEN)
            destacado = 1 if request.form.get(f"destacado_{codigo}") else 0

            conn.execute("""
                UPDATE DATOS_PAPELERIA
                SET destacado = ?
                WHERE codigo = ?
            """, (destacado, codigo))

        conn.commit()

        #  2. IMÁGENES
        codigos_img = request.form.getlist("codigo_imagen[]")

        for codigo in codigos_img:
            archivos = request.files.getlist(f"imagenes_{codigo}[]")

            for archivo in archivos:
                if archivo and archivo.filename != "":
                    nombre_seguro = secure_filename(archivo.filename)

                    ruta_guardar = os.path.join(PRODUCTOS_FOLDER, nombre_seguro)

                    print("Guardando imagen en:", ruta_guardar)

                    archivo.save(ruta_guardar)

                    producto = conn.execute(
                        "SELECT imagen FROM DATOS_PAPELERIA WHERE codigo=?",
                        (codigo,)
                    ).fetchone()

                    imagenes_actuales = []
                    if producto["imagen"]:
                        imagenes_actuales = producto["imagen"].split(',')

                    imagenes_actuales.append(nombre_seguro)

                    conn.execute(
                        "UPDATE DATOS_PAPELERIA SET imagen=? WHERE codigo=?",
                        (",".join(imagenes_actuales), codigo)
                    )

        conn.commit()

        flash("Movimientos, destacados e imágenes guardados correctamente", "success")
        return redirect(url_for("admin_productos"))

    productos = conn.execute("SELECT * FROM DATOS_PAPELERIA ORDER BY descripcion").fetchall()
    conn.close()

    return render_template("admin_productos.html", productos=productos)

@app.route("/actualizar_precio/<codigo>", methods=["POST"])
@login_required
def actualizar_precio(codigo):
    nuevo_precio = request.form.get("nuevo_precio", type=float)
    if nuevo_precio is None:
        return {"ok": False, "error": "Precio no válido"}, 400

    conn = get_db_connection()
    conn.execute("UPDATE DATOS_PAPELERIA SET precio=? WHERE codigo=?", (nuevo_precio, codigo))
    conn.commit()
    conn.close()
    return {"ok": True}
# --- RUTA PARA AJUSTAR INVENTARIO DESDE BOTONES + / - ---
@app.route('/admin/producto/ajustar/<codigo>', methods=['POST'])
@login_required
def ajustar_inventario(codigo):
    cambio = int(request.form.get('cambio', 0))  # Puede ser positivo o negativo

    conn = get_db_connection()
    producto = conn.execute("SELECT inventario FROM DATOS_PAPELERIA WHERE codigo = ?", (codigo,)).fetchone()
    
    if not producto:
        conn.close()
        return {"ok": False, "error": "Producto no encontrado"}, 404

    nuevo_inventario = producto['inventario'] + cambio
    if nuevo_inventario < 0:
        nuevo_inventario = 0  # Nunca permitimos inventario negativo

    conn.execute("UPDATE DATOS_PAPELERIA SET inventario = ? WHERE codigo = ?", (nuevo_inventario, codigo))
    conn.commit()
    conn.close()
    
    return {"ok": True, "nuevo_inventario": nuevo_inventario}
# -----------------------------
# CORTE DE CAJA DIARIO
# -----------------------------
@app.route("/corte_caja")
@login_required
def corte_caja():
    conn = get_db_connection()

    ventas = conn.execute("""
        SELECT id, fecha, total
        FROM VENTAS
        WHERE date(fecha) = date('now','localtime')
    """).fetchall()

    total = conn.execute("""
        SELECT COALESCE(SUM(total),0)
        FROM VENTAS
        WHERE date(fecha) = date('now','localtime')
    """).fetchone()[0]

    conn.close()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("CORTE DEL DÍA", styles['Title']))
    elementos.append(Spacer(1,20))
    elementos.append(Paragraph(f"Total: ${total:.2f}", styles['Normal']))
    elementos.append(Spacer(1,20))

    data = [["ID","Fecha","Total"]]
    for v in ventas:
        data.append([v["id"], v["fecha"], f"${v['total']:.2f}"])

    tabla = Table(data)
    tabla.setStyle(TableStyle([("GRID",(0,0),(-1,-1),1,colors.black)]))

    elementos.append(tabla)
    doc.build(elementos)
    buffer.seek(0)

    return send_file(buffer, download_name="corte.pdf", mimetype="application/pdf")

# -----------------------------
# CIERRE DE CAJA
# -----------------------------
@app.route("/cerrar_caja")
@login_required
def cerrar_caja():
    conn = get_db_connection()
    cursor = conn.cursor()

    total = cursor.execute("SELECT COALESCE(SUM(total),0) FROM VENTAS").fetchone()[0]
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cursor.execute("""
        INSERT INTO cortes_caja (fecha, total, tickets)
        VALUES (?, ?, 0)
    """, (fecha, total))
    corte_id = cursor.lastrowid

    cursor.execute("""
        UPDATE VENTAS
        SET corte_id = ?
        WHERE corte_id IS NULL
    """, (corte_id,))



    conn.commit()
    conn.close()

    flash("Caja cerrada correctamente", "success")
    return redirect(url_for("admin_dashboard"))
@app.route("/corte_caja_rango")
@login_required
def corte_caja_rango():

    inicio = request.args.get("inicio")
    fin = request.args.get("fin")

    conn = get_db_connection()

    ventas = conn.execute("""
        SELECT id, fecha, total
        FROM VENTAS
        WHERE date(fecha) BETWEEN ? AND ?
    """,(inicio,fin)).fetchall()

    total = conn.execute("""
        SELECT COALESCE(SUM(total),0) as total
        FROM VENTAS
        WHERE corte_id IS NULL
        AND date(fecha) BETWEEN ? AND ?
    """,(inicio,fin)).fetchone()["total"]

    tickets = len(ventas)

    conn.close()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)

    styles = getSampleStyleSheet()
    elementos = []

    elementos.append(Paragraph("CORTE DE CAJA", styles['Title']))
    elementos.append(Spacer(1,20))

    elementos.append(Paragraph(f"Periodo: {inicio} a {fin}", styles['Heading2']))
    elementos.append(Spacer(1,20))

    elementos.append(Paragraph(f"Total vendido: ${total:.2f}", styles['Normal']))
    elementos.append(Paragraph(f"Tickets: {tickets}", styles['Normal']))

    elementos.append(Spacer(1,20))

    data = [["ID","Fecha","Total"]]

    for v in ventas:
        data.append([v["id"], v["fecha"], f"${v['total']:.2f}"])

    data.append(["","TOTAL",f"${total:.2f}"])

    tabla = Table(data)

    tabla.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.grey),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("GRID",(0,0),(-1,-1),1,colors.black)
    ]))

    elementos.append(tabla)

    doc.build(elementos)

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=False,
        download_name="corte_caja_rango.pdf",
        mimetype="application/pdf"
    )
# -----------------------------
# REPORTES
# -----------------------------
@app.route("/admin/reportes")
@login_required
def admin_reportes():
    conn = get_db_connection()
    cortes = conn.execute("SELECT * FROM cortes_caja ORDER BY fecha DESC").fetchall()
    conn.close()
    return render_template("reportes.html", cortes=cortes)

@app.route("/admin/detalle_corte/<int:corte_id>")
@login_required
def detalle_corte(corte_id):
    conn = get_db_connection()

    corte = conn.execute(
        "SELECT * FROM cortes_caja WHERE id = ?",
        (corte_id,)
    ).fetchone()

    # Se corrigió el cierre de las comillas triples aquí
    detalles = conn.execute("""
        SELECT vd.descripcion, SUM(vd.cantidad) as cant, SUM(vd.subtotal) as total
        FROM venta_detalle vd
        JOIN ventas v ON vd.venta_id = v.id
        WHERE v.corte_id = ?
        GROUP BY vd.descripcion
    """, (corte_id,)).fetchall()

    conn.close()

    return render_template(
        "detalle_corte.html",
        corte=corte,
        detalles=detalles
    )

# -----------------------------
# HISTORIAL
# -----------------------------
@app.route("/admin/historial")
@login_required
def admin_historial():
    conn = get_db_connection()
    datos = conn.execute("SELECT * FROM HISTORIAL_RECARGAS ORDER BY fecha DESC").fetchall()
    conn.close()
    return render_template("admin_historial.html", historial=datos)

def check_db():
    conn = get_db_connection()
    # Se asegura que el string multilínea esté bien cerrado
    conn.execute('''
        CREATE TABLE IF NOT EXISTS configuracion 
        (id INTEGER PRIMARY KEY, nombre_negocio TEXT, mensaje_ticket TEXT)
    ''')
    
    res = conn.execute("SELECT COUNT(*) FROM configuracion").fetchone()
    if res[0] == 0:
        conn.execute("INSERT INTO configuracion (id, nombre_negocio, mensaje_ticket) VALUES (1, 'Papelería', '¡Gracias por su compra!')")
    conn.commit()
    conn.close()

@app.route('/fotos_productos/<path:filename>')
def servir_foto_externa(filename):
    try:
        return send_from_directory(PRODUCTOS_FOLDER, filename)
    except FileNotFoundError:
        return "Archivo no encontrado", 404

@app.route("/admin/limpiar_imagenes")
@login_required
def ruta_limpiar_imagenes():
    borrados = limpiar_fotos_huerfanas()
    flash(f" Se eliminaron {borrados} imágenes sin uso", "success")
    return redirect(url_for("admin_productos"))

@app.route("/admin/eliminar_imagen/<codigo>/<img>")
@login_required
def eliminar_imagen(codigo, img):
    conn = get_db_connection()
    cursor = conn.cursor()

    producto = cursor.execute(
        "SELECT imagen FROM DATOS_PAPELERIA WHERE codigo = ?",
        (codigo,)
    ).fetchone()

    if producto and producto["imagen"]:
        imagenes = producto["imagen"].split(",")
        imagenes = [i for i in imagenes if i != img]
        nueva_lista = ",".join(imagenes)

        cursor.execute(
            "UPDATE DATOS_PAPELERIA SET imagen = ? WHERE codigo = ?",
            (nueva_lista, codigo)
        )
        conn.commit()

    conn.close()

    ruta = os.path.join(PRODUCTOS_FOLDER, img)
    if os.path.exists(ruta):
        try:
            os.remove(ruta)
        except:
            pass

    flash(" Imagen eliminada", "success")
    return redirect(url_for("admin_productos"))
def actualizar_tabla_configuracion():
    conn = get_db_connection()
    cursor = conn.cursor()

    # Intentar agregar columnas nuevas
    try:
        cursor.execute("ALTER TABLE configuracion ADD COLUMN banner TEXT")
    except:
        pass

    try:
        cursor.execute("ALTER TABLE configuracion ADD COLUMN fondo TEXT")
    except:
        pass

    conn.commit()
    conn.close()
# --- ADMINISTRACIÓN DE ZONAS DE ENTREGA ---

# --- ADMINISTRACIÓN DE ZONAS DE ENTREGA ---


# --- ADMINISTRACIÓN DE ZONAS DE ENTREGA ---

@app.route('/admin/zonas_entrega', methods=['GET', 'POST'])
def admin_zonas_entrega():
    if not session.get('admin_logueado'):
        return redirect('/admin/login')

    conn = get_db_connection()

    # Crear la tabla automáticamente si aún no existe en database.db
    conn.execute('''
        CREATE TABLE IF NOT EXISTS zonas_entrega (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_postal TEXT UNIQUE NOT NULL,
            disponible INTEGER DEFAULT 1
        )
    ''')
    conn.commit()

    if request.method == 'POST':
        nuevo_cp = request.form.get('codigo_postal', '').strip()
        if nuevo_cp and len(nuevo_cp) == 5:
            conn.execute('''
                INSERT INTO zonas_entrega (codigo_postal, disponible) 
                VALUES (?, 1)
                ON CONFLICT(codigo_postal) DO UPDATE SET disponible = 1
            ''', (nuevo_cp,))
            conn.commit()

    zonas = conn.execute('SELECT * FROM zonas_entrega ORDER BY codigo_postal ASC').fetchall()
    conn.close()

    return render_template('admin_zonas.html', zonas=zonas)


@app.route('/admin/zonas_entrega/cambiar_estado/<int:zona_id>')
def cambiar_estado_zona(zona_id):
    if not session.get('admin_logueado'):
        return redirect('/admin/login')

    conn = get_db_connection()
    conn.execute('''
        UPDATE zonas_entrega 
        SET disponible = CASE WHEN disponible = 1 THEN 0 ELSE 1 END 
        WHERE id = ?
    ''', (zona_id,))
    conn.commit()
    conn.close()

    return redirect(url_for('admin_zonas_entrega'))


# -----------------------------
# SERVIDOR
# -----------------------------
def iniciar_flask():
    app.run(debug=False, use_reloader=False)
MODO = "app"  #  CAMBIA AQUÍ: "web" o "app"

if __name__ == "__main__":

    conn = sqlite3.connect(DB_PATH)
def actualizar_tablas():
    conn = get_db_connection()
    cursor = conn.cursor()

    # Agregar columna 'precio_unit' si no existe
    try:
        cursor.execute("ALTER TABLE venta_detalle ADD COLUMN precio_unit REAL")
    except Exception:
        pass

    # Agregar columna 'destacado' si no existe
    try:
        cursor.execute("ALTER TABLE DATOS_PAPELERIA ADD COLUMN destacado INTEGER DEFAULT 0")
        print("Columna 'destacado' agregada correctamente.")
    except Exception:
        pass  # Si ya existe, ignora el error silenciosamente

    conn.commit()
    conn.close()

def iniciar_flask():
    app.run(host="127.0.0.1", port=5000, debug=False, use_reloader=False)

MODO = "app"  # "web" o "app"

if __name__ == "__main__":

    # 1. PREPARAR CARPETAS
    for carpeta in [DB_FOLDER, PRODUCTOS_FOLDER, TICKETS_FOLDER]:
        os.makedirs(carpeta, exist_ok=True)

    # 2. BASE DE DATOS
    inicializar_base_de_datos()
    crear_tablas()
    actualizar_tabla_configuracion()
    actualizar_tablas()
    crear_indices()
    crear_respaldo_db()
    cargar_inventario_inicial()
    parche_base_datos()
    cargar_excel_automatico()

    print("Sistema listo")

    if MODO == "app":
        print("Modo APP iniciado")

        t = threading.Thread(target=iniciar_flask)
        t.daemon = True
        t.start()

        webview.create_window("Papeleria App", "http://127.0.0.1:5000")
        webview.start()

    elif MODO == "web":
        print("Modo WEB iniciado en http://127.0.0.1:5001")
        app.run(host="0.0.0.0", port=5001)

    else:
        print("MODO no válido. Usa 'app' o 'web'")