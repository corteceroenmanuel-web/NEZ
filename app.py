import os
import datetime
from io import BytesIO
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash 
import sqlite3 
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
# Clave secreta para encriptar las sesiones del navegador de forma segura
app.secret_key = "clave_secreta_rolay_negocio"

# CONFIGURACIÓN DE LA BASE DE DATOS (SQLite local)
base_dir = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(base_dir, 'negocio.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# ========================================================
# MODELOS DE LA BASE DE DATOS (ESTRUCTURA DE LAS TABLAS)
# ========================================================

class Producto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    stock = db.Column(db.Integer, nullable=False, default=0)
    precio_venta = db.Column(db.Float, nullable=False)
    costo_compra = db.Column(db.Float, nullable=True)
    proveedor_id = db.Column(db.Integer, db.ForeignKey('proveedor.id'), nullable=True)
    ventas_totales = db.Column(db.Integer, nullable=False, default=0)

class Cliente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    telefono = db.Column(db.String(20), nullable=True)
    deuda_total = db.Column(db.Float, default=0.0)
    deudas = db.relationship('Deuda', backref='cliente', lazy=True)

class Deuda(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    cliente_id = db.Column(db.Integer, db.ForeignKey('cliente.id'), nullable=False)
    tipo = db.Column(db.String(50), nullable=False) 
    monto_inicial = db.Column(db.Float, nullable=False)
    saldo_pendiente = db.Column(db.Float, nullable=False)
    descripcion = db.Column(db.Text, nullable=True) 
    fecha = db.Column(db.DateTime, default=db.func.now())
    estado = db.Column(db.String(20), default="Pendiente") 
    
    abonos = db.relationship('HistorialAbono', backref='deuda', lazy=True)

class HistorialAbono(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    deuda_id = db.Column(db.Integer, db.ForeignKey('deuda.id'), nullable=False)
    monto_usd = db.Column(db.Float, nullable=False)
    monto_original = db.Column(db.Float, nullable=False) # Lo que entregó en físico
    moneda_pago = db.Column(db.String(10), nullable=False) # "USD" o "BS"
    tasa_momento = db.Column(db.Float, nullable=False)
    fecha = db.Column(db.DateTime, default=db.func.now())

class Proveedor(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    telefono = db.Column(db.String(20), nullable=True)
    rubro = db.Column(db.String(100), nullable=True, default="General") 
    deuda_pendiente = db.Column(db.Float, nullable=True, default=0.0)
    cuentas_pagar = db.relationship('CuentaPorPagar', backref='proveedor', lazy=True)

class Ingreso(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    tipo = db.Column(db.String(50), nullable=False) # <--- Aquí se guarda si es "Venta de Producto" o "Monto Externo"
    monto = db.Column(db.Float, nullable=False)
    metodo_pago = db.Column(db.String(50), nullable=False)
    fecha = db.Column(db.String(10), nullable=False)
    cerrado = db.Column(db.Boolean, default=False)

class Gasto(db.Model):
    """ Registro financiero de egresos o costos operativos """
    id = db.Column(db.Integer, primary_key=True)
    tipo = db.Column(db.String(50), nullable=False)
    monto = db.Column(db.Float, nullable=False)
    descripcion = db.Column(db.String(200))
    metodo_pago = db.Column(db.String(20), default='bs') # 'bs' o 'usd_efectivo'
    fecha = db.Column(db.DateTime, default=datetime.datetime.utcnow)
    cerrado = db.Column(db.Boolean, default=False)

class CuentaPorPagar(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    
    # Enlace real con el proveedor
    proveedor_id = db.Column(db.Integer, db.ForeignKey('proveedor.id'), nullable=False)
    
    # Control Multimoneda
    moneda = db.Column(db.String(3), nullable=False, default='USD') # 'USD' o 'BS'
    monto_original = db.Column(db.Float, nullable=False)            # El monto tal cual se negoció
    tasa_factura = db.Column(db.Float, nullable=False, default=1.0) # La tasa de cambio de ese día
    
    # Monto estandarizado en dólares para tu contabilidad global
    monto = db.Column(db.Float, nullable=False)
    
    descripcion = db.Column(db.String(200), nullable=True)
    fecha_limite = db.Column(db.String(20), nullable=True)
    pagado = db.Column(db.Boolean, default=False)

class UsuarioAdmin(db.Model):
    """ Credenciales de seguridad de acceso y preguntas de respaldo para recuperación """
    id = db.Column(db.Integer, primary_key=True)
    password_hash = db.Column(db.String(200), nullable=False)
    pregunta1 = db.Column(db.String(150), nullable=False)
    respuesta1 = db.Column(db.String(150), nullable=False)
    pregunta2 = db.Column(db.String(150), nullable=False)
    respuesta2 = db.Column(db.String(150), nullable=False)
    pregunta3 = db.Column(db.String(150), nullable=False)
    respuesta3 = db.Column(db.String(150), nullable=False)

class Configuracion(db.Model):
    """ Tabla global para almacenar variables del negocio como la tasa cambiaria """
    id = db.Column(db.Integer, primary_key=True)
    tasa_dolar = db.Column(db.Float, default=1.0)

# ========================================================
# RUTAS DE CONTROL DE FLUJO Y SEGURIDAD (AUTENTICACIÓN)
# ========================================================

@app.route('/bienvenida', methods=['GET', 'POST'])
def bienvenida():
    conn = sqlite3.connect('negocio.db') # Cambia por el nombre exacto de tu archivo .db
    cursor = conn.cursor()
    
    # 1. Contamos cuántos usuarios existen en la tabla
    cursor.execute("SELECT COUNT(*) FROM usuarios")
    cantidad_usuarios = cursor.fetchone()[0]
    conn.close()

    if request.method == 'POST':
        usuario_ingresado = request.form.get('usuario')
        password_ingresado = request.form.get('password')

        # CASO A: No hay usuarios en el sistema todavía (Primer encendido)
        if cantidad_usuarios == 0:
            if usuario_ingresado == 'admin' and password_ingresado == 'admin':
                session['primer_inicio_autorizado'] = True
                return redirect(url_for('configurar_dueno')) # Lo mandamos a crear su cuenta
            else:
                flash("Para la configuración inicial use el usuario maestro de fábrica.", "danger")
                return render_template('bienvenida.html')

        # CASO B: Ya el sistema está configurado con usuarios reales
        else:
            if usuario_ingresado == 'admin' and password_ingresado == 'admin':
                flash("La cuenta de fábrica 'admin' ha sido deshabilitada por seguridad.", "danger")
                return render_template('bienvenida.html')
            
            # Buscamos al usuario real en la base de datos
            conn = sqlite3.connect('negocio.db')
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM usuarios WHERE usuario = ? AND password = ?", (usuario_ingresado, password_ingresado))
            user = cursor.fetchone()
            conn.close()

            if user:
                session['logueado'] = True
                session['usuario_actual'] = usuario_ingresado
                return redirect(url_for('menu_seleccion')) # ¡Entra al sistema!
            else:
                flash("Usuario o contraseña incorrectos.", "danger")

    return render_template('bienvenida.html')

@app.route('/configurar_dueno', methods=['GET', 'POST'])
def configurar_dueno():
    # Evita que cualquiera entre a esta ruta si no puso admin/admin primero
    if not session.get('primer_inicio_autorizado'):
        return redirect(url_for('bienvenida'))

    if request.method == 'POST':
        nuevo_usuario = request.form.get('nuevo_usuario')
        nuevo_password = request.form.get('nuevo_password')

        if nuevo_usuario == 'admin':
            flash("No puede usar 'admin' como usuario real. Elija otro nombre.", "danger")
            return render_template('configurar_dueno.html')

        # Usamos un bloque try/except para atrapar el error de duplicado sin que se caiga el sistema
        try:
            conn = sqlite3.connect('negocio.db')
            cursor = conn.cursor()
            cursor.execute("INSERT INTO usuarios (usuario, password) VALUES (?, ?)", (nuevo_usuario, nuevo_password))
            conn.commit()
            conn.close()

            # Limpiamos la sesión temporal y lo logueamos oficialmente
            session.pop('primer_inicio_autorizado', None)
            session['logueado'] = True
            session['usuario_actual'] = nuevo_usuario

            return redirect(url_for('menu_seleccion')) # Va directo al menú principal ya protegido

        except sqlite3.IntegrityError:
            #  Si el usuario ya existe en la base de datos, entra aquí en lugar de dar error
            flash(f"El usuario '{nuevo_usuario}' ya está registrado en el sistema. Intente con otro nombre.", "danger")
            return render_template('configurar_dueno.html')

    return render_template('configurar_dueno.html')
    
@app.route('/menu')
def menu_seleccion():
    """ Pantalla intermedia para elegir entre ventas o finanzas después del login """
    if not session.get('logueado'): 
        return redirect(url_for('bienvenida'))
    return render_template('menu.html')

@app.route('/logout')
def logout():
    """ Destruye la sesión del usuario de forma segura """
    session.pop('logueado', None)
    flash("Has cerrado sesión correctamente.", "info")
    return redirect(url_for('bienvenida'))

@app.route('/registrar_admin', methods=['POST'])
def registrar_admin():
    """ Genera el único usuario maestro de la base de datos con sus preguntas secretas """
    if UsuarioAdmin.query.first():
        flash("El administrador ya está registrado.", "danger")
        return redirect(url_for('bienvenida'))
    
    password = request.form.get('password')
    p1 = request.form.get('pregunta1')
    r1 = request.form.get('respuesta1', '').strip().lower()
    p2 = request.form.get('pregunta2')
    r2 = request.form.get('respuesta2', '').strip().lower()
    p3 = request.form.get('pregunta3')
    r3 = request.form.get('respuesta3', '').strip().lower()

    if not password or not r1 or not r2 or not r3:
        flash("Todos los campos obligatorios deben completarse.", "warning")
        return redirect(url_for('bienvenida'))

    hashed_pw = generate_password_hash(password)
    
    nuevo_usuario = UsuarioAdmin(
        password_hash=hashed_pw,
        pregunta1=p1, respuesta1=r1,
        pregunta2=p2, respuesta2=r2,
        pregunta3=p3, respuesta3=r3
    )
    db.session.add(nuevo_usuario)
    db.session.commit()
    
    session['logueado'] = True
    flash("¡Administrador registrado con éxito! Bienvenido.", "success")
    return redirect(url_for('index'))

@app.route('/recuperar_password', methods=['POST'])
def recuperar_password():
    """ Evalúa las respuestas de seguridad para autorizar el reseteo de la clave corporativa """
    admin = UsuarioAdmin.query.first()
    if not admin:
        return redirect(url_for('bienvenida'))
        
    r1 = request.form.get('respuesta1', '').strip().lower()
    r2 = request.form.get('respuesta2', '').strip().lower()
    r3 = request.form.get('respuesta3', '').strip().lower()
    nueva_pw = request.form.get('nueva_password')
    
    if r1 == admin.respuesta1 and r2 == admin.respuesta2 and r3 == admin.respuesta3:
        admin.password_hash = generate_password_hash(nueva_pw)
        db.session.commit()
        flash("Contraseña restablecida con éxito. Inicia sesión.", "success")
    else:
        flash("Respuestas incorrectas. Verificación de identidad denegada.", "danger")
        
    return redirect(url_for('bienvenida'))

# ========================================================
# PANEL PRINCIPAL (INDEX) Y PROCESOS OPERATIVOS DEL NEGOCIO
# ========================================================

@app.route('/')
def index():
    if not session.get('logueado'):
        return redirect(url_for('bienvenida'))
        
    config = Configuracion.query.first()
    if not config:
        config = Configuracion(tasa_dolar=1.0)
        db.session.add(config)
        db.session.commit()

    tasa_actual = float(config.tasa_dolar)

    ingresos_lista = Ingreso.query.filter_by(cerrado=False).order_by(Ingreso.id.desc()).all()
    gastos_lista = Gasto.query.filter_by(cerrado=False).order_by(Gasto.id.desc()).all()
    
    # -------------------------------------------------------------
    # CONTABILIZACIÓN DE INGRESOS (En Caja)
    # -------------------------------------------------------------
    ingresos_en_dolares = 0.0
    ingresos_en_bs = 0.0

    for i in ingresos_lista:
        monto_registro = float(i.monto or 0.0)
        metodo = str(i.metodo_pago or "").lower().strip()
        
        es_divisa = ("dolar" in metodo or "dólar" in metodo or "divisa" in metodo or "zelle" in metodo or "usd" in metodo or "💵" in metodo)
        if es_divisa:
            ingresos_en_dolares += monto_registro
        else:
            # Si el ingreso se registró en Bs originalmente, lo sumamos directo a Bs
            ingresos_en_bs += monto_registro

    # -------------------------------------------------------------
    # CONTABILIZACIÓN DE GASTOS (En Caja)
    # -------------------------------------------------------------
    gastos_en_dolares = 0.0
    gastos_en_bs = 0.0

    for g in gastos_lista:
        monto_gasto = float(g.monto or 0.0)
        metodo_g = str(g.metodo_pago or "").lower().strip()
        
        es_divisa_g = ("dolar" in metodo_g or "dólar" in metodo_g or "divisa" in metodo_g or "zelle" in metodo_g or "usd" in metodo_g or "💵" in metodo_g)
        if es_divisa_g:
            gastos_en_dolares += monto_gasto
        else:
            # Si el gasto se registró en Bs originalmente, lo sumamos directo a Bs
            gastos_en_bs += monto_gasto

    # -------------------------------------------------------------
    #  GANANCIA NETA (Resta directa por moneda independiente)
    # -------------------------------------------------------------
    ganancia_neta_usd = ingresos_en_dolares - gastos_en_dolares
    ganancia_neta_bs = ingresos_en_bs - gastos_en_bs
    # -------------------------------------------------------------

    productos = Producto.query.all()
    clientes = Cliente.query.all()
    proveedores = Proveedor.query.all()
    cuentas_pagar = CuentaPorPagar.query.all()

    dinero_debe_usd = sum(float(c.monto or 0.0) for c in cuentas_pagar)
    # Calculamos el equivalente total en bolívares de esa deuda
    dinero_debe_bs = dinero_debe_usd * tasa_actual

    productos_ordenados = sorted(productos, key=lambda p: p.ventas_totales or 0, reverse=True)
    productos_estrella = [p for p in productos_ordenados if (p.ventas_totales or 0) > 0][:3]
    productos_frios = [p for p in productos if (p.ventas_totales or 0) == 0]

    return render_template('index.html',
                           productos=productos,
                           clientes=clientes,
                           proveedores=proveedores,
                           ingresos=ingresos_lista,
                           gastos=gastos_lista,
                           cuentas_pagar=cuentas_pagar,
                           dinero_generado=ingresos_en_dolares,
                           
                           # VARIABLES CLAVE PARA TU CUADRO ROJO:
                           dinero_debe=dinero_debe_usd,  # Envía el monto limpio en $
                           dinero_debe_bs=dinero_debe_bs, # Envía el monto convertido a Bs
                           
                           productos_estrella=productos_estrella,
                           productos_frios=productos_frios,
                           tasa_dolar=tasa_actual,
                           ingresos_en_dolares=ingresos_en_dolares,
                           ingresos_en_bs=ingresos_en_bs,
                           gastos_en_dolares=gastos_en_dolares,
                           gastos_en_bs=gastos_en_bs,
                           ganancia_neta_usd=ganancia_neta_usd,
                           ganancia_neta_bs=ganancia_neta_bs)

@app.route('/actualizar_tasa', methods=['POST'])
def actualizar_tasa():
    """ Modifica el multiplicador de conversión Bs/USD de forma global """
    if not session.get('logueado'): return redirect(url_for('bienvenida'))
    nueva_tasa = float(request.form.get('tasa_dolar', 1.0))
    config = Configuracion.query.first()
    if config:
        config.tasa_dolar = nueva_tasa
        db.session.commit()
        flash("Tasa del dólar actualizada correctamente.", "success")
    return redirect(url_for('index'))

@app.route('/agregar_producto', methods=['POST'])
def agregar_producto():
    """ Agrega un nuevo artículo al stock de inventario """
    if not session.get('logueado'): return redirect(url_for('bienvenida'))
    nombre = request.form.get('nombre')
    p_compra = float(request.form.get('precio_compra'))
    p_venta = float(request.form.get('precio_venta'))
    stock = int(request.form.get('stock'))

    nuevo_p = Producto(nombre=nombre, precio_compra=p_compra, precio_venta=p_venta, stock=stock)
    db.session.add(nuevo_p)
    db.session.commit()
    flash("Producto añadido con éxito.", "success")
    return redirect(url_for('index'))

@app.route('/registrar_ingreso', methods=['POST'])
def registrar_ingreso():
    """ Procesa tanto ventas directas de inventario como inyecciones de dinero externas """
    if not session.get('logueado'): 
        return redirect(url_for('bienvenida'))
        
    tipo_formulario = request.form.get('tipo_ingreso')
    metodo = request.form.get('metodo_pago')
    
    config = Configuracion.query.first()
    tasa = config.tasa_dolar if config else 1.0

    import datetime
    fecha_hoy = datetime.date.today().strftime("%Y-%m-%d")

    if tipo_formulario == "producto":
        # 🛒 PROCESAR VENTA DE PRODUCTOS DEL INVENTARIO (Queda igual, ya funciona bien)
        productos_ids = request.form.getlist('productos[]')
        cantidades = request.form.getlist('cantidades[]')
        
        total_venta_usd = 0.0
        detalles_venta = []

        for p_id, cant_str in zip(productos_ids, cantidades):
            if not p_id or not cant_str:
                continue
            
            cant = int(cant_str)
            if cant <= 0:
                continue

            prod = Producto.query.get(int(p_id))
            if prod:
                if prod.stock < cant:
                    flash(f"Error: Stock insuficiente para el producto '{prod.nombre}'.", "danger")
                    return redirect(url_for('index'))
                
                prod.stock -= cant
                prod.ventas_totales = (prod.ventas_totales or 0) + cant
                total_venta_usd += (prod.precio_venta * cant)
                detalles_venta.append(f"{prod.nombre} (x{cant})")

        if total_venta_usd > 0:
            descripcion_final = f"Venta: {', '.join(detalles_venta)}"
            nuevo_ingreso = Ingreso(
                tipo=descripcion_final,
                monto=total_venta_usd,
                metodo_pago=metodo,
                fecha=fecha_hoy,
                cerrado=False
            )
            db.session.add(nuevo_ingreso)
            db.session.commit() # Guardamos en la base de datos
            flash("Venta registrada con éxito.", "success")
        else:
            flash("No seleccionaste ningún producto válido.", "warning")
            return redirect(url_for('index'))

    else:
        # PROCESAR INGRESO EXTERNO 
        monto_ext_str = request.form.get('monto_externo')
        concepto = request.form.get('concepto_externo') or "Otros Ingresos"
        
        if monto_ext_str:
            monto_ext = float(monto_ext_str)
            
            # AQUÍ ESTÁ EL TRUCO INTELIGENTE:
            if metodo == "bs":
                # Si el usuario seleccionó Bolívares, calculamos su equivalente en USD dividiendo por la tasa
                # Así la columna 'monto' guarda el valor real equivalente en dólares
                monto_final_usd = monto_ext / tasa if tasa > 0 else monto_ext
                # Le añadimos una nota al concepto para que en tus tablas recuerdes cuántos Bs eran originalmente
                concepto_final = f"{concepto} ({monto_ext:.2f} Bs)"
            else:
                # Si seleccionó dólares en efectivo, el monto base es exactamente el valor plano digitado
                monto_final_usd = monto_ext
                concepto_final = concepto

            nuevo_ingreso = Ingreso(
                tipo=concepto_final,
                monto=monto_final_usd, # Guardamos el contravalor correcto en USD base
                metodo_pago=metodo,     # Guarda 'bs' o 'usd_efectivo' para saber a qué cuenta va
                fecha=fecha_hoy,
                cerrado=False
            )
            db.session.add(nuevo_ingreso)
            db.session.commit() # Guardamos en la base de datos
            flash("Ingreso externo registrado con éxito.", "success")
        else:
            flash("Monto externo no válido.", "warning")
            return redirect(url_for('index'))

    return redirect(url_for('index'))

@app.route('/registrar_gasto', methods=['POST'])
def registrar_gasto():
    """ Procesa el registro de un egreso o gasto del negocio """
    if not session.get('logueado'): 
        return redirect(url_for('bienvenida'))
        
    # Intentamos leer 'concepto', si no existe 'descripcion', y si no 'tipo'.
    concepto = request.form.get('concepto') or request.form.get('descripcion') or request.form.get('tipo') or "Gasto General"
    
    metodo = request.form.get('metodo_pago')
    monto_str = request.form.get('monto')

    # CAPTURAMOS LA TASA DINÁMICA
    tasa_raw = request.form.get('tasa_gasto')
    tasa_gasto = float(tasa_raw) if tasa_raw and tasa_raw.strip() != "" else 1.0

    import datetime
    fecha_hoy = datetime.date.today() 

    if monto_str:
        monto_original = float(monto_str)
        
        # VALIDACIÓN MULTIMONEDA: Si es en bolívares, hacemos la conversión a USD
        if metodo == 'bs':
            monto_gasto = monto_original / tasa_gasto
        else:
            monto_gasto = monto_original # Si es dólares, pasa directo

        # Creamos el registro en la base de datos con el monto ya en USD
        nuevo_gasto = Gasto(
            tipo=concepto,          # Se guarda en el campo 'tipo' de tu tabla
            monto=round(monto_gasto, 2), # Redondeamos a 2 decimales para evitar decimales infinitos     
            metodo_pago=metodo,     
            fecha=fecha_hoy,        
            cerrado=False
        )
        db.session.add(nuevo_gasto)
        db.session.commit()
        flash("Gasto registrado con éxito.", "success")
    else:
        flash("Monto de gasto no válido.", "warning")

    return redirect(url_for('index'))

@app.route('/agregar_cuenta_pagar', methods=['POST'])
def agregar_cuenta_pagar():
    if not session.get('logueado'): 
        return redirect(url_for('bienvenida'))
        
    proveedor_id = int(request.form.get('proveedor_id'))
    moneda = request.form.get('moneda')
    monto_original = float(request.form.get('monto_original', 0.0))
    descripcion = request.form.get('descripcion')
    fecha_limite = request.form.get('fecha_limite')

    # 🛠️ CORRECCIÓN AQUÍ: Evita el error si el input viene vacío desde el HTML
    tasa_raw = request.form.get('tasa_factura')
    tasa_factura = float(tasa_raw) if tasa_raw and tasa_raw.strip() != "" else 1.0

    # Convertimos a dólares si se registró originalmente en Bolívares
    if moneda == 'BS':
        monto_usd = monto_original / tasa_factura
    else:
        monto_usd = monto_original
        tasa_factura = 1.0  # Si es dólares, la tasa interna siempre es 1

    nueva_deuda = CuentaPorPagar(
        proveedor_id=proveedor_id,
        moneda=moneda,
        monto_original=round(monto_original, 2),
        tasa_factura=round(tasa_factura, 2),
        monto=round(monto_usd, 2),
        descripcion=descripcion,
        fecha_limite=fecha_limite,
        pagado=False
    )
    
    db.session.add(nueva_deuda)
    db.session.commit()
    
    flash("Cuenta por pagar registrada de forma inteligente.", "success")
    return redirect(url_for('finanzas'))

@app.route('/abonar_cuenta_pagar/<int:id>', methods=['POST'])
def abonar_cuenta_pagar(id):
    if not session.get('logueado'): 
        return redirect(url_for('bienvenida'))
    
    cuenta = CuentaPorPagar.query.get_or_404(id)
    
    moneda_pago = request.form.get('moneda_pago')
    monto_abonado = float(request.form.get('monto_abonado', 0.0))
    tasa_pago = float(request.form.get('tasa_pago', 1.0))
    
    # Calculamos el equivalente en dólares de lo que se está pagando en este instante
    if moneda_pago == 'BS':
        abono_usd = monto_abonado / tasa_pago
    else:
        abono_usd = monto_abonado
        tasa_pago = 1.0

    if abono_usd > (cuenta.monto + 0.01):
        flash(f"Error: El abono (${abono_usd:.2f}) supera el saldo pendiente (${cuenta.monto:.2f}).", "danger")
        return redirect(url_for('finanzas'))

    # Restamos del compromiso pendiente
    cuenta.monto = round(cuenta.monto - abono_usd, 2)
    
    # Si ya se cubrió la deuda por completo, la marcamos como resuelta
    if cuenta.monto <= 0.05:
        cuenta.pagado = True
        flash(f"¡Felicidades! Deuda liquidada con {cuenta.proveedor.nombre}.", "success")
    else:
        flash(f"Abono registrado por {monto_abonado} {moneda_pago}.", "success")

    # COMO SALIÓ DINERO REAL, AQUÍ SÍ LO REGISTRAMOS EN TU TABLA DE GASTOS
    detalles_gasto = f"Pago/Abono de deuda a {cuenta.proveedor.nombre}: {monto_abonado:.2f} {moneda_pago}"
    if moneda_pago == 'BS':
        detalles_gasto += f" (Tasa: {tasa_pago})"
        
    nuevo_gasto = Gasto(
        tipo="Proveedor (Pago de Deuda)",
        monto=round(abono_usd, 2),
        descripcion=detalles_gasto,
        metodo_pago="Efectivo Dólar" if moneda_pago == 'USD' else "Transferencia/Pago Móvil Bs",
        cerrado=False
    )
    db.session.add(nuevo_gasto)
    db.session.commit()
    
    return redirect(url_for('finanzas'))

@app.route('/abonar_cliente', methods=['POST'])
def abonar_cliente():
    """ Permite amortizar o liquidar la deuda de un cliente, actualizando su nivel """
    if not session.get('logueado'): return redirect(url_for('bienvenida'))
    clie_id = int(request.form.get('cliente_id'))
    monto_abono = float(request.form.get('monto', 0))
    metodo = request.form.get('metodo_pago')

    clie = Cliente.query.get(clie_id)
    if clie and monto_abono > 0:
        clie.deuda = max(0.0, clie.deuda - monto_abono)
        if clie.deuda == 0:
            clie.nivel = 'Al Día'
            
        nuevo_ingreso = Ingreso(monto=monto_abono, descripcion=f"Abono de cliente: {clie.nombre}", metodo_pago=metodo)
        db.session.add(nuevo_ingreso)
        db.session.commit()
        flash("Abono procesado correctamente.", "success")
    return redirect(url_for('index'))

@app.route('/pagar_cuenta', methods=['POST'])
def pagar_cuenta():
    """ Liquida total o parcialmente una obligación con tasa personalizada e inyecta el gasto en la caja activa """
    if not session.get('logueado'): 
        return redirect(url_for('bienvenida'))
    
    # SALVAVIDAS: Si la base de datos es nueva, aseguramos que exista la configuración de la tasa
    config = Configuracion.query.first()
    if not config:
        config = Configuracion(tasa_dolar=1.0)
        db.session.add(config)
        db.session.commit()
    
    cuenta_id = int(request.form.get('cuenta_id'))
    monto_a_pagar_usd = float(request.form.get('monto_a_pagar', 0.0))
    moneda_pago = request.form.get('moneda_pago')  # 'USD' o 'BS'
    tasa_pago = float(request.form.get('tasa_pago', 1.0)) 

    cuenta = CuentaPorPagar.query.get(cuenta_id)
    if cuenta:
        if monto_a_pagar_usd <= 0 or monto_a_pagar_usd > cuenta.monto:
            flash("Monto de abono inválido.", "danger")
            return redirect(url_for('index'))

        prov = Proveedor.query.get(cuenta.proveedor_id)
        
        # 1. Calcular el monto real que saldrá de los Gastos según la moneda
        if moneda_pago == 'BS':
            monto_gasto_real = monto_a_pagar_usd * tasa_pago
            metodo_gasto = f"Transferencia/Pago Móvil Bs (Tasa: {tasa_pago})"
            detalle_gasto = f"Abono de ${monto_a_pagar_usd:.2f} a deuda de {prov.nombre if prov else 'Proveedor'} pagado en Bs a tasa {tasa_pago}. (Ref: {cuenta.descripcion})"
        else:
            monto_gasto_real = monto_a_pagar_usd
            metodo_gasto = "Efectivo Dólar / Divisa"
            detalle_gasto = f"Abono de ${monto_a_pagar_usd:.2f} a deuda de {prov.nombre if prov else 'Proveedor'} pagado en USD. (Ref: {cuenta.descripcion})"

        # 2. Registrar el Gasto automático (CORREGIDO: Añadido el campo 'tipo')
        nuevo_gasto = Gasto(
            tipo="Proveedor",  # Puedes cambiarlo por "Variable" o "Otros" si usas otra categoría en tus modelos
            monto=monto_gasto_real, 
            descripcion=detalle_gasto, 
            metodo_pago=metodo_gasto,
            cerrado=False
        )
        db.session.add(nuevo_gasto)
        
        # 3. Restar el abono a la deuda general del proveedor
        if prov:
            prov.deuda_pendiente = max(0.0, prov.deuda_pendiente - monto_a_pagar_usd)
            
        # 4. Restar el abono a esta factura/cuenta específica
        cuenta.monto -= monto_a_pagar_usd

        if cuenta.monto <= 0.01:
            db.session.delete(cuenta)
            flash("Obligación liquidada por completo.", "success")
        else:
            flash(f"Abono de ${monto_a_pagar_usd:.2f} procesado con éxito. Saldo restante: ${cuenta.monto:.2f}", "success")
            
        db.session.commit()
        
    return redirect(url_for('index'))

# ========================================================
# GENERADOR DE CIERRE DE CAJA AUTOMÁTICO A EXCEL
# ========================================================

@app.route('/cierre_caja')
def closure_caja():
    """ Genera el libro Excel separando totales por moneda y calcula la ganancia neta del día """
    if not session.get('logueado'): return redirect(url_for('bienvenida'))

    ingresos_dia = Ingreso.query.filter_by(cerrado=False).all()
    gastos_dia = Gasto.query.filter_by(cerrado=False).all()

    if not ingresos_dia and not gastos_dia:
        flash("No hay movimientos activos hoy para realizar un cierre.", "warning")
        return redirect(url_for('index'))

    config = Configuracion.query.first()
    tasa = config.tasa_dolar if config else 1.0

    wb = openpyxl.Workbook()
    
    # Hoja 1: Ingresos
    ws1 = wb.active
    ws1.title = "Ingresos (Ventas)"
    ws1.views.sheetView[0].showGridLines = True
    
    # Hoja 2: Gastos
    ws2 = wb.create_sheet(title="Gastos e Inversiones")
    ws2.views.sheetView[0].showGridLines = True

    # Estilos corporativos
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill_ing = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_fill_gas = PatternFill(start_color="A62B2B", end_color="A62B2B", fill_type="solid")
    title_font = Font(name="Segoe UI", size=16, bold=True, color="1F497D")
    data_font = Font(name="Segoe UI", size=11)
    bold_font = Font(name="Segoe UI", size=11, bold=True)
    summary_header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    
    # Color especial para la ganancia neta (Verde suavizado para éxito)
    ganancia_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    ganancia_font = Font(name="Segoe UI", size=11, bold=True, color="375623")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    
    # === CONSTRUIR HOJA DE INGRESOS ===
    ws1.cell(row=2, column=2, value="REPORTE DIARIO DE INGRESOS - NEZ").font = title_font
    headers1 = ["ID", "Descripción", "Método de Pago", "Monto (USD)"]
    for col_num, header in enumerate(headers1, start=2):
        cell = ws1.cell(row=4, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill_ing
        cell.alignment = Alignment(horizontal="center")

    row_num = 5
    for ing in ingresos_dia:
        ws1.cell(row=row_num, column=2, value=ing.id).font = data_font
        ws1.cell(row=row_num, column=3, value=ing.tipo or "Venta").font = data_font
        ws1.cell(row=row_num, column=4, value=(ing.metodo_pago or 'bs').upper()).font = data_font
        m_cell = ws1.cell(row=row_num, column=5, value=ing.monto)
        m_cell.font = data_font
        m_cell.number_format = '#,##0.00'
        for c in range(2, 6): ws1.cell(row=row_num, column=c).border = thin_border
        row_num += 1

    start_row_ing = 5
    end_row_ing = row_num - 1 if row_num > 5 else 5


    # === CONSTRUIR HOJA DE EGRESOS (GASTOS) ===
    ws2.cell(row=2, column=2, value="REPORTE DIARIO DE GASTOS - NEZ").font = title_font
    headers2 = ["ID", "Descripción", "Método de Pago", "Monto (USD)"]
    for col_num, header in enumerate(headers2, start=2):
        cell = ws2.cell(row=4, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill_gas
        cell.alignment = Alignment(horizontal="center")

    row_num_g = 5
    for gas in gastos_dia:
        ws2.cell(row=row_num_g, column=2, value=gas.id).font = data_font
        ws2.cell(row=row_num_g, column=3, value=gas.descripcion or "Gasto").font = data_font
        ws2.cell(row=row_num_g, column=4, value=(gas.metodo_pago or 'bs').upper()).font = data_font
        m_cell = ws2.cell(row=row_num_g, column=5, value=gas.monto)
        m_cell.font = data_font
        m_cell.number_format = '#,##0.00'
        for c in range(2, 6): ws2.cell(row=row_num_g, column=c).border = thin_border
        row_num_g += 1

    start_row_gas = 5
    end_row_gas = row_num_g - 1 if row_num_g > 5 else 5


    # CUADRO RESUMEN COMPLETO (Se dibuja al final de la Hoja 1 para consolidar todo)
    r_idx = row_num + 2
    ws1.cell(row=r_idx, column=3, value="RESUMEN FINANCIERO DEL DÍA").font = bold_font
    ws1.cell(row=r_idx, column=3).fill = summary_header_fill
    ws1.cell(row=r_idx, column=4, value="BALANCE FINAL").font = bold_font
    ws1.cell(row=r_idx, column=4).fill = summary_header_fill
    ws1.cell(row=r_idx, column=3).border = thin_border
    ws1.cell(row=r_idx, column=4).border = thin_border
    
    # SECCIÓN 1: EFECTIVO (USD)
    ws1.cell(row=r_idx+1, column=3, value="(+) Total Ingresos Efectivo (USD)").font = data_font
    ws1.cell(row=r_idx+1, column=3).border = thin_border
    t_usd_efec = ws1.cell(row=r_idx+1, column=4, value=f'=SUMIF(D{start_row_ing}:D{end_row_ing}, "USD_EFECTIVO", E{start_row_ing}:E{end_row_ing})')
    t_usd_efec.font = data_font
    t_usd_efec.number_format = '"$"#,##0.00'
    ws1.cell(row=r_idx+1, column=4).border = thin_border

    ws1.cell(row=r_idx+2, column=3, value="(-) Total Gastos Efectivo (USD)").font = data_font
    ws1.cell(row=r_idx+2, column=3).border = thin_border
    # Va a buscar los gastos en dólares de la HOJA 2 directamente usando la sintaxis de Excel 'Gastos e Inversiones'!
    t_usd_gas = ws1.cell(row=r_idx+2, column=4, value=f'=SUMIF(\'Gastos e Inversiones\'!D{start_row_gas}:D{end_row_gas}, "USD_EFECTIVO", \'Gastos e Inversiones\'!E{start_row_gas}:E{end_row_gas})')
    t_usd_gas.font = data_font
    t_usd_gas.number_format = '"$"#,##0.00'
    ws1.cell(row=r_idx+2, column=4).border = thin_border

    # Fila de GANANCIA NETA EN USD (Resta del ingreso en efectivo menos gasto en efectivo)
    ws1.cell(row=r_idx+3, column=3, value="(=) GANANCIA NETA EFECTIVO (USD)").font = ganancia_font
    ws1.cell(row=r_idx+3, column=3).fill = ganancia_fill
    ws1.cell(row=r_idx+3, column=3).border = thin_border
    
    g_neto_usd = ws1.cell(row=r_idx+3, column=4, value=f'=D{r_idx+1}-D{r_idx+2}')
    g_neto_usd.font = ganancia_font
    g_neto_usd.fill = ganancia_fill
    g_neto_usd.number_format = '"$"#,##0.00'
    ws1.cell(row=r_idx+3, column=4).border = thin_border


    # SECCIÓN 2: BANCO (Bolívares)
    ws1.cell(row=r_idx+5, column=3, value="(+) Total Ingresos Banco (Bs)").font = data_font
    ws1.cell(row=r_idx+5, column=3).border = thin_border
    t_bs_banco = ws1.cell(row=r_idx+5, column=4, value=f'=SUMIF(D{start_row_ing}:D{end_row_ing}, "BS", E{start_row_ing}:E{end_row_ing}) * {tasa}')
    t_bs_banco.font = data_font
    t_bs_banco.number_format = '#,##0.00" Bs"'
    ws1.cell(row=r_idx+5, column=4).border = thin_border

    ws1.cell(row=r_idx+6, column=3, value="(-) Total Gastos Banco (Bs)").font = data_font
    ws1.cell(row=r_idx+6, column=3).border = thin_border
    t_bs_gas = ws1.cell(row=r_idx+6, column=4, value=f'=SUMIF(\'Gastos e Inversiones\'!D{start_row_gas}:D{end_row_gas}, "BS", \'Gastos e Inversiones\'!E{start_row_gas}:E{end_row_gas}) * {tasa}')
    t_bs_gas.font = data_font
    t_bs_gas.number_format = '#,##0.00" Bs"'
    ws1.cell(row=r_idx+6, column=4).border = thin_border

    # Fila de GANANCIA NETA EN BOLÍVARES (Resta del ingreso en banco menos gasto en banco)
    ws1.cell(row=r_idx+7, column=3, value="(=) GANANCIA NETA BANCO (Bs)").font = ganancia_font
    ws1.cell(row=r_idx+7, column=3).fill = ganancia_fill
    ws1.cell(row=r_idx+7, column=3).border = thin_border
    
    g_neto_bs = ws1.cell(row=r_idx+7, column=4, value=f'=D{r_idx+5}-D{r_idx+6}')
    g_neto_bs.font = ganancia_font
    g_neto_bs.fill = ganancia_fill
    g_neto_bs.number_format = '#,##0.00" Bs"'
    ws1.cell(row=r_idx+7, column=4).border = thin_border


    # Autoajuste automático de anchos de columna para que el texto largo no se corte
    for ws in [ws1, ws2]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 15)

    # El truco: Oculta los datos de las vistas del día marcándolos como cerrados sin eliminarlos de los históricos
    for ing in ingresos_dia: ing.cerrado = True
    for gas in gastos_dia: gas.cerrado = True
    db.session.commit()

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    import datetime
    fecha_str = datetime.datetime.now().strftime("%d-%m-%Y")
    
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"Cierre_Caja_{fecha_str}.xlsx"
    )

# Modifica tu ruta original /finanzas para que busque a los clientes y productos
@app.route('/finanzas')
def finanzas():
    if not session.get('logueado'): return redirect(url_for('bienvenida'))
    
    config = Configuracion.query.first()
    tasa_actual = float(config.tasa_dolar) if config else 1.0

    proveedores = Proveedor.query.all()
    cuentas_pagar = CuentaPorPagar.query.all()
    dinero_debe_usd = sum(float(c.monto or 0.0) for c in cuentas_pagar)

    # NUEVAS CONSULTAS COMPLEMENTARIAS
    clientes = Cliente.query.all()
    todos_productos = Producto.query.all()

    return render_template('finanzas.html',
                           proveedores=proveedores,
                           cuentas_pagar=cuentas_pagar,
                           dinero_debe=dinero_debe_usd,
                           tasa_dolar=tasa_actual,
                           clientes=clientes,
                           todos_productos=todos_productos)


@app.route('/crear_deuda_externa', methods=['POST'])
def crear_deuda_externa():
    nombre = request.form.get('nombre_cliente')
    telefono = request.form.get('telefono_cliente')
    monto = float(request.form.get('monto_deuda', 0.0))
    moneda = request.form.get('moneda')
    descripcion = request.form.get('descripcion')

    config = Configuracion.query.first()
    tasa = float(config.tasa_dolar) if config else 1.0
    
    # Dolarizar si viene en Bs usando la tasa general del mostrador
    monto_usd = monto / tasa if moneda == 'BS' else monto

    # Verificar o crear cliente unico
    cliente = Cliente.query.filter_by(nombre=nombre).first()
    if not cliente:
        cliente = Cliente(nombre=nombre, telefono=telefono, deuda_total=0.0)
        db.session.add(cliente)
        db.session.commit()

    nueva_deuda = Deuda(
        cliente_id=cliente.id,
        tipo="Externa",
        monto_inicial=monto_usd,
        saldo_pendiente=monto_usd,
        descripcion=descripcion
    )
    cliente.deuda_total += monto_usd
    db.session.add(nueva_deuda)
    db.session.commit()
    
    flash("Deuda externa registrada y congelada en dólares.", "success")
    return redirect(url_for('finanzas'))

@app.route('/crear_deuda_productos', methods=['POST'])
def crear_deuda_productos():
    nombre = request.form.get('nombre_cliente')
    telefono = request.form.get('telefono_cliente')
    producto_id = int(request.form.get('producto_id'))
    cantidad = int(request.form.get('cantidad', 1))

    producto = Producto.query.get(producto_id)
    if not producto or producto.stock < cantidad:
        flash("Error: Stock insuficiente para fiar ese producto.", "danger")
        return redirect(url_for('finanzas'))

    monto_total_usd = producto.precio_venta * cantidad

    # Descontar del inventario activo de inmediato
    producto.stock -= cantidad

    cliente = Cliente.query.filter_by(nombre=nombre).first()
    if not cliente:
        cliente = Cliente(nombre=nombre, telefono=telefono, deuda_total=0.0)
        db.session.add(cliente)
        db.session.commit()

    nueva_deuda = Deuda(
        cliente_id=cliente.id,
        tipo="Productos",
        monto_inicial=monto_total_usd,
        saldo_pendiente=monto_total_usd,
        descripcion=f"Fió {cantidad} un. de {producto.nombre} (${producto.precio_venta} c/u)"
    )
    cliente.deuda_total += monto_total_usd
    db.session.add(nueva_deuda)
    db.session.commit()

    flash(f"Productos fiados. Se restaron {cantidad} unidades de {producto.nombre}.", "success")
    return redirect(url_for('finanzas'))




@app.route('/abonar_deuda/<int:cliente_id>', methods=['POST'])
def abonar_deuda(cliente_id):
    monto_abono = float(request.form.get('monto_abono', 0.0))
    moneda = request.form.get('moneda_abono')

    config = Configuracion.query.first()
    tasa = float(config.tasa_dolar) if config else 1.0
    monto_usd = monto_abono / tasa if moneda == 'BS' else monto_abono

    cliente = Cliente.query.get(cliente_id)
    
    # Buscar la primera deuda pendiente que tenga el cliente para ir amortizando
    deuda = Deuda.query.filter_by(cliente_id=cliente_id, estado="Pendiente").first()
    
    if not deuda:
        flash("Este cliente no registra deudas activas individuales.", "warning")
        return redirect(url_for('finanzas'))

    # Registrar el historial analítico
    nuevo_abono = HistorialAbono(
        deuda_id=deuda.id,
        monto_usd=round(monto_usd, 2),
        monto_original=monto_abono,
        moneda_pago=moneda,
        tasa_momento=tasa
    )
    db.session.add(nuevo_abono)

    # Restar de los saldos
    deuda.saldo_pendiente -= monto_usd
    cliente.deuda_total -= monto_usd

    if deuda.saldo_pendiente <= 0.05: # Umbral de centavos por redondeo
        deuda.estado = "Pagada"
        deuda.saldo_pendiente = 0.0

    # EFECTO REBOTAR EN CAJA CHICA: Creamos un Ingreso directo en tu mostrador
    # Suponiendo que tu modelo de Ingresos se llama 'Ingreso' o 'Venta'
    # Creamos un ingreso genérico de caja para que sume a tus ganancias del día
    nuevo_ingreso_caja = Gasto( # Si usas la misma tabla para movimientos positivos/negativos con tipo
        tipo="Ingreso Abono",
        monto=round(monto_usd, 2),
        descripcion=f"Abono de cliente {cliente.nombre} ({moneda})",
        metodo_pago="Efectivo Dólar" if moneda == 'USD' else "Transferencia/Pago Móvil Bs",
        cerrado=False
    )
    db.session.add(nuevo_ingreso_caja)
    db.session.commit()

    flash(f"Abono de ${monto_usd:.2f} USD procesado e integrado a la caja del día.", "success")
    return redirect(url_for('finanzas'))


@app.route('/liquidar_todo/<int:cliente_id>', methods=['POST'])
def liquidar_todo(cliente_id):
    cliente = Cliente.query.get(cliente_id)
    deudas_activas = Deuda.query.filter_by(cliente_id=cliente_id, estado="Pendiente").all()
    
    monto_recuperado = cliente.deuda_total

    for d in deudas_activas:
        d.estado = "Pagada"
        d.saldo_pendiente = 0.0
        
    cliente.deuda_total = 0.0

    # Inyectar el dinero completo recuperado a la caja
    pago_completo_caja = Gasto(
        tipo="Ingreso Liquidación",
        monto=round(monto_recuperado, 2),
        descripcion=f"Liquidación total de cuenta: {cliente.nombre}",
        metodo_pago="Efectivo Dólar",
        cerrado=False
    )
    db.session.add(pago_completo_caja)
    db.session.commit()

    flash(f"Cuenta saldada por completo. Se ingresaron ${monto_recuperado:.2f} USD a caja.", "success")
    return redirect(url_for('finanzas'))


@app.route('/agregar_producto_finanzas', methods=['POST'])
def agregar_producto_finanzas():
    if not session.get('logueado'): 
        return redirect(url_for('bienvenida'))
        
    nombre = request.form.get('nombre')
    moneda_compra = request.form.get('moneda_compra')
    total_factura = float(request.form.get('total_factura', 0.0))
    stock_comprado = int(request.form.get('stock', 0))
    porcentaje_ganancia = float(request.form.get('porcentaje_ganancia', 0.0))
    proveedor_id = int(request.form.get('proveedor_id'))
    
    # Leemos la tasa ingresada en el formulario
    tasa_factura = float(request.form.get('tasa_factura', 1.0))

    # Convertimos el costo usando la tasa específica dada por el proveedor
    if moneda_compra == 'BS':
        total_factura_usd = total_factura / tasa_factura
    else:
        total_factura_usd = total_factura

    # Lógica de cálculo matemático estándar
    costo_unitario_usd = total_factura_usd / stock_comprado
    precio_venta_usd = costo_unitario_usd * (1 + (porcentaje_ganancia / 100))

    # 1. Registramos el producto con sus costos dolarizados reales
    # AGREGAMOS: ventas_totales=0 para que la base de datos sepa que arranca limpio
    nuevo_producto = Producto(
        nombre=nombre,
        stock=stock_comprado,
        costo_compra=round(costo_unitario_usd, 2),
        precio_venta=round(precio_venta_usd, 2),
        proveedor_id=proveedor_id,
        ventas_totales=0  
    )
    db.session.add(nuevo_producto)

    # 2. MODIFICAMOS AQUÍ: Pasamos 'monto=round(total_factura_usd, 2)' 
    # para que en la caja se reste el equivalente real en dólares y no los bolívares completos.
    nuevo_gasto = Gasto(
        tipo="Proveedor",
        monto=round(total_factura_usd, 2),  
        descripcion=f"Compra de {stock_comprado} un. de {nombre} ({total_factura:.2f} Bs a tasa {tasa_factura})",
        metodo_pago="Efectivo Dólar" if moneda_compra == 'USD' else "Transferencia/Pago Móvil Bs",
        cerrado=False
    )
    db.session.add(nuevo_gasto)
    db.session.commit()

    return redirect(url_for('finanzas'))

@app.route('/agregar_proveedor_finanzas', methods=['POST'])
def agregar_proveedor_finanzas():
    if not session.get('logueado'): 
        return redirect(url_for('bienvenida'))
        
    nombre = request.form.get('nombre')
    telefono = request.form.get('telefono')
    rubro = request.form.get('rubro') 

    if nombre:
        #  CORREGIDO: Eliminamos la línea 'deuda_pendiente=0.0' porque 
        # ahora las deudas se calculan solas de forma inteligente con la Fase 3.
        nuevo_proveedor = Proveedor(
            nombre=nombre,
            telefono=telefono,
            rubro=rubro if rubro else "General"
        )
        db.session.add(nuevo_proveedor)
        db.session.commit()
        flash(f"Proveedor '{nombre}' guardado correctamente.", "success")
    else:
        flash("El nombre del proveedor es obligatorio.", "danger")
        
    return redirect(url_for('finanzas'))

def inicializar_base_de_datos():
    conn = sqlite3.connect('negocio.db') # Asegúrate de usar el mismo nombre de archivo .db que tienes en tus rutas
    cursor = conn.cursor()
    
    # Creamos la tabla usuarios con columnas para id, usuario y password
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT NOT NULL UNIQUE,
            password TEXT NOT NULL
        )
    ''')
    
    conn.commit()
    conn.close()

inicializar_base_de_datos()

if __name__ == '__main__':
    
    with app.app_context():
        db.create_all()  
        
    app.run(debug=True)