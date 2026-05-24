from flask import Flask, render_template, request, redirect, url_for
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)

# Configuramos la base de datos local (se guardará como un archivo en tu carpeta)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///negocio.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# =======================================================
# 📊 TABLAS DE LA BASE DE DATOS (MODELOS SQL)
# =======================================================

class Producto(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    precio_venta = db.Column(db.Float, nullable=False)
    stock = db.Column(db.Integer, default=0)
    ventas_totales = db.Column(db.Integer, default=0) # Para identificar el más vendido

class Cliente(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    telefono = db.Column(db.String(20))
    estado_pago = db.Column(db.String(50), default='Al día') # Ej: 'Al día', 'Deudor'
    deuda = db.Column(db.Float, default=0.0)

class Proveedor(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    telefono = db.Column(db.String(20))
    rubro = db.Column(db.String(100)) # Qué producto te surte

# Crear las tablas automáticamente si no existen
with app.app_context():
    db.create_all()

# =======================================================
# 🌐 RUTAS DE LA PÁGINA WEB
# =======================================================

@app.route('/')
def index():
    # Consultamos todos los registros actuales en la base de datos
    productos = Producto.query.order_by(Producto.ventas_totales.desc()).all()
    clientes = Cliente.query.all()
    proveedores = Proveedor.query.all()
    
    # Cálculos matemáticos básicos para las finanzas
    ingresos = sum(p.ventas_totales * p.precio_venta for p in productos)
    deudas_totales = sum(c.deuda for c in clientes)

    # Enviamos estos datos al diseño visual (HTML)
    return render_template('index.html', productos=productos, clientes=clientes, proveedores=proveedores, ingresos=ingresos, deudas=deudas_totales)

if __name__ == '__main__':
    app.run(debug=True)