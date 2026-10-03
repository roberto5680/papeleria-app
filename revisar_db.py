import sqlite3

db = r"C:\abarrotes_app\DATOS_PAPELERIA.db"

conn = sqlite3.connect(db)

tablas = conn.execute(
    "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
).fetchall()

print("BASE:")
print(db)

print("\nTABLAS:")
for tabla in tablas:
    print("-", tabla[0])

if any(tabla[0] == "DATOS_PAPELERIA" for tabla in tablas):
    cantidad = conn.execute(
        "SELECT COUNT(*) FROM DATOS_PAPELERIA"
    ).fetchone()[0]

    print("\nCANTIDAD DE PRODUCTOS:", cantidad)

    productos = conn.execute(
        "SELECT codigo, descripcion, inventario, precio FROM DATOS_PAPELERIA LIMIT 5"
    ).fetchall()

    print("\nPRIMEROS PRODUCTOS:")
    for producto in productos:
        print(producto)
else:
    print("\nNO EXISTE LA TABLA DATOS_PAPELERIA")

conn.close()