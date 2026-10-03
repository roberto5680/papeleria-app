import sqlite3

def mostrar_o_resetear_admin():
    conn = sqlite3.connect('database.db')
    cursor = conn.cursor()

    # 1. Consultar estructura y contenido de ADMIN_USUARIOS
    cursor.execute("PRAGMA table_info(ADMIN_USUARIOS)")
    columnas = [col[1] for col in cursor.fetchall()]
    print("📌 Columnas en ADMIN_USUARIOS:", columnas)

    cursor.execute("SELECT * FROM ADMIN_USUARIOS")
    usuarios = cursor.fetchall()
    
    print("\n👤 Usuarios encontrados:")
    for user in usuarios:
        print(user)

    # Si quieres cambiar la contraseña del usuario existente a 'admin123' (o el usuario que tengas):
    # (Descomenta las líneas de abajo si las contraseñas no están en texto plano o si quieres cambiarla)
    
    # cursor.execute("UPDATE ADMIN_USUARIOS SET password = '123' WHERE usuario = 'admin'") # O ajusta los nombres de columna
    # conn.commit()
    # print("\n✅ Contraseña actualizada.")

    conn.close()

if __name__ == '__main__':
    mostrar_o_resetear_admin()