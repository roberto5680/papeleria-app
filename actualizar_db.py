import sqlite3

def crear_tabla_zonas():
    # Conexión a la base de datos local
    conn = sqlite3.connect('database.db')
    cursor = conn.cursor()
    
    # Crear la tabla de zonas de entrega si no existe
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS zonas_entrega (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_postal VARCHAR(10) UNIQUE NOT NULL,
            disponible BOOLEAN DEFAULT 1
        )
    ''')
    
    # Agregar códigos postales iniciales de prueba
    cps_demo = ['68000', '68010', '68020', '68030']
    for cp in cps_demo:
        cursor.execute('''
            INSERT OR IGNORE INTO zonas_entrega (codigo_postal, disponible)
            VALUES (?, 1)
        ''', (cp,))
        
    conn.commit()
    conn.close()
    print("✅ Tabla 'zonas_entrega' creada e inicializada en la base de datos local.")

if __name__ == '__main__':
    crear_tabla_zonas()