# IMPORTAR LIBRERÍAS NECESARIAS
import streamlit as st
import pandas as pd     # para manejar tablas (DataFrame)
import mysql.connector  # para conectarse a MySQL
from mysql.connector import Error  # para capturar errores de MySQL
import io # para manejar archivos en memoria (Excel)
from openpyxl.styles import numbers # para formatear celdas de Excel

# Configuración de página global (una sola vez por ejecución)
st.set_page_config(page_title="Control de Producción", layout="wide")


# CONFIGURACIÓN DE LA BASE DE DATOS
DB_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'medtgol'
}

# Convierte un DataFrame a bytes de Excel para descargar
def convertir_a_excel(df):
    # Copiamos el DataFrame para no modificar el original
    df_excel = df.copy()
    
    # Convertir cédula a string para evitar notación científica
    df_excel['cedula'] = df_excel['cedula'].astype(str)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_excel.to_excel(writer, index=False, sheet_name='Registros')
        worksheet = writer.sheets['Registros']
        
        # Ajustar ancho de todas las columnas según el texto más largo
        for col in worksheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = min(max_length + 2, 40)  # ancho máximo 40
            worksheet.column_dimensions[col_letter].width = adjusted_width
        
        # Aplicar formato de fecha a la columna 'fecha'
        fecha_col_idx = None
        for idx, col_name in enumerate(df_excel.columns, start=1):
            if col_name == 'fecha':
                fecha_col_idx = idx
                break
        if fecha_col_idx:
            for row in range(2, len(df_excel) + 2):  # fila 1 es encabezado
                cell = worksheet.cell(row=row, column=fecha_col_idx)
                cell.number_format = 'yyyy-mm-dd'
    
    return output.getvalue()

# FUNCIÓN PARA OBTENER LA CONEXIÓN A MySQL
def get_connection():
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        return conn
    except Error as e:
        st.error(f"Error de conexión a la base de datos: {e}")
        return None

def css():
    st.markdown("""
    <style>
        /* Fondo general */
        .stApp {
            background-color: #f5f7fa;
        }
        /* Títulos */
        h1, h2, h3 {
            color: #1e3c72;
            font-family: 'Segoe UI', sans-serif;
        }
        /* Tarjetas para los formularios */
        .css-1r6slb0, .css-1v3fvcr {
            background-color: #ffffff;
            border-radius: 15px;
            padding: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        /* Botones */
        .stButton > button {
            background-color: #1e3c72;
            color: white;
            border-radius: 8px;
            border: none;
            padding: 8px 20px;
            font-weight: bold;
            transition: 0.2s;
        }
        .stButton > button:hover {
            background-color: #0f2b4f;
            transform: scale(1.02);
        }
        /* Dataframe */
        .stDataFrame {
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
        }
        /* Sidebar */
        .css-1d391kg {
            background-color: #e9f0f9;
        }
        /* Inputs */
        .stTextInput > div > div > input, .stNumberInput > div > div > input {
            border-radius: 8px;
            border: 1px solid #ccc;
        }
        /* Mensajes de éxito y error */
        .stAlert {
            border-radius: 10px;
            font-weight: bold;
        }
    </style>
    """, unsafe_allow_html=True)

# FUNCIÓN PARA CARGAR LOS DATOS DESDE LA TABLA "trabajadores"
def cargar_datos(busqueda="", fecha=None):
    conn = get_connection()
    if conn is None:
        return pd.DataFrame()

    query = "SELECT id, cedula, nombres, apellidos, trabajo, cantidad, fecha FROM trabajadores"
    params = []
    condiciones = []

    if busqueda:
        condiciones.append("(nombres LIKE %s OR apellidos LIKE %s OR trabajo LIKE %s OR CAST(cedula AS CHAR) LIKE %s)")
        like = f"%{busqueda}%"
        params.extend([like, like, like, like])

    if fecha:
        condiciones.append("DATE(fecha) = %s")
        params.append(fecha)

    if condiciones:
        query += " WHERE " + " AND ".join(condiciones)

    query += " ORDER BY fecha DESC"
    df = pd.read_sql(query, conn, params=params)
    conn.close()
    return df

# FUNCIÓN PARA AGREGAR UN NUEVO REGISTRO
def agregar_registro(cedula, nombres, apellidos, trabajo, cantidad, fecha):
    conn = get_connection()
    if conn is None:
        return False
    cursor = conn.cursor()
    try:
        sql = "INSERT INTO trabajadores (cedula, nombres, apellidos, trabajo, cantidad, fecha) VALUES (%s, %s, %s, %s, %s, %s)"
        valores = (cedula, nombres, apellidos, trabajo, cantidad, fecha)
        cursor.execute(sql, valores)
        conn.commit()   # guarda los cambios en la BD
        return True
    except Error as e:
        st.error(f"Error al insertar: {e}")
        return False
    finally:
        cursor.close()
        conn.close()

# FUNCIÓN PARA ACTUALIZAR UN REGISTRO EXISTENTE
def actualizar_registro(id_reg, cedula, nombres, apellidos, trabajo, cantidad, fecha):
    conn = get_connection()
    if conn is None:
        return False
    cursor = conn.cursor()
    try:
        sql = "UPDATE trabajadores SET cedula=%s, nombres=%s, apellidos=%s, trabajo=%s, cantidad=%s, fecha=%s WHERE id=%s"
        valores = (cedula, nombres, apellidos, trabajo, cantidad, fecha, id_reg)
        cursor.execute(sql, valores)
        conn.commit()
        return True
    except Error as e:
        st.error(f"Error al actualizar: {e}")
        return False
    finally:
        cursor.close()
        conn.close()

# FUNCIÓN PARA ELIMINAR UN REGISTRO POR SU ID
def eliminar_registro(id_reg):
    conn = get_connection()
    if conn is None:
        return False
    cursor = conn.cursor()
    try:
        cursor.execute("DELETE FROM trabajadores WHERE id=%s", (id_reg,))
        conn.commit()
        return True
    except Error as e:
        st.error(f"Error al eliminar: {e}")
        return False
    finally:
        cursor.close()
        conn.close()

# PANTALLA DE LOGIN
def login():
    css()
    # Centrar el formulario de login dentro del layout wide
    cols = st.columns([1, 2, 1])
    with cols[1]:
        st.title("Inicio de Sesión - Control de Producción")
        st.subheader("Acceso exclusivo administrador")
        
        usuario_input = st.text_input("Usuario")
        contraseña_input = st.text_input("Contraseña", type="password")
        
        if st.button("Iniciar Sesión"):
            if usuario_input == "admin" and contraseña_input == "123456":
                st.session_state["logged_in"] = True   # guardamos que ha iniciado sesión
                st.rerun()  # forzamos a que Streamlit vuelva a ejecutar el script (y ya no mostrará login)
            else:
                st.error("Usuario o contraseña incorrectos")

# APLICACIÓN PRINCIPAL (CRUD)
def app_principal():
    css()
    # Mostrar mensaje pendiente (si existe)
    if "mensaje" in st.session_state:
        tipo, texto = st.session_state["mensaje"]
        if tipo == "exito":
            st.success(texto)
        elif tipo == "error":
            st.error(texto)
        elif tipo == "toast":
            st.toast(texto, icon="✅")
        del st.session_state["mensaje"]  # lo borramos tras mostrarlo
    
    # Logo en la barra lateral
    st.sidebar.image("img/medtgol_logo.png")
    st.title("Registro de Producción Diaria")
    
    # BARRA LATERAL
    with st.sidebar:
        st.header("🔍 Filtrar registros")
        # Cuadro de búsqueda: su valor se guarda en la sesión para que persista
        busqueda = st.text_input("Buscar por nombre, apellido, trabajo o cédula", key="search_input")
        st.markdown("---")
        st.subheader("📅 Filtrar por fecha")
        filtro_fecha = st.date_input("Mostrar registros desde:", value=None, key="fecha_filtro")
        st.markdown("---")
        # Botón para cerrar sesión
        if st.button("Cerrar sesión"):
            st.session_state["logged_in"] = False
            st.rerun()
    
    # Cargamos los datos desde la BD (aplicando filtro de búsqueda si lo hay)
    df = cargar_datos(busqueda, filtro_fecha)
    
    # MOSTRAR TABLA DE REGISTROS
    if df.empty:
        st.info("No hay registros aún. Agrega uno usando el formulario de abajo.")
    else:
        # Dos columnas: título y botón de descarga
        col_titulo, col_boton = st.columns([3, 1])
        with col_titulo:
            st.subheader(f"Listado de trabajadores ({len(df)} registros)")
        with col_boton:
            excel_data = convertir_a_excel(df)
            st.download_button(
                label="Descargar Excel",
                data=excel_data,
                file_name="registros_produccion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        st.dataframe(df, use_container_width=True, hide_index=True)
    
    # FORMULARIO PARA AGREGAR NUEVO REGISTRO
    st.markdown("---")
    st.subheader("Agregar nuevo registro")
    
    # Usamos st.form para que al enviar no se recargue hasta que pulsemos el botón
    with st.form("form_agregar"):
        # Dividimos en dos columnas para mejor disposición
        col1, col2 = st.columns(2)
        with col1:
            cedula_nueva = st.text_input("Cédula", max_chars=10)
            nombres_nuevos = st.text_input("Nombres")
            apellidos_nuevos = st.text_input("Apellidos")
        with col2:
            trabajo_nuevo = st.text_input("Trabajo realizado")
            cantidad_nueva = st.number_input("Cantidad", value=0, step=1)
            fecha_nueva = st.date_input("Fecha")  # devuelve un objeto date
        
        # Botón de envío del formulario
        enviado = st.form_submit_button("Guardar registro")
        if enviado:
            # Validamos que los campos obligatorios no estén vacíos
            if not nombres_nuevos or not apellidos_nuevos or not trabajo_nuevo:
                st.error("❌ Los campos Nombres, Apellidos y Trabajo son obligatorios")
            elif cantidad_nueva < 0:
                st.error("❌ La cantidad no puede ser negativa")
            elif not cedula_nueva.isdigit():
                st.error("❌ La cédula debe contener solo números y no puede estar vacía")
            else:
                cedula_nueva_int = int(cedula_nueva)
                # Llamamos a la función que inserta en la BD
                exito = agregar_registro(cedula_nueva_int, nombres_nuevos, apellidos_nuevos, trabajo_nuevo, cantidad_nueva, fecha_nueva)
                if exito:
                    st.session_state["mensaje"] = ("exito", "✅ Registro agregado correctamente")
                    st.rerun()  # Recargamos para que se vea el nuevo registro en la tabla
                else:
                    st.session_state["mensaje"] = ("error", "❌ No se pudo agregar el registro")
                    st.rerun()
    
    # EDITAR REGISTRO
    st.markdown("---")
    st.subheader("Editar un registro existente")
    
    if not df.empty:
        id_seleccionado = st.selectbox("Selecciona ID del registro a editar", df['id'].tolist())
        registro = df[df['id'] == id_seleccionado].iloc[0]
        
        with st.form("form_editar"):
            col1, col2 = st.columns(2)
            with col1:
                edit_cedula = st.text_input("Cédula", value=str(int(registro['cedula'])), max_chars=10)
                edit_nombres = st.text_input("Nombres", value=registro['nombres'])
                edit_apellidos = st.text_input("Apellidos", value=registro['apellidos'])
            with col2:
                edit_trabajo = st.text_input("Trabajo", value=registro['trabajo'])
                edit_cantidad = st.number_input("Cantidad", value=int(registro['cantidad']), step=1)
                edit_fecha = st.date_input("Fecha", value=pd.to_datetime(registro['fecha']).date())
            
            if st.form_submit_button("Actualizar registro"):
                if not edit_cedula.isdigit():
                    st.error("❌ La cédula debe contener solo números")
                elif edit_cantidad < 0:
                    st.error("❌ La cantidad no puede ser negativa")
                else:
                    edit_cedula_int = int(edit_cedula)
                    if actualizar_registro(id_seleccionado, edit_cedula_int, edit_nombres, edit_apellidos, edit_trabajo, edit_cantidad, edit_fecha):
                        st.session_state["mensaje"] = ("exito", "✅ Registro actualizado")
                        st.rerun()
                    else:
                        st.session_state["mensaje"] = ("error", "❌ Error al actualizar")
                        st.rerun()
    else:
        st.info("No hay registros para editar.")
    
    # ELIMINAR REGISTRO
    st.subheader("Eliminar un registro")
    if not df.empty:
        id_eliminar = st.selectbox("Selecciona ID del registro a eliminar", df['id'].tolist(), key="delete_select")
        # Confirmación con checkbox para evitar borrados accidentales
        confirmar = st.checkbox(f"Confirmar que quiero eliminar el registro con ID {id_eliminar}")
        if st.button("Eliminar definitivamente"):
            if confirmar:
                if eliminar_registro(id_eliminar):
                    st.session_state["mensaje"] = ("exito", f"✅ Registro {id_eliminar} eliminado")
                    st.rerun()
                else:
                    st.session_state["mensaje"] = ("error", "❌ Error al eliminar")
                    st.rerun()
            else:
                st.warning("Marca la casilla de confirmación antes de eliminar.")
    else:
        st.info("No hay registros para eliminar.")

# PUNTO DE ENTRADA PRINCIPAL (esto se ejecuta al iniciar la app)
def main():
    # Inicializamos la variable de sesión "logged_in" si no existe
    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False
    
    # Según el estado, mostramos la pantalla de login o la app principal
    if not st.session_state["logged_in"]:
        login()
    else:
        app_principal()

# Llamamos a main() solo si este archivo es ejecutado directamente
if __name__ == "__main__":
    main()