# MedTrack: Control de Producción Diaria

Aplicación de escritorio para el registro y control de producción diaria en pequeños talleres o fábricas. Desarrollada con **Streamlit** y **MySQL**, permite gestionar de forma centralizada los trabajadores, tareas realizadas, cantidades producidas y fechas.

## Contexto del problema

En muchos entornos productivos pequeños, el seguimiento de la producción diaria se realiza aún con **papel, libretas o archivos Excel dispersos**. Esto genera:

- Información desorganizada y difícil de consultar.
- Pérdida de trazabilidad sobre quién hizo qué tarea y cuándo.
- Errores humanos al transcribir datos.
- Falta de una base de datos centralizada que permita filtrar, buscar y analizar el rendimiento.

**MedTrack** resuelve estos problemas ofreciendo una interfaz sencilla y rápida para registrar, editar, eliminar y consultar la producción, con filtros por nombre y fecha, y exportación a Excel.

## Funcionalidades

- **Login de administrador** (único usuario, fácilmente ampliable).
- **Registro de trabajadores y tareas**: cédula, nombres, apellidos, trabajo realizado, cantidad y fecha.
- **Búsqueda en tiempo real** por nombre, apellido, trabajo o cédula.
- **Filtro por fecha exacta** para ver la producción de un día concreto.
- **Tabla interactiva** con todos los registros.
- **CRUD completo**: agregar, editar y eliminar registros con validaciones.
- **Exportación a Excel** con formato profesional (cédulas como texto, fechas legibles, columnas ajustadas automáticamente).
- **Interfaz limpia y responsive** con CSS personalizado.

## Tecnologías utilizadas

- **Python 3.13+**
- **Streamlit** – Framework para la interfaz gráfica web.
- **MySQL** – Base de datos relacional.
- **pandas** – Manejo y transformación de datos.
- **mysql-connector-python** – Conector a MySQL.
- **openpyxl** – Generación y formateo de archivos Excel.
- **XAMPP** – Entorno local con phpMyAdmin para gestionar la BD.

---

# Instalación y Ejecución

## 1. Instalar dependencias

Después de descargar el repositorio, asegúrate de tener Python 3.13 o superior instalado. Luego ejecuta:

```bash
pip install streamlit pandas mysql-connector-python openpyxl
```

## 2. Configurar la base de datos MySQL

- Instala **XAMPP**
- Inicia los servicios Apache y MySQL desde el panel de XAMPP.
- Abre phpMyAdmin (```http://localhost/phpmyadmin```).
- Crea una base de datos llamada ```medtgol```.
- Ejecuta la siguiente consulta SQL para crear la tabla ```trabajadores```:
```sql
CREATE TABLE trabajadores (
    id INT AUTO_INCREMENT PRIMARY KEY,
    cedula INT(10) NOT NULL,
    nombres VARCHAR(50) NOT NULL,
    apellidos VARCHAR(50) NOT NULL,
    trabajo VARCHAR(50) NOT NULL,
    cantidad INT(10) NOT NULL,
    fecha DATETIME NOT NULL
);
```

## 3. Ejecutar la aplicación

Desde la terminal, en la carpeta del proyecto:

```bash
streamlit run app.py
```

La aplicación se abrirá automáticamente en tu navegador en ```http://localhost:8501```.

## 4. Credenciales de acceso (Login del programa)

- Usuario: ```admin```
- Contraseña: ```123456```

---

# Estructura del Proyecto
```text
MedTrack/
│
├── app.py                  # Código principal de la aplicación
├── img/
│   └── medtgol_logo.png    # Logo
├── README.md               # Este archivo
```

---

# Uso Básico

- **Iniciar sesión** con las credenciales.
- **Agregar un registro:** completa el formulario "Agregar nuevo registro" y haz clic en "Guardar registro".
- **Buscar y filtrar:** escribe en la barra lateral o selecciona una fecha.
- **Editar:** selecciona un ID en el desplegable de "Editar", modifica los campos y actualiza.
- **Eliminar:** selecciona un ID, marca la casilla de confirmación y pulsa "Eliminar definitivamente".
- **Exportar a Excel:** haz clic en el botón "Descargar Excel" (se descargará con los filtros aplicados).
- **Cerrar sesión** con el botón en la barra lateral.

---

# Autores

- Juan José Torres Betancur
- David Felipe Rosales Arcos
- Melissa Micolta Cuellar
- Simon Satizabal Conde
