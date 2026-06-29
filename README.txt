GESTOR DE CONTACTOS
By MSV
Guía de instalación y uso

Descripción:
Este proyecto es una aplicación web realizada con el Framework Flask, el cual permite administrar una lista de contactos. Toda la información se guarda y lee directamente desde un archivo de Excel.

Estructura de archivos:

app.py: Archivo principal que contiene las rutas y la lógica en Python.
contactos.xlsx: Archivo de Excel que funciona como base de datos.
templates/login-page.html: Pantalla para iniciar sesión.
templates/contacts.html: Pantalla principal con la tabla de contactos.
templates/nuevo_contacto.html: Formulario para registrar un contacto nuevo.
templates/editar_contacto.html: Formulario para modificar un contacto existente.
templates/buscar_contacto.html: Pantalla para buscar contactos por nombre o apellido.
templates/perfil.html: Pantalla que muestra el detalle completo de un contacto.
templates/reporte.html: Pantalla con las estadísticas y conteos del sistema.
static/css/: Archivos de diseño y estilos visuales.
static/Images/: Imágenes y logos utilizados en la aplicación.

Instrucciones para ejecutar el proyecto

1. Instalar las librerías necesarias:
Abra la terminal de su computadora y ejecute el siguiente comando para instalar Flask y la herramienta de Excel:
pip install Flask Openpyxl

2. Iniciar la aplicación:
En la misma terminal, ejecute el archivo principal con este comando:
python app.py
O bien, una vez abierto en Visual Studio, presione F5.

3. Abrir el sistema:
Una vez que el servidor esté corriendo, abra su navegador web e ingrese a la siguiente dirección:
http://127.0.0.1:5000
O bien, utilice la combinación de teclas Ctrl + C para abrir este mismo.

Credenciales de acceso
Para ingresar al sistema, debe utilizar los siguientes datos en la pantalla de inicio de sesión:
Usuario: admin
Contraseña: 1234

COTEPECOS
Subárea: Programación para Web
Nivel: Undécimo Año