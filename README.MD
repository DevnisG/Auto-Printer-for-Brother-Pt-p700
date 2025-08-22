# PrintSocketApp

**PrintSocketApp** es una aplicación cliente-servidor pensada como aporte a la comunidad para automatizar la impresión de etiquetas con impresoras Brother (usando el SDK b-PAC).  
La aplicación permite que diferentes PCs envíen sus números de serie a un servidor central, que luego los imprime automáticamente en una plantilla de etiqueta (`.lbx`).

⚠️ **Nota importante:** Este proyecto no tiene ninguna relación con Brother. No es un producto oficial, ni está soportado o aprobado por la empresa. Es solo un ejemplo educativo y comunitario.

---

## 🚀 Características

- Comunicación **cliente-servidor TCP/IP**.
- Obtención automática del número de serie del BIOS de la PC cliente.
- Envío de datos desde el cliente al servidor.
- Impresión automática en una impresora Brother compatible mediante una plantilla `.lbx`.
- Código simple y extensible en **Python**.

---

## 📂 Estructura del proyecto

printer_client/
├── client.py # Script cliente: obtiene el número de serie y lo envía al servidor
└── ip_port.txt # Configuración de IP y puerto del servidor

printer_server/
├── server.py # Script servidor: recibe el dato y ejecuta la impresión
└── nameplate1.lbx # Plantilla de etiqueta para b-PAC (colocar en el Escritorio)

---

## ⚙️ Requisitos

- **Windows** (usa `wmic` y b-PAC SDK).
- **Python 3.8+**
- Librerías Python:
  - `comtypes` (para controlar b-PAC desde Python)
- Tener instalada una impresora Brother compatible con **b-PAC SDK**.
- Archivo de plantilla `.lbx` (ejemplo incluido: `nameplate1.lbx`).

---

## 📥 Instalación

1. Clonar este repositorio o descargar los archivos.
2. Instalar dependencias necesarias:

```bash
pip install comtypes
```
3. Colocar el archivo nameplate1.lbx en el Escritorio del servidor.

---

## ▶️ Uso

1. Configuración

En ip_port.txt editar:

192.168.0.X   # IP del servidor
1234          # Puerto

2. Servidor

Ejecutar en la máquina que tiene conectada la impresora Brother:

```bash
python server.py
```
3. Cliente

En cada PC cliente ejecutar:

```bash
python client.py
```
El cliente enviará su número de serie y la instrucción al servidor, que imprimirá la etiqueta automáticamente.

