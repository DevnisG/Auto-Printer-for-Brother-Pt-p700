import socket
import subprocess

# Leer las líneas desde un archivo de texto
with open('ip_port.txt', 'r') as file:
    lines = file.readlines()

# Extraer la dirección IP y el puerto de las líneas
ip = lines[0].strip()
port = int(lines[1])

# Obtén el número de serie de tu PC
serial_number = subprocess.check_output('wmic bios get serialnumber').decode().split('\n')[1].strip()

# Instrucción para la máquina de impresión
instruction = 'Imprimir'

# Concatena el número de serie y la instrucción
data = f"{serial_number};{instruction}"

# Crea un socket TCP/IP
client_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

try:
    # Conecta con la máquina de impresión
    client_socket.connect((ip, port))

    # Envía los datos
    client_socket.sendall(data.encode())

finally:
    # Cierra la conexión
    client_socket.close()
