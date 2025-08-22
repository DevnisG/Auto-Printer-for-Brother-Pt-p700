import socket
import comtypes.client
import os

# Dirección IP y puerto en los que escuchará la máquina de impresión
ip = '0.0.0.0'  # Escucha en todas las interfaces de red
port = 1234  # Puerto para la comunicación

while True:
    try:
        # Crea un socket TCP/IP
        server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

        # Vincula el socket a la dirección IP y puerto
        server_socket.bind((ip, port))

        # Escucha las conexiones entrantes
        server_socket.listen(1)

        print("Esperando conexiones...")

        # Acepta una conexión entrante
        client_socket, address = server_socket.accept()

        print("Conexión establecida desde:", address)

        # Recibe los datos enviados por la PC
        data = client_socket.recv(1024).decode()

        # Separa el número de serie y la instrucción
        serial_number, instruction = data.split(';')

        # Realiza la impresión con los datos recibidos
        print("Número de serie:", serial_number)
        print("Instrucción:", instruction)

        # Create an instance of the b-PAC object
        bpac = comtypes.client.CreateObject('bpac.Document')

        # Obtener la ruta absoluta del archivo en el escritorio
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        label_file = 'nameplate1.lbx'
        label_path = os.path.join(desktop_path, label_file)

        if bpac.Open(label_path):
            objtext = bpac.GetObject('objName')

            # Chequear si el objeto de texto es válido
            if objtext is not None:
                objtext.Text = serial_number

            bpac.StartPrint(1, 0)
            bpac.PrintOut(1, 0)
            bpac.EndPrint(0, 0)
            bpac.Close(0, 0)
        else:
            print("Error al abrir el archivo de plantilla de etiqueta.")

    except Exception as e:
        print("Error:", str(e))

    finally:
        # Cierra la conexión del cliente
        client_socket.close()
        # Cierra el socket del servidor
        server_socket.close()