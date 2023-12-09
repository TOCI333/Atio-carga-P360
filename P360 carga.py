import pyperclip
import time
import subprocess
import shutil
import winreg
import datetime
import pyautogui
import keyboard
import os
import wmi
from fpdf import FPDF

logic=True

while logic:
    print("CARGAR LAS 1000000000 P360 =()")
    print("Qué desea hacer?: ")

    
    def copy():
        #habilita el portapapeles
        keyboard.press_and_release('win+v')
        #copia al portapales
        time.sleep(.5)
        text=["usrOperadorCG", "0p3r4d0rP4$$", "BaseModelo_2019", "CONTROLGAS-P360"]
        for n in text:
            pyperclip.copy(n)
            time.sleep(0.4)
    
    def acces():
        ruta_carpeta = ["C:\Program Files (x86)\ATIO\ControlGAS"]
        for n in ruta_carpeta:
            subprocess.Popen(f'explorer "{n}"')

    def transfer():
        path1="C:\\Program Files (x86)\\ATIO\\ControlGAS"
        os.remove(path1)

        # Ruta del archivo de origen
        ruta_origen = "E:\\P360 carga\\_internal\\dist\\Controlgas\\ControlGAS"

        # Ruta del archivo de destino
        ruta_destino = "C:\\Program Files (x86)\\ATIO"

        # Copiar el archivo de origen al destino
        shutil.copytree(ruta_origen, ruta_destino)


    def delete_registry_folder(folder_path):
        try:
            #Abre el editor de registro de fondo
            subprocess.run("reg.exe")
            # Abrir la clave del registro
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, folder_path, 0, winreg.KEY_ALL_ACCESS)
            
            # Recorrer todas las subclaves y eliminarlas
            index = 0
            while True:
                subkey_name = winreg.EnumKey(key, index)
                subkey_path = f"{folder_path}\\{subkey_name}"
                delete_registry_folder(subkey_path)  # Llamada recursiva para eliminar subcarpetas
                index += 1
        except OSError as e:
            # No hay más subclaves, se puede eliminar la carpeta
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, folder_path)
            print(f"Carpeta eliminada: {folder_path}")
        except Exception as e:
            print(f"Error al eliminar la carpeta: {folder_path}")
            print(f"Mensaje de error: {str(e)}")

    def load_registry_file():
        try:
            # Abrir la clave del registro
            key = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
            #EL SIGUIENTE CODIGO NO SIRVE JAJA
                    # Cargar el archivo .reg
                    #paths_to_load=["E:\\regs_controlg\\x44424E.reg", "E:\\regs_controlg\\x505744.reg", "E:\\regs_controlg\\x554944.reg"]
                    #winreg.("E:\\regs_controlg\\x44424E.reg")
                    # winreg.LoadKey(key, "None", "E:\\regs_controlg\\x505744.reg")
                    # winreg.LoadKey(key, "None", "E:\\regs_controlg\\x554944.reg")
                    # print("Archivo .reg cargado exitosamente.")
                    # subprocess.run(["reg.exe", "import", "E:\\regs_controlg\\x44424E.reg"], capture_output=True, text=True)
                    # subprocess.run(["reg.exe", "import", "E:\\regs_controlg\\x505744.reg"], capture_output=True, text=True)
                    # subprocess.run(["reg.exe", "import", "E:\\regs_controlg\\x554944.reg"], capture_output=True, text=True)
                    # if subprocess.returncode == 0:
                    #     print("Archivo .reg importado exitosamente.")
                    # else:
                    #     print("Error al importar el archivo .reg.")
                    #     print(subprocess.stderr)
            routes=["E:\\P360 carga\\_internal\\dist\\regs_controlg\\x44424E.reg", "E:\\P360 carga\\_internal\\dist\\regs_controlg\\x505744.reg", "E:\\P360 carga\\_internal\\dist\\regs_controlg\\x554944.reg"]
            for n in routes:
                subprocess.run(["reg.exe", "import", n], capture_output=True, text=True)

            

        except Exception as e:
            print(f"Error al cargar el archivo .reg: {str(e)}")

    def open():
        # Ruta al archivo ODBC32.exe
        ruta = ["C:\\Windows\\SysWOW64\\odbcad32.exe", "C:\\Program Files (x86)\\ATIO\\ControlGAS\\UK.exe", "C:\\Program Files (x86)\\ATIO\\ControlGAS\\UDB.exe", "C:\\Program Files (x86)\\ATIO\\ControlGAS\\SG.exe"]

        # Ejecutar el archivo ODBC32.exe
        for n in ruta:
            subprocess.run(n)  
    

    def capturar_pantalla(event):
        time.sleep(1)
        if event.name=='|':
            # Generar un nombre único para la captura de pantalla con el año, mes, dia y hora
            nombre = datetime.datetime.now().strftime("%Y%m%d%H%M%S") + ".png"
            
            # Tomar la captura de pantalla
            captura = pyautogui.screenshot()
            
            # Guardar la captura de pantalla con el nombre único
            captura.save(nombre)
            # Ruta del archivo de origen
            ruta_origen = nombre

            # Ruta del archivo de destino
            ruta_destino = "E:\\P360 carga\\_internal\\dist"

            # Copiar el archivo de origen al destino
            shutil.copy(ruta_origen, ruta_destino)
            
            # Imprimir el nombre de la captura de pantalla guardada
            print("Se genero la captura")


    def create_pdf_from_images(folder_path, output_path):
        pdf = FPDF()
        image_extensions = ['.jpg', '.jpeg', '.png']

        # Obtener la lista de archivos en la carpeta
        files = os.listdir(folder_path)

        # Ordenar los archivos por nombre
        files.sort()

        # Recorrer los archivos en la carpeta
        for file in files:
            # Obtener la extensión del archivo
            _, extension = os.path.splitext(file)

            # Verificar si el archivo es una imagen
            if extension.lower() in image_extensions:
                # Agregar una página al PDF
                pdf.add_page()

                # Obtener la ruta completa de la imagen
                image_path = os.path.join(folder_path, file)

                # Agregar la imagen al PDF
                pdf.image(image_path, x=10, y=10, w=190)

        # Guardar el PDF en el archivo de salida
        pdf.output(output_path)

    def delete_images_captured():
        image_extensions = ['.jpg', '.jpeg', '.png']
        # Obtener la lista de archivos en la carpeta
        files = os.listdir(folder_path)
        # Ordenar los archivos por nombre
        files.sort()
        # Recorrer los archivos en la carpeta
        for file in files:
            # Obtener la extensión del archivo
            _, extension = os.path.splitext(file)

            # Verificar si el archivo es una imagen
            if extension.lower() in image_extensions:
                # Obtener la ruta completa de la imagen
                image_path = os.path.join(folder_path, file)
                #borra la imagen
                os.remove(image_path)

    def delete_images_captured2():
        image_extensions = ['.jpg', '.jpeg', '.png']
        # Obtener la lista de archivos en la carpeta
        files = os.listdir(folder_path2)
        # Ordenar los archivos por nombre
        files.sort()
        # Recorrer los archivos en la carpeta
        for file in files:
            # Obtener la extensión del archivo
            _, extension = os.path.splitext(file)

            # Verificar si el archivo es una imagen
            if extension.lower() in image_extensions:
                # Obtener la ruta completa de la imagen
                image_path = os.path.join(folder_path2, file)
                #borra la imagen
                os.remove(image_path)                

        
    #codigo para obtener el serialnumber de la pc

    # Crear una instancia de la clase WMI
    c = wmi.WMI()

    # Obtener el número de serie del sistema
    system_serial_number = c.Win32_ComputerSystemProduct()[0].IdentifyingNumber
    #print("Número de serie del sistema:", system_serial_number)

    #paths:
    folder_path_key = "Software\ATIO" 
    # path del registro de atio
    folder_path = 'E:\\P360 carga\\_internal\\dist'
    folder_path2 = 'E:\\P360 carga'
    output_path = f'E:\\P360 carga\\{system_serial_number}.pdf'
    #ambos paths son para las imagenes
    

        
                    
    #Menú
    print("1. copiar contraseñas")
    print("2. acceder a la carpeta de ControlGas")
    print("3. actualizar Controlgas")
    print("4. Abrir programas para comenzar la configuración")
    print("5. Cargar configuración para ControlGas")
    print("6. Borrar registros de ControlGas")
    print("7. Tomar capturas y convertir a PDF")
    print("8. SALIR")
    select=int(input("pon un número: "))
    if select==1:
        copy()
    if select==2:
        acces()
    if select==3:
        transfer()
    if select==4:
        open()
    if select==6:
        delete_registry_folder(folder_path_key)
    if select==5:
        load_registry_file()
    if select==7:
        print("Pulsa la tecla | para hacer una captura y cuando termines pulsa esc para generar el pdf")
        keyboard.on_press(capturar_pantalla)
        keyboard.wait('esc')
        #en esta parte se toman las capturas y se crea el pdf
        create_pdf_from_images(folder_path,output_path)
        delete_images_captured()
        delete_images_captured2()

    if select==8:
        logic=False
    

    
    
    
