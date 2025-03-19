import win32com.client  # Importa la librería para interactuar con Microsoft Word
import os  # Importa el módulo para manejar rutas y archivos

try:
    # Solicita al usuario la ruta de la carpeta donde está el archivo
    usuario_carpeta = input("Ingresa la ruta de la carpeta donde se encuentra el archivo: ")
    # Solicita al usuario el nombre del archivo a convertir
    documento_usuario = input("Ingresa el nombre del archivo a convertir (incluyendo la extensión .docx o .doc): ")

    # Verifica si la carpeta ingresada por el usuario existe
    if os.path.isdir(usuario_carpeta):
        # Une la carpeta y el archivo para obtener la ruta completa del documento de Word
        ruta_completa = os.path.join(usuario_carpeta, documento_usuario)

        # Inicia una instancia oculta de Microsoft Word
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # Evita que se abra la ventana de Word

        # Abre el documento de Word en la ruta especificada
        doc_word = word.Documents.Open(ruta_completa)

        # Define la ruta donde se guardará el archivo PDF
        ruta_pdf = os.path.join(usuario_carpeta, "documento.pdf")

        # Guarda el documento en formato PDF (FileFormat = 17)
        doc_word.SaveAs(ruta_pdf, FileFormat=17)

        # Cierra el documento y la aplicación de Word
        doc_word.Close()
        word.Quit()

        print(f"Documento convertido exitosamente y guardado en: {ruta_pdf}")
    else:
        print("Carpeta o documento no encontrado.")

except Exception as error:
    # Captura y muestra cualquier error que ocurra en el programa
    print(f"Error en el programa: {error}.")
