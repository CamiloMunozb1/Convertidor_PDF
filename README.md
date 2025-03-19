# Convertidor de Word a PDF

Este es un script en Python que permite convertir documentos de **Microsoft Word** (`.docx` o `.doc`) a **PDF** de manera automática.

## 🚀 Características

- Convierte archivos de **Word a PDF** sin necesidad de abrir Word manualmente.
- Utiliza `win32com.client` para interactuar con **Microsoft Word**.
- Permite al usuario ingresar la ruta del archivo a convertir.

## 🛠️ Requisitos

- Tener instalado **Python 3.x**.
- Tener **Microsoft Word** instalado (necesario para que `win32com.client` funcione).
- La biblioteca **pywin32** instalada.

## 📦 Instalación

1. **Clonar el repositorio**:
   ```sh
   git clone https://github.com/TU_USUARIO/TU_REPOSITORIO.git
   cd TU_REPOSITORIO
   
2. **Instalar las dependencias**:
 
        pip install pywin32
   
## Uso
Ejecuta el script y sigue las instrucciones:

        python convertidor.py

El programa pedirá que ingreses la ruta de la carpeta y el archivo a convertir. Luego, generará un archivo PDF en la misma carpeta.

## Notas
-Asegúrate de que el archivo Word que deseas convertir esté cerrado antes de ejecutar el script.-
-Si tienes problemas con permisos, intenta ejecutar Python como administrador.

## Autor
Juan Camilo Muñoz
