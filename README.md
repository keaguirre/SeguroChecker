# Automatizador de Consultas de Seguros - README

Este repositorio contiene una herramienta desarrollada en Python que automatiza el proceso de consulta de seguros en el sitio web de la Aach (Asociación de Aseguradoras de Chile) utilizando el navegador Microsoft Edge. La herramienta escanea una lista de RUTs proporcionados en un archivo Excel y recopila información sobre el estado y detalles de los seguros asociados.

## Características

- Escaneo automático de seguros a partir de RUTs proporcionados en un archivo Excel.
- Uso de Selenium para automatizar la interacción con el navegador Microsoft Edge (versión 115.0.1901.188).
- Extracción de información relevante desde la página web de la Aach.
- Actualización y cálculo automatizado de detalles de seguros en un archivo Excel de salida.
- Identificación de seguros vigentes, fechas de inicio y terminación, y tiempo restante o vencido del seguro.

## Requisitos

- Navegador Microsoft Edge versión 115.0.1901.188.
- Archivo ejecutable proporcionado en la última versión de la [release en GitHub](https://github.com/tu_usuario/tu_repositorio/releases).

## Modo de Uso

1. Descarga la última versión del software desde la [sección de releases](https://github.com/tu_usuario/tu_repositorio/releases) en GitHub.

2. Descomprime el archivo descargado. Encontrarás dos archivos: el ejecutable y un archivo Excel llamado `bbdd.xlsx`.

3. Abre el archivo `bbdd.xlsx` utilizando Microsoft Excel u otra aplicación compatible.

4. En la segunda fila de la columna llamada "RUT", pega o escribe los RUTs que deseas procesar. Agrega un RUT por fila.

5. Guarda los cambios en el archivo `bbdd.xlsx` y ciérralo.

6. Ejecuta el archivo ejecutable proporcionado.

7. Durante el proceso, la herramienta abrirá y cerrará automáticamente ventanas del navegador Microsoft Edge en varias ocasiones para realizar las consultas.

8. **Es importante señalar que, dado que el software aún no está completamente finalizado, se sugiere una supervisión durante la ejecución del software.** Mientras las ventanas del navegador Microsoft Edge se abren y los resultados se procesan, se mostrarán detalles en la consola para facilitar la supervisión.

9. Una vez que el proceso esté completo, recibirás un mensaje de finalización en la interfaz de la aplicación, sugiriendo que cierres la aplicación.

## Contribuciones

Si deseas contribuir a este proyecto, puedes realizar un fork del repositorio, hacer tus modificaciones y luego enviar un pull request. ¡Estamos abiertos a colaboraciones!

## Notas

- Esta herramienta utiliza Selenium para automatizar la navegación en la web. Asegúrate de tener instalado el controlador adecuado y que esté configurado para trabajar con Microsoft Edge versión 115.0.1901.188.

- Asegúrate de tener las dependencias de Python instaladas. Puedes instalarlas utilizando el archivo `requirements.txt` proporcionado.

- Esta herramienta fue desarrollada y probada hasta mi conocimiento en septiembre de 2021. Cualquier cambio en la estructura de la página web de la Aach o en la versión del navegador podría requerir ajustes en el código.

¡Esperamos que esta herramienta sea útil para automatizar tus procesos de consulta de seguros! Si tienes algún problema, sugerencia o consulta, no dudes en abrir un issue en este repositorio.

