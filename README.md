# Herramienta-de-listening-en-Google-Sheet

El código proporcionado contiene tres funciones: "sheet", "user" y "buscador".

La función "sheet" tiene como objetivo agregar una nueva fila en una hoja de cálculo específica dentro de un documento de Google Sheets, utilizando el ID del documento. La función recibe un parámetro "n" que es un arreglo de valores que se agregarán en la nueva fila. La función usa la clase "SpreadsheetApp" para abrir el documento, seleccionar la hoja de cálculo específica y obtener la última fila de esa hoja. Luego, se agrega una nueva fila después de la última fila existente y se establecen los valores del parámetro "n" en esta nueva fila. La función finalmente devuelve el resultado de establecer los valores.

La función "user" tiene como objetivo obtener el correo electrónico del usuario activo en la sesión actual, utilizando la clase "Session". Luego, se extrae el nombre del usuario del correo electrónico mediante una función de reemplazo y se convierte la primera letra del nombre en mayúsculas. Finalmente, la función devuelve el nombre del usuario.

La función "buscador" tiene como objetivo buscar tweets en Twitter utilizando una serie de parámetros específicos. La función utiliza la clase "SpreadsheetApp" para obtener una hoja de cálculo específica en un documento de Google Sheets, donde se encuentran almacenados los parámetros de búsqueda. Luego, utiliza la función "formatDate" para formatear las fechas de inicio y fin de la búsqueda en un formato específico. Se obtienen los demás parámetros de búsqueda de la hoja de cálculo.

La función utiliza una serie de condicionales para determinar si incluir o no tweets de retweets y respuestas en la búsqueda. Luego, utiliza la función "encodeURIComponent" para codificar la consulta de búsqueda y construye la URL para realizar la solicitud a la API de Twitter. La función utiliza la clase "UrlFetchApp" para realizar la solicitud y obtener la respuesta. Si la respuesta es un código de error, se muestra una alerta al usuario para informar del problema. Si la respuesta es exitosa, se convierte el contenido de la respuesta en un objeto JSON. La función "buscador" también tiene una línea con un comentario que indica que se debe agregar un token de acceso para poder hacer la petición.
