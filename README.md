Automatización de Copia de Datos entre Archivos Excel
Este script de Python facilita la transferencia de datos entre múltiples archivos y hojas de Excel. Permite copiar información de una hoja de origen a una hoja de destino, con opciones para mapear columnas por nombre o copiar por posición.
Características
 * Copia de Múltiples Archivos: Configura fácilmente varias tareas de copia entre diferentes pares de archivos Excel.
 * Mapeo de Columnas por Nombre (Opcional): Define explícitamente qué columna del origen corresponde a qué columna del destino usando sus nombres de encabezado. Esto asegura una copia precisa incluso si el orden de las columnas cambia.
 * Copia de Columnas por Posición: Si no se especifica un mapeo, el script copia las columnas basándose en su orden posicional (comportamiento por defecto).
 * Manejo de Encabezados: Permite especificar la fila donde se encuentran los encabezados en el archivo de origen y la fila donde comenzarán los datos en el archivo de destino.
 * Verificación de Archivos: Comprueba la existencia de los archivos de origen y destino antes de procesar, evitando errores.
Requisitos
Asegúrate de tener instaladas las siguientes librerías de Python:
 * pandas
 * openpyxl
Puedes instalarlas usando pip:
pip install pandas openpyxl

Uso
 * Configuración:
   Abre el script excel_automation.py (o como lo hayas nombrado).
   Modifica la lista CONFIGURACION_TAREAS para definir tus tareas de copia. Cada diccionario en la lista representa una tarea y debe incluir:
   * "ARCHIVO_ORIGEN": Ruta completa al archivo Excel de origen.
   * "HOJA_ORIGEN": Nombre de la hoja dentro del archivo de origen.
   * "FILA_ENCABEZADOS_ORIGEN": La fila (índice 0-basado) donde se encuentran los encabezados en el archivo de origen. Por ejemplo, si están en la Fila 6 de Excel, usa 5.
   * "ARCHIVO_DESTINO": Ruta completa al archivo Excel de destino.
   * "HOJA_DESTINO": Nombre de la hoja dentro del archivo de destino.
   * "FILA_INICIO_DATOS_DESTINO": La fila (índice 0-basado) donde deseas que comiencen los datos en el archivo de destino. Por ejemplo, si los datos deben comenzar en la Fila 12 de Excel, usa 11.
   * "COLUMNAS_A_MAPEARD": (Opcional) Un diccionario para mapear columnas. La clave es el nombre de la columna en el origen y el valor es el nombre de la columna correspondiente en el destino. Si se omite o está vacío ({}), la copia se realizará por posición.
     "COLUMNAS_A_MAPEARD": {
    "ColumnaOrigen1": "ColumnaDestino1",
    "OtraColOrigen": "ColumnaDestinoXYZ",
    # ... más mapeos
}

    Importante: Los archivos y hojas de destino DEBEN existir previamente y tener los encabezados definidos en la fila correcta (actualmente configurado para la Fila 11 de Excel).
 * Ejecutar el Script:
   Guarda los cambios y ejecuta el script desde tu terminal:
   python tu_script_de_automatizacion.py

Consideraciones Importantes
 * Indexación: Python y Pandas usan indexación 0-basada para filas, mientras que Excel usa 1-basada. Ajusta FILA_ENCABEZADOS_ORIGEN y FILA_INICIO_DATOS_DESTINO en consecuencia. Por ejemplo, la Fila 1 de Excel es 0 en Python.
 * Archivos de Destino: El script no crea los archivos o hojas de destino. Asegúrate de que existan y tengan la estructura de encabezados correcta antes de ejecutar.
 * Archivos Abiertos: Asegúrate de que los archivos de destino no estén abiertos en Excel mientras el script se está ejecutando, ya que esto podría causar errores de escritura.
 * Encabezados: El script espera que los encabezados del archivo de destino estén en la Fila 11 de Excel (índice 10 en Pandas/Python) para leerlos correctamente y alinear las columnas. Si tus encabezados de destino están en una fila diferente, deberás ajustar manualmente la línea header=10 en el código (dentro del try principal).
Solución de Problemas
 * FileNotFoundError: Verifica que las rutas de ARCHIVO_ORIGEN y ARCHIVO_DESTINO sean correctas y que los archivos existan.
 * KeyError (relacionado con hojas): Asegúrate de que los nombres de HOJA_ORIGEN y HOJA_DESTINO coincidan exactamente con los nombres de las hojas en tus archivos Excel (sensible a mayúsculas y minúsculas).
 * Errores de escritura (BadZipFile, etc.): Lo más común es que el archivo de destino esté abierto en Excel. Cierra el archivo e intenta de nuevo.
 * Datos incorrectos o faltantes:
   * Revisa FILA_ENCABEZADOS_ORIGEN para asegurarte de que Pandas lea los encabezados correctos.
   * Si usas mapeo, verifica que los nombres de columna en COLUMNAS_A_MAPEARD coincidan exactamente con los encabezados del origen y destino.
   * Si no usas mapeo, confirma que la posición de las columnas en origen y destino sea la esperada.
