<h1> Documentación</h1>
 
1-	Pinchar la opción Libro de calificaciones


<img src="https://lh3.googleusercontent.com/pw/AP1GczN43aea-KWNAmXGox0yU9BcLwb8WKb5RKQEXOjgH5yPneOSSUBf4ngt-iA2uP8Oq79tuJLCGQpmPhizFwmLxXoUNjQztTvJ58XCfC74BdHY2MRwTqoKVtcb0X3K5f6QDx7PIRh28i3mvzviHS4rMHig=w598-h259-s-no-gm?authuser=0">



2-	Pinchar la opción descargar libro de calificaciones


<img src="https://lh3.googleusercontent.com/pw/AP1GczNWTDoHsq9B7IgWM1Wg1GKk1Zy9-rnE1oFIxGlvOQisqRk8aRK9UsdU6lvtA0IZu3WcMD4Cx4Iecrfb9l5ADpk209rYUIk6mNwMAZKVEvbgyQSHqfmXea6n8z_rp_ezA6Smr_R4oRtyfkaETRxxrztY=w598-h174-s-no-gm?authuser=0">





3-	En el apartado de Detalles del registro selecciona solo los habilitadores tecnológicos (5 cuestionarios) y en la opción de Tipo de archivo seleccionar la opción Valores separados por coma (.csv)

<img src="https://lh3.googleusercontent.com/pw/AP1GczOtZpySaWJ7KAmWXQ0Sil8armR5MJFnYEr59s6humkndHy-uNwfC29At3a0Q0xKdnm0lOCV6mnuzKiMwY1Gb3JuunuD1pn_FauoAgvA8YXkOjVLZq40aqoJ42w8mUaoEJMt8V9nvO5WJyiqE9QNsKC2=w398-h682-s-no-gm?authuser=0">


📊 Leer Archivos

Este proyecto en Python automatiza el procesamiento de archivos CSV para generar reportes en formato Excel. Se extrae información clave desde los nombres y rutas de los archivos, y se crea un consolidado con un resumen por carrera, facilitando la evaluación del estado de habilitación de estudiantes.

🧩 Funcionalidades

- 🔍 Lectura masiva de archivos .csv desde múltiples carpetas.

- 📁 Extracción automática de:

  - Carrera desde el nombre del archivo.
  
  - Sede desde el nombre de la carpeta donde se encuentra.

- ✂️ Limpieza de columnas innecesarias.

- 📊 Evaluación automática del estado de habilitación basado en las notas de distintos cuestionarios.

- 📄 Generación de archivos Excel individuales por archivo original.

- 📘 Creación de un archivo consolidado con:

  - Todos los datos combinados.

  - Una tabla resumen por carrera.

- ✅ Interfaz gráfica de notificación al finalizar el proceso.

📂 Estructura esperada

```plaintext
Leer archivos/
│
├── archivos/
│   ├── BES/
│   │   ├── archivo1.csv
│   │   └── ...
│   ├── TPC/
│   └── PAP/
│
├── Leer.py
README.md
```

Los archivos CSV deben tener nombres estructurados que permitan identificar la carrera en la posición 5 del nombre (por ejemplo: ..._..._..._..._INNI_...csv → Carrera: INNI).

▶️ Cómo ejecutar

Instala las dependencias:

- pip install pandas openpyxl

Ejecuta el script:

- python Leer.py

Al finalizar, se mostrarán:

- Archivos Excel individuales (uno por CSV).

- Un archivo consolidado.xlsx con toda la información unificada y una tabla resumen.

🛠️ ¿Qué hace cada función?

- listar_archivos_csv()	Busca todos los archivos .csv en una carpeta.

- leer_csv()	Lee un archivo CSV con manejo de errores.

- extraer_carrera()	Toma la carrera desde el nombre del archivo.

- extraer_sede()	Toma la sede desde la ruta de carpeta contenedora.

- procesar_datos()	Limpia columnas, renombra, agrega carrera/sede, y evalúa estado de habilitación.

- guardar_como_excel()	Convierte el DataFrame resultante en un archivo .xlsx.

- procesar_csvs_y_guardar_excel()	Ejecuta todo el flujo por cada archivo CSV.

- unir_excels_y_guardar()	Une todos los Excel en uno solo y genera tabla resumen.

📦 Salida esperada

- Archivos modificados con nombre original mas agregado: *_modificado.xlsx.

- Consolidado: consolidado.xlsx con hoja "Datos" que incluye:

  - Todos los registros.

  - Tabla resumen con conteo de habilitados/no habilitados por carrera (en columna N en archivo xlsx).

📌 Notas adicionales

- La tabla de resumen no incluye la columna Sede.

- Las celdas de resumen tienen bordes aplicados para mejor presentación.

