# Asistente local de archivos para iPhone 13 Pro

Esta guía resume el sistema mostrado en las imágenes: un asistente 100% local para organizar, clasificar, buscar, respaldar y reportar archivos desde la app **Archivos** y **Atajos** de iOS, sin internet, sin nube y sin servicios externos.

## Objetivo

El sistema permite administrar documentos directamente en el iPhone mediante automatizaciones locales. El usuario conserva el control de sus archivos, puede consultar inventarios, detectar duplicados, generar respaldos comprimidos y producir reportes de estado.

## Estructura de carpetas

La carpeta raíz sugerida es `Asistente_Local_Luis` e incluye las siguientes subcarpetas:

| Carpeta | Uso |
| --- | --- |
| `01_Documentos` | PDF, Word, Excel, certificados y documentos generales. |
| `02_Imagenes` | Fotos personales, escaneos e imágenes de trabajo. |
| `03_Videos` | Videos personales o laborales. |
| `04_Audio` | Grabaciones, música y notas de voz. |
| `05_Comprimidos` | Archivos ZIP, RAR, 7z y otros paquetes. |
| `06_Respaldos` | Respaldos automáticos comprimidos. |
| `07_Reportes` | Reportes generados por el sistema. |
| `08_Pendientes` | Archivos pendientes de organizar o revisar. |
| `09_BaseDatos` | Base de datos e inventarios del sistema. |
| `10_Plantillas` | Plantillas reutilizables para reportes. |

## Base de datos local

El inventario principal se guarda como CSV en `Asistente_Local_Luis/09_BaseDatos/inventario.csv`. Cada fila representa un archivo registrado por el sistema.

Columnas recomendadas:

- `Fecha`: fecha de registro o última actualización.
- `Nombre`: nombre del archivo.
- `Tipo`: tipo o extensión normalizada.
- `Tamaño`: tamaño legible del archivo.
- `Ubicación`: ruta local dentro de la estructura del asistente.

Ejemplo:

```csv
Fecha,Nombre,Tipo,Tamaño,Ubicación
2026-08-04,Certificado.pdf,PDF,2.1 MB,/01_Documentos/Certificados
2026-08-04,Contrato.docx,Word,340 KB,/01_Documentos/Word
2026-08-04,Presupuesto.xlsx,Excel,120 KB,/01_Documentos/Excel
```

## Atajos principales

### 1. Organizar Archivos

Escanea una carpeta, obtiene el contenido de cada archivo, detecta su extensión y lo mueve automáticamente a la carpeta correspondiente.

Acciones principales:

1. Obtener contenido de carpeta.
2. Repetir con cada archivo.
3. Obtener nombre y extensión.
4. Evaluar el tipo de archivo.
5. Mover el archivo a la carpeta asignada.

### 2. Actualizar Inventario

Recorre todas las carpetas del sistema y registra cada archivo en `inventario.csv` con su información básica.

Acciones principales:

1. Obtener todas las carpetas.
2. Obtener contenido.
3. Obtener nombre, tipo, tamaño y fecha.
4. Agregar una fila al CSV.

### 3. Buscar Duplicados

Compara nombres y tamaños para identificar archivos repetidos y guarda un reporte en `07_Reportes`.

Acciones principales:

1. Obtener todos los archivos.
2. Comparar nombres.
3. Comparar tamaños.
4. Registrar duplicados.
5. Generar `duplicados.txt`.

### 4. Respaldo Diario

Crea un respaldo completo de carpetas importantes y lo comprime en un ZIP dentro de `06_Respaldos`.

Acciones principales:

1. Seleccionar carpetas.
2. Copiar archivos.
3. Crear archivo ZIP.
4. Guardar el ZIP en `06_Respaldos`.

### 5. Buscar Archivo

Permite buscar documentos por nombre en las carpetas importantes y muestra ubicación, tamaño y tipo.

Acciones principales:

1. Pedir nombre a buscar.
2. Buscar en carpetas configuradas.
3. Mostrar resultados y detalles.

### 6. Estado del Sistema

Genera un reporte general con estadísticas del sistema y lo guarda como RTF en `07_Reportes`.

Acciones principales:

1. Contar archivos por tipo.
2. Calcular tamaños.
3. Crear documento RTF.
4. Guardar `Estado.rtf`.

## Tipos de archivos soportados

- Documentos: PDF, DOCX, XLSX, TXT, RTF.
- Imágenes: JPG, PNG.
- Videos: MP4, MOV.
- Audio: MP3, M4A.
- Comprimidos: ZIP, RAR, 7z.
- Datos: CSV y JSON.

## Flujo general del sistema

1. Escanear archivos desde la app Archivos.
2. Clasificar automáticamente por tipo.
3. Organizar en carpetas.
4. Registrar cada archivo en la base de datos CSV.
5. Generar reportes y estadísticas.
6. Crear respaldos comprimidos.
7. Consultar o buscar documentos cuando sea necesario.

## Mejora de automatizaciones y redireccionamiento automático

El asistente puede reforzarse con reglas de automatización que clasifiquen, redirijan y validen archivos sin intervención manual. La idea es que cada archivo nuevo pase por un flujo estándar: detección, clasificación, redireccionamiento, registro y verificación.

### Reglas de redireccionamiento automático

Configurar el atajo **Organizar Archivos** para evaluar cada archivo por extensión, nombre y ubicación de origen. Según el resultado, el archivo debe moverse automáticamente a la carpeta correspondiente.

| Condición detectada | Destino automático | Acción adicional |
| --- | --- | --- |
| `.pdf`, `.docx`, `.xlsx`, `.txt`, `.rtf` | `01_Documentos` | Registrar tipo documental en el inventario. |
| `.jpg`, `.jpeg`, `.png`, `.heic` | `02_Imagenes` | Conservar fecha de creación si está disponible. |
| `.mp4`, `.mov` | `03_Videos` | Registrar tamaño para reportes de almacenamiento. |
| `.mp3`, `.m4a`, `.wav` | `04_Audio` | Clasificar como audio o nota de voz. |
| `.zip`, `.rar`, `.7z` | `05_Comprimidos` | Marcar como paquete comprimido. |
| `.csv`, `.json` | `09_BaseDatos` | Evitar sobrescribir el inventario principal. |
| Tipo desconocido | `08_Pendientes` | Marcar para revisión manual. |

### Flujo recomendado de automatización

1. Detectar archivos nuevos o modificados en la carpeta de entrada.
2. Obtener nombre, extensión, tamaño, fecha y ruta original.
3. Normalizar la extensión a minúsculas para evitar duplicados de reglas.
4. Aplicar la tabla de redireccionamiento automático.
5. Mover el archivo a la carpeta destino.
6. Agregar o actualizar la entrada en `09_BaseDatos/inventario.csv`.
7. Registrar archivos no reconocidos en un reporte de pendientes.
8. Mostrar un resumen con archivos procesados, movidos, omitidos y pendientes.

### Validaciones antes de mover archivos

Antes de redirigir un archivo, el atajo debe comprobar:

- Que el archivo todavía exista en la ubicación original.
- Que la carpeta destino exista; si no existe, crearla automáticamente.
- Que no haya otro archivo con el mismo nombre en el destino.
- Que el inventario no quede duplicado.
- Que los archivos críticos, como `inventario.csv`, no se sobrescriban accidentalmente.

Cuando exista un conflicto de nombre, usar una estrategia consistente, por ejemplo agregar fecha y hora al nombre del archivo:

```text
Contrato.pdf
Contrato_2026-08-14_1530.pdf
```

### Automatizaciones programadas sugeridas

| Automatización | Frecuencia recomendada | Resultado esperado |
| --- | --- | --- |
| Organizar archivos | Al guardar o importar archivos | Archivos redirigidos a su carpeta correcta. |
| Actualizar inventario | Diario | CSV actualizado con rutas y metadatos. |
| Buscar duplicados | Semanal | Reporte de posibles duplicados. |
| Respaldo comprimido | Diario o semanal | ZIP guardado en `06_Respaldos`. |
| Estado del sistema | Semanal | Reporte RTF con estadísticas generales. |

### Criterio de pruebas y comprobaciones

Ejecutar pruebas automatizadas y validaciones de código en cada cambio que afecte la lógica o el comportamiento de la aplicación. No es necesario ejecutar estas comprobaciones cuando los cambios se limiten exclusivamente a comentarios, documentación o contenido no funcional.

Ejecutar pruebas y comprobaciones de código para cualquier cambio funcional. Omitirlas cuando los cambios afecten únicamente a comentarios o documentación.

## Seguridad y privacidad

El diseño funciona completamente en local:

- No requiere internet.
- No utiliza nube.
- No comparte información con servicios externos.
- Puede protegerse con Face ID o código.
- Mantiene los respaldos dentro del dispositivo, bajo control del usuario.

