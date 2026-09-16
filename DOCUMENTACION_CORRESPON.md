# 📖 Documentación Oficial de Funcionalidades y Manual Técnico de Correspon

Bienvenido a la documentación oficial de **Correspon**, la plataforma de automatización de correspondencia masiva, generación de documentos Word/PDF y envío masivo por Microsoft Outlook.

---

## 📋 Tabla de Contenidos
1. [Arquitectura y Estructura Principal](#1-arquitectura-y-estructura-principal)
2. [Desglose por Pasos del Flujo de Trabajo](#2-desglose-por-pasos-del-flujo-de-trabajo)
   - [Paso 1: Datos Excel y Menús Personalizados](#paso-1-datos-excel-y-menús-personalizados)
   - [Paso 2: Plantillas Word y Previsualizaciones PDF](#paso-2-plantillas-word-y-previsualizaciones-pdf)
   - [Paso 3: Generación y Envío Masivo por Outlook](#paso-3-generación-y-envío-masivo-por-outlook)
   - [Paso 4: Acciones del Proyecto](#paso-4-acciones-del-proyecto)
3. [Menús Personalizados y Hoja Oculta `menu_list_app`](#3-menús-personalizados-y-hoja-oculta-menu_list_app)
4. [Múltiples Destinatarios de Correo en Outlook](#4-múltiples-destinatarios-de-correo-en-outlook)
5. [Editor Central y Sincronización Bidireccional](#5-editor-central-y-sincronización-bidireccional)
6. [Reglas Condicionales de Párrafos y Plantillas](#6-reglas-condicionales-de-párrafos-y-plantillas)
7. [Matriz de Combinación de Ajustes](#7-matriz-de-combinación-de-ajustes)
8. [Script Unificado de Instalación y Ejecución (`Iniciar_Correspon.bat`)](#8-script-unificado-de-instalación-y-ejecución-iniciar_corresponbat)
9. [Gestión de Configuración en Proyectos (`settings.json`)](#9-gestión-de-configuración-en-proyectos-settingsjson)

---

## 1. Arquitectura y Estructura Principal

Correspon organiza la pantalla en una arquitectura de **3 Columnas funcionales**:

- **Panel Izquierdo (Pasos 1 al 4)**: Control del flujo secuencial de trabajo, subida de archivos, ejecución de procesos masivos y administración de la carpeta del proyecto.
- **Viewport / Editor Central**: Espacio interactivo que permite alternar entre **Vista Word** (Editor editable WYSIWYG) y **Vista Previa PDF** (visor multi-página nativo).
- **Panel Derecho (Variables y Reglas)**: Mapeo automático de columnas de Excel como fichas arrastrables (*drag-and-drop*) y lista de reglas de párrafos condicionales.

> [!NOTE]
> Cada proyecto de correspondencia es autónomo: almacena sus plantillas, listados de datos Excel, documentos generados y configuraciones en una carpeta física independiente que contiene un archivo `settings.json`.

---

## 2. Desglose por Pasos del Flujo de Trabajo

### Paso 1: Datos Excel y Menús Personalizados
- **Carga de Archivo `.xlsx` / `.xls` / `.xlsm`**: Extrae automáticamente los encabezados de las columnas para convertirlos en fichas de variables mapeables (`{{ Nombre }}`, `{{ Empresa }}`).
- **Editor de Tabla Excel (`#excel-modal`)**: Modal interactivo para visualizar y modificar filas y columnas. Permite:
  - Añadir y eliminar filas y columnas interactivamente.
  - Asignar menús desplegables a columnas o celdas individuales respaldados por la hoja oculta `menu_list_app`.
  - Activar el interruptor **Campos Obligatorios Habilitados**.
- **Sincronización en Tiempo Real**: Un observador de archivos (*watcher*) detecta modificaciones externas hechas directamente en Excel y refresca los indicadores y datos sin recargar la aplicación.
- **Abrir en Excel**: Abre el archivo de origen en Microsoft Excel nativo preservando fórmulas, estilos y hojas ocultas.

### Paso 2: Plantillas Word y Previsualizaciones PDF
- **Gestión Multi-Plantilla**: Permite cargar una o múltiples plantillas `.docx` en el mismo proyecto.
- **Vista Previa y Selector**: Alterna fácilmente entre las plantillas cargadas para previsualizarlas o editarlas en el área central.
- **Eliminación Física Segura y Limpieza de PDF**:
  - Al presionar **Eliminar Plantilla Actual**, la aplicación cierra cualquier proceso de Word que bloquee el archivo mediante COM, elimina el `.docx` físicamente del disco y purga de inmediato cualquier previsualización PDF residual del almacenamiento temporal y de la vista del navegador.
  - Esto garantiza que al reiniciar la aplicación, las plantillas eliminadas no reaparezcan y la pantalla no quede con visores PDF flotantes sin documento activo.
- **Abrir en Word**: Lanza la plantilla seleccionada en la aplicación nativa de Microsoft Word.
- **Reglas de Plantillas Condicionales**: Configuración para seleccionar la plantilla adecuada en base al valor de una columna.

### Paso 3: Generación y Envío Masivo por Outlook
- **Botón Dividido "Generar Documentos Faltantes" / "Opciones de Generación"**:
  - *Generar Faltantes*: Genera únicamente los registros que no contengan la marca de verificación en Excel.
  - *Opciones (Selección de Documentos)*: Modal con buscador para elegir y regenerar registros específicos.
- **Botón Dividido "Enviar Correos Masivos" / "Selección de Destinatarios"** *(visible con Outlook habilitado)*:
  - *Enviar Masivo*: Procesa y envía los mensajes mediante Microsoft Outlook a los destinatarios pendientes.
  - *Opciones (Selección de Destinatarios)*: Permite filtrar e indicar manualmente qué destinatarios recibirán el correo.
- **Múltiples Campos de Correo**: Capacidad de incluir columnas adicionales (ej. Estudiantes, Profesores, Copias) para enviar a todos los involucrados de la fila.
- **Patrón de Nombre de Archivo**: Define la nomenclatura para los archivos generados (ej: `{{ index }}_{{ Nombre }}_comunicado.docx`).

### Paso 4: Acciones del Proyecto
- **Abrir Carpeta del Proyecto**: Abre el Explorador de Archivos de Windows posicionado en la carpeta raíz del proyecto activo.
- **Eliminar Proyecto**: Cierra los procesos de archivos abiertos y elimina de forma segura la carpeta del proyecto y su historial, limpiando además todas las previsualizaciones PDF asociadas.

---

## 3. Menús Personalizados y Hoja Oculta `menu_list_app`

El sistema incluye soporte integral para menús desplegables personalizados integrados directamente en el archivo Excel del proyecto:

### 3.1. Estructura de la Hoja `menu_list_app`
- En el archivo Excel se crea o actualiza automáticamente una hoja con el nombre **`menu_list_app`**.
- La hoja se configura con `sheet_state = "hidden"` en Excel para que no interfiera visualmente con los datos principales, pero se puede visualizar mostrando hojas ocultas en Excel.
- Cada columna de la hoja representa un menú:
  - **Fila 1 (Encabezado)**: Nombre del menú (ej. `materias`, `roles`, `estados`).
  - **Filas 2 en adelante**: Opciones disponibles para dicho menú.

### 3.2. Gestión de Menús (Manual y desde la App)
1. **Desde la App**: En el modal de la tabla Excel, el botón **"Gestionar Menús (menu_list_app)"** abre un gestor visual donde es posible crear nuevos menús, renombrarlos, añadir o eliminar opciones, y guardarlos de inmediato.
2. **Desde Excel**: Puedes abrir el archivo en Excel nativo, mostrar la hoja oculta `menu_list_app`, agregar columnas con nombres de menú y sus opciones debajo, y volver a guardar. La App leerá y sincronizará los cambios automáticamente.

### 3.3. Asignación a Celdas o Columnas
- **Asignación por Columna**: Al presionar el botón `+ Menú` en la cabecera de una columna, puedes vincular toda la columna a un menú específico.
- **Asignación por Celda**: Al hacer clic en el indicador de menú de una celda o mediante el menú contextual, se puede asignar un menú específico a una celda en particular.
- **Comportamiento en la Celda**:
  - Al asignarse un menú, la celda desactiva la edición manual libre (`contenteditable="false"`).
  - En su lugar, se despliega un `<select>` nativo con las opciones del menú.
  - Al guardar la tabla con el botón "Guardar Cambios en Excel", el valor seleccionado se persiste en la hoja de datos principal del Excel.
  - Puedes desasignar el menú en cualquier momento para restaurar la edición manual.

---

## 4. Múltiples Destinatarios de Correo en Outlook

En la ventana de **Ajustes y Configuración de Outlook**, ahora es posible definir múltiples campos de correo procedentes de las columnas del archivo Excel:

- **Campo Principal de Correo**: Columna base para el envío (ej: `Email`).
- **Campos Adicionales de Correo**: Permite agregar tantas columnas de la tabla como se requiera (ejemplo: `Correo Estudiantes`, `Correo Profesores`, `Correo Tutor`).
- **Procesamiento de Destinatarios**:
  - Durante la generación y envío, el sistema recopila los correos de todas las columnas configuradas para la fila actual.
  - Admite múltiples correos separados por comas (`,`) o punto y coma (`;`) dentro de una misma celda.
  - Aplica un filtro de deduplicación automático para no enviar correos duplicados al mismo destinatario.
  - Asigna la lista completa y limpia de destinatarios al campo `mail.To` en Microsoft Outlook.

---

## 5. Editor Central y Sincronización Bidireccional

### Vista Word (Editor Editable)
- **Superficie de Edición (`contenteditable="true"`)**: Permite ajustar textos, modificar títulos o redactar nuevos párrafos directamente en el lienzo.
- **Arrastrar y Soltar (Drag & Drop)**: Arrastra cualquier ficha de variable desde la barra derecha y suéltala en el punto exacto del cursor. Se insertará una etiqueta no editable atómica (`{{ Variable }}`).
- **Resaltar Variables**: Botón para alternar la visibilidad y colores de los distintivos de variables mapeadas.

### Sincronización Bidireccional (Split Button Guardar)
- **`App → Word`**: Toma el contenido HTML editado en la aplicación, lo convierte y lo guarda directamente en el archivo físico `.docx` alojado en el disco.
- **`Word → App`**: Lee la versión actual del archivo `.docx` en el disco y recarga su contenido actualizado al lienzo de la App.

### Vista Previa PDF
- Renderiza el documento final en un visor de PDF multi-página nativo.
- Oculta automáticamente las herramientas de edición y el botón de resaltar variables para presentar el acabado final impreso.
- Al eliminar plantillas o cambiar de proyecto, el visor se limpia por completo garantizando que no queden previsualizaciones huérfanas flotando.

---

## 6. Reglas Condicionales de Párrafos y Plantillas

> [!TIP]
> Las reglas condicionales permiten personalizar dinámicamente el contenido o la estructura del documento final sin necesidad de duplicar archivos.

1. **Reglas de Párrafos Condicionales**:
   - *Lógica*: `Si [Columna] == [Valor] ➔ Insertar Párrafo Personalizado (Soporta {{ Variables }})`.
   - *Efecto Visual*: La ficha de variable correspondiente en la barra lateral derecha se tiñe de color naranja (`chip-orange`) para identificar rápidamente que contiene reglas de párrafo asociadas.

2. **Reglas de Plantillas Condicionales**:
   - *Lógica*: `Si [Columna] == [Valor] ➔ Usar únicamente la plantilla [Nombre_Plantilla.docx]`.
   - Permite que un mismo proceso masivo alterne entre diferentes formatos de Word según el tipo de cliente o documento.

---

## 7. Matriz de Combinación de Ajustes

| Formato de Salida | Reglas de Plantilla | Envío por Outlook | Fuente del Cuerpo de Correo | Comportamiento Resultante |
| :--- | :--- | :--- | :--- | :--- |
| **Word (.docx)** | Habilitadas | Deshabilitado | N/A | Genera únicamente archivos `.docx` aplicando la plantilla correspondiente según la regla condicional del registro. |
| **PDF (.pdf)** | Habilitadas | Deshabilitado | N/A | Renderiza la plantilla Word elegida por la regla, resuelve las etiquetas Jinja2 y exporta únicamente el archivo `.pdf`. |
| **Ambos (Word y PDF)** | Habilitadas | Deshabilitado | N/A | Genera y conserva **ambos formatos** (`.docx` y `.pdf`) resultantes de la plantilla que cumplió la condición. |
| **PDF (.pdf)** | N/A | Habilitado | **Texto personalizado** | Genera los archivos PDF, los adjunta al correo de Outlook (a todos los destinatarios configurados) y utiliza el texto configurado como cuerpo. |
| **PDF (.pdf)** | N/A | Habilitado | **Usar plantilla Word** | Convierte la plantilla `.docx` a formato HTML enriquecido (via Mammoth) y la establece como `HTMLBody` en Outlook. |

> [!IMPORTANT]
> Si la opción **Campos Obligatorios Habilitados** está activa en la tabla Excel, cualquier fila que contenga un valor vacío en sus columnas obligatorias será omitida durante la generación masiva, emitiendo un reporte de advertencia al finalizar.

---

## 8. Script Unificado de Instalación y Ejecución (`Iniciar_Correspon.bat`)

Para simplificar la administración del sistema, todos los scripts `.bat` dispersos fueron consolidados en un único archivo ejecutable: **`Iniciar_Correspon.bat`**.

### Flujo Automatizado de `Iniciar_Correspon.bat`:
1. **Detección de Ejecutable**: Si existe la versión compilada `dist\Correspon\Correspon.exe`, la inicia inmediatamente.
2. **Verificación de Python**: Comprueba si Python 3 está instalado en el sistema. Si falta, ofrece instalación desatendida mediante `winget`.
3. **Comprobación e Instalación de Dependencias**: Verifica las librerías críticas (`fastapi`, `uvicorn`, `openpyxl`, `pandas`, `docxtpl`, `docx`, `mammoth`, `win32com`). Si alguna falta, ejecuta automáticamente la instalación con `pip -r requirements.txt`.
4. **Acceso Directo en el Escritorio**: Crea o actualiza automáticamente el acceso directo `Correspon.lnk` en el Escritorio del usuario con el icono oficial `app_icon.ico`.
5. **Ejecución del Servidor**: Inicia el backend y el entorno gráfico ejecutando `python main.py`.

---

## 9. Gestión de Configuración en Proyectos (`settings.json`)

Cada proyecto guarda su estado de forma transparente en la ruta `<Carpeta_Proyecto>/settings.json` con la siguiente estructura:

```json
{
  "output_format": "both",
  "output_filename_pattern": "{{ index }}_{{ Nombre }}_documento.docx",
  "outlook_enabled": true,
  "email_column": "Email",
  "email_columns": [
    { "column": "Correo Estudiantes" },
    { "column": "Correo Profesores" }
  ],
  "custom_menu_assignments": {
    "columns": {
      "Materia": "materias",
      "Estado": "estados"
    },
    "cells": {
      "0_Observacion": "observaciones_tipo"
    }
  },
  "check_column": "Enviado",
  "doc_check_column": "Doc_Generado",
  "doc_check_value_mode": "timestamp",
  "doc_check_custom_text": "Generado",
  "all_fields_mandatory": false,
  "template_rules_enabled": true,
  "template_rules": [
    {
      "column": "Tipo_Documento",
      "value": "Certificado",
      "template_filename": "Plantilla_Certificado.docx"
    }
  ],
  "send_mode": "display",
  "attachment_type": "pdf",
  "email_subject": "Notificación Oficial - {{ Nombre }}",
  "email_body_source": "template",
  "email_body_template": "C:\\Ruta\\Plantilla_Cuerpo.docx",
  "custom_paragraph_rules": []
}
```
