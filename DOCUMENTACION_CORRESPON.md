# 📖 Documentación Oficial de Funcionalidades y Manual Técnico de Correspon

Bienvenido a la documentación oficial de **Correspon**, la plataforma de automatización de correspondencia masiva, generación de documentos Word/PDF y envío masivo por Microsoft Outlook.

---

## 📋 Tabla de Contenidos
1. [Arquitectura y Estructura Principal](#1-arquitectura-y-estructura-principal)
2. [Desglose por Pasos del Flujo de Trabajo](#2-desglose-por-pasos-del-flujo-de-trabajo)
   - [Paso 1: Datos Excel](#paso-1-datos-excel)
   - [Paso 2: Plantillas Word](#paso-2-plantillas-word)
   - [Paso 3: Generación y Envío](#paso-3-generación-y-envío)
   - [Paso 4: Acciones del Proyecto](#paso-4-acciones-del-proyecto)
3. [Editor Central y Sincronización Bidireccional](#3-editor-central-y-sincronización-bidireccional)
4. [Reglas Condicionales de Párrafos y Plantillas](#4-reglas-condicionales-de-párrafos-y-plantillas)
5. [Matriz de Combinación de Ajustes](#5-matriz-de-combinación-de-ajustes)
6. [Gestión de Configuración en Proyectos (`settings.json`)](#6-gestión-de-configuración-en-proyectos-settingsjson)

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

### Paso 1: Datos Excel
- **Carga de Archivo `.xlsx` / `.xls` / `.xlsm`**: Extrae automáticamente los encabezados de las columnas para convertirlos en fichas de variables mapeables (`{{ Nombre }}`, `{{ Empresa }}`).
- **Editor de Tabla Excel (`#excel-modal`)**: Modal interactivo para visualizar y modificar filas y columnas. Incluye botones para agregar filas/columnas y el interruptor **Campos Obligatorios Habilitados**.
- **Sincronización en Tiempo Real**: Un observador de archivos (*watcher*) detecta modificaciones externas hechas directamente en Excel y refresca los indicadores y datos sin necesidad de recargar la aplicación.
- **Abrir en Excel**: Abre el archivo de origen en Microsoft Excel nativo.

### Paso 2: Plantillas Word
- **Gestión Multi-Plantilla**: Permite cargar una o múltiples plantillas `.docx` en el mismo proyecto.
- **Vista Previa y Selector**: Alterna fácilmente entre las plantillas cargadas para previsualizarlas o editarlas en el área central.
- **Abrir en Word**: Lanza la plantilla seleccionada en la aplicación nativa de Microsoft Word.
- **Reglas de Plantillas Condicionales**: Botón para configurar qué plantilla utilizar automáticamente según los valores de una columna en Excel.

### Paso 3: Generación y Envío
- **Botón Dividido "Generar Documentos Faltantes" / "Opciones de Generación"**:
  - *Generar Faltantes*: Genera únicamente los registros que no contengan la marca de verificación en Excel.
  - *Opciones (Selección de Documentos)*: Despliega un modal con buscador para elegir y regenerar registros específicos.
- **Botón Dividido "Enviar Correos Masivos" / "Selección de Destinatarios"** *(visible únicamente cuando Outlook está habilitado)*:
  - *Enviar Masivo*: Procesa y envía los mensajes mediante Microsoft Outlook a los destinatarios pendientes.
  - *Opciones (Selección de Destinatarios)*: Permite filtrar e indicar manualmente qué destinatarios recibirán el correo.
- **Patrón de Nombre de Archivo**: Define la plantilla de nomenclatura para los archivos generados (ej: `{{ index }}_{{ Nombre }}_comunicado.docx`).

### Paso 4: Acciones del Proyecto
- **Abrir Carpeta del Proyecto**: Abre el Explorador de Archivos de Windows posicionado en la carpeta raíz del proyecto activo.
- **Eliminar Proyecto**: Cierra los procesos de archivos abiertos y elimina de forma segura la carpeta del proyecto y su historial.

---

## 3. Editor Central y Sincronización Bidireccional

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

---

## 4. Reglas Condicionales de Párrafos y Plantillas

> [!TIP]
> Las reglas condicionales permiten personalizar dinámicamente el contenido o la estructura del documento final sin necesidad de duplicar archivos.

1. **Reglas de Párrafos Condicionales**:
   - *Lógica*: `Si [Columna] == [Valor] ➔ Insertar Párrafo Personalizado (Soporta {{ Variables }})`.
   - *Efecto Visual*: La ficha de variable correspondiente en la barra lateral derecha se tiñe de color naranja (`chip-orange`) para identificar rápidamente que contiene reglas de párrafo asociadas.

2. **Reglas de Plantillas Condicionales**:
   - *Lógica*: `Si [Columna] == [Valor] ➔ Usar únicamente la plantilla [Nombre_Plantilla.docx]`.
   - Permite que un mismo proceso masivo alterne entre diferentes formatos de Word según el tipo de cliente o documento.

---

## 5. Matriz de Combinación de Ajustes

La siguiente tabla resume el comportamiento del motor de generación según la combinación de configuraciones seleccionadas:

| Formato de Salida | Reglas de Plantilla | Envío por Outlook | Fuente del Cuerpo de Correo | Comportamiento Resultante |
| :--- | :--- | :--- | :--- | :--- |
| **Word (.docx)** | Habilitadas | Deshabilitado | N/A | Genera únicamente archivos `.docx` aplicando la plantilla correspondiente según la regla condicional del registro. |
| **PDF (.pdf)** | Habilitadas | Deshabilitado | N/A | Renderiza la plantilla Word elegida por la regla, resuelve las etiquetas Jinja2 y exporta únicamente el archivo `.pdf`. |
| **Ambos (Word y PDF)** | Habilitadas | Deshabilitado | N/A | Genera y conserva **ambos formatos** (`.docx` y `.pdf`) resultantes de la plantilla que cumplió la condición. |
| **PDF (.pdf)** | N/A | Habilitado | **Texto personalizado** | Genera los archivos PDF, los adjunta al correo de Outlook y utiliza el texto configurado en la caja de correo como cuerpo. |
| **PDF (.pdf)** | N/A | Habilitado | **Usar plantilla Word** | Convierte la plantilla `.docx` configurada a formato HTML enriquecido (via Mammoth) y la establece como `HTMLBody` en Outlook. |

> [!IMPORTANT]
> Si la opción **Campos Obligatorios Habilitados** está activa en la tabla Excel, cualquier fila que contenga un valor vacío en sus columnas obligatorias será omitida durante la generación masiva, emitiendo un reporte de advertencia al finalizar.

---

## 6. Gestión de Configuración en Proyectos (`settings.json`)

Cada proyecto guarda su estado de forma transparente en la ruta `<Carpeta_Proyecto>/settings.json` con la siguiente estructura:

```json
{
  "output_format": "both",
  "output_filename_pattern": "{{ index }}_{{ Nombre }}_documento.docx",
  "outlook_enabled": true,
  "email_column": "Email",
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
