# Md2PPT - Conversor de Markdown a PowerPoint

## Descripción

**Md2PPT** es una herramienta de consola en Python que convierte un archivo Markdown en una presentación PowerPoint (`.pptx`) usando una plantilla corporativa como base.

---

## Requisitos

| Requisito | Instalación |
|-----------|-------------|
| Python 3.7+ | Configurar ruta en `Settings.json` |
| python-pptx | `pip install python-pptx` |
| Pillow | `pip install Pillow` |
| lxml | `pip install lxml` (normalmente incluida con python-pptx) |
| PowerShell | Incluido en Windows |
| mmdc (opcional) | `npm install -g @mermaid-js/mermaid-cli` |

> `mmdc` es necesario únicamente si el Markdown contiene bloques ` ```mermaid ``` `.

---

## Uso rápido

Haz doble clic en **`Md2PPT.bat`** o ejecútalo desde la terminal:

```
Md2PPT.bat
```

### Comportamiento automático del .bat

1. Busca en `docs\` el **primer `.md`** (por fecha de modificación más antigua).
2. Busca en `docs\` el **primer `.pptx`** como plantilla.
3. Limpia la plantilla con `LimpiarPlantilla.py` (archivo temporal, se borra al terminar).
4. Genera el PowerPoint con el nombre del título `# H1` del Markdown.
5. Si el archivo ya existe, añade sufijo incremental (`_1`, `_2`, …).

### Estructura de carpetas

```
G:\Victor\Md2PPT\
├── Md2PPT.bat                                          ← Lanzador principal
├── Md2PPT.py                                           ← Script de conversión
├── LimpiarPlantilla.py                                 ← Limpiador de layouts de plantilla
├── GetTitle.ps1                                        ← Auxiliar PowerShell para el nombre de salida
├── Settings.json                                       ← Configuración general (empresa, Python, etc.)
├── requirements.txt                                    ← Dependencias Python
├── README.md                                           ← Esta guía
├── PROMPT GENERACION MD PRESENTACION.md                ← Ejemplo con todas las opciones de sintaxis
└── docs\
    ├── ejemplo_entrada.md                              ← Ejemplo de Markdown de entrada
    └── ejemplo_plantilla.pptx                          ← Ejemplo de plantilla
```

### Configuración (`Settings.json`)

```json
{
  "General": {
    "CompanyName": "Empresa"
  },
  "Python": {
    "InterpreterPath": "E:\\Python\\Python38-32\\python.exe",
    "SearchPaths": [],
    "TestFramework": "unittest"
  }
}
```

| Campo | Descripción |
|-------|-------------|
| `General.CompanyName` | Nombre de empresa que aparece en el pie de cada slide de contenido (`"Empresa  /  Título"`). Si se deja vacío, el pie muestra solo el título. |
| `Python.InterpreterPath` | Ruta absoluta al ejecutable de Python. |

### Uso directo del script Python

```
python Md2PPT.py <entrada.md> <plantilla.pptx> <salida.pptx>
```

---

## Formato del Markdown de entrada

Se puede usar el fichero de "PROMPT GENERACION MD PRESENTACION.md" en la raíz del proyecto para que una IA genere un md con la referencia de sintaxis correcta. Ver también `ejemplo_entrada.md` de docs\ejemplos.

### Referencia de sintaxis

| Sintaxis Markdown | Resultado en PowerPoint |
|---|---|
| `# Título` | **Portada** (layout 0): título, fecha mes/año y diagrama Mermaid opcional |
| `## Sección` | **Portada de sección** (layout 1): título en mayúsculas + número "01", "02"… y una slide de contenido vacía |
| `### Subtítulo` | En la **portada de sección**: se acumula como subtítulo. En el **contenido**: crea una nueva slide de contenido con ese texto como subtítulo |
| Texto normal | Párrafo sin viñeta en la slide de contenido actual |
| `* texto` o `- texto` | Viñeta (bullet point) en la slide de contenido actual |
| ` ```mermaid … ``` ` | Primer bloque → imagen en la portada. Bloques siguientes → imagen en la slide de contenido actual |
| `**texto**` o `__texto__` | **Negrita** |
| `*texto*` o `_texto_` | *Cursiva* |
| `~~texto~~` | ~~Tachado~~ |
| `` `texto` `` | `Código` (fuente Courier New) |

### Reglas de flujo

- Todo el contenido (texto, bullets, diagramas) va a la **slide de contenido activa**.
- Cada `## Sección` crea una portada de sección **y** una slide de contenido inicial.
- Cada `### Subtítulo` crea una **nueva slide de contenido** con ese subtítulo. Si aparece antes del primer contenido de una sección, también se añade a la portada de sección.
- Los bullets y textos van siempre a la slide de contenido activa en ese momento, **en el orden del Markdown**.

---

## Preparación de la plantilla PowerPoint

### Por qué es necesario prepararla

En PowerPoint, cada slide se basa en un **layout** del Patrón de diapositivas. Aunque elimines slides visibles, sus layouts permanecen ocultos. Una plantilla corporativa puede tener 30-40 layouts aunque solo uses 4 slides. `Md2PPT.py` usa los layouts **por índice** (0, 1, 2, 3), por lo que la plantilla debe tener exactamente 4 layouts en el orden correcto.

Además, dentro de cada layout los placeholders se identifican por un número de índice (`idx`) que **varía entre plantillas**: lo que en una plantilla es el cuerpo con `idx 15`, en otra puede ser `idx 11`. `LimpiarPlantilla.py` resuelve esto anotando cada placeholder con un nombre canónico (`MD2PPT_TITLE`, `MD2PPT_BODY`, etc.) basándose en su tamaño y posición, de forma que `Md2PPT.py` pueda encontrarlos de forma dinámica sin depender de números fijos.

### Paso 1 — Preparar las slides en PowerPoint

1. Abre la plantilla en PowerPoint.
2. Elimina todas las diapositivas **excepto las 4 que necesitas**, en este orden:
   - Slide 1 → Portada
   - Slide 2 → Portada de sección (separata)
   - Slide 3 → Contenido con texto
   - Slide 4 → Cierre
3. Guarda la plantilla.

### Paso 2 — Eliminar layouts huérfanos y anotar placeholders

El archivo `Md2PPT.bat` ejecuta automáticamente `LimpiarPlantilla.py`. Si quieres hacerlo manualmente:

```
python LimpiarPlantilla.py docs\plantilla_original.pptx docs\plantilla_limpia.pptx
```

El script realiza dos tareas:

1. **Elimina layouts huérfanos**: detecta qué layouts usan las 4 slides y elimina el resto.
2. **Anota los placeholders**: asigna nombres canónicos (`MD2PPT_TITLE`, `MD2PPT_BODY`, `MD2PPT_SUBTITLE`, etc.) a cada placeholder del layout, distinguiendo los que tienen el mismo tipo por tamaño y posición. Esto permite que `Md2PPT.py` los encuentre correctamente en cualquier plantilla.

La salida muestra el orden resultante y las anotaciones realizadas:

```
Layout 0: 'PORTADA - V - Pruno'
Layout 1: 'SEPARATA - Principal'
Layout 2: 'CONTENIDO - Texto - Ceramico'
  idx=0  -> MD2PPT_TITLE
  idx=14 -> MD2PPT_SUBTITLE
  idx=15 -> MD2PPT_BODY
  idx=16 -> MD2PPT_ANTETITLE
  idx=17 -> MD2PPT_FOOTER
  idx=18 -> MD2PPT_SLIDE_NUMBER
Layout 3: 'CIERRE - Pruno'
```

### Paso 3 — Usar la plantilla limpia

Coloca `plantilla_limpia.pptx` en `docs\` y el `.bat` la usará automáticamente.

> Si la plantilla tiene menos de 4 layouts, `Md2PPT.py` mostrará un error indicando que hay que ejecutar `LimpiarPlantilla.py` primero.

### Inspeccionar layouts con el modo debug

```
python Md2PPT.py docs\entrada.md docs\plantilla.pptx docs\salida.pptx --debug
```

Muestra todos los layouts disponibles y los placeholders de cada slide existente en la plantilla.

---

## Layouts esperados y sus placeholders

`Md2PPT.py` identifica los placeholders de cada layout por los **nombres canónicos** anotados por `LimpiarPlantilla.py`, con fallback a índices estándar si la plantilla no fue anotada.

| Índice | Uso | Placeholders clave |
|--------|-----|--------------------|
| 0 | Portada | `MD2PPT_TITLE` título, `MD2PPT_SUBTITLE` subtítulo, `MD2PPT_DATE` fecha, `MD2PPT_PICTURE` imagen Mermaid |
| 1 | Portada de sección | `MD2PPT_TITLE` título (mayúsculas), `MD2PPT_SUBTITLE` subtítulos H3, número de sección (`idx 10`) |
| 2 | Contenido | `MD2PPT_TITLE` título sección, `MD2PPT_SUBTITLE` subtítulo H3, `MD2PPT_BODY` cuerpo, `MD2PPT_FOOTER` pie, `MD2PPT_SLIDE_NUMBER` nº página |
| 3 | Cierre | Sin modificar |

---

## Nombre del archivo de salida

- Se extrae del título `# H1` del Markdown.
- Se eliminan caracteres no válidos para nombres de archivo en Windows.
- Si ya existe un archivo con ese nombre, se añade `_1`, `_2`, etc.
- Si no hay `# H1`, se usa `resultado.pptx`.

---

## Ayuda

```
python Md2PPT.py --help
```
