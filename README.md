# Punteos + Rúbrica — Google Classroom

Herramienta para generar reportes de calificaciones combinando el archivo de
Google Classroom con el desglose de criterios de rúbrica.

**Resultado:** Una hoja de Google Sheets con:
- Actividades **sin rúbrica** → columna con el punteo directo (azul)
- Actividades **con rúbrica** → columnas por cada criterio + columna de total (verde)

---

## Cómo usar (versión Google Sheets — sin instalar nada)

### 1. Preparar el script

1. Abre **[Google Sheets](https://sheets.google.com)** y crea una hoja nueva.
2. En el menú ve a **Extensiones → Apps Script**.
3. Borra todo el código que aparece por defecto.
4. Copia el contenido del archivo **`Code.gs`** de este repositorio y pégalo.
5. Haz clic en **Guardar** (ícono de disquete o Ctrl+S).
6. Cierra la pestaña de Apps Script y **recarga** la hoja de Google Sheets.
7. Aparecerá un nuevo menú llamado **"Rúbricas Classroom"** en la barra de menús.

> La primera vez que ejecutes cualquier paso, Google te pedirá autorizar el
> script. Haz clic en **"Revisar permisos"** → selecciona tu cuenta → **"Avanzado"**
> → **"Ir a [nombre del proyecto] (no seguro)"** → **"Permitir"**.

---

### Paso 1 — Importar CSV de Classroom

1. En Google Classroom abre el curso, ve a **Calificaciones** y descarga el CSV.
2. Abre el archivo CSV con el **Bloc de notas** (clic derecho → "Abrir con" → Bloc de notas).
3. Selecciona todo (Ctrl+A), copia (Ctrl+C).
4. En la hoja de Google Sheets, ve al menú **Rúbricas Classroom → Paso 1: Importar CSV**.
5. Pega el contenido en el cuadro que aparece y haz clic en **Importar**.

---

### Paso 2 — Generar Plantilla de Rúbricas

Ve al menú **Rúbricas Classroom → Paso 2: Generar Plantilla**.

Se crea la hoja **PLANTILLA** con todas las actividades del curso. Para cada actividad:

- En la columna **"¿Tiene Rúbrica?"** selecciona **SI** o **NO**.
- Para las que tienen rúbrica, rellena los criterios en las **celdas amarillas**:

| Actividad             | Máx. Pts | ¿Tiene Rúbrica? | Criterio 1  | Máx C1 | Criterio 2 | Máx C2 | Criterio 3   | Máx C3 |
|-----------------------|----------|-----------------|-------------|--------|------------|--------|--------------|--------|
| ACTIVIDAD UNO…        | 15       | SI              | Puntualidad | 5      | Contenido  | 5      | Presentación | 5      |
| ASPECTOS UNIDAD UNO   | 10       | SI              | Presentación| 3      | Dominio    | 4      | Participación| 3      |
| TAREAS EN CLASE…      | 15       | NO              |             |        |            |        |              |        |

---

### Paso 3 — Crear Hojas de Punteos

Ve al menú **Rúbricas Classroom → Paso 3: Crear Hojas de Punteos**.

Se crea una hoja **RUB_[nombre de la actividad]** por cada actividad marcada con SI.
Cada hoja tiene:
- Lista de estudiantes ya cargada
- **Celdas amarillas** para ingresar el punteo de cada criterio
- Columna **TOTAL** que se calcula automáticamente

Ingresa los puntajes y cuando termines pasa al paso 4.

---

### Paso 4 — Generar Reporte Final

Ve al menú **Rúbricas Classroom → Paso 4: Generar Reporte Final**.

Se crea la hoja **REPORTE** con todo consolidado:

```
┌──────────┬──────────┬──────────┬──────┬──────────────────────┬──────┬─────────────────────┐
│ Apellido │ Nombre   │ Email    │  %   │  ACTIVIDAD UNO               │  ASPECTOS UNO        │
│          │          │          │      │ Puntual│Cont.│Pres.│TOTAL│Directo│Pres.│Dom.│Part.│TOTAL│
├──────────┼──────────┼──────────┼──────┼────────┼─────┼─────┼─────┼───────┼─────┼────┼─────┼─────┤
│ NÁJERA   │ ANDREA   │ ...      │ 74%  │   5    │  4  │  5  │ 14  │   9   │  3  │ 3  │  1  │  7  │
└──────────┴──────────┴──────────┴──────┴────────┴─────┴─────┴─────┴───────┴─────┴────┴─────┴─────┘
  Columnas azules = puntaje directo   |   Columnas verdes = criterios de rúbrica + total
```

---

## Un curso por hoja

Classroom exporta un CSV por curso/sección. Para manejar varios cursos,
crea una hoja de Google Sheets diferente para cada uno y repite los 4 pasos.

---

## Versión Python (alternativa sin Google Sheets)

Si prefieres trabajar localmente con Excel, el archivo `procesar.py` hace
lo mismo desde la terminal. Requiere Python 3.8+ y las librerías de
`requirements.txt`.

```bash
pip install -r requirements.txt
python procesar.py plantilla  datos_ejemplo/classroom_prog1_4PAA.csv
python procesar.py datos      datos_ejemplo/classroom_prog1_4PAA.csv  plantilla_rubricas.xlsx
python procesar.py reporte    datos_ejemplo/classroom_prog1_4PAA.csv  plantilla_rubricas.xlsx
```
