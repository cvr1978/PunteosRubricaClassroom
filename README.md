# Punteos + Rúbrica — Google Classroom

Herramienta para generar reportes de calificaciones combinando el archivo de
Google Classroom con el desglose de criterios de rúbrica.

**Resultado:** Un archivo Excel con:
- Actividades **sin rúbrica** → columna con el punteo directo
- Actividades **con rúbrica** → columnas por cada criterio + columna de total

---

## Requisitos

```
pip install -r requirements.txt
```

Requiere Python 3.8 o superior.

---

## Flujo de uso (3 pasos)

### Paso 1 — Generar la plantilla de configuración

```bash
python procesar.py plantilla  <archivo_classroom.csv>
```

Ejemplo:
```bash
python procesar.py plantilla  "Calificaciones PROG1 4PAA.csv"
```

Genera `plantilla_rubricas.xlsx` con:
- **Hoja `Configuracion`**: lista de todas las actividades del curso.
  - Llena la columna **"¿Tiene Rúbrica?"** con `SÍ` o `NO`.
  - Para las que tienen rúbrica, escribe el nombre de cada criterio y su punteo máximo.

> **Ejemplo de cómo llenar:**
>
> | Actividad              | Máx. Pts | ¿Tiene Rúbrica? | Criterio 1   | Máx C1 | Criterio 2 | Máx C2 | Criterio 3   | Máx C3 |
> |------------------------|----------|-----------------|--------------|--------|------------|--------|--------------|--------|
> | ACTIVIDAD UNO UNIDAD…  | 15       | SÍ              | Puntualidad  | 5      | Contenido  | 5      | Presentación | 5      |
> | ASPECTOS UNIDAD UNO    | 10       | SÍ              | Presentación | 3      | Dominio    | 4      | Participación| 3      |
> | TAREAS EN CLASE…       | 15       | NO              |              |        |            |        |              |        |

Guarda el archivo después de llenar.

---

### Paso 2 — Generar hojas de ingreso de punteos

```bash
python procesar.py datos  <archivo_classroom.csv>  <plantilla_rubricas.xlsx>
```

Ejemplo:
```bash
python procesar.py datos  "Calificaciones PROG1 4PAA.csv"  plantilla_rubricas.xlsx
```

Agrega una hoja por cada actividad con rúbrica. Cada hoja muestra:
- Nombre y email del estudiante (ya llenado)
- **Celdas amarillas** para ingresar el punteo de cada criterio
- Columna **TOTAL** calculada automáticamente con la suma

Ingresa los punteos de cada criterio para cada estudiante. Guarda el archivo.

---

### Paso 3 — Generar el reporte final

```bash
python procesar.py reporte  <archivo_classroom.csv>  <plantilla_rubricas.xlsx>
```

Ejemplo:
```bash
python procesar.py reporte  "Calificaciones PROG1 4PAA.csv"  plantilla_rubricas.xlsx
```

Genera `reporte_final.xlsx` con:
- Todas las actividades en columnas
- Las actividades con rúbrica muestran cada criterio + el total (en verde)
- Las actividades sin rúbrica muestran el punteo directo de Classroom
- Encabezado con nombre del curso y sección

Para guardar con otro nombre:
```bash
python procesar.py reporte  "Calificaciones PROG1 4PAA.csv"  plantilla_rubricas.xlsx  "Reporte_4PAA_Marzo.xlsx"
```

---

## Un curso por archivo

Classroom exporta un CSV por curso y sección. Para procesar varios cursos,
repite los 3 pasos con cada archivo. Puedes nombrar las plantillas según el
curso para organizarlas:

```
plantilla_rubricas_PROG1_4PAA.xlsx
plantilla_rubricas_PROG1_4PAB.xlsx
plantilla_rubricas_PROG2_5BACHA.xlsx
...
```

---

## Archivo de ejemplo

En la carpeta `datos_ejemplo/` encontrarás:

- `classroom_prog1_4PAA.csv` — archivo de muestra basado en el formato real
  de Google Classroom (PROGRAMACIÓN UNO 4PAA)

Para probar:
```bash
python procesar.py plantilla datos_ejemplo/classroom_prog1_4PAA.csv
# → Llena plantilla_rubricas.xlsx
python procesar.py datos    datos_ejemplo/classroom_prog1_4PAA.csv plantilla_rubricas.xlsx
# → Ingresa punteos en las hojas amarillas
python procesar.py reporte  datos_ejemplo/classroom_prog1_4PAA.csv plantilla_rubricas.xlsx
# → Abre reporte_final.xlsx
```

---

## Estructura del reporte final

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  REPORTE DE CALIFICACIONES — PROGRAMACIÓN UNO    4PAA                       │
├──────────┬──────────┬──────────┬────┬──────────────────────┬─────┬──────────┤
│ Apellido │ Nombre   │ Email    │ %  │  ACTIVIDAD UNO       │ACT2 │ASPECTOS  │
│          │          │          │    │ Puntual│Content│Pres│Total│     │Pres│Dom│Part│Total│
├──────────┼──────────┼──────────┼────┼────────┼───────┼────┼─────┼─────┼────┼───┼────┼─────┤
│ NÁJERA   │ ANDREA   │ ...      │74% │   5    │   4   │  5 │ 14  │  9  │ 3  │ 3 │ 1  │  7  │
│ SALGUERO │ ANGEL    │ ...      │78% │   4    │   5   │  5 │ 14  │  9  │ 2  │ 3 │ 2  │  7  │
└──────────┴──────────┴──────────┴────┴────────┴───────┴────┴─────┴─────┴────┴───┴────┴─────┘
  Columnas azules = sin rúbrica   |   Columnas verdes = con rúbrica (criterios + total)
```
