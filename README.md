# Punteos + Rúbrica — Google Classroom

Herramienta que se conecta **directamente a Google Classroom** y extrae
automáticamente todos los punteos con el desglose de criterios de rúbrica.

**Resultado:** Una hoja REPORTE en Google Sheets con:
- Actividades **sin rúbrica** → columna con punteo directo (azul)
- Actividades **con rúbrica en Classroom** → columnas por criterio + total (verde)

---

## Configuración inicial (una sola vez)

### 1. Crear la hoja de Google Sheets

1. Ve a [sheets.google.com](https://sheets.google.com) y crea una hoja nueva.
2. Menú **Extensiones → Apps Script**.
3. Borra el código que aparece por defecto.
4. Copia el contenido del archivo **`Code.gs`** de este repositorio y pégalo.
5. Guarda (Ctrl+S) y cierra la pestaña de Apps Script.

### 2. Habilitar la API de Classroom

Dentro del editor de Apps Script (antes de cerrar):

1. Clic en **"Services"** (ícono `+` en el panel izquierdo)
2. Busca **"Google Classroom API"**
3. Clic en **Add**
4. Guarda de nuevo

### 3. Recargar la hoja

Recarga la hoja de Google Sheets. Aparecerá el menú **"Rúbricas Classroom"**.

> **Primera ejecución:** Google pedirá autorizar el script.
> Haz clic en "Revisar permisos" → selecciona tu cuenta →
> "Avanzado" → "Ir al proyecto (no seguro)" → "Permitir".
> Esto es normal porque el script accede a tu Classroom.

---

## Uso (2 pasos)

### Paso 1 — Seleccionar curso y sincronizar

**Rúbricas Classroom → 1. Seleccionar curso y sincronizar**

Elige tu curso en el menú desplegable. El script descarga automáticamente:

- Lista de estudiantes
- Todas las actividades con fechas y puntaje máximo
- Rúbricas configuradas en Classroom (criterios y niveles de calificación)
- Calificaciones por estudiante, **con desglose por criterio** si la actividad tiene rúbrica

### Paso 2 — Generar Reporte Final

**Rúbricas Classroom → 2. Generar Reporte Final**

Se crea la hoja **REPORTE**:

```
┌─────────────────────────┬───────────────────────┬───────┬──────────────────────────────────┬──────────┐
│ Estudiante              │ Email                 │ Prom. │  ACTIVIDAD UNO                   │ ACT DOS  │
│                         │                       │       │ Criterio1 │ Criterio2 │  TOTAL   │ Punteo   │
├─────────────────────────┼───────────────────────┼───────┼───────────┼───────────┼──────────┼──────────┤
│ Andrea Nájera           │ anajera@...            │       │     5     │     9     │    14    │    12    │
│ Angel Salguero          │ asalguero@...          │       │     4     │     8     │    12    │    10    │
└─────────────────────────┴───────────────────────┴───────┴───────────┴───────────┴──────────┴──────────┘
  Columnas azules = puntaje directo   |   Columnas verdes = criterios de rúbrica + TOTAL
```

> **Celda vacía** = actividad sin calificar todavía en Classroom.

### Actualizar datos

Cuando agregues nuevas calificaciones en Classroom, usa:

**Rúbricas Classroom → Sincronizar de nuevo (mismo curso)**

No tienes que volver a seleccionar el curso.

---

## Cómo funcionan las rúbricas

El script usa la [Google Classroom API](https://developers.google.com/classroom):

1. Para cada actividad verifica si tiene una **rúbrica configurada en Classroom**
2. Si tiene rúbrica: extrae los criterios y los puntajes que el docente asignó por criterio (`rubricGrades`)
3. Si no tiene rúbrica: usa el puntaje total directo (`assignedGrade`)

> Si una actividad tiene rúbrica definida en Classroom pero aún no fue calificada
> usando los criterios, las columnas de criterio aparecerán vacías aunque haya un punteo total.

---

## Un curso por hoja

Crea una hoja de Google Sheets diferente para cada curso/sección y repite
el Paso 1 en cada una.

