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

---

# ⚽ Campeonato de Fútbol — Tabla de Posiciones interactiva

Módulo adicional para organizar un **campeonato escolar de fútbol** con
control de partidos en Google Sheets y una **tabla de posiciones web**
interactiva y responsiva, lista para embeberse en **Google Sites**.

**Arquitectura:** Google Sheets (datos) + Apps Script Web App (HTML/JS con
Tailwind) + Google Sites (portada que embebe la Web App).

## Características

- **Varios grupos** en paralelo (ej. `Grupo A`, `Grupo B`, `Femenino`) con
  tabla de posiciones independiente por grupo
- Selector de grupo visible al tope de la Web App (se oculta si solo hay uno)
- Tabla de posiciones ordenable (PJ, PG, PE, PP, GF, GC, DG, PTS)
- Cálculo automático: 3 pts victoria, 1 pt empate
- Desempates: Puntos → Diferencia de goles → Goles a favor
- Racha de forma (últimos 5 partidos)
- Listado de partidos filtrable por jornada y por grupo
- Detalle por equipo en modal (historial de partidos)
- Ranking de goles por equipo dentro del grupo
- Resaltado de posiciones de clasificación
- 100% responsivo (móvil, tablet, pantalla)
- Se actualiza al modificar la hoja

## Archivos del módulo

| Archivo | Descripción |
|---|---|
| `Campeonato.gs` | Backend: lectura de hojas, cálculo de tabla, Web App |
| `Index.html`   | Página principal (header, tabs, layout) |
| `Styles.html`  | CSS personalizado (animaciones, estilos) |
| `Script.html`  | Lógica del cliente (render, filtros, modal) |

## Configuración inicial

### 1. Crear una hoja de Sheets nueva para el campeonato

Recomendado usar un archivo **independiente** del de Classroom para no
mezclar datos.

### 2. Copiar los archivos a Apps Script

1. En la hoja: **Extensiones → Apps Script**
2. Crea estos archivos (mismo nombre, sin extensión):
   - `Campeonato.gs` → pega el contenido de `Campeonato.gs`
   - `Index` (HTML) → pega `Index.html`
   - `Styles` (HTML) → pega `Styles.html`
   - `Script` (HTML) → pega `Script.html`
3. Guarda (Ctrl+S).

### 3. Registrar el menú al abrir la hoja

En el editor de Apps Script abre el archivo que controla `onOpen()`
(por ejemplo `Code.gs`) y agrega una llamada a `onOpenCampeonato()`
dentro de tu `onOpen()` existente, **o** si el proyecto solo es del
campeonato, renombra `onOpenCampeonato` a `onOpen`.

### 4. Preparar las hojas

Recarga Google Sheets → menú **⚽ Campeonato → 1. Crear hojas base**.

Se crean 3 hojas:

- **Equipos** — `ID | Nombre | Grado/Sección | Grupo | Escudo (URL)`
- **Partidos** — `Jornada | Fecha | Local | Goles L | Visitante | Goles V | Estado`
- **Config** — nombre del campeonato, colores, logo, clasificados, orden de grupos

La columna **Grupo** de Equipos admite cualquier texto; por defecto trae
validación con `Grupo A`, `Grupo B`, `Femenino`. Los partidos se asignan al
grupo automáticamente según el equipo local.

> Opcional: **⚽ Campeonato → 2. Cargar datos de ejemplo** llena las hojas
> con 6 equipos y 4 jornadas de muestra para ver el resultado inmediato.

## Uso diario

1. Ingresa los partidos en la hoja **Partidos**.
2. Asegúrate de que el **Estado** sea `Jugado` cuando haya resultado
   (si dejas goles en blanco, queda como `Pendiente`).
3. Abre la Web App o pulsa **Actualizar** en la tabla — se recalcula sola.

## Publicar la Web App

En Apps Script:

1. **Implementar → Nueva implementación**
2. Tipo: **Aplicación web**
3. Ejecutar como: **Tu cuenta**
4. Quién tiene acceso: **Cualquier usuario**
5. Copia la **URL de la Web App**

## Embeberla en Google Sites

1. Abre tu sitio en [sites.google.com](https://sites.google.com)
2. **Insertar → Insertar → URL** y pega la URL de la Web App
3. Redimensiona el contenedor a ancho completo
4. Publica el sitio

Google Sites queda como portal (reglas, calendario, noticias) y la tabla
viva vive en el embed, actualizándose al recargar.

## Personalización rápida

En la hoja **Config**:

- `Nombre del Campeonato` → título del header
- `Subtítulo` → texto bajo el título
- `Color Primario` / `Color Secundario` → degradado del header y botones
  (ej. `#1e40af`, `#059669`)
- `Logo (URL)` → URL pública de la imagen del torneo (opcional)
- `Clasifican (primeros N)` → cuántos primeros puestos se resaltan en verde
  (aplica por grupo)
- `Orden de grupos` → lista separada por comas con el orden en que aparecen
  las pestañas de grupo (ej. `Grupo A, Grupo B, Femenino`)

