# Sistema de Gestión Curricular ICM

Aplicación web local para gestionar el currículo de Ingeniería Civil Matemática
(Universidad de Valparaíso, Plan 2025): dashboard de cobertura de competencias,
editor de programas de asignatura, y generación de documentos oficiales
(Matriz de Competencias en Excel, Mapa de Progreso y programas individuales en Word).

Todos los documentos se generan **desde la base de datos**, no desde archivos
Word externos — la base de datos es la única fuente de verdad del sistema.

---

## Requisitos previos

- Python 3.10 o superior
- pip

---

## 1. Activar el entorno virtual

Si es la primera vez que configuras el proyecto, crea el entorno:

```bash
python3 -m venv .venv
```

Luego, cada vez que abras una terminal nueva para trabajar en el proyecto,
actívalo:

```bash
source .venv/bin/activate
```

El prompt de tu terminal debería mostrar `(.venv)` al inicio cuando está
activo. Si en algún momento ves un error como `command 'panel' not found`,
es señal de que esta terminal nueva no tiene el entorno activado — vuelve
a correr el comando de arriba.

---

## 2. Instalar dependencias

Con el entorno activado:

```bash
pip install panel python-docx openpyxl
```

Estas son las tres únicas dependencias externas del proyecto:

| Paquete | Para qué se usa |
|---|---|
| `panel` | Interfaz web (dashboard + editor de programas) |
| `python-docx` | Generar los documentos Word (programas individuales y Mapa de Progreso) |
| `openpyxl` | Generar la Matriz de Competencias en Excel |

---

## 3. Archivos que el sistema necesita encontrar

Antes de correr la app, verifica que existan estas rutas dentro de `data/`:

```bash
ls data/sistema.db          # la base de datos (ya debería existir)
ls data/logo_uv.jpg          # logo institucional (opcional — hay fallback de texto si falta)
ls data/plantilla_matriz.xlsx  # plantilla del Excel (necesaria para generar la Matriz)
```

Si `data/sistema.db` no existe todavía, créala con el schema base:

```bash
python3 src/db_schema.py
```

Esto crea las tablas vacías y precarga las 8 competencias del plan 2025
(CL1, CL2, CE1, CE2, CG1–CG4).

---

## 4. Levantar la aplicación web

Desde la **raíz del proyecto** (no desde dentro de `src/`):

```bash
panel serve src/app.py --show --autoreload
```

Esto:
- Abre automáticamente el navegador en `http://localhost:5006/app`
- `--autoreload` hace que la app se recargue sola cada vez que guardas un
  cambio en `src/app.py` (útil mientras se sigue desarrollando)

Para detener el servidor: `Ctrl+C` en la terminal.

---

## Estructura del proyecto

```
sistema-gestion-curricular-icm/
├── data/
│   ├── sistema.db              # base de datos SQLite — fuente única de verdad
│   ├── logo_uv.jpg              # logo institucional
│   ├── plantilla_matriz.xlsx    # plantilla de referencia para el Excel
│   ├── programas/                # 53 Word originales (solo respaldo histórico,
│   │                              # el sistema ya NO los lee en tiempo de ejecución)
│   ├── programas_json/           # JSONs intermedios de la carga inicial
│   │                              # (necesarios solo si hay que repoblar la BD)
│   └── output/                   # documentos generados por la app (Excel/Word)
├── src/
│   ├── app.py                    # interfaz Panel: dashboard + editor
│   ├── generador_excel.py        # genera la Matriz de Competencias (.xlsx)
│   ├── generador_word.py         # genera programas individuales y Mapa de Progreso (.docx)
│   ├── db_schema.py              # crea el schema de la BD
│   ├── editar.py                 # CLI para vincular/desvincular tributaciones desde terminal
│   ├── verificar_db.py           # inspección rápida del estado de la BD
│   ├── cargar_a_db.py            # carga los JSONs iniciales a la BD (uso histórico, no rutina)
│   ├── extraer.py                # parsea un Word individual a JSON (uso histórico)
│   ├── procesar_todos.py         # corre extraer.py sobre todos los Word (uso histórico)
│   └── limpiar_competencias_fantasma.py  # elimina competencias obsoletas detectadas (CL4, CL5, ND1, ND2)
└── .venv/                        # entorno virtual (no se sube a git)
```

---

## Uso cotidiano

Una vez que la app está corriendo:

- **Dashboard de Cobertura** — muestra las 8 competencias agrupadas en 3
  bloques (Licenciatura, Título Profesional, Sello UV). Al hacer clic en
  una competencia se despliegan sus Resultados de Aprendizaje, cada uno
  con las asignaturas que tributan a él. Los RA sin ninguna asignatura
  cubriéndolos se marcan en rojo con la etiqueta "sin cobertura".

- **Editor de Programas** — selecciona una asignatura para editar su
  identificación, tributación de RAs, unidades y contenidos, metodología
  y evaluaciones. Al guardar, los cambios se persisten directamente en
  `data/sistema.db`.

- **Generar documentos** — desde el dashboard, los botones "Generar Matriz
  de Competencias" y "Generar Mapa de Progreso" crean los archivos en
  `data/output/`. Desde el Editor de Programas, el botón "Generar Word"
  crea el programa individual de la asignatura seleccionada.

---

## Comandos útiles desde terminal

Verificar el estado general de la base de datos:

```bash
python3 src/verificar_db.py
```

Vincular o desvincular una asignatura a un Resultado de Aprendizaje sin
pasar por la interfaz:

```bash
python3 src/editar.py mostrar IMAT322
python3 src/editar.py vincular IMAT322 "CL2, N2, RA1"
python3 src/editar.py desvincular IMAT322 "CL2, N2, RA1"
python3 src/editar.py listar-ras
python3 src/editar.py listar-asignaturas
```

Generar la Matriz de Competencias o un programa individual sin abrir la
interfaz web:

```bash
python3 src/generador_excel.py
python3 src/generador_word.py mapa          # Mapa de Progreso
python3 src/generador_word.py <id_asignatura>  # programa individual por ID
```

---

## Notas importantes
-Los archivos generados por el sistema son guardados dentro de la carpeta 'output'
sistema-gestion-curricular-icm/data/output

- **La base de datos es la fuente de verdad.** Los Word originales en
  `data/programas/` y los JSON en `data/programas_json/` ya no son leídos
  por la aplicación en ningún flujo normal — son respaldo histórico de la
  carga inicial. No los borres hasta confirmar que la BD tiene toda la
  información que necesitas; en caso de que la BD necesite reconstruirse,
  son tu única ruta de regreso.

- **El nivel académico (N1/N2/N3) de cada asignatura se calcula siempre
  desde el semestre**, no se lee de ningún campo de texto libre: semestres
  1–4 son N1, 5–8 son N2, 9–11 son N3.

- Pendiente de información externa (no modificar hasta confirmar con la
  fuente oficial): el número exacto de competencias Sello UV, y la tabla
  de prerrequisitos entre asignaturas.

  
