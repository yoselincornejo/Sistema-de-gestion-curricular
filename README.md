# Sistema de Gestión Curricular ICM

Aplicación web local para gestionar el currículo de Ingeniería Civil Matemática (Universidad de Valparaíso, Plan 2025): dashboard de cobertura de competencias, editor de programas de asignatura y generación de documentos oficiales (Matriz de Competencias en Excel, Mapa de Progreso y programas individuales en Word).

Todos los documentos se generan **desde la base de datos** (`data/sistema.db`), que es la única fuente de verdad del sistema — los Word originales en `data/programas/` son solo respaldo histórico.

---
## RESUMEN DE COMANDOS 

# CLONAR REPO

git clone https://github.com/yoselincornejo/Sistema-de-gestion-curricular.git
cd Sistema-de-gestion-curricular

# ACTIVAR ENTORNO

**macOS / Linux**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

**Windows (PowerShell)**

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

**Windows (CMD)**

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

# INSTALAR LIBRERIAS

pip install -r requirements.txt

# LEVANTAR APP

panel serve src/app.py --show --autoreload

## Requisitos previos

- **Python 3.10 o superior**
- **pip** (incluido con Python)
- **git** (para clonar el repositorio)

---

## 1. Clonar el repositorio

```bash
git clone https://github.com/yoselincornejo/Sistema-de-gestion-curricular.git
cd Sistema-de-gestion-curricular
```

## 2. Crear y activar el entorno virtual

Crea el entorno virtual una sola vez, luego actívalo cada vez que abras una terminal nueva para trabajar en el proyecto.

**macOS / Linux**

```bash
python3 -m venv .venv
source .venv/bin/activate
```

**Windows (PowerShell)**

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

**Windows (CMD)**

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

> El prompt debería mostrar `(.venv)` al inicio cuando el entorno está activo.

## 3. Instalar las librerías necesarias

Con el entorno activado, instala todas las dependencias del proyecto (mismo comando en los tres sistemas):

```bash
pip install -r requirements.txt
```

Esto instala:

| Paquete | Para qué se usa |
|---|---|
| `panel` | Interfaz web (dashboard + editor de programas) |
| `bokeh` | Motor de renderizado usado por Panel |
| `python-docx` | Generar documentos Word (programas y Mapa de Progreso) |
| `openpyxl` | Generar la Matriz de Competencias en Excel |
| `pandas` | Manejo de datos tabulares |

## 4. Verificar/crear la base de datos

Verifica que exista `data/sistema.db`. Si no existe, créala con:

```bash
python3 src/db_schema.py
```

(En Windows, usa `python` en vez de `python3`.)

## 5. Ejecutar la aplicación

Desde la **raíz del proyecto** (no desde dentro de `src/`):

```bash
panel serve src/app.py --show --autoreload
```

- Abre automáticamente el navegador en `http://localhost:5006/app`
- `--autoreload` recarga la app al guardar cambios en `src/app.py`
- Para detener el servidor: `Ctrl+C`

---

## Estructura del proyecto

```
Sistema-de-gestion-curricular/
├── data/
│   ├── sistema.db              # base de datos SQLite — fuente única de verdad
│   ├── logo_uv.jpg             # logo institucional
│   ├── plantilla_matriz.xlsx   # plantilla de referencia para el Excel
│   ├── programas/              # Word originales (respaldo histórico)
│   ├── programas_json/         # JSONs intermedios de la carga inicial
│   └── output/                 # documentos generados por la app
├── src/
│   ├── app.py                  # interfaz Panel: dashboard + editor
│   ├── generador_excel.py      # genera la Matriz de Competencias (.xlsx)
│   ├── generador_word.py       # genera programas y Mapa de Progreso (.docx)
│   ├── db_schema.py            # crea el schema de la BD
│   ├── editar.py               # CLI para vincular/desvincular tributaciones
│   └── verificar_db.py         # inspección rápida del estado de la BD
├── requirements.txt
└── .venv/                      # entorno virtual (no se sube a git)
```

---

## Uso cotidiano

- **Dashboard de Cobertura** — muestra las 8 competencias agrupadas en 3 bloques (Licenciatura, Título Profesional, Sello UV). Al hacer clic se despliegan los Resultados de Aprendizaje y las asignaturas que tributan a cada uno; los RA sin cobertura se marcan en rojo.
- **Editor de Programas** — permite editar identificación, tributación de RAs, unidades, contenidos, metodología y evaluaciones de una asignatura. Los cambios se guardan directamente en `data/sistema.db`.
- **Generar documentos** — desde el dashboard: "Generar Matriz de Competencias" y "Generar Mapa de Progreso". Desde el Editor: "Generar Word" para el programa individual.

## Comandos útiles desde terminal

```bash
# Estado general de la base de datos
python3 src/verificar_db.py

# Vincular / desvincular una asignatura a un Resultado de Aprendizaje
python3 src/editar.py mostrar IMAT322
python3 src/editar.py vincular IMAT322 "CL2, N2, RA1"
python3 src/editar.py desvincular IMAT322 "CL2, N2, RA1"
python3 src/editar.py listar-ras
python3 src/editar.py listar-asignaturas

# Generar documentos sin abrir la interfaz web
python3 src/generador_excel.py
python3 src/generador_word.py mapa              # Mapa de Progreso
python3 src/generador_word.py <id_asignatura>   # programa individual
```

---

## Notas importantes

- **La base de datos es la fuente de verdad.** Los Word en `data/programas/` y los JSON en `data/programas_json/` ya no son leídos por la app en ningún flujo normal; son respaldo por si la BD necesita reconstruirse.
- **El nivel académico (N1/N2/N3)** de cada asignatura se calcula siempre desde el semestre (1–4 → N1, 5–8 → N2, 9–11 → N3), nunca desde texto libre.
- Pendiente de confirmar con la fuente oficial: el número exacto de competencias Sello UV y la tabla de prerrequisitos entre asignaturas.
