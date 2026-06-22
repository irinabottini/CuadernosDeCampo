# CP Cuadernos de Campo

Aplicación corporativa para generar exportables de trabajo a partir de un Excel fuente de cuatro hojas.

## Estado actual

- Genera **Libro de campo - nombre.xlsx**.
- Usa `libro_campo_modelo.xlsm` como plantilla visual.
- Copia y conserva las hojas base del modelo.
- Actualiza `Plano`, `Aplicaciones`, `Tratamientos` y `Datos a completar` con datos del Excel fuente.
- Deja el módulo **Protocolo** preparado para la próxima etapa.

## Uso local

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Abrir `http://127.0.0.1:5000` y subir un Excel como `Vayego 1.xlsx`.

## Deploy en Render

Crear un repositorio en GitHub con estos archivos en la raíz y después crear un `Web Service` en Render.

Configuración manual:

```text
Runtime: Python
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 120
```

El archivo `render.yaml` ya incluye esta configuración como Blueprint opcional.

Render usa `.python-version` para fijar la versión de Python. Este repo usa `3.13`.

## Archivo fuente desde Scout

Antes de generar el libro de campo, exportar desde Scout el Report Set `Protocolo_completo` con estos reportes:

1. `Protocol Audit Trail`
2. `Protocol Data Headers`
3. `Protocol Description`
4. `Protocol Treatments`

Marcar `Logo`, dejar la clasificación como `Internal` y cargar el Excel exactamente como lo descarga Scout, sin editar hojas, nombres ni columnas.

## Reglas implementadas para Libro de Campo

- El Excel fuente debe tener exactamente 4 hojas.
- La hoja `Plano` toma el título desde `A6` de la primera hoja del Excel fuente.
- La hoja `Aplicaciones` mantiene las columnas `A:C` del modelo y llena desde `D` con la tabla detectada en la tercera hoja.
- La hoja `Tratamientos` se arma desde la cuarta hoja, empezando en la fila donde columna `A` dice `Nº` y la siguiente fila dice `Tr.`.
- La hoja `Datos a completar` se arma desde la segunda hoja, eliminando columnas con datos en la fila 25.
- En `Datos a completar` se conservan las filas fuente 3, 11, 12, 13, 15, 16, 19, 20, 22 y 23.
- En `Datos a completar` se agregan las filas `Fenología` y `Fecha`.
- La app pide cantidad de repeticiones antes de generar el Excel, con valor default `4`.
- La grilla de carga se arma por repetición, tratamiento y submuestra; `Parcela` queda vacía.
