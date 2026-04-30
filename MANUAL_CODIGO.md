# Manual del código: `conciliacion.py`

Este documento explica cómo funciona el script línea a línea, pensado para quien quiere entender la lógica y aprender Python leyendo código real.

---

## Índice

1. [Estructura general del archivo](#1-estructura-general-del-archivo)
2. [Imports y constantes globales](#2-imports-y-constantes-globales)
3. [Funciones auxiliares](#3-funciones-auxiliares)
4. [Lectura de archivos](#4-lectura-de-archivos)
5. [Funciones de extracción de datos](#5-funciones-de-extracción-de-datos)
6. [Algoritmo de conciliación de cobros](#6-algoritmo-de-conciliación-de-cobros)
7. [Algoritmo de conciliación de gastos](#7-algoritmo-de-conciliación-de-gastos)
8. [Generación de salidas](#8-generación-de-salidas)
9. [Función main](#9-función-main)
10. [Flujo completo de datos](#10-flujo-completo-de-datos)

---

## 1. Estructura general del archivo

El script es un único archivo `.py` de ~1100 líneas organizado así:

```
docstring del módulo       →  explicación general (líneas 1-28)
imports                    →  librerías que usa (líneas 30-35)
constantes globales        →  rutas y tolerancias (líneas 41-61)
funciones auxiliares       →  helpers pequeños (líneas 64-255)
funciones de lectura       →  parsean los 3 archivos de entrada (líneas 114-210)
funciones de extracción    →  extraen CUIT, nombre, centavos (líneas 213-255)
conciliar_cobros()         →  algoritmo principal de cobros (líneas 258-632)
conciliar_gastos()         →  algoritmo de gastos (líneas 635-725)
marcar_extracto()          →  pinta el Excel con colores (líneas 728-752)
generar_reporte()          →  arma el Excel de reporte (líneas 755-952)
main()                     →  orquesta todo (líneas 955-1075)
```

No hay clases. Todo son funciones. Esto es una decisión deliberada: es más fácil seguir el flujo de un dato a través de funciones simples que a través de métodos de objetos.

---

## 2. Imports y constantes globales

### Líneas 30-35: Imports

```python
import subprocess, re, sys, shutil
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
```

**¿Qué es cada import?**

- `subprocess` — permite ejecutar programas externos. Acá se usa para llamar a LibreOffice y convertir `.xls` a `.xlsx`.
- `re` — módulo de expresiones regulares. Se usa para extraer el CUIT de una descripción de texto libre.
- `sys` — acceso al sistema. Se usa para `sys.exit(1)` cuando hay un error fatal.
- `shutil` — utilidades de archivos. Se usa para copiar el extracto original antes de pintarlo.
- `pathlib.Path` — manejo de rutas de archivo de forma elegante. `Path("carpeta") / "archivo.xlsx"` es más limpio que concatenar strings.
- `datetime` — para timestamps en los nombres de los archivos de salida.
- `openpyxl` — la única dependencia externa. Lee y escribe archivos `.xlsx` de Excel.
  - `load_workbook` — abre un `.xlsx` existente.
  - `Workbook` — crea un `.xlsx` nuevo desde cero.
  - `PatternFill`, `Font`, `Alignment` — estilos de celda (colores, negrita, alineación).
  - `get_column_letter` — convierte número de columna a letra (ej: 3 → "C"). Se importa pero no se usa activamente en el código principal.

### Líneas 41-61: Constantes globales

```python
EXTRACTO_BANCO   = r"extracto.xlsx"
CARPETA_ADA4     = r"."
CARPETA_SALIDA   = r"."
TOLERANCIA_GASTO = 0.50
TOLERANCIA_COBRO = 0.02
```

El prefijo `r` en `r"."` indica un *raw string* — en este caso no cambia nada (no hay barras invertidas que escapar), pero es convención usarlo para rutas.

El `.` significa "carpeta actual" — el script busca los archivos en la misma carpeta donde está.

```python
FILL_VERDE    = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
FILL_AMARILLO = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
FILL_NARANJA  = PatternFill(start_color="FFD699", end_color="FFD699", fill_type="solid")
FILL_ROJO     = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
FILL_HEADER   = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
FONT_HEADER   = Font(bold=True, color="FFFFFF")
```

Los colores se crean una sola vez acá arriba y se reutilizan en todo el código. Los valores hexadecimales son los colores "estilo semáforo" de Excel (los mismos que usa el formato condicional de Excel por defecto). `start_color` y `end_color` son iguales porque es relleno sólido (sin degradado).

---

## 3. Funciones auxiliares

### `encontrar_libreoffice()` — líneas 64-69

```python
def encontrar_libreoffice():
    for p in [r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
        if Path(p).exists():
            return p
    return None
```

Busca el ejecutable de LibreOffice en las dos ubicaciones posibles en Windows (64-bit y 32-bit). `Path(p).exists()` devuelve `True` si el archivo existe en disco. Si no encuentra ninguno, devuelve `None` y el script lo maneja más adelante mostrando un mensaje de error.

### `buscar_archivo_por_prefijo()` — líneas 72-95

```python
def buscar_archivo_por_prefijo(carpeta, prefijo, extensiones):
```

Esta función resuelve un problema real: el nombre completo del archivo ADA4 incluye fecha y hora (`InformeGastos_2024-05-15_09-30.xls`), así que no podemos hardcodear el nombre exacto.

**Paso a paso:**

```python
carpeta = Path(carpeta)
if not carpeta.exists():
    return None
```
Convierte el string a un objeto `Path` y verifica que la carpeta exista.

```python
candidatos = []
for archivo in carpeta.iterdir():
    if not archivo.is_file():
        continue
    nombre = archivo.name
    if not nombre.startswith(prefijo):
        continue
    if not any(nombre.lower().endswith(ext.lower()) for ext in extensiones):
        continue
    candidatos.append(archivo)
```

`carpeta.iterdir()` lista todos los archivos y subcarpetas. Para cada elemento:
- `is_file()` descarta subcarpetas.
- `startswith(prefijo)` filtra por prefijo (ej: `"InformeGastos"`).
- `any(... for ext in extensiones)` verifica que termine en alguna de las extensiones permitidas. El `.lower()` hace la comparación case-insensitive.

```python
candidatos.sort(key=lambda p: p.stat().st_mtime, reverse=True)
return str(candidatos[0])
```

Ordena los candidatos por fecha de modificación (más reciente primero). `p.stat().st_mtime` devuelve el timestamp de modificación como número flotante (segundos desde epoch Unix). `reverse=True` pone el más reciente al inicio. Devuelve el primero (el más reciente).

### `convertir_xls()` — líneas 98-111

```python
def convertir_xls(ruta_xls):
    soffice = encontrar_libreoffice()
    if not soffice:
        raise FileNotFoundError("LibreOffice no encontrado...")
```

`raise` lanza una excepción. Si no encuentra LibreOffice, el script no puede continuar — el `main()` captura esta excepción y muestra el error al usuario.

```python
    r = subprocess.run(
        [soffice, '--headless', '--convert-to', 'xlsx',
         str(xls), '--outdir', str(xls.parent)],
        capture_output=True, text=True)
```

`subprocess.run()` ejecuta un programa externo y espera a que termine. Los argumentos:
- `--headless` — ejecuta LibreOffice sin interfaz gráfica (modo silencioso).
- `--convert-to xlsx` — indica el formato destino.
- `--outdir` — carpeta donde guarda el archivo convertido.
- `capture_output=True` — captura stdout y stderr para poder leerlos.
- `text=True` — decodifica la salida como texto (no bytes).

```python
    out = xls.parent / (xls.stem + '.xlsx')
    if not out.exists():
        raise RuntimeError(f"Falló la conversión:\n{r.stderr}")
    return str(out)
```

`xls.stem` es el nombre sin extensión (ej: `"InformeGastos_2024-05-15"`). Si el archivo `.xlsx` no aparece después de ejecutar LibreOffice, algo falló y se incluye el mensaje de error de stderr.

---

## 4. Lectura de archivos

### `leer_extracto()` — líneas 114-153

```python
def leer_extracto(ruta):
    wb = load_workbook(ruta, read_only=True)
    ws = wb.active
```

`load_workbook(..., read_only=True)` abre el archivo en modo lectura, más eficiente en memoria. `wb.active` obtiene la primera hoja activa.

```python
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        fecha, comp, conc, desc, imp = row[0], row[1], row[2], row[3], row[4]
        if imp is None or isinstance(imp, str):
            continue
```

`ws.iter_rows(values_only=True)` itera fila por fila devolviendo tuplas de valores (no objetos de celda). `enumerate(..., 1)` numera desde 1 (coincide con la fila real del Excel).

El `if imp is None or isinstance(imp, str): continue` salta filas vacías o con texto en la columna de importe (como encabezados que dicen "Importe").

```python
        item = {
            'fila_banco':  i,
            'fecha':       fecha,
            ...
            'tipo':        'cobro' if imp > 0 else 'gasto',
            'es_qr':       False,
            'bruto_qr':    None,
        }
```

Cada movimiento se convierte en un diccionario Python. Los importes positivos son cobros, los negativos son gastos — esta es la convención del banco.

```python
        if 'CREDITO POR COBRO QR' in descripcion.upper():
            h_val = row[7] if len(row) > 7 else None
            f_val = row[5] if len(row) > 5 else None
            if h_val is not None and not isinstance(h_val, str):
                item['es_qr'] = True
                item['bruto_qr'] = float(h_val)
            elif f_val is not None and not isinstance(f_val, str) and float(f_val) < 0:
                item['es_qr'] = True
                item['bruto_qr'] = float(imp) + abs(float(f_val))
```

Los cobros QR tienen dos formatos posibles en el extracto. El código detecta cuál de los dos aplica:
- **Formato 1**: columna H tiene el bruto directamente.
- **Formato 2**: columna F tiene la comisión como número negativo; bruto = neto + |comisión|.

El `descripcion.upper()` convierte a mayúsculas para comparar sin importar si el banco a veces lo escribe en minúsculas.

### `leer_gastos_ada()` — líneas 156-181

```python
def leer_gastos_ada(ruta):
    wb = load_workbook(ruta, read_only=True)
    ws = wb.active
    gastos = []
    for row in ws.iter_rows(values_only=True):
        if not any(v is not None for v in row):
            continue
```

`not any(v is not None for v in row)` es True cuando todos los valores de la fila son `None` — es una fila completamente vacía. Las salta.

```python
        imp = row[3]
        if imp is None or isinstance(imp, str):
            continue
        gastos.append({
            'periodo':    row[0],
            'descripcion': str(row[1] or '').strip(),
            'importe':    float(imp),
            ...
            'usado':      False,
        })
```

`str(row[1] or '').strip()`: 
- `row[1] or ''` devuelve `''` si `row[1]` es `None` (evita pasar `None` a `str()`).
- `.strip()` elimina espacios al inicio y al final.

El campo `'usado': False` es clave para el algoritmo. Se empieza en `False` y se pone en `True` cuando se empareja con un movimiento del banco, evitando que el mismo gasto ADA sea usado dos veces.

### `leer_cobros_ada()` — líneas 184-210

Misma lógica que `leer_gastos_ada`, pero para la CobranzaRapida.

```python
        uf = row[2] if len(row) > 2 else None
        if uf is None or str(uf).strip() in ('', 'UF'):
            continue
        uf_str = str(uf).strip()
        if not any(c.isdigit() for c in uf_str):
            continue
```

Filtra encabezados y filas inválidas. El check `not any(c.isdigit() for c in uf_str)` descarta filas donde el campo UF contiene texto puro sin números.

```python
        cobros.append({
            ...
            'deuda_anterior': float(deuda_ant),
            'movimiento':    float(row[7]) if row[7] is not None else 0,
            'usado':         False,
        })
```

`movimiento` puede ser `None` si la celda está vacía, por eso el `if ... else 0`. Un movimiento negativo en ADA significa que la UF pagó (convención contable: el pago "sale" del saldo deudor).

---

## 5. Funciones de extracción de datos

### `extraer_cuit()` — líneas 213-221

```python
def extraer_cuit(descripcion):
    m = re.search(r'-?(\d{11})-?', descripcion)
    if m:
        return m.group(1)
    m = re.search(r'(\d{2}-\d{8}-\d{1})', descripcion)
    if m:
        return m.group(1).replace('-', '')
    return None
```

`re.search()` busca el patrón en cualquier parte del string (no solo al inicio).

- Primer patrón `r'-?(\d{11})-?'`: busca 11 dígitos seguidos, con guiones opcionales alrededor. Los paréntesis `(\d{11})` crean un *grupo de captura* — `m.group(1)` devuelve solo los 11 dígitos, sin los guiones.
- Segundo patrón: algunos bancos usan el formato `20-12345678-9`. Se captura y se eliminan los guiones con `.replace('-', '')`.

### `extraer_nombre_pagador()` — líneas 224-245

```python
def extraer_nombre_pagador(descripcion):
    desc = re.sub(r'\d{11}', '', descripcion).upper()
    desc = re.sub(r'[/,.\-]', ' ', desc)
    tokens = [t for t in desc.split() if len(t) >= 3 and t.isalpha()]
    excluir = {'TRANSF', 'INTER', 'DISTINTO', ...}
    tokens = [t for t in tokens if t not in excluir]
    return tokens[0] if tokens else None
```

`re.sub(patrón, reemplazo, string)` reemplaza todas las coincidencias del patrón en el string.

La lógica:
1. Eliminar el CUIT del texto (para no confundirlo con un nombre).
2. Reemplazar separadores (`/`, `,`, `.`, `-`) por espacios.
3. Dividir en tokens y quedarse con los de 3+ letras puros (sin números).
4. Filtrar palabras del vocabulario bancario que no son nombres.
5. Devolver el primero — suele ser el apellido.

Esta función no es perfecta (el lenguaje natural es difícil), pero captura el caso más común con suficiente precisión.

### `obtener_centavos()` — línea 248-250

```python
def obtener_centavos(importe):
    return round((importe - int(importe)) * 100)
```

`int(importe)` trunca hacia cero. Para `179304.07`: `int(179304.07)` = `179304`. La diferencia es `0.07`. Multiplicado por 100 = `7.0`. `round()` convierte a entero limpio.

**¿Por qué `round()` en vez de `int()`?** Porque la aritmética de punto flotante en las computadoras no es exacta. `0.07 * 100` a veces da `6.999999999...` en vez de `7.0`, y `int()` truncaría eso a `6`. El `round()` evita ese error.

### `es_cobro_qr()` — líneas 253-255

```python
def es_cobro_qr(descripcion):
    return 'CREDITO POR COBRO QR' in (descripcion or '').upper()
```

El `(descripcion or '')` maneja el caso donde `descripcion` es `None`. En Python, `None or ''` devuelve `''`, que es un string vacío — nunca va a contener el texto buscado, así que la función devuelve `False` sin errores.

---

## 6. Algoritmo de conciliación de cobros

Esta es la función central del script. Tiene ~380 líneas porque maneja 5 estrategias con sus edge cases.

### Estructura general — líneas 258-279

```python
def conciliar_cobros(cobros_banco, cobros_ada):
    resultados = []
    procesados = set()

    def imp_match(cb):
        if cb.get('es_qr') and cb.get('bruto_qr') is not None:
            return cb['bruto_qr']
        return cb['importe']

    TOLERANCIA_CUOTAS = 1.00
```

`procesados` es un `set` de índices de cobros del banco ya resueltos. Los `set` en Python tienen búsqueda O(1) — verificar si un índice está en el conjunto es instantáneo sin importar cuántos elementos tenga.

`imp_match()` es una función auxiliar definida *dentro* de `conciliar_cobros`. En Python esto se llama *función anidada* y puede acceder a variables del scope externo. Su propósito: para cobros QR, el importe de comparación es el bruto (no el neto que llegó al banco).

`TOLERANCIA_CUOTAS = 1.00` se define acá (no en el nivel global) porque es específica del contexto de cuotas.

### Estrategia 1: COBROS QR — líneas 292-316

```python
for idx, cb in enumerate(cobros_banco):
    if idx in procesados:
        continue
    if not (cb.get('es_qr') and cb.get('bruto_qr') is not None):
        continue
    bruto    = cb['bruto_qr']
    centavos = obtener_centavos(bruto)
    for ca in cobros_ada:
        uf_num = int(re.sub(r'\D', '', ca['uf']) or '0')
        if uf_num == centavos and not ca['usado']:
            if abs(ca['deuda_anterior'] - bruto) <= TOLERANCIA_COBRO:
                ...
                ca['usado'] = True
                procesados.add(idx)
                break
```

`re.sub(r'\D', '', ca['uf'])` elimina todo lo que no sea dígito del campo `uf` (`\D` = "non-digit"). Si `uf` es `"0007"`, el resultado es `"7"`. El `or '0'` maneja el caso improbable de que quede string vacío.

`procesados.add(idx)` marca el cobro como resuelto. `ca['usado'] = True` marca la UF ADA como resuelta. El `break` sale del loop de cobros ADA cuando encuentra match (no sigue buscando).

### Estrategia 2: MATCH ÚNICO — líneas 318-374

```python
for ca in cobros_ada:
    if ca['usado'] or ca['movimiento'] == 0:
        continue
    movimiento_abs = abs(ca['movimiento'])
    uf_num = int(re.sub(r'\D', '', ca['uf']) or '0')

    cand_con_centavos = []
    cand_sin_centavos = []
    for idx, cb in enumerate(cobros_banco):
        if idx in procesados:
            continue
        imp = imp_match(cb)
        if abs(imp - movimiento_abs) <= TOLERANCIA_CUOTAS:
            if obtener_centavos(imp) == uf_num:
                cand_con_centavos.append(idx)
            else:
                cand_sin_centavos.append(idx)

    candidatos = cand_con_centavos if cand_con_centavos else cand_sin_centavos

    if len(candidatos) == 1:
        ...
```

Esta estrategia itera sobre las UFs de ADA (no sobre los cobros del banco). Para cada UF con movimiento registrado, busca cobros del banco cuyo importe coincida.

La separación en `cand_con_centavos` y `cand_sin_centavos` implementa una prioridad: si hay un cobro cuyo importe además tiene los centavos correctos, ese es más confiable que uno que solo coincide en monto. Pero si no hay ninguno con centavos, acepta los sin centavos.

`if len(candidatos) == 1` es la condición clave: solo concilia si hay exactamente un candidato. Si hay 2 cobros que coinciden en importe, el script no puede saber cuál es el correcto — los deja pasar a estrategias posteriores.

El mensaje de detalle se genera según la relación entre `movimiento_abs` y `deuda_anterior`:

```python
if diff_deuda <= TOLERANCIA_COBRO:
    detalle = ''                          # pagó exacto
elif movimiento_abs < ca['deuda_anterior']:
    detalle = f"Pago parcial: pagó..."    # pagó menos de la deuda
else:
    detalle = f"Pagó de más:..."          # pagó más de la deuda
```

### Estrategia 3: PAGOS EN CUOTAS — líneas 376-534

#### Helper `intentar_match_cuotas()` — líneas 388-411

```python
def intentar_match_cuotas(indices, ca, ref):
    total = sum(cobros_banco[i]['importe'] for i in indices)
    n_cuotas = len(indices)
    ref_monto = abs(ca['movimiento']) if ref == 'movimiento ADA' else ca['deuda_anterior']
    cuits = {extraer_cuit(cobros_banco[i]['descripcion']) for i in indices}
    cuits.discard(None)
    cuit_comun = cuits.pop() if len(cuits) == 1 else None
```

`{...}` con llaves y una expresión dentro es un *set comprehension* — crea un conjunto (sin duplicados). Si todos los cobros tienen el mismo CUIT, el conjunto tiene un solo elemento.

`cuits.discard(None)` elimina `None` del conjunto si está (algunos cobros no tienen CUIT en la descripción). A diferencia de `.remove()`, `.discard()` no lanza error si el elemento no existe.

```python
    for i, idx in enumerate(
            sorted(indices, key=lambda x: cobros_banco[x]['fecha'] or datetime.min), 1):
```

Ordena los cobros por fecha antes de numerar las cuotas (cuota 1/2 va al pago más antiguo). `datetime.min` es la fecha mínima posible — si `fecha` es `None`, se usa ese valor para que quede al inicio del orden.

#### Estrategia 3a: mismo CUIT — líneas 413-467

```python
grupos_cuit = {}
for idx, cb in enumerate(cobros_banco):
    if cb.get('es_qr'):
        continue
    cuit = extraer_cuit(cb['descripcion'])
    if cuit:
        grupos_cuit.setdefault(cuit, []).append(idx)
```

`dict.setdefault(key, default)` devuelve el valor para esa clave si existe, o lo inserta con el valor por defecto y lo devuelve. Es un patrón común para agrupar: si el CUIT ya tiene una lista, agrega el índice; si no, crea la lista y agrega.

```python
for cuit, indices in grupos_cuit.items():
    if len(indices) < 2:
        continue
    total = sum(cobros_banco[i]['importe'] for i in indices)
    centavos_suma = obtener_centavos(total)
```

Solo procesa grupos con 2+ cobros. Los checks que siguen verifican si los centavos de la suma apuntan a una UF disponible Y la suma coincide con la deuda o movimiento.

**Prioridad 2 (fallback sin centavos de la suma):**

```python
if uf_match is None:
    for ca in cobros_ada:
        ...
        if not all(obtener_centavos(cobros_banco[i]['importe']) == uf_num
                   for i in indices):
            continue
```

`all(...)` devuelve `True` solo si todos los elementos del iterable son verdaderos. Acá verifica que *todos* los cobros del grupo tengan los centavos de la UF (ej: todos son `.19` y la UF es 19). Si alguno no coincide, descarta.

#### Estrategia 3b: sin CUIT, con centavos — líneas 469-501

```python
from itertools import combinations
for ca in cobros_ada:
    if ca['usado'] or ca['movimiento'] == 0:
        continue
    ...
    disponibles = [idx for idx in range(len(cobros_banco))
                   if idx not in procesados and not cobros_banco[idx].get('es_qr')]
    for n in (2, 3):
        for combo in combinations(disponibles, n):
            total = sum(cobros_banco[i]['importe'] for i in combo)
            if obtener_centavos(total) != uf_num:
                continue
            if abs(total - mov_abs) <= TOLERANCIA_COBRO:
                intentar_match_cuotas(list(combo), ca, 'movimiento ADA')
```

`combinations(disponibles, n)` genera todas las combinaciones posibles de `n` elementos del listado. Si hay 5 cobros disponibles y `n=2`, genera todas las parejas posibles (10 combinaciones). Si `n=3`, genera ternas (10 combinaciones).

El `from itertools import combinations` está dentro de la función para evitar importar al nivel global algo que solo se usa acá (aunque funciona igual de bien en cualquiera de los dos lugares).

La tolerancia acá es `TOLERANCIA_COBRO` ($0,02), más estricta que la de estrategia 2 ($1,00). Esto evita que cobros espurios (como una transferencia accidental de $1) participen en sumas que casualmente cuadren con una UF.

#### Estrategia 3c: sin CUIT ni centavos, suma exacta + anchor — líneas 503-534

Igual que 3b pero sin el check de centavos de la suma. En cambio, exige que **al menos un cobro** del grupo tenga centavos coincidentes con la UF:

```python
if not any(obtener_centavos(cobros_banco[i]['importe']) == uf_num
           for i in combo):
    continue
```

`any(...)` devuelve `True` si al menos un elemento del iterable es verdadero. El "anchor" de centavos previene que una coincidencia puramente numérica (sin ninguna señal de que los pagadores intentaban pagar esa UF) genere un falso positivo.

### Estrategia 4: CENTAVOS — líneas 536-628

```python
for idx, cb in enumerate(cobros_banco):
    if idx in procesados:
        continue

    imp = imp_match(cb)
    centavos = obtener_centavos(imp)

    uf_candidata = None

    for ca in cobros_ada:
        uf_num = int(re.sub(r'\D', '', ca['uf']) or '0')
        if uf_num == centavos and not ca['usado']:
            if ca['movimiento'] == 0:
                if abs(ca['deuda_anterior'] - imp) <= TOLERANCIA_COBRO:
                    uf_candidata = ca
                break
            uf_candidata = ca
            break
```

La regla "ADA es la verdad" se implementa acá:
- Si `movimiento == 0` (ADA dice que no pagó), solo acepta el cobro si el importe es exactamente igual a la deuda (pago reciente no registrado todavía).
- El `break` después del `if` negativo es importante: si ADA dice que no pagó y el importe no coincide, `uf_candidata` queda en `None` **y no se busca otra UF**. La intención es explícita: si los centavos apuntan a esa UF pero ADA no registra pago, el cobro queda sin match.

**Fallback por importe aproximado:**

```python
if uf_candidata is None:
    imp_entero = int(imp)
    candidatos = []
    for ca in cobros_ada:
        if ca['usado'] or ca['movimiento'] == 0:
            continue
        diff_entero = abs(int(ca['deuda_anterior']) - imp_entero)
        if diff_entero <= 200:
            candidatos.append((diff_entero, ca))
    if candidatos:
        candidatos.sort(key=lambda x: x[0])
        uf_candidata = candidatos[0][1]
```

Si no hubo match por centavos, intenta un match "por importe aproximado" con tolerancia de $200 (solo para UFs con movimiento registrado). Guarda pares `(diferencia, UF)` y elige el de menor diferencia. `candidatos[0][1]` es el segundo elemento del primer par — la UF.

**Determinación del estado final:**

```python
diff_deuda = abs(uf_candidata['deuda_anterior'] - imp)
if diff_deuda <= TOLERANCIA_COBRO:
    estado = 'CONCILIADO'
    detalle = ''
else:
    mov_abs = abs(uf_candidata['movimiento'])
    if mov_abs > 0 and mov_abs > imp and abs(mov_abs - imp) > TOLERANCIA_COBRO:
        estado = 'CONCILIADO'
        detalle = f"Pago parcial: este cobro de {imp:,.2f} es parte del total..."
    else:
        estado = 'DIFERENCIA'
        ...
```

El caso `mov_abs > imp` detecta cuando ADA ya tiene registrado un monto mayor que este cobro individual — significa que hubo otros pagos previos y este es solo una parte.

**Al final:**

```python
resultados.sort(key=lambda r: r['banco']['fila_banco'])
return resultados
```

Reordena todos los resultados por fila del extracto original. Esto es necesario porque las estrategias procesan en orden distinto al del archivo.

---

## 7. Algoritmo de conciliación de gastos

### `conciliar_gastos()` — líneas 635-725

```python
def conciliar_gastos(gastos_banco, gastos_ada):
    resultados = [None] * len(gastos_banco)
```

A diferencia de cobros (que usa `resultados.append()`), acá se pre-asigna la lista con `None`s del tamaño correcto. Esto permite asignar por índice `resultados[i] = {...}` en cualquier orden.

```python
    grupos_banco = {}
    for i, gb in enumerate(gastos_banco):
        clave = round(abs(gb['importe']), 2)
        grupos_banco.setdefault(clave, []).append(i)
```

Agrupa los gastos del banco por importe. `abs()` convierte negativos a positivos (los gastos vienen negativos del extracto). `round(..., 2)` redondea a 2 decimales para que `$5000.0` y `$5000.00` sean la misma clave.

```python
    for clave, indices in grupos_banco.items():
        candidatos_ada = [g for g in gastos_ada
                          if not g['usado'] and abs(g['importe'] - clave) <= TOLERANCIA_GASTO]
```

*List comprehension*: `[expresión for item in iterable if condición]`. Crea una nueva lista filtrando. Acá: todos los gastos ADA no usados cuyo importe esté dentro de la tolerancia.

**Caso 2 (N en banco = N en ADA):**

```python
        if len(indices) == len(candidatos_ada):
            indices_ord = sorted(indices, key=lambda i: gastos_banco[i]['fecha'] or datetime.min)
            def orden_ada(g):
                if g['fecha_fact'] and hasattr(g['fecha_fact'], 'year'):
                    return (0, g['fecha_fact'])
                return (1, g['id_gasto'])
            candidatos_ord = sorted(candidatos_ada, key=orden_ada)

            for i, ga in zip(indices_ord, candidatos_ord):
```

`zip(lista1, lista2)` combina dos listas elemento a elemento. `for i, ga in zip(...)` itera sobre pares `(índice_banco, gasto_ada)`.

La función `orden_ada` devuelve una tupla para ordenar: primero los que tienen fecha real `(0, fecha)`, luego los que solo tienen ID `(1, id)`. Python compara tuplas lexicográficamente (primero el primer elemento, luego el segundo si el primero es igual).

`hasattr(g['fecha_fact'], 'year')` verifica si el campo fecha es un objeto `datetime` real (que tiene `.year`) o solo un string vacío/texto.

---

## 8. Generación de salidas

### `marcar_extracto()` — líneas 728-752

```python
def marcar_extracto(ruta_extracto, resultados_cobros, resultados_gastos, ruta_salida):
    shutil.copy2(ruta_extracto, ruta_salida)
    wb = load_workbook(ruta_salida)
    ws = wb.active
```

`shutil.copy2()` copia el archivo preservando metadatos (fecha de modificación). Se trabaja sobre la copia, nunca sobre el original.

```python
    por_fila = {}
    for r in resultados_cobros + resultados_gastos:
        por_fila[r['banco']['fila_banco']] = r
```

Crea un diccionario indexado por número de fila del extracto. `resultados_cobros + resultados_gastos` concatena las dos listas (funciona con listas en Python).

```python
    for fila, res in por_fila.items():
        ...
        for col in range(1, 9):
            ws.cell(row=fila, column=col).fill = color
```

`range(1, 9)` genera números del 1 al 8 (el 9 es exclusivo). Pinta las columnas A a H de cada fila.

### `generar_reporte()` — líneas 755-952

La función crea un `Workbook` nuevo con 4 hojas. El patrón se repite para cada hoja:

```python
ws = wb.active
ws.title = "Cobros"
headers = ['Estado', 'Fecha Banco', ...]
for col, h in enumerate(headers, 1):
    c = ws.cell(row=1, column=col, value=h)
    c.fill = FILL_HEADER
    c.font = FONT_HEADER
    c.alignment = Alignment(horizontal='center')
```

`enumerate(headers, 1)` devuelve pares `(1, 'Estado')`, `(2, 'Fecha Banco')`, etc. `ws.cell(row, column, value)` escribe en la celda y devuelve el objeto celda para poder aplicarle estilos.

```python
for i, r in enumerate(resultados_cobros, 2):
    b = r['banco']
    a = r['ada']
    ws.cell(row=i, column=2, value=b['fecha']).number_format = 'DD/MM/YYYY'
```

El método `.number_format` formatea la celda. Sin esto, las fechas se mostrarían como números seriales de Excel (ej: 45400) en vez de `15/05/2024`.

```python
    ws.cell(row=i, column=7, value=a['uf'] if a else '')
```

`a['uf'] if a else ''` es un operador ternario. Si `a` es `None` (cobro sin match), pone string vacío en vez de lanzar un `TypeError` por intentar acceder a `None['uf']`.

**Hoja 4: Resumen** se crea con `wb.create_sheet("Resumen", 0)` — el `0` la inserta en la primera posición (tab más a la izquierda en Excel).

```python
ws4 = wb.create_sheet("Resumen", 0)
...
ok_cob = sum(1 for r in resultados_cobros if r['estado'] == 'CONCILIADO')
```

`sum(1 for r in ... if condición)` es una forma idiomática Python de contar cuántos elementos cumplen una condición. Equivale a `len([r for r in ... if condición])` pero sin crear la lista intermedia.

---

## 9. Función `main()`

```python
def main():
    print("=" * 65)
```

`"=" * 65` repite el string 65 veces. Un truco simple para dibujar separadores en consola.

```python
    if not Path(EXTRACTO_BANCO).exists():
        print(f"\n❌ Archivo no encontrado...")
        input("\nPresioná Enter para cerrar...")
        sys.exit(1)
```

`input()` pausa la ejecución hasta que el usuario presiona Enter. Es necesario porque el `.exe` de Windows cierra la ventana de consola inmediatamente al terminar — sin este `input()`, el usuario no podría leer el mensaje de error.

```python
    INFORME_GASTOS = buscar_archivo_por_prefijo(
        CARPETA_ADA4, "InformeGastos", [".xls", ".xlsx"])
```

Las variables `INFORME_GASTOS` y `COBRANZA_RAPIDA` se definen localmente en `main()` (no son globales) porque su valor se determina en tiempo de ejecución buscando archivos.

```python
    ruta_gastos = INFORME_GASTOS
    if INFORME_GASTOS.lower().endswith('.xls'):
        print("\nConvirtiendo informe de gastos (.xls → .xlsx)...")
        try:
            ruta_gastos = convertir_xls(INFORME_GASTOS)
```

Si el informe de gastos es `.xls` (el formato viejo de Excel), lo convierte antes de leerlo. `openpyxl` solo entiende `.xlsx`.

```python
    res_cobros = conciliar_cobros(cobros_banco, cobros_ada)
    res_gastos = conciliar_gastos(gastos_banco, gastos_ada)
```

Las listas `cobros_ada` y `gastos_ada` se pasan a las funciones y los campos `usado` se modifican dentro. En Python, los objetos mutables (como listas y diccionarios) se pasan por referencia — las funciones modifican los mismos diccionarios que vive en `main()`, no copias.

```python
if __name__ == "__main__":
    main()
```

Este bloque al final del archivo es un patrón estándar de Python. `__name__` es una variable especial que vale `"__main__"` cuando el script se ejecuta directamente (`python conciliacion.py`) y vale el nombre del módulo cuando se importa desde otro script. Esto permite que el archivo sea importable como módulo sin ejecutar `main()` automáticamente.

---

## 10. Flujo completo de datos

```
DISCO
  extracto.xlsx ─────────────────────────────────────┐
  InformeGastos_*.xls  →  [convertir_xls]  →  .xlsx ─┤
  CobranzaRapida_*.xlsx ─────────────────────────────┘
                                                       │
                          [leer_extracto]  ←───────────┤
                          [leer_gastos_ada] ←──────────┤
                          [leer_cobros_ada] ←──────────┘
                                │
                 Lista de dicts  │  Lista de dicts
                 (cobros_banco)  │  (cobros_ada, gastos_ada)
                                │
              ┌─────────────────┴─────────────────┐
              │                                   │
    [conciliar_cobros]               [conciliar_gastos]
    5 estrategias                    Match por importe
    set `procesados`                 1-a-1 por fecha
    campo `usado` en ADA             campo `usado` en ADA
              │                                   │
         res_cobros                          res_gastos
         (lista de resultados               (lista de resultados
          con estado y detalle)              con estado y detalle)
              │                                   │
              └─────────────────┬─────────────────┘
                                │
                 ┌──────────────┴──────────────┐
                 │                             │
       [marcar_extracto]            [generar_reporte]
       Copia + pinta colores        4 hojas Excel
                 │                             │
         Extracto_Conciliado_        ConciliacionBancaria_
         YYYY-MM-DD_HH-MM.xlsx       YYYY-MM-DD_HH-MM.xlsx
```

### Tipos de datos que fluyen

Cada elemento de `cobros_banco` es un diccionario como este:
```python
{
    'fila_banco':  3,
    'fecha':       datetime(2024, 5, 15),
    'comprobante': '0012345',
    'concepto':    '100',
    'descripcion': 'TRANSF.INTER.DISTINTO TIT.-20289692607-CORNIGLIO/FEDERICO',
    'importe':     179304.07,
    'tipo':        'cobro',
    'es_qr':       False,
    'bruto_qr':    None,
}
```

Cada elemento de `cobros_ada` es:
```python
{
    'uf':            '0007',
    'depto':         '3',
    'piso':          'B',
    'propietario':   'CORNIGLIO FEDERICO',
    'deuda_anterior': 179304.07,
    'movimiento':    -179304.07,   # negativo = pagó
    'usado':         False,        # se vuelve True cuando matchea
}
```

Cada elemento de `res_cobros` (resultado) es:
```python
{
    'banco':   { ... },             # el dict del banco
    'ada':     { ... } | None,      # el dict de ADA (None si SIN MATCH)
    'estado':  'CONCILIADO',        # o PAGO EN CUOTAS, DIFERENCIA, SIN MATCH
    'detalle': 'Pago parcial: ...',  # string vacío si está todo bien
    'cuit':    '20289692607',        # o None si no había CUIT en la descripción
}
```

---

## Conceptos Python usados en este código

| Concepto | Ejemplo en el código |
|----------|---------------------|
| List comprehension | `[m for m in movimientos if m['tipo'] == 'cobro']` |
| Dict comprehension / set comprehension | `{extraer_cuit(cobros_banco[i]['descripcion']) for i in indices}` |
| Generator expression | `sum(1 for r in resultados_cobros if r['estado'] == 'CONCILIADO')` |
| Función anidada | `imp_match()` y `intentar_match_cuotas()` dentro de `conciliar_cobros()` |
| Operador ternario | `a['uf'] if a else ''` |
| Unpacking | `fecha, comp, conc, desc, imp = row[0], row[1], row[2], row[3], row[4]` |
| `or` como valor por defecto | `str(row[1] or '').strip()` |
| `set` para búsqueda O(1) | `if idx in procesados` |
| `enumerate` con inicio | `enumerate(headers, 1)` |
| `zip` para iterar pares | `for i, ga in zip(indices_ord, candidatos_ord)` |
| f-strings | `f"Pago parcial: pagó {movimiento_abs:,.2f} de {ca['deuda_anterior']:,.2f}"` |
| Formato numérico en f-string | `:,.2f` → separador de miles y 2 decimales |
| `pathlib.Path` | `Path(carpeta) / (xls.stem + '.xlsx')` |
| `raise` y `try/except` | En `convertir_xls()` y `main()` |
| `__name__ == "__main__"` | Bloque final del archivo |
