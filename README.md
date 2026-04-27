# Conciliador Bancario ADA4

Script Python que automatiza la conciliación bancaria diaria para administraciones de consorcios en Buenos Aires. Cruza el extracto del banco contra los registros del sistema de gestión ADA4 y produce un reporte coloreado listo para revisar.

---

## Qué hace

Toma tres archivos de entrada:

| Archivo | Origen | Contenido |
|---------|--------|-----------|
| `extracto.xlsx` | Home banking | Movimientos del día (cobros y gastos) |
| `InformeGastos_*.xls` | ADA4 | Gastos cargados en el sistema |
| `CobranzaRapida_*.xlsx` | ADA4 | Deudas y pagos registrados por unidad funcional |

Y produce dos archivos de salida:

| Archivo | Contenido |
|---------|-----------|
| `ConciliacionBancaria_YYYY-MM-DD_HH-MM.xlsx` | Reporte detallado con 4 hojas: Cobros, Gastos, Pendientes ADA4, Resumen |
| `Extracto_Conciliado_YYYY-MM-DD_HH-MM.xlsx` | Copia del extracto con cada fila pintada según su estado |

### Colores del reporte

| Color | Estado | Significado |
|-------|--------|-------------|
| Verde | CONCILIADO | Match confirmado contra ADA4 |
| Naranja | PAGO EN CUOTAS | Varios cobros agrupados a una misma unidad funcional |
| Amarillo | DIFERENCIA / AMBIGUO | Match con desvío o varias coincidencias posibles |
| Rojo | SIN MATCH | No se encontró correspondencia |

---

## Cómo funciona

### Cobros

Aplica 5 estrategias en orden de prioridad. Una vez que un cobro queda resuelto no se reutiliza:

1. **QR** — matching por importe bruto (la comisión bancaria se descuenta aparte).
2. **Match único** — el importe coincide con el movimiento registrado en ADA4 para una UF.
3. **Cuotas** — 2 o 3 cobros cuya suma coincide con la deuda de una UF:
   - Mismo CUIT pagador con centavos de la suma apuntando a la UF.
   - CUITs distintos con centavos exactos + tolerancia estricta.
   - CUITs distintos con suma exacta y al menos un cobro con centavos coincidentes (anchor).
4. **Centavos** — los centavos del importe identifican la UF (truco local: `.07` → UF 0007).
5. **Sin match** — el cobro no tiene contraparte en ADA4.

**Filosofía clave:** ADA4 es la fuente autoritativa. Si ADA dice que una unidad no pagó, el script no le asigna ningún cobro del banco aunque los centavos coincidan, salvo que el importe sea exactamente igual a la deuda (pago reciente que ADA aún no procesó).

### Gastos

Match por importe con tolerancia de $0,50. Si hay N gastos idénticos en el banco y N en ADA4, se emparejan 1 a 1 por fecha. Si las cantidades difieren, se marca como AMBIGUO.

---

## Requisitos

- **Python 3.10+**
- **openpyxl** — lectura y escritura de `.xlsx`
- **LibreOffice** — para convertir el `InformeGastos_*.xls` a `.xlsx` (se invoca automáticamente)

```bash
pip install openpyxl
```

LibreOffice debe estar instalado en `C:\Program Files\LibreOffice` o `C:\Program Files (x86)\LibreOffice`.

---

## Uso

1. Colocá los tres archivos de entrada en la misma carpeta que el script (o el `.exe`):
   - `extracto.xlsx`
   - `InformeGastos_*.xls` (o `.xlsx`)
   - `CobranzaRapida_*.xlsx`

2. Ejecutá el script:

```bash
python conciliacion.py
```

El script imprime un resumen en consola y deja los dos archivos de salida en la misma carpeta.

---

## Compilar a .exe

```bash
pip install pyinstaller
pyinstaller conciliacion.spec
```

El `.exe` resultante no requiere Python instalado. LibreOffice sigue siendo necesario en la máquina donde corra.

---

## Estructura del proyecto

```
conciliacion.py       # Script principal
conciliacion.spec     # Configuración de PyInstaller
CLAUDE.md             # Documentación técnica detallada para desarrollo
```

---

## Tolerancias de matching

| Concepto | Valor | Motivo |
|----------|-------|--------|
| Cobros individuales | $0,02 | Deben ser muy exactos |
| Sumas de cuotas | $1,00 | Acumulación de redondeos centavo a centavo |
| Gastos | $0,50 | Redondeos en liquidaciones de proveedores |

Estos valores fueron calibrados con datasets reales. No conviene modificarlos sin un caso concreto que lo justifique.
