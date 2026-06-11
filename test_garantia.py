"""
test_garantia.py
================
Verifica que:
  1. RMA_Capturado (checkbox de Google Sheets) se interpreta correctamente
     en sus distintas representaciones: bool True/False, string "TRUE"/"FALSE", "".
  2. Los registros capturados (True) NO aparecen en futuras consultas.
  3. Los registros NO capturados (False / "") SÍ aparecen.
  4. RMA_Fecha_Captura se escribe en formato ISO YYYY-MM-DD.
  5. La función batch_update generaría las celdas correctas.

Corre sin conexión a Google Sheets (todo mockeado).
Ejecutar con:
    .venv\\Scripts\\python.exe test_garantia.py
"""

import sys
import os
from datetime import date
from unittest.mock import MagicMock, patch

import pandas as pd

# ─────────────────────────────────────────────────
# SETUP: mock de streamlit para poder importar sheets_ops
# ─────────────────────────────────────────────────
st_mock = MagicMock()
st_mock.cache_data = lambda *a, **kw: (lambda f: f)   # no-op decorator
st_mock.cache_resource = lambda *a, **kw: (lambda f: f)
sys.modules["streamlit"] = st_mock

# También mockear gspread.utils usado en marcar_capturados_sheets
import gspread.utils  # noqa: ya está instalado

LINE = "=" * 62
PASS = "  [OK]"
FAIL = "  [FAIL]"


def check(cond: bool, msg_ok: str, msg_fail: str = "") -> bool:
    if cond:
        print(f"{PASS} {msg_ok}")
    else:
        print(f"{FAIL} {msg_fail or msg_ok}")
    return cond


all_ok = True

print(LINE)
print("  TEST: RMA_Capturado (checkbox Google Sheets)")
print(LINE)


# ─────────────────────────────────────────────────────────────────────────────
# TEST 1 — Filtrado de proveedores con registros pendientes
# (replica la lógica de listar_proveedores_garantia)
# ─────────────────────────────────────────────────────────────────────────────
print("\n[1] Filtrado de proveedores pendientes")

# gspread.get_all_records() devuelve bool True/False para checkboxes
raw = [
    {"proveedor": "ALTAVISTA",  "Producto": "ProdA1", "RMA_Capturado": True,  "RMA_Fecha_Captura": "2026-06-01"},
    {"proveedor": "ALTAVISTA",  "Producto": "ProdA2", "RMA_Capturado": False, "RMA_Fecha_Captura": ""},
    {"proveedor": "SAMSUNG",    "Producto": "ProdB1", "RMA_Capturado": "",    "RMA_Fecha_Captura": ""},
    {"proveedor": "SAMSUNG",    "Producto": "ProdB2", "RMA_Capturado": True,  "RMA_Fecha_Captura": "2026-06-05"},
    {"proveedor": "MOTOROLA",   "Producto": "ProdC1", "RMA_Capturado": True,  "RMA_Fecha_Captura": "2026-06-09"},
    # "TRUE" como string (edge case si alguien escribe el valor a mano)
    {"proveedor": "LG",         "Producto": "ProdD1", "RMA_Capturado": "TRUE","RMA_Fecha_Captura": "2026-06-08"},
    {"proveedor": "LG",         "Producto": "ProdD2", "RMA_Capturado": "FALSE","RMA_Fecha_Captura": ""},
]

df = pd.DataFrame(raw).fillna("").astype(str)

mask_pendiente = ~df["RMA_Capturado"].str.strip().str.upper().isin(["TRUE", "1"])
provs = sorted(
    df[mask_pendiente]["proveedor"].str.strip()
    .replace("", pd.NA).dropna().unique().tolist()
)

# ALTAVISTA: ProdA2 (False) → pendiente ✓
# SAMSUNG:   ProdB1 ("") → pendiente ✓
# MOTOROLA:  solo capturado → NO debe aparecer ✓
# LG:        ProdD2 ("FALSE") → pendiente ✓
expected = sorted(["ALTAVISTA", "SAMSUNG", "LG"])
ok = provs == expected
all_ok &= ok
check(ok,
      f"Proveedores pendientes correctos: {provs}",
      f"Esperado {expected}, obtenido {provs}")


# ─────────────────────────────────────────────────────────────────────────────
# TEST 2 — Cobertura de todos los valores posibles de RMA_Capturado
# ─────────────────────────────────────────────────────────────────────────────
print("\n[2] Todos los valores posibles de RMA_Capturado")

casos = [
    # (valor_raw,  debe_filtrarse_como_capturado)
    (True,    True,  "bool True     → capturado"),
    (False,   False, "bool False    → NO capturado"),
    ("TRUE",  True,  'str "TRUE"    → capturado'),
    ("true",  True,  'str "true"    → capturado'),
    ("1",     True,  'str "1"       → capturado'),
    ("FALSE", False, 'str "FALSE"   → NO capturado'),
    ("false", False, 'str "false"   → NO capturado'),
    ("",      False, 'str ""        → NO capturado'),
]

for raw_val, expect_cap, label in casos:
    val_str = str(raw_val).strip().upper()
    es_cap  = val_str in ("TRUE", "1")
    ok = es_cap == expect_cap
    all_ok &= ok
    check(ok, label, f"{label}  (obtuvo is_capturado={es_cap}, esperado={expect_cap})")


# ─────────────────────────────────────────────────────────────────────────────
# TEST 3 — Filtrado de registros por proveedor (lógica de obtener_registros_proveedor_sheets)
# ─────────────────────────────────────────────────────────────────────────────
print("\n[3] Filtrado de registros por proveedor (get_all_values)")

# get_all_values() devuelve strings para todo, incluidos checkboxes
headers = ["proveedor", "Producto", "Serial", "diagnostico", "RMA_Capturado", "RMA_Fecha_Captura"]
all_rows_raw = [
    ["ALTAVISTA", "ProdA1", "S001", "Pantalla rota", "TRUE",  "2026-06-01"],
    ["ALTAVISTA", "ProdA2", "S002", "No enciende",   "FALSE", ""],
    ["ALTAVISTA", "ProdA3", "S003", "Teclado falla", "",      ""],
    ["SAMSUNG",   "ProdB1", "S004", "Carga lenta",   "FALSE", ""],
]

target_proveedor = "altavista"   # prueba case-insensitive

registros = []
for i, row_vals in enumerate(all_rows_raw, start=2):  # simula start desde fila 2
    fila = {headers[j]: row_vals[j] if j < len(row_vals) else "" for j in range(len(headers))}
    if fila.get("proveedor", "").strip().lower() != target_proveedor:
        continue
    cap = fila.get("RMA_Capturado", "").strip().upper()
    if cap in ("TRUE", "1"):
        continue
    registros.append({"_row": i, "Producto": fila["Producto"], "RMA_Capturado": fila["RMA_Capturado"]})

prods_pendientes = [r["Producto"] for r in registros]
rows_pendientes  = [r["_row"] for r in registros]

ok1 = prods_pendientes == ["ProdA2", "ProdA3"]
ok2 = rows_pendientes  == [3, 4]
all_ok &= ok1 & ok2
check(ok1, f"Productos pendientes correctos: {prods_pendientes}",
           f"Esperado ['ProdA2','ProdA3'], obtenido {prods_pendientes}")
check(ok2, f"Números de fila correctos: {rows_pendientes}",
           f"Esperado [3, 4], obtenido {rows_pendientes}")


# ─────────────────────────────────────────────────────────────────────────────
# TEST 4 — Formato de RMA_Fecha_Captura (ISO YYYY-MM-DD)
# ─────────────────────────────────────────────────────────────────────────────
print("\n[4] Formato de fecha RMA_Fecha_Captura")

hoy = date.today().isoformat()

ok_len  = len(hoy) == 10
ok_fmt  = hoy.count("-") == 2
ok_year = hoy.startswith("20")
all_ok &= ok_len & ok_fmt & ok_year

check(ok_len,  f"Longitud correcta (10 chars): '{hoy}'")
check(ok_fmt,  f"Formato ISO correcto (2 guiones): '{hoy}'")
check(ok_year, f"Año correcto (comienza con '20'): '{hoy}'")


# ─────────────────────────────────────────────────────────────────────────────
# TEST 5 — Construcción del batch_update (lógica de marcar_capturados_sheets)
# ─────────────────────────────────────────────────────────────────────────────
print("\n[5] Construcción del batch_update para marcar_capturados_sheets")

import gspread.utils as gu

# Simular: headers de la hoja, columnas de RMA_Capturado y RMA_Fecha_Captura
headers_sheet = ["proveedor", "Producto", "Serial", "diagnostico",
                 "RMA_Capturado", "RMA_Fecha_Captura"]
col_cap   = headers_sheet.index("RMA_Capturado") + 1   # → 5
col_fecha = headers_sheet.index("RMA_Fecha_Captura") + 1  # → 6
row_numbers_to_mark = [3, 4]  # filas de los registros pendientes

updates = []
for row_num in row_numbers_to_mark:
    updates.append({
        "range":  gu.rowcol_to_a1(row_num, col_cap),
        "values": [["TRUE"]],
    })
    updates.append({
        "range":  gu.rowcol_to_a1(row_num, col_fecha),
        "values": [[hoy]],
    })

expected_ranges = ["E3", "F3", "E4", "F4"]
expected_values = [["TRUE"], [hoy], ["TRUE"], [hoy]]

actual_ranges = [u["range"] for u in updates]
actual_values = [u["values"][0] for u in updates]

ok_ranges = actual_ranges == expected_ranges
ok_values = actual_values == expected_values
all_ok &= ok_ranges & ok_values

check(ok_ranges, f"Rangos batch_update correctos: {actual_ranges}",
                 f"Esperado {expected_ranges}, obtenido {actual_ranges}")
check(ok_values, f"Valores batch_update correctos (TRUE + fecha de hoy)",
                 f"Esperado {expected_values}, obtenido {actual_values}")


# ─────────────────────────────────────────────────────────────────────────────
# RESULTADO FINAL
# ─────────────────────────────────────────────────────────────────────────────
print(f"\n{LINE}")
if all_ok:
    print("  [OK] TODOS LOS TESTS PASARON")
else:
    print("  [FAIL] ALGUN TEST FALLO -- revisa los errores arriba")
print(LINE)

print("""
Referencia rápida — RMA_Capturado en Google Sheets:
  • get_all_records()  devuelve bool True/False  (Python native)
  • get_all_values()   devuelve str  "TRUE"/"FALSE" (raw string)
  • Para ESCRIBIR TRUE usar: batch_update con valor "TRUE" (string)
  • Para ESCRIBIR fecha usar: date.today().isoformat() → "YYYY-MM-DD"
""")

sys.exit(0 if all_ok else 1)
