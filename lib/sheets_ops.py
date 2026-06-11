"""
sheets_ops.py
Operaciones de lectura/escritura sobre Google Sheets.
Usa @st.cache_data para reducir llamadas a la API en recargas de página.
"""
import streamlit as st
import pandas as pd
import gspread

from lib.google_client import get_google_clients

SHEET_RMA = "RMA_Proveedores"
SHEET_PROV = "Proveedores"

COLUMNAS_RMA = [
    "ID_Registro",
    "Proveedor",
    "Fecha_carga",
    "Cantidad_Articulos",
    "Link_Excel_Detalle",
    "Estado",
    "Numero_Envio",
    "Fecha_Envio",
    "Numero_Caso_RMA",
    "Fecha_Recepcion",
    "Resolucion",
    "Detalle_Resolucion",
    "Fecha_Resolucion",
    "Link_Carpeta_Drive",
    "Links_Adjuntos",
]


# ═══════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════

def _open_worksheet(sheet_id: str, sheet_name: str) -> gspread.Worksheet:
    _, gc = get_google_clients()
    sh = gc.open_by_key(sheet_id)
    try:
        return sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=2000, cols=len(COLUMNAS_RMA))
        ws.append_row(COLUMNAS_RMA, value_input_option="USER_ENTERED")
        return ws


def _ensure_header(ws: gspread.Worksheet) -> list[str]:
    """Asegura que la hoja tenga encabezado. Retorna la lista de headers."""
    first_row = ws.row_values(1)
    if not first_row or first_row[0] != "ID_Registro":
        ws.insert_row(COLUMNAS_RMA, index=1, value_input_option="USER_ENTERED")
        return COLUMNAS_RMA
    return first_row


# ═══════════════════════════════════════════════════════
#  LECTURA CON CACHE
# ═══════════════════════════════════════════════════════

@st.cache_data(ttl=60, show_spinner=False)
def leer_registros(estado: str, sheet_id: str) -> pd.DataFrame:
    """
    Lee registros de RMA_Proveedores filtrados por Estado.
    Cache de 60 segundos para reducir llamadas a la API.
    """
    ws = _open_worksheet(sheet_id, SHEET_RMA)
    data = ws.get_all_records(default_blank="")
    if not data:
        return pd.DataFrame(columns=COLUMNAS_RMA)
    df = pd.DataFrame(data)
    for col in COLUMNAS_RMA:
        if col not in df.columns:
            df[col] = ""
    df = df.fillna("").astype(str)
    if estado:
        df = df[df["Estado"].str.strip() == estado]
    return df.reset_index(drop=True)


@st.cache_data(ttl=300, show_spinner=False)
def leer_proveedores(sheet_id: str) -> list:
    """
    Lee la hoja 'Proveedores' del mismo spreadsheet.
    Cache de 5 minutos (cambian poco).
    Retorna lista de dicts con al menos la clave 'nombre'.
    """
    try:
        _, gc = get_google_clients()
        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet(SHEET_PROV)
        data = ws.get_all_records(default_blank="")
        return data
    except Exception:
        return []


# ═══════════════════════════════════════════════════════
#  ESCRITURA
# ═══════════════════════════════════════════════════════

def upsert_registro(id_registro: str, campos: dict, sheet_id: str) -> None:
    """
    Inserta o actualiza una fila por ID_Registro.
    Si ya existe la fila, actualiza solo los campos indicados.
    Si no existe, agrega una fila nueva con todos los campos.
    """
    ws = _open_worksheet(sheet_id, SHEET_RMA)
    headers = _ensure_header(ws)

    all_values = ws.get_all_values()
    if not all_values:
        all_values = [headers]

    id_col_idx = headers.index("ID_Registro") if "ID_Registro" in headers else 0

    # Buscar fila existente
    target_row = None
    for i, row in enumerate(all_values[1:], start=2):
        cell_val = row[id_col_idx] if len(row) > id_col_idx else ""
        if cell_val == id_registro:
            target_row = i
            break

    if target_row:
        # Actualizar campos existentes (batch para reducir llamadas)
        updates = []
        for col_name, value in campos.items():
            if col_name in headers:
                col_num = headers.index(col_name) + 1
                updates.append({
                    "range": gspread.utils.rowcol_to_a1(target_row, col_num),
                    "values": [[str(value) if value is not None else ""]],
                })
        if updates:
            ws.batch_update(updates, value_input_option="USER_ENTERED")
    else:
        # Nueva fila
        nueva_fila = [""] * len(headers)
        nueva_fila[id_col_idx] = id_registro
        for col_name, value in campos.items():
            if col_name in headers:
                nueva_fila[headers.index(col_name)] = str(value) if value is not None else ""
        ws.append_row(nueva_fila, value_input_option="USER_ENTERED")


def actualizar_registro(id_registro: str, campos: dict, sheet_id: str) -> None:
    """Alias semántico de upsert para actualizaciones parciales."""
    upsert_registro(id_registro, campos, sheet_id)


# ═══════════════════════════════════════════════════════
#  LIMPIAR CACHE
# ═══════════════════════════════════════════════════════

def clear_cache() -> None:
    leer_registros.clear()
    leer_proveedores.clear()
    leer_productos_garantia.clear()


# ═══════════════════════════════════════════════════════
#  GARANTÍA — Sección migrada desde Airtable
#  Hoja: "RMA ALTAVISTA"  (misma que PRODUCTOS_SHEET_NAME en app.py)
#  Columnas esperadas:
#    proveedor, Producto, Serial, diagnostico,
#    RMA_Capturado (TRUE/FALSE/""), RMA_Fecha_Captura (YYYY-MM-DD/"")
# ═══════════════════════════════════════════════════════

SHEET_GARANTIA = "RMA ALTAVISTA"

COLUMNAS_GARANTIA = [
    "proveedor",
    "Producto",
    "Serial",
    "diagnostico",
    "RMA_Capturado",
    "RMA_Fecha_Captura",
]


@st.cache_data(ttl=120, show_spinner=False)
def leer_productos_garantia(sheet_id: str) -> pd.DataFrame:
    """
    Lee todos los registros de la hoja 'RMA ALTAVISTA'.
    Retorna un DataFrame con las columnas de COLUMNAS_GARANTIA disponibles.
    Cache de 2 minutos.
    """
    try:
        _, gc = get_google_clients()
        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet(SHEET_GARANTIA)
        data = ws.get_all_records(default_blank="")
        if not data:
            return pd.DataFrame(columns=COLUMNAS_GARANTIA)
        df = pd.DataFrame(data)
        # Asegurar columnas mínimas
        for col in COLUMNAS_GARANTIA:
            if col not in df.columns:
                df[col] = ""
        return df.fillna("").astype(str)
    except Exception:
        return pd.DataFrame(columns=COLUMNAS_GARANTIA)


def listar_proveedores_garantia(sheet_id: str) -> list[str]:
    """
    Retorna lista ordenada de proveedores únicos con registros pendientes
    (RMA_Capturado vacío o distinto de TRUE/True/1).
    """
    df = leer_productos_garantia(sheet_id)
    if df.empty or "proveedor" not in df.columns:
        return []
    # Filtrar no capturados
    mask_no_capturado = ~df["RMA_Capturado"].str.strip().str.upper().isin(["TRUE", "1"])
    proveedores = (
        df[mask_no_capturado]["proveedor"]
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    return sorted(proveedores)


def obtener_registros_proveedor_sheets(
    sheet_id: str, proveedor: str
) -> list[dict]:
    """
    Retorna todos los registros de 'RMA ALTAVISTA' donde:
      - proveedor == proveedor (comparación case-insensitive, trimmed)
      - RMA_Capturado no es TRUE/1
    Cada elemento incluye: {"_row": <row_number>, "Producto", "Serial", "diagnostico", "proveedor"}
    """
    try:
        _, gc = get_google_clients()
        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet(SHEET_GARANTIA)
        all_values = ws.get_all_values()
        if not all_values:
            return []
        headers = all_values[0]
        registros = []
        prov_norm = proveedor.strip().lower()
        for i, row in enumerate(all_values[1:], start=2):
            fila = {headers[j]: row[j] if j < len(row) else "" for j in range(len(headers))}
            if fila.get("proveedor", "").strip().lower() != prov_norm:
                continue
            capturado = fila.get("RMA_Capturado", "").strip().upper()
            if capturado in ("TRUE", "1"):
                continue
            registros.append({
                "_row":        i,
                "Producto":    fila.get("Producto", "").strip(),
                "Serial":      fila.get("Serial", "").strip(),
                "diagnostico": fila.get("diagnostico", "").strip(),
                "proveedor":   fila.get("proveedor", "").strip(),
            })
        return registros
    except Exception:
        return []


def marcar_capturados_sheets(
    sheet_id: str, row_numbers: list[int]
) -> tuple[int, list[str]]:
    """
    Marca como capturados los registros en las filas indicadas.
    Escribe RMA_Capturado = TRUE y RMA_Fecha_Captura = hoy (YYYY-MM-DD).
    Retorna (cantidad_actualizados, lista_de_errores).
    """
    from datetime import date
    errores: list[str] = []
    actualizados = 0
    try:
        _, gc = get_google_clients()
        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet(SHEET_GARANTIA)
        headers = ws.row_values(1)
        try:
            col_cap  = headers.index("RMA_Capturado") + 1
            col_fecha = headers.index("RMA_Fecha_Captura") + 1
        except ValueError as e:
            return 0, [f"Columna no encontrada: {e}"]

        hoy = date.today().isoformat()
        updates = []
        for row_num in row_numbers:
            updates.append({
                "range": gspread.utils.rowcol_to_a1(row_num, col_cap),
                "values": [["TRUE"]],
            })
            updates.append({
                "range": gspread.utils.rowcol_to_a1(row_num, col_fecha),
                "values": [[hoy]],
            })
        if updates:
            ws.batch_update(updates, value_input_option="USER_ENTERED")
            actualizados = len(row_numbers)
    except Exception as exc:
        errores.append(str(exc))
    return actualizados, errores
