"""
drive_ops.py
Operaciones sobre Google Drive con las siguientes optimizaciones:

1. BATCH QUERY: listar_excels_todas_carpetas() reemplaza N llamadas individuales
   por UNA sola llamada a la API que trae todos los archivos de todas las carpetas
   de proveedor de una vez.

2. COUNT CACHE: contar_articulos_excel_cached() guarda el conteo de filas
   en un dict en memoria (nivel módulo), indexado por (file_id + modifiedTime).
   Si el archivo no cambió, no se vuelve a descargar.

3. FILTRO DE SUBCARPETAS: La query de Drive usa '{folder_id}' in parents,
   que por diseño NO es recursiva. Los archivos en /procesado o /finalizado
   tienen esos IDs como parent, no el folder raíz del proveedor.
   => La Tabla 1 ve solo archivos sueltos en la carpeta del proveedor.

4. PARSEO ROBUSTO DE LINKS: extraer_file_id_de_link y extraer_folder_id_de_link
   manejan todos los formatos de URL de Google Drive.
"""

import io
import json
import os
import re
from datetime import datetime

import pandas as pd
from googleapiclient.http import MediaIoBaseDownload

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ── Cache en memoria (sobrevive recargas de página, se pierde al reiniciar el server) ──
_COUNT_CACHE: dict[str, int | None] = {}

# ── Cache persistente en disco (sobrevive reinicios del server) ──
_DISK_CACHE_PATH = "/tmp/rma_count_cache.json"


def _load_disk_cache() -> dict:
    try:
        if os.path.exists(_DISK_CACHE_PATH):
            with open(_DISK_CACHE_PATH) as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _save_disk_cache(cache: dict) -> None:
    try:
        with open(_DISK_CACHE_PATH, "w") as f:
            json.dump(cache, f)
    except Exception:
        pass


# Precarga el cache de disco al importar el módulo
_DISK_CACHE: dict = _load_disk_cache()


# ═══════════════════════════════════════════════════════
#  PARSEO DE LINKS
# ═══════════════════════════════════════════════════════

def extraer_file_id_de_link(link: str) -> str:
    """
    Extrae el file ID de cualquier formato de URL de Google Drive/Docs.
    Formatos soportados:
      - https://drive.google.com/file/d/{ID}/view
      - https://docs.google.com/spreadsheets/d/{ID}/edit
      - https://docs.google.com/document/d/{ID}/edit
      - https://drive.google.com/open?id={ID}
      - https://drive.google.com/uc?id={ID}
    """
    if not link or not isinstance(link, str):
        return ""
    link = link.strip()

    # Patrón /d/{ID}/ — cubre file, spreadsheets, document, presentation
    m = re.search(r"/d/([a-zA-Z0-9_-]{10,})", link)
    if m:
        return m.group(1)

    # Patrón ?id={ID} o &id={ID}
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]{10,})", link)
    if m:
        return m.group(1)

    return ""


def extraer_folder_id_de_link(link: str) -> str:
    """
    Extrae el folder ID de una URL de carpeta de Google Drive.
    Formatos soportados:
      - https://drive.google.com/drive/folders/{ID}
      - https://drive.google.com/drive/u/0/folders/{ID}
      - https://drive.google.com/open?id={ID}
    """
    if not link or not isinstance(link, str):
        return ""
    link = link.strip()

    # /folders/{ID}
    m = re.search(r"/folders/([a-zA-Z0-9_-]{10,})", link)
    if m:
        return m.group(1)

    # ?id={ID}
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]{10,})", link)
    if m:
        return m.group(1)

    return ""


def obtener_parent_desde_drive(drive_client, file_id: str) -> str:
    """
    Fallback: si el link de carpeta no parsea, obtiene el parent directo desde Drive.
    Retorna el primer parent del archivo, o '' si falla.
    """
    try:
        meta = (
            drive_client.files()
            .get(fileId=file_id, fields="parents", supportsAllDrives=True)
            .execute()
        )
        parents = meta.get("parents", [])
        return parents[0] if parents else ""
    except Exception:
        return ""


# ═══════════════════════════════════════════════════════
#  LISTAR CARPETAS
# ═══════════════════════════════════════════════════════

def listar_carpetas_proveedor(drive_client, root_folder_id: str, proveedores: list) -> list:
    """
    Lista carpetas válidas de proveedor dentro de root_folder_id.
    Formato esperado: "ID - NOMBRE" (ej: "40 - CEVEN S.A.").
    Solo retorna las que coinciden con la lista de proveedores de la planilla.
    """
    nombres_validos = {str(p.get("nombre", "")).strip().upper() for p in proveedores if p.get("nombre")}

    query = (
        f"'{root_folder_id}' in parents "
        f"and mimeType='application/vnd.google-apps.folder' "
        f"and trashed=false"
    )
    results = _paginate_files(
        drive_client, query, "files(id, name, webViewLink)"
    )

    carpetas = []
    for f in results:
        nombre = f["name"].strip()
        m = re.match(r"^(\d+)\s*-\s*(.+)$", nombre)
        if m:
            id_prov = m.group(1).strip()
            nombre_prov = m.group(2).strip()
            if nombre_prov.upper() in nombres_validos:
                carpetas.append(
                    {
                        "folder_id": f["id"],
                        "folder_link": f.get("webViewLink", ""),
                        "nombre": nombre_prov,
                        "id_proveedor": id_prov,
                    }
                )
    return carpetas


def listar_carpetas_invalidas(drive_client, root_folder_id: str, proveedores: list) -> list:
    """Carpetas en el root que NO corresponden a ningún proveedor válido."""
    nombres_validos = {str(p.get("nombre", "")).strip().upper() for p in proveedores if p.get("nombre")}

    query = (
        f"'{root_folder_id}' in parents "
        f"and mimeType='application/vnd.google-apps.folder' "
        f"and trashed=false"
    )
    results = _paginate_files(drive_client, query, "files(id, name)")

    invalidas = []
    for f in results:
        nombre = f["name"].strip()
        m = re.match(r"^(\d+)\s*-\s*(.+)$", nombre)
        if not m or m.group(2).strip().upper() not in nombres_validos:
            invalidas.append(
                {"carpeta": nombre, "motivo": "No coincide con ningún proveedor registrado"}
            )
    return invalidas


# ═══════════════════════════════════════════════════════
#  LISTAR EXCELS — UNA SOLA LLAMADA PARA TODAS LAS CARPETAS
# ═══════════════════════════════════════════════════════

def listar_excels_todas_carpetas(drive_client, carpetas: list) -> dict:
    """
    OPTIMIZACIÓN PRINCIPAL: reemplaza N llamadas API por UNA sola.

    Recibe lista de carpetas (con campo 'folder_id') y devuelve
    un dict { folder_id: [lista de archivos excel] }.

    Drive API: '{id}' in parents NO es recursiva.
    => Solo devuelve archivos directos en cada carpeta de proveedor,
       NO los que están en /procesado o /finalizado.
    """
    if not carpetas:
        return {}

    folder_ids = [c["folder_id"] for c in carpetas]

    # Construimos un OR query (Drive soporta esto eficientemente)
    parents_clause = " or ".join(f"'{fid}' in parents" for fid in folder_ids)
    query = f"({parents_clause}) and mimeType='{XLSX_MIME}' and trashed=false"

    results = _paginate_files(
        drive_client, query, "files(id, name, webViewLink, modifiedTime, parents)"
    )

    # Indexar por folder_id
    by_folder: dict = {fid: [] for fid in folder_ids}
    for f in results:
        for parent in f.get("parents", []):
            if parent in by_folder:
                by_folder[parent].append(f)
                break

    return by_folder


def listar_excels_carpeta(drive_client, folder_id: str) -> list:
    """Versión individual (usada cuando se necesita una carpeta específica)."""
    query = (
        f"'{folder_id}' in parents "
        f"and mimeType='{XLSX_MIME}' "
        f"and trashed=false"
    )
    return _paginate_files(drive_client, query, "files(id, name, webViewLink, modifiedTime)")


def listar_excels_modificados_desde(
    drive_client, folder_ids: list[str], desde_iso: str
) -> dict:
    """
    Para refresh incremental: retorna solo archivos modificados después de 'desde_iso'
    (formato RFC 3339, ej: '2025-06-01T12:00:00').
    Retorna dict { folder_id: [archivos] }.
    """
    if not folder_ids:
        return {}
    parents_clause = " or ".join(f"'{fid}' in parents" for fid in folder_ids)
    query = (
        f"({parents_clause}) "
        f"and mimeType='{XLSX_MIME}' "
        f"and modifiedTime > '{desde_iso}' "
        f"and trashed=false"
    )
    results = _paginate_files(
        drive_client, query, "files(id, name, webViewLink, modifiedTime, parents)"
    )
    by_folder: dict = {fid: [] for fid in folder_ids}
    for f in results:
        for parent in f.get("parents", []):
            if parent in by_folder:
                by_folder[parent].append(f)
                break
    return by_folder


# ═══════════════════════════════════════════════════════
#  VALIDACIÓN DE NOMBRE
# ═══════════════════════════════════════════════════════

def validar_nombre_excel(nombre: str) -> dict:
    """
    Valida formato: "PROVEEDOR - DD-MM-YY.xlsx" o "PROVEEDOR - DD-MM-YYYY.xlsx"
    Rechaza automáticamente archivos que ya tienen sufijos de procesado.
    """
    base = nombre.rsplit(".", 1)[0].strip()
    base_lower = base.lower()

    # Archivos ya procesados — ignorar silenciosamente en Tabla 1
    for sufijo in ("enviado", "finalizado", " - nota de credito", " - cambio", " - rechazado"):
        if sufijo in base_lower:
            return {
                "valido": False,
                "motivo_error": f"Archivo ya procesado (contiene '{sufijo}')",
            }

    m = re.match(r"^(.+?)\s*-\s*(\d{2}-\d{2}-(?:\d{2}|\d{4}))\s*$", base, re.IGNORECASE)
    if not m:
        return {
            "valido": False,
            "motivo_error": "Nombre no sigue el formato 'PROVEEDOR - DD-MM-YY.xlsx'",
        }

    proveedor = m.group(1).strip()
    fecha_str = m.group(2).strip()

    for fmt in ("%d-%m-%y", "%d-%m-%Y"):
        try:
            fecha = datetime.strptime(fecha_str, fmt)
            return {
                "valido": True,
                "proveedor": proveedor,
                "fecha": fecha.strftime("%Y-%m-%d"),
                "motivo_error": "",
            }
        except ValueError:
            continue

    return {
        "valido": False,
        "motivo_error": f"Fecha '{fecha_str}' inválida (usar DD-MM-YY o DD-MM-YYYY)",
    }


# ═══════════════════════════════════════════════════════
#  CONTAR ARTÍCULOS — CON CACHE EN MEMORIA Y DISCO
# ═══════════════════════════════════════════════════════

def contar_articulos_excel(drive_client, file_id: str) -> int | None:
    """Descarga el Excel y cuenta filas (sin header). Sin cache."""
    try:
        buf = _descargar_archivo(drive_client, file_id)
        df = pd.read_excel(buf, nrows=5000)
        return max(0, len(df))
    except Exception:
        return None


def contar_articulos_excel_cached(
    drive_client, file_id: str, modified_time: str
) -> int | None:
    """
    Versión con cache de dos niveles:
    1. Memoria (rápido, se pierde al reiniciar el proceso)
    2. Disco /tmp (persiste entre reinicios del server)

    La clave es (file_id + modifiedTime): si el archivo no cambió,
    no se vuelve a descargar.
    """
    cache_key = f"{file_id}__{modified_time}"

    # Nivel 1: memoria
    if cache_key in _COUNT_CACHE:
        return _COUNT_CACHE[cache_key]

    # Nivel 2: disco
    if cache_key in _DISK_CACHE:
        count = _DISK_CACHE[cache_key]
        _COUNT_CACHE[cache_key] = count
        return count

    # Sin cache: descargar y contar
    count = contar_articulos_excel(drive_client, file_id)
    _COUNT_CACHE[cache_key] = count
    _DISK_CACHE[cache_key] = count
    _save_disk_cache(_DISK_CACHE)
    return count


# ═══════════════════════════════════════════════════════
#  LEER EXCEL COMPLETO
# ═══════════════════════════════════════════════════════

def leer_excel_completo(drive_client, file_id: str) -> pd.DataFrame | None:
    """Descarga el Excel desde Drive y retorna como DataFrame."""
    try:
        buf = _descargar_archivo(drive_client, file_id)
        return pd.read_excel(buf)
    except Exception:
        return None


# ═══════════════════════════════════════════════════════
#  OPERACIONES DE ARCHIVO
# ═══════════════════════════════════════════════════════

def renombrar_archivo(drive_client, file_id: str, nuevo_nombre: str) -> None:
    drive_client.files().update(
        fileId=file_id,
        body={"name": nuevo_nombre},
        supportsAllDrives=True,
    ).execute()


def mover_archivo(drive_client, file_id: str, nuevo_parent_id: str) -> None:
    file_meta = (
        drive_client.files()
        .get(fileId=file_id, fields="parents", supportsAllDrives=True)
        .execute()
    )
    old_parents = ",".join(file_meta.get("parents", []))
    drive_client.files().update(
        fileId=file_id,
        addParents=nuevo_parent_id,
        removeParents=old_parents,
        supportsAllDrives=True,
        fields="id, parents",
    ).execute()


def asegurar_subcarpeta(drive_client, parent_id: str, nombre: str) -> str:
    """Crea la subcarpeta si no existe. Retorna su ID."""
    query = (
        f"'{parent_id}' in parents "
        f"and name='{nombre}' "
        f"and mimeType='application/vnd.google-apps.folder' "
        f"and trashed=false"
    )
    results = drive_client.files().list(
        q=query,
        fields="files(id)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = results.get("files", [])
    if files:
        return files[0]["id"]

    folder = drive_client.files().create(
        body={
            "name": nombre,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        },
        fields="id",
        supportsAllDrives=True,
    ).execute()
    return folder["id"]


def obtener_nombre_archivo(drive_client, file_id: str) -> str:
    """Obtiene el nombre actual de un archivo desde Drive."""
    try:
        meta = (
            drive_client.files()
            .get(fileId=file_id, fields="name", supportsAllDrives=True)
            .execute()
        )
        return meta.get("name", "")
    except Exception:
        return ""


# ═══════════════════════════════════════════════════════
#  HELPERS INTERNOS
# ═══════════════════════════════════════════════════════

def _paginate_files(drive_client, query: str, fields: str) -> list:
    """Maneja paginación de Drive API automáticamente."""
    items = []
    page_token = None
    while True:
        params = {
            "q": query,
            "fields": f"nextPageToken, {fields}",
            "supportsAllDrives": True,
            "includeItemsFromAllDrives": True,
            "pageSize": 1000,
        }
        if page_token:
            params["pageToken"] = page_token
        response = drive_client.files().list(**params).execute()
        items.extend(response.get("files", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            break
    return items


def _descargar_archivo(drive_client, file_id: str) -> io.BytesIO:
    """Descarga un archivo de Drive como BytesIO."""
    request = drive_client.files().get_media(
        fileId=file_id, supportsAllDrives=True
    )
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf
