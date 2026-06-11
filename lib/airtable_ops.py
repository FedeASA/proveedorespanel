"""
lib/airtable_ops.py
Operaciones con Airtable usando la REST API v0 directamente (sin SDK externo).

Base:  appjlLix1HpBwnhpS
Table: tblNnoXdIsLFN92Mr
Campos relevantes:
  - proveedor        (texto)
  - Producto         (texto)
  - Serial           (texto)
  - diagnostico      (texto)
  - RMA_Capturado    (checkbox)
  - RMA_Fecha_Captura (fecha, formato ISO)
"""

from __future__ import annotations

import time
from datetime import date
from typing import Optional

import requests

BASE_ID  = "appjlLix1HpBwnhpS"
TABLE_ID = "tblNnoXdIsLFN92Mr"
API_URL  = f"https://api.airtable.com/v0/{BASE_ID}/{TABLE_ID}"

CAMPOS_EXPORT = ["Producto", "Serial", "diagnostico", "proveedor"]


def _headers(api_key: str) -> dict:
    return {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}


# ── LECTURA ────────────────────────────────────────────────────────────────────

def listar_proveedores_unicos(api_key: str) -> list[str]:
    """
    Devuelve la lista ordenada y deduplicada de valores del campo 'proveedor'.
    Pagina automáticamente hasta obtener todos los registros.
    """
    proveedores: set[str] = set()
    params: dict = {
        "fields[]": ["proveedor"],
        "filterByFormula": "NOT({proveedor} = '')",
        "pageSize": 100,
    }
    offset = None
    while True:
        if offset:
            params["offset"] = offset
        resp = requests.get(API_URL, headers=_headers(api_key), params=params, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        for rec in data.get("records", []):
            val = str(rec.get("fields", {}).get("proveedor", "")).strip()
            if val:
                proveedores.add(val)
        offset = data.get("offset")
        if not offset:
            break
    return sorted(proveedores)


def obtener_registros_proveedor(api_key: str, proveedor: str) -> list[dict]:
    """
    Retorna todos los registros donde 'proveedor' == proveedor
    y RMA_Capturado es falso (o está vacío).

    Cada elemento del resultado es:
      {"id": <airtable record id>, "Producto": ..., "Serial": ..., "diagnostico": ...}
    """
    formula = (
        f"AND("
        f"LOWER(TRIM({{proveedor}}))=LOWER(TRIM('{proveedor}')), "
        f"OR({{RMA_Capturado}}=0, {{RMA_Capturado}}=FALSE(), {{RMA_Capturado}}='')"
        f")"
    )
    params: dict = {
        "fields[]": ["proveedor", "Producto", "Serial", "diagnostico"],
        "filterByFormula": formula,
        "pageSize": 100,
    }
    registros: list[dict] = []
    offset = None
    while True:
        if offset:
            params["offset"] = offset
        resp = requests.get(API_URL, headers=_headers(api_key), params=params, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        for rec in data.get("records", []):
            f = rec.get("fields", {})
            registros.append({
                "id":          rec["id"],
                "Producto":    str(f.get("Producto",    "") or "").strip(),
                "Serial":      str(f.get("Serial",      "") or "").strip(),
                "diagnostico": str(f.get("diagnostico", "") or "").strip(),
                "proveedor":   str(f.get("proveedor",   "") or "").strip(),
            })
        offset = data.get("offset")
        if not offset:
            break
    return registros


# ── ESCRITURA ──────────────────────────────────────────────────────────────────

def marcar_como_capturados(api_key: str, record_ids: list[str]) -> tuple[int, list[str]]:
    """
    Actualiza los registros indicados en lotes de 10 (límite de la API)
    seteando:
      - RMA_Capturado    = True
      - RMA_Fecha_Captura = hoy (ISO YYYY-MM-DD)

    Retorna (cantidad_actualizados, lista_de_errores).
    """
    hoy = date.today().isoformat()
    errores: list[str] = []
    actualizados = 0

    # Airtable acepta hasta 10 registros por PATCH
    for i in range(0, len(record_ids), 10):
        batch = record_ids[i: i + 10]
        payload = {
            "records": [
                {
                    "id": rid,
                    "fields": {
                        "RMA_Capturado":     True,
                        "RMA_Fecha_Captura": hoy,
                    },
                }
                for rid in batch
            ]
        }
        try:
            resp = requests.patch(
                API_URL,
                headers=_headers(api_key),
                json=payload,
                timeout=20,
            )
            resp.raise_for_status()
            actualizados += len(resp.json().get("records", []))
        except Exception as exc:
            errores.append(str(exc))
        # Respetar rate-limit de Airtable (5 req/s)
        time.sleep(0.22)

    return actualizados, errores
