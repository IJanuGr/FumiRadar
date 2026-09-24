"""
FumiRadar v0.1 — detecta locales gastronómicos recién abiertos usando SerpApi.

Cómo correrlo:
  1. pip install requests openpyxl
  2. export SERPAPI_KEY="tu_key"      (en Windows: set SERPAPI_KEY=tu_key)
  3. python fumiradar.py

Resultado: fumiradar_san_isidro.xlsx con los candidatos ordenados
del más nuevo al más viejo.
"""
import os
import sys
from datetime import datetime, timezone

import requests
from openpyxl import Workbook

# ------------------------------------------------------------------
# CONFIGURACIÓN (lo único que vas a tocar seguido)
# ------------------------------------------------------------------
API_KEY = os.environ.get("SERPAPI_KEY")
RUBROS = ["cafetería", "pizzería", "restaurante"]          # empezá con pocos para cuidar créditos
LOCALIDADES = ["San Isidro", "Martínez", "Acassuso", "Beccar", "Boulogne", "Villa Adelina"]
MAX_RESENAS = 30        # más reseñas que esto = local viejo, ni lo miramos
MAX_DIAS = 120          # la reseña más vieja tiene que ser de hace menos de esto
MAX_CREDITOS = 70       # freno de mano: el script corta si llega a este número de llamadas
CADENAS = ["café martínez", "mi gusto", "le blé", "havanna", "grido", "mostaza",
           "mcdonald", "burger king", "starbucks", "subway", "blossom", "kentucky",
           "la farola", "big pizza", "almacén de pizzas", "bonafide"]

creditos_usados = 0


def llamar_serpapi(params):
    """Hace UNA llamada a SerpApi y suma un crédito. Si pasamos el límite, corta todo."""
    global creditos_usados
    if creditos_usados >= MAX_CREDITOS:
        raise RuntimeError("Llegué al límite de créditos (MAX_CREDITOS). Corto acá.")
    params = {**params, "api_key": API_KEY, "hl": "es", "gl": "ar"}
    resp = requests.get("https://serpapi.com/search.json", params=params, timeout=30)
    creditos_usados += 1
    resp.raise_for_status()
    return resp.json()


def buscar_locales(rubro, localidad):
    """Paso 1: busca en Google Maps 'rubro + localidad' y devuelve la lista de locales."""
    data = llamar_serpapi({
        "engine": "google_maps",
        "type": "search",
        "q": f"{rubro} {localidad}, Buenos Aires, Argentina",
    })
    return data.get("local_results", [])


def fecha_resena_mas_vieja(data_id):
    """
    Paso 3: trae las reseñas ordenadas de más nueva a más vieja y recorre
    todas las páginas. La última reseña que aparece es la más vieja,
    o sea, más o menos cuándo abrió el local.
    """
    params = {"engine": "google_maps_reviews", "data_id": data_id, "sort_by": "newestFirst"}
    mas_vieja = None
    while True:
        data = llamar_serpapi(params)
        for r in data.get("reviews", []):
            iso = r.get("iso_date")
            if iso:
                fecha = datetime.fromisoformat(iso.replace("Z", "+00:00"))
                if mas_vieja is None or fecha < mas_vieja:
                    mas_vieja = fecha
        token = data.get("serpapi_pagination", {}).get("next_page_token")
        if not token:
            return mas_vieja
        # A partir de la 2da página se pueden pedir 20 reseñas por llamada
        params = {**params, "next_page_token": token, "num": 20}


def es_cadena(nombre):
    """Paso 2b: descarta cadenas y franquicias (ya tienen fumigador corporativo)."""
    n = nombre.lower()
    return any(c in n for c in CADENAS)


def texto_para_cliente(dias):
    """Lo que ve el fumigador: solo el dato, nunca cómo lo sacamos."""
    if dias is None:
        return "Local muy nuevo, todavía sin reseñas"
    if dias < 30:
        return "Abrió hace pocas semanas"
    meses = round(dias / 30)
    return f"Abrió hace aproximadamente {meses} {'mes' if meses == 1 else 'meses'}"


def main():
    if not API_KEY:
        sys.exit("Falta la variable de entorno SERPAPI_KEY.")

    hoy = datetime.now(timezone.utc)
    vistos = set()          # para no procesar dos veces el mismo local
    candidatos = []

    # Paso 1 y 2: buscar y filtrar por cantidad de reseñas (esto es barato)
    for rubro in RUBROS:
        for loc in LOCALIDADES:
            for local in buscar_locales(rubro, loc):
                did = local.get("data_id")
                nombre = local.get("title", "")
                n_resenas = local.get("reviews", 0) or 0
                if not did or did in vistos or es_cadena(nombre) or n_resenas > MAX_RESENAS:
                    continue
                vistos.add(did)
                candidatos.append({**local, "_rubro": rubro, "_n": n_resenas})

    print(f"Candidatos con ≤{MAX_RESENAS} reseñas: {len(candidatos)} (créditos usados: {creditos_usados})")

    # Paso 3: solo a los candidatos les pedimos las reseñas (esto es lo caro)
    resultados = []
    for c in sorted(candidatos, key=lambda x: x["_n"]):   # primero los de menos reseñas
        try:
            fecha = fecha_resena_mas_vieja(c["data_id"]) if c["_n"] > 0 else None
        except RuntimeError as e:
            print(e)
            break
        dias = (hoy - fecha).days if fecha else None
        if dias is not None and dias > MAX_DIAS:
            continue
        resultados.append((c, fecha, dias))

    # Paso 4: exportar a Excel, del más nuevo al más viejo
    resultados.sort(key=lambda x: x[2] if x[2] is not None else -1)
    wb = Workbook()
    ws = wb.active
    ws.title = "Prospectos"
    ws.append(["Local", "Rubro", "Dirección", "Teléfono", "Web", "Reseñas",
               "Reseña más vieja", "Días", "Para el cliente"])
    for c, fecha, dias in resultados:
        ws.append([c.get("title"), c["_rubro"], c.get("address"), c.get("phone"),
                   c.get("website"), c["_n"],
                   fecha.strftime("%d/%m/%Y") if fecha else "sin reseñas",
                   dias, texto_para_cliente(dias)])
    salida = "fumiradar_san_isidro.xlsx"
    wb.save(salida)
    print(f"Listo: {len(resultados)} locales en {salida}. Créditos usados: {creditos_usados}")


if __name__ == "__main__":
    main()
