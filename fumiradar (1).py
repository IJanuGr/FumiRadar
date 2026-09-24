"""
FumiRadar v0.2 — detecta locales gastronómicos recién abiertos usando SerpApi.

Cómo correrlo:
  1. pip install -r requirements.txt
  2. Poné tu key en un archivo .env (copiá .env.example) o en la variable
     de entorno SERPAPI_KEY. Nunca la escribas en este archivo.
  3. Prueba chica (1 rubro, 1 localidad, 5 créditos, guarda respuestas crudas):
       python fumiradar.py --prueba
  4. Corrida completa:
       python fumiradar.py

Resultado: fumiradar_san_isidro.xlsx con los candidatos ordenados
del más nuevo al más viejo.
"""
import argparse
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

import requests
from openpyxl import Workbook

# ------------------------------------------------------------------
# CONFIGURACIÓN (lo único que vas a tocar seguido)
# ------------------------------------------------------------------
RUBROS = ["cafetería", "pizzería", "restaurante"]          # empezá con pocos para cuidar créditos
LOCALIDADES = ["San Isidro", "Martínez", "Acassuso", "Beccar", "Boulogne", "Villa Adelina"]
MAX_RESENAS = 30        # más reseñas que esto = local viejo, ni lo miramos
MAX_DIAS = 120          # la reseña más vieja tiene que ser de hace menos de esto
MAX_CREDITOS = 70       # freno de mano: el script corta si llega a este número de llamadas
CADENAS = ["café martínez", "mi gusto", "le blé", "havanna", "grido", "mostaza",
           "mcdonald", "burger king", "starbucks", "subway", "blossom", "kentucky",
           "la farola", "big pizza", "almacén de pizzas", "bonafide"]

CARPETA_CRUDO = Path("crudo")   # acá se guardan las respuestas crudas en modo --prueba

creditos_usados = 0
guardar_crudo = False


def cargar_env(ruta=".env"):
    """Lee KEY=valor de un archivo .env (si existe) sin pisar variables ya definidas."""
    if not os.path.exists(ruta):
        return
    with open(ruta, encoding="utf-8") as f:
        for linea in f:
            linea = linea.strip()
            if linea and not linea.startswith("#") and "=" in linea:
                clave, valor = linea.split("=", 1)
                os.environ.setdefault(clave.strip(), valor.strip().strip('"').strip("'"))


def llamar_serpapi(params):
    """Hace UNA llamada a SerpApi y suma un crédito. Si pasamos el límite, corta todo."""
    global creditos_usados
    if creditos_usados >= MAX_CREDITOS:
        raise RuntimeError("Llegué al límite de créditos (MAX_CREDITOS). Corto acá.")
    params = {**params, "api_key": os.environ["SERPAPI_KEY"], "hl": "es", "gl": "ar"}
    resp = requests.get("https://serpapi.com/search.json", params=params, timeout=30)
    creditos_usados += 1
    data = resp.json() if resp.headers.get("content-type", "").startswith("application/json") else {}

    if guardar_crudo:
        CARPETA_CRUDO.mkdir(exist_ok=True)
        archivo = CARPETA_CRUDO / f"{creditos_usados:02d}_{params['engine']}.json"
        archivo.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  [crudo] respuesta guardada en {archivo}")

    # SerpApi avisa los problemas con un campo "error" (key inválida, sin créditos, etc.)
    if resp.status_code != 200:
        raise SystemExit(f"SerpApi respondió {resp.status_code}: {data.get('error', resp.text[:200])}")
    if "error" in data:
        # "Google hasn't returned any results" no es grave: es una búsqueda vacía
        print(f"  Aviso de SerpApi: {data['error']}")
    return data


def buscar_locales(rubro, localidad):
    """Paso 1: busca en Google Maps 'rubro + localidad' y devuelve la lista de locales."""
    data = llamar_serpapi({
        "engine": "google_maps",
        "type": "search",
        "q": f"{rubro} {localidad}, Buenos Aires, Argentina",
    })
    return data.get("local_results", [])


def fecha_resena_mas_vieja(data_id, hoy):
    """
    Paso 3: trae las reseñas ordenadas de más nueva a más vieja y recorre
    las páginas. La última reseña que aparece es la más vieja,
    o sea, más o menos cuándo abrió el local.

    Ahorro de créditos: si en una página ya aparece una reseña más vieja
    que MAX_DIAS, el local se va a descartar igual, así que no pedimos
    más páginas.
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
        if mas_vieja is not None and (hoy - mas_vieja).days > MAX_DIAS:
            return mas_vieja    # ya sabemos que es viejo: no gastamos más créditos
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
        return "Fecha de apertura sin confirmar"
    if dias < 30:
        return "Abrió hace pocas semanas"
    meses = round(dias / 30)
    return f"Abrió hace aproximadamente {meses} {'mes' if meses == 1 else 'meses'}"


def main():
    global MAX_CREDITOS, RUBROS, LOCALIDADES, guardar_crudo

    parser = argparse.ArgumentParser(description="FumiRadar: locales gastronómicos nuevos")
    parser.add_argument("--prueba", action="store_true",
                        help="1 rubro, 1 localidad, 5 créditos y guarda respuestas crudas")
    parser.add_argument("--max-creditos", type=int, help="pisa MAX_CREDITOS")
    args = parser.parse_args()

    if args.prueba:
        RUBROS, LOCALIDADES, MAX_CREDITOS, guardar_crudo = RUBROS[:1], LOCALIDADES[:1], 5, True
    if args.max_creditos is not None:
        MAX_CREDITOS = args.max_creditos

    cargar_env()
    if not os.environ.get("SERPAPI_KEY"):
        sys.exit("Falta SERPAPI_KEY (en el archivo .env o como variable de entorno).")

    print(f"Rubros: {RUBROS} | Localidades: {LOCALIDADES} | Límite: {MAX_CREDITOS} créditos")
    hoy = datetime.now(timezone.utc)
    vistos = set()          # para no procesar dos veces el mismo local
    candidatos = []

    # Paso 1 y 2: buscar y filtrar por cantidad de reseñas (esto es barato)
    try:
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
    except RuntimeError as e:
        print(e)

    print(f"Candidatos con ≤{MAX_RESENAS} reseñas: {len(candidatos)} (créditos usados: {creditos_usados})")

    # Paso 3: solo a los candidatos les pedimos las reseñas (esto es lo caro)
    resultados = []
    for c in sorted(candidatos, key=lambda x: x["_n"]):   # primero los de menos reseñas
        try:
            fecha = fecha_resena_mas_vieja(c["data_id"], hoy) if c["_n"] > 0 else None
        except RuntimeError as e:
            print(e)
            break
        dias = (hoy - fecha).days if fecha else None
        if dias is not None and dias > MAX_DIAS:
            continue
        resultados.append((c, fecha, dias))

    # Paso 4: exportar a Excel, del más nuevo al más viejo.
    # Los que no tienen reseñas van AL FINAL: no tenemos cómo confirmar su fecha.
    resultados.sort(key=lambda x: (x[2] is None, x[2] or 0))
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
    salida = "fumiradar_prueba.xlsx" if args.prueba else "fumiradar_san_isidro.xlsx"
    wb.save(salida)
    con_fecha = sum(1 for r in resultados if r[2] is not None)
    print(f"Listo: {len(resultados)} locales en {salida} ({con_fecha} con fecha confirmada). "
          f"Créditos usados: {creditos_usados}")


if __name__ == "__main__":
    main()
