"""
FumiRadar v0.3 — detecta locales gastronómicos recién abiertos usando SerpApi.

Cómo correrlo:
  1. pip install -r requirements.txt
  2. Poné tu key en un archivo .env (copiá .env.example) o en la variable
     de entorno SERPAPI_KEY. Nunca la escribas en este archivo.
  3. Prueba chica (1 rubro, 1 zona, 5 créditos, guarda respuestas crudas):
       python fumiradar.py --prueba
  4. Corrida completa:
       python fumiradar.py

Resultado: fumiradar_san_isidro.xlsx con los candidatos ordenados
del más nuevo al más viejo.

Cómo decide si un local es nuevo:
  - Filtro barato (gratis, viene en la búsqueda): pocas reseñas y no es cadena
    ni sucursal de una marca conocida.
  - Medición (cuesta créditos): la fecha de la reseña MÁS VIEJA. Un local recién
    abierto recibe su primera reseña en las primeras 1-2 semanas, así que esa
    fecha es una buena aproximación de cuándo abrió.
"""
import argparse
import json
import os
import re
import sys
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

import requests
from openpyxl import Workbook

# ------------------------------------------------------------------
# CONFIGURACIÓN (lo único que vas a tocar seguido)
# ------------------------------------------------------------------
RUBROS = ["restaurante", "parrilla", "pizzería", "hamburguesería", "sushi", "cafetería",
          "panadería", "heladería", "cervecería", "rotisería", "pastas", "empanadas"]
# Zonas agrupadas para ahorrar búsquedas (Google ya devuelve lo que hay alrededor)
ZONAS = ["San Isidro", "Martínez", "Boulogne"]
# Búsquedas extra: Google Maps también busca en el texto de las reseñas,
# y los locales nuevos tienen reseñas que dicen "nuevo", "recién abrió", etc.
BUSQUEDAS_NUEVOS = ["restaurante nuevo", "recién inaugurado comida"]

MAX_RESENAS = 80        # más reseñas que esto = local viejo, ni lo miramos
MAX_DIAS = 120          # la reseña más vieja tiene que ser de hace menos de esto
MAX_CREDITOS = 70       # freno de mano: el script corta si llega a este número de llamadas
MIN_RESENAS_CONFIANZA = 5   # con menos reseñas, puede ser un local viejo recién cargado en Maps

CADENAS = ["café martínez", "mi gusto", "le blé", "havanna", "grido", "mostaza",
           "mcdonald", "burger king", "starbucks", "subway", "blossom", "kentucky",
           "la farola", "big pizza", "almacén de pizzas", "bonafide", "freddo", "rapanui",
           "lucciano", "persicco", "chungo", "cremolatti", "club de la milanesa", "tostado",
           "la panera rosa", "dean & dennys", "wendy", "popeyes", "sushi club", "sushi pop",
           "antares", "patagonia", "temple", "kansas", "tea connection", "green eat",
           "betos", "pertutti", "la continental", "moshi moshi", "juan valdez", "tienda de café"]
# Palabras que se sacan del nombre para detectar sucursales ("Marca San Isidro" = "Marca")
SUFIJOS_SUCURSAL = ["san isidro", "martinez", "acassuso", "beccar", "boulogne", "villa adelina",
                    "zona norte", "sucursal", "s.i"]

# Solo nos quedamos con locales cuya dirección esté en estas localidades
# (Google a veces trae locales de otros partidos, ej. "Parrilla Martinez" en La Matanza)
LOCALIDADES_VALIDAS = ["san isidro", "martinez", "acassuso", "beccar", "boulogne", "villa adelina"]
# Nombres que quedan genéricos al sacarles la zona ("Panadería Martínez" -> "panaderia"):
# no sirven para detectar marcas
GENERICOS = {"restaurante", "restaurant", "resto", "parrilla", "pizzeria", "hamburgueseria",
             "sushi", "cafeteria", "cafe", "panaderia", "heladeria", "cerveceria", "rotiseria",
             "pastas", "empanadas", "bar", "confiteria", "bodegon"}

CARPETA_CRUDO = Path("crudo")   # acá se guardan las respuestas crudas en modo --prueba
ARCHIVO_CACHE = Path("cache.json")   # búsquedas y fechas ya consultadas: no se pagan dos veces
DIAS_CACHE_BUSQUEDA = 7

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


def cargar_cache():
    if ARCHIVO_CACHE.exists():
        return json.loads(ARCHIVO_CACHE.read_text(encoding="utf-8"))
    return {"busquedas": {}, "resenas": {}}


def guardar_cache(cache):
    ARCHIVO_CACHE.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")


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


def buscar_locales(consulta, cache, hoy):
    """Paso 1: busca en Google Maps y devuelve la lista de locales (usa el caché si es reciente)."""
    guardada = cache["busquedas"].get(consulta)
    if guardada and (hoy - datetime.fromisoformat(guardada["fecha"])).days < DIAS_CACHE_BUSQUEDA:
        return guardada["local_results"]
    data = llamar_serpapi({
        "engine": "google_maps",
        "type": "search",
        "q": f"{consulta}, Buenos Aires, Argentina",
    })
    resultados = data.get("local_results", [])
    cache["busquedas"][consulta] = {"fecha": hoy.isoformat(), "local_results": resultados}
    return resultados


def fecha_resena_mas_vieja(data_id, hoy):
    """
    Paso 3: trae las reseñas ordenadas de más nueva a más vieja y recorre
    las páginas. La reseña más vieja es, más o menos, cuándo abrió el local.

    Ojo: Google ordena por fecha de ÚLTIMA EDICIÓN, así que una reseña vieja
    editada puede aparecer primero. Por eso tomamos el mínimo de iso_date
    (fecha de creación) de todo lo que vimos, no la última de la lista.

    Ahorro de créditos: si ya apareció una reseña más vieja que MAX_DIAS,
    el local se descarta igual, así que no pedimos más páginas. Un local viejo
    con pocas reseñas recibe pocas por mes, así que casi siempre se descarta
    con la primera página (1 crédito).
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


def sin_acentos(texto):
    return unicodedata.normalize("NFKD", texto.lower()).encode("ascii", "ignore").decode()


def es_cadena(nombre):
    """Paso 2b: descarta cadenas y franquicias conocidas (ya tienen fumigador corporativo)."""
    n = sin_acentos(nombre)
    return any(sin_acentos(c) in n for c in CADENAS)


def nombre_base(titulo):
    """
    'Brooklyn Bakery - San Isidro' y 'Brooklyn Bakery (Beccar)' -> 'brooklyn bakery'.
    Sirve para detectar sucursales nuevas de marcas que ya existen.
    """
    t = sin_acentos(titulo)
    t = re.split(r"\s+[-–|(]", t)[0]            # corta "Marca - Sucursal", "Marca (Zona)"
    for sufijo in SUFIJOS_SUCURSAL:
        t = re.sub(rf"\b{re.escape(sufijo)}\b", " ", t)
    return " ".join(re.sub(r"[^a-z0-9 ]", " ", t).split())


def es_sucursal_de_marca(local, marcas_establecidas):
    base = nombre_base(local.get("title", ""))
    # nombres muy cortos o genéricos ("cafe", "panaderia") darían falsos positivos
    return len(base) >= 5 and base not in GENERICOS and base in marcas_establecidas


def en_la_zona(local):
    direccion = sin_acentos(local.get("address") or "")
    return any(loc in direccion for loc in LOCALIDADES_VALIDAS)


def confianza(n_resenas, dias):
    if dias is None:
        return "baja (sin reseñas)"
    if n_resenas < MIN_RESENAS_CONFIANZA:
        return "baja (pocas reseñas: puede ser un local viejo recién cargado en Maps)"
    return "alta"


def texto_para_cliente(dias):
    """Lo que ve el fumigador: solo el dato, nunca cómo lo sacamos."""
    if dias is None:
        return "Fecha de apertura sin confirmar"
    if dias < 30:
        return "Abrió hace pocas semanas"
    meses = round(dias / 30)
    return f"Abrió hace aproximadamente {meses} {'mes' if meses == 1 else 'meses'}"


def main():
    global MAX_CREDITOS, RUBROS, ZONAS, BUSQUEDAS_NUEVOS, guardar_crudo

    parser = argparse.ArgumentParser(description="FumiRadar: locales gastronómicos nuevos")
    parser.add_argument("--prueba", action="store_true",
                        help="1 rubro, 1 zona, 5 créditos y guarda respuestas crudas")
    parser.add_argument("--max-creditos", type=int, help="pisa MAX_CREDITOS")
    args = parser.parse_args()

    if args.prueba:
        RUBROS, ZONAS, BUSQUEDAS_NUEVOS, MAX_CREDITOS, guardar_crudo = RUBROS[:1], ZONAS[:1], [], 5, True
    if args.max_creditos is not None:
        MAX_CREDITOS = args.max_creditos

    cargar_env()
    if not os.environ.get("SERPAPI_KEY"):
        sys.exit("Falta SERPAPI_KEY (en el archivo .env o como variable de entorno).")

    consultas = [f"{b} {z}" for b in BUSQUEDAS_NUEVOS for z in ZONAS]   # primero las de "nuevo"
    consultas += [f"{r} {z}" for r in RUBROS for z in ZONAS]
    print(f"{len(consultas)} búsquedas ({len(RUBROS)} rubros × {len(ZONAS)} zonas "
          f"+ {len(BUSQUEDAS_NUEVOS) * len(ZONAS)} de 'nuevos') | Límite: {MAX_CREDITOS} créditos")

    hoy = datetime.now(timezone.utc)
    cache = cargar_cache()

    # Paso 1: juntar TODOS los locales que aparecen (también los viejos: sirven
    # para detectar marcas establecidas)
    todos = {}              # data_id -> local
    de_busqueda_nuevos = set()
    try:
        for consulta in consultas:
            for local in buscar_locales(consulta, cache, hoy):
                did = local.get("data_id")
                if did and did not in todos:
                    todos[did] = local
                    if any(consulta.startswith(b) for b in BUSQUEDAS_NUEVOS):
                        de_busqueda_nuevos.add(did)
    except RuntimeError as e:
        print(e)
    guardar_cache(cache)
    print(f"Locales únicos encontrados: {len(todos)} (créditos usados: {creditos_usados})")

    # Paso 2: filtros gratis
    marcas_establecidas = {nombre_base(l.get("title", "")) for l in todos.values()
                           if (l.get("reviews") or 0) > MAX_RESENAS}
    candidatos, descartes = [], {"fuera de zona": 0, "muchas reseñas": 0, "cadena": 0,
                                 "sucursal de marca": 0}
    for did, local in todos.items():
        n = local.get("reviews", 0) or 0
        if not en_la_zona(local):
            descartes["fuera de zona"] += 1
        elif n > MAX_RESENAS:
            descartes["muchas reseñas"] += 1
        elif es_cadena(local.get("title", "")):
            descartes["cadena"] += 1
        elif es_sucursal_de_marca(local, marcas_establecidas):
            descartes["sucursal de marca"] += 1
            print(f"  Descarto sucursal de marca conocida: {local.get('title')}")
        else:
            candidatos.append({**local, "_n": n, "_nuevo": did in de_busqueda_nuevos})
    print(f"Descartes gratis: {descartes}")
    print(f"Candidatos a revisar: {len(candidatos)}")

    # Paso 3: fecha de la reseña más vieja (esto es lo caro).
    # Prioridad: los que salieron en búsquedas de "nuevo", después los de menos reseñas.
    resultados, sin_revisar = [], 0
    for c in sorted(candidatos, key=lambda x: (not x["_nuevo"], x["_n"])):
        did = c["data_id"]
        if did in cache["resenas"]:
            iso = cache["resenas"][did]
            fecha = datetime.fromisoformat(iso) if iso else None
        elif c["_n"] == 0:
            fecha = None
        else:
            try:
                fecha = fecha_resena_mas_vieja(did, hoy)
            except RuntimeError:
                sin_revisar += 1
                continue
            cache["resenas"][did] = fecha.isoformat() if fecha else None
        dias = (hoy - fecha).days if fecha else None
        if dias is not None and dias > MAX_DIAS:
            continue
        resultados.append((c, fecha, dias))
    guardar_cache(cache)
    if sin_revisar:
        print(f"Se acabaron los créditos: quedaron {sin_revisar} candidatos sin revisar "
              f"(la próxima corrida sigue desde ahí gracias al caché).")

    # Paso 4: exportar a Excel, del más nuevo al más viejo.
    # Los que no tienen reseñas van AL FINAL: no tenemos cómo confirmar su fecha.
    resultados.sort(key=lambda x: (x[2] is None, x[2] or 0))
    wb = Workbook()
    ws = wb.active
    ws.title = "Prospectos"
    ws.append(["Local", "Tipo", "Dirección", "Teléfono", "Web", "Reseñas",
               "Reseña más vieja", "Días", "Confianza", "Para el cliente", "Google Maps"])
    for c, fecha, dias in resultados:
        ws.append([c.get("title"), c.get("type"), c.get("address"), c.get("phone"),
                   c.get("website"), c["_n"],
                   fecha.strftime("%d/%m/%Y") if fecha else "sin reseñas",
                   dias, confianza(c["_n"], dias), texto_para_cliente(dias),
                   f"https://www.google.com/maps/place/?q=place_id:{c.get('place_id')}"])
    salida = "fumiradar_prueba.xlsx" if args.prueba else "fumiradar_san_isidro.xlsx"
    wb.save(salida)
    con_fecha = sum(1 for r in resultados if r[2] is not None)
    print(f"Listo: {len(resultados)} locales en {salida} ({con_fecha} con fecha confirmada). "
          f"Créditos usados: {creditos_usados}")


if __name__ == "__main__":
    main()
