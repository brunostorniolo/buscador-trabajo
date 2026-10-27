"""
Buscador automático de ofertas laborales.

Junta ofertas de varios portales, las puntúa según tu perfil (config.json)
y guarda las que superan el puntaje mínimo en ofertas.db (SQLite).

Uso:
    python buscador.py              # busca y guarda ofertas nuevas
    python buscador.py --exportar   # además genera ofertas.csv para abrir en Excel
"""

import csv
import html
import json
import re
import sqlite3
import sys
import unicodedata
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from pathlib import Path

CARPETA = Path(__file__).resolve().parent
CONFIG = CARPETA / "config.json"
BASE = CARPETA / "ofertas.db"
CSV = CARPETA / "ofertas.csv"
PAGINA = CARPETA / "docs" / "index.html"
PLANTILLA = CARPETA / "plantilla.html"

ARGENTINA = timezone(timedelta(hours=-3))

USER_AGENT = "Mozilla/5.0 (buscador-personal-de-empleo)"


# ---------------------------------------------------------------- utilidades

def descargar(url, json_=True):
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept": "*/*"})
    with urllib.request.urlopen(req, timeout=30) as r:
        datos = r.read()
    return json.loads(datos) if json_ else datos


def limpiar_html(texto):
    texto = re.sub(r"<[^>]+>", " ", texto or "")
    return re.sub(r"\s+", " ", html.unescape(texto)).strip()


def normalizar(texto):
    """Minúsculas y sin tildes, para comparar palabras clave."""
    texto = unicodedata.normalize("NFKD", (texto or "").lower())
    return "".join(c for c in texto if not unicodedata.combining(c))


def contiene(texto_norm, palabra):
    return re.search(r"\b" + re.escape(normalizar(palabra)) + r"\b", texto_norm) is not None


def oferta(fuente, id_externo, titulo, empresa, url, descripcion="",
           ubicacion="", publicada="", salario="", remoto=True):
    return {
        "id": f"{fuente}:{id_externo}",
        "fuente": fuente,
        "titulo": (titulo or "").strip(),
        "empresa": (empresa or "").strip(),
        "url": url or "",
        "descripcion": limpiar_html(descripcion)[:5000],
        "ubicacion": (ubicacion or "").strip(),
        "publicada": publicada or "",
        "salario": salario or "",
        "remoto": 1 if remoto else 0,
    }


# ---------------------------------------------------------------- fuentes

def fuente_remoteok(cfg):
    datos = descargar("https://remoteok.com/api")
    for j in datos[1:]:  # el primer elemento es un aviso legal
        salario = ""
        if j.get("salary_min"):
            salario = f"USD {j['salary_min']}-{j.get('salary_max') or ''}/año"
        yield oferta("RemoteOK", j.get("id"), j.get("position"), j.get("company"),
                     j.get("url"), j.get("description"), j.get("location") or "Remoto",
                     j.get("date", "")[:10], salario)


def fuente_remotive(cfg):
    datos = descargar("https://remotive.com/api/remote-jobs")
    for j in datos.get("jobs", []):
        yield oferta("Remotive", j.get("id"), j.get("title"), j.get("company_name"),
                     j.get("url"), j.get("description"),
                     j.get("candidate_required_location") or "Remoto",
                     (j.get("publication_date") or "")[:10], j.get("salary"))


def fuente_weworkremotely(cfg):
    categorias = ["remote-design-jobs", "remote-sales-and-marketing-jobs", "all-other-remote-jobs"]
    for cat in categorias:
        raiz = ET.fromstring(descargar(f"https://weworkremotely.com/categories/{cat}.rss", json_=False))
        for item in raiz.iter("item"):
            titulo_completo = item.findtext("title") or ""
            empresa, _, titulo = titulo_completo.partition(":")
            if not titulo:
                empresa, titulo = "", titulo_completo
            enlace = item.findtext("link") or ""
            fecha = item.findtext("pubDate") or ""
            try:
                fecha = datetime.strptime(fecha[:25], "%a, %d %b %Y %H:%M:%S").strftime("%Y-%m-%d")
            except ValueError:
                pass
            yield oferta("WeWorkRemotely", enlace.rsplit("/", 1)[-1], titulo, empresa, enlace,
                         item.findtext("description"), item.findtext("region") or "Remoto", fecha)


def fuente_jobicy(cfg):
    filtros = [{"tag": t} for t in cfg.get("tags_jobicy", [])]
    filtros += [{"geo": g} for g in cfg.get("geos_jobicy", [])]
    for filtro in filtros:
        q = urllib.parse.urlencode({"count": 50, **filtro})
        datos = descargar(f"https://jobicy.com/api/v2/remote-jobs?{q}")
        for j in datos.get("jobs", []):
            salario = ""
            if j.get("annualSalaryMin"):
                salario = f"{j.get('salaryCurrency', '')} {j['annualSalaryMin']}-{j.get('annualSalaryMax', '')}/año"
            yield oferta("Jobicy", j.get("id"), j.get("jobTitle"), j.get("companyName"),
                         j.get("url"), j.get("jobDescription") or j.get("jobExcerpt"),
                         j.get("jobGeo") or "Remoto", (j.get("pubDate") or "")[:10], salario)


def fuente_getonbrd(cfg):
    for termino in cfg.get("busquedas_getonbrd", []):
        q = urllib.parse.urlencode({"query": termino, "per_page": 50, "expand": '["company"]'})
        datos = descargar(f"https://www.getonbrd.com/api/v0/search/jobs?{q}")
        for j in datos.get("data", []):
            a = j.get("attributes", {})
            empresa = (((a.get("company") or {}).get("data") or {}).get("attributes") or {}).get("name", "")
            url = (j.get("links") or {}).get("public_url") or f"https://www.getonbrd.com/jobs/{j.get('id')}"
            publicada = ""
            if a.get("published_at"):
                publicada = datetime.fromtimestamp(a["published_at"], timezone.utc).strftime("%Y-%m-%d")
            salario = ""
            if a.get("min_salary"):
                salario = f"USD {a['min_salary']}-{a.get('max_salary') or ''}/mes"
            ubic = "Remoto" if a.get("remote") else ", ".join(a.get("countries") or [])
            desc = " ".join(filter(None, [a.get("description"), a.get("functions"), a.get("desirable")]))
            yield oferta("GetOnBoard", j.get("id"), a.get("title"), empresa, url, desc,
                         ubic, publicada, salario, bool(a.get("remote")))


FUENTES = [fuente_remoteok, fuente_remotive, fuente_weworkremotely, fuente_jobicy, fuente_getonbrd]


# ---------------------------------------------------------------- puntaje

PALABRAS_ES = {"de", "la", "que", "el", "en", "los", "del", "las", "por", "para", "con", "una",
               "nuestro", "nuestra", "buscamos", "experiencia", "trabajo", "equipo", "somos"}
PALABRAS_EN = {"the", "and", "to", "of", "you", "with", "for", "our", "we", "will", "are",
               "is", "your", "team", "experience", "looking"}


def en_espanol(o):
    palabras = re.findall(r"[a-z]+", normalizar(f"{o['titulo']} {o['descripcion']}"))
    es = sum(p in PALABRAS_ES for p in palabras)
    en = sum(p in PALABRAS_EN for p in palabras)
    return es > en


def te_sirve(o, cfg):
    """Solo ofertas en Argentina, o remotas en español."""
    en_argentina = any(contiene(normalizar(o["ubicacion"]), p) for p in cfg["argentina"])
    if o["remoto"]:
        return en_argentina or en_espanol(o)
    return en_argentina


def puntuar(o, cfg):
    """Devuelve (puntaje 0-100, perfil principal, palabras encontradas) o None si se descarta."""
    titulo = normalizar(o["titulo"])
    if any(contiene(titulo, p) for p in cfg["excluir_si_el_titulo_dice"]):
        return None
    if not te_sirve(o, cfg):
        return None

    texto = normalizar(o["descripcion"])
    ubicacion = normalizar(o["ubicacion"])
    todo = f"{titulo} {texto} {ubicacion}"
    ubicacion_sirve = not ubicacion or any(contiene(ubicacion, p) for p in cfg["ubicaciones_que_te_sirven"])

    # presencial fuera de Argentina: no sirve
    if not o["remoto"] and not ubicacion_sirve:
        return None

    # Solo cuenta un perfil si alguna palabra aparece en el título del puesto;
    # las de la descripción suman pero no alcanzan solas (evita mucho ruido).
    mejor_perfil, mejor, encontradas = "", 0, set()
    for perfil, datos in cfg["perfiles"].items():
        en_titulo = [p for p in datos["palabras"] if contiene(titulo, p)]
        if not en_titulo:
            continue
        en_texto = [p for p in datos["palabras"] if p not in en_titulo and contiene(texto, p)]
        puntos = (40 + 15 * (len(en_titulo) - 1) + 5 * len(en_texto)) * datos["prioridad"]
        encontradas.update(en_titulo + en_texto)
        if puntos > mejor:
            mejor_perfil, mejor = perfil, puntos

    if mejor == 0:
        return None

    if any(contiene(todo, p) for p in cfg["bonus"]["palabras"]):
        mejor += cfg["bonus"]["puntos"]
    if any(contiene(todo, p) for p in cfg["restricciones_geograficas"]["palabras"]):
        mejor += cfg["restricciones_geograficas"]["puntos"]
    if not ubicacion_sirve:
        mejor += cfg["penalizacion_otra_ubicacion"]

    return max(0, min(100, round(mejor))), mejor_perfil, ", ".join(sorted(encontradas))


# ---------------------------------------------------------------- base de datos

ESQUEMA = """
CREATE TABLE IF NOT EXISTS ofertas (
    id            TEXT PRIMARY KEY,
    fuente        TEXT,
    titulo        TEXT,
    empresa       TEXT,
    perfil        TEXT,
    puntaje       INTEGER,
    palabras      TEXT,
    ubicacion     TEXT,
    remoto        INTEGER,
    salario       TEXT,
    publicada     TEXT,
    url           TEXT,
    descripcion   TEXT,
    estado        TEXT DEFAULT 'nueva',   -- nueva, interesa, postulada, entrevista, rechazada, descartada
    notas         TEXT DEFAULT '',
    encontrada_en TEXT
);
CREATE INDEX IF NOT EXISTS idx_puntaje ON ofertas(puntaje DESC);
CREATE INDEX IF NOT EXISTS idx_estado  ON ofertas(estado);
"""


def guardar(con, o):
    """Inserta la oferta si no existía. Nunca pisa el estado ni las notas de una ya guardada."""
    cur = con.execute(
        """INSERT OR IGNORE INTO ofertas
           (id, fuente, titulo, empresa, perfil, puntaje, palabras, ubicacion, remoto,
            salario, publicada, url, descripcion, encontrada_en)
           VALUES (:id, :fuente, :titulo, :empresa, :perfil, :puntaje, :palabras, :ubicacion,
                   :remoto, :salario, :publicada, :url, :descripcion, :encontrada_en)""", o)
    return cur.rowcount == 1


def exportar_csv(con):
    filas = con.execute(
        """SELECT puntaje, perfil, titulo, empresa, ubicacion, salario, publicada, fuente,
                  estado, url FROM ofertas ORDER BY estado = 'nueva' DESC, puntaje DESC""").fetchall()
    with open(CSV, "w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig para que Excel lea las tildes
        w = csv.writer(f, delimiter=";")
        w.writerow(["Puntaje", "Perfil", "Puesto", "Empresa", "Ubicación", "Salario",
                    "Publicada", "Fuente", "Estado", "Link"])
        w.writerows(filas)
    print(f"\nExportado a {CSV.name} ({len(filas)} ofertas)")


def generar_pagina(con, ahora):
    """Arma docs/index.html con todas las ofertas, para ver desde el celular."""
    con.row_factory = sqlite3.Row
    filas = con.execute(
        """SELECT titulo, empresa, perfil, puntaje, ubicacion, salario, publicada, fuente, url,
                  estado, encontrada_en, substr(descripcion, 1, 400) AS resumen
           FROM ofertas WHERE estado != 'descartada' ORDER BY puntaje DESC""").fetchall()
    con.row_factory = None
    datos = json.dumps([dict(f) for f in filas], ensure_ascii=False).replace("</", "<\\/")
    pagina = (PLANTILLA.read_text(encoding="utf-8")
              .replace("__OFERTAS__", datos)
              .replace("__ACTUALIZADO__", ahora))
    PAGINA.parent.mkdir(exist_ok=True)
    PAGINA.write_text(pagina, encoding="utf-8")
    print(f"Página generada: {PAGINA.relative_to(CARPETA)} ({len(filas)} ofertas)")


# ---------------------------------------------------------------- principal

def main():
    sys.stdout.reconfigure(encoding="utf-8")
    cfg = json.loads(CONFIG.read_text(encoding="utf-8"))
    ahora = datetime.now(ARGENTINA).strftime("%Y-%m-%d %H:%M")

    con = sqlite3.connect(BASE)
    con.executescript(ESQUEMA)

    # saca las ofertas sin tocar que ya no cumplen los filtros (por si cambiaste config.json)
    con.row_factory = sqlite3.Row
    viejas = [dict(f) for f in con.execute("SELECT * FROM ofertas WHERE estado = 'nueva'")]
    con.row_factory = None
    for o in viejas:
        if not te_sirve(o, cfg):
            con.execute("DELETE FROM ofertas WHERE id = ?", (o["id"],))

    nuevas, vistas = [], set()
    for fuente in FUENTES:
        nombre = fuente.__name__.replace("fuente_", "")
        revisadas = 0
        try:
            for o in fuente(cfg):
                if o["id"] in vistas:
                    continue
                vistas.add(o["id"])
                revisadas += 1
                resultado = puntuar(o, cfg)
                if not resultado or resultado[0] < cfg["puntaje_minimo"]:
                    continue
                o["puntaje"], o["perfil"], o["palabras"] = resultado
                o["encontrada_en"] = ahora
                if guardar(con, o):
                    nuevas.append(o)
            print(f"  {nombre:<15} {revisadas:>5} ofertas revisadas")
        except Exception as e:  # si un portal falla, seguimos con los demás
            print(f"  {nombre:<15} ERROR: {e}")
    con.commit()

    print(f"\n{len(nuevas)} ofertas nuevas que te pueden interesar:\n")
    for o in sorted(nuevas, key=lambda x: -x["puntaje"]):
        print(f"  [{o['puntaje']:>3}] {o['titulo'][:60]} — {o['empresa'][:30]}  ({o['perfil']})")
        print(f"        {o['url']}")

    total = con.execute("SELECT COUNT(*) FROM ofertas").fetchone()[0]
    print(f"\nTotal en la base: {total} ofertas")

    if "--exportar" in sys.argv:
        exportar_csv(con)
    generar_pagina(con, ahora)
    con.close()


if __name__ == "__main__":
    main()
