import feedparser
import pandas as pd
import os
import smtplib
import requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# =====================================================
# CONFIG
# =====================================================

EMAIL_ORIGEN = os.getenv("EMAIL_ORIGEN")
CLAVE_APP = os.getenv("CLAVE_APP")
EMAIL_DESTINO = [x.strip() for x in os.getenv("EMAIL_DESTINO").split(",")]

archivo_excel = "historial_paz_total.xlsx"
archivo_links = "links_vistos.txt"

# =====================================================
# FILTROS
# =====================================================

def es_relevante(titulo):

    t = titulo.lower()

    actores = [
        "eln",
        "emc",
        "disidencias",
        "marquetalia",
        "agc",
        "clan del golfo",
        "mesa",
        "negociación",
        "diálogo"
    ]

    variables = [
        "paz total",
        "cese al fuego",
        "acuerdo",
        "participación",
        "delegación",
        "garante",
        "mesa"
    ]

    return (
        any(a in t for a in actores)
        and any(v in t for v in variables)
    )

# =====================================================
# CLASIFICACIONES
# =====================================================

def clasificar(titulo):

    t = titulo.lower()

    if "eln" in t:
        return "ELN"

    elif "disidencias" in t or "emc" in t:
        return "Disidencias"

    elif "petro" in t or "gobierno" in t:
        return "Gobierno"

    elif "cese al fuego" in t:
        return "Cese al fuego"

    elif "negociación" in t:
        return "Negociación"

    else:
        return "General"

# =====================================================
# ACTORES
# =====================================================

def detectar_actor(titulo):

    t = titulo.lower()

    if "eln" in t:
        return "ELN"

    elif "marquetalia" in t:
        return "Segunda Marquetalia"

    elif "emc" in t or "disidencias" in t:
        return "EMC / Disidencias"

    elif "agc" in t or "clan del golfo" in t:
        return "AGC"

    elif "urbana" in t or "barrial" in t:
        return "Estructuras urbanas"

    else:
        return "No identificado"

# =====================================================
# PROCESO
# =====================================================

def tipo_proceso(titulo):

    t = titulo.lower()

    if "mesa" in t or "diálogo" in t:
        return "Multilateral"

    elif "cese unilateral" in t:
        return "Unilateral"

    else:
        return "No claro"

# =====================================================
# VARIABLES
# =====================================================

def detectar_variable(titulo):

    t = titulo.lower()

    if "cese al fuego" in t:
        return "Cese al fuego"

    elif "agenda" in t:
        return "Agenda"

    elif "delegación" in t:
        return "Delegaciones"

    elif "garante" in t:
        return "Garantes"

    elif "participación" in t:
        return "Participación social"

    elif "acuerdo" in t:
        return "Acuerdos"

    else:
        return "General"

# =====================================================
# TIPOLOGÍA
# =====================================================

def detectar_tipologia(tipo, link, titulo):

    l = link.lower()
    t = titulo.lower()

    if tipo == "Académico":
        return "Paper académico"

    elif "youtube" in l:
        return "Video"

    elif "spotify" in l or "podcast" in t:
        return "Podcast"

    elif "pares" in l or "indepaz" in l:
        return "Informe / observatorio"

    elif "opinion" in l or "columna" in t:
        return "Opinión"

    else:
        return "Noticia"

# =====================================================
# TERRITORIO
# =====================================================

def detectar_territorio(titulo):

    t = titulo.lower()

    territorios = [
        "medellín",
        "antioquia",
        "buenaventura",
        "chocó",
        "cauca",
        "nariño",
        "catatumbo",
        "arauca",
        "valle de aburrá",
        "quibdó"
    ]

    for territorio in territorios:

        if territorio in t:
            return territorio.title()

    return "No identificado"

# =====================================================
# RESUMEN
# =====================================================

def resumir(titulo):

    palabras = titulo.split()

    return " ".join(palabras[:12]) + "..."

# =====================================================
# CREAR EXCEL
# =====================================================

if not os.path.exists(archivo_excel):

    df_base = pd.DataFrame(columns=[
        "Fecha",
        "Tipo",
        "Tipologia",
        "Actor",
        "Proceso",
        "Variable",
        "Territorio",
        "Categoria",
        "Titulo",
        "Resumen",
        "Link"
    ])

    df_base.to_excel(archivo_excel, index=False)

# =====================================================
# FUENTES
# =====================================================

fuentes = {

    "Noticias":
    "https://news.google.com/rss/search?q=%22Paz+Total%22+Colombia+after:2023&hl=es-419&gl=CO&ceid=CO:es-419",

    "Académico":
    "https://news.google.com/rss/search?q=%22Paz+Total%22+Colombia+(site:scielo.org.co+OR+site:redalyc.org+OR+site:doi.org+OR+site:dialnet.unirioja.es)+after:2023&hl=es-419&gl=CO&ceid=CO:es-419",

    "Podcast":
    "https://news.google.com/rss/search?q=%22Paz+Total%22+(podcast+OR+Spotify+OR+audio)&hl=es-419&gl=CO&ceid=CO:es-419",

    "YouTube":
    "https://www.youtube.com/feeds/videos.xml?search_query=Paz+Total+Colombia"
}

# =====================================================
# LINKS
# =====================================================

if os.path.exists(archivo_links):

    with open(archivo_links, "r", encoding="utf-8") as f:
        links_vistos = set(f.read().splitlines())

else:
    links_vistos = set()

nuevos = []

# =====================================================
# RSS
# =====================================================

for tipo, url in fuentes.items():

    feed = feedparser.parse(url)

    for item in feed.entries[:10]:

        titulo = item.title.strip()
        link = item.link.strip()

        if link not in links_vistos and es_relevante(titulo):

            nuevos.append({

                "Fecha":
                datetime.now().strftime("%Y-%m-%d %H:%M"),

                "Tipo":
                tipo,

                "Tipologia":
                detectar_tipologia(tipo, link, titulo),

                "Actor":
                detectar_actor(titulo),

                "Proceso":
                tipo_proceso(titulo),

                "Variable":
                detectar_variable(titulo),

                "Territorio":
                detectar_territorio(titulo),

                "Categoria":
                clasificar(titulo),

                "Titulo":
                titulo,

                "Resumen":
                resumir(titulo),

                "Link":
                link
            })

            links_vistos.add(link)

# =====================================================
# OPENALEX
# =====================================================

try:

    url = "https://api.openalex.org/works?search=%22Paz%20Total%22%20Colombia&filter=from_publication_date:2023-01-01&sort=publication_date:desc&per-page=10"

    r = requests.get(url).json()

    for item in r["results"]:

        titulo = item["title"]
        link = item["id"]

        if link not in links_vistos and es_relevante(titulo):

            nuevos.append({

                "Fecha":
                datetime.now().strftime("%Y-%m-%d %H:%M"),

                "Tipo":
                "Académico",

                "Tipologia":
                detectar_tipologia("Académico", link, titulo),

                "Actor":
                detectar_actor(titulo),

                "Proceso":
                tipo_proceso(titulo),

                "Variable":
                detectar_variable(titulo),

                "Territorio":
                detectar_territorio(titulo),

                "Categoria":
                clasificar(titulo),

                "Titulo":
                titulo,

                "Resumen":
                resumir(titulo),

                "Link":
                link
            })

            links_vistos.add(link)

except:
    print("OpenAlex falló")

# =====================================================
# PARES + INDEPAZ
# =====================================================

try:

    especiales = [
        "https://www.pares.com.co/feed",
        "https://indepaz.org.co/feed"
    ]

    for url in especiales:

        feed = feedparser.parse(url)

        for item in feed.entries[:5]:

            titulo = item.title.strip()
            link = item.link.strip()

            if link not in links_vistos and es_relevante(titulo):

                nuevos.append({

                    "Fecha":
                    datetime.now().strftime("%Y-%m-%d %H:%M"),

                    "Tipo":
                    "Centro de investigación",

                    "Tipologia":
                    detectar_tipologia(
                        "Centro de investigación",
                        link,
                        titulo
                    ),

                    "Actor":
                    detectar_actor(titulo),

                    "Proceso":
                    tipo_proceso(titulo),

                    "Variable":
                    detectar_variable(titulo),

                    "Territorio":
                    detectar_territorio(titulo),

                    "Categoria":
                    clasificar(titulo),

                    "Titulo":
                    titulo,

                    "Resumen":
                    resumir(titulo),

                    "Link":
                    link
                })

                links_vistos.add(link)

except:
    print("Centros fallaron")

# =====================================================
# EXCEL
# =====================================================

df_nuevo = pd.DataFrame(nuevos)

df_viejo = pd.read_excel(archivo_excel)

total = pd.concat([df_viejo, df_nuevo]).drop_duplicates(subset=["Link"])

total.to_excel(archivo_excel, index=False)

# =====================================================
# GUARDAR LINKS
# =====================================================

with open(archivo_links, "w", encoding="utf-8") as f:

    for link in links_vistos:
        f.write(link + "\n")

# =====================================================
# CORREO (SE MANTIENE IGUAL)
# =====================================================

html = f"""
<h2>📌 Reporte Diario - Paz Total Colombia</h2>
<p>{datetime.now().strftime("%Y-%m-%d %H:%M")}</p>
"""

if nuevos:

    html += "<table border='1'>"
    html += "<tr><th>Actor</th><th>Proceso</th><th>Variable</th><th>Título</th></tr>"

    for item in nuevos:

        html += f"""
        <tr>
            <td>{item['Actor']}</td>
            <td>{item['Proceso']}</td>
            <td>{item['Variable']}</td>
            <td><a href="{item['Link']}">{item['Titulo']}</a></td>
        </tr>
        """

    html += "</table>"

else:

    html += "<p>Sin novedades hoy</p>"

msg = MIMEMultipart()

msg["Subject"] = f"📌 Paz Total | {len(nuevos)} hallazgos"

msg["From"] = EMAIL_ORIGEN

msg["To"] = ", ".join(EMAIL_DESTINO)

msg.attach(MIMEText(html, "html"))

server = smtplib.SMTP("smtp.gmail.com", 587)

server.starttls()

server.login(EMAIL_ORIGEN, CLAVE_APP)

server.sendmail(
    EMAIL_ORIGEN,
    EMAIL_DESTINO,
    msg.as_string()
)

server.quit()

print("Correo enviado correctamente")
