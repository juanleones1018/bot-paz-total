import feedparser
import pandas as pd
import os
import smtplib
import requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# =====================================================
# CONFIGURACIÓN
# =====================================================

EMAIL_ORIGEN = os.getenv("EMAIL_ORIGEN")
CLAVE_APP = os.getenv("CLAVE_APP")
EMAIL_DESTINO = [x.strip() for x in os.getenv("EMAIL_DESTINO").split(",")]

archivo_excel = "historial_paz_total.xlsx"
archivo_links = "links_vistos.txt"

# =====================================================
# CREAR EXCEL SI NO EXISTE
# =====================================================

if not os.path.exists(archivo_excel):
    df_base = pd.DataFrame(columns=["Fecha","Tipo","Titulo","Link"])
    df_base.to_excel(archivo_excel, index=False)

# =====================================================
# FUENTES
# =====================================================

fuentes = {
    "Noticias": "https://news.google.com/rss/search?q=%22Paz+Total%22+Colombia&hl=es-419&gl=CO&ceid=CO:es-419",

    "Académicos RSS": "https://news.google.com/rss/search?q=%22Paz+Total%22+site:scielo.org.co+OR+site:redalyc.org+OR+site:doi.org&hl=es-419&gl=CO&ceid=CO:es-419",

    "PDFs": "https://news.google.com/rss/search?q=%22Paz+Total%22+filetype:pdf&hl=es-419&gl=CO&ceid=CO:es-419",

    "YouTube": "https://www.youtube.com/feeds/videos.xml?search_query=Paz+Total+Colombia"
}

# =====================================================
# CARGAR LINKS
# =====================================================

if os.path.exists(archivo_links):
    with open(archivo_links, "r", encoding="utf-8") as f:
        links_vistos = set(f.read().splitlines())
else:
    links_vistos = set()

nuevos = []

# =====================================================
# RSS NORMAL
# =====================================================

for tipo, url in fuentes.items():
    feed = feedparser.parse(url)

    for item in feed.entries[:10]:
        titulo = item.title.strip()
        link = item.link.strip()

        if link not in links_vistos:
            nuevos.append({
                "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Tipo": tipo,
                "Titulo": titulo,
                "Link": link
            })
            links_vistos.add(link)

# =====================================================
# NUEVA FUENTE PRO: OPENALEX (PAPERS REALES)
# =====================================================

try:
    url = "https://api.openalex.org/works?search=Paz%20Total%20Colombia&per-page=5"
    r = requests.get(url).json()

    for item in r["results"]:
        titulo = item["title"]
        link = item["id"]

        if link not in links_vistos:
            nuevos.append({
                "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Tipo": "Académico PRO",
                "Titulo": titulo,
                "Link": link
            })
            links_vistos.add(link)

except:
    print("OpenAlex falló, continuando...")

# =====================================================
# GUARDAR EXCEL
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
# CORREO
# =====================================================

html = f"""
<h2>📌 Reporte Diario - Paz Total Colombia</h2>
<p>Fecha: {datetime.now().strftime("%Y-%m-%d %H:%M")}</p>
"""

if nuevos:

    html += f"<p><b>{len(nuevos)} novedades detectadas</b></p>"

    html += "<table border='1'><tr><th>Tipo</th><th>Título</th><th>Link</th></tr>"

    for item in nuevos:
        html += f"<tr><td>{item['Tipo']}</td><td>{item['Titulo']}</td><td><a href='{item['Link']}'>Abrir</a></td></tr>"

    html += "</table>"

else:
    html += "<p>📭 Sin novedades hoy, bot funcionando correctamente.</p>"

msg = MIMEMultipart()

fecha = datetime.now().strftime("%d/%m")

if nuevos:
    msg["Subject"] = f"📌 Paz Total {fecha} | {len(nuevos)} novedades"
else:
    msg["Subject"] = f"📌 Paz Total {fecha} | Sin novedades"

msg["From"] = EMAIL_ORIGEN
msg["To"] = ", ".join(EMAIL_DESTINO)

msg.attach(MIMEText(html, "html", "utf-8"))

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(EMAIL_ORIGEN, CLAVE_APP)
server.sendmail(EMAIL_ORIGEN, EMAIL_DESTINO, msg.as_string())
server.quit()

print("Correo enviado correctamente")
