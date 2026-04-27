import feedparser
import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# =====================================================
# CONFIGURACIÓN
# =====================================================

# CORREO QUE ENVÍA
EMAIL_ORIGEN = "Juanleones1018@gmail.com"

# CLAVE DE APLICACIÓN GMAIL (NO contraseña normal)
CLAVE_APP = "qixi xrfa rdhq clwi"

# CORREO QUE RECIBE
EMAIL_DESTINO = [
    "Juanleones1018@gmail.com",
    "chaverra01@gmail.com",
    "geraldine.ramirez1@udea.edu.co","german.valencia@udea.edu.co", "maria.gvasquez@udea.edu.co"

]

# ARCHIVO EXCEL
archivo_excel = r"C:\BotPazTotal\historial_paz_total.xlsx"

# ARCHIVO CONTROL LINKS YA ENVIADOS
archivo_links = r"C:\BotPazTotal\links_vistos.txt"

# =====================================================
# FUENTES PRO
# =====================================================

fuentes = {
    "Noticias": "https://news.google.com/rss/search?q=%22Paz+Total%22+Colombia&hl=es-419&gl=CO&ceid=CO:es-419",

    "Académicos": "https://news.google.com/rss/search?q=%22Paz+Total%22+site:scielo.org.co+OR+site:redalyc.org+OR+site:doi.org+OR+site:researchgate.net&hl=es-419&gl=CO&ceid=CO:es-419",

    "PDFs": "https://news.google.com/rss/search?q=%22Paz+Total%22+filetype:pdf&hl=es-419&gl=CO&ceid=CO:es-419",

    "YouTube": "https://www.youtube.com/feeds/videos.xml?search_query=Paz+Total+Colombia"
}

# =====================================================
# CARGAR LINKS YA REVISADOS
# =====================================================

if os.path.exists(archivo_links):
    with open(archivo_links, "r", encoding="utf-8") as f:
        links_vistos = set(f.read().splitlines())
else:
    links_vistos = set()

# =====================================================
# BUSCAR NUEVOS RESULTADOS
# =====================================================

nuevos = []

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
# GUARDAR EXCEL
# =====================================================

if nuevos:

    df_nuevo = pd.DataFrame(nuevos)

    if os.path.exists(archivo_excel):
        viejo = pd.read_excel(archivo_excel)
        total = pd.concat([viejo, df_nuevo]).drop_duplicates(subset=["Link"])
    else:
        total = df_nuevo

    total.to_excel(archivo_excel, index=False)

# =====================================================
# GUARDAR LINKS
# =====================================================

with open(archivo_links, "w", encoding="utf-8") as f:
    for link in links_vistos:
        f.write(link + "\n")

# =====================================================
# ENVIAR CORREO SI HAY NOVEDADES
# =====================================================

html = """
<h2>📌 Reporte Diario - Paz Total Colombia</h2>
<p>Fecha: """ + datetime.now().strftime("%Y-%m-%d %H:%M") + """</p>
"""

if nuevos:

    html += f"<p><b>Se detectaron {len(nuevos)} novedades.</b></p>"

    html += """
    <table border="1" cellpadding="6" cellspacing="0">
    <tr>
        <th>Tipo</th>
        <th>Título</th>
        <th>Link</th>
    </tr>
    """

    for item in nuevos:
        html += f"""
        <tr>
            <td>{item['Tipo']}</td>
            <td>{item['Titulo']}</td>
            <td><a href="{item['Link']}">Abrir</a></td>
        </tr>
        """

    html += "</table>"

else:
    html += """
    <p>📭 Hoy no se detectaron novedades nuevas relevantes.</p>
    <p>✅ El bot sigue funcionando correctamente.</p>
    """

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

print("✅ Correo diario enviado")

