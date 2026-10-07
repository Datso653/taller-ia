# -*- coding: utf-8 -*-
"""
Scraper de convocatorias · demo del Taller de IA (Secretaría de Juventud FUBA, 07/10/2026)

Qué hace, paso a paso:
  1. Entra a las secciones de becas de Fulbright Argentina (fulbright.edu.ar).
  2. Junta los links de todas las convocatorias que aparecen.
  3. Entra a cada una y lee: fecha de cierre, para quién es y los requisitos.
  4. Compara la fecha de cierre con la de hoy: abierta, vencida o "revisar".
  5. Arma una planilla de Excel ordenada por cierre y un archivo de recordatorios (.ics).

Uso:
  python scraper_convocatorias.py              en vivo, desde internet
  python scraper_convocatorias.py --offline    plan B: usa las páginas guardadas en cache/
  python scraper_convocatorias.py --no-abrir   no abre el Excel al terminar

El robots.txt de fulbright.edu.ar permite leer estas páginas. Igual vamos despacio:
una pausa entre pedido y pedido, para no cargar el sitio.
"""
import datetime as dt
import html
import os
import re
import sys
import time
import urllib.request
from concurrent.futures import ThreadPoolExecutor

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
SITIO = 'https://fulbright.edu.ar'
CATEGORIAS = ['abiertas', 'estudiantes-de-grado', 'graduados', 'docentes', 'investigadores']
PAUSA = 0.4                      # segundos entre pedidos
CARPETA = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(CARPETA, 'cache')
NAVEGADOR = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
             '(KHTML, like Gecko) Chrome/129.0 Safari/537.36')
OFFLINE = '--offline' in sys.argv

MESES = {'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6, 'julio': 7,
         'agosto': 8, 'septiembre': 9, 'setiembre': 9, 'octubre': 10, 'noviembre': 11, 'diciembre': 12}
FECHA = r'(\d{1,2})\s+de\s+(' + '|'.join(MESES) + r')(?:\s+(?:de\s+|del\s+)?(\d{4}))?'

# Colores para la terminal
os.system('')  # activa los colores en la consola de Windows
VERDE, ROJO, AMARILLO, GRIS, NEGRITA, FIN = '\033[92m', '\033[91m', '\033[93m', '\033[90m', '\033[1m', '\033[0m'


# ---------------------------------------------------------------------------
# 1. Bajar páginas (con copia local para el modo sin internet)
# ---------------------------------------------------------------------------
def nombre_en_cache(url):
    """Cada página se guarda en cache/ con un nombre sacado de su dirección."""
    partes = url.rstrip('/').split('/')
    if partes[-2] == 'course-category':
        return os.path.join(CACHE, 'categoria_' + partes[-1] + '.html')
    return os.path.join(CACHE, partes[-1][:80] + '.html')


def descargar(url):
    archivo = nombre_en_cache(url)
    if OFFLINE:
        if not os.path.exists(archivo):
            return ''
        with open(archivo, encoding='utf-8') as f:
            return f.read()
    pedido = urllib.request.Request(url, headers={'User-Agent': NAVEGADOR})
    with urllib.request.urlopen(pedido, timeout=30) as r:
        pagina = r.read().decode('utf-8', 'replace')
    os.makedirs(CACHE, exist_ok=True)
    with open(archivo, 'w', encoding='utf-8') as f:
        f.write(pagina)
    time.sleep(PAUSA)
    return pagina


# ---------------------------------------------------------------------------
# 2. Leer el contenido de una página
# ---------------------------------------------------------------------------
def texto_plano(pagina):
    """Saca el HTML y deja el texto, un párrafo o ítem de lista por línea."""
    s = re.sub(r'<(script|style|nav|header|footer)[^>]*>.*?</\1>', ' ', pagina, flags=re.S)
    s = re.sub(r'<li[^>]*>', '\n• ', s)
    s = re.sub(r'<(br|/p|/li|/h\d|/div|/tr|/td|/ul|/ol)[^>]*>', '\n', s)
    s = html.unescape(re.sub(r'<[^>]+>', '', s))
    return [re.sub(r'\s+', ' ', l).strip() for l in s.split('\n') if l.strip()]


def links_de_convocatorias(pagina):
    """En una página de categoría, devuelve {link: título} de cada convocatoria."""
    encontrados = {}
    for link, titulo in re.findall(r'<a[^>]+href="(' + SITIO + r'/course/[^"]+)"[^>]*>(.*?)</a>', pagina, re.S):
        titulo = re.sub(r'\s+', ' ', html.unescape(re.sub(r'<[^>]+>', ' ', titulo))).strip()
        if titulo:
            encontrados.setdefault(link, titulo)
    return encontrados


def buscar_cierre(texto):
    """Busca la fecha de cierre. Devuelve (fecha o None, cómo lo dice la página).
    Si la página no dice el año, la fecha vuelve como None: no lo inventamos."""
    patrones = [
        r'cierre(?:\s+de\s+(?:la\s+)?(?:convocatoria|postulaci[oó]n))?\s*:\s*(?:[a-záéíóú]+\s+)?' + FECHA,  # "Cierre de la convocatoria: 25 de marzo de 2026"
        FECHA + r'\s*:\s*cierre',                                                                       # "10 de noviembre: Cierre de convocatoria"
        r'hasta\s+el\s*:?\s*' + FECHA,                                                                  # "Hasta el 9 de octubre de 2025"
    ]
    for patron in patrones:
        m = re.search(patron, texto, re.I)
        if m:
            dia, mes, anio = m.groups()[-3:]
            dicho = '%s de %s' % (int(dia), mes.lower())
            if not anio:
                return None, dicho
            try:
                return dt.date(int(anio), MESES[mes.lower()], int(dia)), dicho
            except ValueError:
                return None, dicho
    return None, ''


def para_quien(pagina):
    nombres = re.findall(r'class="eltdf-course-category"[^>]*>([^<]+)</a>', pagina)
    nombres = [n.strip() for n in nombres if n.strip().lower() != 'abiertas']
    return ', '.join(dict.fromkeys(nombres))


def requisitos(lineas, cuantos=3):
    """Los primeros ítems de la sección Requisitos."""
    inicios = [n for n, l in enumerate(lineas) if re.match(r'^requisitos\b', l, re.I)]
    if not inicios:
        return ''
    corte = re.compile(r'^(beneficios|calendario|cronograma|c[oó]mo postular|documentaci[oó]n|proceso de selecci)', re.I)
    items = []
    for l in lineas[inicios[-1] + 1:]:
        if corte.match(l):
            break
        if l.startswith('•'):
            items.append(l.lstrip('• ').strip())
    if not items:
        return ''
    return ' · '.join(i if len(i) <= 110 else i[:107] + '...' for i in items[:cuantos])


# ---------------------------------------------------------------------------
# 3. Decidir el estado de cada convocatoria
# ---------------------------------------------------------------------------
def estado(cierre, dicho, hoy):
    if cierre:
        return ('ABIERTA', (cierre - hoy).days) if cierre >= hoy else ('VENCIDA', None)
    return ('REVISAR', None) if dicho else ('SIN FECHA', None)


ORDEN = {'ABIERTA': 0, 'REVISAR': 1, 'SIN FECHA': 2, 'VENCIDA': 3}


def clave_orden(c):
    if c['estado'] == 'ABIERTA':
        return (0, c['cierre'].toordinal())
    if c['estado'] == 'VENCIDA':
        return (3, -c['cierre'].toordinal())
    return (ORDEN[c['estado']], 0)


# ---------------------------------------------------------------------------
# 4. Guardar: planilla de Excel y recordatorios
# ---------------------------------------------------------------------------
def guardar_excel(convocatorias, ruta, hoy):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    libro = Workbook()
    hoja = libro.active
    hoja.title = 'Convocatorias'
    columnas = [('Estado', 12), ('Cierre', 16), ('Faltan', 9), ('Convocatoria', 52), ('Para quién', 22),
                ('¿La web dice "abierta"?', 14), ('Requisitos (los primeros)', 70), ('Link', 40)]
    hoja.append([c for c, _ in columnas])
    for i, (_, ancho) in enumerate(columnas, start=1):
        hoja.column_dimensions[hoja.cell(1, i).column_letter].width = ancho
        celda = hoja.cell(1, i)
        celda.font = Font(bold=True)
        celda.fill = PatternFill('solid', fgColor='EBB416')
        celda.alignment = Alignment(wrap_text=True, vertical='center')
    colores = {'ABIERTA': 'C6EFCE', 'REVISAR': 'FFF2CC', 'SIN FECHA': 'FFF2CC', 'VENCIDA': 'E7E6E6'}

    for c in convocatorias:
        cierre = c['cierre'].strftime('%d/%m/%Y') if c['cierre'] else (c['dicho'] + ' (sin año)' if c['dicho'] else '—')
        faltan = '%d días' % c['faltan'] if c['faltan'] is not None else ''
        hoja.append([c['estado'], cierre, faltan, c['titulo'], c['para_quien'],
                     'Sí' if c['web_dice_abierta'] else 'No', c['requisitos'], c['link']])
        fila = hoja.max_row
        hoja.cell(fila, 1).fill = PatternFill('solid', fgColor=colores[c['estado']])
        hoja.cell(fila, 1).font = Font(bold=True)
        hoja.cell(fila, 4).hyperlink = c['link']
        hoja.cell(fila, 4).font = Font(color='1F4E79', underline='single')
        if c['web_dice_abierta'] and c['estado'] == 'VENCIDA':
            hoja.cell(fila, 6).fill = PatternFill('solid', fgColor='F4CCCC')   # la web dice abierta, pero ya cerró
        for col in range(1, len(columnas) + 1):
            hoja.cell(fila, col).alignment = Alignment(wrap_text=True, vertical='top')

    hoja.freeze_panes = 'A2'
    hoja.auto_filter.ref = hoja.dimensions
    hoja.cell(hoja.max_row + 2, 1).value = 'Revisado el %s. Fuente: %s' % (hoy.strftime('%d/%m/%Y'), SITIO)
    libro.save(ruta)


def guardar_recordatorios(convocatorias, ruta, hoy):
    """Un recordatorio 3 días antes del cierre de cada convocatoria abierta (archivo .ics)."""
    eventos = []
    for c in convocatorias:
        if c['estado'] != 'ABIERTA':
            continue
        aviso = max(hoy, c['cierre'] - dt.timedelta(days=3))
        eventos += ['BEGIN:VEVENT',
                    'UID:%s@taller-ia' % abs(hash(c['link'])),
                    'DTSTAMP:%s' % dt.datetime.now().strftime('%Y%m%dT%H%M%S'),
                    'DTSTART;VALUE=DATE:%s' % aviso.strftime('%Y%m%d'),
                    'SUMMARY:Cierra el %s: %s' % (c['cierre'].strftime('%d/%m'), c['titulo'][:60]),
                    'DESCRIPTION:%s' % c['link'],
                    'END:VEVENT']
    contenido = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Taller de IA//Convocatorias//ES'] + eventos + ['END:VCALENDAR']
    with open(ruta, 'w', encoding='utf-8', newline='') as f:
        f.write('\r\n'.join(contenido) + '\r\n')
    return len(eventos) // 7


# ---------------------------------------------------------------------------
# Todo junto
# ---------------------------------------------------------------------------
def main():
    hoy = dt.date.today()
    print('\n%sBuscando convocatorias en Fulbright Argentina%s  (hoy: %s)%s\n'
          % (NEGRITA, FIN, hoy.strftime('%d/%m/%Y'), '  [modo sin internet]' if OFFLINE else ''))

    # Paso 1 y 2: recorrer las secciones y juntar los links
    todas, abiertas_segun_la_web = {}, set()
    for categoria in CATEGORIAS:
        links = links_de_convocatorias(descargar('%s/course-category/%s/' % (SITIO, categoria)))
        print('  Sección %-22s %2d convocatorias' % (categoria, len(links)))
        for link, titulo in links.items():
            todas.setdefault(link, titulo)
            if categoria == 'abiertas':
                abiertas_segun_la_web.add(link)
    print('\n  %d convocatorias distintas. Entro a cada una...\n' % len(todas))

    # Paso 3 y 4: entrar a cada convocatoria y leerla
    convocatorias = []
    with ThreadPoolExecutor(max_workers=4) as grupo:            # de a 4 páginas a la vez
        paginas = grupo.map(descargar, list(todas))
    for n, ((link, titulo), pagina) in enumerate(zip(todas.items(), paginas), start=1):
        lineas = texto_plano(pagina)
        cierre, dicho = buscar_cierre('\n'.join(lineas))
        est, faltan = estado(cierre, dicho, hoy)
        convocatorias.append({'titulo': titulo, 'link': link, 'cierre': cierre, 'dicho': dicho,
                              'estado': est, 'faltan': faltan, 'para_quien': para_quien(pagina),
                              'requisitos': requisitos(lineas), 'web_dice_abierta': link in abiertas_segun_la_web})
        color = {'ABIERTA': VERDE, 'VENCIDA': GRIS, 'REVISAR': AMARILLO, 'SIN FECHA': AMARILLO}[est]
        detalle = (cierre.strftime('%d/%m/%Y') if cierre else (dicho + ' (no dice el año)' if dicho else 'no encontré la fecha'))
        extra = ' · faltan %d días' % faltan if faltan is not None else ''
        aviso = ('  %s<- la web la lista como abierta%s' % (ROJO, FIN)) if (link in abiertas_segun_la_web and est == 'VENCIDA') else ''
        print('  [%2d/%d] %-58s %s%-9s%s %s%s%s' % (n, len(todas), titulo[:58], color, est, FIN, detalle, extra, aviso))

    # Ordenar y guardar
    convocatorias.sort(key=clave_orden)
    excel = os.path.join(CARPETA, 'convocatorias_%s.xlsx' % hoy.isoformat())
    guardar_excel(convocatorias, excel, hoy)
    n_avisos = guardar_recordatorios(convocatorias, os.path.join(CARPETA, 'recordatorios.ics'), hoy)

    cuenta = {e: sum(1 for c in convocatorias if c['estado'] == e) for e in ORDEN}
    print('\n%sResumen%s' % (NEGRITA, FIN))
    print('  La web lista %d como "abiertas".' % len(abiertas_segun_la_web))
    print('  Abiertas de verdad hoy: %s%d%s · para revisar: %d · vencidas: %d'
          % (VERDE, cuenta['ABIERTA'], FIN, cuenta['REVISAR'] + cuenta['SIN FECHA'], cuenta['VENCIDA']))
    print('  Planilla: %s' % os.path.basename(excel))
    print('  Recordatorios: recordatorios.ics (%d, tres días antes de cada cierre)\n' % n_avisos)

    if '--no-abrir' not in sys.argv and hasattr(os, 'startfile'):
        os.startfile(excel)


if __name__ == '__main__':
    main()
