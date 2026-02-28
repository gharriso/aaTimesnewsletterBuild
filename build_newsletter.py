"""
AATimes newsletter auto-builder.

Fetches all data from aatimes.org.au and builds the weekly newsletter.
Run every Saturday/Sunday with no edits:

    python3 build_newsletter.py

Output: aatimes_YYYYMMDD.docx  (YYYYMMDD = upcoming Monday)
"""
import zipfile, re, os, io, urllib.request
from datetime import date, timedelta
from html import unescape as html_unescape

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC      = os.path.join(BASE_DIR, 'sampleNewsletter.docx')
EVENTS_URL   = 'https://aatimes.org.au/events/'
MEETINGS_URL = 'https://aatimes.org.au/meetings/'
UA = {'User-Agent': 'Mozilla/5.0 (compatible; AATimes-builder/2.0)'}

MONTH_MAP = {
    'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,
    'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12,
}
MONTH_NAMES = {v: k.capitalize() for k, v in MONTH_MAP.items()}   # 1→'Jan'
FULL_MONTHS = [
    'January','February','March','April','May','June',
    'July','August','September','October','November','December',
]


# ── Date utilities ────────────────────────────────────────────────────────────

def next_monday():
    """Return the date of the upcoming Monday."""
    today = date.today()
    days  = (7 - today.weekday()) % 7 or 7
    return today + timedelta(days=days)

def ordinal(n):
    if 11 <= (n % 100) <= 13:
        return f'{n}th'
    return f'{n}' + {1:'st', 2:'nd', 3:'rd'}.get(n % 10, 'th')

def fmt_date(d):
    """'Monday 2nd Mar'"""
    return f'{d.strftime("%A")} {ordinal(d.day)} {d.strftime("%b")}'

def fmt_month_year(d):
    """'March 2026'"""
    return f'{FULL_MONTHS[d.month - 1]} {d.year}'

def parse_date(text, today_year=None):
    """
    Parse 'Monday 2nd Mar', '27 Feb', '14th April 2026' → date.
    Returns None on failure.
    """
    if today_year is None:
        today_year = date.today().year
    m = re.search(r'(\d{1,2})(?:st|nd|rd|th)?\s+(\w{3,9})(?:\s+(\d{4}))?',
                  text, re.IGNORECASE)
    if not m:
        return None
    day   = int(m.group(1))
    mon   = MONTH_MAP.get(m.group(2).lower()[:3])
    if not mon:
        return None
    year  = int(m.group(3)) if m.group(3) else today_year
    today = date.today()
    try:
        d = date(year, mon, day)
        # If the resulting date is more than 6 months in the past, try next year
        if d < today - timedelta(days=180):
            d = date(year + 1, mon, day)
        return d
    except ValueError:
        return None

def extract_time(text):
    """Extract '7:30pm', '10:30am', etc. from text."""
    m = re.search(r'\b(\d{1,2}(?::\d{2})?\s*(?:am|pm))\b', text, re.IGNORECASE)
    return m.group(1).strip() if m else ''


# ── HTML utilities ────────────────────────────────────────────────────────────

def fetch(url):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=20) as r:
        return r.read().decode('utf-8', errors='replace')

def strip_tags(fragment, br_to_newline=True):
    """Strip HTML tags; optionally replace <br> with newlines."""
    if br_to_newline:
        fragment = re.sub(r'<br\s*/?>', '\n', fragment, flags=re.IGNORECASE)
        fragment = re.sub(r'</p>', '\n', fragment, flags=re.IGNORECASE)
    fragment = re.sub(r'<[^>]+>', ' ', fragment)
    return html_unescape(fragment)

def text_lines(fragment):
    """Return non-empty stripped text lines from an HTML fragment."""
    return [l.strip() for l in strip_tags(fragment).splitlines() if l.strip()]

def xml_safe(text):
    """Make text safe for XML content (after HTML-unescaping)."""
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;'))

def get_div_block(html, pos):
    """
    Starting at pos (where '<div' begins), return the complete div block
    including its matching closing tag.
    """
    depth = 0
    i = pos
    n = len(html)
    while i < n:
        if re.match(r'<div\b', html[i:]):
            depth += 1
            end = html.find('>', i)
            i = end + 1 if end >= 0 else i + 4
        elif html[i:i+6].lower() == '</div>':
            depth -= 1
            if depth == 0:
                return html[pos:i + 6]
            i += 6
        else:
            i += 1
    return html[pos:]


# ── Events page scraper ───────────────────────────────────────────────────────

def scrape_events(html):
    """
    Parse all event blocks from the events listing page.
    Returns list of dicts: date_start, date_end, date_lines, time_str,
                           title, description, venue, image_url
    """
    events = []
    today_year = date.today().year

    # Find every <div that has both 'row' and 'event' in its class attribute
    for m in re.finditer(r'<div\b[^>]*\bclass=[\'"]([^\'"]*)[\'"][^>]*>', html):
        classes = m.group(1).split()
        if 'row' not in classes or 'event' not in classes:
            continue
        block = get_div_block(html, m.start())
        ev = _parse_event_block(block, today_year)
        if ev:
            events.append(ev)

    return events

def _parse_event_block(block, today_year):
    """Parse one event div block → dict, or None if unparseable."""
    # ── Date ──
    date_m = re.search(
        r"class=['\"][^'\"]*\bdate\b[^'\"]*['\"][^>]*>(.*?)</a>",
        block, re.DOTALL | re.IGNORECASE)
    if not date_m:
        return None
    raw_date = re.sub(r'\s+', ' ', strip_tags(date_m.group(1))).strip()
    # Remove timezone
    raw_date = re.sub(r'\b(?:AEDT|AEST)\b', '', raw_date).strip()

    # Is this a date range?
    is_range = ' - ' in raw_date
    time_str = extract_time(raw_date)

    if is_range:
        halves     = raw_date.split(' - ', 1)
        date_start = parse_date(halves[0], today_year)
        date_end   = parse_date(halves[1], today_year)
        # Build display lines
        start_str  = fmt_date(date_start) if date_start else halves[0].strip()
        end_str    = fmt_date(date_end)   if date_end   else halves[1].strip()
        date_lines = [start_str + ' \u2013', end_str]  # en-dash
        if time_str:
            date_lines.append(time_str)
    else:
        date_start = parse_date(raw_date, today_year)
        if not date_start:
            return None
        date_end   = None
        date_lines = [fmt_date(date_start)]
        if time_str:
            date_lines.append(time_str)

    # ── Title ──
    title_m = re.search(
        r"class=['\"][^'\"]*\btitle\b[^'\"]*['\"][^>]*>.*?<a[^>]*>(.*?)</a>",
        block, re.DOTALL | re.IGNORECASE)
    if not title_m:
        return None
    title = strip_tags(title_m.group(1)).strip()
    if not title:
        return None

    # ── Description ──
    desc_m = re.search(
        r"class=['\"][^'\"]*\bdescription\b[^'\"]*['\"][^>]*>(.*?)</div>",
        block, re.DOTALL | re.IGNORECASE)
    description = text_lines(desc_m.group(1)) if desc_m else []

    # ── Venue ──
    venue = []
    venue_m = re.search(
        r'VENUE</b></small></div>(.*?)(?=<a\b[^>]*maps\.google|</div>\s*</div>)',
        block, re.DOTALL | re.IGNORECASE)
    if venue_m:
        for div_m in re.finditer(r'<div[^>]*>(.*?)</div>',
                                 venue_m.group(1), re.DOTALL | re.IGNORECASE):
            line = strip_tags(div_m.group(1)).strip()
            if line and 'direction' not in line.lower():
                venue.append(line)

    # ── Image ──
    img_m = re.search(
        r"(?:src|href)=['\"]([^'\"]*calendar\.aatimes\.org\.au[^'\"]+\.(?:jpe?g|png))['\"]",
        block, re.IGNORECASE)
    image_url = img_m.group(1).strip() if img_m else None

    return {
        'date_start':  date_start,
        'date_end':    date_end,
        'date_lines':  date_lines,
        'time_str':    time_str,
        'title':       title,
        'description': description,
        'venue':       venue,
        'image_url':   image_url,
        'has_flyer':   False,   # filled in later after image download
    }


# ── Meetings page scraper ─────────────────────────────────────────────────────

def scrape_meeting_changes(html):
    """
    Parse new and changed meetings from /meetings/.
    Returns {'new': [...], 'changed': [...]}
    Each entry: {'location', 'day_time', 'title', 'details'}
    """
    result = {'new': [], 'changed': []}

    for m in re.finditer(r'<div\b[^>]*\bclass=[\'"]([^\'"]*)[\'"][^>]*>', html):
        classes = m.group(1).split()
        if 'meeting-box' not in classes:
            continue
        kind = ('new'     if 'meeting-new'    in classes else
                'changed' if 'meeting-change' in classes else None)
        if not kind:
            continue
        block = get_div_block(html, m.start())
        entry = _parse_meeting_block(block)
        if entry:
            result[kind].append(entry)

    return result

def _parse_meeting_block(block):
    """Parse one meeting-box div → dict or None."""
    # Title (first h4 > a, skip the badge)
    title_m = re.search(r'<h4\b[^>]*>.*?<a[^>]*>(.*?)</a>', block, re.DOTALL | re.IGNORECASE)
    if not title_m:
        return None
    title = strip_tags(title_m.group(1)).strip()
    # Remove badge text if any
    title = re.sub(r'\b(?:New|Changed)\b', '', title, flags=re.IGNORECASE).strip()
    if not title:
        return None

    # Day/time: <h4 class='day_time'>
    dt_m = re.search(
        r"class=['\"]day_time['\"][^>]*>.*?<a[^>]*>(.*?)</a>",
        block, re.DOTALL | re.IGNORECASE)
    if not dt_m:
        return None
    raw_dt = re.sub(r'\s+', ' ', strip_tags(dt_m.group(1))).strip()
    raw_dt = re.sub(r'\b(?:AEDT|AEST)\b', '', raw_dt).strip()
    # raw_dt is like "Monday 7:00pm"
    # Extract day name + time
    day_m = re.search(r'(\w+day)', raw_dt, re.IGNORECASE)
    time_s = extract_time(raw_dt)
    day_time_str = ' '.join(filter(None, [
        day_m.group(1) if day_m else '',
        time_s,
    ]))

    # Venue / address block
    addr_m = re.search(r"class=['\"]address_block['\"][^>]*>(.*?)</div>",
                       block, re.DOTALL | re.IGNORECASE)
    venue_name = ''
    address    = ''
    if addr_m:
        addr_block = addr_m.group(1)
        h5_m = re.search(r'<h5[^>]*>(.*?)</h5>', addr_block, re.DOTALL | re.IGNORECASE)
        venue_name = strip_tags(h5_m.group(1)).strip() if h5_m else ''
        # Remaining text after h5
        addr_text = re.sub(r'<h5[^>]*>.*?</h5>', '', addr_block, flags=re.DOTALL | re.IGNORECASE)
        addr_lines = text_lines(addr_text)
        address = ', '.join(addr_lines) if addr_lines else ''

    # Location = suburb from address (last comma-delimited part before VIC/NSW etc.)
    suburb_m = re.search(r'([A-Z][a-zA-Z\s]+)\s+VIC\b', address)
    location = suburb_m.group(1).strip() if suburb_m else venue_name.split(',')[0].strip()

    details = []
    if venue_name:
        details.append(venue_name)
    if address:
        details.append(address)

    return {
        'location': location,
        'day_time': [day_time_str],
        'title':    title,
        'details':  details,
    }


# ── Image downloading ─────────────────────────────────────────────────────────

MAX_EVENT_LINES = 220   # approx 2 pages of 2-column A4 at 9pt (~110 lines/page)

IMG_W_EMU = 2238000   # ~2.45" width per image column
COL_H     = [-25400, 2270125, 4625340]   # 3-column H positions (EMU from margin)
ROW1_V    = 1168000                      # ~1.28" from page top (EMU)

def get_jpeg_dims(data):
    i = 2
    while i < len(data) - 9:
        if data[i] != 0xFF:
            break
        if data[i + 1] in (0xC0, 0xC1, 0xC2):
            h = (data[i+5] << 8) | data[i+6]
            w = (data[i+7] << 8) | data[i+8]
            return w, h
        i += 2 + ((data[i+2] << 8) | data[i+3])
    return None, None

def download_image(url):
    """Download image bytes from URL. Returns (data, width_px, height_px)."""
    req = urllib.request.Request(url.strip(), headers=UA)
    with urllib.request.urlopen(req, timeout=15) as r:
        data = r.read()
    w, h = get_jpeg_dims(data)
    return data, (w or 768), (h or 1086)

def scale_emu(wp, hp):
    cx = IMG_W_EMU
    cy = int(IMG_W_EMU * hp / wp)
    return cx, cy


# ── Page-2 image anchor XML ───────────────────────────────────────────────────

def make_image_anchor(h_emu, v_emu, cx, cy, rId, idx):
    a1 = f'{0xA0000000 + idx * 4:08X}'
    a2 = f'{0xA0000001 + idx * 4:08X}'
    b1 = f'{0xB0000000 + idx * 4:08X}'
    b2 = f'{0xB0000001 + idx * 4:08X}'
    oid = 200000 + idx * 10
    iid = 200001 + idx * 10
    pid = 200002 + idx * 10
    name = f'FlierImage{idx + 1}'
    return (
        f'<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
        f'<wp:anchor distT="45720" distB="45720" distL="114300" distR="114300"'
        f' simplePos="0" relativeHeight="{251658240 + idx}" behindDoc="0"'
        f' locked="1" layoutInCell="1" allowOverlap="1"'
        f' wp14:anchorId="{a1}" wp14:editId="{a2}">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="margin"><wp:posOffset>{h_emu}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{v_emu}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="9525" b="8255"/>'
        f'<wp:wrapSquare wrapText="bothSides"/>'
        f'<wp:docPr id="{oid}" name="{name}"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wps:wsp><wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr>'
        f'<wps:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
        f'<a:ln w="15875"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        f'<a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a:ln></wps:spPr>'
        f'<wps:txbx><w:txbxContent><w:p w14:paraId="{b1}" w14:textId="{b2}">'
        f'<w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"'
        f' wp14:anchorId="{b2}" wp14:editId="{b1}">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{iid}" name="{name} inline"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{pid}" name="{name} pic"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill rotWithShape="1"><a:blip r:embed="{rId}"/>'
        f'<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln><a:noFill/></a:ln></pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic></wp:inline>'
        f'</w:drawing></w:r></w:p></w:txbxContent></wps:txbx>'
        f'<wps:bodyPr rot="0" vert="horz" wrap="square"'
        f' lIns="7200" tIns="7200" rIns="7200" bIns="7200" anchor="t" anchorCtr="0">'
        f'<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>'
        f'<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>0</wp14:pctWidth></wp14:sizeRelH>'
        f'<wp14:sizeRelV relativeFrom="margin"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV>'
        f'</wp:anchor></w:drawing></mc:Choice></mc:AlternateContent>'
    )


# ── OOXML row builders ────────────────────────────────────────────────────────

def make_header_row(text):
    """Grey month banner."""
    t = xml_safe(text)
    return (
        f'<w:tr w14:paraId="AAAAAAAA" w14:textId="77777777">'
        f'<w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr>'
        f'<w:tcW w:w="5308" w:type="dxa"/><w:gridSpan w:val="4"/>'
        f'<w:tcBorders>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        f'</w:tcBorders>'
        f'<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>'
        f'</w:tcPr><w:p><w:pPr>'
        f'<w:pStyle w:val="AADate"/><w:ind w:left="0"/><w:jc w:val="left"/>'
        f'<w:rPr><w:rFonts w:cs="Arial"/><w:bCs/>'
        f'<w:color w:val="800000"/><w:szCs w:val="18"/></w:rPr>'
        f'</w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:cs="Arial"/><w:bCs/>'
        f'<w:color w:val="800000"/><w:szCs w:val="18"/></w:rPr>'
        f'<w:t>{t}</w:t></w:r>'
        f'</w:p></w:tc></w:tr>'
    )

def make_event_row(date_lines, title, details, has_flyer=False):
    """
    date_lines: list of (text, color) tuples.  color None = blue, 'red' = red.
    title:      plain text
    details:    list of plain-text strings ('' = blank line)
    """
    date_paras = ''
    for txt, color in date_lines:
        t = xml_safe(txt)
        if color == 'red':
            date_paras += (
                f'<w:p><w:pPr><w:pStyle w:val="AADate"/></w:pPr>'
                f'<w:r><w:rPr><w:color w:val="C00000"/></w:rPr>'
                f'<w:t xml:space="preserve">{t}</w:t></w:r></w:p>'
            )
        else:
            date_paras += (
                f'<w:p><w:pPr><w:pStyle w:val="AADate"/></w:pPr>'
                f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'
            )
    date_paras += '<w:p><w:pPr><w:pStyle w:val="AADate"/></w:pPr></w:p>'

    event_paras = (
        f'<w:p><w:pPr><w:pStyle w:val="AAHeading"/></w:pPr>'
        f'<w:r><w:t>{xml_safe(title)}</w:t></w:r></w:p>'
        f'<w:p/>'
    )
    for line in details:
        if line == '':
            event_paras += '<w:p/>'
        else:
            event_paras += f'<w:p><w:r><w:t>{xml_safe(line)}</w:t></w:r></w:p>'

    return (
        f'<w:tr w14:paraId="BBBBBBBB" w14:textId="77777777">'
        f'<w:trPr><w:cantSplit/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="1668" w:type="dxa"/>'
        f'<w:gridSpan w:val="2"/></w:tcPr>{date_paras}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="3640" w:type="dxa"/>'
        f'<w:gridSpan w:val="2"/></w:tcPr>{event_paras}</w:tc>'
        f'</w:tr>'
    )

def make_section_header_row(text):
    """Grey section header (New Meetings / Recently changed / etc.) – 8pt."""
    t = xml_safe(text)
    return (
        f'<w:tr w14:paraId="CCCCCCCC" w14:textId="77777777">'
        f'<w:tblPrEx><w:tblBorders>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="auto"/>'
        f'<w:insideH w:val="single" w:sz="6" w:space="0" w:color="auto"/>'
        f'</w:tblBorders></w:tblPrEx>'
        f'<w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr>'
        f'<w:tcW w:w="5298" w:type="dxa"/><w:gridSpan w:val="4"/>'
        f'<w:tcBorders>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        f'<w:left w:val="single" w:sz="6" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="nil"/>'
        f'<w:right w:val="single" w:sz="6" w:space="0" w:color="auto"/>'
        f'</w:tcBorders>'
        f'<w:shd w:val="clear" w:color="auto" w:fill="E6E6E6"/>'
        f'</w:tcPr><w:p><w:pPr>'
        f'<w:pStyle w:val="AADate"/><w:jc w:val="left"/>'
        f'<w:rPr><w:color w:val="CC0066"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
        f'</w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:cs="Arial"/><w:bCs/><w:color w:val="800000"/>'
        f'<w:kern w:val="32"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
        f'<w:t xml:space="preserve">{t}  </w:t></w:r>'
        f'</w:p></w:tc></w:tr>'
    )

def make_new_meeting_row(location, day_time, title, details):
    """Yellow (FFFFCC) row – 8pt – for New / Changed / Closed meetings."""
    SZ = '<w:sz w:val="16"/><w:szCs w:val="16"/>'
    BORDERS = '<w:tcBorders><w:top w:val="nil"/><w:bottom w:val="nil"/></w:tcBorders>'
    FILL    = '<w:shd w:val="clear" w:color="auto" w:fill="FFFFCC"/>'

    day_paras = ''.join(
        f'<w:p><w:pPr><w:pStyle w:val="AADate"/><w:rPr>{SZ}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{SZ}</w:rPr>'
        f'<w:t xml:space="preserve">{xml_safe(part)}</w:t></w:r></w:p>'
        for part in day_time
    )

    detail_paras = (
        f'<w:p><w:pPr><w:pStyle w:val="AAHeading"/><w:rPr>{SZ}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{SZ}</w:rPr>'
        f'<w:t xml:space="preserve">{xml_safe(title)} </w:t></w:r></w:p>'
    )
    for line in details:
        if line == '':
            detail_paras += f'<w:p><w:pPr><w:rPr>{SZ}</w:rPr></w:pPr></w:p>'
        else:
            detail_paras += (
                f'<w:p><w:pPr><w:rPr>{SZ}</w:rPr></w:pPr>'
                f'<w:r><w:rPr>{SZ}</w:rPr>'
                f'<w:t>{xml_safe(line)}</w:t></w:r></w:p>'
            )

    return (
        f'<w:tr w14:paraId="DDDDDDDD" w14:textId="77777777">'
        f'<w:trPr><w:cantSplit/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="1365" w:type="dxa"/>{BORDERS}{FILL}</w:tcPr>'
        f'<w:p><w:pPr><w:pStyle w:val="AADate"/><w:rPr>{SZ}</w:rPr></w:pPr>'
        f'<w:r><w:rPr>{SZ}</w:rPr><w:t>{xml_safe(location)}</w:t></w:r></w:p>'
        f'</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="968" w:type="dxa"/>'
        f'<w:gridSpan w:val="2"/>{BORDERS}{FILL}</w:tcPr>'
        f'{day_paras}</w:tc>'
        f'<w:tc><w:tcPr><w:tcW w:w="2965" w:type="dxa"/>{BORDERS}{FILL}</w:tcPr>'
        f'{detail_paras}</w:tc>'
        f'</w:tr>'
    )


# ── Header date updater ───────────────────────────────────────────────────────

def update_header_date(xml, target_date):
    """Update the date in a header XML to match target_date."""
    day     = str(target_date.day)
    ord_sfx = ordinal(target_date.day)[len(day):]   # 'st', 'nd', 'rd', 'th'
    month   = FULL_MONTHS[target_date.month - 1]
    year    = str(target_date.year)

    # Day number – just before the superscript ordinal run
    xml = re.sub(
        r'<w:t>(\d+)</w:t>(</w:r><w:r[^>]*><w:rPr>.*?<w:vertAlign)',
        rf'<w:t>{day}</w:t>\2',
        xml, count=1, flags=re.DOTALL,
    )
    # Ordinal superscript
    xml = re.sub(
        r'(<w:vertAlign w:val="superscript"/></w:rPr><w:t>)\w+(</w:t>)',
        rf'\g<1>{ord_sfx}\2',
        xml, count=1,
    )
    # Month — format A: with xml:space="preserve"
    xml = re.sub(
        r'<w:t xml:space="preserve"> (?:January|February|March|April|May|June|'
        r'July|August|September|October|November|December)</w:t>',
        f'<w:t xml:space="preserve"> {month}</w:t>',
        xml, count=1,
    )
    # Month — format B: without xml:space (some headers)
    xml = re.sub(
        r'<w:t>(?:January|February|March|April|May|June|'
        r'July|August|September|October|November|December)</w:t>',
        f'<w:t>{month}</w:t>',
        xml, count=1,
    )
    # Year — format A: split " 202" + "X"
    xml = re.sub(
        r'(> 202</w:t>.*?<w:t>)\d(</w:t>)',
        rf'\g<1>{year[-1]}\2',
        xml, count=1, flags=re.DOTALL,
    )
    # Year — format B: combined " 20YY"
    xml = re.sub(
        r'(<w:t xml:space="preserve"> )20\d\d(</w:t>)',
        rf'\g<1>{year}\2',
        xml, count=1,
    )
    return xml


# ── Page estimator ────────────────────────────────────────────────────────────

def estimate_event_lines(ev):
    """
    Estimate the number of text lines an event row will occupy.
    The row height is the taller of the two cells.
    """
    date_col    = len(ev['date_lines']) + 2          # date lines + trailing blank
    detail_col  = (
        2                                            # title + blank para
        + len(ev['description'])
        + (1 if ev['description'] and ev['venue'] else 0)
        + len(ev['venue'])
    )
    return max(date_col, detail_col)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    monday = next_monday()
    dst_name = f'aatimes{monday.strftime("%Y%m%d")}.docx'
    dst      = os.path.join(BASE_DIR, dst_name)

    print(f'Building newsletter for Monday {fmt_date(monday)} {monday.year}')
    print(f'Output: {dst_name}')
    print()

    # ── Fetch and parse events ──
    print('Fetching events from aatimes.org.au/events/ ...')
    try:
        events_html = fetch(EVENTS_URL)
        all_events  = scrape_events(events_html)
        print(f'  Found {len(all_events)} events total.')
    except Exception as e:
        print(f'  ERROR fetching events: {e}')
        all_events = []

    # Filter to events starting on or after next Monday
    upcoming = [ev for ev in all_events if ev['date_start'] and ev['date_start'] >= monday]
    upcoming.sort(key=lambda e: e['date_start'])

    # Cap to ~2 pages of content
    total_lines = 0
    last_month  = None
    capped      = []
    for ev in upcoming:
        month_label  = fmt_month_year(ev['date_start'])
        header_lines = 1 if month_label != last_month else 0
        ev_lines     = estimate_event_lines(ev)
        if total_lines + header_lines + ev_lines > MAX_EVENT_LINES:
            break
        capped.append(ev)
        total_lines += header_lines + ev_lines
        last_month   = month_label
    upcoming = capped
    print(f'  {len(upcoming)} upcoming from {fmt_date(monday)} (capped at ~2 pages).')

    # ── Fetch and parse meeting changes ──
    print('Fetching meeting changes from aatimes.org.au/meetings/ ...')
    try:
        meetings_html = fetch(MEETINGS_URL)
        changes = scrape_meeting_changes(meetings_html)
        print(f'  New: {len(changes["new"])},  Changed: {len(changes["changed"])}')
    except Exception as e:
        print(f'  ERROR fetching meetings: {e}')
        changes = {'new': [], 'changed': []}

    # ── Download flyer images (up to 6 for page 2 grid) ──
    print('Downloading flyer images ...')
    flyer_events  = [ev for ev in upcoming if ev['image_url']][:6]
    image_store   = {}   # rId → bytes
    image_meta    = {}   # rId → (fn, wp, hp)

    for i, ev in enumerate(flyer_events):
        rId = f'rId{200 + i}'
        fn  = f'img_auto_{i:02d}.jpg'
        try:
            data, wp, hp = download_image(ev['image_url'])
            image_store[rId] = data
            image_meta[rId]  = (fn, wp, hp)
            ev['has_flyer']  = True
            ev['_rId']       = rId
            ev['_fn']        = fn
            ev['_wp']        = wp
            ev['_hp']        = hp
            print(f'  ✓ {ev["title"][:50]}  ({wp}×{hp})')
        except Exception as e:
            print(f'  ✗ {ev["title"][:50]}: {e}')

    # ── Build table rows ──
    rows        = []
    last_month  = None

    for ev in upcoming:
        # Month header when month changes
        month_label = fmt_month_year(ev['date_start'])
        if month_label != last_month:
            rows.append(make_header_row(month_label))
            last_month = month_label

        # Date column lines: (text, color)  — time is plain, only "See Next Page" is red
        date_col = [(dl, None) for dl in ev['date_lines']]
        if ev.get('has_flyer'):
            date_col.append(('See Next Page', 'red'))

        # Details column: description first, then blank line, then venue
        details = list(ev['description'])
        if ev['venue']:
            if details:
                details.append('')
            details.extend(ev['venue'])

        rows.append(make_event_row(date_col, ev['title'], details))

    # ── New meetings section ──
    if changes['new']:
        rows.append(make_section_header_row('New Meetings'))
        for m in changes['new']:
            rows.append(make_new_meeting_row(
                m['location'], m['day_time'], m['title'], m['details']))

    # ── Recently changed meetings section ──
    if changes['changed']:
        rows.append(make_section_header_row('Recently changed meetings'))
        for m in changes['changed']:
            rows.append(make_new_meeting_row(
                m['location'], m['day_time'], m['title'], m['details']))

    # ── Build page-2 image anchor XML ──
    image_runs   = ''
    row1_max_cy  = 0
    for i, ev in enumerate(flyer_events):
        if not ev.get('has_flyer'):
            continue
        rId       = ev['_rId']
        cx, cy    = scale_emu(ev['_wp'], ev['_hp'])
        col_idx   = i % 3
        row_idx   = i // 3
        h_emu     = COL_H[col_idx]
        if row_idx == 0:
            v_emu = ROW1_V
            row1_max_cy = max(row1_max_cy, cy)
        else:
            v_emu = ROW1_V + row1_max_cy + 100000
        anchor_xml = make_image_anchor(h_emu, v_emu, cx, cy, rId, i)
        image_runs += (
            f'<w:r><w:rPr><w:noProof/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
            f'{anchor_xml}</w:r>'
        )

    # ── Section properties ──
    SECT0 = (
        '<w:sectPr>'
        '<w:headerReference w:type="default" r:id="rId12"/>'
        '<w:footerReference w:type="default" r:id="rId13"/>'
        '<w:headerReference w:type="first" r:id="rId14"/>'
        '<w:footerReference w:type="first" r:id="rId15"/>'
        '<w:type w:val="continuous"/>'
        '<w:pgSz w:w="11909" w:h="16834" w:code="9"/>'
        '<w:pgMar w:top="1080" w:right="509" w:bottom="540" w:left="576"'
        ' w:header="360" w:footer="225" w:gutter="0"/>'
        '<w:cols w:num="2" w:space="372"/>'
        '<w:titlePg/><w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )
    SECT1 = (
        '<w:sectPr>'
        '<w:headerReference w:type="default" r:id="rId27"/>'
        '<w:footerReference w:type="default" r:id="rId28"/>'
        '<w:headerReference w:type="first" r:id="rId29"/>'
        '<w:footerReference w:type="first" r:id="rId30"/>'
        '<w:type w:val="continuous"/>'
        '<w:pgSz w:w="11909" w:h="16834" w:code="9"/>'
        '<w:pgMar w:top="1080" w:right="509" w:bottom="540" w:left="576"'
        ' w:header="360" w:footer="225" w:gutter="0"/>'
        '<w:cols w:space="372"/><w:docGrid w:linePitch="360"/>'
        '</w:sectPr>'
    )

    # ── Assemble document XML ──
    with zipfile.ZipFile(SRC) as z:
        original_doc = z.read('word/document.xml').decode('utf-8')

    pre_m = re.search(r'^(.*?)(?=<w:tr[ >])', original_doc, re.DOTALL)
    pre   = pre_m.group(1)

    post = (
        '</w:tbl>'
        '<w:p>'
        f'<w:pPr><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>{SECT0}</w:pPr>'
        f'{image_runs}'
        '</w:p>'
        '<w:p>'
        f'<w:pPr>{SECT1}</w:pPr>'
        '</w:p>'
        '</w:body></w:document>'
    )

    new_doc = pre + '\n'.join(rows) + post

    # ── Read template, apply all changes, write output ──
    with zipfile.ZipFile(SRC) as zin:
        file_data = {name: zin.read(name) for name in zin.namelist()}

    file_data['word/document.xml'] = new_doc.encode('utf-8')

    # Update header dates
    for hdr in ('word/header2.xml', 'word/header3.xml'):
        if hdr in file_data:
            file_data[hdr] = update_header_date(
                file_data[hdr].decode('utf-8'), monday).encode('utf-8')

    # Add flyer image files
    IMG_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
    rels_xml = file_data['word/_rels/document.xml.rels'].decode('utf-8')
    for rId, data in image_store.items():
        fn, wp, hp = image_meta[rId]
        file_data[f'word/media/{fn}'] = data
        rel = f'<Relationship Id="{rId}" Type="{IMG_REL}" Target="media/{fn}"/>'
        rels_xml = rels_xml.replace('</Relationships>', rel + '</Relationships>')
    file_data['word/_rels/document.xml.rels'] = rels_xml.encode('utf-8')

    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in file_data.items():
            zout.writestr(name, data)

    print(f'\nDone! {len(upcoming)} events, '
          f'{len(changes["new"])} new meetings, '
          f'{len(changes["changed"])} changed, '
          f'{len(image_store)} flyer images.')
    print(f'Written to {dst_name}')


if __name__ == '__main__':
    main()
