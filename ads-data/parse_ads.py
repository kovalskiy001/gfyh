"""Parse saved Facebook Ad Library pages (Chrome 'Webpage, Complete') into CSV/JSON."""
import sys, re, json, csv, glob, os
from urllib.parse import urlparse, parse_qs, unquote
from bs4 import BeautifulSoup, NavigableString

ID_RE = re.compile(r'(?:ID|Ідентифікатор)\s*[Бб]ібліотеки|ID библиотеки|Library ID')
MONTHS = {'січ':1,'лют':2,'бер':3,'кві':4,'тра':5,'чер':6,'лип':7,'сер':8,'вер':9,'жов':10,'лис':11,'гру':12,
          'янв':1,'фев':2,'мар':3,'апр':4,'мая':5,'май':5,'июн':6,'июл':7,'авг':8,'сен':9,'окт':10,'ноя':11,'дек':12}
DATE_RE = re.compile(r'(\d{1,2})\s+([^\s\d]+)\s+(\d{4})')

def text_of(el):
    return ' '.join(s.strip() for s in el.find_all(string=True) if s.strip() and s.strip() != '​')

def body_text(el):
    out = []
    for c in el.descendants:
        if isinstance(c, NavigableString):
            out.append(str(c))
        elif c.name == 'br':
            out.append('\n')
    return re.sub(r'\n{3,}', '\n\n', ''.join(out)).strip()

def parse_date(s):
    m = DATE_RE.search(s)
    if not m: return ''
    mon = MONTHS.get(m.group(2).lower()[:3])
    return f'{m.group(3)}-{mon:02d}-{int(m.group(1)):02d}' if mon else ''

def unwrap_link(href):
    if 'l.facebook.com/l.php' in href:
        u = parse_qs(urlparse(href).query).get('u', [''])[0]
        return unquote(u)
    return href

def clean_url(u):
    p = urlparse(u)
    q = [kv for kv in p.query.split('&') if kv and not kv.startswith('fbclid=')]
    return p._replace(query='&'.join(q)).geturl()

def dest_type(u):
    host = urlparse(u).netloc.lower()
    if not u: return 'no link (page like / message)'
    if host == 'fb.me': return 'fb.me (likely lead form)'
    if 'instagram.com' in host: return 'instagram'
    return 'website'

def parse_page(path):
    raw = open(path, encoding='utf-8', errors='replace').read()
    saved = re.search(r'saved from url=\(\d+\)(\S+)', raw)
    page_id = re.search(r'view_all_page_id=(\d+)', saved.group(1)).group(1) if saved else ''
    soup = BeautifulSoup(raw, 'lxml')
    for t in soup(['script', 'style']): t.decompose()
    full = soup.get_text('\n')
    rc = re.search(r'Результат[иы]:\s*~?\s*([\d\s]+)', full)
    results = int(rc.group(1).split()[0]) if rc else None
    ads = []
    for rank, node in enumerate(soup.find_all(string=ID_RE), 1):
        card = node.parent
        while card.parent is not None and len(card.parent.find_all(string=ID_RE)) == 1:
            card = card.parent
        lib_id = re.search(r'(\d{6,})', node).group(1)
        # header = everything before the <hr> separating meta from creative
        hr = card.find('hr')
        header_txt = ''
        for el in card.descendants:
            if el is hr: break
            if isinstance(el, NavigableString) and el.strip() and el.strip() != '​':
                header_txt += el.strip() + '\n'
        status = header_txt.split('\n', 1)[0]
        date_line = next((l for l in header_txt.split('\n') if DATE_RE.search(l)), '')
        runtime = re.search(r'(?:час діяльності|время активности)[^:]*:\s*(.+)$', date_line)
        n_ads = re.search(r'(\d+)\s+(?:оголошен|объявлен|ads)', header_txt)
        multi = bool(re.search(r'кілька версій|несколько версий|multiple versions', header_txt))
        # creative
        adv_a = card.find('a', href=re.compile(r'^https://www\.facebook\.com/(?!ads/)'))
        bodies = card.find_all('div', style=re.compile(r'white-space:\s*pre-wrap'))
        links = []
        body_links = []
        for a in card.find_all('a', href=re.compile(r'l\.facebook\.com/l\.php|^https?://(?!www\.facebook\.com)')):
            if any(b in a.parents for b in bodies):
                body_links.append(clean_url(unwrap_link(a['href'])))
                continue
            dest = clean_url(unwrap_link(a['href']))
            clamps = a.find_all('div', style=re.compile(r'line-clamp'))
            caps = [text_of(c) for c in clamps if 'line-height: 12px' in c.get('style', '')]
            tits = [text_of(c) for c in clamps if 'line-height: 14px' in c.get('style', '')]
            cta_box = a.find('div', style=re.compile(r'max-width:\s*120px'))
            links.append({'url': dest,
                          'caption': next((c for c in caps if c), ''),
                          'headline': tits[0] if tits else '',
                          'description': tits[1] if len(tits) > 1 else '',
                          'cta': text_of(cta_box) if cta_box else ''})
        # de-dup links (same URL appears in image + text block)
        seen, uniq = set(), []
        for l in links:
            key = (l['url'], l['headline'])
            if key in seen: continue
            seen.add(key); uniq.append(l)
        videos = card.find_all('video')
        durs = re.findall(r'0:00\s*/\s*(\d+:\d{2})', text_of(card))
        imgs = [i['src'] for i in card.find_all('img') if '_n.' in i.get('src', '')]
        if len(uniq) > 1 and len({l['headline'] for l in uniq}) > 1:
            media = 'carousel'
        elif videos:
            media = 'video'
        elif imgs[1:]:
            media = 'image'
        else:
            media = 'text/other'
        primary = uniq[0] if uniq else {'url': body_links[0] if body_links else '', 'caption': '', 'headline': '', 'description': '', 'cta': ''}
        cta_any = primary['cta'] or next((text_of(b) for b in card.find_all('div', style=re.compile(r'max-width:\s*120px')) if text_of(b)), '')
        ads.append({
            'page_id': page_id,
            'advertiser': text_of(adv_a) if adv_a else '',
            'advertiser_fb': adv_a['href'].split('?')[0] if adv_a else '',
            'advertiser_slug': adv_a['href'].split('?')[0].rstrip('/').rsplit('/', 1)[-1] if adv_a else '',
            'rank_on_page': rank,
            'library_id': lib_id,
            'status': status,
            'start_date': parse_date(date_line),
            'runtime_note': runtime.group(1) if runtime else '',
            'multiple_versions': multi,
            'ads_using_creative': int(n_ads.group(1)) if n_ads else 1,
            'media_type': media,
            'video_duration': durs[0] if durs else '',
            'body_text': body_text(bodies[0]) if bodies else '',
            'destination_url': primary['url'],
            'destination_domain': urlparse(primary['url']).netloc.lower().removeprefix('www.'),
            'destination_type': dest_type(primary['url']),
            'link_caption': primary['caption'],
            'headline': primary['headline'],
            'description': primary['description'],
            'cta': cta_any,
            'ad_library_url': f'https://www.facebook.com/ads/library/?id={lib_id}',
            'local_images': '|'.join(imgs[1:]),
            'source_file': os.path.basename(path),
        })
    return page_id, results, ads

if __name__ == '__main__':
    src_dir, out_dir = sys.argv[1], sys.argv[2]
    all_ads, summary = [], []
    for path in sorted(glob.glob(os.path.join(src_dir, 'p*.html')), key=lambda p: int(re.search(r'p(\d+)', os.path.basename(p)).group(1))):
        pid, res, ads = parse_page(path)
        all_ads += ads
        summary.append({'file': os.path.basename(path), 'page_id': pid,
                        'advertiser': ads[0]['advertiser'] if ads else '',
                        'advertiser_fb': ads[0]['advertiser_fb'] if ads else '',
                        'results_shown_on_page': res, 'ads_captured': len(ads),
                        'ads_incl_versions': sum(a['ads_using_creative'] for a in ads)})
    with open(os.path.join(out_dir, 'ads.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.DictWriter(f, fieldnames=list(all_ads[0].keys()))
        w.writeheader(); w.writerows(all_ads)
    json.dump(all_ads, open(os.path.join(out_dir, 'ads.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    json.dump(summary, open(os.path.join(out_dir, 'pages_summary.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    for s in summary: print(s)
