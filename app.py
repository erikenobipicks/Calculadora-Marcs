import base64, hashlib, hmac, secrets, os, json, time, unicodedata, math
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, jsonify, send_file, g, has_request_context)
from datetime import datetime, timedelta
from functools import wraps
from urllib.parse import urlencode, quote as urllib_quote
from urllib import error as urllib_error
from urllib import request as urllib_request
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, HRFlowable)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os as _os
_FONT_DIR = _os.path.join(_os.path.dirname(__file__), 'static', 'fonts')
pdfmetrics.registerFont(TTFont('DejaVu', _os.path.join(_FONT_DIR, 'DejaVuSans.ttf')))
pdfmetrics.registerFont(TTFont('DejaVu-Bold', _os.path.join(_FONT_DIR, 'DejaVuSans-Bold.ttf')))
import io

app = Flask(__name__)
# SECRET_KEY és obligatori. Abans hi havia un fallback a
# secrets.token_hex(32), cosa que feia que cada restart generés un secret
# nou i descarregués totes les sessions existents. Ara exigim que estigui
# a l'env i fallem a l'arrencada si falta. És preferible un crash
# explícit a la fase de deploy que un funcionament inestable.
_SECRET_KEY = (os.environ.get('SECRET_KEY') or '').strip()
if not _SECRET_KEY:
    raise RuntimeError(
        "SECRET_KEY env var is required. Set it on the deploy "
        "(Railway → Variables → SECRET_KEY) with a long random string, "
        "e.g. `python3 -c 'import secrets; print(secrets.token_urlsafe(64))'`."
    )
app.secret_key = _SECRET_KEY

# Railway (i altres proxies) envien HTTPS al client però HTTP cap al Flask.
# ProxyFix fa que Flask respecti X-Forwarded-Proto/Host per detectar HTTPS.
# Imprescindible si SESSION_COOKIE_SECURE=True, sinó les cookies no es creen.
from werkzeug.middleware.proxy_fix import ProxyFix
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

# Flags de cookie bàsics, sempre. Abans només s'activaven quan hi havia
# SESSION_COOKIE_DOMAIN, deixant configuracions sense subdomini amb
# cookies sense Secure. ProxyFix (vegeu dalt) ja garanteix que Flask
# detecta HTTPS correctament darrere del proxy.
app.config['SESSION_COOKIE_SECURE'] = True
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'
app.config['SESSION_COOKIE_HTTPONLY'] = True
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(days=30)

# Cookie compartida amb reusrevela.cat per SSO entre subdominis
# (.reusrevela.cat permet compartir entre reusrevela.cat i calculadora.reusrevela.cat)
_COOKIE_DOMAIN = os.environ.get('SESSION_COOKIE_DOMAIN', '').strip() or None
if _COOKIE_DOMAIN:
    app.config['SESSION_COOKIE_DOMAIN'] = _COOKIE_DOMAIN

MOLDURA_IMAGE_EXTS = ('jpg', 'jpeg', 'png', 'webp', 'gif')
MOLDURA_COLOR_FILTERS = [
    ('daurat', 'Daurades'),
    ('plata', 'Plata'),
    ('negre', 'Negres'),
    ('blanc', 'Blanques'),
    ('marro', 'Marrons'),
    ('fusta', 'Fusta / natural'),
    ('gris', 'Grises'),
    ('blau', 'Blaves'),
    ('verd', 'Verdes'),
    ('vermell', 'Vermelles'),
]
MOLDURA_COLOR_KEYWORDS = {
    'daurat': ('daurat', 'daurada', 'dorat', 'dorada', 'or ', ' oro', 'oro ', ' orro', 'or vell'),
    'plata': ('plata', 'argent', 'silver'),
    'negre': ('negre', 'negra', 'negro', 'mate'),
    'blanc': ('blanc', 'blanca', 'blanco', 'blanca'),
    'marro': ('marron', 'marró', 'caoba', 'wengé', 'wenge', 'noguera'),
    'fusta': ('fusta', 'natural', 'pi', 'pino', 'roure', 'roble', 'bambu', 'canya', 'fresno'),
    'gris': ('gris', 'antracita'),
    'blau': ('blau', 'azul', 'marino'),
    'verd': ('verd', 'verde', 'oliva'),
    'vermell': ('vermell', 'rojo', 'granate', 'burdeos'),
}
MOLDURA_COLOR_KEYWORDS['marro'] = ('marro', 'marron', 'caoba', 'wenge', 'noguera')
MOLDURA_GRUIX_FILTERS = [
    ('fina', 'Fines fins a 2 cm'),
    ('mitjana', 'Mitjanes de 2 a 4 cm'),
    ('gruixuda', 'Gruixudes de 4 a 6 cm'),
    ('extra', 'Extra de mes de 6 cm'),
]

COMMERCIAL_MARGIN_DEFAULTS = {
    'general': 60.0,
    'frames': 60.0,
    'canvas': 60.0,
    'prints': 60.0,
    'foam': 60.0,
    'laminate_foam': 60.0,
    'fine_art': 60.0,
    'albums': 60.0,
}

DEFAULT_BRAND_COLOR = '#1A6B45'
DEFAULT_BRAND_SECONDARY_COLOR = '#C8873A'
DEFAULT_BRAND_MENU_COLOR = DEFAULT_BRAND_COLOR


def _normalize_hex_color(value, default=DEFAULT_BRAND_COLOR):
    text = str(value or '').strip()
    if not text:
        return default
    if not text.startswith('#'):
        text = '#' + text
    if len(text) != 7:
        return default
    try:
        int(text[1:], 16)
    except ValueError:
        return default
    return text.upper()


def _mix_with_white(hex_color, ratio=0.88):
    base = _normalize_hex_color(hex_color)
    ratio = min(max(float(ratio), 0.0), 1.0)
    r = int(base[1:3], 16)
    g = int(base[3:5], 16)
    b = int(base[5:7], 16)
    nr = int(r + (255 - r) * ratio)
    ng = int(g + (255 - g) * ratio)
    nb = int(b + (255 - b) * ratio)
    return f'#{nr:02X}{ng:02X}{nb:02X}'


def _mix_hex(hex_color, target_hex, ratio=0.5):
    base = _normalize_hex_color(hex_color)
    target = _normalize_hex_color(target_hex, '#FFFFFF')
    ratio = min(max(float(ratio), 0.0), 1.0)
    r = int(base[1:3], 16)
    g = int(base[3:5], 16)
    b = int(base[5:7], 16)
    tr = int(target[1:3], 16)
    tg = int(target[3:5], 16)
    tb = int(target[5:7], 16)
    nr = int(r + (tr - r) * ratio)
    ng = int(g + (tg - g) * ratio)
    nb = int(b + (tb - b) * ratio)
    return f'#{nr:02X}{ng:02X}{nb:02X}'


def _hex_luminance(hex_color):
    base = _normalize_hex_color(hex_color)
    r = int(base[1:3], 16) / 255.0
    g = int(base[3:5], 16) / 255.0
    b = int(base[5:7], 16) / 255.0
    return (0.2126 * r) + (0.7152 * g) + (0.0722 * b)


def _contrast_text_color(hex_color):
    return '#1C1B18' if _hex_luminance(hex_color) >= 0.62 else '#FFFFFF'


def _current_brand_palette():
    brand_color = DEFAULT_BRAND_COLOR
    secondary_color = DEFAULT_BRAND_SECONDARY_COLOR
    menu_color = DEFAULT_BRAND_MENU_COLOR
    if has_request_context():
        brand_color = _normalize_hex_color(session.get('brand_color', DEFAULT_BRAND_COLOR))
        secondary_color = _normalize_hex_color(
            session.get('brand_color_secondary', DEFAULT_BRAND_SECONDARY_COLOR),
            DEFAULT_BRAND_SECONDARY_COLOR,
        )
        menu_color = _normalize_hex_color(
            session.get('brand_color_menu', brand_color or DEFAULT_BRAND_MENU_COLOR),
            brand_color or DEFAULT_BRAND_MENU_COLOR,
        )
    nav_text_color = _contrast_text_color(menu_color)
    nav_muted_color = (
        _mix_hex(menu_color, '#FFFFFF', 0.62)
        if nav_text_color == '#FFFFFF'
        else _mix_hex(menu_color, '#1C1B18', 0.58)
    )
    nav_pill_color = (
        'rgba(255,255,255,.14)'
        if nav_text_color == '#FFFFFF'
        else 'rgba(28,27,24,.10)'
    )
    return {
        'brand_color': brand_color,
        'brand_color_light': _mix_with_white(brand_color, 0.88),
        'brand_secondary_color': secondary_color,
        'brand_secondary_color_light': _mix_with_white(secondary_color, 0.86),
        'nav_color': menu_color,
        'nav_text_color': nav_text_color,
        'nav_muted_color': nav_muted_color,
        'nav_pill_color': nav_pill_color,
    }


@app.context_processor
def inject_brand_theme():
    return _current_brand_palette()


@app.context_processor
def inject_pendents_albara():
    if session.get('is_admin'):
        row = query('''SELECT COUNT(*) as n FROM comandes
                       WHERE observacions LIKE '%[ACCEPTAT]%'
                         AND (fd_albara IS NULL OR fd_albara='')''', one=True)
        return {'nav_pendents_albara': row['n'] if row else 0}
    return {'nav_pendents_albara': 0}

LAMINATE_ONLY_PRICES = {
    '20x30': 7.35,
    '24x30': 7.35,
    '24x36': 8.15,
    '28x35': 9.55,
    '30x30': 9.55,
    '30x40': 9.95,
    '30x45': 10.35,
    '30x50': 11.35,
    '35x50': 12.20,
    '30x60': 12.60,
    '40x50': 12.50,
    '40x60': 13.90,
    '50x50': 16.40,
    '50x60': 21.30,
    '50x70': 22.00,
    '50x75': 22.60,
    '50x80': 24.20,
    '60x60': 22.00,
    '60x70': 20.15,
    '60x80': 23.50,
    '60x90': 25.95,
    '60x100': 31.00,
    '70x100': 33.00,
    '80x100': 35.00,
    '80x120': 31.00,
    '80x150': 36.00,
    '80x180': 56.75,
    '80x200': 60.00,
    '90x100': 24.80,
    '90x120': 27.55,
    '90x150': 36.25,
    '90x180': 59.00,
    '90x200': 74.65,
    '100x100': 32.00,
    '100x150': 36.75,
    '100x200': 85.00,
}


def _safe_moldura_ref(ref):
    ref = (ref or '').strip().lower()
    safe = ''.join(ch if ch.isalnum() else '-' for ch in ref)
    return safe.strip('-')


def _to_public_photo_url(value):
    value = str(value or '').strip().replace('\\', '/')
    if not value:
        return ''
    if value.startswith(('http://', 'https://', 'data:', '/')):
        return value
    if value.startswith('static/'):
        return '/' + value
    full = os.path.join(app.root_path, value.lstrip('/').replace('/', os.sep))
    if os.path.isfile(full):
        return '/' + value.lstrip('/')
    return ''


def _find_local_moldura_photo(ref):
    fotos_dir = os.path.join(app.root_path, 'static', 'fotos')
    if not ref or not os.path.isdir(fotos_dir):
        return ''

    ref_strip = ref.strip()
    # Supplier codes use separators: "1717-206" or "4017/530" → photo is "1717206.jpg"
    ref_clean = ref_strip.replace('-', '').replace('/', '')

    candidates = []
    for stem in [ref_strip, ref_strip.lower(), ref_strip.upper(), _safe_moldura_ref(ref_strip),
                 ref_clean, ref_clean.lower()]:
        if stem and stem not in candidates:
            candidates.append(stem)

    # Handle leading-zero suffix: "1815-0171" (clean "18150171") → also try "1815171"
    # "7184/0086" (clean "71840086") → also try "7184086" (remove exactly ONE leading zero)
    if len(ref_clean) == 8 and ref_clean[:4].isdigit() and ref_clean[4] == '0':
        alt = ref_clean[:4] + ref_clean[5:]  # remove the single leading zero at position 4
        if alt and alt not in candidates:
            candidates.append(alt)

    for stem in candidates:
        for ext in MOLDURA_IMAGE_EXTS:
            path = os.path.join(fotos_dir, f'{stem}.{ext}')
            if os.path.isfile(path):
                return f'/static/fotos/{stem}.{ext}'
    return ''


def _resolve_moldura_photo(ref, foto, ref2=''):
    """Resolve photo for a moldura. Tries: stored foto URL, then local file by
    referencia (stripping separators), then local file by ref2 (supplier code)."""
    public_url = _to_public_photo_url(foto)
    if public_url:
        return public_url
    local = _find_local_moldura_photo(ref)
    if local:
        return local
    if ref2:
        local2 = _find_local_moldura_photo(ref2)
        if local2:
            return local2
    return ''


def _serialize_moldura(row):
    if not row:
        return None
    data = dict(row)
    data['foto'] = _resolve_moldura_photo(
        data.get('referencia', ''), data.get('foto', ''), ref2=data.get('ref2', ''))
    return data


def _serialize_moldures(rows):
    return [_serialize_moldura(row) for row in (rows or [])]


def _normalize_text(value):
    text = str(value or '').strip().lower()
    return ''.join(
        ch for ch in unicodedata.normalize('NFKD', text)
        if not unicodedata.combining(ch)
    )


def _matches_moldura_color(desc, color):
    if not color:
        return True
    keywords = MOLDURA_COLOR_KEYWORDS.get(color, ())
    text = ' ' + _normalize_text(desc) + ' '
    return any(keyword in text for keyword in keywords)


def _matches_moldura_gruix(gruix, bucket):
    value = float(gruix or 0)
    if not bucket:
        return True
    if bucket == 'fina':
        return value <= 2
    if bucket == 'mitjana':
        return 2 < value <= 4
    if bucket == 'gruixuda':
        return 4 < value <= 6
    if bucket == 'extra':
        return value > 6
    return True


def _query_moldures(q='', proveidor='', cols='*'):
    sql = f"SELECT {cols} FROM moldures"
    args = []
    where = []

    if q:
        like = f'%{q.lower()}%'
        where.append("""(
            LOWER(referencia) LIKE ?
            OR LOWER(COALESCE(descripcio, '')) LIKE ?
            OR LOWER(COALESCE(proveidor, '')) LIKE ?
        )""")
        args.extend([like, like, like])

    if proveidor:
        where.append("proveidor=?")
        args.append(proveidor)

    if where:
        sql += " WHERE " + " AND ".join(where)

    sql += " ORDER BY referencia"
    return query(sql, args)


def _moldura_photo_path(ref, foto):
    public_url = _resolve_moldura_photo(ref, foto)
    if not public_url.startswith('/static/'):
        return ''
    rel = public_url[len('/static/'):]
    full = os.path.join(app.root_path, 'static', rel.replace('/', os.sep))
    return full if os.path.isfile(full) else ''


def _save_moldura_photo(upload, referencia):
    if not upload or not getattr(upload, 'filename', ''):
        return ''

    filename = upload.filename.rsplit('/', 1)[-1].rsplit('\\', 1)[-1]
    if '.' not in filename:
        raise ValueError('La imatge ha de tenir extensio (jpg, png, webp o gif).')

    ext = filename.rsplit('.', 1)[-1].lower()
    if ext not in MOLDURA_IMAGE_EXTS:
        raise ValueError('Format d\'imatge no valid. Usa JPG, PNG, WEBP o GIF.')

    stem = _safe_moldura_ref(referencia) or 'moldura'
    fotos_dir = os.path.join(app.root_path, 'static', 'fotos')
    os.makedirs(fotos_dir, exist_ok=True)

    for old_ext in MOLDURA_IMAGE_EXTS:
        old_path = os.path.join(fotos_dir, f'{stem}.{old_ext}')
        if os.path.exists(old_path):
            try:
                os.remove(old_path)
            except OSError:
                pass

    upload.save(os.path.join(fotos_dir, f'{stem}.{ext}'))
    return f'/static/fotos/{stem}.{ext}'

def _iter_conflict_check_files(base):
    """Yield text/source files where merge markers should never appear."""
    exts = {'.py', '.html', '.css', '.js', '.md', '.txt', '.toml', '.yml', '.yaml'}
    skip_dirs = {'.git', '__pycache__', '.pytest_cache', '.mypy_cache', '.venv', 'venv'}

    for root, dirs, files in os.walk(base):
        dirs[:] = [d for d in dirs if d not in skip_dirs]
        for name in files:
            if os.path.splitext(name)[1].lower() in exts:
                yield os.path.join(root, name)


def _assert_no_conflict_markers():
    """Fail fast if unresolved merge conflict markers are present in source files."""
    base = os.path.dirname(__file__)
    found = []

    for path in _iter_conflict_check_files(base):
        rel = os.path.relpath(path, base)
        try:
            with open(path, 'r', encoding='utf-8') as f:
                for n, line in enumerate(f, start=1):
                    stripped = line.strip()
                    if line.startswith('<<<<<<< ') or line.startswith('>>>>>>> '):
                        found.append(f"{rel}:{n}: {stripped}")
                    elif line.startswith('=======') and stripped == '=======':
                        found.append(f"{rel}:{n}: =======")
        except (FileNotFoundError, UnicodeDecodeError):
            continue

    if found:
        details = "\n".join(found[:20])
        more = f"\n... and {len(found)-20} more" if len(found) > 20 else ""
        raise RuntimeError(f"Merge conflict markers detected:\n{details}{more}")

_assert_no_conflict_markers()

# â"€â"€ DB layer: PostgreSQL (production) or SQLite (local) â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
DATABASE_URL = os.environ.get('DATABASE_URL')

if DATABASE_URL:
    import psycopg2
    import psycopg2.extras
    # Railway uses postgres:// but psycopg2 needs postgresql://
    if DATABASE_URL.startswith('postgres://'):
        DATABASE_URL = DATABASE_URL.replace('postgres://', 'postgresql://', 1)
    USE_PG = True
else:
    import sqlite3
    USE_PG = False
    DB = os.path.join(os.path.dirname(__file__), 'objectiu.db')

def get_db():
    if 'db' not in g:
        if USE_PG:
            g.db = psycopg2.connect(
                DATABASE_URL,
                cursor_factory=psycopg2.extras.RealDictCursor,
                connect_timeout=10,
                options='-c statement_timeout=30000',
            )
        else:
            g.db = sqlite3.connect(DB)
            g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop('db', None)
    if db:
        try: db.close()
        except: pass

def _fix_sql(sql):
    """Convert SQLite ? placeholders to PostgreSQL %s, and escape % in LIKE"""
    if USE_PG:
        import re as _re
        # First escape existing % that are part of LIKE patterns (not already %s or %%)
        # Replace LIKE '...%...' patterns to use %%
        sql = _re.sub(r"LIKE '([^']*)'", lambda m: "LIKE '" + m.group(1).replace('%','%%') + "'", sql)
        sql = sql.replace('?', '%s')
    return sql

def query(sql, args=(), one=False):
    db = get_db()
    sql = _fix_sql(sql)
    if USE_PG:
        # Rollback any failed transaction before executing
        if db.status != 1:  # 1 = STATUS_READY
            try: db.rollback()
            except: pass
        cur = db.cursor()
        cur.execute(sql, list(args))
        r = cur.fetchall()
        # Convert to list of dicts
        r = [dict(row) for row in r]
        return (r[0] if r else None) if one else r
    else:
        cur = db.execute(sql, args)
        r = cur.fetchall()
        return (r[0] if r else None) if one else r

def _pg_fix_insert(sql):
    """Convert SQLite INSERT OR IGNORE/REPLACE to PostgreSQL syntax"""
    sql = sql.replace('INSERT OR IGNORE INTO', 'INSERT INTO').replace(
          'INSERT OR REPLACE INTO', 'INSERT INTO')
    # Add ON CONFLICT DO NOTHING for IGNORE (will be added below per-call)
    return sql

def execute(sql, args=()):
    db = get_db()
    if USE_PG:
        sql2 = _fix_sql(sql)
        is_ignore  = 'INSERT OR IGNORE'  in sql.upper()
        is_replace = 'INSERT OR REPLACE' in sql.upper()
        sql2 = _pg_fix_insert(sql2)
        if 'INSERT' in sql2.upper():
            sql2 = sql2.rstrip().rstrip(';')
            if is_ignore:
                sql2 += ' ON CONFLICT DO NOTHING'
            # Only add RETURNING id for tables that have serial id
            has_id = any(t in sql2.upper() for t in ['INTO USUARIS','INTO COMANDES','INTO MOLDURES',
                         'INTO VIDRES','INTO PASSPARTOUT','INTO ENCOLAT_PRO','INTO IMPRESSIO'])
            if has_id and 'RETURNING' not in sql2.upper():
                sql2 += ' RETURNING id'
            cur = db.cursor()
            cur.execute(sql2, list(args))
            db.commit()
            if has_id:
                row = cur.fetchone()
                return row['id'] if row else None
            return None
        else:
            cur = db.cursor()
            cur.execute(sql2, list(args))
            db.commit()
            return None
    else:
        cur = db.execute(sql, args)
        db.commit()
        return cur.lastrowid

from werkzeug.security import generate_password_hash, check_password_hash


def hash_pw(pw):
    """Hash d'una contrasenya amb pbkdf2-sha256 + sal aleatòria (werkzeug).
    Els hashes resultants comencen per 'pbkdf2:sha256:...' i contenen '$'
    separant els camps, cosa que els distingeix dels hashes legacy (64
    chars hex).
    """
    return generate_password_hash(pw or '', method='pbkdf2:sha256', salt_length=16)


def _hash_pw_legacy(pw):
    """SHA-256 sense sal. Només per verificar hashes antics que encara no
    han migrat. Cap ruta hauria d'escriure res nou amb això."""
    return hashlib.sha256((pw or '').encode()).hexdigest()


def _is_legacy_hash(stored):
    """Un hash legacy és un string de 64 caràcters hex (SHA-256 sense sal
    emmagatzemat com a hexdigest). Els hashes werkzeug sempre contenen
    el separador '$'."""
    if not isinstance(stored, str) or not stored:
        return False
    return '$' not in stored


def verify_pw(stored, provided):
    """True si la contrasenya 'provided' és correcta per al hash 'stored',
    independentment de si és un hash legacy (SHA-256) o un de werkzeug."""
    if not stored or provided is None:
        return False
    if _is_legacy_hash(stored):
        return hmac.compare_digest(stored, _hash_pw_legacy(provided))
    try:
        return check_password_hash(stored, provided)
    except Exception:
        return False


def maybe_upgrade_hash(user_id, stored, provided):
    """Si l'usuari acaba de passar un login amb un hash legacy, aprofitem
    per rehashejar amb el format nou i desar-ho a la BD. Migració
    progressiva: no cal cap operació massiva, simplement a cada login
    correcte es moderniza un usuari més.

    No fa fallar el login si l'UPDATE falla — només registra warning.
    """
    if not _is_legacy_hash(stored) or not user_id:
        return
    try:
        execute('UPDATE usuaris SET password=? WHERE id=?', [hash_pw(provided), user_id])
    except Exception as exc:
        app.logger.warning('password_rehash_failed user_id=%s detail=%s', user_id, exc)


def _is_admin_session():
    return bool(session.get('is_admin'))


def _row_get(row, key, default=None):
    if row is None:
        return default
    try:
        return row[key]
    except Exception:
        pass
    try:
        return row.get(key, default)
    except Exception:
        return default


def _user_access_status(user):
    status = str(_row_get(user, 'access_status', '') or '').strip().lower()
    return status or 'active'


def _get_marge_value(u):
    """Marge professional efectiu. Prioritat: marge_pro_pct > marge > 60.
    Fa servir 'is not None' en lloc de 'or' perquè un valor explícit de 0
    és legítim (admin, col·laborador intern) i no hauria de caure al
    fallback de 60."""
    val = _row_get(u, 'marge_pro_pct')
    if val is not None:
        return float(val)
    val = _row_get(u, 'marge')
    if val is not None:
        return float(val)
    return 60.0


def _get_marge_impressio_value(u):
    """Marge d'impressió. Prioritat: marge_impressio_pro_pct > marge_impressio > 0.
    Mateix raonament que _get_marge_value — 0 és legítim."""
    val = _row_get(u, 'marge_impressio_pro_pct')
    if val is not None:
        return float(val)
    val = _row_get(u, 'marge_impressio')
    if val is not None:
        return float(val)
    return 0.0


def _user_is_allowed(user):
    return bool(user) and (bool(_row_get(user, 'is_admin', 0)) or _user_access_status(user) == 'active')


def _clean_profile_type(value):
    allowed = {'professional', 'studio', 'gallery', 'association'}
    value = (value or '').strip().lower()
    return value if value in allowed else 'professional'


def _bridge_api_token():
    return os.environ.get('PUBLIC_BRIDGE_TOKEN', '').strip()


def _bridge_signing_secret():
    return os.environ.get('BRIDGE_LOGIN_SECRET', '').strip() or _bridge_api_token()


def _main_site_url():
    return os.environ.get('MAIN_SITE_URL', 'https://reusrevela.cat').strip().rstrip('/')


def _safe_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return float(default)


def get_config_value(clau, default=None):
    """Read a single value from the config key-value table."""
    r = query("SELECT valor FROM config WHERE clau=?", [clau], one=True)
    return r['valor'] if r else default


def calcular_pvd(preu_cost, categoria):
    """cost × (1 + marge_admin_pct/100). categoria: 'moldures'|'vidres'|'passpartu'|'encolat'"""
    if preu_cost is None:
        return None
    marge = float(get_config_value(f'marge_admin_{categoria}_pct', '60'))
    return round(preu_cost * (1 + marge / 100), 4)


def calcular_preu_marc(amplada, alcada, gruix, preu_cost, merma_pct=10.0, minim_cm=100.0):
    """Calcula cost i PVD d'un marc a partir de preu_cost per cm lineal.
    Retorna dict {'cost': float, 'pvd': float} o None si dades insuficients."""
    if not all([amplada, alcada, preu_cost]):
        return None
    perimetre = 2 * (amplada + alcada)
    long_bruta = (perimetre + (gruix or 0) * 8) * (1 + (merma_pct or 10.0) / 100)
    long_final = max(long_bruta, minim_cm or 100.0)
    cost = round(long_final * preu_cost / 100, 4)
    pvd = calcular_pvd(cost, 'moldures')
    return {'cost': cost, 'pvd': pvd}


def _normalize_commercial_margins(raw=None, frame_margin=None, print_margin=None):
    data = raw if isinstance(raw, dict) else {}
    general_margin = max(_safe_float(data.get('general'), frame_margin if frame_margin is not None else COMMERCIAL_MARGIN_DEFAULTS['general']), 0.0)
    frame_value = max(_safe_float(data.get('frames'), frame_margin if frame_margin is not None else general_margin), 0.0)
    print_seed = print_margin if print_margin is not None else frame_value
    print_value = max(_safe_float(data.get('prints'), print_seed), 0.0)
    normalized = {
        'general': general_margin,
        'frames': frame_value,
        'canvas': max(_safe_float(data.get('canvas'), frame_value), 0.0),
        'prints': print_value,
        'foam': general_margin,
        'laminate_foam': general_margin,
        'fine_art': max(_safe_float(data.get('fine_art'), frame_value), 0.0),
        'albums': max(_safe_float(data.get('albums'), frame_value), 0.0),
    }
    return normalized


def _load_user_commercial_margins(user_row):
    payload = {}
    raw = user_row['margins_json'] if user_row and 'margins_json' in user_row.keys() else ''
    if raw:
        try:
            payload = json.loads(raw)
        except (TypeError, ValueError):
            payload = {}
    frame_margin = user_row['marge'] if user_row and user_row['marge'] is not None else COMMERCIAL_MARGIN_DEFAULTS['frames']
    print_margin = None
    if user_row and 'marge_impressio' in user_row.keys() and user_row['marge_impressio'] is not None:
        print_margin = user_row['marge_impressio']
    return _normalize_commercial_margins(payload, frame_margin=frame_margin, print_margin=print_margin)


def _format_margin_for_view(value):
    value = _safe_float(value, 0.0)
    return int(value) if float(value).is_integer() else round(value, 2)


def _sync_private_commercial_settings(frame_margin, print_margin, margins=None):
    api_token = _bridge_api_token()
    base = _main_site_url()
    if not api_token or not base:
        return {'attempted': False, 'reason': 'missing_config'}

    normalized = _normalize_commercial_margins(margins or {}, frame_margin=frame_margin, print_margin=print_margin)
    payload = dict(normalized)
    req = urllib_request.Request(
        f'{base}/api/private/commercial-settings-sync',
        data=json.dumps(payload).encode('utf-8'),
        headers={
            'Content-Type': 'application/json',
            'X-Bridge-Token': api_token,
            # Cloudflare davant de reusrevela.cat pot bloquejar el UA per
            # defecte de Python-urllib. Enviar un UA propi permet whitelist
            # i fa visibles els hits de sortida.
            'User-Agent': 'calculadora-marcs-bridge/1.0 (+https://calculadora.reusrevela.cat)',
        },
        method='POST',
    )
    try:
        with urllib_request.urlopen(req, timeout=12) as resp:
            body = resp.read().decode('utf-8')
            return {'attempted': True, 'ok': True, 'response': json.loads(body or '{}')}
    except urllib_error.HTTPError as exc:
        detail = exc.read().decode('utf-8', errors='ignore')
        return {'attempted': True, 'ok': False, 'status': exc.code, 'detail': detail}
    except (urllib_error.URLError, TimeoutError, ValueError) as exc:
        return {'attempted': True, 'ok': False, 'detail': str(exc)}


def _urlsafe_b64encode(raw):
    return base64.urlsafe_b64encode(raw).decode().rstrip('=')


def _urlsafe_b64decode(text):
    text = str(text or '').strip()
    if not text:
        return b''
    padding = '=' * (-len(text) % 4)
    return base64.urlsafe_b64decode((text + padding).encode())


def _safe_next_path(value, default='/'):
    text = str(value or '').strip()
    if not text.startswith('/') or text.startswith('//') or '://' in text:
        return default
    return text


def _build_bridge_token(data):
    secret = _bridge_signing_secret()
    if not secret:
        raise RuntimeError('Missing bridge signing secret.')
    payload = _urlsafe_b64encode(json.dumps(data, separators=(',', ':'), ensure_ascii=True).encode())
    signature = _urlsafe_b64encode(hmac.new(secret.encode(), payload.encode(), hashlib.sha256).digest())
    return f'{payload}.{signature}'


def _read_bridge_token(token, max_age=120):
    secret = _bridge_signing_secret()
    if not secret or '.' not in str(token or ''):
        return None

    payload, signature = token.split('.', 1)
    expected = _urlsafe_b64encode(hmac.new(secret.encode(), payload.encode(), hashlib.sha256).digest())
    if not hmac.compare_digest(signature, expected):
        return None

    try:
        data = json.loads(_urlsafe_b64decode(payload).decode())
    except Exception:
        return None

    issued_at = int(data.get('iat') or 0)
    if not issued_at or (time.time() - issued_at) > max_age:
        return None
    return data


def _load_empresa_nom_for_session(user):
    try:
        nom_emp = user['nom_empresa'] if user['nom_empresa'] else ''
        if not nom_emp:
            r_cfg = query("SELECT valor FROM config WHERE clau='empresa_nom'", one=True)
            nom_emp = r_cfg['valor'] if r_cfg and r_cfg['valor'] else ''
        return nom_emp or 'Calculadora'
    except Exception:
        return 'Calculadora'


def _build_web_return_url(source=None, lang=None):
    base = _main_site_url()
    if not base:
        return ''

    source = str(source or '').strip().lower()
    lang = (lang or 'ca').strip().lower() or 'ca'
    path = '/area-privada' if source in {'private_area', 'web_private', 'area_privada', 'web'} else '/'
    if path == '/':
        return f'{base}/?lang={lang}'
    return f'{base}{path}?lang={lang}'


def _current_web_return_url():
    lang = session.get('bridge_lang') or request.args.get('lang') or 'ca'
    base = _main_site_url()
    return f'{base}/area-privada?lang={lang}'


def _current_web_order_url():
    base = _main_site_url()
    lang = (session.get('bridge_lang') or request.args.get('lang') or 'ca').strip().lower() or 'ca'
    return f'{base}/area-privada/comanda?lang={lang}'


def _current_web_module_url(path):
    base = _main_site_url()
    lang = (session.get('bridge_lang') or request.args.get('lang') or 'ca').strip().lower() or 'ca'
    clean_path = '/' + str(path or '').strip().lstrip('/')
    return f'{base}{clean_path}?lang={lang}'


def _needs_setup(user_id):
    try:
        u2 = query('SELECT setup_done FROM usuaris WHERE id=?', [user_id], one=True)
        return bool(u2) and not bool(_row_get(u2, 'setup_done', 0))
    except Exception as exc:
        print(f'setup check error: {exc}')
        return False


def _start_user_session(user):
    # Usem _row_get per a TOTS els camps: si el schema no té una columna (p.ex.
    # brand_color en bases antigues) no volem un 500 durant el login.
    session['user_id'] = _row_get(user, 'id')
    session['username'] = _row_get(user, 'username', '')
    session['is_admin'] = bool(_row_get(user, 'is_admin', 0))
    session['nom'] = _row_get(user, 'nom', '') or ''
    session['access_status'] = _user_access_status(user)
    session['profile_type'] = _clean_profile_type(_row_get(user, 'profile_type', 'professional'))
    session['empresa_nom'] = _load_empresa_nom_for_session(user)
    session['brand_color'] = _normalize_hex_color(_row_get(user, 'brand_color', DEFAULT_BRAND_COLOR))
    session['brand_color_secondary'] = _normalize_hex_color(
        _row_get(user, 'brand_color_secondary', DEFAULT_BRAND_SECONDARY_COLOR),
        DEFAULT_BRAND_SECONDARY_COLOR,
    )
    session['brand_color_menu'] = _normalize_hex_color(
        _row_get(user, 'brand_color_menu', session.get('brand_color', DEFAULT_BRAND_MENU_COLOR)),
        session.get('brand_color', DEFAULT_BRAND_MENU_COLOR),
    )


def _can_access_comanda(row):
    return bool(row) and (_is_admin_session() or row['user_id'] == session.get('user_id'))


def _get_comanda_for_session(cid, *, fields='*'):
    row = query(f'SELECT {fields} FROM comandes WHERE id=?', [cid], one=True)
    if not _can_access_comanda(row):
        return None
    return row


def _get_comanda_by_sessio_for_session(sessio_id, *, fields='id, user_id'):
    row = query(f'SELECT {fields} FROM comandes WHERE sessio_id=? LIMIT 1', [sessio_id], one=True)
    if not _can_access_comanda(row):
        return None
    return row


_INTERMOL_REFS = [
    ('1021271','1021'),('1021298','1021'),('1520136','1520'),('1520170','1520'),('1520180','1520'),
    ('1528003','1528'),('17172011','1717'),('17172012','1717'),('17172013','1717'),('17172014','1717'),
    ('17172015','1717'),('17172016','1717'),('1717206','1717'),('1717212','1717'),('1717216','1717'),
    ('1717261','1717'),('1717263','1717'),('1717268','1717'),('1717270','1717'),('1717272','1717'),
    ('1717281','1717'),('1717295','1717'),('1717330','1717'),('1717336','1717'),('1717374','1717'),
    ('1717380','1717'),('1717540','1717'),('1717541','1717'),('1717543','1717'),('1717549','1717'),
    ('1717653','1717'),('1717673','1717'),('1717975','1717'),('1717985','1717'),('18141008','1814'),
    ('18141010','1814'),('1814640','1814'),('1815171','1815'),('1815181','1815'),('18183012','1818'),
    ('18183016','1818'),('1840003','1840'),('1840007','1840'),('18413011','1841'),('18413012','1841'),
    ('18413013','1841'),('1930536','1930'),('1930540','1930'),('1930975','1930'),('1930985','1930'),
    ('2020003','2020'),('2020007','2020'),('2127172','2127'),('2127485','2127'),('2322481','2322'),
    ('2322536','2322'),('2558640','2558'),('2816530','2816'),('2816536','2816'),('28253012','2825'),
    ('28253016','2825'),('3031003','3031'),('3031007','3031'),('3131530','3131'),('3131536','3131'),
    ('3131540','3131'),('3131653','3131'),('3131673','3131'),('3237171','3237'),('3237181','3237'),
    ('3237675','3237'),('4017380','4017'),('4017530','4017'),('4017540','4017'),('4017546','4017'),
    ('4017549','4017'),('4017638','4017'),('4017985','4017'),('4144171','4144'),('4411629','4411'),
    ('4413629','4413'),('5064614','5064'),('5067714','5067'),('5285714','5285'),('5301183','5301'),
    ('5301806','5301'),('5302667','5302'),('5302806','5302'),('5303624','5303'),('5303720','5303'),
    ('5515201','5515'),('5515203','5515'),('5515238','5515'),('5670236','5670'),('5753616','5753'),
    ('5753617','5753'),('5753718','5753'),('5757718','5757'),('61831021','6183'),('61861022','6186'),
    ('6201571','6201'),('6201581','6201'),('6202670','6202'),('64121019','6214'),('64124010','6214'),
    ('64125004','6214'),('6344082','6344'),('6412626','6412'),('6486171','6486'),('6486181','6486'),
    ('6486629','6486'),('6486727','6486'),('66121024','6612'),('66121025','6612'),('66123008','6612'),
    ('66124011','6612'),('66125006','6612'),('69921027','6992'),('69931027','6993'),('69931028','6993'),
    ('69951026','6995'),('69951028','6995'),('7017638','7017'),('7027627','7027'),('7184076','7184'),
    ('7184086','7184'),('7184602','7184'),('7184702','7184'),('7532581','7532'),('7560481','7560'),
    ('7801471','7801'),('7901480','7901'),('7901870','7901'),('7901979','7901'),('7902976','7902'),
    ('7902979','7902'),('8517374','8517'),('8517651','8517'),('8537171','8537'),('8537181','8537'),
    ('8537655','8537'),('8547603','8547'),('8547675','8547'),('8547703','8547'),('85574001','8557'),
]

_PROECO_PREUS = [
    ('PROECO1520', 7.81),  ('PROECO1824', 9.24),  ('PROECO2020', 7.81),
    ('PROECO2025', 8.80),  ('PROECO2030', 10.67), ('PROECO2323', 8.80),
    ('PROECO2430', 11.55), ('PROECO2436', 12.65), ('PROECO2835', 14.52),
    ('PROECO3030', 14.52), ('PROECO3040', 15.51), ('PROECO3045', 16.50),
    ('PROECO3050', 17.93), ('PROECO3060', 20.35), ('PROECO3550', 19.36),
    ('PROECO4050', 20.35), ('PROECO4060', 21.34), ('PROECO5050', 23.21),
    ('PROECO5060', 25.19), ('PROECO5070', 29.04), ('PROECO5075', 31.90),
    ('PROECO60100', 43.56),('PROECO6060', 29.04), ('PROECO6070', 33.88),
    ('PROECO6080', 36.74), ('PROECO6090', 40.70), ('PROECO70100', 48.40),
    ('PROECO80100', 50.38),('PROECO80120', 55.22),('PROECO80150', 61.05),
    ('PROECO80180', 87.12),('PROECO80200', 116.16),('PROECO90100', 53.24),
    ('PROECO90120', 62.92),('PROECO90150', 72.60),('PROECO90180', 92.95),
    ('PROECO90200', 125.84),
]

def _seed_proeco_preus(db, use_pg=False):
    """Inserta els preus de ProEco si no existeixen."""
    ok = 0
    if use_pg:
        cur = db.cursor()
        for ref, preu in _PROECO_PREUS:
            try:
                cur.execute(
                    "INSERT INTO proeco (referencia, preu) VALUES (%s, %s) ON CONFLICT (referencia) DO NOTHING",
                    [ref, preu]
                )
                ok += 1
            except Exception as e:
                print(f'ProEco seed PG skip {ref}: {e}')
    else:
        for ref, preu in _PROECO_PREUS:
            try:
                db.execute(
                    "INSERT OR IGNORE INTO proeco (referencia, preu) VALUES (?, ?)",
                    [ref, preu]
                )
                ok += 1
            except Exception as e:
                print(f'ProEco seed SQLite skip {ref}: {e}')
    print(f'ProEco seed: {ok}/{len(_PROECO_PREUS)} preus inserits')


def _seed_intermol_moldures(db, use_pg=False):
    """Neteja registres Intermol inserits per error (codi proveïdor com a referència primària).
    Les fotos es resolen dinàmicament via _find_local_moldura_photo: el codi intern
    "1717-206" troba automàticament "1717206.jpg" eliminant el guió/barra.
    """
    refs = [r for r, _ in _INTERMOL_REFS]
    deleted = 0

    if use_pg:
        cur = db.cursor()
        for ref in refs:
            try:
                cur.execute(
                    "DELETE FROM moldures WHERE referencia=%s "
                    "AND preu_taller=0 AND cost=0 AND COALESCE(descripcio,'')=''",
                    [ref]
                )
                deleted += cur.rowcount
            except Exception as e:
                print(f'Intermol cleanup PG {ref}: {e}')
        try:
            db.commit()
        except Exception:
            pass
    else:
        for ref in refs:
            try:
                db.execute(
                    "DELETE FROM moldures WHERE referencia=? "
                    "AND preu_taller=0 AND cost=0 AND COALESCE(descripcio,'')=''",
                    [ref]
                )
            except Exception as e:
                print(f'Intermol cleanup SQLite {ref}: {e}')
        try:
            db.commit()
        except Exception:
            pass
    print(f'Intermol cleanup: {deleted} registres erronis eliminats (fotos resoltes per referencia)')


def _fix_ref2_errors(db, use_pg=False):
    """Corregeix errors de transcripció coneguts al camp ref2 de moldures."""
    fixes = [
        # (referencia, ref2_erroni, ref2_correcte)
        ('M2314', '6183/1022', '6183/1021'),  # foto 61831021.jpg disponible
        ('M2315', '6214/1019', '6412/1019'),  # 6214 era 6412 invertit
        ('M2316', '6214/4010', '6412/4010'),
        ('M2317', '6214/5004', '6412/5004'),
    ]
    fixed = 0
    if use_pg:
        cur = db.cursor()
        for ref, wrong, correct in fixes:
            try:
                cur.execute(
                    "UPDATE moldures SET ref2=%s WHERE referencia=%s AND ref2=%s",
                    [correct, ref, wrong])
                fixed += cur.rowcount
            except Exception as e:
                print(f'ref2 fix PG {ref}: {e}')
        try:
            db.commit()
        except Exception:
            pass
    else:
        for ref, wrong, correct in fixes:
            try:
                db.execute(
                    "UPDATE moldures SET ref2=? WHERE referencia=? AND ref2=?",
                    [correct, ref, wrong])
                fixed += 1
            except Exception as e:
                print(f'ref2 fix SQLite {ref}: {e}')
        try:
            db.commit()
        except Exception:
            pass
    if fixed:
        print(f'ref2 fixes: {fixed} correccions aplicades')


def _seed_admin_if_configured(db):
    admin_user = os.environ.get('ADMIN_USERNAME', '').strip()
    admin_pass = os.environ.get('ADMIN_PASSWORD', '').strip()
    admin_name = os.environ.get('ADMIN_NAME', 'Administrador').strip() or 'Administrador'

    if not admin_user or not admin_pass:
        print('Admin bootstrap skipped: set ADMIN_USERNAME and ADMIN_PASSWORD to create the first admin.')
        return

    if USE_PG:
        cur = db.cursor()
        cur.execute('SELECT id FROM usuaris WHERE username=%s', [admin_user])
        if not cur.fetchone():
            cur.execute(
                'INSERT INTO usuaris (username,password,nom,is_admin) VALUES (%s,%s,%s,%s)',
                [admin_user, hash_pw(admin_pass), admin_name, 1]
            )
            db.commit()
            print(f"Admin creat des de variables d'entorn: usuari={admin_user}")
    else:
        cur = db.execute('SELECT id FROM usuaris WHERE username=?', [admin_user])
        if not cur.fetchone():
            db.execute(
                'INSERT INTO usuaris (username,password,nom,is_admin) VALUES (?,?,?,1)',
                [admin_user, hash_pw(admin_pass), admin_name]
            )
            db.commit()
            print(f"Admin creat des de variables d'entorn: usuari={admin_user}")

# â"€â"€ Auth decorators â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
_CANONICAL_HOST = os.environ.get('CANONICAL_HOST', '').strip().lower()


@app.before_request
def _redirect_legacy_host():
    """Redirigeix 301 qualsevol host no canònic (p.ex. calculadora.objectiufotografs.com)
    cap a CANONICAL_HOST (p.ex. calculadora.reusrevela.cat). Opt-in via env var.
    Preserva path i query string. Només actua amb GET/HEAD."""
    if not _CANONICAL_HOST:
        return
    if request.method not in ('GET', 'HEAD'):
        return
    current_host = (request.host or '').lower()
    # No redirigim si ja estem al host canònic
    if current_host == _CANONICAL_HOST:
        return
    # No redirigim localhost/127.0.0.1/railway.app internal domains
    if current_host.startswith(('localhost', '127.', '0.0.0.0')) or 'railway' in current_host:
        return
    target = f'https://{_CANONICAL_HOST}{request.full_path}'
    # full_path acaba amb '?' si no hi ha query → netegem
    if target.endswith('?'):
        target = target[:-1]
    # 302 (temporal) en lloc de 301 perquè els navegadors no el cachegin
    # permanentment i puguin reintentar si la configuració canvia.
    return redirect(target, code=302)


@app.before_request
def _sso_from_shared_cookie():
    """SSO fallback: si ve una sessió del Repo B (reusrevela.cat) via cookie
    compartida amb 'private_professional' però sense 'user_id' (sessió de A),
    carregar l'usuari automàticament perquè no hagi de tornar a loguejar-se."""
    if session.get('user_id'):
        return
    pp = session.get('private_professional') or {}
    username = (pp.get('username') or '').strip().lower()
    if not username:
        return
    try:
        user = query('SELECT * FROM usuaris WHERE username=?', [username], one=True)
    except Exception:
        return
    if not user:
        return
    try:
        if not _user_is_allowed(user):
            return
    except Exception:
        pass
    # Activar la sessió del Repo A reusant les dades de B
    try:
        _start_user_session(user)
    except Exception:
        session['user_id'] = user['id']
        session['is_admin'] = bool(user.get('is_admin', 0))
        session['nom'] = user.get('nom') or ''


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            next_path = request.full_path[:-1] if request.full_path.endswith('?') else request.full_path
            login_lang = (
                (session.get('bridge_lang') or '').strip().lower()
                or request.accept_languages.best_match(['ca', 'es', 'en'])
                or 'ca'
            )
            return redirect(url_for('login', next=_safe_next_path(next_path, '/'), source='calc', lang=login_lang))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        user = query('SELECT * FROM usuaris WHERE id=?', [session['user_id']], one=True)
        if not user or not user['is_admin']:
            flash('Accés restringit a administradors.', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated

# ── Passpartú: càlcul dinàmic de preu ────────────────────────────────────────
_SHEET_W      = 80       # ample full en cm
_SHEET_H      = 120      # alt full en cm
_SHEET_PRICE  = 7.0      # preu full en €
_COST_CM2     = 0.000729 # cost per cm²
_MULT_PETITA  = 8.0      # multiplicador fins a _LIMIT_A cm²
_MULT_MITJANA = 5.0      # multiplicador entre _LIMIT_A i _LIMIT_B cm²
_MULT_GRAN    = 3.5      # multiplicador per sobre de _LIMIT_B cm²
_LIMIT_A      = 900      # cm² (~30x30)
_LIMIT_B      = 4800     # cm² (~60x80)
_MIN_PRICE    = 5.50     # preu mínim per peça

def _peces_per_full(ew, eh):
    # Quantes peces caben en un full 80x120, provant les dues orientacions.
    n1 = math.floor(_SHEET_W / ew) * math.floor(_SHEET_H / eh)
    n2 = math.floor(_SHEET_W / eh) * math.floor(_SHEET_H / ew)
    return max(1, n1, n2)

def calcular_precio_passpartu(ew, eh):
    """Legacy: retorna PVD directament (multiplicadors ja inclouen marge).
    Mantingut per compatibilitat — nou codi hauria d'usar calcular_cost_passpartu()."""
    area = ew * eh
    n = _peces_per_full(ew, eh)
    coste_cm2  = area * _COST_CM2
    coste_hoja = _SHEET_PRICE / n
    coste_base = max(coste_cm2, coste_hoja)
    if area <= _LIMIT_A:
        mult = _MULT_PETITA
        t = (area - _LIMIT_A * 0.92) / (_LIMIT_A * 0.16)
        if 0 < t < 1:
            mult = _MULT_PETITA + t * (_MULT_MITJANA - _MULT_PETITA)
    elif area <= _LIMIT_B:
        mult = _MULT_MITJANA
        t = (area - _LIMIT_B * 0.92) / (_LIMIT_B * 0.16)
        if 0 < t < 1:
            mult = _MULT_MITJANA + t * (_MULT_GRAN - _MULT_MITJANA)
    else:
        mult = _MULT_GRAN
    return round(max(coste_base * mult, _MIN_PRICE), 2)


def _closest_passpartu_taula(amplada, alcada, prefix='1PAS'):
    """Min-contain lookup on the passpartout table. Returns row dict or None."""
    rows = [dict(r) for r in query('SELECT * FROM passpartout')]
    return _find_closest(rows, amplada, alcada, prefix=prefix)


def calcular_cost_passpartu(amplada, alcada, tipus='simple', finestres_extra=0):
    """Calcula cost i PVD del passpartú: primer busca a taula, si no, fórmula.
    Retorna dict {'cost', 'pvd', 'origen': 'taula'|'formula'}."""
    cost_hora    = float(get_config_value('cost_hora_taller', '25'))
    temps_simple = float(get_config_value('passpartu_temps_simple', '9'))
    temps_doble  = float(get_config_value('passpartu_temps_doble', '16'))
    temps_extra  = float(get_config_value('passpartu_temps_finestra', '3.5'))
    cost_cm2     = float(get_config_value('passpartu_cost_cm2', '0.000620'))
    minim_mat    = float(get_config_value('passpartu_minim_material', '0.50'))
    marge_pas    = float(get_config_value('marge_admin_passpartu_pct', '60'))

    # 1. Min-contain sobre taula (mides estàndard)
    prefix = '1PAS' if tipus == 'simple' else 'DOBPAS'
    fila = _closest_passpartu_taula(amplada, alcada, prefix)
    if fila:
        pc = _row_get(fila, 'preu_cost')
        if pc is not None:
            cost = float(pc) + finestres_extra * (temps_extra * cost_hora / 60)
            pvd = round(cost * (1 + marge_pas / 100), 4)
            return {'cost': round(cost, 4), 'pvd': pvd, 'origen': 'taula', 'ref': fila['referencia']}

    # 2. Fallback fórmula per a mides fora de taula
    area = amplada * alcada
    cost_mat = max(area * cost_cm2, minim_mat)
    minuts = temps_doble if tipus == 'doble' else temps_simple
    cost_mo = minuts * cost_hora / 60
    cost_extra = finestres_extra * (temps_extra * cost_hora / 60)
    cost = round(cost_mat + cost_mo + cost_extra, 4)
    pvd = round(cost * (1 + marge_pas / 100), 4)
    return {'cost': cost, 'pvd': pvd, 'origen': 'formula', 'ref': f'pas-{amplada}x{alcada}'}


def _cost_muntatge(amplada, alcada, cost_cm2, temps_base, temps_var_cm2):
    """Base function shared by foam and laminat. Returns workshop cost (material + labor)."""
    cost_hora = float(get_config_value('cost_hora_taller', '25'))
    area = amplada * alcada
    mat = area * cost_cm2
    temps = temps_base + (area * temps_var_cm2)
    mo = temps * cost_hora / 60
    return round(mat + mo, 4)


def _closest_encolat_taula(amplada, alcada, prefix):
    """Min-contain on encolat_pro filtered by prefix (ENC or PRO).
    Returns cheapest row whose dimensions physically cover (amplada, alcada)."""
    rows = [dict(r) for r in query('SELECT * FROM encolat_pro')]
    return _find_closest(rows, amplada, alcada, prefix=prefix)


def calcular_cost_foam(amplada, alcada):
    """Encolat en foam adhesiu (ProEco és àlies del mateix producte).
    1. Min-contain sobre encolat_pro (refs ENC%)
    2. Fórmula fallback per mides fora de taula"""
    marge = float(get_config_value('marge_admin_encolat_pct', '60'))

    fila = _closest_encolat_taula(amplada, alcada, prefix='ENC')
    if fila and _row_get(fila, 'preu_cost') is not None:
        cost = float(fila['preu_cost'])
        pvd = round(cost * (1 + marge / 100), 4)
        return {'cost': cost, 'pvd': pvd, 'preu': pvd, 'origen': 'taula', 'ref': fila['referencia']}

    # Fallback fórmula
    cost_cm2   = float(get_config_value('foam_cost_cm2', '0.001143'))
    temps_base = float(get_config_value('foam_temps_base_min', '9'))
    temps_var  = float(get_config_value('foam_temps_var_cm2', '0.0015'))
    cost = _cost_muntatge(amplada, alcada, cost_cm2, temps_base, temps_var)
    pvd = round(cost * (1 + marge / 100), 4)
    return {'cost': cost, 'pvd': pvd, 'preu': pvd, 'origen': 'formula', 'ref': f'foam-{amplada}x{alcada}'}


def calcular_cost_laminat(amplada, alcada, tipus='semibrillo'):
    """Laminat Protter. tipus: 'semibrillo' | 'mate'
    1. Min-contain sobre encolat_pro (refs PRO%) — usa preu_cost com a base semibrillo
       i aplica diferencial per mate
    2. Fórmula fallback per mides fora de taula"""
    marge = float(get_config_value('marge_admin_encolat_pct', '60'))
    cost_cm2_semi = float(get_config_value('laminat_semibrillo_cost_cm2', '0.000504'))
    cost_cm2_mate = float(get_config_value('laminat_mate_cost_cm2', '0.000685'))
    temps_base    = float(get_config_value('laminat_temps_base_min', '12'))
    temps_var     = float(get_config_value('laminat_temps_var_cm2', '0.0012'))
    cost_cm2 = cost_cm2_mate if tipus == 'mate' else cost_cm2_semi

    fila = _closest_encolat_taula(amplada, alcada, prefix='PRO')
    if fila and _row_get(fila, 'preu_cost') is not None:
        cost_base = float(fila['preu_cost'])
        if tipus == 'mate':
            area = amplada * alcada
            extra_mat = area * (cost_cm2_mate - cost_cm2_semi)
            cost = round(cost_base + extra_mat, 4)
        else:
            cost = cost_base
        pvd = round(cost * (1 + marge / 100), 4)
        return {'cost': cost, 'pvd': pvd, 'preu': pvd, 'tipus': tipus, 'origen': 'taula', 'ref': fila['referencia']}

    # Fallback fórmula
    cost = _cost_muntatge(amplada, alcada, cost_cm2, temps_base, temps_var)
    pvd = round(cost * (1 + marge / 100), 4)
    return {'cost': cost, 'pvd': pvd, 'preu': pvd, 'tipus': tipus, 'origen': 'formula', 'ref': f'laminat-{tipus}-{amplada}x{alcada}'}


def calcular_cost_protter(amplada, alcada, tipus='semibrillo'):
    """Protter = foam adhesiu + laminat.
    Suma els costs independents de foam i laminat.
    tipus: 'semibrillo' | 'mate'"""
    marge = float(get_config_value('marge_admin_encolat_pct', '60'))
    c_foam = calcular_cost_foam(amplada, alcada)['cost']
    c_lam = calcular_cost_laminat(amplada, alcada, tipus=tipus)['cost']
    cost = round(c_foam + c_lam, 4)
    pvd = round(cost * (1 + marge / 100), 4)
    return {'cost': cost, 'pvd': pvd, 'preu': pvd, 'tipus': tipus,
            'origen': 'composicio', 'ref': f'protter-{tipus}-{amplada}x{alcada}'}


def _closest_vidre_taula(amplada, alcada, prefix=''):
    """Min-contain sobre taula vidres filtrat per prefix.
    prefix: ''=vidre simple (exclou DV-/MIR-), 'DV-'=doble vidre, 'MIR-'=mirall."""
    rows = [dict(r) for r in query('SELECT * FROM vidres')]
    if prefix == '':
        # Exclou DV- i MIR-
        rows = [r for r in rows if not (r['referencia'].upper().startswith('DV-') or r['referencia'].upper().startswith('MIR-'))]
        return _find_closest(rows, amplada, alcada)
    return _find_closest(rows, amplada, alcada, prefix=prefix)


def calcular_cost_vidre(amplada, alcada):
    """Vidre simple tallat a mida.
    1. Min-contain sobre vidres (exclou DV-/MIR-)
    2. Fórmula fallback: material × cm² + temps tall (base + lineal × perímetre)"""
    marge     = float(get_config_value('marge_admin_vidres_pct', '60'))
    cost_cm2  = float(get_config_value('vidre_cost_cm2', '0.002880'))
    t_base    = float(get_config_value('vidre_temps_base_min', '3'))
    t_lineal  = float(get_config_value('vidre_temps_lineal_m', '0.5'))
    cost_hora = float(get_config_value('cost_hora_taller', '25'))

    fila = _closest_vidre_taula(amplada, alcada, prefix='')
    if fila and _row_get(fila, 'preu_cost') is not None:
        cost = float(fila['preu_cost'])
        pvd = round(cost * (1 + marge / 100), 4)
        return {'cost': cost, 'pvd': pvd, 'preu': pvd, 'origen': 'taula', 'ref': fila['referencia']}

    # Fallback fórmula
    area = amplada * alcada
    perimetre_m = 2 * (amplada + alcada) / 100
    mat = area * cost_cm2
    temps = t_base + (t_lineal * perimetre_m)
    mo = temps * cost_hora / 60
    cost = round(mat + mo, 4)
    pvd = round(cost * (1 + marge / 100), 4)
    return {'cost': cost, 'pvd': pvd, 'preu': pvd, 'origen': 'formula', 'ref': f'vidre-{amplada}x{alcada}'}


def calcular_cost_doble_vidre(amplada, alcada):
    """Doble vidre: dos vidres simples + muntatge.
    1. Min-contain sobre DV-
    2. Fórmula: (cost_vidre_simple × 2) + cost_muntatge"""
    marge     = float(get_config_value('marge_admin_vidres_pct', '60'))
    t_muntat  = float(get_config_value('vidre_dv_muntatge_min', '5'))
    cost_hora = float(get_config_value('cost_hora_taller', '25'))

    fila = _closest_vidre_taula(amplada, alcada, prefix='DV-')
    if fila and _row_get(fila, 'preu_cost') is not None:
        cost = float(fila['preu_cost'])
        pvd = round(cost * (1 + marge / 100), 4)
        return {'cost': cost, 'pvd': pvd, 'preu': pvd, 'origen': 'taula', 'ref': fila['referencia']}

    # Fallback: 2 × vidre simple + muntatge
    simple = calcular_cost_vidre(amplada, alcada)
    mo_muntat = t_muntat * cost_hora / 60
    cost = round(simple['cost'] * 2 + mo_muntat, 4)
    pvd = round(cost * (1 + marge / 100), 4)
    return {'cost': cost, 'pvd': pvd, 'preu': pvd, 'origen': 'formula', 'ref': f'dv-{amplada}x{alcada}'}


def calcular_cost_mirall(amplada, alcada):
    """Mirall tallat a mida pel proveïdor (Cristaleria Abrio).
    Preu 5mm amb recàrrec energètic: 31,53 €/m² (= 0.003153 €/cm²).
    Facturació en múltiples de 6 dm². Sense MO de tall (el talla el proveïdor).

    1. Fórmula real sobre cost de factura (primer)
    2. Min-contain sobre taula MIR- com a fallback si la fórmula no aplica."""
    marge     = float(get_config_value('marge_admin_vidres_pct', '60'))
    cost_cm2  = float(get_config_value('mirall_cost_cm2', '0.003153'))
    multiplo  = float(get_config_value('mirall_multiplo_dm2', '6'))

    if amplada > 0 and alcada > 0 and cost_cm2 > 0 and multiplo > 0:
        # Àrea facturada (arrodonida al múltiple de dm² superior)
        area_real_dm2 = (amplada * alcada) / 100.0
        area_fact_dm2 = math.ceil(area_real_dm2 / multiplo) * multiplo
        area_fact_cm2 = area_fact_dm2 * 100.0
        cost = round(area_fact_cm2 * cost_cm2, 4)
        pvd = round(cost * (1 + marge / 100), 4)
        return {
            'cost': cost,
            'pvd': pvd,
            'preu': pvd,
            'area_real_cm2': round(amplada * alcada, 2),
            'area_fact_cm2': round(area_fact_cm2, 2),
            'origen': 'formula',
            'ref': f'mir-{amplada}x{alcada}',
        }

    # Fallback: min-contain sobre taula MIR-
    fila = _closest_vidre_taula(amplada, alcada, prefix='MIR-')
    if fila and _row_get(fila, 'preu_cost') is not None:
        cost = float(fila['preu_cost'])
        pvd = round(cost * (1 + marge / 100), 4)
        return {'cost': cost, 'pvd': pvd, 'preu': pvd, 'origen': 'taula_estimat', 'ref': fila['referencia']}

    return {'cost': 0.0, 'pvd': 0.0, 'preu': 0.0, 'origen': 'no_data', 'ref': f'mir-{amplada}x{alcada}'}


# â"€â"€ Routes: Auth â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
@app.route('/login', methods=['GET', 'POST'])
def login():
    next_path = _safe_next_path(request.values.get('next'), url_for('index'))
    login_source = (request.values.get('source') or session.get('bridge_source') or 'calc').strip().lower() or 'calc'
    login_lang = (
        (request.values.get('lang') or session.get('bridge_lang') or '').strip().lower()
        or request.accept_languages.best_match(['ca', 'es', 'en'])
        or 'ca'
    )
    web_return_url = _current_web_return_url()
    if request.method == 'POST':
        # Els usuaris es desen sempre en minúscules (signup + bridge). Normalitzem
        # aquí per evitar que "Juan@X.com" o " juan@x.com " fallin per case/espais.
        username = (request.form.get('username') or '').strip().lower()
        password = request.form.get('password') or ''
        try:
            user = query('SELECT * FROM usuaris WHERE lower(username)=?',
                         [username], one=True)
            credentials_ok = bool(user) and verify_pw(user['password'], password)
        except Exception as exc:
            app.logger.exception('login_db_error user=%s detail=%s', username[:3] + '***', exc)
            flash("No hem pogut validar l'accés ara mateix. Torna-ho a provar en uns segons.", 'error')
            return render_template(
                'login.html',
                web_return_url=web_return_url,
                login_next=next_path,
                login_source=login_source,
                login_lang=login_lang,
            )
        if credentials_ok:
            if not _user_is_allowed(user):
                status = _user_access_status(user)
                if status == 'pending':
                    flash("El teu accés encara està pendent de validació.", 'error')
                else:
                    flash("El teu accés està bloquejat. Contacta amb l'administrador.", 'error')
                return render_template(
                    'login.html',
                    web_return_url=web_return_url,
                    login_next=next_path,
                    login_source=login_source,
                    login_lang=login_lang,
                )
            # Migració progressiva del hash: si aquest usuari encara tenia
            # un SHA-256 legacy, aprofitem aquest login correcte per
            # rehashejar-lo amb pbkdf2. Silenciós si falla.
            maybe_upgrade_hash(user['id'], user['password'], password)
            try:
                _start_user_session(user)
                session['bridge_source'] = login_source
                session['bridge_lang'] = login_lang
                needs_setup = _needs_setup(user['id'])
            except Exception as exc:
                app.logger.exception('login_session_error user_id=%s detail=%s', _row_get(user, 'id'), exc)
                session.clear()
                flash("Hi ha hagut un error iniciant la sessió. Si continua, contacta amb el taller.", 'error')
                return render_template(
                    'login.html',
                    web_return_url=web_return_url,
                    login_next=next_path,
                    login_source=login_source,
                    login_lang=login_lang,
                )
            if needs_setup:
                return redirect(url_for('setup'))
            return redirect(next_path)
        flash('Usuari o contrasenya incorrectes.', 'error')
    return render_template(
        'login.html',
        web_return_url=web_return_url,
        login_next=next_path,
        login_source=login_source,
        login_lang=login_lang,
    )

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/api/public/professional-signup', methods=['POST'])
def public_professional_signup():
    expected_token = os.environ.get('PUBLIC_SIGNUP_TOKEN', '').strip()
    provided_token = request.headers.get('X-Signup-Token', '').strip()

    if not expected_token or provided_token != expected_token:
        return jsonify({'ok': False, 'error': 'forbidden'}), 403

    data = request.get_json(silent=True) or {}
    name = (data.get('name') or '').strip()
    email = (data.get('email') or '').strip().lower()
    phone = (data.get('phone') or '').strip()
    business_name = (data.get('business_name') or '').strip()
    web_url = (data.get('web_url') or '').strip()
    instagram = (data.get('instagram') or '').strip()
    fiscal_id = (data.get('fiscal_id') or '').strip()
    subject = (data.get('subject') or '').strip()
    message = (data.get('message') or '').strip()
    profile_type = _clean_profile_type(data.get('profile_type'))

    if not name or not email:
        return jsonify({'ok': False, 'error': 'missing_fields'}), 400

    if '@' not in email or '.' not in email.split('@')[-1]:
        return jsonify({'ok': False, 'error': 'invalid_email'}), 400

    notes = "\n".join([
        f"Alta des de la web principal el {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        f"Nom: {name}",
        f"Email: {email}",
        f"Telèfon: {phone or '-'}",
        f"Assumpte: {subject or '-'}",
        f"Perfil: {profile_type}",
        f"Empresa: {business_name or '-'}",
        f"Web: {web_url or '-'}",
        f"Instagram: {instagram or '-'}",
        f"CIF/NIF: {fiscal_id or '-'}",
        "Missatge:",
        message or '-',
    ])

    existing = query('SELECT id, is_admin, access_status FROM usuaris WHERE username=?', [email], one=True)
    if existing:
        current_status = _user_access_status(existing)
        next_status = current_status if current_status in ('active', 'blocked') else 'pending'
        execute(
            'UPDATE usuaris SET nom=?, nom_empresa=?, profile_type=?, web_url=?, instagram=?, fiscal_id=?, notes_validacio=?, access_status=? WHERE id=?',
            [name, business_name, profile_type, web_url, instagram, fiscal_id, notes, next_status, existing['id']]
        )
        return jsonify({'ok': True, 'action': 'updated', 'status': next_status})

    temp_password = secrets.token_urlsafe(12)
    execute(
        'INSERT INTO usuaris (username, password, nom, is_admin, nom_empresa, access_status, profile_type, web_url, instagram, fiscal_id, notes_validacio) VALUES (?,?,?,?,?,?,?,?,?,?,?)',
        [email, hash_pw(temp_password), name, 0, business_name, 'pending', profile_type, web_url, instagram, fiscal_id, notes]
    )
    return jsonify({'ok': True, 'action': 'created', 'status': 'pending'})


@app.route('/api/public/bridge-login', methods=['POST'])
def public_bridge_login():
    expected_token = _bridge_api_token()
    provided_token = request.headers.get('X-Bridge-Token', '').strip()

    if not expected_token or provided_token != expected_token:
        return jsonify({'ok': False, 'error': 'forbidden'}), 403

    data = request.get_json(silent=True) or {}
    username = (data.get('username') or data.get('email') or '').strip().lower()
    password = data.get('password') or ''
    lang = (data.get('lang') or 'ca').strip().lower()
    source = (data.get('source') or 'web_private').strip().lower()
    service = (data.get('service') or '').strip().lower()
    next_path = _safe_next_path(data.get('next') or '', '/calculadora' if service == 'frames' else '/')

    if not username or not password:
        return jsonify({'ok': False, 'error': 'missing_credentials'}), 400

    # lower(username) per consistència amb /login i bridge-refresh: un
    # usuari antic desat amb majúscules a la BD podia fallar aquí tot
    # entrant bé per la resta d'endpoints.
    user = query('SELECT * FROM usuaris WHERE lower(username)=?', [username], one=True)
    if not user or not verify_pw(user['password'], password):
        return jsonify({'ok': False, 'error': 'invalid_credentials'}), 401

    if not _user_is_allowed(user):
        return jsonify({'ok': False, 'error': _user_access_status(user)}), 403

    # Migració progressiva del hash: aprofitem aquest login via bridge
    # per rehashejar si encara era un SHA-256 legacy.
    maybe_upgrade_hash(user['id'], user['password'], password)

    token = _build_bridge_token({
        'uid': user['id'],
        'next': next_path,
        'lang': lang,
        'source': source,
        'service': service,
        'iat': int(time.time()),
    })
    host = request.host_url.rstrip('/')
    return jsonify({
        'ok': True,
        'redirect_url': f'{host}{url_for("bridge_auth")}?token={token}',
        'next': next_path,
    })


@app.route('/api/public/bridge-refresh', methods=['POST'])
def public_bridge_refresh():
    """Emet un token de bridge curt de durada per a un usuari que ja ha
    validat credencials a la web pública (reusrevela-web).

    Es protegeix amb X-Bridge-Token (el mateix que bridge-login): si algú
    té aquest secret, ja pot emetre tokens a lliure voluntat via
    bridge-login. Per tant aquest endpoint no obre cap superfície nova,
    només permet refrescar el token d'SSO (p.ex. per entrar a la
    calculadora des de l'àrea privada passats els 120s de vida del token
    original) sense demanar la contrasenya una altra vegada.
    """
    expected_token = _bridge_api_token()
    provided_token = request.headers.get('X-Bridge-Token', '').strip()
    if not expected_token or provided_token != expected_token:
        return jsonify({'ok': False, 'error': 'forbidden'}), 403

    data = request.get_json(silent=True) or {}
    username = (data.get('username') or data.get('email') or '').strip().lower()
    lang = (data.get('lang') or 'ca').strip().lower()
    source = (data.get('source') or 'web_private').strip().lower()
    service = (data.get('service') or '').strip().lower()
    next_path = _safe_next_path(data.get('next') or '', '/calculadora' if service == 'frames' else '/')

    if not username:
        return jsonify({'ok': False, 'error': 'missing_username'}), 400

    user = query('SELECT id, is_admin, access_status FROM usuaris WHERE lower(username)=?', [username], one=True)
    if not user:
        return jsonify({'ok': False, 'error': 'not_found'}), 404

    if not _user_is_allowed(user):
        return jsonify({'ok': False, 'error': _user_access_status(user)}), 403

    token = _build_bridge_token({
        'uid': user['id'],
        'next': next_path,
        'lang': lang,
        'source': source,
        'service': service,
        'iat': int(time.time()),
    })
    host = request.host_url.rstrip('/')
    return jsonify({
        'ok': True,
        'redirect_url': f'{host}{url_for("bridge_auth")}?token={token}',
        'next': next_path,
    })


@app.route('/api/public/professional-summary', methods=['POST'])
def public_professional_summary():
    try:
        expected_token = _bridge_api_token()
        provided_token = request.headers.get('X-Bridge-Token', '').strip()

        if not expected_token or provided_token != expected_token:
            return jsonify({'ok': False, 'error': 'unauthorized'}), 401

        data = request.get_json(silent=True) or {}
        username = (data.get('username') or '').strip().lower()
        if not username:
            return jsonify({'ok': False, 'error': 'not_found'}), 404

        user = query(
            'SELECT id, nom, nom_empresa, profile_type, access_status, marge, marge_impressio, margins_json FROM usuaris WHERE lower(username)=?',
            [username],
            one=True,
        )
        if not user:
            return jsonify({'ok': False, 'error': 'not_found'}), 404

        margins = _load_user_commercial_margins(user)

        recent_quotes_rows = query(
            '''SELECT id, data, num_pressupost, client_nom, preu_final, pagat, entregat, pendent
               FROM comandes
               WHERE user_id = ?
               ORDER BY data DESC
               LIMIT 5''',
            [user['id']],
        ) or []

        recent_quotes = []
        for row in recent_quotes_rows:
            recent_quotes.append({
                'id': row['id'],
                'date': row['data'] or '',
                'num_pressupost': row['num_pressupost'] or '',
                'client_nom': row['client_nom'] or '',
                'preu_final': float(row['preu_final'] or 0),
                'pagat': bool(row['pagat']),
                'entregat': bool(row['entregat']),
                'pendent': bool(row['pendent']),
            })

        return jsonify({
            'ok': True,
            'name': user['nom'] or '',
            'business_name': user['nom_empresa'] or '',
            'profile_type': _clean_profile_type(user['profile_type']),
            'access_status': _user_access_status(user),
            'recent_quotes': recent_quotes,
            # Marges del calc (font autoritativa). El web hauria de mostrar aquests
            # valors en lloc del JSON local, que ara fa només de cache.
            'margins': margins,
            'marge': float(user['marge']) if user['marge'] is not None else 60.0,
            'marge_impressio': float(user['marge_impressio']) if user['marge_impressio'] is not None else 0.0,
        })
    except Exception as exc:
        print(f'professional-summary error: {exc}')
        return jsonify({
            'ok': False,
            'error': 'internal_error',
            'name': '',
            'business_name': '',
            'profile_type': 'professional',
            'access_status': 'pending',
            'recent_quotes': [],
        }), 200


@app.route('/api/public/commercial-settings-sync', methods=['POST'])
def public_commercial_settings_sync():
    expected_token = _bridge_api_token()
    provided_token = request.headers.get('X-Bridge-Token', '').strip()
    if not expected_token or provided_token != expected_token:
        return jsonify({'ok': False, 'error': 'forbidden'}), 403

    data = request.get_json(silent=True) or {}
    username = (data.get('username') or '').strip().lower()
    if not username:
        return jsonify({'ok': False, 'error': 'missing_username'}), 400

    user = query('SELECT id FROM usuaris WHERE lower(username)=?', [username], one=True)
    if not user:
        return jsonify({'ok': False, 'error': 'user_not_found'}), 404

    margins = _normalize_commercial_margins(
        data.get('margins') if isinstance(data.get('margins'), dict) else data,
        frame_margin=data.get('marge'),
        print_margin=data.get('marge_impressio'),
    )
    marge = margins['frames']
    marge_impressio = margins['prints']
    execute(
        'UPDATE usuaris SET marge=?, marge_impressio=?, margins_json=? WHERE id=?',
        [marge, marge_impressio, json.dumps(margins, ensure_ascii=True), user['id']]
    )
    return jsonify({'ok': True, 'username': username, 'marge': marge, 'marge_impressio': marge_impressio, 'margins': margins})


@app.route('/api/public/pricing', methods=['GET'])
def public_pricing():
    """Exposa les tarifes professionals (sense marge aplicat) perquè
    la web les pugui consumir i aplicar-hi el marge propi de cada client.

    Autenticació: capçalera X-Bridge-Token (mateix token que la resta
    d'endpoints de bridge).

    Retorna:
      impressio     — taula completa de còpies fotogràfiques (ref, preu, descripcio)
      laminate_only — preus de laminat sol per mida (ref, preu)
      encolat_pro   — preus de muntatge encolat i protter per mida (ref, preu, tipus)
    """
    expected_token = _bridge_api_token()
    provided_token = request.headers.get('X-Bridge-Token', '').strip()
    if not expected_token or provided_token != expected_token:
        return jsonify({'ok': False, 'error': 'forbidden'}), 403

    # Còpies fotogràfiques
    impressio_rows = query('SELECT referencia, preu, descripcio FROM impressio ORDER BY preu') or []
    impressio = [
        {'ref': r['referencia'], 'preu': float(r['preu'] or 0), 'descripcio': r['descripcio'] or ''}
        for r in impressio_rows
    ]

    # Laminat sol (hardcoded en constants, igual que a la calculadora)
    laminate_only = [
        {'ref': ref, 'preu': float(preu)}
        for ref, preu in sorted(LAMINATE_ONLY_PRICES.items(),
                                key=lambda x: float(x[1]))
    ]

    # Encolat professional i protter (taula encolat_pro, prefix ENC / PRO)
    encolat_rows = query('SELECT referencia, preu, preu_cost FROM encolat_pro ORDER BY preu') or []
    encolat_pro = []
    for r in encolat_rows:
        ref = r['referencia'] or ''
        tipus = 'protter' if ref.upper().startswith('PRO') else 'encolat'
        mida = ref[3:] if ref.upper().startswith('PRO') or ref.upper().startswith('ENC') else ref
        pc = _row_get(r, 'preu_cost')
        pvd = calcular_pvd(pc, 'encolat') if pc is not None else None
        # NO exposar preu_cost al Repo B — és dada interna de l'admin.
        # Només enviem el PVD (preu "taller" que paga el professional).
        encolat_pro.append({
            'ref': ref,
            'mida': mida,
            'preu': pvd if pvd is not None else float(r['preu'] or 0),
            'tipus': tipus,
        })

    return jsonify({
        'ok': True,
        'impressio': impressio,
        'laminate_only': laminate_only,
        'encolat_pro': encolat_pro,
    })


@app.route('/api/public/compute', methods=['GET'])
def public_compute():
    """Calcula un preu base (sense marge del client, sense IVA) per un producte concret.

    Auth: X-Bridge-Token (mateix que la resta de bridge endpoints).

    Query params:
      kind        = 'impressio' | 'laminate' | 'protter' | 'frame'
      width_cm    = int (obligatori)
      height_cm   = int (obligatori)
      paper       = string opcional (impressió: 'lustre','silk','fine_art',...)
      finish      = string opcional ('none','laminate','protter','foam')
      moldura_id  = string opcional (només si kind=frame)
      qty         = int opcional, defecte 1

    Resposta: {ok, kind, width_cm, height_cm, base_price, vat_rate, breakdown}
    """
    expected_token = _bridge_api_token()
    provided_token = request.headers.get('X-Bridge-Token', '').strip()
    if not expected_token or provided_token != expected_token:
        return jsonify({'ok': False, 'error': 'forbidden'}), 403

    kind = (request.args.get('kind') or '').strip().lower()
    if kind not in ('impressio', 'laminate', 'protter', 'frame'):
        return jsonify({'ok': False, 'error': 'unknown_kind'}), 400

    try:
        w = float(request.args.get('width_cm', 0))
        h = float(request.args.get('height_cm', 0))
    except (TypeError, ValueError):
        return jsonify({'ok': False, 'error': 'invalid_size'}), 400
    if w <= 0 or h <= 0:
        return jsonify({'ok': False, 'error': 'invalid_size'}), 400

    try:
        qty = max(1, int(request.args.get('qty', 1)))
    except (TypeError, ValueError):
        qty = 1

    paper = (request.args.get('paper') or '').strip().lower()
    finish = (request.args.get('finish') or 'none').strip().lower()
    moldura_id = (request.args.get('moldura_id') or '').strip()

    breakdown = {}
    base_price = 0.0

    if kind == 'impressio':
        # Impressió: usa la taula amb min-contain (no té preu_cost, ja porta marge)
        imp = _imp_closest(w, h)
        if not imp:
            return jsonify({'ok': False, 'error': 'impressio_not_found'}), 404
        base_price = float(imp.get('preu', 0))
        breakdown['impressio'] = base_price
        # Acabats opcionals
        if finish == 'laminate':
            lam = _laminate_only_closest(w, h)
            if lam:
                breakdown['finish'] = float(lam.get('preu', 0))
                base_price += breakdown['finish']
        elif finish == 'protter':
            prot = calcular_cost_laminat(w, h, tipus='semibrillo')
            breakdown['finish'] = prot['pvd']
            base_price += prot['pvd']
        elif finish == 'foam':
            foam = calcular_cost_foam(w, h)
            breakdown['finish'] = foam['pvd']
            base_price += foam['pvd']

    elif kind == 'laminate':
        lam = _laminate_only_closest(w, h)
        if not lam:
            return jsonify({'ok': False, 'error': 'laminate_not_found'}), 404
        base_price = float(lam.get('preu', 0))
        breakdown['laminate'] = base_price

    elif kind == 'protter':
        r = calcular_cost_laminat(w, h, tipus='semibrillo')
        base_price = r['pvd']
        breakdown['protter'] = r['pvd']
        breakdown['origen'] = r['origen']

    elif kind == 'frame':
        if not moldura_id:
            return jsonify({'ok': False, 'error': 'missing_moldura_id'}), 400
        m = query('SELECT preu_taller, preu_cost, gruix, merma_pct, minim_cm FROM moldures WHERE LOWER(referencia)=LOWER(?)', [moldura_id], one=True)
        if not m:
            return jsonify({'ok': False, 'error': 'moldura_not_found'}), 404
        pc = _row_get(m, 'preu_cost')
        gruix = float(_row_get(m, 'gruix') or 0)
        merma = float(_row_get(m, 'merma_pct') or 10.0)
        minim = float(_row_get(m, 'minim_cm') or 100.0)
        if pc is not None:
            marc = calcular_preu_marc(w, h, gruix, float(pc), merma_pct=merma, minim_cm=minim)
            if not marc:
                return jsonify({'ok': False, 'error': 'compute_failed'}), 500
            base_price = marc['pvd']
            breakdown['moldura'] = marc['pvd']
            breakdown['cost'] = marc['cost']
        else:
            # Fallback a preu_taller com €/cm lineal
            preu_cm = float(_row_get(m, 'preu_taller') or 0)
            perimetre = 2 * (w + h)
            longitud = (perimetre + gruix * 8) * (1 + merma / 100)
            longitud = max(longitud, minim)
            base_price = round(longitud * preu_cm / 100, 4)
            breakdown['moldura'] = base_price

    # Apply quantity
    total = round(base_price * qty, 4)

    return jsonify({
        'ok': True,
        'kind': kind,
        'width_cm': w,
        'height_cm': h,
        'qty': qty,
        'base_price': total,
        'vat_rate': 0.21,
        'breakdown': breakdown,
    })


@app.route('/auth/bridge')
def bridge_auth():
    payload = _read_bridge_token(request.args.get('token'))
    if not payload:
        flash("L'accés unificat ha caducat o no és vàlid. Torna a iniciar sessió.", 'error')
        return redirect(url_for('login'))

    user = query('SELECT * FROM usuaris WHERE id=?', [payload.get('uid')], one=True)
    if not _user_is_allowed(user):
        flash("No hem pogut validar aquest accés. Contacta amb administració si cal.", 'error')
        return redirect(url_for('login'))

    _start_user_session(user)
    session['bridge_source'] = str(payload.get('source') or '').strip().lower()
    session['bridge_lang'] = str(payload.get('lang') or 'ca').strip().lower()
    target = _safe_next_path(payload.get('next'), '/')
    if _needs_setup(user['id']) and target != url_for('logout'):
        return redirect(url_for('setup'))
    return redirect(target)

# â"€â"€ Routes: App principal â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€

@app.route('/ajuda')
@login_required
def ajuda():
    return render_template('ajuda.html')

@app.route('/setup')
@login_required
def setup():
    return render_template('setup.html')

@app.route('/api/setup-done', methods=['POST'])
@login_required
def setup_done():
    try:
        execute('UPDATE usuaris SET setup_done=1 WHERE id=?', [session['user_id']])
    except:
        pass
    return jsonify({'ok': True})

@app.route('/')
@login_required
def index():
    try:
        u = query('SELECT setup_done FROM usuaris WHERE id=?', [session['user_id']], one=True)
        if u and not bool(_row_get(u, 'setup_done', 0)):
            return redirect(url_for('setup'))
    except:
        pass
    web_url = _current_web_return_url()
    if web_url:
        return redirect(web_url)
    return redirect(url_for('calculadora'))


@app.route('/calculadora')
@login_required
def calculadora():
    try:
        u = query('SELECT setup_done FROM usuaris WHERE id=?', [session['user_id']], one=True)
        if u and not bool(_row_get(u, 'setup_done', 0)):
            return redirect(url_for('setup'))
    except:
        pass
    user = query('SELECT brand_color, marge_pro_pct, marge, marge_impressio_pro_pct, marge_impressio FROM usuaris WHERE id=?', [session['user_id']], one=True)
    brand_color = _normalize_hex_color(_row_get(user, 'brand_color', DEFAULT_BRAND_COLOR))
    marge_pro_actiu = get_config_value('marge_pro_actiu', '1') == '1'
    marge_pro = _get_marge_value(user) if marge_pro_actiu else 0.0
    marge_imp_pro = _get_marge_impressio_value(user) if marge_pro_actiu else 0.0
    return render_template('calculadora.html',
                           web_return_url=_current_web_return_url(),
                           web_order_url=_current_web_order_url(),
                           color_filters=MOLDURA_COLOR_FILTERS,
                           gruix_filters=MOLDURA_GRUIX_FILTERS,
                           brand_color=brand_color,
                           brand_color_light=_mix_with_white(brand_color),
                           marge_pro_actiu=marge_pro_actiu,
                           marge_pro=marge_pro,
                           marge_impressio_pro=marge_imp_pro)

@app.route('/api/lookup')
@login_required
def lookup():
    ref = request.args.get('ref', '').strip()
    tipus = request.args.get('tipus', 'moldura')
    is_admin = session.get('is_admin')
    if tipus == 'moldura':
        try:
            r = query('SELECT preu_taller, preu_cost, gruix, merma_pct, minim_cm, descripcio, foto, ref2 FROM moldures WHERE LOWER(referencia)=LOWER(?)', [ref], one=True)
            print(f"lookup moldura ref={ref} result={r}")
            if r:
                pvd = calcular_pvd(_row_get(r, 'preu_cost'), 'moldures')
                preu = pvd if pvd is not None else r['preu_taller']
                resp = {'ok': True, 'preu': preu, 'gruix': r['gruix'],
                        'merma_pct': _row_get(r, 'merma_pct', 10.0),
                        'minim_cm': _row_get(r, 'minim_cm', 100.0),
                        'descripcio': r['descripcio'], 'foto': _resolve_moldura_photo(ref, r['foto'], ref2=_row_get(r,'ref2',''))}
                if is_admin:
                    resp['preu_cost'] = _row_get(r, 'preu_cost')
                return jsonify(resp)
        except Exception as e:
            print(f"lookup ERROR: {e}")
            return jsonify({'ok': False, 'error': str(e)})
    elif tipus == 'vidre':
        r = query('SELECT preu, preu_cost FROM vidres WHERE LOWER(referencia)=LOWER(?)', [ref], one=True)
        if r:
            pvd = calcular_pvd(_row_get(r, 'preu_cost'), 'vidres')
            preu = pvd if pvd is not None else r['preu']
            resp = {'ok': True, 'preu': preu}
            if is_admin: resp['preu_cost'] = _row_get(r, 'preu_cost')
            return jsonify(resp)
    elif tipus == 'passpartout':
        try:
            parts = ref.lower().replace('pas-','').split('x')
            ew, eh = float(parts[0]), float(parts[1])
            is_doble = 'doble' in ref.lower() or 'dob' in ref.lower()
            r = calcular_cost_passpartu(ew, eh, tipus='doble' if is_doble else 'simple')
            resp = {'ok': True, 'preu': r['pvd'], 'origen': r['origen']}
            if is_admin:
                resp['preu_cost'] = r['cost']
            return jsonify(resp)
        except Exception:
            return jsonify({'ok': False})
    elif tipus == 'encolat':
        r = query('SELECT preu, preu_cost FROM encolat_pro WHERE LOWER(referencia)=LOWER(?)', [ref], one=True)
        if r:
            pvd = calcular_pvd(_row_get(r, 'preu_cost'), 'encolat')
            preu = pvd if pvd is not None else r['preu']
            resp = {'ok': True, 'preu': preu}
            if is_admin: resp['preu_cost'] = _row_get(r, 'preu_cost')
            return jsonify(resp)
    elif tipus == 'impressio':
        r = query('SELECT preu, descripcio FROM impressio WHERE LOWER(referencia)=LOWER(?)', [ref], one=True)
        if r: return jsonify({'ok': True, 'preu': r['preu'], 'descripcio': r['descripcio']})
    return jsonify({'ok': False})

@app.route('/api/refs')
@login_required
def refs():
    tipus = request.args.get('tipus', 'moldura')
    tables = {'moldura': ('moldures', 'referencia'),
              'vidre': ('vidres', 'referencia'),
              'passpartout': ('passpartout', 'referencia'),
              'encolat': ('encolat_pro', 'referencia'),
              'impressio': ('impressio', 'referencia')}
    if tipus not in tables: return jsonify([])
    t, col = tables[tipus]
    rows = query(f'SELECT {col} FROM {t} ORDER BY {col}')
    col = list(rows[0].keys())[0] if rows else 'referencia'
    return jsonify([r[col] for r in rows])


@app.route('/api/pendents-albara')
@login_required
def api_pendents_albara():
    if not session.get('is_admin'):
        return jsonify({'n': 0})
    row = query('''SELECT COUNT(*) as n FROM comandes
                   WHERE observacions LIKE '%[ACCEPTAT]%'
                     AND (fd_albara IS NULL OR fd_albara='')''', one=True)
    return jsonify({'n': row['n'] if row else 0})


@app.route('/api/feedback', methods=['POST'])
@login_required
def api_feedback():
    d = request.get_json(force=True) or {}
    missatge = (d.get('missatge') or '').strip()
    if not missatge:
        return jsonify({'ok': False, 'error': 'Cal escriure un missatge'}), 400
    tipus = d.get('tipus', 'millora')
    if tipus not in ('error', 'millora', 'altre'):
        tipus = 'millora'
    pagina = (d.get('pagina') or '')[:200]
    from datetime import datetime
    execute("INSERT INTO feedback (user_id, tipus, missatge, pagina, data) VALUES (?,?,?,?,?)",
            [session['user_id'], tipus, missatge[:2000], pagina, datetime.now().strftime('%Y-%m-%d %H:%M')])
    return jsonify({'ok': True})


@app.route('/admin/feedback')
@admin_required
def admin_feedback():
    rows = query('''SELECT f.*, u.nom as usuari_nom FROM feedback f
                    JOIN usuaris u ON f.user_id=u.id
                    ORDER BY f.id DESC LIMIT 100''')
    # Mark all as read
    execute("UPDATE feedback SET llegit=1 WHERE llegit=0")
    return render_template('admin_feedback.html', feedbacks=rows)


@app.route('/api/feedback/count')
@login_required
def api_feedback_count():
    if not session.get('is_admin'):
        return jsonify({'n': 0})
    row = query("SELECT COUNT(*) as n FROM feedback WHERE llegit=0", one=True)
    return jsonify({'n': row['n'] if row else 0})


@app.route('/admin/preus-cost')
@admin_required
def admin_preus_cost():
    """Llistat de preus de cost amb filtres per taula, proveïdor i verificat."""
    taules_valides = ['moldures', 'vidres', 'encolat_pro', 'passpartout']
    taula = request.args.get('taula', 'moldures')
    if taula not in taules_valides:
        taula = 'moldures'
    proveidor = request.args.get('proveidor', '').strip()
    verificat = request.args.get('verificat', '')

    preu_orig = 'preu_taller' if taula == 'moldures' else 'preu'
    cols = f'referencia, descripcio, preu_cost, preu_cost_ant, data_cost, cost_verificat, notes_cost, {preu_orig} as preu_original'
    if taula == 'moldures':
        cols += ', proveidor'

    conditions, args = [], []
    if proveidor and taula == 'moldures':
        conditions.append("LOWER(proveidor) LIKE LOWER(?)")
        args.append(f'%{proveidor}%')
    if verificat == '1':
        conditions.append("cost_verificat = 1")
    elif verificat == '0':
        conditions.append("(cost_verificat = 0 OR cost_verificat IS NULL)")

    where = (' WHERE ' + ' AND '.join(conditions)) if conditions else ''
    sql = f'SELECT {cols} FROM {taula}{where} ORDER BY cost_verificat ASC, referencia ASC'
    rows = query(sql, args)

    # Proveïdors per al filtre (només moldures)
    proveidors = []
    if taula == 'moldures':
        prov_rows = query("SELECT DISTINCT proveidor FROM moldures WHERE proveidor IS NOT NULL AND proveidor != '' ORDER BY proveidor")
        proveidors = [r['proveidor'] for r in (prov_rows or [])]

    # Compute PVD + estadístiques per taula
    cat_map = {'moldures': 'moldures', 'vidres': 'vidres', 'encolat_pro': 'encolat', 'passpartout': 'passpartu'}
    categoria = cat_map.get(taula, 'moldures')
    for r in rows:
        pc = _row_get(r, 'preu_cost')
        r['pvd'] = calcular_pvd(pc, categoria) if pc is not None else None

    stats = {}
    for t in taules_valides:
        s = query(f"SELECT COUNT(*) as total, SUM(CASE WHEN cost_verificat=1 THEN 1 ELSE 0 END) as verificats FROM {t}", one=True)
        stats[t] = {'total': s['total'] or 0, 'verificats': s['verificats'] or 0}

    cfg = {r['clau']: r['valor'] for r in (query("SELECT clau, valor FROM config") or [])}
    return render_template('admin_preus_cost.html', rows=rows, taula=taula,
                           proveidor=proveidor, verificat=verificat, proveidors=proveidors, stats=stats, config=cfg)


@app.route('/admin/preus-cost/update', methods=['POST'])
@admin_required
def admin_preus_cost_update():
    """Actualitza el preu de cost d'un producte i registra a l'historial."""
    d = request.get_json(force=True) or {}
    taula = d.get('taula', 'moldures')
    if taula not in ('moldures', 'vidres', 'encolat_pro', 'passpartout'):
        return jsonify({'ok': False, 'error': 'Taula no vàlida'}), 400
    referencia = (d.get('referencia') or '').strip()
    if not referencia:
        return jsonify({'ok': False, 'error': 'Falta referència'}), 400
    try:
        preu_cost_nou = round(float(d.get('preu_cost_nou', 0)), 4)
    except (TypeError, ValueError):
        return jsonify({'ok': False, 'error': 'Preu no vàlid'}), 400
    notes = (d.get('notes') or '')[:500]

    # Read current cost
    row = query(f'SELECT preu_cost FROM {taula} WHERE referencia=?', [referencia], one=True)
    if not row:
        return jsonify({'ok': False, 'error': f'Referència {referencia} no trobada a {taula}'}), 404
    preu_cost_ant = _row_get(row, 'preu_cost')

    from datetime import datetime
    avui = datetime.now().strftime('%Y-%m-%d')

    # Update the product
    execute(f'''UPDATE {taula} SET preu_cost_ant=?, preu_cost=?, data_cost=?,
                usuari_cost_id=?, cost_verificat=1, notes_cost=?
                WHERE referencia=?''',
            [preu_cost_ant, preu_cost_nou, avui, session['user_id'], notes, referencia])

    # Also update legacy preu/preu_taller to keep backward compat
    cat_map = {'moldures': 'moldures', 'vidres': 'vidres', 'encolat_pro': 'encolat', 'passpartout': 'passpartu'}
    pvd = calcular_pvd(preu_cost_nou, cat_map.get(taula, 'moldures'))
    if taula == 'moldures':
        execute('UPDATE moldures SET preu_taller=? WHERE referencia=?', [pvd, referencia])
    else:
        execute(f'UPDATE {taula} SET preu=? WHERE referencia=?', [pvd, referencia])

    # Insert audit trail
    execute('''INSERT INTO historial_preus_cost (taula, referencia, preu_cost_antic, preu_cost_nou, usuari_id, data, notes)
               VALUES (?,?,?,?,?,?,?)''',
            [taula, referencia, preu_cost_ant, preu_cost_nou, session['user_id'], avui, notes])

    marge_admin = float(get_config_value(f'marge_admin_{cat_map.get(taula, "moldures")}_pct', '60'))
    return jsonify({'ok': True, 'preu_cost_nou': preu_cost_nou, 'pvd': pvd,
                     'preu_cost_ant': preu_cost_ant, 'marge_aplicat': marge_admin})


@app.route('/admin/preus-cost/historial')
@admin_required
def admin_preus_cost_historial():
    """Historial de canvis de preu de cost per una referència."""
    referencia = request.args.get('referencia', '').strip()
    taula = request.args.get('taula', '')
    conditions, args = [], []
    if referencia:
        conditions.append("referencia=?")
        args.append(referencia)
    if taula:
        conditions.append("taula=?")
        args.append(taula)
    where = ' WHERE ' + ' AND '.join(conditions) if conditions else ''
    rows = query(f'''SELECT h.*, u.nom as usuari_nom
                     FROM historial_preus_cost h
                     LEFT JOIN usuaris u ON h.usuari_id=u.id
                     {where}
                     ORDER BY h.id DESC LIMIT 200''', args)
    return jsonify([dict(r) for r in (rows or [])])


@app.route('/api/moldura-options')
@login_required
def moldura_options():
    rows = query("""SELECT referencia, gruix, descripcio, foto, ref2
                    FROM moldures
                    ORDER BY referencia""")
    return jsonify(_serialize_moldures(rows))

@app.route('/api/marge')
@login_required
def get_marge():
    u = query('SELECT marge, marge_pro_pct, marge_impressio, marge_impressio_pro_pct, nom_empresa, empresa_adreca, empresa_tel, margins_json, brand_color, brand_color_secondary, brand_color_menu FROM usuaris WHERE id=?', [session['user_id']], one=True)
    # Prioritat clara: marge_pro_pct > marge (legacy) > 60.
    # margins_json NO s'ha de fer servir per al marge general — només per a
    # categories específiques (albums, canvas…) si se n'afegeixen en el futur.
    marge_pro_actiu = get_config_value('marge_pro_actiu', '0') == '1'
    if marge_pro_actiu:
        marge = _get_marge_value(u)
    else:
        marge = 0.0
    # marge_impressio es calcula sempre, independentment del flag, perquè
    # marge_impressio_pro_pct és un valor específic de la impressió que no
    # hauria de quedar a 0 només pel fet que el flag d'overrides està off.
    marge_imp = _get_marge_impressio_value(u)
    nom_emp = u['nom_empresa'] if u and u['nom_empresa'] else ''
    brand_color = _normalize_hex_color(_row_get(u, 'brand_color', DEFAULT_BRAND_COLOR))
    brand_color_secondary = _normalize_hex_color(
        _row_get(u, 'brand_color_secondary', DEFAULT_BRAND_SECONDARY_COLOR),
        DEFAULT_BRAND_SECONDARY_COLOR,
    )
    brand_color_menu = _normalize_hex_color(
        _row_get(u, 'brand_color_menu', brand_color),
        brand_color,
    )
    margins = _load_user_commercial_margins(u)
    cfg_rows = query("SELECT clau, valor FROM config WHERE clau LIKE 'empresa_%'")
    cfg = {r['clau']: r['valor'] for r in (cfg_rows or [])}
    if not nom_emp:
        nom_emp = cfg.get('empresa_nom', 'Objectiu Emmarcació')
    user_adreca = _row_get(u, 'empresa_adreca', '') or ''
    user_tel    = _row_get(u, 'empresa_tel', '') or ''
    return jsonify({
        'marge': marge,
        'marge_impressio': marge_imp,
        'margins': margins,
        'brand_color': brand_color,
        'empresa_nom': nom_emp,
        'empresa_adreca': user_adreca if user_adreca else cfg.get('empresa_adreca',''),
        'empresa_tel':    user_tel    if user_tel    else cfg.get('empresa_tel',''),
    })


@app.route('/api/logo', methods=['POST'])
@login_required
def upload_logo():
    f = request.files.get('logo')
    if not f: return jsonify({'ok': False})
    import base64
    data = f.read()
    ext = f.filename.rsplit('.', 1)[-1].lower() if '.' in f.filename else 'png'
    mime = 'image/png' if ext == 'png' else 'image/jpeg'
    b64 = 'data:' + mime + ';base64,' + base64.b64encode(data).decode()
    # Save per user
    try:
        execute('UPDATE usuaris SET logo_b64=? WHERE id=?', [b64, session['user_id']])
    except:
        pass
    return jsonify({'ok': True})

@app.route('/api/logo', methods=['GET'])
@login_required
def get_logo():
    r = query('SELECT logo_b64 FROM usuaris WHERE id=?', [session['user_id']], one=True)
    return jsonify({'url': _row_get(r, 'logo_b64', '') or ''})

@app.route('/static/logo-preview')
@login_required
def logo_preview():
    import base64
    r = query("SELECT valor FROM config WHERE clau='empresa_logo'", one=True)
    if not r or not r['valor']: return '', 404
    data_url = r['valor']
    if data_url.startswith('data:'):
        header, b64 = data_url.split(',', 1)
        mime = header.split(':')[1].split(';')[0]
        data = base64.b64decode(b64)
        from flask import Response
        return Response(data, mimetype=mime)
    return '', 404

@app.route('/api/empresa', methods=['POST'])
@admin_required
def api_empresa():
    # Nota: aquest endpoint és només per admins i actualitza el config global
    # que actua de valors per defecte. Els clients no-admin guarden les seves
    # dades pròpies a usuaris.empresa_adreca / empresa_tel via /api/desar-marge.
    d = request.json or {}
    nom    = d.get('nom','')
    adreca = d.get('adreca','')
    tel    = d.get('tel','')
    execute("UPDATE config SET valor=? WHERE clau='empresa_nom'",    [nom])
    execute("UPDATE config SET valor=? WHERE clau='empresa_adreca'", [adreca])
    execute("UPDATE config SET valor=? WHERE clau='empresa_tel'",    [tel])
    return jsonify({'ok': True})

@app.route('/api/desar-marge', methods=['POST'])
@login_required
def desar_marge():
    d = request.get_json(silent=True) or {}
    m = float(d.get('marge', 60))
    mi = float(d.get('marge_impressio', 100))
    ne = d.get('nom_empresa', '')
    nf = d.get('nom_fiscal', '')
    fi = d.get('fiscal_id', '')
    ea = d.get('empresa_adreca', '')
    et = d.get('empresa_tel', '')
    brand_color = _normalize_hex_color(d.get('brand_color', DEFAULT_BRAND_COLOR))
    brand_color_secondary = _normalize_hex_color(
        d.get('brand_color_secondary', DEFAULT_BRAND_SECONDARY_COLOR),
        DEFAULT_BRAND_SECONDARY_COLOR,
    )
    brand_color_menu = _normalize_hex_color(
        d.get('brand_color_menu', brand_color),
        brand_color,
    )
    margins = _normalize_commercial_margins(
        d.get('margins') if isinstance(d.get('margins'), dict) else d,
        frame_margin=m,
        print_margin=mi,
    )
    execute(
        'UPDATE usuaris SET marge=?, marge_impressio=?, nom_empresa=?, nom_fiscal=?, fiscal_id=?, empresa_adreca=?, empresa_tel=?, margins_json=?, brand_color=?, brand_color_secondary=?, brand_color_menu=? WHERE id=?',
        [
            margins['frames'],
            margins['prints'],
            ne,
            nf,
            fi,
            ea,
            et,
            json.dumps(margins, ensure_ascii=True),
            brand_color,
            brand_color_secondary,
            brand_color_menu,
            session['user_id'],
        ]
    )
    if ne: session['empresa_nom'] = ne
    session['brand_color'] = brand_color
    session['brand_color_secondary'] = brand_color_secondary
    session['brand_color_menu'] = brand_color_menu
    _sync_private_commercial_settings(margins['frames'], margins['prints'], margins=margins)
    return jsonify({'ok': True, 'margins': margins})

# â"€â"€ Routes: Guardar comanda i historial â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
@app.route('/guardar', methods=['POST'])
@login_required
def guardar():
    d = request.json
    comanda_id = d.get('comanda_id')  # if set → UPDATE existing row

    # Snapshot dels marges del professional en el moment de guardar.
    # marge_pro_snap congela el % aplicat perquè futurs canvis a usuaris.marge_pro_pct
    # no alterin la reimpressió de pressupostos antics.
    marge_pro_actiu = get_config_value('marge_pro_actiu', '1') == '1'
    usuari = query('SELECT marge_pro_pct, marge FROM usuaris WHERE id=?', [session['user_id']], one=True)
    if marge_pro_actiu:
        marge_pro_snap = _get_marge_value(usuari)
    else:
        marge_pro_snap = 0.0

    qty = float(d.get('quantitat') or 1) or 1.0
    cost_produccio = float(d.get('cost_produccio') or 0)
    # pvd_unitari: cost per unit amb marge admin aplicat (el que retorna /api/closest com a 'pvd').
    # cost_produccio ja ve amb marge admin aplicat (suma de CLOSEST[k].preu × qty), per tant:
    pvd_unitari = round(cost_produccio / qty, 4) if qty else 0.0
    # cost_unitari: cost workshop sense marge admin. Requereix desglòs per categoria que el
    # payload actual no envia; queda NULL fins que el frontend enviï el detall.
    cost_unitari = None

    # Common field values
    vals_comuns = [
        d.get('client_nom',''), d.get('client_tel',''),
        d.get('pre_marc',''), d.get('marc_principal',''),
        d.get('amplada',0), d.get('alcada',0), d.get('copia',0),
        d.get('encolat',''), d.get('vidre',''), d.get('passpartout',''),
        d.get('passpartu_ref',''), d.get('impressio',''),
        1 if d.get('revers_peu') else 0, d.get('revers_peu_preu', 0),
        d.get('tipus_peca','fotografia'), d.get('tipus_peca_detall',''),
        d.get('final_amplada',0), d.get('final_alcada',0),
        d.get('marge',60), d.get('descompte',0), d.get('quantitat',1),
        d.get('preu_net',0), d.get('preu_final',0),
        cost_produccio,
        d.get('entrega',0), d.get('pendent',0),
        d.get('observacions',''), d.get('opcio_nom','Opció A'), d.get('lang','ca'),
        cost_unitari, pvd_unitari, marge_pro_snap,
    ]

    if comanda_id:
        # Verify ownership before updating
        existing = query('SELECT id, num_pressupost, sessio_id FROM comandes WHERE id=? AND user_id=?',
                         [comanda_id, session['user_id']], one=True)
        if existing:
            execute('''UPDATE comandes SET
                client_nom=?, client_tel=?, pre_marc=?, marc_principal=?,
                amplada=?, alcada=?, copia=?,
                encolat=?, vidre=?, passpartout=?, passpartu_ref=?, impressio=?,
                revers_peu=?, revers_peu_preu=?,
                tipus_peca=?, tipus_peca_detall=?, final_amplada=?, final_alcada=?,
                marge=?, descompte=?, quantitat=?,
                preu_net=?, preu_final=?, cost_produccio=?,
                entrega=?, pendent=?, observacions=?, opcio_nom=?, lang=?,
                cost_unitari=?, pvd_unitari=?, marge_pro_snap=?
                WHERE id=? AND user_id=?''',
                vals_comuns + [comanda_id, session['user_id']])
            return jsonify({'ok': True, 'id': existing['id'],
                            'sessio_id': existing['sessio_id'],
                            'num': existing['num_pressupost']})
        # If not found (wrong user or deleted), fall through to INSERT

    sessio_id = d.get('sessio_id') or secrets.token_hex(8)
    num_pressupost = generar_num_pressupost()
    cid = execute('''INSERT INTO comandes
        (user_id, data, client_nom, client_tel,
         pre_marc, marc_principal, amplada, alcada, copia,
         encolat, vidre, passpartout, passpartu_ref, impressio,
         revers_peu, revers_peu_preu,
         tipus_peca, tipus_peca_detall, final_amplada, final_alcada,
         marge, descompte, quantitat,
         preu_net, preu_final, cost_produccio, entrega, pendent, observacions,
         sessio_id, opcio_nom, num_pressupost, lang,
         cost_unitari, pvd_unitari, marge_pro_snap)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',
        [session['user_id'], datetime.now().strftime('%d/%m/%Y %H:%M')] +
        vals_comuns[:27] + [sessio_id] + [vals_comuns[27]] + [num_pressupost] + [vals_comuns[28]] +
        vals_comuns[29:32]
    )
    return jsonify({'ok': True, 'id': cid, 'sessio_id': sessio_id, 'num': num_pressupost})



@app.route('/sessio/<sessio_id>/pagat', methods=['POST'])
@login_required
def marcar_pagat(sessio_id):
    c = _get_comanda_by_sessio_for_session(sessio_id)
    if not c:
        return jsonify({'ok': False, 'error': 'No autoritzat'}), 403
    pagat = (request.json or {}).get('pagat', 1)
    execute('UPDATE comandes SET pagat=? WHERE sessio_id=?', [pagat, sessio_id])
    return jsonify({'ok': True})

@app.route('/sessio/<sessio_id>/entregat', methods=['POST'])
@login_required
def marcar_entregat(sessio_id):
    c = _get_comanda_by_sessio_for_session(sessio_id)
    if not c:
        return jsonify({'ok': False, 'error': 'No autoritzat'}), 403
    entregat = (request.json or {}).get('entregat', 1)
    execute('UPDATE comandes SET entregat=? WHERE sessio_id=?', [entregat, sessio_id])
    return jsonify({'ok': True})

@app.route('/comanda/<int:cid>/liquidar', methods=['POST'])
@login_required
def liquidar_comanda(cid):
    c = _get_comanda_for_session(cid, fields='id, user_id')
    if not c:
        return jsonify({'ok': False, 'error': 'No autoritzat'}), 403
    execute('UPDATE comandes SET entrega=preu_final, pendent=0, pagat=1 WHERE id=?', [cid])
    return jsonify({'ok': True})


@app.route('/admin/eliminar-tot', methods=['POST'])
@admin_required
def eliminar_tot():
    uid = request.json.get('user_id')
    sessio_ids = request.json.get('sessio_ids', [])
    if sessio_ids:
        for sid in sessio_ids:
            execute('DELETE FROM comandes WHERE sessio_id=?', [sid])
    elif uid:
        execute('DELETE FROM comandes WHERE user_id=?', [uid])
    return jsonify({'ok': True})

@app.route('/comanda/<int:cid>/eliminar', methods=['POST'])
@login_required
def eliminar_comanda(cid):
    c = _get_comanda_for_session(cid, fields='id, user_id')
    if not c:
        return jsonify({'ok': False, 'error': 'No autoritzat'})
    execute('DELETE FROM comandes WHERE id=?', [cid])
    return jsonify({'ok': True})

@app.route('/comanda/<int:cid>/acceptar', methods=['POST'])
@login_required
def acceptar_comanda(cid):
    c = _get_comanda_for_session(cid, fields='id, user_id')
    if not c:
        return jsonify({'ok': False, 'error': 'No autoritzat'}), 403
    estat = (request.json or {}).get('estat', 'acceptat')
    execute('UPDATE comandes SET observacions = CASE WHEN observacions IS NULL OR observacions=\'\' THEN ? ELSE observacions || \' | \' || ? END WHERE id=?',
            [f'[{estat.upper()}]', f'[{estat.upper()}]', cid])
    return jsonify({'ok': True})


# ── Catàleg de motllures ──────────────────────────────────────────────────
@app.route('/cataleg')
@login_required
def cataleg():
    q = request.args.get('q', '').strip()
    color = request.args.get('color', '').strip().lower()
    gruix = request.args.get('gruix', '').strip().lower()
    sql = """SELECT referencia, gruix, descripcio, foto, ref2
             FROM moldures"""
    args = []
    if q:
        sql += """ WHERE LOWER(referencia) LIKE ?
                   OR LOWER(COALESCE(descripcio, '')) LIKE ?"""
        args = [f'%{q.lower()}%', f'%{q.lower()}%']
    sql += " ORDER BY referencia"
    moldures = _serialize_moldures(query(sql, args))
    moldures = [
        m for m in moldures
        if _matches_moldura_color(m.get('descripcio', ''), color)
        and _matches_moldura_gruix(m.get('gruix', 0), gruix)
    ]
    total = query("SELECT COUNT(*) as n FROM moldures", one=True)
    return render_template('cataleg.html', moldures=moldures, q=q,
                           total=total['n'] if total else 0,
                           color=color, gruix=gruix,
                           color_filters=MOLDURA_COLOR_FILTERS,
                           gruix_filters=MOLDURA_GRUIX_FILTERS)

@app.route('/admin/run-migrations')
@admin_required
def admin_run_migrations():
    """TEMPORAL: únic punt d'inicialització/migració de la BD.

    Es crida manualment per un admin després de cada deploy amb canvis de schema.
    Fa en ordre:
      1. init_db() — CREATE TABLE + ALTER ADD COLUMN + seeds de config + admin bootstrap
      2. ALTER TABLE ADD COLUMN IF NOT EXISTS redundants (safety net)
      3. Seeds i backfill pesats (ProEco, Intermol cleanup, v2 price backfill)
      4. Conversions de tipus per a DBs desplegades abans (cost_verificat BOOLEAN→INTEGER)
      5. Renames de columnes històriques (historial_preus_cost)

    Cada pas amb try/except + rollback individual perquè una fallada no bloqui la resta.
    Eliminar aquesta ruta un cop la BD estigui alineada a totes les instàncies."""

    resultats = []
    db = get_db()

    # 1) init_db() (CREATE TABLE + ALTER + seeds lleugers)
    try:
        init_db()
        resultats.append("OK: init_db()")
    except Exception as e:
        try:
            db.rollback()
        except Exception:
            pass
        resultats.append(f"SKIP init_db(): {str(e)[:120]}")

    alteracions = [
        # Moldures
        "ALTER TABLE moldures ADD COLUMN IF NOT EXISTS preu_cost DECIMAL(8,4)",
        "ALTER TABLE moldures ADD COLUMN IF NOT EXISTS merma_pct DECIMAL(4,2) DEFAULT 10.00",
        "ALTER TABLE moldures ADD COLUMN IF NOT EXISTS minim_cm DECIMAL(6,1) DEFAULT 100.0",
        "ALTER TABLE moldures ADD COLUMN IF NOT EXISTS preu_cost_ant DECIMAL(8,4)",
        "ALTER TABLE moldures ADD COLUMN IF NOT EXISTS data_cost DATE",
        "ALTER TABLE moldures ADD COLUMN IF NOT EXISTS usuari_cost_id INTEGER",
        "ALTER TABLE moldures ADD COLUMN IF NOT EXISTS notes_cost TEXT",
        "ALTER TABLE moldures ADD COLUMN IF NOT EXISTS cost_verificat INTEGER DEFAULT 0",
        # Vidres
        "ALTER TABLE vidres ADD COLUMN IF NOT EXISTS preu_cost DECIMAL(8,4)",
        "ALTER TABLE vidres ADD COLUMN IF NOT EXISTS preu_cost_ant DECIMAL(8,4)",
        "ALTER TABLE vidres ADD COLUMN IF NOT EXISTS data_cost DATE",
        "ALTER TABLE vidres ADD COLUMN IF NOT EXISTS usuari_cost_id INTEGER",
        "ALTER TABLE vidres ADD COLUMN IF NOT EXISTS notes_cost TEXT",
        "ALTER TABLE vidres ADD COLUMN IF NOT EXISTS cost_verificat INTEGER DEFAULT 0",
        # Encolat
        "ALTER TABLE encolat_pro ADD COLUMN IF NOT EXISTS preu_cost DECIMAL(8,4)",
        "ALTER TABLE encolat_pro ADD COLUMN IF NOT EXISTS preu_cost_ant DECIMAL(8,4)",
        "ALTER TABLE encolat_pro ADD COLUMN IF NOT EXISTS data_cost DATE",
        "ALTER TABLE encolat_pro ADD COLUMN IF NOT EXISTS usuari_cost_id INTEGER",
        "ALTER TABLE encolat_pro ADD COLUMN IF NOT EXISTS notes_cost TEXT",
        "ALTER TABLE encolat_pro ADD COLUMN IF NOT EXISTS cost_verificat INTEGER DEFAULT 0",
        # Passpartout
        "ALTER TABLE passpartout ADD COLUMN IF NOT EXISTS preu_cost DECIMAL(8,4)",
        "ALTER TABLE passpartout ADD COLUMN IF NOT EXISTS preu_cost_ant DECIMAL(8,4)",
        "ALTER TABLE passpartout ADD COLUMN IF NOT EXISTS data_cost DATE",
        "ALTER TABLE passpartout ADD COLUMN IF NOT EXISTS usuari_cost_id INTEGER",
        "ALTER TABLE passpartout ADD COLUMN IF NOT EXISTS notes_cost TEXT",
        "ALTER TABLE passpartout ADD COLUMN IF NOT EXISTS cost_verificat INTEGER DEFAULT 0",
        # Comandes
        "ALTER TABLE comandes ADD COLUMN IF NOT EXISTS cost_unitari DECIMAL(8,4)",
        "ALTER TABLE comandes ADD COLUMN IF NOT EXISTS pvd_unitari DECIMAL(8,4)",
        "ALTER TABLE comandes ADD COLUMN IF NOT EXISTS marge_admin_snap DECIMAL(5,2)",
        "ALTER TABLE comandes ADD COLUMN IF NOT EXISTS marge_pro_snap DECIMAL(5,2)",
        # Usuaris
        "ALTER TABLE usuaris ADD COLUMN IF NOT EXISTS marge_pro_pct DECIMAL(5,2)",
        "ALTER TABLE usuaris ADD COLUMN IF NOT EXISTS marge_impressio_pro_pct DECIMAL(5,2)",
        # Historial de preus de cost — noms de columna alineats amb el codi
        # existent (taula, preu_cost_antic, data, notes).
        """CREATE TABLE IF NOT EXISTS historial_preus_cost (
            id SERIAL PRIMARY KEY,
            taula VARCHAR(50) NOT NULL,
            referencia VARCHAR(50) NOT NULL,
            preu_cost_antic DECIMAL(8,4) NOT NULL,
            preu_cost_nou DECIMAL(8,4) NOT NULL,
            data TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
            usuari_id INTEGER NOT NULL,
            notes TEXT
        )""",
    ]

    resultats = []
    db = get_db()
    for sql in alteracions:
        try:
            execute(sql)
            resultats.append(f"OK: {sql[:80].strip()}…")
        except Exception as e:
            # A PG, una sentència fallida deixa la transacció en estat aborted
            # i tota la resta del loop fallaria. Rollback explícit per recuperar.
            try:
                db.rollback()
            except Exception:
                pass
            resultats.append(f"SKIP: {str(e)[:120]}")

    # Operacions pesades: totes mogudes aquí des de init_db() per no penjar
    # l'arrencada dels workers a PG. Cadascuna amb el seu try/except + rollback.
    for nom, fn in [
        ("_seed_proeco_preus",      lambda: _seed_proeco_preus(db, use_pg=USE_PG)),
        ("_seed_intermol_moldures", lambda: _seed_intermol_moldures(db, use_pg=USE_PG)),
        ("_fix_ref2_errors",        lambda: _fix_ref2_errors(db, use_pg=USE_PG)),
        ("_run_v2_price_backfill",  lambda: _run_v2_price_backfill(db)),
    ]:
        try:
            fn()
            try:
                db.commit()
            except Exception:
                pass
            resultats.append(f"OK: {nom}")
        except Exception as e:
            try:
                db.rollback()
            except Exception:
                pass
            resultats.append(f"SKIP {nom}: {str(e)[:120]}")

    # Catch-up migrations per a DBs desplegades amb schema intermedi.
    # Només s'executen a PG (SQLite no té tipus estrictes i aquestes fallen
    # allà però el SKIP és inofensiu).
    if USE_PG:
        # BOOLEAN → INTEGER: a PG cal DROP DEFAULT, ALTER TYPE amb USING, SET DEFAULT.
        # Si la columna ja és INTEGER les sentències fallen i es registren com a SKIP.
        bool_to_int = []
        for t in ('moldures', 'vidres', 'encolat_pro', 'passpartout'):
            bool_to_int += [
                f"ALTER TABLE {t} ALTER COLUMN cost_verificat DROP DEFAULT",
                f"ALTER TABLE {t} ALTER COLUMN cost_verificat TYPE INTEGER "
                f"USING (CASE WHEN cost_verificat THEN 1 ELSE 0 END)",
                f"ALTER TABLE {t} ALTER COLUMN cost_verificat SET DEFAULT 0",
            ]
        for sql in bool_to_int:
            try:
                execute(sql)
                resultats.append(f"OK: {sql[:80].strip()}…")
            except Exception as e:
                try:
                    db.rollback()
                except Exception:
                    pass
                resultats.append(f"SKIP: {str(e)[:120]}")

        # historial_preus_cost: renoms per alinear noms antics → nous.
        renames = [
            "ALTER TABLE historial_preus_cost RENAME COLUMN taula_origen TO taula",
            "ALTER TABLE historial_preus_cost RENAME COLUMN preu_cost_ant TO preu_cost_antic",
            "ALTER TABLE historial_preus_cost RENAME COLUMN data_canvi TO data",
            "ALTER TABLE historial_preus_cost RENAME COLUMN motiu TO notes",
        ]
        for sql in renames:
            try:
                execute(sql)
                resultats.append(f"OK: {sql[:80].strip()}…")
            except Exception as e:
                try:
                    db.rollback()
                except Exception:
                    pass
                resultats.append(f"SKIP: {str(e)[:120]}")

    return "<br>".join(resultats)


@app.route('/admin/db-status')
def admin_db_status():
    """TEMPORAL: diagnòstic complet de l'estat de la BD.

    Accés permès si es compleix qualsevol de les dues condicions:
      1. Sessió d'admin al cookie (funcionament normal).
      2. Paràmetre ?token=<valor> que coincideixi amb l'env var DB_STATUS_TOKEN.

    La via 2 existeix per poder usar la ruta quan l'auth o la BD estan
    trencades. Configura DB_STATUS_TOKEN a Railway i visita-la així:
        /admin/db-status?token=<valor>

    Si DB_STATUS_TOKEN no està definida, només serveix via admin session.
    Eliminar aquesta ruta un cop la BD estigui estable."""
    token_env = (os.environ.get('DB_STATUS_TOKEN') or '').strip()
    token_req = (request.args.get('token') or '').strip()
    is_admin  = bool(session.get('is_admin'))
    token_ok  = bool(token_env) and hmac.compare_digest(token_env, token_req)
    if not (is_admin or token_ok):
        return 'No autoritzat', 403

    import traceback
    try:
        result = {}

        if USE_PG:
            taules_sql = (
                "SELECT table_name FROM information_schema.tables "
                "WHERE table_schema='public' ORDER BY table_name"
            )
        else:
            taules_sql = (
                "SELECT name AS table_name FROM sqlite_master "
                "WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name"
            )
        taules_rows = query(taules_sql) or []
        result['taules'] = [r['table_name'] for r in taules_rows]

        # Comptar files per taula (protegit per si alguna fa error)
        counts = {}
        for t in result['taules']:
            try:
                r = query(f"SELECT COUNT(*) AS n FROM {t}", one=True)
                counts[t] = r['n'] if r else 0
            except Exception as e:
                counts[t] = f'error: {str(e)[:60]}'
        result['counts'] = counts

        def _cols(taula):
            if USE_PG:
                rows = query(
                    "SELECT column_name, data_type FROM information_schema.columns "
                    "WHERE table_name=%s ORDER BY ordinal_position",
                    [taula],
                ) or []
                return [{'nom': r['column_name'], 'tipus': r['data_type']} for r in rows]
            else:
                rows = query(f'PRAGMA table_info({taula})') or []
                return [{'nom': r['name'], 'tipus': r['type']} for r in rows]

        result['usuaris_columnes']  = _cols('usuaris')
        result['moldures_columnes'] = _cols('moldures')
        result['comandes_columnes'] = _cols('comandes')
        result['historial_columnes'] = _cols('historial_preus_cost')

        try:
            result['usuaris_mostra'] = query(
                'SELECT id, username, nom, is_admin, marge, '
                'marge_pro_pct, marge_impressio_pro_pct FROM usuaris LIMIT 3'
            ) or []
        except Exception as e:
            result['usuaris_mostra'] = f'error: {str(e)[:120]}'

        try:
            cfg_rows = query(
                "SELECT clau, valor FROM config WHERE clau IN ("
                "'marge_pro_actiu','marge_defecte','migration_v2_done',"
                "'marge_admin_moldures_pct','cost_hora_taller')"
            ) or []
            result['config'] = {r['clau']: r['valor'] for r in cfg_rows}
        except Exception as e:
            result['config'] = f'error: {str(e)[:120]}'

        result['use_pg'] = USE_PG
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500


@app.route('/admin/debug-fotos')
@admin_required
def admin_debug_fotos():
    """Mostra per a cada motllura: referencia, ref2, foto resolta i si existeix el fitxer."""
    rows = query('SELECT referencia, ref2, foto, proveidor FROM moldures ORDER BY referencia')
    out = []
    for r in (rows or []):
        ref  = _row_get(r, 'referencia', '') or ''
        ref2 = _row_get(r, 'ref2', '') or ''
        foto = _row_get(r, 'foto', '') or ''
        prov = _row_get(r, 'proveidor', '') or ''
        resolved = _resolve_moldura_photo(ref, foto, ref2=ref2)
        out.append({'ref': ref, 'ref2': ref2, 'proveidor': prov,
                    'foto_db': foto, 'foto_resolta': resolved})
    # Return as plain text table for easy reading
    lines = ['referencia | ref2 | proveidor | foto_resolta']
    lines.append('-' * 80)
    no_foto = []
    for o in out:
        if o['foto_resolta']:
            lines.append(f"OK  {o['ref']:<25} ref2={o['ref2']:<20} -> {o['foto_resolta']}")
        else:
            no_foto.append(o)
    lines.append('')
    lines.append(f'--- SENSE FOTO ({len(no_foto)}) ---')
    for o in no_foto:
        lines.append(f"    {o['ref']:<25} ref2={o['ref2']:<20} prov={o['proveidor']}")
    lines.append('')
    lines.append(f'Total: {len(out)} motllures, amb foto: {len(out)-len(no_foto)}, sense: {len(no_foto)}')
    return '<pre>' + '\n'.join(lines) + '</pre>'


@app.route('/admin/cataleg')
@admin_required
def admin_cataleg():
    q = request.args.get('q', '').strip()
    proveidor = request.args.get('proveidor', '').strip()
    moldures = _serialize_moldures(_query_moldures(q=q, proveidor=proveidor))
    proveidors = query("SELECT DISTINCT proveidor FROM moldures WHERE proveidor!='' ORDER BY proveidor")
    total = query("SELECT COUNT(*) as n FROM moldures", one=True)
    pdf_params = {}
    if q:
        pdf_params['q'] = q
    if proveidor:
        pdf_params['proveidor'] = proveidor
    pdf_url = url_for('admin_cataleg_pdf')
    if pdf_params:
        pdf_url += '?' + urlencode(pdf_params)
    return render_template('admin_cataleg.html', moldures=moldures, q=q,
                           proveidor=proveidor, proveidors=proveidors,
                           total=total['n'] if total else 0,
                           pdf_url=pdf_url)


@app.route('/admin/cataleg/pdf')
@admin_required
def admin_cataleg_pdf():
    q = request.args.get('q', '').strip()
    proveidor = request.args.get('proveidor', '').strip()
    moldures = _serialize_moldures(_query_moldures(q=q, proveidor=proveidor))
    pdf = crear_pdf_cataleg_admin(moldures, q=q, proveidor=proveidor)
    return send_file(pdf, mimetype='application/pdf', as_attachment=False,
                     download_name='cataleg-motllures-admin.pdf')

@app.route('/admin/cataleg/nou', methods=['GET','POST'])
@admin_required
def admin_moldura_nova():
    if request.method == 'POST':
        d = request.form
        existing = query('SELECT referencia FROM moldures WHERE referencia=?', [d['referencia']], one=True)
        if existing:
            return render_template('admin_moldura_form.html', error='Ja existeix una motllura amb aquesta referència.', moldura=d, nova=True)
        foto = d.get('foto', '').strip()
        try:
            uploaded_photo = _save_moldura_photo(request.files.get('foto_fitxer'), d['referencia'])
        except ValueError as e:
            return render_template('admin_moldura_form.html', error=str(e), moldura=d, nova=True)
        if uploaded_photo:
            foto = uploaded_photo
        execute("""INSERT INTO moldures (referencia,preu_taller,gruix,cost,proveidor,ref2,ubicacio,descripcio,foto)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                [d['referencia'], float(d.get('preu_taller',0)),
                 float(d.get('gruix',0)), float(d.get('cost',0)),
                 d.get('proveidor',''), d.get('ref2',''),
                 d.get('ubicacio',''), d.get('descripcio',''), foto])
        return redirect(url_for('admin_cataleg'))
    return render_template('admin_moldura_form.html', moldura={}, nova=True, error=None)

@app.route('/admin/cataleg/<ref>/editar', methods=['GET','POST'])
@admin_required
def admin_moldura_editar(ref):
    moldura = query('SELECT * FROM moldures WHERE referencia=?', [ref], one=True)
    if not moldura:
        return redirect(url_for('admin_cataleg'))
    if request.method == 'POST':
        d = request.form
        foto = d.get('foto', '').strip()
        try:
            uploaded_photo = _save_moldura_photo(request.files.get('foto_fitxer'), ref)
        except ValueError as e:
            moldura_data = dict(moldura)
            moldura_data.update(d)
            moldura_data['foto'] = foto or _resolve_moldura_photo(ref, _row_get(moldura, 'foto', ''))
            return render_template('admin_moldura_form.html', error=str(e), moldura=moldura_data, nova=False)
        if uploaded_photo:
            foto = uploaded_photo
        execute("""UPDATE moldures SET preu_taller=?,gruix=?,cost=?,proveidor=?,
                   ref2=?,ubicacio=?,descripcio=?,foto=? WHERE referencia=?""",
                [float(d.get('preu_taller',0)), float(d.get('gruix',0)),
                 float(d.get('cost',0)), d.get('proveidor',''),
                 d.get('ref2',''), d.get('ubicacio',''),
                 d.get('descripcio',''), foto, ref])
        return redirect(url_for('admin_cataleg'))
    return render_template('admin_moldura_form.html', moldura=_serialize_moldura(moldura), nova=False, error=None)

@app.route('/admin/cataleg/<ref>/eliminar', methods=['POST'])
@admin_required
def admin_moldura_eliminar(ref):
    execute('DELETE FROM moldures WHERE referencia=?', [ref])
    return jsonify({'ok': True})

@app.route('/admin/cataleg/<ref>/toggle-actiu', methods=['POST'])
@admin_required
def admin_moldura_toggle(ref):
    m = query('SELECT actiu FROM moldures WHERE referencia=?', [ref], one=True)
    nou = 0 if (m and _row_get(m, 'actiu', 1)) else 1
    try:
        execute('UPDATE moldures SET actiu=? WHERE referencia=?', [nou, ref])
    except:
        pass
    return jsonify({'ok': True, 'actiu': nou})

# â"€â"€ API: buscar moldura per ref exacte (autocomplete) â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
@app.route('/admin/cataleg/api/cerca')
@admin_required
def api_cerca_moldura():
    q = request.args.get('q','').strip()
    if not q: return jsonify([])
    rows = query("""SELECT referencia, preu_taller, gruix, descripcio, proveidor 
                    FROM moldures WHERE LOWER(referencia) LIKE ? OR LOWER(descripcio) LIKE ?
                    ORDER BY referencia LIMIT 20""",
                 [f'%{q.lower()}%', f'%{q.lower()}%'])
    return jsonify([dict(r) for r in rows])

@app.route('/admin/taules-preus')
@admin_required
def admin_taules_preus():
    """Mostra totes les taules de preus i simula lookups per mides personalitzades."""
    w = request.args.get('w', 0, type=float)
    h = request.args.get('h', 0, type=float)

    vidres_rows   = [dict(r) for r in query('SELECT referencia, preu FROM vidres ORDER BY referencia')   or []]
    encolat_rows  = [dict(r) for r in query('SELECT referencia, preu FROM encolat_pro ORDER BY referencia') or []]

    lookup = {}
    if w > 0 and h > 0:
        def sim(rows, prefix=None):
            r = _find_closest(rows, w, h, prefix=prefix)
            return {'ref': r['referencia'], 'preu': r['preu']} if r else None
        lookup = {
            'vidre':       sim([r for r in vidres_rows if not r['referencia'].upper().startswith(('DV-','MIR-'))]),
            'doble_vidre': sim(vidres_rows, prefix='DV-'),
            'mirall':      sim(vidres_rows, prefix='MIR-'),
            'encolat':     sim(encolat_rows, prefix='ENC'),
            'protter':     sim(encolat_rows, prefix='PRO'),
        }

    def rows_html(rows, highlight_ref=None):
        lines = ['<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-size:13px">',
                 '<tr><th>Referència</th><th>Preu (€)</th></tr>']
        for r in rows:
            bg = ' style="background:#ffe082"' if highlight_ref and r['referencia'] == highlight_ref else ''
            lines.append(f'<tr{bg}><td>{r["referencia"]}</td><td>{r["preu"]}</td></tr>')
        lines.append('</table>')
        return '\n'.join(lines)

    vid_hl   = lookup.get('vidre',   {}).get('ref') if lookup else None
    enc_hl   = lookup.get('encolat', {}).get('ref') if lookup else None
    pro_hl   = lookup.get('protter', {}).get('ref') if lookup else None

    html = f"""<!DOCTYPE html><html><head><meta charset="UTF-8">
<title>Taules de preus — Admin</title>
<style>body{{font-family:Arial,sans-serif;padding:2rem;background:#f5f6fa}}
h2{{margin-top:2rem}}
.lookup{{background:#e8f5e9;border:1px solid #a5d6a7;padding:1rem;border-radius:8px;margin-bottom:1.5rem;font-size:14px}}
.lookup b{{color:#1b5e20}}</style></head><body>
<h1>Taules de preus</h1>
<form method="get">
  Mida marc: <input name="w" value="{w or ''}" placeholder="Amplada" style="width:70px"> ×
             <input name="h" value="{h or ''}" placeholder="Alçada"  style="width:70px"> cm
  <button type="submit">Simular lookup</button>
</form>
"""
    if lookup:
        def lrow(k, label):
            v = lookup.get(k)
            return f'<li><b>{label}:</b> {v["ref"]} → {v["preu"]}€</li>' if v else f'<li><b>{label}:</b> no trobat</li>'
        html += f'<div class="lookup"><b>Lookup per {w}×{h} cm:</b><ul>'
        html += lrow('vidre',       'Vidre')
        html += lrow('doble_vidre', 'Doble vidre')
        html += lrow('mirall',      'Mirall')
        html += lrow('encolat',     'Encolat')
        html += lrow('protter',     'Protter')
        html += '</ul></div>'

    html += f'<h2>Vidres ({len(vidres_rows)} files)</h2>' + rows_html(vidres_rows, vid_hl)
    html += f'<h2>Encolat / Protter ({len(encolat_rows)} files)</h2>' + rows_html(encolat_rows, enc_hl or pro_hl)
    html += '</body></html>'
    return html


@app.route('/historial')
@login_required
def historial():
    filtre_uid_raw = request.args.get('user_id', '').strip()
    filtre_uid = int(filtre_uid_raw) if filtre_uid_raw.isdigit() else None
    filtre_all  = filtre_uid_raw == 'all'
    filtre_albara = request.args.get('albara') == 'pendent'
    if session.get('is_admin'):
        if filtre_albara:
            comandes = query('''SELECT c.*, u.nom as usuari_nom FROM comandes c
                               JOIN usuaris u ON c.user_id=u.id
                               WHERE c.observacions LIKE '%[ACCEPTAT]%'
                                 AND (c.fd_albara IS NULL OR c.fd_albara='')
                               ORDER BY c.id DESC''')
        elif filtre_all:
            comandes = query('''SELECT c.*, u.nom as usuari_nom FROM comandes c
                               JOIN usuaris u ON c.user_id=u.id
                               ORDER BY c.id DESC''')
        elif filtre_uid:
            comandes = query('''SELECT c.*, u.nom as usuari_nom FROM comandes c
                               JOIN usuaris u ON c.user_id=u.id
                               WHERE c.user_id=? ORDER BY c.id DESC''', [filtre_uid])
        else:
            # Per defecte l'admin veu les seves pròpies comandes
            filtre_uid = session['user_id']
            comandes = query('''SELECT c.*, u.nom as usuari_nom FROM comandes c
                               JOIN usuaris u ON c.user_id=u.id
                               WHERE c.user_id=? ORDER BY c.id DESC''', [filtre_uid])
        usuaris_list = query('SELECT id, nom, username FROM usuaris WHERE is_admin=0 ORDER BY nom')
        n_pendents_albara = query('''SELECT COUNT(*) as n FROM comandes
                                    WHERE observacions LIKE '%[ACCEPTAT]%'
                                      AND (fd_albara IS NULL OR fd_albara='')''', one=True)
        n_pendents_albara = n_pendents_albara['n'] if n_pendents_albara else 0
    else:
        usuaris_list = []
        comandes = query('''SELECT c.*, u.nom as usuari_nom FROM comandes c
                           JOIN usuaris u ON c.user_id=u.id
                           WHERE c.user_id=? ORDER BY c.id DESC''', [session['user_id']])
    # Group by sessio_id
    sessions = {}
    for c in comandes:
        sid = c['sessio_id'] or str(c['id'])
        if sid not in sessions:
            sessions[sid] = []
        d = dict(c)
        sessions[sid].append(d)
    sessio_list = list(sessions.values())
    # Add pagat/entregat flag to first item of each session
    for grp in sessio_list:
        grp[0]['pagat']    = any(op.get('pagat')    for op in grp)
        grp[0]['entregat'] = any(op.get('entregat') for op in grp)
    return render_template('historial.html', comandes=comandes, sessio_list=sessio_list,
                           usuaris_list=usuaris_list if session.get('is_admin') else [],
                           filtre_uid=filtre_uid, filtre_all=filtre_all if session.get('is_admin') else False,
                           filtre_albara=filtre_albara if session.get('is_admin') else False,
                           n_pendents_albara=n_pendents_albara if session.get('is_admin') else 0,
                           web_return_url=_current_web_return_url())

@app.route('/pdf-comparativa/<sessio_id>')
@login_required
def pdf_comparativa(sessio_id):
    comandes = query('''SELECT * FROM comandes WHERE sessio_id=? ORDER BY id''', [sessio_id])
    if not comandes:
        return 'No trobat', 404
    if not session.get('is_admin') and comandes[0]['user_id'] != session['user_id']:
        return 'No autoritzat', 403
    pdf = crear_pdf_comparativa([dict(c) for c in comandes])
    nom = comandes[0]['client_nom'] or 'comparativa'
    return send_file(pdf, mimetype='application/pdf',
                     download_name=f"comparativa_{nom}.pdf")

@app.route('/pdf/<int:comanda_id>')
@login_required
def generar_pdf(comanda_id):
    c = query('SELECT * FROM comandes WHERE id=?', [comanda_id], one=True)
    if not c: return 'No trobat', 404
    if not session.get('is_admin') and c['user_id'] != session['user_id']:
        return 'No autoritzat', 403
    pdf = crear_pdf(dict(c))
    return send_file(pdf, mimetype='application/pdf',
                     download_name=f"pressupost_{c['client_nom']}_{comanda_id}.pdf")

# â"€â"€ Routes: Admin â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
@app.route('/admin')
@admin_required
def admin():
    usuaris = query('SELECT * FROM usuaris ORDER BY nom')
    config = {r['clau']: r['valor'] for r in query('SELECT * FROM config')}
    impressio = query('SELECT * FROM impressio ORDER BY preu')
    passpartous = query('SELECT referencia, color, textura, descripcio FROM passpartout ORDER BY referencia') or []
    proeco_rows = query('SELECT referencia, preu FROM proeco ORDER BY referencia') or []
    return render_template('admin.html', usuaris=usuaris, config=config, impressio=impressio,
                           passpartous=passpartous, proeco_rows=proeco_rows)

@app.route('/admin/usuari', methods=['POST'])
@admin_required
def admin_usuari():
    action = request.form.get('action')
    if action == 'crear':
        execute('INSERT INTO usuaris (username, password, nom, is_admin, access_status, profile_type, web_url, instagram, fiscal_id, notes_validacio) VALUES (?,?,?,?,?,?,?,?,?,?)',
                [request.form['username'], hash_pw(request.form['password']),
                 request.form['nom'], int(request.form.get('is_admin', 0)),
                 request.form.get('access_status', 'active'),
                 request.form.get('profile_type', 'professional'),
                 request.form.get('web_url', '').strip(),
                 request.form.get('instagram', '').strip(),
                 request.form.get('fiscal_id', '').strip(),
                 request.form.get('notes_validacio', '').strip()])
        flash(f"Usuari '{request.form['nom']}' creat.", 'ok')
    elif action == 'eliminar':
        execute('DELETE FROM usuaris WHERE id=?', [request.form['uid']])
        flash('Usuari eliminat.', 'ok')
    elif action == 'canviar_pw':
        execute('UPDATE usuaris SET password=? WHERE id=?',
                [hash_pw(request.form['password']), request.form['uid']])
        flash('Contrasenya actualitzada.', 'ok')
    elif action == 'actualitzar_estat':
        execute('UPDATE usuaris SET access_status=?, profile_type=?, web_url=?, instagram=?, fiscal_id=?, notes_validacio=? WHERE id=?',
                [request.form.get('access_status', 'active'),
                 request.form.get('profile_type', 'professional'),
                 request.form.get('web_url', '').strip(),
                 request.form.get('instagram', '').strip(),
                 request.form.get('fiscal_id', '').strip(),
                 request.form.get('notes_validacio', '').strip(),
                 request.form['uid']])
        flash('Perfil professional actualitzat.', 'ok')
    return redirect(url_for('admin'))

@app.route('/admin/config', methods=['POST'])
@admin_required
def admin_config():
    execute("UPDATE config SET valor=? WHERE clau='marge_defecte'",
            [request.form.get('marge', 60)])
    # Toggle marge pro
    mpa = '1' if request.form.get('marge_pro_actiu') else '0'
    execute("INSERT OR REPLACE INTO config (clau, valor) VALUES ('marge_pro_actiu', ?)", [mpa])
    if request.form.get('save_gmail'):
        gu = request.form.get('gmail_user','').strip()
        gp = request.form.get('gmail_pass','').strip().replace(' ','')
        if gu:
            execute('INSERT OR REPLACE INTO config (clau, valor) VALUES ("gmail_user", ?)', [gu])
        if gp:
            execute('INSERT OR REPLACE INTO config (clau, valor) VALUES ("gmail_pass", ?)', [gp])
    flash('Configuració desada.', 'ok')
    return redirect(url_for('admin'))

# ── Factura Directa ───────────────────────────────────────────────────────
_FD_TOKEN   = os.environ.get('FACTURADIRECTA_TOKEN', '')
_FD_COMPANY = os.environ.get('FACTURADIRECTA_COMPANY', '')
_FD_BASE    = 'https://app.facturadirecta.com/api'

def _fd_headers():
    return {
        'facturadirecta-api-key': _FD_TOKEN,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
    }

def _fd_get(path):
    url = f'{_FD_BASE}/{_FD_COMPANY}/{path}'
    req = urllib_request.Request(url, headers=_fd_headers())
    try:
        with urllib_request.urlopen(req, timeout=10) as r:
            return json.loads(r.read().decode())
    except urllib_error.HTTPError as e:
        return {'_error': e.code, '_msg': e.read().decode()}

def _fd_post(path, data):
    url = f'{_FD_BASE}/{_FD_COMPANY}/{path}'
    body = json.dumps(data).encode()
    req = urllib_request.Request(url, data=body, headers=_fd_headers(), method='POST')
    try:
        with urllib_request.urlopen(req, timeout=10) as r:
            return json.loads(r.read().decode())
    except urllib_error.HTTPError as e:
        return {'_error': e.code, '_msg': e.read().decode()}

def _fd_cerca_contacte(nom=None, nif=None):
    """Cerca un contacte a FD per NIF o per nom."""
    if nif:
        res = _fd_get(f'contacts?fiscalId={urllib_quote(nif)}')
        items = res.get('items') or res.get('data') or (res if isinstance(res, list) else [])
        if items:
            return items[0]
    if nom:
        res = _fd_get(f'contacts?search={urllib_quote(nom)}')
        items = res.get('items') or res.get('data') or (res if isinstance(res, list) else [])
        if items:
            return items[0]
    return None

def _fd_extract_contact_id(r):
    """Extreu l'ID d'un contacte de la resposta de l'API de FD."""
    if not r:
        return ''
    c = r.get('content') or {}
    m = c.get('main') or {}
    return (r.get('id') or r.get('uuid') or r.get('contactId') or r.get('_id') or
            c.get('uuid') or c.get('id') or c.get('contactId') or
            m.get('id') or m.get('uuid') or '')

def _fd_crear_contacte(nom, nif=None, telefon=None):
    main = {
        'name': nom, 'country': 'ES', 'currency': 'EUR',
        'accounts': {'client': '430000', 'clientCredit': '438000'},
    }
    if nif:     main['fiscalId'] = nif
    if telefon: main['phone']    = telefon
    res = _fd_post('contacts', {'content': {'type': 'contact', 'main': main}})
    if '_error' in (res or {}):
        return res
    # Si la resposta ja porta ID, la retornem directament
    if _fd_extract_contact_id(res):
        return res
    # Si no té ID però tenim NIF, busquem per NIF (cerca exacta, mai per nom)
    if nif:
        trobat = _fd_cerca_contacte(nif=nif)
        if trobat:
            return trobat
    # Retornem la resposta original (sense ID); l'error es gestionarà a dalt
    print(f'FD crear_contacte: resposta sense ID (nom={nom}, nif={nif}): {json.dumps(res, ensure_ascii=False)}')
    return res

def _fd_crear_albara(contact_id, linies, notes='', data_doc=None):
    if not data_doc:
        data_doc = datetime.now().strftime('%Y-%m-%d')
    main = {
        'contact':   contact_id,
        'currency':  'EUR',
        'baseState': 'pending',
        'docNumber': {'series': 'AL'},
        'lines':     linies,
    }
    if data_doc:
        main['date'] = data_doc
    if notes:
        main['notes'] = notes
    return _fd_post('deliveryNotes', {'content': {'type': 'deliveryNote', 'main': main}})

@app.route('/api/crear-albara', methods=['POST'])
@login_required
def api_crear_albara():
    if not _FD_TOKEN or not _FD_COMPANY:
        return jsonify({'ok': False, 'error': 'Factura Directa no configurat (variables d\'entorn)'}), 503

    d = request.get_json(force=True) or {}
    client_nom   = (d.get('client_nom') or '').strip()
    client_tel   = (d.get('client_tel') or '').strip()
    cost_prod    = float(d.get('cost_produccio') or 0)
    preu_net     = float(d.get('preu_net') or 0)
    preu_final   = float(d.get('preu_final') or 0)
    marc         = (d.get('marc') or '').strip()
    opcions_text = (d.get('opcions') or '').strip()
    observacions = (d.get('observacions') or '').strip()
    num_pressupost = (d.get('num_pressupost') or '').strip()
    quantitat    = float(d.get('quantitat') or 1)
    revers_peu   = bool(d.get('revers_peu'))
    mode_preu    = (d.get('mode_preu') or 'pvp').strip().lower()  # 'pvp' or 'cost'

    # Només l'admin (Reus Revela) pot crear albarans a FD
    user = query('SELECT nom, nom_empresa, nom_fiscal, fiscal_id, is_admin FROM usuaris WHERE id=?',
                 [session['user_id']], one=True)
    is_admin = bool(_row_get(user, 'is_admin', 0))
    if not is_admin:
        return jsonify({'ok': False, 'error': 'Només l\'administrador pot crear albarans a Factura Directa.'}), 403

    # El contacte FD és el CLIENT (nom del client de la calculadora)
    # El NIF del client és opcional — si s'envia, s'usa per cercar exacte a FD
    nif_client = (d.get('client_nif') or '').strip()
    nom_fd     = client_nom

    if not nom_fd:
        return jsonify({'ok': False, 'error': 'Cal omplir el nom del client abans de crear l\'albarà.'}), 400

    # Buscar per NIF (exacte) o crear nou — no busquem per nom per evitar falsos positius
    contacte = _fd_cerca_contacte(nif=nif_client) if nif_client else None
    if not contacte:
        contacte = _fd_crear_contacte(nom_fd, nif=nif_client or None, telefon=client_tel or None)
    if '_error' in (contacte or {}):
        return jsonify({'ok': False, 'error': f'Error contacte FD {contacte.get("_error")}: {contacte.get("_msg","")}'}), 500

    contact_id = _fd_extract_contact_id(contacte)
    if not contact_id:
        print(f'FD contacte sense ID (api_crear_albara): {json.dumps(contacte, ensure_ascii=False)}')
        return jsonify({'ok': False, 'error': f'Contacte FD sense ID. Resposta: {json.dumps(contacte, ensure_ascii=False)}'}), 500

    # Línies de l'albarà
    desc_marc = f'Marc {marc}' if marc else 'Emmarcació'
    parts = []
    if opcions_text:
        parts.append(opcions_text)
    if revers_peu:
        parts.append('Revers amb peu')
    if parts:
        desc_marc += f' · {", ".join(parts)}'

    # mode_preu: 'pvp' uses preu_net (PVP sense IVA), 'cost' uses cost_produccio
    base_total = preu_net if mode_preu == 'pvp' else cost_prod
    unit_price = round(base_total / quantitat, 2) if quantitat > 0 else round(base_total, 2)
    linies = [{
        'text':      desc_marc,
        'quantity':  float(quantitat),
        'unitPrice': unit_price,
        'tax':       ['S_IVA_21'],
    }]

    notes_parts = []
    if num_pressupost:
        notes_parts.append(f'Pressupost: {num_pressupost}')
    if observacions:
        notes_parts.append(f'Obs: {observacions}')
    notes = ' | '.join(notes_parts)

    albara = _fd_crear_albara(contact_id, linies, notes=notes)
    if '_error' in (albara or {}):
        return jsonify({'ok': False, 'error': f'Error albarà FD {albara.get("_error")}: {albara.get("_msg","")}'}), 500

    num_albara = albara.get('number') or albara.get('documentNumber') or albara.get('id', '—')
    return jsonify({'ok': True, 'albara': num_albara, 'contact': nom_fd})


@app.route('/api/albara-de-comanda', methods=['POST'])
@admin_required
def api_albara_de_comanda():
    """Crea un albarà FD a partir d'una sessió de comanda d'un client professional."""
    if not _FD_TOKEN or not _FD_COMPANY:
        return jsonify({'ok': False, 'error': 'Factura Directa no configurat (variables d\'entorn)'}), 503

    d = request.get_json(force=True) or {}
    sessio_id = (d.get('sessio_id') or '').strip()
    if not sessio_id:
        return jsonify({'ok': False, 'error': 'Falta sessio_id'}), 400

    # Read all lines of this session
    comandes = query(
        '''SELECT c.*, u.nom as usuari_nom, u.nom_empresa, u.nom_fiscal, u.fiscal_id, u.empresa_tel
           FROM comandes c JOIN usuaris u ON c.user_id=u.id
           WHERE c.sessio_id=? ORDER BY c.id''', [sessio_id])
    if not comandes:
        return jsonify({'ok': False, 'error': 'Sessió no trobada'}), 404

    c0 = comandes[0]

    # The FD contact is the professional client (the user who owns the session)
    # Use their nom_fiscal > nom_empresa > nom
    nom_fiscal   = (_row_get(c0, 'nom_fiscal', '') or '').strip()
    nom_empresa  = (_row_get(c0, 'nom_empresa', '') or '').strip()
    usuari_nom   = (_row_get(c0, 'usuari_nom', '') or '').strip()
    fiscal_id    = (_row_get(c0, 'fiscal_id', '') or '').strip()
    empresa_tel  = (_row_get(c0, 'empresa_tel', '') or '').strip()

    nom_fd = nom_fiscal or nom_empresa or usuari_nom
    if not nom_fd:
        return jsonify({'ok': False, 'error': 'El client no té nom fiscal ni nom d\'empresa configurat.'}), 400

    # Lookup by NIF only (exact), never by name
    contacte = _fd_cerca_contacte(nif=fiscal_id) if fiscal_id else None
    if not contacte:
        contacte = _fd_crear_contacte(nom_fd, nif=fiscal_id or None, telefon=empresa_tel or None)
    if '_error' in (contacte or {}):
        return jsonify({'ok': False, 'error': f'Error contacte FD {contacte.get("_error")}: {contacte.get("_msg","")}'}), 500

    contact_id = _fd_extract_contact_id(contacte)
    if not contact_id:
        print(f'FD contacte sense ID (api_albara_de_comanda): {json.dumps(contacte, ensure_ascii=False)}')
        return jsonify({'ok': False, 'error': f'Contacte FD sense ID. Resposta: {json.dumps(contacte, ensure_ascii=False)}'}), 500

    # Build albaran lines — one per comanda row
    linies = []
    notes_parts = []
    for com in comandes:
        marc         = (_row_get(com, 'marc_principal', '') or '').strip()
        pre_marc     = (_row_get(com, 'pre_marc', '') or '').strip()
        passpartout  = (_row_get(com, 'passpartout', '') or '').strip()
        vidre        = (_row_get(com, 'vidre', '') or '').strip()
        opcio_nom    = (_row_get(com, 'opcio_nom', '') or '').strip()
        num_pres     = (_row_get(com, 'num_pressupost', '') or '').strip()
        observacions = (_row_get(com, 'observacions', '') or '').strip()
        quantitat    = int(_row_get(com, 'quantitat', 1) or 1)
        cost_prod    = float(_row_get(com, 'cost_produccio', 0) or 0)
        client_nom   = (_row_get(com, 'client_nom', '') or '').strip()

        revers_peu   = str(_row_get(com, 'revers_peu', '') or '').strip().lower() in ('1', 'true', 'yes', 'on')
        encolat      = (_row_get(com, 'encolat', '') or '').strip()
        impressio    = (_row_get(com, 'impressio', '') or '').strip()

        parts_opc = []
        if passpartout and passpartout != 'cap': parts_opc.append(passpartout)
        if vidre:                                parts_opc.append(vidre)
        if pre_marc and pre_marc != '-':         parts_opc.append(f'+ {pre_marc}')
        if encolat and encolat != '-':           parts_opc.append(encolat)
        if revers_peu:                           parts_opc.append('Revers amb peu')
        if impressio and impressio != '-':       parts_opc.append(impressio)
        desc_marc = f'Marc {marc}' if marc else 'Emmarcació'
        if parts_opc:
            desc_marc += f' · {", ".join(parts_opc)}'
        if opcio_nom and opcio_nom != 'Opció A':
            desc_marc += f' ({opcio_nom})'
        if client_nom:
            desc_marc = f'[{client_nom}] ' + desc_marc

        # cost_produccio is total for all units (cTot*qty), divide by qty for unit price
        unit_cost = round(cost_prod / quantitat, 2) if quantitat > 0 else round(cost_prod, 2)
        linies.append({
            'text':      desc_marc,
            'quantity':  float(quantitat),
            'unitPrice': unit_cost,
            'tax':       ['S_IVA_21'],
        })
        if num_pres and num_pres not in notes_parts:
            notes_parts.append(f'Pressupost: {num_pres}')
        if observacions:
            notes_parts.append(f'Obs: {observacions}')

    notes = ' | '.join(notes_parts)

    albara = _fd_crear_albara(contact_id, linies, notes=notes)
    if '_error' in (albara or {}):
        return jsonify({'ok': False, 'error': f'Error albarà FD {albara.get("_error")}: {albara.get("_msg","")}'}), 500

    num_albara = albara.get('number') or albara.get('documentNumber') or albara.get('id', '—')

    # Store albaran number on the session rows
    execute("UPDATE comandes SET fd_albara=? WHERE sessio_id=?", [str(num_albara), sessio_id])

    return jsonify({'ok': True, 'albara': num_albara, 'contact': nom_fd})


@app.route('/api/albara-individual', methods=['POST'])
@admin_required
def api_albara_individual():
    """Crea un albarà FD per una comanda individual (no per sessió sencera)."""
    if not _FD_TOKEN or not _FD_COMPANY:
        return jsonify({'ok': False, 'error': 'Factura Directa no configurat (variables d\'entorn)'}), 503

    d = request.get_json(force=True) or {}
    comanda_id = d.get('comanda_id')
    if not comanda_id:
        return jsonify({'ok': False, 'error': 'Falta comanda_id'}), 400

    mode_preu = (d.get('mode_preu') or 'pvp').strip().lower()  # 'pvp' (default) or 'cost'

    com = query(
        '''SELECT c.*, u.nom as usuari_nom, u.nom_empresa, u.nom_fiscal, u.fiscal_id, u.empresa_tel, u.is_admin as owner_is_admin
           FROM comandes c JOIN usuaris u ON c.user_id=u.id
           WHERE c.id=?''', [comanda_id], one=True)
    if not com:
        return jsonify({'ok': False, 'error': 'Comanda no trobada'}), 404

    # Determine FD contact: if admin's own order, use client_nom; if professional's order, use nom_fiscal/nom_empresa
    owner_is_admin = bool(_row_get(com, 'owner_is_admin', 0))
    if owner_is_admin:
        # Admin's own order → the FD contact is the end client
        nom_fd    = (_row_get(com, 'client_nom', '') or '').strip()
        fiscal_id = ''
        telefon   = (_row_get(com, 'client_tel', '') or '').strip()
    else:
        # Professional's order → the FD contact is the professional
        nom_fiscal  = (_row_get(com, 'nom_fiscal', '') or '').strip()
        nom_empresa = (_row_get(com, 'nom_empresa', '') or '').strip()
        usuari_nom  = (_row_get(com, 'usuari_nom', '') or '').strip()
        nom_fd      = nom_fiscal or nom_empresa or usuari_nom
        fiscal_id   = (_row_get(com, 'fiscal_id', '') or '').strip()
        telefon     = (_row_get(com, 'empresa_tel', '') or '').strip()

    if not nom_fd:
        return jsonify({'ok': False, 'error': 'Cal un nom de contacte per crear l\'albarà.'}), 400

    contacte = _fd_cerca_contacte(nif=fiscal_id) if fiscal_id else None
    if not contacte:
        contacte = _fd_crear_contacte(nom_fd, nif=fiscal_id or None, telefon=telefon or None)
    if '_error' in (contacte or {}):
        return jsonify({'ok': False, 'error': f'Error contacte FD {contacte.get("_error")}: {contacte.get("_msg","")}'}), 500

    contact_id = _fd_extract_contact_id(contacte)
    if not contact_id:
        return jsonify({'ok': False, 'error': f'Contacte FD sense ID.'}), 500

    # Build line item
    marc        = (_row_get(com, 'marc_principal', '') or '').strip()
    pre_marc    = (_row_get(com, 'pre_marc', '') or '').strip()
    passpartout = (_row_get(com, 'passpartout', '') or '').strip()
    vidre       = (_row_get(com, 'vidre', '') or '').strip()
    encolat     = (_row_get(com, 'encolat', '') or '').strip()
    impressio   = (_row_get(com, 'impressio', '') or '').strip()
    revers_peu  = str(_row_get(com, 'revers_peu', '') or '').strip().lower() in ('1', 'true', 'yes', 'on')
    quantitat   = int(_row_get(com, 'quantitat', 1) or 1)
    cost_prod   = float(_row_get(com, 'cost_produccio', 0) or 0)
    preu_net    = float(_row_get(com, 'preu_net', 0) or 0)
    client_nom  = (_row_get(com, 'client_nom', '') or '').strip()
    opcio_nom   = (_row_get(com, 'opcio_nom', '') or '').strip()

    parts_opc = []
    if passpartout and passpartout != 'cap': parts_opc.append(passpartout)
    if vidre:                                parts_opc.append(vidre)
    if pre_marc and pre_marc != '-':         parts_opc.append(f'+ {pre_marc}')
    if encolat and encolat != '-':           parts_opc.append(encolat)
    if revers_peu:                           parts_opc.append('Revers amb peu')
    if impressio and impressio != '-':       parts_opc.append(impressio)
    desc_marc = f'Marc {marc}' if marc else 'Emmarcació'
    if parts_opc:
        desc_marc += f' · {", ".join(parts_opc)}'
    if opcio_nom and opcio_nom != 'Opció A':
        desc_marc += f' ({opcio_nom})'
    if client_nom and not owner_is_admin:
        desc_marc = f'[{client_nom}] ' + desc_marc

    # mode_preu: 'pvp' uses preu_net (PVP sense IVA), 'cost' uses cost_produccio
    base_total = preu_net if mode_preu == 'pvp' else cost_prod
    unit_price = round(base_total / quantitat, 2) if quantitat > 0 else round(base_total, 2)
    linies = [{
        'text':      desc_marc,
        'quantity':  float(quantitat),
        'unitPrice': unit_price,
        'tax':       ['S_IVA_21'],
    }]

    notes_parts = []
    num_pres = (_row_get(com, 'num_pressupost', '') or '').strip()
    if num_pres:
        notes_parts.append(f'Pressupost: {num_pres}')
    observacions = (_row_get(com, 'observacions', '') or '').strip()
    if observacions:
        notes_parts.append(f'Obs: {observacions}')
    notes = ' | '.join(notes_parts)

    albara = _fd_crear_albara(contact_id, linies, notes=notes)
    if '_error' in (albara or {}):
        return jsonify({'ok': False, 'error': f'Error albarà FD {albara.get("_error")}: {albara.get("_msg","")}'}), 500

    num_albara = albara.get('number') or albara.get('documentNumber') or albara.get('id', '—')
    execute("UPDATE comandes SET fd_albara=? WHERE id=?", [str(num_albara), comanda_id])

    return jsonify({'ok': True, 'albara': num_albara, 'contact': nom_fd})


@app.route('/api/passpartous')
@login_required
def api_passpartous():
    """Llista de referències de passpartú (color/textura) per al selector de la calculadora."""
    rows = query('SELECT referencia, color, textura, descripcio FROM passpartout ORDER BY referencia') or []
    return jsonify([{
        'ref': r['referencia'],
        'color': r['color'] or '',
        'textura': r['textura'] or '',
        'descripcio': r['descripcio'] or '',
    } for r in rows])


@app.route('/admin/passpartous', methods=['POST'])
@admin_required
def admin_passpartous():
    action = request.form.get('action')
    if action == 'crear':
        ref = request.form.get('referencia', '').strip().upper()
        color = request.form.get('color', '').strip()
        textura = request.form.get('textura', '').strip()
        descripcio = request.form.get('descripcio', '').strip()
        if ref:
            execute('INSERT OR REPLACE INTO passpartout (referencia, color, textura, descripcio) VALUES (?,?,?,?)',
                    [ref, color, textura, descripcio])
            flash(f'Referència {ref} desada.', 'ok')
    elif action == 'eliminar':
        ref = request.form.get('ref', '').strip()
        execute('DELETE FROM passpartout WHERE referencia=?', [ref])
        flash('Referència eliminada.', 'ok')
    return redirect(url_for('admin'))


@app.route('/admin/proeco', methods=['POST'])
@admin_required
def admin_proeco():
    action = request.form.get('action')
    if action == 'crear':
        ref = request.form.get('referencia', '').strip().upper()
        preu = float(request.form.get('preu', 0))
        if ref:
            execute('INSERT OR REPLACE INTO proeco (referencia, preu) VALUES (?,?)', [ref, preu])
            flash(f'ProEco {ref} desat.', 'ok')
    elif action == 'eliminar':
        execute('DELETE FROM proeco WHERE referencia=?', [request.form.get('ref')])
        flash('Preu ProEco eliminat.', 'ok')
    elif action == 'editar':
        ref = request.form.get('ref', '').strip().upper()
        preu = float(request.form.get('preu', 0))
        execute('UPDATE proeco SET preu=? WHERE referencia=?', [preu, ref])
        flash(f'ProEco {ref} actualitzat.', 'ok')
    return redirect(url_for('admin'))


@app.route('/admin/impressio', methods=['POST'])
@admin_required
def admin_impressio():
    action = request.form.get('action')
    if action == 'crear':
        ref = request.form.get('referencia','').strip().upper()
        desc = request.form.get('descripcio','').strip()
        preu = float(request.form.get('preu', 0))
        execute('INSERT OR REPLACE INTO impressio VALUES (?,?,?)', [ref, desc, preu])
        flash(f'Format {ref} desat.', 'ok')
    elif action == 'eliminar':
        execute('DELETE FROM impressio WHERE referencia=?', [request.form.get('ref')])
        flash('Format eliminat.', 'ok')
    return redirect(url_for('admin'))

@app.route('/admin/foto', methods=['POST'])
@admin_required
def admin_foto():
    ref = request.form.get('referencia', '').strip().lower()
    f = request.files.get('foto')
    if f and ref:
        try:
            foto = _save_moldura_photo(f, ref)
            execute('UPDATE moldures SET foto=? WHERE LOWER(referencia)=?', [foto, ref])
            flash(f'Foto de {ref} pujada.', 'ok')
        except ValueError as e:
            flash(str(e), 'error')
    return redirect(url_for('admin'))

# â"€â"€ PDF generator â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€

def generar_num_pressupost():
    """Genera número tipus RR-2503-001"""
    from datetime import datetime
    # Get initials from empresa_nom in config
    r = query("SELECT valor FROM config WHERE clau='empresa_nom'", one=True)
    nom = r['valor'] if r and r['valor'] else 'XX'
    # Take first letters of each word (max 3)
    inicials = ''.join(w[0].upper() for w in nom.split()[:3] if w)[:3] or 'XX'
    # YYMM
    yymm = datetime.now().strftime('%y%m')
    prefix = f"{inicials}-{yymm}-"
    # Find last number with same prefix
    r2 = query("SELECT num_pressupost FROM comandes WHERE num_pressupost LIKE ? ORDER BY id DESC LIMIT 1",
               [prefix + '%'], one=True)
    if r2 and r2['num_pressupost']:
        try:
            last_n = int(r2['num_pressupost'].split('-')[-1])
            return f"{prefix}{last_n+1:03d}"
        except:
            pass
    return f"{prefix}001"


def crear_pdf_comparativa(comandes):
    from reportlab.lib.pagesizes import landscape
    buf = io.BytesIO()
    PAGE = landscape(A4)
    W_page, H_page = PAGE
    margin = 15*mm
    W = W_page - 2*margin
    n = len(comandes)
    col_lbl = 38*mm
    col_w = (W - col_lbl) / n

    doc = SimpleDocTemplate(buf, pagesize=PAGE,
                            rightMargin=margin, leftMargin=margin,
                            topMargin=12*mm, bottomMargin=15*mm)
    DARK  = colors.HexColor("#1C1B18")
    ACC   = colors.HexColor("#1A6B45")
    LIG   = colors.HexColor("#F5F6FA")
    BRD   = colors.HexColor("#E5E2DB")
    GREEN = colors.HexColor("#1A6B45")
    RED   = colors.HexColor("#B84040")
    WHITE = colors.white

    def p(txt, bold=False, size=10, color=DARK, align='LEFT'):
        st = ParagraphStyle('x', fontName='DejaVu-Bold' if bold else 'DejaVu',
                            fontSize=size, textColor=color,
                            alignment={'LEFT':0,'CENTER':1,'RIGHT':2}[align])
        return Paragraph(str(txt), st)

    story = []

    lang = (comandes[0].get('lang') or 'ca').lower()
    t = PDF_T.get(lang, PDF_T['ca'])

    # Header
    c0 = comandes[0]
    _u_comp = query('SELECT nom_empresa FROM usuaris WHERE id=?', [c0.get('user_id',0)], one=True)
    nom_empresa_comp = (_row_get(_u_comp, 'nom_empresa', '') or '') if _u_comp else ''
    if not nom_empresa_comp:
        _r_comp = query("SELECT valor FROM config WHERE clau='empresa_nom'", one=True)
        nom_empresa_comp = (_r_comp['valor'] if _r_comp else '') or 'Reus Revela'
    header = Table([[
        p(f"{t['comparativa']}", bold=True, size=14, color=colors.white),
        p(nom_empresa_comp, size=9, color=colors.HexColor("#B2BEC3"), align='RIGHT')
    ]], colWidths=[W*0.65, W*0.35])
    header.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1), DARK),
        ('TOPPADDING',(0,0),(-1,-1), 10), ('BOTTOMPADDING',(0,0),(-1,-1), 10),
        ('LEFTPADDING',(0,0),(-1,-1), 12), ('RIGHTPADDING',(0,0),(-1,-1), 12),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
    ]))
    story.append(header)

    client_info = Table([[
        p(f"{t['client']}: {c0['client_nom'] or '—'}", bold=True, size=11),
        p(f"{t['data']}: {c0['data']}", size=10, align='RIGHT')
    ]], colWidths=[W*0.6, W*0.4])
    client_info.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1), LIG),
        ('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),
        ('LEFTPADDING',(0,0),(-1,-1),12),('RIGHTPADDING',(0,0),(-1,-1),12),
        ('BOX',(0,0),(-1,-1),0.5,BRD),
    ]))
    story.append(client_info)
    from reportlab.platypus import Spacer as Sp
    story.append(Sp(1,4*mm))

    # Column headers (with label column)
    hdr = [p('', bold=True, size=10, color=WHITE)] +           [p(c.get('opcio_nom','Opció'), bold=True, size=11, color=WHITE, align='CENTER') for c in comandes]
    hdr_row = Table([hdr], colWidths=[col_lbl]+[col_w]*n)
    hdr_row.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1), DARK),
        ('TOPPADDING',(0,0),(-1,-1),9),('BOTTOMPADDING',(0,0),(-1,-1),9),
        ('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),
        ('BOX',(0,0),(-1,-1),0.5,BRD),
        ('INNERGRID',(0,1),(-1,-1),0.5,colors.HexColor("#3d4d5e")),
    ]))
    story.append(hdr_row)

    # Photo row (frame images for each option)
    import os as _os_comp
    from reportlab.platypus import Image as RLImageComp
    _foto_cells = [p('', size=8)]
    _has_foto_comp = False
    for _c in comandes:
        _cell = p('', size=8)
        if _c.get('marc_principal'):
            _r_f = query('SELECT foto, ref2 FROM moldures WHERE LOWER(referencia)=LOWER(?)',
                         [_c['marc_principal']], one=True)
            _foto_url = _resolve_moldura_photo(_c['marc_principal'], _r_f['foto'] if _r_f else '',
                                               ref2=(_row_get(_r_f,'ref2','') or '') if _r_f else '')
            if _foto_url.startswith('/static/'):
                _rel = _foto_url.lstrip('/')
                _full = _os_comp.path.join(app.root_path, 'static', _rel.replace('static/',''))
                if _os_comp.path.exists(_full):
                    try:
                        _img_c = RLImageComp(_full, width=min(col_w - 8*mm, 35*mm), height=22*mm)
                        _cell = _img_c
                        _has_foto_comp = True
                    except Exception:
                        pass
        _foto_cells.append(_cell)
    if _has_foto_comp:
        foto_row_tbl = Table([_foto_cells], colWidths=[col_lbl]+[col_w]*n)
        foto_row_tbl.setStyle(TableStyle([
            ('BACKGROUND',(0,0),(-1,-1), LIG),
            ('BOX',(0,0),(-1,-1),0.5,BRD),
            ('INNERGRID',(0,0),(-1,-1),0.3,BRD),
            ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
            ('LEFTPADDING',(0,0),(-1,-1),4),('RIGHTPADDING',(0,0),(-1,-1),4),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
            ('ALIGN',(1,0),(-1,0),'CENTER'),
        ]))
        story.append(foto_row_tbl)

    # Build comparison rows
    def val_clean(c, key):
        v = c.get(key,'') or ''
        return '—' if v in ['','-','None'] else str(v)

    fields = [
        (t['marc_principal'], 'marc_principal',   False),
        (t['tipus_peca'],     'piece_type',       False),
        (t['mida_final'],     'final_size',      False),
        (t['mides_foto'],     'photo_size',      False),
        (t['muntatge'],       'muntatge_label',  False),
        (t['vidre_mirall'],   'proteccio_label', False),
        (t['interior'],       'interior_label',  False),
        (t['revers_peu'],     'revers_peu_label', False),
        (t['impressio'],      'impressio_label', False),
        (t['observacions'],   'observacions',    False),
        (t['preu_net_label'], 'preu_net',        True),
        (t['iva'],            'iva',             True),
        (t['total_iva'],      'preu_final',      True),
        (t['entrega'],        'entrega',         True),
        (t['pendent_short'],  'pendent',         True),
    ]

    rows = []
    for lbl, key, is_price in fields:
        lbl_color = DARK if not is_price else colors.HexColor("#1A6B45") if key=='preu_final' else                     RED if key=='pendent' else colors.HexColor("#6B6860")
        row = [p(lbl, bold=is_price, size=9, color=lbl_color)]
        for c in comandes:
            if key == 'marc_principal':
                val = c.get('marc_principal') or t['sense_marc']
            elif key == 'piece_type':
                val = _display_piece_type(c, t)
            elif key == 'final_size':
                val = _final_size_text(c, sep=' x ', with_unit=True)
            elif key == 'photo_size':
                val = _photo_size_text(c, sep=' x ', with_unit=True)
            elif key == 'muntatge_label':
                val = _display_muntatge(c, t)
            elif key == 'proteccio_label':
                val = _display_proteccio(c, t)
            elif key == 'interior_label':
                val = _display_interior(c, t)
            elif key == 'revers_peu_label':
                val = _display_revers_peu(c, t)
            elif key == 'impressio_label':
                val = _display_impressio(c, t)
            elif key == 'iva':
                pn = float(c.get('preu_net',0) or 0)
                val = f"{pn*0.21:.2f} €"
            elif is_price:
                v = float(c.get(key,0) or 0)
                val = f"{v:.2f} €"
            else:
                val = val_clean(c, key)
            bold_cell = is_price and key in ('preu_final','pendent')
            cell_color = GREEN if key=='preu_final' else RED if key=='pendent' else DARK
            row.append(p(val, bold=bold_cell, size=9 if not bold_cell else 11,
                        color=cell_color, align='CENTER'))
        rows.append(row)

    col_widths = [col_lbl] + [col_w]*n
    detail_table = Table(rows, colWidths=col_widths)
    price_start = len([f for f in fields if not f[2]])  # index where prices start
    detail_table.setStyle(TableStyle([
        ('ROWBACKGROUNDS',(0,0),(-1,-1),[LIG, WHITE]),
        ('BOX',(0,0),(-1,-1),0.5,BRD),
        ('INNERGRID',(0,0),(-1,-1),0.3,BRD),
        ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
        ('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),
        ('BACKGROUND',(0,0),(0,-1),colors.HexColor("#F2F0EC")),
        ('FONTNAME',(0,0),(0,-1),'DejaVu-Bold'),
        ('LINEABOVE',(0,price_start),(-1,price_start),1.5,ACC),
        ('BACKGROUND',(0,price_start+2),(-1,price_start+2),colors.HexColor("#E8F3EE")),
        ('BACKGROUND',(0,price_start+4),(-1,price_start+4),colors.HexColor("#FAEAEA")),
    ]))
    story.append(detail_table)

    doc.build(story)
    buf.seek(0)
    return buf


def crear_pdf_cataleg_admin(moldures, q='', proveidor=''):
    from reportlab.lib.pagesizes import landscape
    from reportlab.platypus import Image as RLImage

    buf = io.BytesIO()
    page = landscape(A4)
    margin = 10 * mm
    width = page[0] - (margin * 2)
    doc = SimpleDocTemplate(
        buf,
        pagesize=page,
        rightMargin=margin,
        leftMargin=margin,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )

    dark = colors.HexColor("#1C1B18")
    green = colors.HexColor("#1A6B45")
    border = colors.HexColor("#E5E2DB")
    light = colors.HexColor("#F5F4F1")
    muted = colors.HexColor("#6B6860")

    def p(txt, *, bold=False, size=8.5, color=dark, align=0):
        return Paragraph(
            str(txt),
            ParagraphStyle(
                name=f"cat-{size}-{int(bold)}-{align}",
                fontName='DejaVu-Bold' if bold else 'DejaVu',
                fontSize=size,
                leading=size + 2,
                textColor=color,
                alignment=align,
            )
        )

    story = [
        p("Cataleg de motllures - Admin", bold=True, size=16),
        Spacer(1, 3 * mm),
        p(
            f"Total de motllures: {len(moldures)}"
            + (f" | Cerca: {q}" if q else "")
            + (f" | Proveidor: {proveidor}" if proveidor else ""),
            size=9,
            color=muted,
        ),
        Spacer(1, 5 * mm),
    ]

    headers = [[
        p("Imatge", bold=True, size=8, color=colors.white, align=TA_CENTER),
        p("Referencia", bold=True, size=8, color=colors.white, align=TA_CENTER),
        p("Descripcio", bold=True, size=8, color=colors.white, align=TA_CENTER),
        p("Preu taller", bold=True, size=8, color=colors.white, align=TA_CENTER),
        p("Gruix", bold=True, size=8, color=colors.white, align=TA_CENTER),
        p("Cost", bold=True, size=8, color=colors.white, align=TA_CENTER),
        p("Proveidor", bold=True, size=8, color=colors.white, align=TA_CENTER),
        p("Ubicacio", bold=True, size=8, color=colors.white, align=TA_CENTER),
        p("Ref2", bold=True, size=8, color=colors.white, align=TA_CENTER),
    ]]

    rows = []
    for moldura in moldures:
        photo_path = _moldura_photo_path(moldura.get('referencia', ''), moldura.get('foto', ''))
        if photo_path:
            photo_cell = RLImage(photo_path, width=16 * mm, height=16 * mm)
        else:
            photo_cell = p("Sense foto", size=7, color=muted, align=TA_CENTER)

        rows.append([
            photo_cell,
            p(moldura.get('referencia', '-') or '-', bold=True, size=8),
            p(moldura.get('descripcio', '-') or '-', size=8),
            p(f"{float(moldura.get('preu_taller', 0) or 0):.2f} EUR/m", size=8, align=TA_RIGHT),
            p(f"{float(moldura.get('gruix', 0) or 0):.1f} cm", size=8, align=TA_RIGHT),
            p(f"{float(moldura.get('cost', 0) or 0):.2f} EUR", size=8, align=TA_RIGHT),
            p(moldura.get('proveidor', '-') or '-', size=8),
            p(moldura.get('ubicacio', '-') or '-', size=8),
            p(moldura.get('ref2', '-') or '-', size=8),
        ])

    if not rows:
        rows.append([p("Sense resultats", size=9, color=muted)] + [''] * 8)

    table = Table(
        headers + rows,
        repeatRows=1,
        colWidths=[18 * mm, 24 * mm, 66 * mm, 22 * mm, 16 * mm, 18 * mm, 30 * mm, 25 * mm, 24 * mm],
    )
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), dark),
        ('BOX', (0, 0), (-1, -1), 0.5, border),
        ('INNERGRID', (0, 0), (-1, -1), 0.35, border),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, light]),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(table)

    doc.build(story)
    buf.seek(0)
    return buf

# â"€â"€ Translations for PDF â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
PDF_T = {
    'ca': {
        'pressupost': 'PRESSUPOST',
        'client': 'Client',
        'telefon': 'Telèfon',
        'data': 'Data',
        'marc': 'Marc',
        'premarc': 'Pre-Marc',
        'mides': 'Mides',
        'mida_final': 'Mida final',
        'tipus_peca': 'Peça a emmarcar',
        'muntatge': 'Muntatge',
        'proteccio': 'Protecció',
        'interior': 'Interior',
        'revers_peu': 'Revers amb peu',
        'impressio': 'Impressió',
        'observacions': 'Observacions',
        'preu_net': 'Preu net (sense IVA)',
        'iva': 'IVA 21%',
        'total': 'Total amb IVA',
        'entrega': 'Entrega a compte',
        'pendent': 'Pendent de cobrar',
        'num': 'Num. Pressupost',
        'comparativa': 'COMPARATIVA DE PRESSUPOSTOS',
        'marc_principal': 'Marc principal',
        'mides_foto': 'Mides peça (cm)',
        'vidre_mirall': 'Vidre / Mirall',
        'preu_net_label': 'Preu net',
        'total_iva': 'TOTAL amb IVA',
        'pendent_short': 'PENDENT',
        'sense_marc': '(sense marc)',
        'encolat_label': 'Encolat',
        'laminat_label': 'Laminat',
        'protter': 'Protter',
        'vidre_label': 'Vidre',
        'doble_vidre': 'Doble vidre',
        'mirall': 'Mirall',
        'passpartu_label': 'Passpartú',
        'doble_passpartu': 'Doble passpartú',
        'proeco_label': 'ProEco',
        'inclosa': 'Inclosa',
        'piece_photo': 'Fotografia',
        'piece_lamina': 'Làmina',
        'piece_painting_unstretched': 'Pintura sense bastidor',
        'piece_painting_stretched': 'Pintura amb bastidor',
        'piece_puzzle': 'Puzzle',
        'piece_cross_stitch': 'Punt de creu',
        'conservar_vidre': 'Conservar vidre existent',
        'preu_net_pvp': 'Preu net PVP (sense IVA)',
        'descompte_sobre_pvp': 'Descompte {pct}% sobre PVP',
        'total_pvp_iva': 'TOTAL PVP amb IVA',
        'pendent_cobrar': 'PENDENT de cobrar',
    },
    'es': {
        'pressupost': 'PRESUPUESTO',
        'client': 'Cliente',
        'telefon': 'Teléfono',
        'data': 'Fecha',
        'marc': 'Marco',
        'premarc': 'Pre-Marco',
        'mides': 'Medidas',
        'mida_final': 'Medida final',
        'tipus_peca': 'Pieza a enmarcar',
        'muntatge': 'Montaje',
        'proteccio': 'Protección',
        'interior': 'Interior',
        'revers_peu': 'Reverso con pie',
        'impressio': 'Impresión',
        'observacions': 'Observaciones',
        'preu_net': 'Precio neto (sin IVA)',
        'iva': 'IVA 21%',
        'total': 'Total con IVA',
        'entrega': 'Pago a cuenta',
        'pendent': 'Pendiente de cobro',
        'num': 'Num. Presupuesto',
        'comparativa': 'COMPARATIVA DE PRESUPUESTOS',
        'marc_principal': 'Marco principal',
        'mides_foto': 'Medidas pieza (cm)',
        'vidre_mirall': 'Vidrio / Espejo',
        'preu_net_label': 'Precio neto',
        'total_iva': 'TOTAL con IVA',
        'pendent_short': 'PENDIENTE',
        'sense_marc': '(sin marco)',
        'encolat_label': 'Encolado',
        'laminat_label': 'Laminado',
        'protter': 'Protter',
        'vidre_label': 'Vidrio',
        'doble_vidre': 'Doble vidrio',
        'mirall': 'Espejo',
        'passpartu_label': 'Passpartú',
        'doble_passpartu': 'Doble passpartú',
        'proeco_label': 'ProEco',
        'inclosa': 'Incluida',
        'piece_photo': 'Fotografía',
        'piece_lamina': 'Lámina',
        'piece_painting_unstretched': 'Pintura sin bastidor',
        'piece_painting_stretched': 'Pintura con bastidor',
        'piece_puzzle': 'Puzzle',
        'piece_cross_stitch': 'Punto de cruz',
        'conservar_vidre': 'Conservar vidrio existente',
        'preu_net_pvp': 'Precio neto PVP (sin IVA)',
        'descompte_sobre_pvp': 'Descuento {pct}% sobre PVP',
        'total_pvp_iva': 'TOTAL PVP con IVA',
        'pendent_cobrar': 'PENDIENTE de cobro',
    },
    'en': {
        'pressupost': 'QUOTE',
        'client': 'Client',
        'telefon': 'Phone',
        'data': 'Date',
        'marc': 'Frame',
        'premarc': 'Inner frame',
        'mides': 'Dimensions',
        'mida_final': 'Final size',
        'tipus_peca': 'Item to frame',
        'muntatge': 'Mounting',
        'proteccio': 'Protection',
        'interior': 'Interior',
        'revers_peu': 'Backing with stand',
        'impressio': 'Print',
        'observacions': 'Notes',
        'preu_net': 'Net price (excl. VAT)',
        'iva': 'VAT 21%',
        'total': 'Total incl. VAT',
        'entrega': 'Deposit',
        'pendent': 'Outstanding',
        'num': 'Quote No.',
        'comparativa': 'QUOTE COMPARISON',
        'marc_principal': 'Main frame',
        'mides_foto': 'Item size (cm)',
        'vidre_mirall': 'Glass / Mirror',
        'preu_net_label': 'Net price',
        'total_iva': 'TOTAL incl. VAT',
        'pendent_short': 'OUTSTANDING',
        'sense_marc': '(no frame)',
        'encolat_label': 'Mounting',
        'laminat_label': 'Laminate only',
        'protter': 'Protter',
        'vidre_label': 'Glass',
        'doble_vidre': 'Double glass',
        'mirall': 'Mirror',
        'passpartu_label': 'Mat',
        'doble_passpartu': 'Double mat',
        'proeco_label': 'ProEco',
        'inclosa': 'Included',
        'piece_photo': 'Photograph',
        'piece_lamina': 'Poster',
        'piece_painting_unstretched': 'Unstretched painting',
        'piece_painting_stretched': 'Stretched painting',
        'piece_puzzle': 'Puzzle',
        'piece_cross_stitch': 'Cross stitch',
        'conservar_vidre': 'Keep existing glass',
        'preu_net_pvp': 'Net retail price (excl. VAT)',
        'descompte_sobre_pvp': 'Discount {pct}% on retail price',
        'total_pvp_iva': 'TOTAL retail incl. VAT',
        'pendent_cobrar': 'OUTSTANDING amount',
    }
}

def crear_pdf(c):
    import os as _os
    from reportlab.platypus import Image as RLImage
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=15*mm, leftMargin=15*mm,
                            topMargin=12*mm, bottomMargin=15*mm)
    W = A4[0] - 30*mm
    DARK  = colors.HexColor("#1C1B18")
    GREEN = colors.HexColor("#1A6B45")
    AMBER = colors.HexColor("#C8873A")
    RED   = colors.HexColor("#B84040")
    LIG   = colors.HexColor("#F5F6FA")
    BRD   = colors.HexColor("#E5E2DB")
    WHITE = colors.white

    def p(txt, bold=False, size=10, color=DARK, align='LEFT'):
        st = ParagraphStyle('x', fontName='DejaVu-Bold' if bold else 'DejaVu',
                            fontSize=size, textColor=color,
                            alignment={'LEFT':0,'CENTER':1,'RIGHT':2}[align],
                            leading=size*1.4)
        return Paragraph(str(txt) if txt else '—', st)

    def fila(lbl, val, color_val=None):
        return [p(lbl, bold=True, size=9, color=colors.HexColor("#6B6860")),
                p(str(val) if val and val not in ['-',''] else '—', size=10,
                  color=color_val or DARK)]

    story = []

    # ── Capçalera ─────────────────────────────────────────────────────────
    # Get empresa info for this user
    u_data = query('SELECT nom_empresa, empresa_adreca, empresa_tel, brand_color FROM usuaris WHERE id=?', [c.get('user_id',0)], one=True)
    nom_empresa = ''
    if u_data and _row_get(u_data, 'nom_empresa', ''):
        nom_empresa = _row_get(u_data, 'nom_empresa', '')
    green_hex = _normalize_hex_color(_row_get(u_data, 'brand_color', DEFAULT_BRAND_COLOR))
    if not nom_empresa:
        _r = query("SELECT valor FROM config WHERE clau='empresa_nom'", one=True)
        nom_empresa = (_r['valor'] if _r else '') or 'Reus Revela'
    adreca = _row_get(u_data, 'empresa_adreca', '') or ''
    if not adreca:
        r_adr = query("SELECT valor FROM config WHERE clau='empresa_adreca'", one=True)
        adreca = (r_adr['valor'] if r_adr else '') or 'C/ Mare Molas, 26 · Reus'

    GREEN = colors.HexColor(green_hex)
    lang = (c.get('lang') or 'ca').lower()
    t = PDF_T.get(lang, PDF_T['ca'])

    header = Table([[
        p(t['pressupost'], bold=True, size=20, color=WHITE),
        p(nom_empresa + '\n' + adreca, size=8,
          color=colors.HexColor("#9E9B94"), align='RIGHT')
    ]], colWidths=[W*0.55, W*0.45])
    header.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1), DARK),
        ('TOPPADDING',(0,0),(-1,-1),14),('BOTTOMPADDING',(0,0),(-1,-1),14),
        ('LEFTPADDING',(0,0),(-1,-1),14),('RIGHTPADDING',(0,0),(-1,-1),14),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
    ]))
    story.append(header)
    story.append(Spacer(1, 3*mm))

    # â"€â"€ Logo (si existeix) â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
    try:
        u_logo = query('SELECT logo_b64 FROM usuaris WHERE id=?', [c.get('user_id',0)], one=True)
        logo_data_url = _row_get(u_logo, 'logo_b64', '') or ''
        if logo_data_url and logo_data_url.startswith('data:'):
            import base64 as _b64
            data_url = logo_data_url
            _, b64data = data_url.split(',', 1)
            img_data = _b64.b64decode(b64data)
            from reportlab.platypus import Image as RLImg2
            from PIL import Image as PILImg
            import io as _io2
            # Get real dimensions to preserve aspect ratio
            pil = PILImg.open(_io2.BytesIO(img_data))
            orig_w, orig_h = pil.size
            max_h = 25 * mm
            max_w = 60 * mm
            ratio = min(max_w / orig_w, max_h / orig_h)
            logo_w = orig_w * ratio
            logo_h = orig_h * ratio
            logo_img = RLImg2(io.BytesIO(img_data), width=logo_w, height=logo_h)
            logo_img.hAlign = 'CENTER'
            story.append(Spacer(1, 3*mm))
            story.append(logo_img)
            story.append(Spacer(1, 3*mm))
    except Exception as _e:
        print(f"Logo PDF error: {_e}")

    # â"€â"€ Dades client + data â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
    opcio_txt = c.get('opcio_nom','') or ''
    num_pres = c.get('num_pressupost','') or ''
    t1_rows = [
        fila(t['num']+':', num_pres, color_val=colors.HexColor('#1A6B45')) if num_pres else None,
        fila(t['client']+':', c['client_nom'] or '—'),
        fila(t['telefon']+':', c['client_tel'] or '—'),
        fila(t['data']+':', c['data']),
    ]
    t1_rows = [r for r in t1_rows if r is not None]
    if opcio_txt and opcio_txt != 'Opció A':
        t1_rows.append(fila('Opció:', opcio_txt))

    t1 = Table(t1_rows, colWidths=[W*0.32, W*0.68])
    t1.setStyle(TableStyle([
        ('ROWBACKGROUNDS',(0,0),(-1,-1),[LIG, WHITE]),
        ('BOX',(0,0),(-1,-1),0.5,BRD),('INNERGRID',(0,0),(-1,-1),0.3,BRD),
        ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6),
        ('LEFTPADDING',(0,0),(-1,-1),10),('RIGHTPADDING',(0,0),(-1,-1),10),
    ]))
    story.append(t1)
    story.append(Spacer(1, 5*mm))

    # â"€â"€ Foto del marc (si existeix) â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
    # Foto del marc + referencia (side-by-side)
    foto_path = None
    marc_ref = c.get('marc_principal') or ''
    marc_descripcio = ''
    if marc_ref:
        r = query('SELECT foto, descripcio, ref2 FROM moldures WHERE LOWER(referencia)=LOWER(?)',
                  [marc_ref], one=True)
        marc_descripcio = (_row_get(r, 'descripcio', '') or '') if r else ''
        foto_url = _resolve_moldura_photo(marc_ref, r['foto'] if r else '',
                                          ref2=(_row_get(r, 'ref2', '') or '') if r else '')
        if foto_url.startswith('/static/'):
            rel = foto_url.lstrip('/')
            full = _os.path.join(app.root_path, 'static', rel.replace('static/',''))
            if _os.path.exists(full):
                foto_path = full

    if foto_path:
        try:
            _ref_lbl = marc_descripcio or marc_ref
            _img_rl = RLImage(foto_path, width=45*mm, height=32*mm)
            _ref_cell_rows = [[p(t['marc_principal'], bold=True, size=8,
                                  color=colors.HexColor("#6B6860"))],
                               [p(_ref_lbl, bold=True, size=11)]]
            if marc_descripcio:
                _ref_cell_rows.append([p(marc_ref, size=8,
                                         color=colors.HexColor("#6B6860"))])
            _ref_cell = Table(_ref_cell_rows, colWidths=[W - 55*mm])
            _ref_cell.setStyle(TableStyle([
                ('VALIGN',(0,0),(-1,-1),'TOP'),
                ('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),
                ('LEFTPADDING',(0,0),(-1,-1),0),('RIGHTPADDING',(0,0),(-1,-1),0),
            ]))
            foto_side = Table([[_img_rl, _ref_cell]], colWidths=[50*mm, W - 50*mm])
            foto_side.setStyle(TableStyle([
                ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('BACKGROUND',(0,0),(-1,-1), LIG),
                ('BOX',(0,0),(-1,-1),0.5,BRD),
                ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6),
                ('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),
            ]))
            story.append(foto_side)
            story.append(Spacer(1, 4*mm))
        except Exception as _e:
            print(f"Foto marc PDF error: {_e}")

    # Detall de la comanda
    final_size = _final_size_text(c, with_unit=True)
    photo_size = _photo_size_text(c, with_unit=True)
    piece_type = _display_piece_type(c, t)
    muntatge_label = _display_muntatge(c, t)
    proteccio_label = _display_proteccio(c, t)
    interior_label = _display_interior(c, t)
    revers_peu_label = _display_revers_peu(c, t)
    impressio_label = _display_impressio(c, t)

    det_rows = []
    if piece_type:
        det_rows.append(fila(t['tipus_peca']+':', piece_type))
    if final_size:
        det_rows.append(fila(t['mida_final']+':', final_size))
    if photo_size:
        det_rows.append(fila(t['mides_foto']+':', photo_size))
    if not foto_path:
        # Show marc as text row only when there's no photo (photo already shown above)
        det_rows.append(fila(t['marc_principal']+':',
                             marc_ref or t['sense_marc']))
    if muntatge_label:
        det_rows.append(fila(t['muntatge']+':', muntatge_label))
    if proteccio_label:
        det_rows.append(fila(t['vidre_mirall']+':', proteccio_label))
    if interior_label:
        det_rows.append(fila(t['interior']+':', interior_label))
    if revers_peu_label:
        det_rows.append(fila(t['revers_peu']+':', revers_peu_label))
    if impressio_label:
        det_rows.append(fila(t['impressio']+':', impressio_label))
    if c.get('observacions'):
        det_rows.append(fila(t['observacions']+':', c['observacions']))

    if det_rows:
        t2 = Table(det_rows, colWidths=[W*0.38, W*0.62])
        t2.setStyle(TableStyle([
            ('ROWBACKGROUNDS',(0,0),(-1,-1),[LIG, WHITE]),
            ('BOX',(0,0),(-1,-1),0.5,BRD),('INNERGRID',(0,0),(-1,-1),0.3,BRD),
            ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6),
            ('LEFTPADDING',(0,0),(-1,-1),10),('RIGHTPADDING',(0,0),(-1,-1),10),
        ]))
        story.append(t2)
    story.append(Spacer(1, 5*mm))

    # ── Resum econòmic ────────────────────────────────────────────────────
    desc_pct = float(c.get('descompte') or 0)
    pnet  = float(c.get('preu_net')   or 0)
    pfin  = float(c.get('preu_final') or 0)
    piva  = pnet * 1.21
    pent  = float(c.get('entrega')    or 0)
    ppend = float(c.get('pendent')    or 0)

    t3_data = [
        [p(t['preu_net_pvp'], bold=True, size=9, color=colors.HexColor("#6B6860")),
         p(f'{pnet:.2f} €', size=10, align='RIGHT')],
    ]
    if desc_pct > 0:
        desc_eur = pnet * (desc_pct/100)
        t3_data.append([
            p(t['descompte_sobre_pvp'].format(pct=int(desc_pct)), bold=True, size=9, color=colors.HexColor("#C8873A")),
            p(f'- {desc_eur:.2f} €', size=10, color=colors.HexColor("#C8873A"), align='RIGHT'),
        ])
    t3_data += [
        [p(t['iva'], bold=True, size=9, color=colors.HexColor("#6B6860")),
         p(f'{(pnet*(1-desc_pct/100))*0.21:.2f} €', size=10, align='RIGHT')],
        [p(t['total_pvp_iva'], bold=True, size=11, color=GREEN),
         p(f'{pfin:.2f} €', bold=True, size=14, color=GREEN, align='RIGHT')],
        [p(t['entrega'], bold=True, size=9, color=colors.HexColor("#6B6860")),
         p(f'{pent:.2f} €', size=10, align='RIGHT')],
        [p(t['pendent_cobrar'], bold=True, size=11, color=RED),
         p(f'{ppend:.2f} €', bold=True, size=14, color=RED, align='RIGHT')],
    ]
    t3 = Table(t3_data, colWidths=[W*0.6, W*0.4])
    # Find total row index (index 2 or 3 depending on discount)
    total_idx = 3 if desc_pct > 0 else 2
    pend_idx  = total_idx + 2
    t3.setStyle(TableStyle([
        ('BACKGROUND',(0,total_idx),(-1,total_idx), colors.HexColor("#E8F3EE")),
        ('BACKGROUND',(0,pend_idx), (-1,pend_idx),  colors.HexColor("#FAEAEA")),
        ('ROWBACKGROUNDS',(0,0),(-1,-1),[LIG, colors.white]),
        ('BOX',(0,0),(-1,-1),0.5,BRD),
        ('INNERGRID',(0,0),(-1,-1),0.3,BRD),
        ('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6),
        ('LEFTPADDING',(0,0),(-1,-1),8),('RIGHTPADDING',(0,0),(-1,-1),8),
        ('LINEABOVE',(0,total_idx),(-1,total_idx),1.5,colors.HexColor("#1A6B45")),
        ('LINEABOVE',(0,pend_idx), (-1,pend_idx), 1.5,colors.HexColor("#B84040")),
    ]))
    story.append(t3)
    story.append(Spacer(1, 6*mm))
    story.append(HRFlowable(width=W, thickness=0.5, color=BRD))
    story.append(Spacer(1, 2*mm))
    story.append(p('Objectiu Emmarcació · C/ Mare Molas, 26 · Reus', size=8,
                   color=colors.HexColor("#636E72"), align='CENTER'))

    doc.build(story)
    buf.seek(0)
    return buf



@app.route('/ajustos')
@login_required
def ajustos():
    u = query(
        'SELECT marge, marge_impressio, nom_empresa, nom_fiscal, fiscal_id, empresa_adreca, empresa_tel, margins_json, brand_color, brand_color_secondary, brand_color_menu FROM usuaris WHERE id=?',
        [session['user_id']],
        one=True,
    )
    marge_actual = float(u['marge']) if u and u['marge'] is not None else 60
    marge_imp = float(u['marge_impressio']) if u and u['marge_impressio'] is not None else 0
    if float(marge_actual).is_integer():
        marge_actual = int(marge_actual)
    if float(marge_imp).is_integer():
        marge_imp = int(marge_imp)
    margins = _load_user_commercial_margins(u)
    margin_entries = [
        {'key': 'general', 'label': 'General', 'description': 'Marge base per a productes generals i acabats que no tenen marge propi.', 'value': _format_margin_for_view(margins['general'])},
        {'key': 'frames', 'label': 'Marcs', 'description': 'Marge principal de la calculadora de marcs.', 'value': _format_margin_for_view(margins['frames'])},
        {'key': 'canvas', 'label': 'Llenços', 'description': 'S\'utilitza al privat per a llenços i fine art si no es defineix un altre marge.', 'value': _format_margin_for_view(margins['canvas'])},
        {'key': 'prints', 'label': 'Impressió fotogràfica', 'description': 'S\'aplica a còpia fotogràfica i serveix també de base per a acabats d\'impressió.', 'value': _format_margin_for_view(margins['prints'])},
        {'key': 'foam', 'label': 'Foam', 'description': 'Permet separar el marge de foam del de la impressió si ho necessites.', 'value': _format_margin_for_view(margins['foam'])},
        {'key': 'laminate_foam', 'label': 'Laminat + foam', 'description': 'Per si voleu treballar aquesta combinació amb un marge propi.', 'value': _format_margin_for_view(margins['laminate_foam'])},
        {'key': 'fine_art', 'label': 'Fine art', 'description': 'Marge específic per a papers fine art i treballs més cuidats.', 'value': _format_margin_for_view(margins['fine_art'])},
        {'key': 'albums', 'label': 'Àlbums', 'description': 'Preparat per quan l\'àrea privada també gestioni àlbums amb el mateix compte.', 'value': _format_margin_for_view(margins['albums'])},
    ]
    nom_emp = u['nom_empresa'] if u and u['nom_empresa'] else ''
    brand_color = _normalize_hex_color(_row_get(u, 'brand_color', DEFAULT_BRAND_COLOR))
    brand_color_secondary = _normalize_hex_color(
        _row_get(u, 'brand_color_secondary', DEFAULT_BRAND_SECONDARY_COLOR),
        DEFAULT_BRAND_SECONDARY_COLOR,
    )
    brand_color_menu = _normalize_hex_color(
        _row_get(u, 'brand_color_menu', brand_color),
        brand_color,
    )
    cfg_rows = query("SELECT clau, valor FROM config WHERE clau LIKE 'empresa_%'")
    cfg = {r['clau']: r['valor'] for r in (cfg_rows or [])}
    if not nom_emp:
        nom_emp = cfg.get('empresa_nom', '')
    user_adreca  = _row_get(u, 'empresa_adreca', '') or ''
    user_tel     = _row_get(u, 'empresa_tel', '') or ''
    nom_fiscal   = _row_get(u, 'nom_fiscal', '') or ''
    fiscal_id    = _row_get(u, 'fiscal_id', '') or ''
    return render_template('ajustos.html', marge_actual=marge_actual, marge_imp=marge_imp,
                           margin_entries=[
                               dict(entry, description='S\'aplica a la fotografia impresa. Foam, laminat + foam i ProEco treballen amb el marge general.') if entry['key'] == 'prints' else entry
                               for entry in margin_entries if entry['key'] not in ('foam', 'laminate_foam')
                           ],
                           nom_empresa=nom_emp,
                           nom_fiscal=nom_fiscal,
                           fiscal_id=fiscal_id,
                           brand_color=brand_color,
                           brand_color_secondary=brand_color_secondary,
                           brand_color_menu=brand_color_menu,
                           empresa_adreca=user_adreca if user_adreca else cfg.get('empresa_adreca',''),
                           empresa_tel=user_tel if user_tel else cfg.get('empresa_tel',''))


import re as _re

def _parse_dims(ref):
    m = _re.search(r'(\d+)[xX](\d+)', ref)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = _re.search(r'[A-Z]+(\d+)$', ref)
    if m:
        d = m.group(1); mid = len(d)//2
        return int(d[:mid]), int(d[mid:])
    return None, None

def _fmt_measure(value):
    try:
        num = float(value or 0)
    except (TypeError, ValueError):
        return ''
    if num.is_integer():
        return str(int(num))
    return f"{num:.1f}".rstrip('0').rstrip('.')

def _format_size_text(w, h, sep=' × ', with_unit=False):
    if w in (None, '') or h in (None, ''):
        return ''
    try:
        wf = float(w)
        hf = float(h)
    except (TypeError, ValueError):
        return ''
    if wf <= 0 or hf <= 0:
        return ''
    text = f"{_fmt_measure(wf)}{sep}{_fmt_measure(hf)}"
    return text + (' cm' if with_unit else '')

def _photo_size_text(c, sep=' × ', with_unit=False):
    return _format_size_text(c.get('amplada'), c.get('alcada'), sep=sep, with_unit=with_unit)

def _final_size_text(c, sep=' × ', with_unit=False):
    text = _format_size_text(c.get('final_amplada'), c.get('final_alcada'), sep=sep, with_unit=with_unit)
    if text:
        return text
    ref = (c.get('passpartout') or '').strip()
    if ref and ref not in ('-', 'CONSERVAR'):
        w, h = _parse_dims(ref)
        text = _format_size_text(w, h, sep=sep, with_unit=with_unit)
        if text:
            return text
    return _photo_size_text(c, sep=sep, with_unit=with_unit)

def _display_piece_type(c, t):
    value = (c.get('tipus_peca') or '').strip().lower()
    if not value:
        return ''
    labels = {
        'fotografia': t['piece_photo'],
        'lamina': t['piece_lamina'],
        'pintura_sense_bastidor': t['piece_painting_unstretched'],
        'pintura_amb_bastidor': t['piece_painting_stretched'],
        'puzzle': t['piece_puzzle'],
        'punt_de_creu': t['piece_cross_stitch'],
    }
    return labels.get(value, value.replace('_', ' ').title())


def _display_piece_detail(c, lang='ca'):
    lang = (lang or 'ca').lower()
    piece_type = (c.get('tipus_peca') or '').strip().lower()
    detail = (c.get('tipus_peca_detall') or '').strip().lower()
    labels = {
        'ca': {
            ('fotografia', 'client'): 'La porta el client',
            ('fotografia', 'laboratori'): 'La imprimeix el laboratori',
            ('puzzle', 'sobre_base'): 'Ja va sobre una base',
            ('puzzle', 'sense_base'): 'Cal afegir suport abans d’emmarcar-lo',
            ('pintura_sense_bastidor', ''): 'Es resol amb encolat i es pot protegir amb vidre',
            ('pintura_amb_bastidor', ''): 'Pot portar vidre si es vol protegir la peça',
        },
        'es': {
            ('fotografia', 'client'): 'La trae el cliente',
            ('fotografia', 'laboratori'): 'La imprime el laboratorio',
            ('puzzle', 'sobre_base'): 'Ya va sobre una base',
            ('puzzle', 'sense_base'): 'Hay que añadir soporte antes de enmarcarlo',
            ('pintura_sense_bastidor', ''): 'Se trabaja con encolado y puede protegerse con vidrio',
            ('pintura_amb_bastidor', ''): 'Puede llevar vidrio si se quiere proteger la pieza',
        },
        'en': {
            ('fotografia', 'client'): 'Provided by the client',
            ('fotografia', 'laboratori'): 'Printed by the lab',
            ('puzzle', 'sobre_base'): 'Already mounted on a backing board',
            ('puzzle', 'sense_base'): 'A backing support must be added before framing',
            ('pintura_sense_bastidor', ''): 'Mounted with adhesive and optionally protected with glass',
            ('pintura_amb_bastidor', ''): 'Glass can be added if the piece needs protection',
        },
    }
    current = labels.get(lang, labels['ca'])
    return current.get((piece_type, detail)) or current.get((piece_type, '')) or ''

def _display_muntatge(c, t):
    ref = (c.get('encolat') or '').strip().upper()
    if not ref or ref == '-':
        return ''
    if ref.startswith('LAM'):
        return t['laminat_label']
    if ref.startswith('PRO'):
        return t['protter']
    return t['encolat_label']

def _display_proteccio(c, t):
    ref = (c.get('vidre') or '').strip().upper()
    if not ref or ref == '-':
        return ''
    if ref == 'CONSERVAR':
        return t['conservar_vidre']
    if ref.startswith('MIR'):
        return t['mirall']
    if ref.startswith('DV-'):
        return t['doble_vidre']
    return t['vidre_label']

def _display_interior(c, t):
    ref = (c.get('passpartout') or '').strip().upper()
    if not ref or ref == '-':
        return ''
    if ref.startswith('DOBPAS'):
        return t['doble_passpartu']
    if ref.startswith('PROECO'):
        return t['proeco_label']
    return t['passpartu_label']

def _display_revers_peu(c, t):
    raw = c.get('revers_peu')
    if isinstance(raw, bool):
        enabled = raw
    else:
        enabled = str(raw or '').strip().lower() in ('1', 'true', 'yes', 'on')
    if not enabled:
        return ''
    price = _safe_float(c.get('revers_peu_preu'), 0.0)
    return f'{price:.2f} €'

def _display_impressio(c, t):
    ref = (c.get('impressio') or '').strip()
    if not ref or ref == '-':
        return ''
    return t['inclosa']

def _find_closest(rows, w, h, prefix=None):
    """Among all sizes that contain (w, h), return the cheapest one.
    Ties broken by smallest perimeter overshoot to avoid unnecessary waste.
    This guarantees a larger frame never costs more than a smaller one due to
    non-monotonic price tables."""
    candidates = []
    for row in rows:
        ref = row['referencia']
        if prefix and not ref.upper().startswith(prefix.upper()):
            continue
        rw, rh = _parse_dims(ref)
        if rw is None: continue
        best_ov = None
        for fw, fh in [(rw,rh),(rh,rw)]:
            if fw >= w and fh >= h:
                ov = (fw-w)+(fh-h)
                if best_ov is None or ov < best_ov:
                    best_ov = ov
        if best_ov is not None:
            candidates.append((float(row.get('preu', 0) or 0), best_ov, dict(row)))
    if not candidates:
        return None
    # Sort by price first, then by overshoot to minimise waste among equally-priced options
    candidates.sort(key=lambda x: (x[0], x[1]))
    return candidates[0][2]

def _find_closest_area(rows, w, h, prefix=None, exclude_multi=None):
    """Closest by surface area — works for all products"""
    target = w * h
    best, best_diff = None, float('inf')
    for row in rows:
        ref = row['referencia']
        if prefix and not ref.upper().startswith(prefix.upper()): continue
        if exclude_multi and any(ref.upper().startswith(e.upper()) for e in exclude_multi): continue
        rw, rh = _parse_dims(ref)
        if rw is None: continue
        diff = abs(rw * rh - target)
        if diff < best_diff:
            best_diff = diff
            best = dict(row)
    return best

def _find_min_contain(rows, w, h, prefix=None):
    """Among all formats that physically contain (w, h), return the cheapest.
    Ties broken by smallest area to minimise waste.
    Fallback: largest available if nothing fits."""
    candidates = []
    for row in rows:
        ref = row['referencia']
        if prefix and not ref.upper().startswith(prefix.upper()): continue
        rw, rh = _parse_dims(ref)
        if rw is None: continue
        fits = (rw >= w and rh >= h) or (rh >= w and rw >= h)
        if fits:
            candidates.append((float(row.get('preu', 0) or 0), rw * rh, row))
    if candidates:
        candidates.sort(key=lambda x: (x[0], x[1]))
        return candidates[0][2]
    # Fallback: if nothing contains it, take the largest available
    return max(rows, key=lambda r: (_parse_dims(r['referencia'])[0] or 0) * (_parse_dims(r['referencia'])[1] or 0), default=None)

def _imp_closest(fw, fh):
    rows = [dict(r) for r in query('SELECT * FROM impressio')]
    r = _find_min_contain(rows, fw, fh)
    if r:
        return {'ref': r['referencia'], 'preu': r.get('preu', 0)}
    return None


def _laminate_only_closest(fw, fh):
    rows = []
    for ref, price in LAMINATE_ONLY_PRICES.items():
        rows.append({'referencia': ref, 'preu': float(price)})
    r = _find_min_contain(rows, fw, fh)
    if r:
        return {'ref': r['referencia'], 'preu': r.get('preu', 0)}
    return None

@app.route('/api/closest')
@login_required
def api_closest():
    w = float(request.args.get('w', 0))
    h = float(request.args.get('h', 0))
    foto_w = float(request.args.get('foto_w', w))
    foto_h = float(request.args.get('foto_h', h))
    tipus_laminat = request.args.get('laminat', 'semibrillo')  # 'semibrillo' | 'mate'
    if w <= 0 or h <= 0:
        return jsonify({})

    def _build_result(r, preu_col='preu'):
        """Build a result dict with preu_cost if available."""
        res = {'ref': r['referencia'], 'preu': r.get(preu_col, 0)}
        if r.get('preu_cost') is not None:
            res['preu_cost'] = float(r['preu_cost'])
        return res

    def closest(table, prefix=None, preu_col='preu', exclude_prefix=None, exclude_multi=None, w_override=None, h_override=None):
        rows = [dict(r) for r in query(f'SELECT * FROM {table}')]
        if exclude_multi:
            rows = [r for r in rows if not any(r['referencia'].upper().startswith(e.upper()) for e in exclude_multi)]
        elif exclude_prefix:
            rows = [r for r in rows if not r['referencia'].upper().startswith(exclude_prefix.upper())]
        uw = w_override if w_override is not None else w
        uh = h_override if h_override is not None else h
        r = _find_closest(rows, uw, uh, prefix)
        return _build_result(r, preu_col) if r else None

    def ca(table, fw, fh, prefix=None, exclude_multi=None, preu_col='preu'):
        """Closest by surface area for all products"""
        rows = [dict(r) for r in query(f'SELECT * FROM {table}')]
        r = _find_closest_area(rows, fw, fh, prefix=prefix, exclude_multi=exclude_multi)
        return _build_result(r, preu_col) if r else None

    # Tots els productes físics (vidre, passpartú, encolat, protter) usen
    # min-contain: la mida ha de cobrir físicament el marc. Això garanteix
    # que una peça més gran no pot sortir mai més barata que una de més petita.
    # Impressió: min-contain també (el paper ha de cobrir la foto).
    def cc(table, cw, ch, prefix=None, exclude_multi=None, preu_col='preu'):
        """Closest by min overshoot (must contain the size)"""
        rows = [dict(r) for r in query(f'SELECT * FROM {table}')]
        if exclude_multi:
            rows = [r for r in rows if not any(r['referencia'].upper().startswith(e.upper()) for e in exclude_multi)]
        r = _find_closest(rows, cw, ch, prefix)
        return _build_result(r, preu_col) if r else None

    def _pvd_result(res, categoria):
        """Add pvd field and update preu from preu_cost if available."""
        if res and res.get('preu_cost') is not None:
            pvd = calcular_pvd(res['preu_cost'], categoria)
            if pvd is not None:
                res['pvd'] = pvd
                res['preu'] = pvd  # backward compat: JS reads 'preu'
        return res

    def _fn_result(r):
        """Format foam/laminat/passpartu result for closest API."""
        return {'ref': r['ref'], 'preu': r['pvd'], 'pvd': r['pvd'], 'preu_cost': r['cost'], 'origen': r['origen']}

    # Foam (encolat simple). ProEco: àlies de compatibilitat històrica.
    foam = _fn_result(calcular_cost_foam(w, h))
    # Laminat sol (sense foam): retornem ambdós acabats
    laminat_semi = _fn_result(calcular_cost_laminat(w, h, tipus='semibrillo'))
    laminat_mate = _fn_result(calcular_cost_laminat(w, h, tipus='mate'))
    # Protter = foam + laminat: retornem ambdós acabats
    protter_semi = _fn_result(calcular_cost_protter(w, h, tipus='semibrillo'))
    protter_mate = _fn_result(calcular_cost_protter(w, h, tipus='mate'))
    protter_actual = protter_mate if tipus_laminat == 'mate' else protter_semi
    laminat_actual = laminat_mate if tipus_laminat == 'mate' else laminat_semi

    result = {
        'encolat':      foam,
        'proeco':       foam,  # àlies històric, mateix cost
        'protter':      protter_actual,
        'protter_semi': protter_semi,
        'protter_mate': protter_mate,
        'laminat':      laminat_actual,
        'laminat_semi': laminat_semi,
        'laminat_mate': laminat_mate,
        'vidre':        _fn_result(calcular_cost_vidre(w, h)),
        'doble_vidre':  _fn_result(calcular_cost_doble_vidre(w, h)),
        'mirall':       _fn_result(calcular_cost_mirall(w, h)),
        'passpartu':    _fn_result(calcular_cost_passpartu(w, h, tipus='simple')),
        'doble_pas':    _fn_result(calcular_cost_passpartu(w, h, tipus='doble')),
        'impressio':    _imp_closest(foto_w, foto_h),
    }
    return jsonify(result)

# â"€â"€ Routes: Email (mailto) â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
@app.route('/mailto-data', methods=['POST'])
@login_required
def mailto_data():
    d = request.json or {}
    cid = d.get('comanda_id')
    nom_comerc = d.get('nom_comerc', session.get('nom',''))
    c = _get_comanda_for_session(cid)
    if not c:
        return jsonify({'ok': False, 'error': 'No autoritzat'}), 403
    lang = (c.get('lang') or 'ca').lower()
    MAILTO_T = {
        'ca': {
            'greet': 'Bon dia,',
            'intro': "Us faig arribar el pressupost d'emmarcació per al client {client}.",
            'detail': 'DETALL DE LA COMANDA:',
            'mides': 'Mides peça',
            'mida_final': 'Mida final',
            'tipus_peca': 'Peça a emmarcar',
            'vidre': 'Vidre',
            'passpartout': 'Passpartout/ProEco',
            'encolat': 'Encolat',
            'revers_peu': 'Revers amb peu',
            'impressio': 'Impressió',
            'inclosa': 'Inclosa',
            'encolat_label': 'Encolat',
            'laminat_label': 'Laminat',
            'protter': 'Protter',
            'vidre_label': 'Vidre',
            'doble_vidre': 'Doble vidre',
            'mirall': 'Mirall',
            'passpartu_label': 'Passpartú',
            'doble_passpartu': 'Doble passpartú',
            'proeco_label': 'ProEco',
            'piece_photo': 'Fotografia',
            'piece_lamina': 'Làmina',
            'piece_painting_unstretched': 'Pintura sense bastidor',
            'piece_painting_stretched': 'Pintura amb bastidor',
            'piece_puzzle': 'Puzzle',
            'piece_cross_stitch': 'Punt de creu',
            'preu': 'PREU FINAL: {price:.2f} EUR (IVA inclòs)',
            'pendent': 'Pendent de cobrar: {pend:.2f} EUR',
            'obs': 'Observacions: {obs}',
            'attach': 'Trobareu el pressupost en PDF adjunt.',
            'bye': 'Atentament,',
            'subject': 'Pressupost emmarcació - {client}'
        },
        'es': {
            'greet': 'Buenos días,',
            'intro': 'Os envío el presupuesto de enmarcación para el cliente {client}.',
            'detail': 'DETALLE DEL PEDIDO:',
            'mides': 'Medidas pieza',
            'mida_final': 'Medida final',
            'tipus_peca': 'Pieza a enmarcar',
            'vidre': 'Vidrio',
            'passpartout': 'Passpartout/ProEco',
            'encolat': 'Montaje',
            'revers_peu': 'Reverso con pie',
            'impressio': 'Impresión',
            'inclosa': 'Incluida',
            'encolat_label': 'Encolado',
            'laminat_label': 'Laminado',
            'protter': 'Protter',
            'vidre_label': 'Vidrio',
            'doble_vidre': 'Doble vidrio',
            'mirall': 'Espejo',
            'passpartu_label': 'Passpartú',
            'doble_passpartu': 'Doble passpartú',
            'proeco_label': 'ProEco',
            'piece_photo': 'Fotografía',
            'piece_lamina': 'Lámina',
            'piece_painting_unstretched': 'Pintura sin bastidor',
            'piece_painting_stretched': 'Pintura con bastidor',
            'piece_puzzle': 'Puzzle',
            'piece_cross_stitch': 'Punto de cruz',
            'preu': 'PRECIO FINAL: {price:.2f} EUR (IVA incluido)',
            'pendent': 'Pendiente de cobro: {pend:.2f} EUR',
            'obs': 'Observaciones: {obs}',
            'attach': 'Encontraréis el presupuesto en PDF adjunto.',
            'bye': 'Atentamente,',
            'subject': 'Presupuesto enmarcación - {client}'
        },
        'en': {
            'greet': 'Good morning,',
            'intro': 'Please find the framing quote for client {client}.',
            'detail': 'QUOTE DETAILS:',
            'mides': 'Item size',
            'mida_final': 'Final size',
            'tipus_peca': 'Item to frame',
            'vidre': 'Glass',
            'passpartout': 'Passpartout/ProEco',
            'encolat': 'Mounting',
            'revers_peu': 'Backing with stand',
            'impressio': 'Print',
            'inclosa': 'Included',
            'encolat_label': 'Mounting',
            'laminat_label': 'Laminate only',
            'protter': 'Protter',
            'vidre_label': 'Glass',
            'doble_vidre': 'Double glass',
            'mirall': 'Mirror',
            'passpartu_label': 'Mat',
            'doble_passpartu': 'Double mat',
            'proeco_label': 'ProEco',
            'piece_photo': 'Photograph',
            'piece_lamina': 'Poster',
            'piece_painting_unstretched': 'Unstretched painting',
            'piece_painting_stretched': 'Stretched painting',
            'piece_puzzle': 'Puzzle',
            'piece_cross_stitch': 'Cross stitch',
            'preu': 'FINAL PRICE: {price:.2f} EUR (VAT included)',
            'pendent': 'Outstanding: {pend:.2f} EUR',
            'obs': 'Notes: {obs}',
            'attach': 'The quote PDF is attached.',
            'bye': 'Kind regards,',
            'subject': 'Framing quote - {client}'
        }
    }
    tt = MAILTO_T.get(lang, MAILTO_T['ca'])
    final_size = _final_size_text(c, sep=' x ', with_unit=True)
    photo_size = _photo_size_text(c, sep=' x ', with_unit=True)
    piece_type = _display_piece_type(c, tt)
    piece_detail = _display_piece_detail(c, lang)
    if piece_type and piece_detail:
        piece_type = f"{piece_type} · {piece_detail}"
    proteccio_label = _display_proteccio(c, tt)
    interior_label = _display_interior(c, tt)
    revers_peu_label = _display_revers_peu(c, tt)
    muntatge_label = _display_muntatge(c, tt)
    impressio_label = _display_impressio(c, tt)
    lines = [
        tt['greet'],
        "",
        tt['intro'].format(client=c['client_nom']),
        "",
        tt['detail'],
    ]
    if piece_type:
        lines.append(f"  {tt['tipus_peca']}: {piece_type}")
    if final_size:
        lines.append(f"  {tt['mida_final']}: {final_size}")
    if photo_size:
        lines.append(f"  {tt['mides']}: {photo_size}")
    if proteccio_label:
        lines.append(f"  {tt['vidre']}: {proteccio_label}")
    if interior_label:
        lines.append(f"  {tt['passpartout']}: {interior_label}")
    if revers_peu_label:
        lines.append(f"  {tt['revers_peu']}: {revers_peu_label}")
    if muntatge_label:
        lines.append(f"  {tt['encolat']}: {muntatge_label}")
    if impressio_label:
        lines.append(f"  {tt['impressio']}: {tt['inclosa']}")
    lines += [
        "",
        tt['preu'].format(price=c['preu_final']),
    ]
    if c['pendent'] and float(c['pendent']) > 0.01:
        lines.append(tt['pendent'].format(pend=c['pendent']))
    if c['observacions']:
        lines.append(tt['obs'].format(obs=c['observacions']))
    lines += [
        "",
        tt['attach'],
        "",
        tt['bye'],
        f"{nom_comerc}",
    ]
    body = "%0D%0A".join(lines)
    subject = tt['subject'].format(client=c['client_nom'])
    mailto = f"mailto:reusrevela@gmail.com?subject={subject}&body={body}"
    return jsonify({'ok': True, 'mailto': mailto})

@app.route('/enviar-email', methods=['POST'])
@login_required
def enviar_email():
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    d = request.json or {}
    cid = d.get('comanda_id')
    nom_comerc = d.get('nom_comerc', session.get('nom',''))
    nota = d.get('nota','')
    c = _get_comanda_for_session(cid)
    if not c:
        return jsonify({'ok': False, 'error': 'No autoritzat'}), 403
    cfg = {r['clau']: r['valor'] for r in query('SELECT * FROM config')}
    gmail_user = cfg.get('gmail_user','')
    gmail_pass = cfg.get('gmail_pass','')
    if not gmail_user or not gmail_pass:
        return jsonify({'ok': False, 'error': "Configura el Gmail a l'admin primer."})
    dest = 'reusrevela@gmail.com'
    obs = c['observacions'] or ''
    pdf_lang = PDF_T.get((c.get('lang') or 'ca').lower(), PDF_T['ca'])
    final_size = _final_size_text(c, sep=' x ', with_unit=True) or '-'
    photo_size = _photo_size_text(c, sep=' x ', with_unit=True) or '-'
    piece_type = _display_piece_type(c, pdf_lang)
    piece_detail = _display_piece_detail(c, (c.get('lang') or 'ca').lower())
    if piece_type and piece_detail:
        piece_type = f"{piece_type} · {piece_detail}"
    proteccio_label = _display_proteccio(c, pdf_lang) or '-'
    interior_label = _display_interior(c, pdf_lang) or '-'
    revers_peu_label = _display_revers_peu(c, pdf_lang) or '-'
    muntatge_label = _display_muntatge(c, pdf_lang) or '-'
    impressio_label = _display_impressio(c, pdf_lang) or '-'
    nota_html = f"<p style='font-family:sans-serif;color:#C8873A'><b>Nota:</b> {nota}</p>" if nota else ""
    html = f"""
    <h2 style='color:#1A6B45;font-family:sans-serif;border-bottom:2px solid #1A6B45;padding-bottom:8px'>
        Nou pressupost d'emmarcacio</h2>
    <p style='font-family:sans-serif;font-size:15px'><b>Comers:</b> {nom_comerc}</p>
    <p style='font-family:sans-serif;font-size:15px'><b>Client:</b> {c['client_nom']} &nbsp;·&nbsp; {c['client_tel'] or '-'}</p>
    <p style='font-family:sans-serif;font-size:15px'><b>Data:</b> {c['data']}</p>
    <hr style='border:1px solid #E5E2DB;margin:12px 0'>
    {"<p style='font-family:sans-serif;font-size:14px'><b>" + pdf_lang['tipus_peca'] + ":</b> " + piece_type + "</p>" if piece_type else ""}
    <p style='font-family:sans-serif;font-size:14px'><b>{pdf_lang['mida_final']}:</b> {final_size}</p>
    <p style='font-family:sans-serif;font-size:14px'><b>{pdf_lang['mides_foto']}:</b> {photo_size}</p>
    <p style='font-family:sans-serif;font-size:14px'><b>{pdf_lang['muntatge']}:</b> {muntatge_label}</p>
    <p style='font-family:sans-serif;font-size:14px'><b>{pdf_lang['vidre_mirall']}:</b> {proteccio_label}</p>
    <p style='font-family:sans-serif;font-size:14px'><b>{pdf_lang['interior']}:</b> {interior_label}</p>
    <p style='font-family:sans-serif;font-size:14px'><b>{pdf_lang['revers_peu']}:</b> {revers_peu_label}</p>
    <p style='font-family:sans-serif;font-size:14px'><b>{pdf_lang['impressio']}:</b> {impressio_label}</p>
    {"<p style='font-family:sans-serif;font-size:14px'><b>Obs:</b> " + obs + "</p>" if obs else ""}
    <hr style='border:1px solid #E5E2DB;margin:12px 0'>
    <p style='font-family:sans-serif;font-size:16px'><b>Preu final:</b>
        <span style='color:#1A6B45;font-size:22px;font-weight:bold'>{c['preu_final']:.2f} EUR</span></p>
    <p style='font-family:sans-serif;font-size:14px'><b>Pendent:</b> {c['pendent']:.2f} EUR</p>
    {nota_html}
    <p style='font-family:sans-serif;font-size:12px;color:#9E9B94;margin-top:24px'>
        Enviat des de l'aplicacio Objectiu Emmarcacio</p>
    """
    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = f"Pressupost #{cid} - {nom_comerc} - {c['client_nom']}"
        msg['From'] = gmail_user
        msg['To'] = dest
        msg.attach(MIMEText(html, 'html'))
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
            s.login(gmail_user, gmail_pass)
            s.sendmail(gmail_user, dest, msg.as_string())
        return jsonify({'ok': True})
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)})

# â"€â"€ Init DB â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€â"€
def init_db():
    with app.app_context():
        db = get_db()
        if USE_PG:
            # Use a SEPARATE connection with autocommit for DDL.
            # Timeouts evitten que un bloqueig a PG penji l'arrencada del worker.
            import psycopg2 as _pg2
            ddl_conn = _pg2.connect(
                DATABASE_URL,
                connect_timeout=10,
                options='-c statement_timeout=30000',
            )
            ddl_conn.autocommit = True
            ddl_cur = ddl_conn.cursor()
            ddl = [
                """CREATE TABLE IF NOT EXISTS usuaris (
                    id SERIAL PRIMARY KEY, username TEXT UNIQUE NOT NULL,
                    password TEXT NOT NULL, nom TEXT NOT NULL,
                    is_admin INTEGER DEFAULT 0, marge REAL DEFAULT 60,
                    marge_impressio REAL DEFAULT 100, nom_empresa TEXT DEFAULT '',
                    brand_color TEXT DEFAULT '#1A6B45',
                    brand_color_secondary TEXT DEFAULT '#C8873A',
                    brand_color_menu TEXT DEFAULT '#1A6B45',
                    margins_json TEXT DEFAULT '',
                    access_status TEXT DEFAULT 'active',
                    profile_type TEXT DEFAULT 'professional',
                    web_url TEXT DEFAULT '',
                    instagram TEXT DEFAULT '',
                    fiscal_id TEXT DEFAULT '',
                    notes_validacio TEXT DEFAULT ''
                )""",
                """CREATE TABLE IF NOT EXISTS impressio (
                    referencia TEXT PRIMARY KEY, descripcio TEXT, preu REAL)""",
                """CREATE TABLE IF NOT EXISTS moldures (
                    referencia TEXT PRIMARY KEY, preu_taller REAL, gruix REAL,
                    cost REAL, proveidor TEXT, ref2 TEXT, ubicacio TEXT,
                    descripcio TEXT, foto TEXT)""",
                """CREATE TABLE IF NOT EXISTS vidres (
                    referencia TEXT PRIMARY KEY, preu REAL)""",
                """CREATE TABLE IF NOT EXISTS passpartout (
                    referencia TEXT PRIMARY KEY, preu REAL)""",
                """CREATE TABLE IF NOT EXISTS encolat_pro (
                    referencia TEXT PRIMARY KEY, preu REAL)""",
                """CREATE TABLE IF NOT EXISTS proeco (
                    referencia TEXT PRIMARY KEY, preu REAL)""",
                """CREATE TABLE IF NOT EXISTS comandes (
                    id SERIAL PRIMARY KEY, user_id INTEGER, data TEXT,
                    client_nom TEXT, client_tel TEXT, pre_marc TEXT,
                    marc_principal TEXT, amplada REAL, alcada REAL, copia REAL,
                    encolat TEXT, vidre TEXT, passpartout TEXT, impressio TEXT,
                    revers_peu INTEGER DEFAULT 0,
                    revers_peu_preu REAL DEFAULT 0,
                    tipus_peca TEXT DEFAULT '',
                    tipus_peca_detall TEXT DEFAULT '',
                    final_amplada REAL DEFAULT 0, final_alcada REAL DEFAULT 0,
                    marge REAL, descompte REAL, quantitat REAL,
                    preu_net REAL, preu_final REAL, entrega REAL, pendent REAL,
                    observacions TEXT, sessio_id TEXT,
                    opcio_nom TEXT DEFAULT 'Opció A')""",
                """CREATE TABLE IF NOT EXISTS config (
                    clau TEXT PRIMARY KEY, valor TEXT)""",
            ]
            for s in ddl:
                try:
                    ddl_cur.execute(s)
                    print("DDL OK:", s[:40])
                except Exception as e:
                    print("DDL err:", e)
            for tbl, col, typ in [
                ('comandes','impressio','TEXT'),
                ('comandes','revers_peu','INTEGER DEFAULT 0'),
                ('comandes','revers_peu_preu','REAL DEFAULT 0'),
                ('comandes','tipus_peca',"TEXT DEFAULT ''"),
                ('comandes','tipus_peca_detall',"TEXT DEFAULT ''"),
                ('comandes','final_amplada','REAL DEFAULT 0'),
                ('comandes','final_alcada','REAL DEFAULT 0'),
                ('comandes','sessio_id','TEXT'),
                ('comandes','opcio_nom','TEXT'),
                ('usuaris','nom_empresa',"TEXT DEFAULT ''"),
                ('usuaris','empresa_adreca',"TEXT DEFAULT ''"),
                ('usuaris','empresa_tel',"TEXT DEFAULT ''"),
                ('usuaris','brand_color',"TEXT DEFAULT '#1A6B45'"),
                ('usuaris','brand_color_secondary',"TEXT DEFAULT '#C8873A'"),
                ('usuaris','brand_color_menu',"TEXT DEFAULT '#1A6B45'"),
                ('usuaris','margins_json',"TEXT DEFAULT ''"),
                ('usuaris','setup_done','INTEGER DEFAULT 0'),
                ('usuaris','logo_b64','TEXT'),
                ('usuaris','marge_impressio_setup','INTEGER DEFAULT 0'),
                ('usuaris','access_status',"TEXT DEFAULT 'active'"),
                ('usuaris','profile_type',"TEXT DEFAULT 'professional'"),
                ('usuaris','web_url',"TEXT DEFAULT ''"),
                ('usuaris','instagram',"TEXT DEFAULT ''"),
                ('usuaris','fiscal_id',"TEXT DEFAULT ''"),
                ('usuaris','notes_validacio',"TEXT DEFAULT ''"),
                ('comandes','num_pressupost','TEXT'),
                ('comandes','pagat','INTEGER DEFAULT 0'),
                ('comandes','entregat','INTEGER DEFAULT 0'),
                ('comandes','lang','TEXT DEFAULT \'ca\''),
                ('comandes','foto_comanda','TEXT'),
                ('comandes','foto_ts','REAL'),
                ('comandes','passpartu_ref',"TEXT DEFAULT ''"),
                ('comandes','cost_produccio','REAL DEFAULT 0'),
                ('comandes','fd_albara',"TEXT DEFAULT ''"),
                ('usuaris','nom_fiscal',"TEXT DEFAULT ''"),
                # --- v2 price model: cost columns ---
                ('moldures','preu_cost','REAL'),
                ('moldures','merma_pct','REAL DEFAULT 10.0'),
                ('moldures','minim_cm','REAL DEFAULT 100.0'),
                ('moldures','preu_cost_ant','REAL'),
                ('moldures','data_cost','TEXT'),
                ('moldures','usuari_cost_id','INTEGER'),
                ('moldures','notes_cost','TEXT'),
                ('moldures','cost_verificat','INTEGER DEFAULT 0'),
                ('vidres','preu_cost','REAL'),
                ('vidres','preu_cost_ant','REAL'),
                ('vidres','data_cost','TEXT'),
                ('vidres','usuari_cost_id','INTEGER'),
                ('vidres','notes_cost','TEXT'),
                ('vidres','cost_verificat','INTEGER DEFAULT 0'),
                ('encolat_pro','preu_cost','REAL'),
                ('encolat_pro','preu_cost_ant','REAL'),
                ('encolat_pro','data_cost','TEXT'),
                ('encolat_pro','usuari_cost_id','INTEGER'),
                ('encolat_pro','notes_cost','TEXT'),
                ('encolat_pro','cost_verificat','INTEGER DEFAULT 0'),
                ('passpartout','preu_cost','REAL'),
                ('passpartout','preu_cost_ant','REAL'),
                ('passpartout','data_cost','TEXT'),
                ('passpartout','usuari_cost_id','INTEGER'),
                ('passpartout','notes_cost','TEXT'),
                ('passpartout','cost_verificat','INTEGER DEFAULT 0'),
                # --- v2 price model: order snapshots ---
                ('comandes','cost_unitari','REAL'),
                ('comandes','pvd_unitari','REAL'),
                ('comandes','marge_admin_snap','REAL'),
                ('comandes','marge_pro_snap','REAL'),
                # --- v2 price model: user margin aliases ---
                ('usuaris','marge_pro_pct','REAL DEFAULT 60'),
                ('usuaris','marge_impressio_pro_pct','REAL DEFAULT 100'),
            ]:
                try:
                    ddl_cur.execute(f"ALTER TABLE {tbl} ADD COLUMN IF NOT EXISTS {col} {typ}")
                except Exception as e:
                    print("ALTER skip:", e)
            # --- v2: historial_preus_cost audit table ---
            try:
                ddl_cur.execute("""CREATE TABLE IF NOT EXISTS historial_preus_cost (
                    id SERIAL PRIMARY KEY,
                    taula TEXT NOT NULL,
                    referencia TEXT NOT NULL,
                    preu_cost_antic REAL,
                    preu_cost_nou REAL,
                    usuari_id INTEGER,
                    data TEXT,
                    notes TEXT
                )""")
            except Exception as e:
                print("historial_preus_cost skip:", e)
            try:
                ddl_cur.execute("""CREATE TABLE IF NOT EXISTS feedback (
                    id SERIAL PRIMARY KEY,
                    user_id INTEGER,
                    tipus TEXT NOT NULL DEFAULT 'millora',
                    missatge TEXT NOT NULL,
                    pagina TEXT,
                    data TEXT,
                    llegit INTEGER DEFAULT 0
                )""")
            except Exception as e:
                print("feedback table skip:", e)
            ddl_cur.close()
            ddl_conn.close()
            print("DDL done, checking admin...")
            # Now use normal connection for DML
            cur = db.cursor()
            cur.execute("INSERT INTO config (clau,valor) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                       ['marge_defecte','60'])
            cur.execute("INSERT INTO config (clau,valor) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                       ['empresa_nom','Reus Revela'])
            cur.execute("INSERT INTO config (clau,valor) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                       ['empresa_adreca',''])
            cur.execute("INSERT INTO config (clau,valor) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                       ['empresa_tel',''])
            # --- v2 price model: admin margin config ---
            for k, v in [('marge_admin_moldures_pct','60'), ('marge_admin_vidres_pct','60'),
                         ('marge_admin_passpartu_pct','60'), ('marge_admin_encolat_pct','60'),
                         ('marge_pvp_suggerit_pct','40'),
                         ('marge_pro_actiu','1'),
                         ('cost_hora_taller','25.00'),
                         ('passpartu_temps_simple','9'),
                         ('passpartu_temps_doble','16'),
                         ('passpartu_temps_finestra','3.5'),
                         ('passpartu_cost_cm2','0.000620'),
                         ('passpartu_minim_material','0.50'),
                         ('foam_cost_cm2','0.001143'),
                         ('foam_temps_base_min','9'),
                         ('foam_temps_var_cm2','0.0015'),
                         ('laminat_semibrillo_cost_cm2','0.000504'),
                         ('laminat_mate_cost_cm2','0.000685'),
                         ('laminat_temps_base_min','12'),
                         ('laminat_temps_var_cm2','0.0012'),
                         ('vidre_cost_cm2','0.002880'),
                         ('vidre_temps_base_min','3'),
                         ('vidre_temps_lineal_m','0.5'),
                         ('vidre_dv_muntatge_min','5'),
                         ('mirall_cost_cm2','0.003153'),
                         ('mirall_multiplo_dm2','6')]:
                cur.execute("INSERT INTO config (clau,valor) VALUES (%s,%s) ON CONFLICT DO NOTHING", [k, v])
            db.commit()
            _seed_admin_if_configured(db)
            # NOTA: tota operació pesada (_seed_proeco_preus, _seed_intermol_moldures,
            # _fix_ref2_errors, _run_v2_price_backfill) s'invoca ara només des de
            # /admin/run-migrations. init_db() es queda amb CREATE/ALTER + config seeds
            # + admin bootstrap perquè l'arrencada del worker no penji a PG.
            db.commit()
        else:
            db.executescript('''
                CREATE TABLE IF NOT EXISTS usuaris (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password TEXT NOT NULL,
                    nom TEXT NOT NULL,
                    is_admin INTEGER DEFAULT 0,
                    marge REAL DEFAULT 60,
                    marge_impressio REAL DEFAULT 100,
                    nom_empresa TEXT DEFAULT '',
                    brand_color TEXT DEFAULT '#1A6B45',
                    brand_color_secondary TEXT DEFAULT '#C8873A',
                    brand_color_menu TEXT DEFAULT '#1A6B45',
                    margins_json TEXT DEFAULT '',
                    access_status TEXT DEFAULT 'active',
                    profile_type TEXT DEFAULT 'professional',
                    web_url TEXT DEFAULT '',
                    instagram TEXT DEFAULT '',
                    fiscal_id TEXT DEFAULT '',
                    notes_validacio TEXT DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS impressio (
                    referencia TEXT PRIMARY KEY, descripcio TEXT, preu REAL
                );
                CREATE TABLE IF NOT EXISTS moldures (
                    referencia TEXT PRIMARY KEY,
                    preu_taller REAL, gruix REAL, cost REAL,
                    proveidor TEXT, ref2 TEXT, ubicacio TEXT,
                    descripcio TEXT, foto TEXT
                );
                CREATE TABLE IF NOT EXISTS vidres (
                    referencia TEXT PRIMARY KEY, preu REAL
                );
                CREATE TABLE IF NOT EXISTS passpartout (
                    referencia TEXT PRIMARY KEY, preu REAL
                );
                CREATE TABLE IF NOT EXISTS encolat_pro (
                    referencia TEXT PRIMARY KEY, preu REAL
                );
                CREATE TABLE IF NOT EXISTS proeco (
                    referencia TEXT PRIMARY KEY, preu REAL
                );
                CREATE TABLE IF NOT EXISTS comandes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER, data TEXT,
                    client_nom TEXT, client_tel TEXT,
                    pre_marc TEXT, marc_principal TEXT,
                    amplada REAL, alcada REAL, copia REAL,
                    encolat TEXT, vidre TEXT, passpartout TEXT, impressio TEXT,
                    revers_peu INTEGER DEFAULT 0,
                    revers_peu_preu REAL DEFAULT 0,
                    tipus_peca TEXT DEFAULT '',
                    final_amplada REAL DEFAULT 0, final_alcada REAL DEFAULT 0,
                    marge REAL, descompte REAL, quantitat REAL,
                    preu_net REAL, preu_final REAL,
                    entrega REAL, pendent REAL, observacions TEXT,
                    sessio_id TEXT, opcio_nom TEXT DEFAULT 'Opció A'
                );
                CREATE TABLE IF NOT EXISTS config (
                    clau TEXT PRIMARY KEY, valor TEXT
                );
            ''')
            db.commit()
            for sql in [
                "ALTER TABLE usuaris ADD COLUMN setup_done INTEGER DEFAULT 0",
                "ALTER TABLE usuaris ADD COLUMN logo_b64 TEXT",
                "ALTER TABLE usuaris ADD COLUMN brand_color TEXT DEFAULT '#1A6B45'",
                "ALTER TABLE usuaris ADD COLUMN brand_color_secondary TEXT DEFAULT '#C8873A'",
                "ALTER TABLE usuaris ADD COLUMN brand_color_menu TEXT DEFAULT '#1A6B45'",
                "ALTER TABLE usuaris ADD COLUMN margins_json TEXT DEFAULT ''",
                "ALTER TABLE usuaris ADD COLUMN marge_impressio_setup INTEGER DEFAULT 0",
                "ALTER TABLE usuaris ADD COLUMN access_status TEXT DEFAULT 'active'",
                "ALTER TABLE usuaris ADD COLUMN profile_type TEXT DEFAULT 'professional'",
                "ALTER TABLE usuaris ADD COLUMN web_url TEXT DEFAULT ''",
                "ALTER TABLE usuaris ADD COLUMN instagram TEXT DEFAULT ''",
                "ALTER TABLE usuaris ADD COLUMN fiscal_id TEXT DEFAULT ''",
                "ALTER TABLE usuaris ADD COLUMN nom_fiscal TEXT DEFAULT ''",
                "ALTER TABLE usuaris ADD COLUMN notes_validacio TEXT DEFAULT ''",
                "ALTER TABLE comandes ADD COLUMN num_pressupost TEXT",
                "ALTER TABLE comandes ADD COLUMN revers_peu INTEGER DEFAULT 0",
                "ALTER TABLE comandes ADD COLUMN revers_peu_preu REAL DEFAULT 0",
                "ALTER TABLE comandes ADD COLUMN pagat INTEGER DEFAULT 0",
                "ALTER TABLE comandes ADD COLUMN entregat INTEGER DEFAULT 0",
                "ALTER TABLE comandes ADD COLUMN lang TEXT DEFAULT 'ca'",
                "ALTER TABLE comandes ADD COLUMN foto_comanda TEXT",
                "ALTER TABLE comandes ADD COLUMN foto_ts REAL",
                "ALTER TABLE comandes ADD COLUMN passpartu_ref TEXT DEFAULT ''",
                "ALTER TABLE comandes ADD COLUMN cost_produccio REAL DEFAULT 0",
                "ALTER TABLE comandes ADD COLUMN fd_albara TEXT DEFAULT ''",
                "ALTER TABLE passpartout ADD COLUMN color TEXT DEFAULT ''",
                "ALTER TABLE passpartout ADD COLUMN textura TEXT DEFAULT ''",
                "ALTER TABLE passpartout ADD COLUMN descripcio TEXT DEFAULT ''",
                # --- v2 price model: cost columns ---
                "ALTER TABLE moldures ADD COLUMN preu_cost REAL",
                "ALTER TABLE moldures ADD COLUMN merma_pct REAL DEFAULT 10.0",
                "ALTER TABLE moldures ADD COLUMN minim_cm REAL DEFAULT 100.0",
                "ALTER TABLE moldures ADD COLUMN preu_cost_ant REAL",
                "ALTER TABLE moldures ADD COLUMN data_cost TEXT",
                "ALTER TABLE moldures ADD COLUMN usuari_cost_id INTEGER",
                "ALTER TABLE moldures ADD COLUMN notes_cost TEXT",
                "ALTER TABLE moldures ADD COLUMN cost_verificat INTEGER DEFAULT 0",
                "ALTER TABLE vidres ADD COLUMN preu_cost REAL",
                "ALTER TABLE vidres ADD COLUMN preu_cost_ant REAL",
                "ALTER TABLE vidres ADD COLUMN data_cost TEXT",
                "ALTER TABLE vidres ADD COLUMN usuari_cost_id INTEGER",
                "ALTER TABLE vidres ADD COLUMN notes_cost TEXT",
                "ALTER TABLE vidres ADD COLUMN cost_verificat INTEGER DEFAULT 0",
                "ALTER TABLE encolat_pro ADD COLUMN preu_cost REAL",
                "ALTER TABLE encolat_pro ADD COLUMN preu_cost_ant REAL",
                "ALTER TABLE encolat_pro ADD COLUMN data_cost TEXT",
                "ALTER TABLE encolat_pro ADD COLUMN usuari_cost_id INTEGER",
                "ALTER TABLE encolat_pro ADD COLUMN notes_cost TEXT",
                "ALTER TABLE encolat_pro ADD COLUMN cost_verificat INTEGER DEFAULT 0",
                "ALTER TABLE passpartout ADD COLUMN preu_cost REAL",
                "ALTER TABLE passpartout ADD COLUMN preu_cost_ant REAL",
                "ALTER TABLE passpartout ADD COLUMN data_cost TEXT",
                "ALTER TABLE passpartout ADD COLUMN usuari_cost_id INTEGER",
                "ALTER TABLE passpartout ADD COLUMN notes_cost TEXT",
                "ALTER TABLE passpartout ADD COLUMN cost_verificat INTEGER DEFAULT 0",
                # --- v2 price model: order snapshots ---
                "ALTER TABLE comandes ADD COLUMN cost_unitari REAL",
                "ALTER TABLE comandes ADD COLUMN pvd_unitari REAL",
                "ALTER TABLE comandes ADD COLUMN marge_admin_snap REAL",
                "ALTER TABLE comandes ADD COLUMN marge_pro_snap REAL",
                # --- v2 price model: user margin aliases ---
                "ALTER TABLE usuaris ADD COLUMN marge_pro_pct REAL DEFAULT 60",
                "ALTER TABLE usuaris ADD COLUMN marge_impressio_pro_pct REAL DEFAULT 100",
            ]:
                try:
                    db.execute(sql)
                except Exception:
                    pass
            db.commit()
            # --- v2: historial_preus_cost audit table ---
            db.execute("""CREATE TABLE IF NOT EXISTS historial_preus_cost (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                taula TEXT NOT NULL,
                referencia TEXT NOT NULL,
                preu_cost_antic REAL,
                preu_cost_nou REAL,
                usuari_id INTEGER,
                data TEXT,
                notes TEXT
            )""")
            db.execute("""CREATE TABLE IF NOT EXISTS feedback (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                tipus TEXT NOT NULL DEFAULT 'millora',
                missatge TEXT NOT NULL,
                pagina TEXT,
                data TEXT,
                llegit INTEGER DEFAULT 0
            )""")
            db.commit()
            db.execute("INSERT OR IGNORE INTO config (clau,valor) VALUES ('marge_defecte','60')")
            db.execute("INSERT OR IGNORE INTO config (clau,valor) VALUES ('empresa_nom','Reus Revela')")
            db.execute("INSERT OR IGNORE INTO config (clau,valor) VALUES ('empresa_adreca','')")
            db.execute("INSERT OR IGNORE INTO config (clau,valor) VALUES ('empresa_tel','')")
            # --- v2 price model: admin margin config ---
            for k, v in [('marge_admin_moldures_pct','60'), ('marge_admin_vidres_pct','60'),
                         ('marge_admin_passpartu_pct','60'), ('marge_admin_encolat_pct','60'),
                         ('marge_pvp_suggerit_pct','40'),
                         ('marge_pro_actiu','1'),
                         ('cost_hora_taller','25.00'),
                         ('passpartu_temps_simple','9'),
                         ('passpartu_temps_doble','16'),
                         ('passpartu_temps_finestra','3.5'),
                         ('passpartu_cost_cm2','0.000620'),
                         ('passpartu_minim_material','0.50'),
                         ('foam_cost_cm2','0.001143'),
                         ('foam_temps_base_min','9'),
                         ('foam_temps_var_cm2','0.0015'),
                         ('laminat_semibrillo_cost_cm2','0.000504'),
                         ('laminat_mate_cost_cm2','0.000685'),
                         ('laminat_temps_base_min','12'),
                         ('laminat_temps_var_cm2','0.0012'),
                         ('vidre_cost_cm2','0.002880'),
                         ('vidre_temps_base_min','3'),
                         ('vidre_temps_lineal_m','0.5'),
                         ('vidre_dv_muntatge_min','5'),
                         ('mirall_cost_cm2','0.003153'),
                         ('mirall_multiplo_dm2','6')]:
                db.execute("INSERT OR IGNORE INTO config (clau,valor) VALUES (?,?)", [k, v])
            db.commit()
            _seed_admin_if_configured(db)
            # Veure nota a la branca PG: tota operació pesada s'executa ara només
            # via /admin/run-migrations.
            db.commit()


def _run_v2_price_backfill(db):
    """Populate preu_cost from existing prices and copy marge to new columns. Idempotent."""
    check = query("SELECT valor FROM config WHERE clau='migration_v2_done'", one=True)
    if check:
        return
    print("Running v2 price model backfill...")
    # Read current default margin to reverse-calculate cost
    cfg = query("SELECT valor FROM config WHERE clau='marge_defecte'", one=True)
    divisor = 1 + float(cfg['valor'] if cfg else 60) / 100  # e.g. 1.60

    # PG needs ::numeric cast for ROUND with precision; SQLite handles ROUND(real, int) natively
    if USE_PG:
        rnd = lambda col: f'ROUND(({col} / %s)::numeric, 4)'
    else:
        rnd = lambda col: f'ROUND({col} / ?, 4)'
    # Moldures: preu_cost = preu_taller / divisor
    execute(f"UPDATE moldures SET preu_cost = {rnd('preu_taller')} WHERE preu_cost IS NULL AND preu_taller IS NOT NULL", [divisor])
    # Vidres
    execute(f"UPDATE vidres SET preu_cost = {rnd('preu')} WHERE preu_cost IS NULL AND preu IS NOT NULL", [divisor])
    # Encolat
    execute(f"UPDATE encolat_pro SET preu_cost = {rnd('preu')} WHERE preu_cost IS NULL AND preu IS NOT NULL", [divisor])
    # Passpartout
    execute(f"UPDATE passpartout SET preu_cost = {rnd('preu')} WHERE preu_cost IS NULL AND preu IS NOT NULL", [divisor])
    # Copy existing margins to new alias columns (always overwrite defaults on first migration)
    execute("UPDATE usuaris SET marge_pro_pct = marge WHERE marge IS NOT NULL")
    execute("UPDATE usuaris SET marge_impressio_pro_pct = marge_impressio WHERE marge_impressio IS NOT NULL")
    # Mark migration as done
    if USE_PG:
        execute("INSERT INTO config (clau, valor) VALUES (%s, %s) ON CONFLICT DO NOTHING", ['migration_v2_done', '1'])
    else:
        execute("INSERT OR IGNORE INTO config (clau, valor) VALUES (?, ?)", ['migration_v2_done', '1'])
    db.commit()
    print("v2 price backfill complete.")


# init_db() ja NO s'executa automàticament a cap request. S'invoca explícitament
# des de /admin/run-migrations perquè els workers arranquin sempre sense tocar BD
# (evita penjades d'arrencada a PG). Cal visitar /admin/run-migrations com a admin
# després de cada deploy amb canvis d'schema.

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
