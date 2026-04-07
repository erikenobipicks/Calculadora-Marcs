import base64, hashlib, hmac, secrets, os, json, time, unicodedata
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, jsonify, send_file, g)
from datetime import datetime
from functools import wraps
from urllib.parse import urlencode
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, HRFlowable)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
import io

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', secrets.token_hex(32))

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

    candidates = []
    for stem in [ref.strip(), ref.strip().lower(), ref.strip().upper(), _safe_moldura_ref(ref)]:
        if stem and stem not in candidates:
            candidates.append(stem)

    for stem in candidates:
        for ext in MOLDURA_IMAGE_EXTS:
            path = os.path.join(fotos_dir, f'{stem}.{ext}')
            if os.path.isfile(path):
                return f'/static/fotos/{stem}.{ext}'
    return ''


def _resolve_moldura_photo(ref, foto):
    public_url = _to_public_photo_url(foto)
    return public_url or _find_local_moldura_photo(ref)


def _serialize_moldura(row):
    if not row:
        return None
    data = dict(row)
    data['foto'] = _resolve_moldura_photo(data.get('referencia', ''), data.get('foto', ''))
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

# ── DB layer: PostgreSQL (production) or SQLite (local) ───────────────────
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
            g.db = psycopg2.connect(DATABASE_URL, cursor_factory=psycopg2.extras.RealDictCursor)
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

def hash_pw(pw): return hashlib.sha256(pw.encode()).hexdigest()


def _is_admin_session():
    return bool(session.get('is_admin'))


def _user_access_status(user):
    status = ''
    if user:
        try:
            status = (user.get('access_status') or '').strip().lower()
        except Exception:
            status = ''
    return status or 'active'


def _user_is_allowed(user):
    return bool(user) and (bool(user.get('is_admin')) or _user_access_status(user) == 'active')


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
    source = session.get('bridge_source') or request.args.get('source') or 'web'
    lang = session.get('bridge_lang') or request.args.get('lang') or 'ca'
    return _build_web_return_url(source, lang)


def _current_web_order_url():
    base = _main_site_url()
    lang = (session.get('bridge_lang') or request.args.get('lang') or 'ca').strip().lower() or 'ca'
    return f'{base}/area-privada/comanda?lang={lang}'


def _needs_setup(user_id):
    try:
        u2 = query('SELECT setup_done FROM usuaris WHERE id=?', [user_id], one=True)
        return bool(u2) and not u2.get('setup_done')
    except Exception as exc:
        print(f'setup check error: {exc}')
        return False


def _start_user_session(user):
    session['user_id'] = user['id']
    session['username'] = user['username']
    session['is_admin'] = bool(user['is_admin'])
    session['nom'] = user['nom']
    session['access_status'] = _user_access_status(user)
    session['profile_type'] = user.get('profile_type', 'professional') if user else 'professional'
    session['empresa_nom'] = _load_empresa_nom_for_session(user)


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

# ── Auth decorators ───────────────────────────────────────────────────────
def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
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

# ── Routes: Auth ─────────────────────────────────────────────────────────
@app.route('/login', methods=['GET', 'POST'])
def login():
    web_return_url = _current_web_return_url()
    if request.method == 'POST':
        user = query('SELECT * FROM usuaris WHERE username=?',
                     [request.form['username']], one=True)
        if user and user['password'] == hash_pw(request.form['password']):
            if not _user_is_allowed(user):
                status = _user_access_status(user)
                if status == 'pending':
                    flash("El teu accés encara està pendent de validació.", 'error')
                else:
                    flash("El teu accés està bloquejat. Contacta amb l'administrador.", 'error')
                return render_template('login.html', web_return_url=web_return_url)
            _start_user_session(user)
            session['bridge_source'] = (request.args.get('source') or session.get('bridge_source') or '').strip().lower()
            session['bridge_lang'] = (request.args.get('lang') or session.get('bridge_lang') or 'ca').strip().lower()
            if _needs_setup(user['id']):
                return redirect(url_for('setup'))
            return redirect(url_for('index'))
        flash('Usuari o contrasenya incorrectes.', 'error')
    return render_template('login.html', web_return_url=web_return_url)

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

    user = query('SELECT * FROM usuaris WHERE username=?', [username], one=True)
    if not user or user['password'] != hash_pw(password):
        return jsonify({'ok': False, 'error': 'invalid_credentials'}), 401

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


@app.route('/auth/bridge')
def bridge_auth():
    payload = _read_bridge_token(request.args.get('token'))
    if not payload:
        flash("L'accÃ©s unificat ha caducat o no Ã©s vÃ lid. Torna a iniciar sessiÃ³.", 'error')
        return redirect(url_for('login'))

    user = query('SELECT * FROM usuaris WHERE id=?', [payload.get('uid')], one=True)
    if not _user_is_allowed(user):
        flash("No hem pogut validar aquest accÃ©s. Contacta amb administraciÃ³ si cal.", 'error')
        return redirect(url_for('login'))

    _start_user_session(user)
    session['bridge_source'] = str(payload.get('source') or '').strip().lower()
    session['bridge_lang'] = str(payload.get('lang') or 'ca').strip().lower()
    target = _safe_next_path(payload.get('next'), '/')
    if _needs_setup(user['id']) and target != url_for('logout'):
        return redirect(url_for('setup'))
    return redirect(target)

# ── Routes: App principal ─────────────────────────────────────────────────

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
        if u and not u.get('setup_done'):
            return redirect(url_for('setup'))
    except:
        pass
    user = query('SELECT nom, username, nom_empresa, profile_type, access_status, web_url, instagram, fiscal_id, notes_validacio FROM usuaris WHERE id=?', [session['user_id']], one=True)
    return render_template('portal.html', user=user, web_return_url=_current_web_return_url())


@app.route('/calculadora')
@login_required
def calculadora():
    try:
        u = query('SELECT setup_done FROM usuaris WHERE id=?', [session['user_id']], one=True)
        if u and not u.get('setup_done'):
            return redirect(url_for('setup'))
    except:
        pass
    return render_template('calculadora.html',
                           web_return_url=_current_web_return_url(),
                           web_order_url=_current_web_order_url(),
                           color_filters=MOLDURA_COLOR_FILTERS,
                           gruix_filters=MOLDURA_GRUIX_FILTERS)

@app.route('/api/lookup')
@login_required
def lookup():
    ref = request.args.get('ref', '').strip()
    tipus = request.args.get('tipus', 'moldura')
    if tipus == 'moldura':
        try:
            r = query('SELECT preu_taller, gruix, descripcio, foto FROM moldures WHERE LOWER(referencia)=LOWER(?)', [ref], one=True)
            print(f"lookup moldura ref={ref} result={r}")
            if r:
                return jsonify({'ok': True, 'preu': r['preu_taller'], 'gruix': r['gruix'],
                                'descripcio': r['descripcio'], 'foto': _resolve_moldura_photo(ref, r['foto'])})
        except Exception as e:
            print(f"lookup ERROR: {e}")
            return jsonify({'ok': False, 'error': str(e)})
    elif tipus == 'vidre':
        r = query('SELECT preu FROM vidres WHERE LOWER(referencia)=LOWER(?)', [ref], one=True)
        if r: return jsonify({'ok': True, 'preu': r['preu']})
    elif tipus == 'passpartout':
        r = query('SELECT preu FROM passpartout WHERE LOWER(referencia)=LOWER(?)', [ref], one=True)
        if r: return jsonify({'ok': True, 'preu': r['preu']})
    elif tipus == 'encolat':
        r = query('SELECT preu FROM encolat_pro WHERE LOWER(referencia)=LOWER(?)', [ref], one=True)
        if r: return jsonify({'ok': True, 'preu': r['preu']})
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


@app.route('/api/moldura-options')
@login_required
def moldura_options():
    rows = query("""SELECT referencia, gruix, descripcio
                    FROM moldures
                    ORDER BY referencia""")
    return jsonify([dict(r) for r in rows])

@app.route('/api/marge')
@login_required
def get_marge():
    u = query('SELECT marge, marge_impressio, nom_empresa FROM usuaris WHERE id=?', [session['user_id']], one=True)
    marge = float(u['marge']) if u and u['marge'] is not None else 60
    marge_imp = float(u['marge_impressio']) if u and u['marge_impressio'] is not None else 0
    nom_emp = u['nom_empresa'] if u and u['nom_empresa'] else ''
    cfg_rows = query("SELECT clau, valor FROM config WHERE clau LIKE 'empresa_%'")
    cfg = {r['clau']: r['valor'] for r in (cfg_rows or [])}
    if not nom_emp:
        nom_emp = cfg.get('empresa_nom', 'Objectiu Emmarcació')
    return jsonify({
        'marge': marge,
        'marge_impressio': marge_imp,
        'empresa_nom': nom_emp,
        'empresa_adreca': cfg.get('empresa_adreca',''),
        'empresa_tel': cfg.get('empresa_tel',''),
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
    return jsonify({'url': r['logo_b64'] if r and r.get('logo_b64') else ''})

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
    execute('UPDATE usuaris SET marge=?, marge_impressio=?, nom_empresa=? WHERE id=?', [m, mi, ne, session['user_id']])
    if ne: session['empresa_nom'] = ne
    return jsonify({'ok': True})

# ── Routes: Guardar comanda i historial ──────────────────────────────────
@app.route('/guardar', methods=['POST'])
@login_required
def guardar():
    d = request.json
    sessio_id = d.get('sessio_id') or secrets.token_hex(8)
    num_pressupost = generar_num_pressupost()
    cid = execute('''INSERT INTO comandes
        (user_id, data, client_nom, client_tel,
         pre_marc, marc_principal, amplada, alcada, copia,
         encolat, vidre, passpartout, impressio,
         tipus_peca, final_amplada, final_alcada,
         marge, descompte, quantitat,
         preu_net, preu_final, entrega, pendent, observacions,
         sessio_id, opcio_nom, num_pressupost, lang)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', [
        session['user_id'], datetime.now().strftime('%d/%m/%Y %H:%M'),
        d.get('client_nom',''), d.get('client_tel',''),
        d.get('pre_marc',''), d.get('marc_principal',''),
        d.get('amplada',0), d.get('alcada',0), d.get('copia',0),
        d.get('encolat',''), d.get('vidre',''), d.get('passpartout',''),
        d.get('impressio',''),
        d.get('tipus_peca','fotografia'),
        d.get('final_amplada',0), d.get('final_alcada',0),
        d.get('marge',60), d.get('descompte',0), d.get('quantitat',1),
        d.get('preu_net',0), d.get('preu_final',0),
        d.get('entrega',0), d.get('pendent',0),
        d.get('observacions',''),
        sessio_id, d.get('opcio_nom','Opció A'), num_pressupost, d.get('lang','ca')
    ])
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
    sql = """SELECT referencia, gruix, descripcio, foto
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
            moldura_data['foto'] = foto or _resolve_moldura_photo(ref, moldura.get('foto', ''))
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
    nou = 0 if (m and m.get('actiu', 1)) else 1
    try:
        execute('UPDATE moldures SET actiu=? WHERE referencia=?', [nou, ref])
    except:
        pass
    return jsonify({'ok': True, 'actiu': nou})

# ── API: buscar moldura per ref exacte (autocomplete) ─────────────────────
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

@app.route('/historial')
@login_required
def historial():
    filtre_uid = request.args.get('user_id', type=int)
    if session.get('is_admin'):
        if filtre_uid:
            comandes = query('''SELECT c.*, u.nom as usuari_nom FROM comandes c
                               JOIN usuaris u ON c.user_id=u.id
                               WHERE c.user_id=? ORDER BY c.id DESC''', [filtre_uid])
        else:
            comandes = query('''SELECT c.*, u.nom as usuari_nom FROM comandes c
                               JOIN usuaris u ON c.user_id=u.id
                               ORDER BY c.id DESC''')
        usuaris_list = query('SELECT id, nom, username FROM usuaris WHERE is_admin=0 ORDER BY nom')
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
                           filtre_uid=filtre_uid,
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

# ── Routes: Admin ─────────────────────────────────────────────────────────
@app.route('/admin')
@admin_required
def admin():
    usuaris = query('SELECT * FROM usuaris ORDER BY nom')
    config = {r['clau']: r['valor'] for r in query('SELECT * FROM config')}
    impressio = query('SELECT * FROM impressio ORDER BY preu')
    return render_template('admin.html', usuaris=usuaris, config=config, impressio=impressio)

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
    if request.form.get('save_gmail'):
        gu = request.form.get('gmail_user','').strip()
        gp = request.form.get('gmail_pass','').strip().replace(' ','')
        if gu:
            execute('INSERT OR REPLACE INTO config (clau, valor) VALUES ("gmail_user", ?)', [gu])
        if gp:
            execute('INSERT OR REPLACE INTO config (clau, valor) VALUES ("gmail_pass", ?)', [gp])
    flash('Configuració desada.', 'ok')
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

# ── PDF generator ─────────────────────────────────────────────────────────

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
        st = ParagraphStyle('x', fontName='Helvetica-Bold' if bold else 'Helvetica',
                            fontSize=size, textColor=color,
                            alignment={'LEFT':0,'CENTER':1,'RIGHT':2}[align])
        return Paragraph(str(txt), st)

    story = []

    lang = (comandes[0].get('lang') or 'ca').lower()
    t = PDF_T.get(lang, PDF_T['ca'])

    # Header
    c0 = comandes[0]
    header = Table([[
        p(f"{t['comparativa']}", bold=True, size=14, color=colors.white),
        p(f"Objectiu · Reus", size=9, color=colors.HexColor("#B2BEC3"), align='RIGHT')
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

    # Build comparison rows
    def val_clean(c, key):
        v = c.get(key,'') or ''
        return '—' if v in ['','-','None'] else str(v)

    fields = [
        (t['tipus_peca'],     'piece_type',       False),
        (t['mida_final'],     'final_size',      False),
        (t['mides_foto'],     'photo_size',      False),
        (t['muntatge'],       'muntatge_label',  False),
        (t['vidre_mirall'],   'proteccio_label', False),
        (t['interior'],       'interior_label',  False),
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
            if key == 'piece_type':
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
        ('FONTNAME',(0,0),(0,-1),'Helvetica-Bold'),
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
                fontName='Helvetica-Bold' if bold else 'Helvetica',
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

# ── Translations for PDF ─────────────────────────────────────────────────
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
        st = ParagraphStyle('x', fontName='Helvetica-Bold' if bold else 'Helvetica',
                            fontSize=size, textColor=color,
                            alignment={'LEFT':0,'CENTER':1,'RIGHT':2}[align],
                            leading=size*1.4)
        return Paragraph(str(txt) if txt else '—', st)

    def fila(lbl, val, color_val=None):
        return [p(lbl, bold=True, size=9, color=colors.HexColor("#6B6860")),
                p(str(val) if val and val not in ['-',''] else '—', size=10,
                  color=color_val or DARK)]

    story = []

    # ── Capçalera ────────────────────────────────────────────────────────
    # Get empresa info for this user
    u_data = query('SELECT nom_empresa FROM usuaris WHERE id=?', [c.get('user_id',0)], one=True)
    nom_empresa = ''
    if u_data and u_data.get('nom_empresa'):
        nom_empresa = u_data['nom_empresa']
    if not nom_empresa:
        _r = query("SELECT valor FROM config WHERE clau='empresa_nom'", one=True)
        nom_empresa = (_r['valor'] if _r else '') or 'Reus Revela'
    r_adr  = query("SELECT valor FROM config WHERE clau='empresa_adreca'", one=True)
    adreca = (r_adr['valor'] if r_adr else '') or 'C/ Mare Molas, 26 · Reus'

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

    # ── Logo (si existeix) ────────────────────────────────────────────────
    try:
        u_logo = query('SELECT logo_b64 FROM usuaris WHERE id=?', [c.get('user_id',0)], one=True)
        logo_data_url = (u_logo['logo_b64'] if u_logo and u_logo.get('logo_b64') else '')
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

    # ── Dades client + data ───────────────────────────────────────────────
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

    # ── Foto del marc (si existeix) ────────────────────────────────────────
    foto_path = None
    if c.get('marc_principal'):
        r = query('SELECT foto FROM moldures WHERE LOWER(referencia)=LOWER(?)',
                  [c['marc_principal']], one=True)
        foto_url = _resolve_moldura_photo(c['marc_principal'], r['foto'] if r else '')
        if foto_url.startswith('/static/'):
            rel = foto_url.lstrip('/')
            full = _os.path.join(app.root_path, 'static', rel.replace('static/',''))
            if _os.path.exists(full):
                foto_path = full

    if foto_path:
        try:
            img = RLImage(foto_path, width=40*mm, height=30*mm)
            img.hAlign = 'LEFT'
            story.append(img)
            story.append(Spacer(1, 3*mm))
        except:
            pass

    # ── Detall de la comanda ──────────────────────────────────────────────
    det_rows = []
    final_size = _final_size_text(c, with_unit=True)
    photo_size = _photo_size_text(c, with_unit=True)
    piece_type = _display_piece_type(c, t)
    muntatge_label = _display_muntatge(c, t)
    proteccio_label = _display_proteccio(c, t)
    interior_label = _display_interior(c, t)
    impressio_label = _display_impressio(c, t)

    if final_size:
        det_rows.append(fila(t['mida_final']+':', final_size))
    if photo_size:
        det_rows.append(fila(t['mides_foto']+':', photo_size))
    if muntatge_label:
        det_rows.append(fila(t['muntatge']+':', muntatge_label))
    if proteccio_label:
        det_rows.append(fila(t['vidre_mirall']+':', proteccio_label))
    if interior_label:
        det_rows.append(fila(t['interior']+':', interior_label))
    if impressio_label:
        det_rows.append(fila(t['impressio']+':', impressio_label))
    det_rows.append(fila(t['marc_principal']+':',
                         c['marc_principal'] if c.get('marc_principal') else t['sense_marc']))
    det_rows.append(fila(t['mides_foto']+':', f"{int(c['amplada'])} × {int(c['alcada'])}"))
    if c.get('encolat') and c['encolat'] not in ['-','']:
        det_rows.append(fila(t['muntatge']+':', c['encolat']))
    if c.get('vidre') and c['vidre'] not in ['-','CONSERVAR','']:
        det_rows.append(fila(t['vidre_mirall']+':', c['vidre']))
    elif c.get('vidre') == 'CONSERVAR':
        det_rows.append(fila(t['proteccio']+':', t['conservar_vidre']))
    if c.get('passpartout') and c['passpartout'] not in ['-','']:
        det_rows.append(fila(t['interior']+':', c['passpartout']))
    if c.get('impressio') and c['impressio'] not in ['-','']:
        det_rows.append(fila(t['impressio']+':', c['impressio']))
    if c.get('observacions'):
        det_rows.append(fila(t['observacions']+':', c['observacions']))

    det_rows = []
    if piece_type:
        det_rows.append(fila(t['tipus_peca']+':', piece_type))
    if final_size:
        det_rows.append(fila(t['mida_final']+':', final_size))
    if photo_size:
        det_rows.append(fila(t['mides_foto']+':', photo_size))
    if muntatge_label:
        det_rows.append(fila(t['muntatge']+':', muntatge_label))
    if proteccio_label:
        det_rows.append(fila(t['vidre_mirall']+':', proteccio_label))
    if interior_label:
        det_rows.append(fila(t['interior']+':', interior_label))
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
    u = query('SELECT marge, marge_impressio, nom_empresa FROM usuaris WHERE id=?', [session['user_id']], one=True)
    marge_actual = float(u['marge']) if u and u['marge'] is not None else 60
    marge_imp = float(u['marge_impressio']) if u and u['marge_impressio'] is not None else 0
    if float(marge_actual).is_integer():
        marge_actual = int(marge_actual)
    if float(marge_imp).is_integer():
        marge_imp = int(marge_imp)
    nom_emp = u['nom_empresa'] if u and u['nom_empresa'] else ''
    cfg_rows = query("SELECT clau, valor FROM config WHERE clau LIKE 'empresa_%'")
    cfg = {r['clau']: r['valor'] for r in (cfg_rows or [])}
    if not nom_emp:
        nom_emp = cfg.get('empresa_nom', '')
    return render_template('ajustos.html', marge_actual=marge_actual, marge_imp=marge_imp,
                           nom_empresa=nom_emp,
                           empresa_adreca=cfg.get('empresa_adreca',''),
                           empresa_tel=cfg.get('empresa_tel',''))


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
    }
    return labels.get(value, value.replace('_', ' ').title())

def _display_muntatge(c, t):
    ref = (c.get('encolat') or '').strip().upper()
    if not ref or ref == '-':
        return ''
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

def _display_impressio(c, t):
    ref = (c.get('impressio') or '').strip()
    if not ref or ref == '-':
        return ''
    return t['inclosa']

def _find_closest(rows, w, h, prefix=None):
    best, best_score = None, float('inf')
    for row in rows:
        ref = row['referencia']
        if prefix and not ref.upper().startswith(prefix.upper()):
            continue
        rw, rh = _parse_dims(ref)
        if rw is None: continue
        for fw, fh in [(rw,rh),(rh,rw)]:
            if fw >= w and fh >= h:
                score = (fw-w)+(fh-h)
                if score < best_score:
                    best_score = score
                    best = dict(row)
    return best

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
    """Smallest format that physically contains the dimensions (can rotate).
    Used for impressio: the paper must be >= the photo in both dimensions."""
    best, best_area = None, float('inf')
    for row in rows:
        ref = row['referencia']
        if prefix and not ref.upper().startswith(prefix.upper()): continue
        rw, rh = _parse_dims(ref)
        if rw is None: continue
        # Try both orientations
        fits = (rw >= w and rh >= h) or (rh >= w and rw >= h)
        if fits:
            area = rw * rh
            if area < best_area:
                best_area = area
                best = row
    # Fallback: if nothing contains it, take the largest available
    if best is None:
        best = max(rows, key=lambda r: (_parse_dims(r['referencia'])[0] or 0) * (_parse_dims(r['referencia'])[1] or 0), default=None)
    return best

def _imp_closest(fw, fh):
    rows = [dict(r) for r in query('SELECT * FROM impressio')]
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
    if w <= 0 or h <= 0:
        return jsonify({})

    def closest(table, prefix=None, preu_col='preu', exclude_prefix=None, exclude_multi=None, w_override=None, h_override=None):
        rows = [dict(r) for r in query(f'SELECT * FROM {table}')]
        if exclude_multi:
            rows = [r for r in rows if not any(r['referencia'].upper().startswith(e.upper()) for e in exclude_multi)]
        elif exclude_prefix:
            rows = [r for r in rows if not r['referencia'].upper().startswith(exclude_prefix.upper())]
        uw = w_override if w_override is not None else w
        uh = h_override if h_override is not None else h
        r = _find_closest(rows, uw, uh, prefix)
        if r:
            return {'ref': r['referencia'], 'preu': r.get(preu_col, 0)}
        return None

    def ca(table, fw, fh, prefix=None, exclude_multi=None, preu_col='preu'):
        """Closest by surface area for all products"""
        rows = [dict(r) for r in query(f'SELECT * FROM {table}')]
        r = _find_closest_area(rows, fw, fh, prefix=prefix, exclude_multi=exclude_multi)
        if r:
            return {'ref': r['referencia'], 'preu': r.get(preu_col, 0)}
        return None

    # Encolat/Protter: min-contain (producte ha de cobrir físicament el marc)
    # Vidre/Passpartú/ProEco: àrea propera (permet mides no estàndard)
    # Impressió: àrea propera (format comercial més similar)
    def cc(table, cw, ch, prefix=None, exclude_multi=None, preu_col='preu'):
        """Closest by min overshoot (must contain the size)"""
        rows = [dict(r) for r in query(f'SELECT * FROM {table}')]
        if exclude_multi:
            rows = [r for r in rows if not any(r['referencia'].upper().startswith(e.upper()) for e in exclude_multi)]
        r = _find_closest(rows, cw, ch, prefix)
        if r:
            return {'ref': r['referencia'], 'preu': r.get(preu_col, 0)}
        return None

    result = {
        'encolat':      cc('encolat_pro', w, h, prefix='ENC'),
        'protter':      cc('encolat_pro', w, h, prefix='PRO'),
        'vidre':        ca('vidres',      w, h, exclude_multi=['DV-','MIR-']),
        'doble_vidre':  ca('vidres',      w, h, prefix='DV-'),
        'mirall':       ca('vidres',      w, h, prefix='MIR-'),
        'passpartu':    ca('passpartout', w, h, prefix='1PAS'),
        'doble_pas':    ca('passpartout', w, h, prefix='DOBPAS'),
        'proeco':       ca('passpartout', w, h, prefix='PROECO'),
        'impressio':    _imp_closest(foto_w, foto_h),
    }
    return jsonify(result)

# ── Routes: Email (mailto) ───────────────────────────────────────────────
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
            'impressio': 'Impressió',
            'inclosa': 'Inclosa',
            'encolat_label': 'Encolat',
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
            'impressio': 'Impresión',
            'inclosa': 'Incluida',
            'encolat_label': 'Encolado',
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
            'impressio': 'Print',
            'inclosa': 'Included',
            'encolat_label': 'Mounting',
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
    proteccio_label = _display_proteccio(c, tt)
    interior_label = _display_interior(c, tt)
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
    proteccio_label = _display_proteccio(c, pdf_lang) or '-'
    interior_label = _display_interior(c, pdf_lang) or '-'
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

# ── Init DB ───────────────────────────────────────────────────────────────
def init_db():
    with app.app_context():
        db = get_db()
        if USE_PG:
            # Use a SEPARATE connection with autocommit for DDL
            import psycopg2 as _pg2
            ddl_conn = _pg2.connect(DATABASE_URL)
            ddl_conn.autocommit = True
            ddl_cur = ddl_conn.cursor()
            ddl = [
                """CREATE TABLE IF NOT EXISTS usuaris (
                    id SERIAL PRIMARY KEY, username TEXT UNIQUE NOT NULL,
                    password TEXT NOT NULL, nom TEXT NOT NULL,
                    is_admin INTEGER DEFAULT 0, marge REAL DEFAULT 60,
                    marge_impressio REAL DEFAULT 100, nom_empresa TEXT DEFAULT '',
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
                """CREATE TABLE IF NOT EXISTS comandes (
                    id SERIAL PRIMARY KEY, user_id INTEGER, data TEXT,
                    client_nom TEXT, client_tel TEXT, pre_marc TEXT,
                    marc_principal TEXT, amplada REAL, alcada REAL, copia REAL,
                    encolat TEXT, vidre TEXT, passpartout TEXT, impressio TEXT,
                    tipus_peca TEXT DEFAULT '',
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
                ('comandes','tipus_peca',"TEXT DEFAULT ''"),
                ('comandes','final_amplada','REAL DEFAULT 0'),
                ('comandes','final_alcada','REAL DEFAULT 0'),
                ('comandes','sessio_id','TEXT'),
                ('comandes','opcio_nom','TEXT'),
                ('usuaris','nom_empresa',"TEXT DEFAULT ''"),
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
            ]:
                try:
                    ddl_cur.execute(f"ALTER TABLE {tbl} ADD COLUMN IF NOT EXISTS {col} {typ}")
                except Exception as e:
                    print("ALTER skip:", e)
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
            db.commit()
            _seed_admin_if_configured(db)
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
                CREATE TABLE IF NOT EXISTS comandes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER, data TEXT,
                    client_nom TEXT, client_tel TEXT,
                    pre_marc TEXT, marc_principal TEXT,
                    amplada REAL, alcada REAL, copia REAL,
                    encolat TEXT, vidre TEXT, passpartout TEXT, impressio TEXT,
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
                "ALTER TABLE usuaris ADD COLUMN marge_impressio_setup INTEGER DEFAULT 0",
                "ALTER TABLE usuaris ADD COLUMN access_status TEXT DEFAULT 'active'",
                "ALTER TABLE usuaris ADD COLUMN profile_type TEXT DEFAULT 'professional'",
                "ALTER TABLE usuaris ADD COLUMN web_url TEXT DEFAULT ''",
                "ALTER TABLE usuaris ADD COLUMN instagram TEXT DEFAULT ''",
                "ALTER TABLE usuaris ADD COLUMN fiscal_id TEXT DEFAULT ''",
                "ALTER TABLE usuaris ADD COLUMN notes_validacio TEXT DEFAULT ''",
                "ALTER TABLE comandes ADD COLUMN num_pressupost TEXT",
                "ALTER TABLE comandes ADD COLUMN pagat INTEGER DEFAULT 0",
                "ALTER TABLE comandes ADD COLUMN entregat INTEGER DEFAULT 0",
                "ALTER TABLE comandes ADD COLUMN lang TEXT DEFAULT 'ca'",
            ]:
                try:
                    db.execute(sql)
                except Exception:
                    pass
            db.commit()
            db.execute("INSERT OR IGNORE INTO config (clau,valor) VALUES ('marge_defecte','60')")
            db.execute("INSERT OR IGNORE INTO config (clau,valor) VALUES ('empresa_nom','Reus Revela')")
            db.execute("INSERT OR IGNORE INTO config (clau,valor) VALUES ('empresa_adreca','')")
            db.execute("INSERT OR IGNORE INTO config (clau,valor) VALUES ('empresa_tel','')")
            db.commit()
            _seed_admin_if_configured(db)

# Init DB via before_first_request equivalent
_db_initialized = False

@app.before_request
def ensure_db():
    global _db_initialized
    if not _db_initialized:
        try:
            init_db()
            # Ensure config rows exist
            for clau, valor in [('empresa_nom','Reus Revela'), ('empresa_adreca',''), ('empresa_tel',''), ('marge_defecte','60')]:
                try:
                    execute('INSERT OR IGNORE INTO config (clau,valor) VALUES (?,?)', [clau, valor])
                except Exception as _ce:
                    pass  # Row may already exist
            _db_initialized = True
            print("init_db OK")
        except Exception as e:
            import traceback
            print(f"init_db ERROR: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
