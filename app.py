import hashlib, secrets, os, json
from flask import (Flask, render_template, request, redirect, url_for,
                   session, flash, jsonify, send_file, g)
from datetime import datetime
from functools import wraps
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
    if request.method == 'POST':
        user = query('SELECT * FROM usuaris WHERE username=?',
                     [request.form['username']], one=True)
        if user and user['password'] == hash_pw(request.form['password']):
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['is_admin'] = bool(user['is_admin'])
            session['nom'] = user['nom']
            # Load empresa_nom for nav
            try:
                nom_emp = user['nom_empresa'] if user['nom_empresa'] else ''
                if not nom_emp:
                    r_cfg = query("SELECT valor FROM config WHERE clau='empresa_nom'", one=True)
                    nom_emp = r_cfg['valor'] if r_cfg and r_cfg['valor'] else ''
                session['empresa_nom'] = nom_emp or 'Calculadora'
            except:
                session['empresa_nom'] = 'Calculadora'
            # First time setup
            try:
                u2 = query('SELECT setup_done FROM usuaris WHERE id=?', [user['id']], one=True)
                if u2 and not u2.get('setup_done'):
                    return redirect(url_for('setup'))
            except Exception as _se:
                print(f'setup check error: {_se}')
            return redirect(url_for('index'))
        flash('Usuari o contrasenya incorrectes.', 'error')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

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
    return render_template('calculadora.html')

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
                                'descripcio': r['descripcio'], 'foto': r['foto'] or ''})
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
@login_required
def api_empresa():
    d = request.json
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
    d = request.json
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
         marge, descompte, quantitat,
         preu_net, preu_final, entrega, pendent, observacions,
         sessio_id, opcio_nom, num_pressupost, lang)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', [
        session['user_id'], datetime.now().strftime('%d/%m/%Y %H:%M'),
        d.get('client_nom',''), d.get('client_tel',''),
        d.get('pre_marc',''), d.get('marc_principal',''),
        d.get('amplada',0), d.get('alcada',0), d.get('copia',0),
        d.get('encolat',''), d.get('vidre',''), d.get('passpartout',''),
        d.get('impressio',''),
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
    pagat = request.json.get('pagat', 1)
    execute('UPDATE comandes SET pagat=? WHERE sessio_id=?', [pagat, sessio_id])
    return jsonify({'ok': True})

@app.route('/sessio/<sessio_id>/entregat', methods=['POST'])
@login_required
def marcar_entregat(sessio_id):
    entregat = request.json.get('entregat', 1)
    execute('UPDATE comandes SET entregat=? WHERE sessio_id=?', [entregat, sessio_id])
    return jsonify({'ok': True})

@app.route('/comanda/<int:cid>/liquidar', methods=['POST'])
@login_required
def liquidar_comanda(cid):
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
    c = query('SELECT user_id FROM comandes WHERE id=?', [cid], one=True)
    if not c:
        return jsonify({'ok': False, 'error': 'No trobada'})
    if not session.get('is_admin') and c['user_id'] != session['user_id']:
        return jsonify({'ok': False, 'error': 'No autoritzat'})
    execute('DELETE FROM comandes WHERE id=?', [cid])
    return jsonify({'ok': True})

@app.route('/comanda/<int:cid>/acceptar', methods=['POST'])
@login_required
def acceptar_comanda(cid):
    c = query('SELECT user_id FROM comandes WHERE id=?', [cid], one=True)
    if not c:
        return jsonify({'ok': False})
    if not session.get('is_admin') and c['user_id'] != session['user_id']:
        return jsonify({'ok': False})
    estat = request.json.get('estat', 'acceptat')
    execute('UPDATE comandes SET observacions = CASE WHEN observacions IS NULL OR observacions=\'\' THEN ? ELSE observacions || \' | \' || ? END WHERE id=?',
            [f'[{estat.upper()}]', f'[{estat.upper()}]', cid])
    return jsonify({'ok': True})

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
                           filtre_uid=filtre_uid)

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
        execute('INSERT INTO usuaris (username, password, nom, is_admin) VALUES (?,?,?,?)',
                [request.form['username'], hash_pw(request.form['password']),
                 request.form['nom'], int(request.form.get('is_admin', 0))])
        flash(f"Usuari '{request.form['nom']}' creat.", 'ok')
    elif action == 'eliminar':
        execute('DELETE FROM usuaris WHERE id=?', [request.form['uid']])
        flash('Usuari eliminat.', 'ok')
    elif action == 'canviar_pw':
        execute('UPDATE usuaris SET password=? WHERE id=?',
                [hash_pw(request.form['password']), request.form['uid']])
        flash('Contrasenya actualitzada.', 'ok')
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
        ext = f.filename.rsplit('.', 1)[-1].lower()
        nom = f'{ref}.{ext}'
        path = os.path.join(app.root_path, 'static', 'fotos', nom)
        f.save(path)
        execute('UPDATE moldures SET foto=? WHERE LOWER(referencia)=?', [f'/static/fotos/{nom}', ref])
        flash(f'Foto de {ref} pujada.', 'ok')
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

    # Header
    c0 = comandes[0]
    header = Table([[
        p(f"COMPARATIVA DE PRESSUPOSTOS", bold=True, size=14, color=colors.white),
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
        p(f"Client: {c0['client_nom'] or '—'}", bold=True, size=11),
        p(f"Data: {c0['data']}", size=10, align='RIGHT')
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
        ('Marc principal',   'marc_principal', False),
        (t['premarc'],         'pre_marc',       False),
        ('Mides foto (cm)',  None,             False),
        (t['muntatge'],         'encolat',        False),
        ('Vidre / Mirall',   'vidre',          False),
        (t['interior'],         'passpartout',    False),
        (t['impressio'],        'impressio',      False),
        (t['observacions'],     'observacions',   False),
        ('Preu net',         'preu_net',       True),
        (t['iva'],          None,             True),
        ('TOTAL amb IVA',    'preu_final',     True),
        ('Entrega compte',   'entrega',        True),
        ('PENDENT',          'pendent',        True),
    ]

    rows = []
    for lbl, key, is_price in fields:
        lbl_color = DARK if not is_price else colors.HexColor("#1A6B45") if key=='preu_final' else                     RED if key=='pendent' else colors.HexColor("#6B6860")
        row = [p(lbl, bold=is_price, size=9, color=lbl_color)]
        for c in comandes:
            if key is None and lbl == 'Mides foto (cm)':
                val = f"{int(c.get('amplada',0))} x {int(c.get('alcada',0))}"
            elif key is None and lbl == t['iva']:
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
    },
    'es': {
        'pressupost': 'PRESUPUESTO',
        'client': 'Cliente',
        'telefon': 'Teléfono',
        'data': 'Fecha',
        'marc': 'Marco',
        'premarc': 'Pre-Marco',
        'mides': 'Medidas',
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
        if r and r['foto']:
            rel = r['foto'].lstrip('/')
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
    if c.get('pre_marc') and c['pre_marc'] not in ['-','']:
        det_rows.append(fila('Pre-Marc:', c['pre_marc']))
    det_rows.append(fila('Marc principal:',
                         c['marc_principal'] if c.get('marc_principal') else '(sense marc)'))
    det_rows.append(fila('Mides foto (cm):', f"{int(c['amplada'])} × {int(c['alcada'])}"))
    if c.get('encolat') and c['encolat'] not in ['-','']:
        det_rows.append(fila('Muntatge:', c['encolat']))
    if c.get('vidre') and c['vidre'] not in ['-','CONSERVAR','']:
        det_rows.append(fila('Vidre / Mirall:', c['vidre']))
    elif c.get('vidre') == 'CONSERVAR':
        det_rows.append(fila('Vidre:', 'Conservar vidre existent'))
    if c.get('passpartout') and c['passpartout'] not in ['-','']:
        det_rows.append(fila('Passpartout / ProEco:', c['passpartout']))
    if c.get('impressio') and c['impressio'] not in ['-','']:
        det_rows.append(fila('Impressió fotogràfica:', c['impressio']))
    if c.get('observacions'):
        det_rows.append(fila('Observacions:', c['observacions']))

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
        [p('Preu net PVP (sense IVA)', bold=True, size=9, color=colors.HexColor("#6B6860")),
         p(f'{pnet:.2f} €', size=10, align='RIGHT')],
    ]
    if desc_pct > 0:
        desc_eur = pnet * (desc_pct/100)
        t3_data.append([
            p(f'Descompte {int(desc_pct)}% sobre PVP', bold=True, size=9, color=colors.HexColor("#C8873A")),
            p(f'- {desc_eur:.2f} €', size=10, color=colors.HexColor("#C8873A"), align='RIGHT'),
        ])
    t3_data += [
        [p(t['iva'], bold=True, size=9, color=colors.HexColor("#6B6860")),
         p(f'{(pnet*(1-desc_pct/100))*0.21:.2f} €', size=10, align='RIGHT')],
        [p(f'TOTAL PVP amb IVA', bold=True, size=11, color=GREEN),
         p(f'{pfin:.2f} €', bold=True, size=14, color=GREEN, align='RIGHT')],
        [p(t['entrega'], bold=True, size=9, color=colors.HexColor("#6B6860")),
         p(f'{pent:.2f} €', size=10, align='RIGHT')],
        [p('PENDENT de cobrar', bold=True, size=11, color=RED),
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
    marge_actual = int(u['marge']) if u and u['marge'] is not None else 60
    marge_imp = int(u['marge_impressio']) if u and u['marge_impressio'] is not None else 0
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
    d = request.json
    cid = d.get('comanda_id')
    nom_comerc = d.get('nom_comerc', session.get('nom',''))
    c = query('SELECT * FROM comandes WHERE id=?', [cid], one=True)
    if not c:
        return jsonify({'ok': False})
    lines = [
        f"Bon dia,",
        f"",
        f"Us faig arribar el pressupost d'emmarcació per al client {c['client_nom']}.",
        f"",
        f"DETALL DE LA COMANDA:",
        f"  Marc: {c['marc_principal']}",
    ]
    if c['pre_marc'] and c['pre_marc'] != '-':
        lines.append(f"  Pre-Marc: {c['pre_marc']}")
    lines += [
        f"  Mides: {int(c['amplada'])} x {int(c['alcada'])} cm",
    ]
    if c['vidre'] and c['vidre'] != '-':
        lines.append(f"  Vidre: {c['vidre']}")
    if c['passpartout'] and c['passpartout'] != '-':
        lines.append(f"  Passpartout/ProEco: {c['passpartout']}")
    if c['encolat'] and c['encolat'] != '-':
        lines.append(f"  Encolat: {c['encolat']}")
    lines += [
        f"",
        f"PREU FINAL: {c['preu_final']:.2f} EUR (IVA inclòs)",
    ]
    if c['pendent'] and float(c['pendent']) > 0.01:
        lines.append(f"Pendent de cobrar: {c['pendent']:.2f} EUR")
    if c['observacions']:
        lines.append(f"Observacions: {c['observacions']}")
    lines += [
        f"",
        f"Trobareu el pressupost en PDF adjunt.",
        f"",
        f"Atentament,",
        f"{nom_comerc}",
    ]
    body = "%0D%0A".join(lines)
    subject = f"Pressupost emmarcació - {c['client_nom']}"
    mailto = f"mailto:reusrevela@gmail.com?subject={subject}&body={body}"
    return jsonify({'ok': True, 'mailto': mailto})

@app.route('/enviar-email', methods=['POST'])
@login_required
def enviar_email():
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    d = request.json
    cid = d.get('comanda_id')
    nom_comerc = d.get('nom_comerc', session.get('nom',''))
    nota = d.get('nota','')
    c = query('SELECT * FROM comandes WHERE id=?', [cid], one=True)
    if not c:
        return jsonify({'ok': False, 'error': 'Comanda no trobada'})
    cfg = {r['clau']: r['valor'] for r in query('SELECT * FROM config')}
    gmail_user = cfg.get('gmail_user','')
    gmail_pass = cfg.get('gmail_pass','')
    if not gmail_user or not gmail_pass:
        return jsonify({'ok': False, 'error': "Configura el Gmail a l'admin primer."})
    dest = 'reusrevela@gmail.com'
    obs = c['observacions'] or ''
    nota_html = f"<p style='font-family:sans-serif;color:#C8873A'><b>Nota:</b> {nota}</p>" if nota else ""
    html = f"""
    <h2 style='color:#1A6B45;font-family:sans-serif;border-bottom:2px solid #1A6B45;padding-bottom:8px'>
        Nou pressupost d'emmarcacio</h2>
    <p style='font-family:sans-serif;font-size:15px'><b>Comers:</b> {nom_comerc}</p>
    <p style='font-family:sans-serif;font-size:15px'><b>Client:</b> {c['client_nom']} &nbsp;·&nbsp; {c['client_tel'] or '-'}</p>
    <p style='font-family:sans-serif;font-size:15px'><b>Data:</b> {c['data']}</p>
    <hr style='border:1px solid #E5E2DB;margin:12px 0'>
    <p style='font-family:sans-serif;font-size:14px'><b>Marc principal:</b> {c['marc_principal']}</p>
    <p style='font-family:sans-serif;font-size:14px'><b>Mides:</b> {c['amplada']} x {c['alcada']} cm</p>
    <p style='font-family:sans-serif;font-size:14px'><b>Vidre:</b> {c['vidre'] or '-'}</p>
    <p style='font-family:sans-serif;font-size:14px'><b>Passpartout:</b> {c['passpartout'] or '-'}</p>
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
                    marge_impressio REAL DEFAULT 100, nom_empresa TEXT DEFAULT ''
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
                ('comandes','sessio_id','TEXT'),
                ('comandes','opcio_nom','TEXT'),
                ('usuaris','nom_empresa',"TEXT DEFAULT ''"),
                ('usuaris','setup_done','INTEGER DEFAULT 0'),
                ('usuaris','logo_b64','TEXT'),
                ('usuaris','marge_impressio_setup','INTEGER DEFAULT 0'),
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
            cur.execute("SELECT id FROM usuaris WHERE username=%s", ['admin'])
            if not cur.fetchone():
                cur.execute("INSERT INTO usuaris (username,password,nom,is_admin) VALUES (%s,%s,%s,%s)",
                           ['admin', hash_pw('admin123'), 'Administrador', 1])
                cur.execute("INSERT INTO config (clau,valor) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                           ['marge_defecte','60'])
                cur.execute("INSERT INTO config (clau,valor) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                           ['empresa_nom','Reus Revela'])
                cur.execute("INSERT INTO config (clau,valor) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                           ['empresa_adreca',''])
                cur.execute("INSERT INTO config (clau,valor) VALUES (%s,%s) ON CONFLICT DO NOTHING",
                           ['empresa_tel',''])
                db.commit()
                print("Admin creat: usuari=admin / contrasenya=admin123")
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
                    nom_empresa TEXT DEFAULT ''
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
            admin = db.execute("SELECT id FROM usuaris WHERE username='admin'").fetchone()
            if not admin:
                db.execute('INSERT INTO usuaris (username,password,nom,is_admin) VALUES (?,?,?,1)',
                           ['admin', hash_pw('admin123'), 'Administrador'])
                db.execute("INSERT OR IGNORE INTO config (clau,valor) VALUES ('marge_defecte','60')")
                db.commit()
                print("Admin creat: usuari=admin / contrasenya=admin123")

# Init DB via before_first_request equivalent
_db_initialized = False

@app.before_request
def ensure_db():
    global _db_initialized
    if not _db_initialized:
        _db_initialized = True
        try:
            init_db()
            # Ensure config rows exist
            for clau, valor in [('empresa_nom','Reus Revela'), ('empresa_adreca',''), ('empresa_tel',''), ('marge_defecte','60')]:
                try:
                    execute('INSERT OR IGNORE INTO config (clau,valor) VALUES (?,?)', [clau, valor])
                except Exception as _ce:
                    pass  # Row may already exist
            print("init_db OK")
        except Exception as e:
            import traceback
            print(f"init_db ERROR: {e}")
            traceback.print_exc()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
