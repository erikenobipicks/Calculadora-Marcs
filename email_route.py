# Pega aquest bloc a app.py just abans de "# ── Init DB"

EMAIL_ROUTE = '''
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
        return jsonify({'ok': False, 'error': "Configura el Gmail a l\'admin primer."})
    dest = 'reusrevela@gmail.com'
    html = f"""<h2 style='color:#1A6B45;font-family:sans-serif'>Nou pressupost d\'emmarcació</h2>
<p style='font-family:sans-serif'><b>Comerç:</b> {nom_comerc}</p>
<p style='font-family:sans-serif'><b>Client:</b> {c['client_nom']} · {c['client_tel'] or '—'}</p>
<p style='font-family:sans-serif'><b>Data:</b> {c['data']}</p>
<hr><p style='font-family:sans-serif'><b>Marc:</b> {c['marc_principal']}</p>
<p style='font-family:sans-serif'><b>Mides:</b> {c['amplada']} × {c['alcada']} cm</p>
<p style='font-family:sans-serif'><b>Preu final:</b> <span style='color:#1A6B45;font-size:18px'>{c['preu_final']:.2f} €</span></p>
<p style='font-family:sans-serif'><b>Pendent:</b> {c['pendent']:.2f} €</p>
{'<p style=font-family:sans-serif><b>Nota:</b> '+nota+'</p>' if nota else ''}
"""
    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = f"Pressupost #{cid} — {nom_comerc} — {c['client_nom']}"
        msg['From'] = gmail_user
        msg['To'] = dest
        msg.attach(MIMEText(html,'html'))
        with smtplib.SMTP_SSL('smtp.gmail.com',465) as s:
            s.login(gmail_user,gmail_pass)
            s.sendmail(gmail_user,dest,msg.as_string())
        return jsonify({'ok':True})
    except Exception as e:
        return jsonify({'ok':False,'error':str(e)})
'''
