
from flask import Flask, render_template, request, send_file, abort
from docx import Document
from io import BytesIO
from collections import Counter
from threading import Lock
from time import time
import re, secrets

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024
CACHE = {}
LOCK = Lock()
STOP = set('the and with for are you your from that this will have has our their job work years experience strong ability skills knowledge required preferred using within across into about including must should candidate team teams role responsible good excellent develop development support supporting business project projects solutions systems technology technologies professional'.split())
MODES = {'ats': 'ATS-Friendly', 'recruiter': 'Recruiter-Friendly', 'hands-on': 'Hands-On Experience'}

def words(s):
    return re.findall(r'(?<!\w)[a-z][a-z0-9+#./-]{1,}(?!\w)', s.lower())

def terms(s):
    return {w for w in words(s) if w not in STOP and len(w) > 2}

def extract(doc):
    parts = [p.text for p in doc.paragraphs]
    for t in doc.tables:
        for r in t.rows:
            parts.extend(c.text for c in r.cells)
    return '\n'.join(parts)

def relevance(text, wanted, mode):
    found = terms(text) & wanted
    score = len(found)
    if mode == 'hands-on':
        verbs = ('implemented', 'developed', 'designed', 'configured', 'troubleshot', 'optimized', 'migrated', 'automated', 'integrated', 'tested', 'debugged', 'deployed', 'built')
        score += 1.5 * sum(v in text.lower() for v in verbs)
    elif mode == 'recruiter':
        score += min(len(text.split()) / 30, 1.5)
    return score

def is_bullet(p):
    st = (p.style.name or '').lower()
    return 'list' in st or 'bullet' in st or p.text.lstrip().startswith(('•', '-', '–', '*')) or p._p.pPr is not None and p._p.pPr.numPr is not None

def tailor(original, jd, mode):
    doc = Document(BytesIO(original))
    wanted = terms(jd)
    # Move whole XML paragraph nodes, preserving original runs, bullets, fonts and styling.
    # Only reorder contiguous bullet paragraphs. Never invent or append unverified skills.
    paragraphs = list(doc.paragraphs)
    i = 0
    while i < len(paragraphs):
        if not is_bullet(paragraphs[i]):
            i += 1
            continue
        j = i
        while j < len(paragraphs) and is_bullet(paragraphs[j]):
            j += 1
        block = paragraphs[i:j]
        if len(block) > 1:
            ranked = sorted(enumerate(block), key=lambda x: (-relevance(x[1].text, wanted, mode), x[0]))
            parent = block[0]._p.getparent()
            index = parent.index(block[0]._p)
            for p in block:
                parent.remove(p._p)
            for offset, (_, p) in enumerate(ranked):
                parent.insert(index + offset, p._p)
        i = j
    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()

@app.get('/')
def home():
    return render_template('index.html')

@app.post('/analyze')
def analyze():
    jd = request.form.get('jd', '').strip()
    f = request.files.get('resume')
    if not jd or not f or not (f.filename or '').lower().endswith('.docx'):
        return render_template('index.html', error='Upload a .docx resume and paste a job description.', jd=jd)
    raw = f.read()
    if not raw:
        return render_template('index.html', error='The uploaded file is empty.', jd=jd)
    try:
        doc = Document(BytesIO(raw))
        resume = extract(doc)
    except Exception:
        return render_template('index.html', error='Could not read the Word file.', jd=jd)
    wanted, present = terms(jd), terms(resume)
    matched = sorted(wanted & present)
    missing = sorted(wanted - present)
    score = round(100 * len(matched) / max(len(wanted), 1))
    token = secrets.token_urlsafe(24)
    with LOCK:
        now = time()
        for key in list(CACHE):
            if now - CACHE[key]['created'] > 1800:
                del CACHE[key]
        CACHE[token] = {'created': now, 'raw': raw, 'jd': jd}
    return render_template('index.html', score=score, found=matched[:70], missing=missing[:70],
                           token=token, jd=jd, preview=resume[:3500],
                           message='Estimated keyword overlap only — not an official ATS score. No skills or experience are fabricated.')

@app.get('/download/<token>/<mode>')
def download(token, mode):
    if mode not in MODES:
        abort(404)
    with LOCK:
        entry = CACHE.get(token)
    if not entry or time() - entry['created'] > 1800:
        abort(404, description='Resume expired. Please upload again.')
    data = tailor(entry['raw'], entry['jd'], mode)
    return send_file(BytesIO(data), as_attachment=True, download_name=f'CareerLift_{mode}.docx',
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

if __name__ == '__main__':
    app.run(debug=True)
