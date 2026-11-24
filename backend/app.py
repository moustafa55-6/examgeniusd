import os, re, json, time, uuid, random, sqlite3
from functools import wraps
from flask import Flask, request, jsonify, g
from werkzeug.exceptions import HTTPException
from werkzeug.security import generate_password_hash, check_password_hash
from itsdangerous import URLSafeTimedSerializer
from dotenv import load_dotenv
import pypdf, docx
from pptx import Presentation
from openai import OpenAI

load_dotenv()
B = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, static_folder=os.path.join(B, '..', 'frontend'), static_url_path='')
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024
ser = URLSafeTimedSerializer(os.getenv('SECRET_KEY', 'dev-change-me'))
MODEL = os.getenv('OPENAI_MODEL', 'gemini-2.5-flash')
DATA = os.getenv('DATA_DIR', B); os.makedirs(DATA, exist_ok=True)
UP = os.path.join(DATA, 'uploads'); os.makedirs(UP, exist_ok=True)
DB = os.path.join(DATA, 'examgenius.db')
MAXC = 60000
SCHEMA = """
create table if not exists users(id integer primary key, name text, email text unique, pw text, created text default current_timestamp);
create table if not exists lectures(id integer primary key, user_id int, title text, filename text, path text, text text, summary text, topics text, created text default current_timestamp);
create table if not exists exams(id integer primary key, user_id int, lecture_id int, config text, questions text, created text default current_timestamp);
create table if not exists attempts(id integer primary key, user_id int, exam_id int, lecture_id int, mode text, score real, correct int, total int, seconds int, detail text, created text default current_timestamp);
"""
SCHEMA += """
create table if not exists battles(id integer primary key, code text unique, host_id int, title text, questions text, per_q int, status text default 'lobby', started_at real, created text default current_timestamp);
create table if not exists battle_players(battle_id int, user_id int, name text, primary key(battle_id,user_id));
create table if not exists battle_answers(battle_id int, user_id int, qidx int, correct int, points int, primary key(battle_id,user_id,qidx));
"""
with sqlite3.connect(DB) as c:
    c.executescript(SCHEMA)
    for t, col in (('attempts', 'label text'), ('lectures', 'prep text')):  # migrations for databases created by v1
        try: c.execute(f'alter table {t} add column {col}')
        except sqlite3.OperationalError: pass

def db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB); g.db.row_factory = sqlite3.Row
    return g.db
@app.teardown_appcontext
def _close(_):
    d = g.pop('db', None)
    if d: d.close()
def rows(sql, a=()): return [dict(r) for r in db().execute(sql, a).fetchall()]
def one(sql, a=()):
    r = db().execute(sql, a).fetchone(); return dict(r) if r else None
def run(sql, a=()):
    c = db().execute(sql, a); db().commit(); return c.lastrowid
def err(m, c=400): return jsonify(error=m), c

@app.errorhandler(Exception)
def _err(e):
    if isinstance(e, HTTPException): return err(e.description, e.code)
    return err(str(e), 500)

def auth(f):
    @wraps(f)
    def w(*a, **k):
        try: g.uid = ser.loads(request.headers.get('Authorization', '').replace('Bearer ', ''), max_age=7 * 86400)
        except Exception: return err('Please log in again', 401)
        return f(*a, **k)
    return w

def ask(system, user):
    if not os.getenv('OPENAI_API_KEY'): raise RuntimeError('OPENAI_API_KEY is not set on the server')
 r = OpenAI(base_url="https://generativelanguage.googleapis.com/v1beta/openai/").chat.completions.create(model=MODEL, response_format={'type': 'json_object'}, temperature=0.3,        messages=[{'role': 'system', 'content': system}, {'role': 'user', 'content': user}])
    return json.loads(r.choices[0].message.content)

# ---------- auth ----------
def token(u): return ser.dumps(u['id'])
def pub(u): return {'id': u['id'], 'name': u['name'], 'email': u['email'], 'created': u['created']}
@app.post('/api/register')
def register():
    d = request.get_json() or {}; n, e, p = (d.get(k, '').strip() for k in ('name', 'email', 'password'))
    if not n or '@' not in e or len(p) < 8: return err('Enter a name, a valid email and a password of at least 8 characters')
    if one('select 1 from users where email=?', (e.lower(),)): return err('That email is already registered')
    i = run('insert into users(name,email,pw) values(?,?,?)', (n, e.lower(), generate_password_hash(p)))
    u = one('select * from users where id=?', (i,)); return jsonify(token=token(u), user=pub(u))
@app.post('/api/login')
def login():
    d = request.get_json() or {}; u = one('select * from users where email=?', (d.get('email', '').lower().strip(),))
    if not u or not check_password_hash(u['pw'], d.get('password', '')): return err('Wrong email or password', 401)
    return jsonify(token=token(u), user=pub(u))
@app.get('/api/me')
@auth
def me(): return jsonify(user=pub(one('select * from users where id=?', (g.uid,))))

# ---------- lectures ----------
def extract(path, ext):
    if ext == 'pdf': return '\n'.join(f'[Page {i+1}]\n{p.extract_text() or ""}' for i, p in enumerate(pypdf.PdfReader(path).pages))
    if ext == 'docx':
        d = docx.Document(path)
        return '\n'.join([p.text for p in d.paragraphs] + [c.text for t in d.tables for r in t.rows for c in r.cells])
    return '\n'.join(f'[Slide {i+1}]\n' + '\n'.join(s.text_frame.text for s in sl.shapes if s.has_text_frame) for i, sl in enumerate(Presentation(path).slides))

@app.post('/api/lectures')
@auth
def upload():
    f = request.files.get('file')
    if not f or '.' not in f.filename: return err('Choose a file to upload')
    ext = f.filename.rsplit('.', 1)[-1].lower()
    if ext not in ('pdf', 'docx', 'pptx'): return err('Only PDF, DOCX and PPTX files are supported')
    d = os.path.join(UP, str(g.uid)); os.makedirs(d, exist_ok=True)
    path = os.path.join(d, uuid.uuid4().hex + '.' + ext); f.save(path)
    text = extract(path, ext)
    if len(text.strip()) < 100: os.remove(path); return err('No readable text found. Scanned or image-only files are not supported.')
    a = ask('You analyze university lectures. Use only the text provided. Return JSON.',
            'Return {"title":str,"summary":str (max 120 words),"topics":[{"name":str,"key_points":[str]}]} with 5-12 topics covering the important definitions, concepts and key information.\n\nLECTURE:\n' + text[:MAXC])
    i = run('insert into lectures(user_id,title,filename,path,text,summary,topics) values(?,?,?,?,?,?,?)',
            (g.uid, a.get('title') or f.filename, f.filename, path, text, a.get('summary', ''), json.dumps(a.get('topics', []))))
    return jsonify(id=i)
@app.get('/api/lectures')
@auth
def lectures(): return jsonify(lectures=rows('select id,title,filename,created from lectures where user_id=? order by id desc', (g.uid,)))
@app.get('/api/lectures/<int:i>')
@auth
def lecture(i):
    l = one('select id,title,filename,summary,topics,created from lectures where id=? and user_id=?', (i, g.uid))
    if not l: return err('Lecture not found', 404)
    l['topics'] = json.loads(l['topics']); return jsonify(lecture=l)
@app.delete('/api/lectures/<int:i>')
@auth
def del_lecture(i):
    l = one('select path from lectures where id=? and user_id=?', (i, g.uid))
    if l:
        if os.path.exists(l['path']): os.remove(l['path'])
        for t in ('attempts', 'exams'): run(f'delete from {t} where lecture_id=? and user_id=?', (i, g.uid))
        run('delete from lectures where id=?', (i,))
    return jsonify(ok=True)

# ---------- exams ----------
TYPES = ('mcq', 'tf', 'fill', 'match', 'short', 'scenario')
GRADER = "You are a fair university professor. Grade each student answer from 0 to 1 using only the rubric and model answer; partial credit is allowed."
def public(qs):
    out = []
    for x in qs:
        p = {k: x.get(k) for k in ('type', 'topic', 'q', 'options', 'lecture_id')}
        if x['type'] == 'match': p['left'] = [a for a, _ in x['pairs']]; p['right'] = x['right']
        out.append(p)
    return out

PROF = {'Easy': 'Easy Professor: basic questions testing fundamental understanding of core definitions and ideas.',
        'Normal': 'Normal Professor: standard university-level questions.',
        'Strict': 'Strict Professor: challenging questions that need deeper understanding and precise answers.',
        'Expert': 'Expert Professor: advanced questions requiring analysis, critical thinking and problem-solving.'}
MIX = '35% mcq, 15% tf, 15% fill, 10% match, 15% short, 10% scenario (scenario = a short case to analyze)'
QSCHEMA = ('Return {"questions":[{"type":"mcq|tf|fill|match|short|scenario","topic":str,"q":str,"options":[4 strings, mcq only],"pairs":[[left,right]...3-5 pairs, match only],'
           '"answer":(mcq: index 0-3; tf: true/false; fill: the missing word or phrase; short/scenario: model answer),"explanation":str,"reference":"page/slide marker or topic from the lecture","rubric":[3-4 criteria, short/scenario only]}]}')

def lectures_for(d):
    ids = d.get('lecture_ids') or [d.get('lecture_id')]
    ls = [one('select * from lectures where id=? and user_id=?', (i, g.uid)) for i in dict.fromkeys(ids)]
    return [l for l in ls if l]

def gen_questions(ls, n, diff, prof, lang, focus=None, mix=MIX):
    share = MAXC // len(ls)
    text = '\n\n'.join(f'=== LECTURE {i+1}: {l["title"]} ===\n' + l['text'][:share] for i, l in enumerate(ls))
    multi = f'The text holds {len(ls)} lectures. Spread the questions evenly across all lectures and their topics, and add "lecture":<lecture number> to every question. ' if len(ls) > 1 else ''
    foc = f"Focus only on these topics: {', '.join(focus)}. " if focus else ''
    r = ask('You are a university professor writing an exam. Use ONLY the lecture text provided. Never invent facts that are not in it; if the lecture lacks information, leave it out. Test understanding and application, not memorization. Return JSON.',
            f'Write exactly {n} questions, difficulty "{diff}", style: {PROF[prof]} Language: {lang}. {multi}{foc}Mix the types roughly: {mix}.\n{QSCHEMA}\n\nLECTURE:\n{text}')
    qs = [x for x in r.get('questions', []) if x.get('type') in TYPES and x.get('q') and ('pairs' in x if x['type'] == 'match' else 'answer' in x)][:n]
    for x in qs:
        k = str(x.get('lecture', 1)); k = int(k) - 1 if k.isdigit() and 0 < int(k) <= len(ls) else 0
        x['lecture_id'], x['lecture'] = ls[k]['id'], ls[k]['title']
        if x['type'] == 'match': x['right'] = [b for _, b in x['pairs']]; random.shuffle(x['right'])
    return qs

def build_exam(d):
    ls = lectures_for(d)
    if not ls: return err('Lecture not found', 404)
    n = int(d.get('count', 10)); prof = d.get('professor', 'Normal'); prof = prof if prof in PROF else 'Normal'
    diff = d.get('difficulty', 'Medium'); lang = d.get('language', 'English')
    qs = gen_questions(ls, n, diff, prof, lang, d.get('topics'))
    if not qs: return err('Could not generate questions from this lecture. Try again.')
    final = len(ls) > 1
    title = 'Final Exam: ' + ' + '.join(l['title'][:30] for l in ls) if final else ls[0]['title']
    cfg = {'count': n, 'difficulty': diff, 'professor': prof, 'duration': int(d.get('duration', 30)), 'language': lang,
           'mode': d.get('mode', 'exam'), 'title': title, 'final': final, 'lecture_ids': [l['id'] for l in ls]}
    i = run('insert into exams(user_id,lecture_id,config,questions) values(?,?,?,?)', (g.uid, ls[0]['id'], json.dumps(cfg), json.dumps(qs)))
    return jsonify(id=i, lecture_id=ls[0]['id'], questions=public(qs), duration=cfg['duration'], mode=cfg['mode'], lang=lang, title=title)

@app.post('/api/exams')
@auth
def make_exam(): return build_exam(request.get_json(silent=True) or {})

def norm(s): return re.sub(r'\W+', ' ', str(s).lower()).strip()
def auto(x, a):
    t = x['type']
    try:
        if t == 'mcq': return 1.0 if a is not None and a != '' and int(a) == int(x['answer']) else 0.0
        if t == 'tf': return 1.0 if str(a).lower() == str(x['answer']).lower() else 0.0
        if t == 'fill': return 1.0 if a and norm(a) == norm(x['answer']) else 0.0
        if t == 'match':
            s = a or []; return sum(1 for i, (_, r) in enumerate(x['pairs']) if i < len(s) and s[i] == r) / len(x['pairs'])
    except (ValueError, TypeError): pass
    return 0.0
def fmt(x, a):
    if a in (None, '', []): return '(no answer)'
    try:
        if x['type'] == 'mcq': return x['options'][int(a)]
        if x['type'] == 'match': return '; '.join(f'{l} → {a[i] if i < len(a) and a[i] else "?"}' for i, (l, _) in enumerate(x['pairs']))
    except (ValueError, IndexError, TypeError): return str(a)
    return str(a)
def get_exam(eid):
    e = one('select * from exams where id=? and user_id=?', (eid, g.uid))
    if e: e['questions'] = json.loads(e['questions']); e['config'] = json.loads(e['config'])
    return e

@app.post('/api/exams/<int:eid>/check')
@auth
def check(eid):
    e = get_exam(eid); d = request.get_json() or {}
    if not e: return err('Exam not found', 404)
    x = e['questions'][int(d['index'])]; a = d.get('answer'); comment = ''
    if x['type'] in ('short', 'scenario'):
        if str(a or '').strip():
            r = ask(GRADER, json.dumps({'question': x['q'], 'rubric': x.get('rubric'), 'model_answer': x['answer'], 'student': a}) + '\nReturn {"score":0-1,"comment":str}')
            s = max(0, min(1, float(r.get('score', 0)))); comment = r.get('comment', '')
        else: s = 0.0
    else: s = auto(x, a)
    ca = [r for _, r in x['pairs']] if x['type'] == 'match' else x['answer']
    return jsonify(score=s, comment=comment, correct=fmt(x, ca if x['type'] != 'mcq' else x['answer']), explanation=x.get('explanation'), reference=x.get('reference'))

@app.post('/api/explain')
@auth
def explain():
    d = request.get_json() or {}; l = one('select text from lectures where id=? and user_id=?', (d.get('lecture_id'), g.uid))
    if not l: return err('Lecture not found', 404)
    r = ask('You are a patient tutor. Explain using only the lecture text provided; if the lecture does not cover it, say so. Return {"explanation": str}.',
            f'Explain this concept simply: {str(d.get("text", ""))[:1000]}\n\nLECTURE:\n' + l['text'][:MAXC])
    return jsonify(explanation=r.get('explanation', ''))

@app.post('/api/exams/<int:eid>/submit')
@auth
def submit(eid):
    e = get_exam(eid)
    if not e: return err('Exam not found', 404)
    d = request.get_json() or {}; ans = d.get('answers') or {}; qs = e['questions']; n = len(qs)
    sc = [None] * n; cm = {}; pend = []; tp = {}
    for i, x in enumerate(qs):
        a = ans.get(str(i))
        if x['type'] in ('short', 'scenario'): pend.append({'i': i, 'question': x['q'], 'rubric': x.get('rubric'), 'model_answer': x['answer'], 'student': a or ''})
        else: sc[i] = auto(x, a)
    for i, s in enumerate(sc):
        if s is not None: tp.setdefault(qs[i]['topic'], []).append(s)
    r = ask(GRADER + ' Then give personalized feedback like a professor. Empty student answers score 0. Return JSON.',
        json.dumps({'short_answers': pend, 'topic_scores_so_far_pct': {t: round(sum(v) / len(v) * 100) for t, v in tp.items()}}) +
        '\nReturn {"grades":[{"i":int,"score":0-1,"comment":str}],"strong":[topic],"weak":[topic],"revise":[topic],"tips":[3-5 personalized study recommendations]}')
    for gr in r.get('grades', []):
        try: sc[int(gr['i'])] = max(0, min(1, float(gr['score']))); cm[int(gr['i'])] = gr.get('comment', '')
        except (KeyError, ValueError, IndexError): pass
    sc = [s or 0.0 for s in sc]
    review, topics = [], {}
    for i, x in enumerate(qs):
        topics.setdefault(x['topic'], []).append(sc[i])
        ca = [b for _, b in x['pairs']] if x['type'] == 'match' else x['answer']
        review.append({'type': x['type'], 'topic': x['topic'], 'lecture_id': x.get('lecture_id'), 'lecture': x.get('lecture'), 'q': x['q'], 'student': fmt(x, ans.get(str(i))), 'correct': fmt(x, ca),
                       'explanation': x.get('explanation'), 'reference': x.get('reference'), 'score': sc[i], 'comment': cm.get(i, '')})
    pct = round(sum(sc) / n * 100, 1); ok = sum(1 for s in sc if s >= 0.5)
    lec = {}
    for r_ in review: lec.setdefault((r_['lecture_id'], r_['lecture'] or e['config'].get('title')), []).append(r_['score'])
    detail = {'lectures': [{'lecture': k[1], 'pct': round(sum(v) / len(v) * 100)} for k, v in lec.items()], 'review': review, 'topics': [{'topic': t, 'pct': round(sum(v) / len(v) * 100)} for t, v in topics.items()],
              'overall': {k: r.get(k, []) for k in ('strong', 'weak', 'revise', 'tips')}}
    aid = run('insert into attempts(user_id,exam_id,lecture_id,mode,score,correct,total,seconds,detail,label) values(?,?,?,?,?,?,?,?,?,?)',
              (g.uid, eid, e['lecture_id'], e['config']['mode'], pct, ok, n, int(d.get('seconds', 0)), json.dumps(detail), e['config'].get('title')))
    return jsonify(id=aid)

@app.get('/api/attempts/<int:i>')
@auth
def attempt(i):
    a = one('select a.*,coalesce(a.label,l.title) as title from attempts a join lectures l on l.id=a.lecture_id where a.id=? and a.user_id=?', (i, g.uid))
    if not a: return err('Result not found', 404)
    a.update(json.loads(a.pop('detail'))); return jsonify(attempt=a)

@app.get('/api/stats')
@auth
def stats():
    at = rows('select a.id,a.score,a.seconds,a.created,a.lecture_id,a.detail,coalesce(a.label,l.title) as title from attempts a join lectures l on l.id=a.lecture_id where a.user_id=? order by a.id', (g.uid,))
    sc = [a['score'] for a in at]; bl, bt = {}, {}
    for a in at:
        bl.setdefault(a['title'], []).append(a['score'])
        for r in json.loads(a['detail'])['review']:
            t = bt.setdefault(r['topic'], [0, 0, a['lecture_id']]); t[0] += r['score']; t[1] += 1; t[2] = r.get('lecture_id') or a['lecture_id']
    topics = sorted([{'topic': k, 'pct': round(v[0] / v[1] * 100), 'lecture_id': v[2]} for k, v in bt.items()], key=lambda x: x['pct'])
    imp = None
    if len(sc) >= 2:
        k = min(3, len(sc) // 2); f, l = sum(sc[:k]) / k, sum(sc[-k:]) / k; imp = round((l - f) / f * 100, 1) if f else None
    return jsonify(total=len(at), avg=round(sum(sc) / len(sc), 1) if sc else 0, hours=round(sum(a['seconds'] for a in at) / 3600, 1),
        recent=[{k: a[k] for k in ('id', 'title', 'score', 'created')} for a in at[-5:][::-1]],
        series=[{'date': a['created'][:10], 'score': a['score']} for a in at],
        by_lecture=[{'title': k, 'avg': round(sum(v) / len(v), 1)} for k, v in bl.items()], by_topic=topics,
        improvement=imp, weakest=topics[0] if topics else None, strongest=topics[-1] if topics else None)

# ---------- exam preparation assistant ----------
@app.get('/api/lectures/<int:i>/prep')
@auth
def prep(i):
    l = one('select id,text,prep from lectures where id=? and user_id=?', (i, g.uid))
    if not l: return err('Lecture not found', 404)
    if not l['prep'] or request.args.get('refresh'):
        r = ask('You are an exam preparation coach. Use ONLY the lecture text. Judge importance from how strongly the lecture emphasizes a concept (repetition, space given, definitions, links to other concepts). Never claim to know or predict actual exam questions. Return JSON.',
                'Return {"topics":[{"name":str,"emphasis":"high|medium","why":str (why it deserves revision, citing how the lecture emphasizes it),"definitions":[{"term":str,"definition":str}],"relationships":[str],"practice":[{"q":str,"hint":str}]}]} with 5-10 topics, most emphasized first.\n\nLECTURE:\n' + l['text'][:MAXC])
        l['prep'] = json.dumps(r); run('update lectures set prep=? where id=?', (l['prep'], i))
    return jsonify(prep=json.loads(l['prep']))

# ---------- AI weakness detector ----------
def history():
    out = []
    for a in rows('select id,lecture_id,detail from attempts where user_id=? order by id', (g.uid,)):
        out += [{**r, 'lecture_id': r.get('lecture_id') or a['lecture_id']} for r in json.loads(a['detail'])['review']]
    return out
def weak_info():
    h = history(); t = {}
    for r in h:
        x = t.setdefault(r['topic'], {'topic': r['topic'], 'sum': 0, 'n': 0, 'wrong': 0})
        x['sum'] += r['score']; x['n'] += 1; x['wrong'] += r['score'] < .5
    weak = sorted([{'topic': x['topic'], 'pct': round(x['sum'] / x['n'] * 100), 'asked': x['n'], 'mistakes': x['wrong']} for x in t.values() if x['sum'] / x['n'] < .6], key=lambda x: x['pct'])[:8]
    names = {w['topic'] for w in weak}; lec = {}
    for r in h:
        if r['topic'] in names and r['score'] < .5: lec[r['lecture_id']] = lec.get(r['lecture_id'], 0) + 1
    return h, weak, lec
@app.get('/api/weakness')
@auth
def weakness():
    h, weak, lec = weak_info()
    if not h: return jsonify(empty=True)
    lects = [{'id': i, 'title': (one('select title from lectures where id=?', (i,)) or {'title': '?'})['title'], 'mistakes': m} for i, m in sorted(lec.items(), key=lambda x: -x[1])]
    out = {'weak': weak, 'lectures': lects, 'summary': '', 'gaps': []}
    miss = [{k: r[k] for k in ('topic', 'q', 'student', 'correct')} for r in h if r['score'] < .5][-15:]
    if weak and miss:
        r = ask('You are an academic coach. Use only the evidence given. Return JSON.', json.dumps({'weak_topics': weak, 'recent_mistakes': miss}) +
                '\nReturn {"summary":str (2-3 sentences),"gaps":[{"topic":str,"gap":str (the specific misunderstanding),"fix":str (what to review)}]}')
        out['summary'] = r.get('summary', ''); out['gaps'] = r.get('gaps', [])
    return jsonify(out)
@app.post('/api/weakness/quiz')
@auth
def weakness_quiz():
    _, weak, lec = weak_info()
    if not weak or not lec: return err('No weak topics yet. Keep taking exams and check back.')
    return build_exam({'lecture_ids': list(lec), 'topics': [w['topic'] for w in weak], 'count': 10, 'mode': 'practice', 'professor': 'Normal', 'duration': 15})

# ---------- study battle ----------
def bget(code):
    b = one('select * from battles where code=?', (str(code).upper().strip(),))
    if b: b['questions'] = json.loads(b['questions'])
    return b
def member(b): return b and one('select 1 from battle_players where battle_id=? and user_id=?', (b['id'], g.uid))
def bstate(b):
    n, st, idx, left, starts = len(b['questions']), b['status'], -1, 0, 0
    if st == 'running':
        el = time.time() - b['started_at']
        if el < 0: starts = -el
        else:
            idx = int(el // b['per_q'])
            if idx >= n: st, idx = 'finished', -1; run('update battles set status=? where id=?', ('finished', b['id']))
            else: left = b['per_q'] - (el - idx * b['per_q'])
    return st, idx, left, starts
def board(b):
    return rows('select p.user_id,p.name,coalesce(sum(a.points),0) score,coalesce(sum(a.correct),0) correct from battle_players p left join battle_answers a on a.battle_id=p.battle_id and a.user_id=p.user_id where p.battle_id=? group by p.user_id order by score desc, p.name', (b['id'],))
@app.post('/api/battles')
@auth
def battle_new():
    d = request.get_json(silent=True) or {}; ls = lectures_for(d)
    if not ls: return err('Choose at least one lecture')
    n = min(max(int(d.get('count', 5)), 3), 20); per = min(max(int(d.get('seconds', 20)), 10), 60)
    qs = [x for x in gen_questions(ls, n, 'Medium', 'Normal', d.get('language', 'English'), None, '100% mcq') if x['type'] == 'mcq']
    if not qs: return err('Could not generate questions. Try again.')
    code = ''.join(random.choices('ABCDEFGHJKLMNPQRSTUVWXYZ23456789', k=5))
    bid = run('insert into battles(code,host_id,title,questions,per_q) values(?,?,?,?,?)', (code, g.uid, ' + '.join(l['title'][:30] for l in ls), json.dumps(qs), per))
    run('insert into battle_players(battle_id,user_id,name) values(?,?,?)', (bid, g.uid, one('select name from users where id=?', (g.uid,))['name']))
    return jsonify(code=code)
@app.post('/api/battles/join')
@auth
def battle_join():
    b = bget((request.get_json(silent=True) or {}).get('code', ''))
    if not b: return err('No room with that code', 404)
    if b['status'] != 'lobby' and not member(b): return err('This battle has already started')
    run('insert or ignore into battle_players(battle_id,user_id,name) values(?,?,?)', (b['id'], g.uid, one('select name from users where id=?', (g.uid,))['name']))
    return jsonify(code=b['code'])
@app.get('/api/battles/<code>')
@auth
def battle_state(code):
    b = bget(code)
    if not member(b): return err('Room not found', 404)
    st, idx, left, starts = bstate(b)
    out = {'code': b['code'], 'title': b['title'], 'status': st, 'host': b['host_id'] == g.uid, 'me': g.uid, 'total': len(b['questions']), 'per_q': b['per_q'],
           'idx': idx, 'left': round(left, 1), 'starts_in': round(starts, 1), 'board': board(b)}
    if st == 'running' and idx >= 0:
        x = b['questions'][idx]; out['question'] = {'q': x['q'], 'options': x['options'], 'topic': x['topic']}
        out['answered'] = one('select correct,points from battle_answers where battle_id=? and user_id=? and qidx=?', (b['id'], g.uid, idx))
    return jsonify(out)
@app.post('/api/battles/<code>/start')
@auth
def battle_start(code):
    b = bget(code)
    if not member(b): return err('Room not found', 404)
    if b['host_id'] != g.uid: return err('Only the host can start the battle', 403)
    if b['status'] == 'lobby': run('update battles set status=?, started_at=? where id=?', ('running', time.time() + 5, b['id']))
    return jsonify(ok=True)
@app.post('/api/battles/<code>/answer')
@auth
def battle_answer(code):
    b = bget(code)
    if not member(b): return err('Room not found', 404)
    st, idx, left, _ = bstate(b); d = request.get_json(silent=True) or {}
    if st != 'running' or idx < 0 or d.get('idx') != idx: return err('That question is closed')
    x = b['questions'][idx]
    try: ok = int(d.get('answer')) == int(x['answer'])
    except (TypeError, ValueError): ok = False
    pts = 500 + int(500 * left / b['per_q']) if ok else 0
    try: run('insert into battle_answers(battle_id,user_id,qidx,correct,points) values(?,?,?,?,?)', (b['id'], g.uid, idx, int(ok), pts))
    except sqlite3.IntegrityError: return err('You already answered this question')
    return jsonify(correct=ok, points=pts, right=int(x['answer']))

@app.get('/healthz')
def health(): return 'ok'

@app.get('/')
def index(): return app.send_static_file('index.html')

if __name__ == '__main__': app.run(debug=os.getenv('FLASK_DEBUG') == '1', port=int(os.getenv('PORT', 5000)))
