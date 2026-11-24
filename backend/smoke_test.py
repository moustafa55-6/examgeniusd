"""Offline end-to-end test: mocks the OpenAI call, exercises every endpoint. Run: python smoke_test.py"""
import os, sys, re, io, json, time, sqlite3, tempfile
os.environ['DATA_DIR'] = tempfile.mkdtemp(); sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx, app as A
def fake(system, user):
    if 'analyze university lectures' in system: return {'title': 'Lecture', 'summary': 's', 'topics': [{'name': 'T', 'key_points': ['k']}]}
    if 'writing an exam' in system:
        n = int(re.search(r'exactly (\d+) questions', user).group(1)); qs = []
        for i in range(n):
            t = 'mcq' if '100% mcq' in user else A.TYPES[i % 6]
            x = {'type': t, 'topic': f'Topic {i % 3}', 'q': f'Q{i}?', 'explanation': 'e', 'reference': 'Page 1', 'lecture': 1 + i % 2}
            x.update({'mcq': dict(options=list('abcd'), answer=1), 'tf': dict(answer=True), 'fill': dict(answer='word'), 'match': dict(pairs=[['l1', 'r1'], ['l2', 'r2']], answer=''), 'short': dict(answer='m', rubric=['r']), 'scenario': dict(answer='m', rubric=['r'])}[t]); qs.append(x)
        return {'questions': qs}
    if 'short_answers' in user:
        p = json.loads(user.split('\nReturn')[0]); return {'grades': [{'i': a['i'], 'score': .5, 'comment': 'ok'} for a in p['short_answers']], 'strong': ['Topic 1'], 'weak': ['Topic 0'], 'revise': ['Topic 0'], 'tips': ['t1']}
    if '"score":0-1,"comment"' in user: return {'score': .8, 'comment': 'good'}
    if 'patient tutor' in system: return {'explanation': 'explained'}
    if 'exam preparation coach' in system: return {'topics': [{'name': 'N', 'emphasis': 'high', 'why': 'w', 'definitions': [], 'relationships': [], 'practice': [{'q': 'q', 'hint': 'h'}]}]}
    if 'academic coach' in system: return {'summary': 'sum', 'gaps': [{'topic': 'Topic 0', 'gap': 'g', 'fix': 'f'}]}
    raise Exception('unexpected prompt: ' + system[:50])
A.ask = fake; c = A.app.test_client()
def call(m, p, tok=None, body=None, **kw):
    r = getattr(c, m)(p, json=body, headers={'Authorization': 'Bearer ' + tok} if tok else {}, **kw); return r.status_code, (r.get_json(silent=True) or {})
def user(n): s, d = call('post', '/api/register', body={'name': n, 'email': n + '@x.io', 'password': 'password1'}); assert s == 200, d; return d['token']
def doc(t):
    d = docx.Document(); d.add_paragraph((t + ' lorem ipsum dolor sit amet. ') * 10); b = io.BytesIO(); d.save(b); b.seek(0); return b
a, b, z = user('amy'), user('bob'), user('zed')
ids = []
for n in ('one', 'two'):
    s, d = call('post', '/api/lectures', a, data={'file': (doc(n), n + '.docx')}, content_type='multipart/form-data'); assert s == 200, d; ids.append(d['id'])
s, d = call('get', f'/api/lectures/{ids[0]}/prep', a); assert s == 200 and d['prep']['topics'][0]['why'], d
s, d = call('post', '/api/exams', a, {'lecture_id': ids[0], 'count': 6, 'professor': 'Strict'}); assert s == 200 and len(d['questions']) == 6, d
assert json.loads(A.sqlite3.connect(A.DB).execute('select config from exams where id=?', (d['id'],)).fetchone()[0])['professor'] == 'Strict'
s, r = call('post', f"/api/exams/{d['id']}/check", a, {'index': 4, 'answer': 'text'}); assert s == 200 and r['score'] == .8, r
s, r = call('post', f"/api/exams/{d['id']}/submit", a, {'answers': {'0': 1, '1': 'true', '2': 'word', '3': ['r1', 'r2'], '4': 'x', '5': 'y'}, 'seconds': 90}); assert s == 200, r
s, r = call('get', f"/api/attempts/{r['id']}", a); assert r['attempt']['score'] > 60 and r['attempt']['review'], r
s, d = call('post', '/api/exams', a, {'lecture_ids': ids, 'count': 6, 'professor': 'Expert', 'duration': 120}); assert s == 200 and {q['lecture_id'] for q in d['questions']} == set(ids), d
s, r = call('post', f"/api/exams/{d['id']}/submit", a, {'answers': {}, 'seconds': 30}); assert s == 200, r
s, r = call('get', f"/api/attempts/{r['id']}", a); at = r['attempt']; assert at['title'].startswith('Final Exam') and len(at['lectures']) == 2 and at['score'] < 30, at
s, st = call('get', '/api/stats', a); assert s == 200 and st['total'] == 2 and st['weakest'], st
s, w = call('get', '/api/weakness', a); assert s == 200 and w['weak'] and w['lectures'] and w['gaps'], w
s, q = call('post', '/api/weakness/quiz', a, {}); assert s == 200 and q['mode'] == 'practice' and q['questions'], q
assert call('get', '/api/weakness', z)[1].get('empty')
code = call('post', '/api/battles', a, {'lecture_ids': ids, 'count': 3, 'seconds': 20})[1]['code']
assert call('post', '/api/battles/join', b, {'code': code.lower()})[0] == 200
assert call('get', f'/api/battles/{code}', z)[0] == 404
s, st = call('get', f'/api/battles/{code}', b); assert st['status'] == 'lobby' and len(st['board']) == 2 and not st['host'], st
assert call('post', f'/api/battles/{code}/start', b)[0] == 403 and call('post', f'/api/battles/{code}/start', a)[0] == 200
s, st = call('get', f'/api/battles/{code}', b); assert st['status'] == 'running' and st['starts_in'] > 0, st
sql = sqlite3.connect(A.DB); sql.execute('update battles set started_at=?', (time.time() - 1,)); sql.commit()
s, st = call('get', f'/api/battles/{code}', b); assert st['idx'] == 0 and 'answer' not in json.dumps(st['question']), st
s, r1 = call('post', f'/api/battles/{code}/answer', b, {'idx': 0, 'answer': 1}); assert r1['correct'] and r1['points'] >= 500, r1
assert call('post', f'/api/battles/{code}/answer', b, {'idx': 0, 'answer': 1})[0] == 400
assert not call('post', f'/api/battles/{code}/answer', a, {'idx': 0, 'answer': 0})[1]['correct']
s, st = call('get', f'/api/battles/{code}', a); assert st['board'][0]['name'] == 'bob' and st['board'][0]['score'] == r1['points'], st
sql.execute('update battles set started_at=?', (time.time() - 1000,)); sql.commit()
s, st = call('get', f'/api/battles/{code}', a); assert st['status'] == 'finished', st
assert call('post', f'/api/battles/{code}/answer', a, {'idx': 0, 'answer': 1})[0] == 400
print('ALL BACKEND TESTS PASSED')
