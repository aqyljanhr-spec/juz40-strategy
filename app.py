"""JUZ40 Strategy — FastAPI/SQLite standalone event session server."""
import os,json,sqlite3,threading,datetime,io,zipfile,urllib.request,hmac,hashlib,base64,time,secrets
import psycopg
from psycopg_pool import ConnectionPool
from pathlib import Path
from fastapi import FastAPI,HTTPException,WebSocket,WebSocketDisconnect,Request
from fastapi.responses import FileResponse,StreamingResponse,JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from fastapi import Cookie
from openpyxl import Workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

ROOT=Path(__file__).parent
STAFF=json.loads((ROOT/'staff.json').read_text())
GROUPS=list(dict.fromkeys(p['group'] for p in STAFF))
DEFAULT_CAPTAINS={g:next(x['id'] for x in STAFF if x['group']==g) for g in GROUPS}
DB=Path(os.environ.get('JUZ40_DB',str(ROOT/'strategy.sqlite3')))
DATABASE_URL=os.environ.get('DATABASE_URL','')
app=FastAPI(title='JUZ40 Strategy Platform')
ADMIN_NAME='Дүйсенғалиұлы Ақылжан'
ADMIN_PASSWORD=os.environ.get('SUPER_ADMIN_PASSWORD','')
ADMIN_SECRET=os.environ.get('SUPER_ADMIN_SESSION_SECRET','')
if not ADMIN_PASSWORD or len(ADMIN_PASSWORD)<12 or not ADMIN_SECRET or len(ADMIN_SECRET)<32:
 raise RuntimeError('Set SUPER_ADMIN_PASSWORD (12+ chars) and SUPER_ADMIN_SESSION_SECRET (32+ chars) in Render Environment')
def admin_cookie():
 expiry=str(int(time.time())+8*3600)
 digest=hmac.new(ADMIN_SECRET.encode(),expiry.encode(),hashlib.sha256).hexdigest()
 return expiry+'.'+digest
def require_admin(request:Request):
 token=request.cookies.get('juz40_superadmin','')
 try:
  expiry,mac=token.split('.',1)
  if int(expiry)<time.time() or not hmac.compare_digest(hmac.new(ADMIN_SECRET.encode(),expiry.encode(),hashlib.sha256).hexdigest(),mac):raise ValueError()
 except (ValueError,TypeError):raise HTTPException(401,'Әкімші ретінде кіріңіз')
def verify_origin(request:Request):
 origin=request.headers.get('origin')
 if origin:
  from urllib.parse import urlsplit
  expected=f'{request.url.scheme}://{request.headers.get("host", "")}'
  # Render terminates HTTPS in a reverse proxy
  forwarded=request.headers.get('x-forwarded-proto',request.url.scheme)
  expected=f'{forwarded}://{request.headers.get("host", "")}'
  if origin!=expected:raise HTTPException(403,'Қате Origin')
class AdminCredentials(BaseModel):
 password:str
@app.post('/api/admin/login')
def admin_login(credentials:AdminCredentials,request:Request):
 verify_origin(request)
 if not hmac.compare_digest(credentials.password,ADMIN_PASSWORD):raise HTTPException(401,'Құпиясөз дұрыс емес')
 response=JSONResponse({'ok':True,'name':ADMIN_NAME})
 response.set_cookie('juz40_superadmin',admin_cookie(),httponly=True,secure=True,samesite='strict',max_age=8*3600,path='/')
 return response
@app.get('/api/admin/me')
def admin_me(request:Request):
 require_admin(request)
 return {'name':ADMIN_NAME,'role':'super_admin'}
@app.post('/api/admin/logout')
def admin_logout(request:Request):
 verify_origin(request)
 response=JSONResponse({'ok':True})
 response.delete_cookie('juz40_superadmin',path='/')
 return response

lock=threading.RLock()
clients={'test':set(),'live':set()}
presence={'test':{},'live':{}}
presence_lock=threading.Lock()

class Action(BaseModel):
    mode:str='test'
    actor:int=0
    action:str
    data:dict={}

def now(): return datetime.datetime.now(datetime.timezone.utc).isoformat()
STAGE_PRESETS = {
 'issues': {'title':'Жеке мәселелерді жазу','minutes':10,'block':0,'task':'Бөлімдегі негізгі мәселелерді жеке жазыңыз.'},
 'vote': {'title':'Мәселелерге дауыс беру','minutes':5,'block':0,'task':'Өз тобыңыздағы ең маңызды үш мәселені таңдаңыз.'},
 'solutions': {'title':'ТОП-3 шешімдерді әзірлеу','minutes':15,'block':0,'task':'Үш басым мәселенің нақты шешімін бірге дайындаңыз.'},
 'vision': {'title':'Vision · 1 / 3 / 5 жыл','minutes':20,'block':1,'task':'Бөлім мен JUZ40-тың болашағын талқылаңыз.'},
 'changes': {'title':'Қосу / Өзгерту / Алып тастау','minutes':15,'block':2,'task':'Қандай жұмысты жеңілдетіп, автоматтандыруға болатынын анықтаңыз.'},
 'risks': {'title':'ҰБТ өзгерістері мен тәуекелдер','minutes':15,'block':3,'task':'ҰБТ өзгерістері мен алдын алу жоспарын талқылаңыз.'},
 'awards': {'title':'Ең белсенділерге дауыс беру','minutes':5,'block':None,'task':'Жалпы номинацияға басқа топтан, топтыққа өз тобыңыздан кандидат таңдаңыз.'},
}
STAGE_KEYS=list(STAGE_PRESETS)

def default_live_stage():
 duration=STAGE_PRESETS['issues']['minutes']*60000
 return {'screen':'work','stage_id':'issues','block':0,'presets':{key:val['minutes'] for key,val in STAGE_PRESETS.items()},
         'duration_ms':duration,'remaining_ms':duration,'deadline_ms':None,'running':False}

def normalize_live_stage(state):
 stage=state.setdefault('live_stage',default_live_stage())
 old_schema='stage_id' not in stage
 old_block=stage.get('block',0)
 defaults=default_live_stage()
 for key,val in defaults.items():stage.setdefault(key,val)
 # Upgrade existing LIVE/TEST sessions without changing a timer already in progress.
 if old_schema or stage.get('stage_id') not in STAGE_PRESETS:
  stage['stage_id']={0:'issues',1:'vision',2:'changes',3:'risks'}.get(old_block,'issues')
 if not isinstance(stage.get('presets'),dict):stage['presets']=defaults['presets'].copy()
 for key,val in defaults['presets'].items():
  if not isinstance(stage['presets'].get(key),int) or not 1<=stage['presets'][key]<=240:
   stage['presets'][key]=val
 stage['block']=STAGE_PRESETS[stage['stage_id']]['block']
 return stage

def choose_stage(stage, stage_id):
 if stage_id not in STAGE_PRESETS:raise HTTPException(400,'Белгісіз кезең')
 if stage.get('running'):raise HTTPException(409,'Алдымен таймерді паузаға қойыңыз')
 stage['stage_id']=stage_id
 stage['block']=STAGE_PRESETS[stage_id]['block']
 duration=stage['presets'][stage_id]*60000
 stage['duration_ms']=duration
 stage['remaining_ms']=duration
 stage['deadline_ms']=None
 stage['running']=False

def apply_stage_command(state, payload):
 stage=normalize_live_stage(state)
 op=str(payload.get('operation',''))
 current_ms=int(time.time()*1000)
 def remaining():
  return max(0,int(stage['deadline_ms'])-current_ms) if stage['running'] and stage['deadline_ms'] is not None else max(0,int(stage['remaining_ms']))
 if op=='select_stage':
  choose_stage(stage,str(payload.get('stage_id','')))
 elif op=='set_presets':
  proposed=payload.get('presets')
  if not isinstance(proposed,dict) or set(proposed)!=set(STAGE_KEYS):raise HTTPException(400,'Барлық кезеңнің уақыты көрсетілуі тиіс')
  for key,value in proposed.items():
   if isinstance(value,bool) or not isinstance(value,int) or not 1<=value<=240:
    raise HTTPException(400,'Әр кезең 1–240 минут болуы тиіс')
  stage['presets']=proposed.copy()
  if not stage['running']:
   duration=proposed[stage['stage_id']]*60000
   stage['duration_ms']=duration
   stage['remaining_ms']=duration
   stage['deadline_ms']=None
 elif op=='set_duration':
  try:minutes=int(payload.get('minutes',0))
  except (ValueError,TypeError):raise HTTPException(400,'Уақыт дұрыс емес')
  if not 1<=minutes<=240:raise HTTPException(400,'1–240 минут аралығы')
  if stage['running']:raise HTTPException(409,'Алдымен таймерді паузаға қойыңыз')
  stage['presets'][stage['stage_id']]=minutes
  stage['duration_ms']=stage['remaining_ms']=minutes*60000
  stage['deadline_ms']=None
 elif op in ('start','resume'):
  if stage['running']:return
  rest=remaining()
  if op=='start' and rest<=0:rest=int(stage['duration_ms'])
  if rest<=0:raise HTTPException(409,'Таймер біткен. Қайта орнатыңыз')
  stage['remaining_ms']=rest
  stage['deadline_ms']=current_ms+rest
  stage['running']=True
 elif op=='pause':
  if not stage['running']:return
  stage['remaining_ms']=remaining()
  stage['deadline_ms']=None
  stage['running']=False
 elif op=='reset':
  stage['running']=False
  stage['duration_ms']=stage['presets'][stage['stage_id']]*60000
  stage['remaining_ms']=int(stage['duration_ms'])
  stage['deadline_ms']=None
 elif op=='add':
  try:minutes=int(payload.get('minutes',0))
  except (ValueError,TypeError):raise HTTPException(400,'Уақыт дұрыс емес')
  if minutes not in (1,5):raise HTTPException(400,'Тек 1 немесе 5 минут қосуға болады')
  if stage['running']:
   stage['deadline_ms']=max(current_ms,int(stage['deadline_ms']))+minutes*60000
   stage['remaining_ms']=max(0,int(stage['deadline_ms'])-current_ms)
  else:stage['remaining_ms']=min(4*3600000,remaining()+minutes*60000)
  stage['duration_ms']=max(int(stage['duration_ms']),stage['remaining_ms'])
 else:raise HTTPException(400,'Таймер командасы белгісіз')

def prepare_stage_on_change(state, stage_id):
 # Admin's actual stage changes prime the matching clock, never auto-start it.
 stage=normalize_live_stage(state)
 if not stage['running']:choose_stage(stage,stage_id)

def default():
 return {'staff':[{**p,'captain':p['id']==DEFAULT_CAPTAINS[p['group']]} for p in STAFF], 'joined':[], 'live_stage':default_live_stage(), 'drafts':{},'activity':{},'awards_open':False,'awards_closed':False,'award_votes':{'overall':{},'group':{}}, 'opened':[False]*4,'closed':[False]*4,'phase':'writing','issues':[],'votes':{},'shared':{},'finished':{},'session_closed':False,'reports':[],'audit':[],'revision':0}
_pg_pool=None
_pg_lock=threading.Lock()
def connection():
  global _pg_pool
  if DATABASE_URL:
   if _pg_pool is None:
    with _pg_lock:
     if _pg_pool is None:
      pool=ConnectionPool(DATABASE_URL,min_size=1,max_size=5,timeout=12,open=True)
      with pool.connection() as c:
       c.execute('CREATE TABLE IF NOT EXISTS state (mode TEXT PRIMARY KEY,payload TEXT NOT NULL)')
      _pg_pool=pool
   return _pg_pool.connection()
  DB.parent.mkdir(parents=True,exist_ok=True)
  c=sqlite3.connect(str(DB),timeout=15)
  c.execute('CREATE TABLE IF NOT EXISTS state (mode TEXT PRIMARY KEY,payload TEXT NOT NULL)')
  return c

def load(mode):
 with connection() as c:
  if DATABASE_URL:r=c.execute('SELECT payload FROM state WHERE mode=%s',(mode,)).fetchone()
  else:r=c.execute('SELECT payload FROM state WHERE mode=?',(mode,)).fetchone()
 return json.loads(r[0]) if r else default()

def save(mode,s):
 with connection() as c:
  if DATABASE_URL:c.execute('INSERT INTO state(mode,payload) VALUES(%s,%s) ON CONFLICT(mode) DO UPDATE SET payload=excluded.payload',(mode,json.dumps(s,ensure_ascii=False)))
  else:c.execute('INSERT INTO state(mode,payload) VALUES(?,?) ON CONFLICT(mode) DO UPDATE SET payload=excluded.payload',(mode,json.dumps(s,ensure_ascii=False)))

def role(s,actor):
 p=next((p for p in s['staff'] if p['id']==actor),None)
 if not p: raise HTTPException(403,'Қызметкер тізімде жоқ')
 return p

def check_action(s,p,action):
 if action.startswith('admin_'): return # Authentication is enforced by the API route
 if s['session_closed'] and action not in ['report_save','award_vote']:raise HTTPException(409,'Сессия жабық')
 if action in ('award_vote',):
  if not s.get('awards_open') or s.get('awards_closed'):raise HTTPException(409,'Марапаттау дауысы қазір жабық')
 if action in ('issue_add','issue_edit','vote'):
  block=0
 elif action.startswith(('shared:','finish:','draft:')):
  block=int(action.split(':',1)[1])
 else:
  block=None
 if block is not None and (block<0 or block>=4 or not s['opened'][block] or s['closed'][block]):
  raise HTTPException(409,'Бұл блок жабық немесе әлі ашылмаған')

def mutate(s,p,action,d):
 uid=p['id'];group=p['group'];phase=s['phase']
 if action=='join':
  if uid not in s['joined']:s['joined'].append(uid)
 elif action=='issue_add':
  if phase!='writing':raise HTTPException(409,'Мәселе жазу кезеңі аяқталды')
  txt=str(d.get('text','')).strip()
  if not 3<=len(txt)<=3000:raise HTTPException(400,'Мәселені 3–3000 таңба аралығында жазыңыз')
  if len([x for x in s['issues'] if x['author']==uid])>=3:raise HTTPException(400,'Ең көбі 3 мәселе')
  n=max([x['id'] for x in s['issues']],default=0)+1;s['issues'].append({'id':n,'author':uid,'group':group,'text':txt})
 elif action=='issue_edit':
  if phase!='writing':raise HTTPException(409,'Мәселелер бекітілген')
  issue=next((x for x in s['issues'] if x['id']==int(d.get('id',-1)) and x['author']==uid),None)
  if not issue:raise HTTPException(404,'Мәселе табылмады')
  txt=str(d.get('text','')).strip()
  if len(txt)<3 or len(txt)>3000:raise HTTPException(400,'3–3000 таңба қажет')
  issue['text']=txt
 elif action=='vote':
  if phase!='voting':raise HTTPException(409,'Дауыс беру кезеңі ашық емес')
  ids=d.get('ids',[])
  if len(ids)!=3 or len(set(ids))!=3 or any(not any(x['id']==i and x['group']==group for x in s['issues']) for i in ids):raise HTTPException(400,'Өз тобыңыздың дәл 3 түрлі мәселесін таңдаңыз')
  s['votes'][str(uid)]=ids
 elif action=='award_vote':
  category=str(d.get('category',''));target_id=int(d.get('target',0))
  if category not in ('overall','group'):raise HTTPException(400,'Номинация қате')
  target=role(s,target_id)
  if target_id==uid:raise HTTPException(400,'Өзіңізге дауыс бере алмайсыз')
  if category=='group' and target['group']!=group:raise HTTPException(400,'Тек өз тобыңызға дауыс беріңіз')
  if category=='overall' and target['group']==group:raise HTTPException(400,'Басқа топтың қатысушысын таңдаңыз')
  s.setdefault('award_votes',{'overall':{},'group':{}}).setdefault(category,{})[str(uid)]=target_id
 elif action.startswith('draft:'):
  b=int(action.split(':')[1]);key=str(d.get('key',''))
  allowed={1:['department_1','department_3','department_5','company_1','company_3','company_5','publisher'],2:['answer'],3:['answer']}
  if p.get('captain'):raise HTTPException(403,'Капитан қаралама жазбайды')
  if b not in allowed or key not in allowed[b]:raise HTTPException(400,'Қате қаралама өрісі')
  if key=='publisher' and p['dept']!='Әдістеме':raise HTTPException(403,'Өріс тек Әдістеме үшін')
  if s['closed'][b] or s['finished'].get(group+'|'+str(b)):raise HTTPException(409,'Блок бекітілген')
  val=str(d.get('text',''))
  if len(val)>20000:raise HTTPException(400,'Мәтін тым ұзын')
  s.setdefault('drafts',{})[str(uid)+'|'+str(b)+'|'+key]=val
 elif action.startswith('shared:'):
  block=int(action.split(':')[1]);key=str(d.get('key',''))
  if not p.get('captain'):raise HTTPException(403,'Тек капитан жаза алады')
  if s['closed'][block] or s['finished'].get(group+'|'+str(block)):raise HTTPException(409,'Блок өңдеуге жабық')
  allowed={0:['solution_0','solution_1','solution_2','choice_2'],1:['department_1','department_3','department_5','company_1','company_3','company_5','publisher'],2:['answer'],3:['answer']}
  if key not in allowed[block] or key=='publisher' and p['dept']!='Әдістеме':raise HTTPException(400,'Қате өріс')
  if block==0 and phase!='final':raise HTTPException(409,'Қорытынды кезеңі ашылмаған')
  text=str(d.get('text',''))
  if len(text)>20000:raise HTTPException(400,'Мәтін тым ұзын')
  s['shared'][group+'|'+str(block)+'|'+key]=text
 elif action.startswith('finish:'):
  b=int(action.split(':')[1]);
  if not p.get('captain'):raise HTTPException(403,'Тек капитан')
  if s['closed'][b]:raise HTTPException(409,'Блок жабық')
  if b==0 and s['phase']!='final':raise HTTPException(409,'Қорытынды кезеңі ашылмаған')
  s['finished'][group+'|'+str(b)]=True
 elif action=='admin_awards':
  op=d.get('operation')
  if op=='open':s['awards_open']=True;s['awards_closed']=False;prepare_stage_on_change(s,'awards')
  elif op=='close':s['awards_open']=False;s['awards_closed']=True
  elif op=='reopen':s['awards_open']=True;s['awards_closed']=False
  else:raise HTTPException(400,'Қате мәртебе')
 elif action=='admin_live_stage':
  apply_stage_command(s,d)
 elif action=='admin_rename':
  target=role(s,int(d.get('id',0)));name=' '.join(str(d.get('name','')).split())
  if not 3<=len(name)<=140:raise HTTPException(400,'Аты-жөні 3–140 таңба болуы тиіс')
  if any(x['id']!=target['id'] and x['name'].casefold()==name.casefold() for x in s['staff']):raise HTTPException(409,'Бұл аты-жөн тізімде бар')
  target['name']=name
 elif action=='admin_stage':
  phase2=d.get('phase');
  if phase2 not in ['writing','voting','final']:raise HTTPException(400,'Қате кезең')
  s['phase']=phase2
  prepare_stage_on_change(s,{'writing':'issues','voting':'vote','final':'solutions'}[phase2])
 elif action=='admin_block':
  b=int(d['block']);s['opened'][b]=bool(d['open']);s['closed'][b]=bool(d.get('close',False)) if not d['open'] else False
  if d['open']:prepare_stage_on_change(s,{0:{'writing':'issues','voting':'vote','final':'solutions'}[s['phase']],1:'vision',2:'changes',3:'risks'}[b])
 elif action=='admin_reopen':
  b=int(d['block']);g=str(d['group']);s['finished'].pop(g+'|'+str(b),None);s['closed'][b]=False;s['opened'][b]=True
 elif action=='admin_assign':
  target=role(s,int(d['id']));g=str(d['group'])
  if g not in GROUPS or g.split(' · ')[0]!=target['dept']:raise HTTPException(400,'Топ бөлімге сай емес')
  old=target['group'];target['group']=g
  if target['captain']:target['captain']=False
  if not any(x['captain'] and x['group']==old for x in s['staff']):pass
 elif action=='admin_captain':
  target=role(s,int(d['id']));
  for x in s['staff']:
   if x['group']==target['group']:x['captain']=x['id']==target['id']
 elif action=='admin_edit':
  g=d['group'];b=int(d['block']);key=d['key'];val=str(d['text'])
  if g not in GROUPS or len(val)>20000:raise HTTPException(400,'Қате дерек')
  s['shared'][g+'|'+str(b)+'|'+str(key)]=val
 elif action=='admin_finish':
  s['session_closed']=True;s['opened']=[False]*4;s['closed']=[True]*4
 elif action=='admin_resume':s['session_closed']=False
 elif action=='report_save':
  text=str(d.get('text',''))
  if len(text)>40000:raise HTTPException(400,'Есеп тым ұзын')
  s['reports'].append({'version':len(s['reports'])+1,'text':text,'at':now(),'editor':uid})
 else:raise HTTPException(400,'Белгісіз әрекет')

@app.get('/api/roster/{mode}')
def roster_public(mode:str):
 if mode not in clients:raise HTTPException(400,'mode')
 s=load(mode)
 return {'staff':[{'id':p['id'],'name':p['name'],'dept':p['dept'],'group':p['group'],'captain':p['captain']} for p in s['staff']]}

@app.get('/api/state/{mode}')
def get_state(mode:str, request:Request, actor:int=0, admin:bool=False):
 if mode not in clients:raise HTTPException(400,'mode')
 s=load(mode)
 normalize_live_stage(s)
 s['server_now_ms']=int(time.time()*1000)
 if admin:require_admin(request)
 if not actor and not admin:raise HTTPException(401,'Қатысушыны таңдаңыз')
 if actor and not admin:
  p=role(s,actor);g=p['group'];s['issues']=[i for i in s['issues'] if i['group']==g]
  s['votes']={k:v for k,v in s['votes'].items() if any(int(k)==x['id'] and x['group']==g for x in s['staff'])} if s['phase']=='final' else {str(actor):s['votes'].get(str(actor),[])}
  s['shared']={k:v for k,v in s['shared'].items() if k.startswith(g+'|')}
  s['reports']=[];s['audit']=[]
  s['drafts']={k:v for k,v in s.get('drafts',{}).items() if k.startswith(str(actor)+'|') or (p.get('captain') and any(k.startswith(str(member['id'])+'|') for member in s['staff'] if member['group']==g))}
  s['activity']={}
  votes=s.get('award_votes',{'overall':{},'group':{}})
  s['award_votes']={category:({str(actor):votes.get(category,{}).get(str(actor))} if not s.get('awards_closed') else {}) for category in ('overall','group')}
  if not s.get('awards_closed'):s.pop('award_results',None)
 if admin:
  with presence_lock:
   s['online_ids']=[int(i) for i,t in presence[mode].items() if time.monotonic()-t['at']<75]
   s['active_views']={str(i):v['view'] for i,v in presence[mode].items() if time.monotonic()-v['at']<75}
 if admin and s.get('awards_closed'):
  s['award_results']={category:{str(p['id']):sum(1 for target in s.get('award_votes',{}).get(category,{}).values() if target==p['id']) for p in s['staff']} for category in ('overall','group')}
 return s

class Heartbeat(BaseModel):
 mode:str='test'
 actor:int
 view:str='home'
@app.get('/api/stage/{mode}')
def get_stage_public(mode:str):
 if mode not in clients:raise HTTPException(400,'mode')
 stage=normalize_live_stage(load(mode))
 sid=stage['stage_id']
 return {'stage':stage,'title':STAGE_PRESETS[sid]['title'],'task':STAGE_PRESETS[sid]['task'],
         'server_now_ms':int(time.time()*1000),'groups':len(GROUPS),'participants':len(STAFF)}

@app.post('/api/heartbeat')
def heartbeat(h:Heartbeat,request:Request):
 verify_origin(request)
 if h.mode not in presence:raise HTTPException(400,'mode')
 if h.actor==-1:require_admin(request)
 elif h.actor not in {p['id'] for p in STAFF}:raise HTTPException(403,'Қатысушы жоқ')
 with presence_lock:presence[h.mode][str(h.actor)]={'at':time.monotonic(),'view':h.view[:32]}
 return {'ok':True}

@app.post('/api/action')
async def post_action(payload:Action,request:Request):
 if payload.mode not in clients:raise HTTPException(400,'mode')
 verify_origin(request)
 admin_action=payload.action.startswith('admin_') or payload.action=='report_save'
 if admin_action:require_admin(request)
 with lock:
  s=load(payload.mode);p={'id':0,'name':ADMIN_NAME,'group':'','dept':'Әкімшілік'} if admin_action else role(s,payload.actor)
  check_action(s,p,payload.action)
  mutate(s,p,payload.action,payload.data)
  if not admin_action:s.setdefault('activity',{})[str(payload.actor)]={'at':now(),'action':payload.action}
  s['revision']+=1
  s['audit'].append({'at':now(),'actor':p['name'],'action':payload.action,'data':{k:(v[:100] if isinstance(v,str) else v) for k,v in payload.data.items()}})
  s['audit']=s['audit'][-1000:]
  save(payload.mode,s)
 dead=[]
 for ws in list(clients[payload.mode]):
  try:await ws.send_json({'revision':s['revision']})
  except:dead.append(ws)
 for w in dead:clients[payload.mode].discard(w)
 return {'ok':True,'revision':s['revision']}

@app.websocket('/api/live/{mode}')
async def live(websocket:WebSocket,mode:str):
 if mode not in clients:await websocket.close(code=1008);return
 await websocket.accept();clients[mode].add(websocket)
 try:
  while True:await websocket.receive_text()
 except WebSocketDisconnect:pass
 finally:clients[mode].discard(websocket)

@app.get('/api/export/{mode}.xlsx')
def excel(mode:str,request:Request):
 if mode not in clients:raise HTTPException(400,'mode')
 require_admin(request)
 s=load(mode);wb=Workbook();ws=wb.active;ws.title='Қатысушылар';ws.append(['Аты-жөні','Бөлім','Топ','Капитан','Check-in'])
 for p in s['staff']:ws.append([p['name'],p['dept'],p['group'],'Иә' if p['captain'] else 'Жоқ','Иә' if p['id'] in s['joined'] else 'Жоқ'])
 w=wb.create_sheet('Мәселелер');w.append(['Топ','Қызметкер','Мәселе','Дауыс'])
 for x in s['issues']:w.append([x['group'],next((p['name'] for p in s['staff'] if p['id']==x['author']),''),x['text'],sum(x['id'] in votes for votes in s['votes'].values())])
 w=wb.create_sheet('Топ жауаптары');w.append(['Топ','Блок','Өріс','Жауап'])
 for k,v in s['shared'].items():
  g,b,f=k.split('|',2);w.append([g,int(b)+1,f,v])
 w=wb.create_sheet('AI есеп нұсқалары');w.append(['Нұсқа','Уақыты','Мәтін'])
 for r in s['reports']:w.append([r['version'],r['at'],r['text']])
 for sh in wb:
  sh.freeze_panes='A2';sh.auto_filter.ref=sh.dimensions
  for col in sh.columns:
   letter=col[0].column_letter;sh.column_dimensions[letter].width=min(58,max(16,max(len(str(c.value or '')) for c in col[:50])+3))
 data=io.BytesIO();wb.save(data);data.seek(0)
 return StreamingResponse(data,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':f'attachment; filename="juz40-{mode}.xlsx"'})

@app.get('/api/team/{mode}/{group_no}/pdf')
def pdf(mode:str,group_no:int,actor:int):
 if mode not in clients:raise HTTPException(400,'mode')
 s=load(mode);p=role(s,actor);g=p['group']
 if GROUPS.index(g)!=group_no:raise HTTPException(403,'Тек өз тобыңыз')
 from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle,KeepTogether,PageBreak
 from reportlab.lib.styles import ParagraphStyle
 from reportlab.lib import colors
 from reportlab.lib.enums import TA_LEFT
 from reportlab.lib.utils import simpleSplit
 from xml.sax.saxutils import escape
 font='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
 boldfont='/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'
 if os.path.isfile(font):pdfmetrics.registerFont(TTFont('JuzBody',font))
 if os.path.isfile(boldfont):pdfmetrics.registerFont(TTFont('JuzBold',boldfont))
 f='JuzBody' if os.path.isfile(font) else 'Helvetica';fb='JuzBold' if os.path.isfile(boldfont) else 'Helvetica-Bold'
 ink=colors.HexColor('#173723');green=colors.HexColor('#149447');soft=colors.HexColor('#F0F8F2');muted=colors.HexColor('#617867')
 buf=io.BytesIO();doc=SimpleDocTemplate(buf,pagesize=A4,rightMargin=40,leftMargin=40,topMargin=49,bottomMargin=48,title='JUZ40 Strategy · '+g)
 title=ParagraphStyle('JTitle',fontName=fb,fontSize=24,leading=31,textColor=ink,spaceAfter=14)
 heading=ParagraphStyle('JHeading',fontName=fb,fontSize=15,leading=21,textColor=ink,spaceBefore=19,spaceAfter=10)
 sub=ParagraphStyle('JSub',fontName=fb,fontSize=11,leading=17,textColor=ink,spaceAfter=6)
 body=ParagraphStyle('JBody',fontName=f,fontSize=10,leading=16,textColor=ink,spaceAfter=8,wordWrap='CJK')
 small=ParagraphStyle('JSmall',fontName=f,fontSize=9,leading=14,textColor=muted)
 def P(v,style=body):return Paragraph(escape(str(v or '—')).replace('\n','<br/>'),style)
 def panel(label,value):
  data=[[P(label,sub)],[P(value,body)]]
  t=Table(data,colWidths=[A4[0]-80],hAlign='LEFT');t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),soft),('BOX',(0,0),(-1,-1),0.6,colors.HexColor('#D5E8D9')),('TOPPADDING',(0,0),(-1,0),12),('BOTTOMPADDING',(0,-1),(-1,-1),13),('LEFTPADDING',(0,0),(-1,-1),13),('RIGHTPADDING',(0,0),(-1,-1),13)]));return t
 story=[Spacer(1,26),P('JUZ40 STRATEGY',ParagraphStyle('Brand',parent=sub,fontSize=17,textColor=green)),Spacer(1,18),P('Стратегиялық сессия қорытындысы',title),P(g,heading),P('PRODUCT & IT  •  Әдістеме / Сапа / IT',small),Spacer(1,25),panel('Топтың стратегиялық есебі','Төрт блок бойынша сақталған жауаптар және негізгі қорытындылар'),Spacer(1,24),P('TEST — сынақ нәтижелері' if mode=='test' else 'LIVE — стратегиялық сессия',small),PageBreak()]
 titles=['1. Қазіргі жағдай: ТОП-3 мәселе','2. Vision · 1 / 3 / 5 жыл','3. Қосу / Өзгерту / Алып тастау','4. ҰБТ өзгерістері және тәуекелдер']
 labels={'department_1':'Бөлім — 1 жыл','department_3':'Бөлім — 3 жыл','department_5':'Бөлім — 5 жыл','company_1':'JUZ40 компаниясы — 1 жыл','company_3':'JUZ40 компаниясы — 3 жыл','company_5':'JUZ40 компаниясы — 5 жыл','publisher':'JUZ40 Баспасының болашағы','answer':'Топтың ортақ ұсынысы'}
 for b,t in enumerate(titles):
  story.append(P(t,heading))
  if b==0:
   issues=[dict(x) for x in s['issues'] if x['group']==g]
   for x in issues:x['count']=sum(x['id'] in vote for vote in s['votes'].values())
   issues.sort(key=lambda x:(-x['count'],x['id']))
   if not issues:story.append(P('Әзірге мәселелер жоқ',small))
   for i,x in enumerate(issues[:3]):story.append(panel('№'+str(i+1)+' · '+str(x['count'])+' дауыс · '+x['text'],s['shared'].get(g+'|0|solution_'+str(i),'Шешім енгізілмеген')));story.append(Spacer(1,8))
  else:
   entries=[(k.split('|',2)[2],v) for k,v in s['shared'].items() if k.startswith(g+'|'+str(b)+'|')]
   if not entries:story.append(P('Әзірге жауап енгізілмеген',small))
   for key,val in entries:story.append(panel(labels.get(key,key),val));story.append(Spacer(1,9))
  story.append(Spacer(1,14))
 def decorate(canvas,doc):
  canvas.saveState();w,h=A4;canvas.setStrokeColor(colors.HexColor('#DDE9E0'));canvas.line(40,35,w-40,35);canvas.setFont(f,8);canvas.setFillColor(muted);canvas.drawString(40,24,'JUZ40 STRATEGY  •  '+('TEST' if mode=='test' else 'LIVE'));canvas.drawRightString(w-40,24,str(doc.page));canvas.restoreState()
 doc.build(story,onFirstPage=decorate,onLaterPages=decorate);buf.seek(0)
 return StreamingResponse(buf,media_type='application/pdf',headers={'Content-Disposition':f'attachment; filename="juz40-group-{group_no+1}.pdf"'})

class AIRequest(BaseModel):
 mode:str='live'
 scope:str='all'
 kind:str='strategic'
 actor:int=0
@app.post('/api/ai')
def ai(req:AIRequest,request:Request):
 require_admin(request)
 verify_origin(request)
 if not os.environ.get('OPENAI_API_KEY'):raise HTTPException(503,'AI API кілті орнатылмаған. OPENAI_API_KEY қажет.')
 s=load(req.mode)
 group_names=GROUPS if req.scope=='all' else [g for g in GROUPS if g.startswith(req.scope+' ·')]
 answers={k:v for k,v in s['shared'].items() if any(k.startswith(g+'|') for g in group_names)}
 prompt='JUZ40 стратегиялық сессиясы. Қазақ тілінде '+req.kind+' талдау жасаңыз. Тек төмендегі бастапқы жауаптарға сүйеніңіз; қолдау таппаған тұжырымдарды гипотеза деп белгілеңіз; әр ұсыныстың дереккөз-тобын көрсетіңіз; нақты адам мен KPI ойдан қоспаңыз.\n'+json.dumps(answers,ensure_ascii=False)
 body={'model':os.getenv('OPENAI_MODEL','gpt-4.1-mini'),'messages':[{'role':'user','content':prompt}],'temperature':0.2}
 request=urllib.request.Request('https://api.openai.com/v1/chat/completions',data=json.dumps(body).encode(),headers={'Authorization':'Bearer '+os.environ['OPENAI_API_KEY'],'Content-Type':'application/json'})
 try:
  with urllib.request.urlopen(request,timeout=60) as res:data=json.load(res)
  return {'report':data['choices'][0]['message']['content']}
 except Exception as e:raise HTTPException(502,'AI қызметі жауап бермеді: '+str(e)[:180])

app.mount('/',StaticFiles(directory=str(ROOT/'static'),html=True),name='static')
