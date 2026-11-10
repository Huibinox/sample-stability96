from __future__ import annotations
import contextlib
import hashlib
import json
import os
import queue
import socket
import threading
import time
import unicodedata
import uuid
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook

STAMP = '%Y-%m-%d %H:%M:%S'
HEADERS = {
    'bas': ['样品编号', '样本编号', 'BAS ID', 'BAS号'],
    'project': ['BAS课题号', '课题号', '项目号', '项目编号'],
    'collected': ['Collection Date and Time', '采样时间', '采血时间'],
    'date': ['Collection Date', '采样日期', '采血日期'],
    'time': ['Collection Time', '采样时分', '采血时分'],
}

def norm(value):
    normalized = unicodedata.normalize('NFKC', str(value or '')).casefold()
    return ''.join(c for c in normalized if not c.isspace() and unicodedata.category(c) != 'Cf')

def text(value):
    if value is None: return ''
    if isinstance(value, float) and value.is_integer(): return str(int(value))
    return str(value).strip()

def atomic_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + '.' + uuid.uuid4().hex + '.tmp')
    try:
        with temp.open('x', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.flush(); os.fsync(f.fileno())
        os.replace(temp, path)
    finally:
        if temp.exists(): temp.unlink()

def parse_date(value):
    if isinstance(value, datetime): return value.date()
    if hasattr(value, 'year') and hasattr(value, 'day'): return value
    try: return datetime.fromisoformat(text(value).replace('/', '-')).date()
    except ValueError: return None

def parse_time(value):
    if isinstance(value, datetime): return value.time().replace(microsecond=0)
    if hasattr(value, 'hour'): return value.replace(microsecond=0)
    if isinstance(value, (float, int)) and 0 <= value < 1:
        return (datetime(2000, 1, 1) + timedelta(seconds=round(value * 86400))).time()
    for fmt in ('%H:%M:%S', '%H:%M'):
        try: return datetime.strptime(text(value), fmt).time()
        except ValueError: pass
    return None

def collection(combined, day, clock):
    full, warning = None, ''
    if isinstance(combined, datetime): full = combined.replace(microsecond=0)
    elif ':' in text(combined):
        try: full = datetime.fromisoformat(text(combined).replace('/', '-'))
        except ValueError: pass
    d, t = parse_date(day), parse_time(clock)
    separate = datetime.combine(d, t) if d and t is not None else None
    if full and full.tzinfo: full = full.replace(tzinfo=None)
    if full and separate and full != separate:
        warning = '合并采样时间与分列日期/时间不一致，请核对'
    chosen = full or separate
    date_only = d or parse_date(combined)
    if not chosen and (text(combined) or text(clock)):
        warning = warning or '采样时间缺失或格式无法识别，请手动核对'
    return (chosen.strftime(STAMP) if chosen else '',
            date_only.isoformat() if date_only else '', warning)

def read_excel(path):
    path = Path(path)
    quick_check = norm(path.stem).startswith('quickchecksampleitems')
    rows, errors = [], []
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        found = False
        for sheet in book:
            sheet.reset_dimensions()
            mapping = None
            for number, cells in enumerate(sheet.iter_rows(), 1):
                row = tuple(cell.value for cell in cells)
                if mapping is None:
                    if number > 30: break
                    candidates = {norm(v): i for i, v in enumerate(row) if v is not None}
                    trial = {key: next((candidates[norm(a)] for a in aliases if norm(a) in candidates), None)
                             for key, aliases in HEADERS.items()}
                    if quick_check and len(row) >= 40 and (
                        norm(row[0]) == 'recordid' or
                        any(trial[key] == index for key, index in
                            (('collected', 30), ('date', 31), ('time', 32)))
                    ) and (trial['bas'] is None or trial['project'] is None):
                        # Confirmed Quick Check export layout. Indices are zero-based.
                        trial.update(bas=24, project=39)
                        for key, index in (('collected', 30), ('date', 31), ('time', 32)):
                            if trial[key] is None:
                                trial[key] = index
                    if trial['bas'] is not None and trial['project'] is not None:
                        mapping = trial; found = True
                    continue
                def val(key):
                    index = mapping[key]
                    return row[index] if index is not None and index < len(row) else None
                if not any(v is not None for v in row): continue
                bas, project = text(val('bas')), text(val('project'))
                if not bas or not project:
                    errors.append(f'{sheet.title} 第{number}行：缺少样品编号或课题号'); continue
                combined = val('collected')
                ci = mapping['collected']
                if isinstance(combined, datetime) and ci is not None:
                    fmt = cells[ci].number_format.lower()
                    if not any(c in fmt for c in ('h', 's')):
                        combined = combined.date()
                collected, date_only, issue = collection(combined, val('date'), val('time'))
                rows.append(dict(bas=bas, project=project, collected=collected, date=date_only,
                                 issue=issue, source=f'{path.name} / {sheet.title} / 第{number}行'))
        if not found: raise ValueError('前30行未识别到样品编号和项目号。Quick Check 导出请保留原文件名、表头及 Y/AN 列。')
    finally: book.close()
    return rows, errors

class SharedStore:
    def __init__(self, root):
        self.root = Path(root)
        self.meta = self.root / '_稳定期助手'
        self.events = self.meta / '操作记录'

    @contextlib.contextmanager
    def lock(self):
        if not self.root.is_dir(): raise OSError('共享文件夹不可访问')
        self.events.mkdir(parents=True, exist_ok=True)
        lock = self.meta / '写入锁'
        deadline = time.monotonic() + 8
        while True:
            try: lock.mkdir(); break
            except FileExistsError:
                if time.monotonic() > deadline:
                    raise OSError('共享目录忙或遗留写入锁；请稍后重试。持续出现时由管理员确认所有客户端退出后检查“写入锁”文件夹')
                time.sleep(.12)
        try: yield
        finally: lock.rmdir()

    def read_events(self):
        if not self.root.is_dir(): raise OSError('共享文件夹不可访问')
        values = []
        if self.events.exists():
            for path in self.events.glob('*.json'):
                try:
                    event = json.loads(path.read_text(encoding='utf-8'))
                    if not isinstance(event.get('seq'), int) or event.get('kind') not in ('import', 'complete', 'undo', 'edit'):
                        raise ValueError('操作记录格式不正确')
                    values.append(event)
                except Exception as exc: raise OSError(f'操作记录无法读取：{path.name}：{exc}') from exc
        return sorted(values, key=lambda e: (e['seq'], e['id']))

    def append_locked(self, events, kind, payload, actor, event_id=None):
        event_id = event_id or uuid.uuid4().hex
        if any(e['id'] == event_id for e in events): return None
        event = dict(id=event_id, seq=max((e['seq'] for e in events), default=0) + 1,
                     kind=kind, payload=payload, actor=actor, at=datetime.now().strftime(STAMP), machine=socket.gethostname())
        atomic_json(self.events / (event_id + '.json'), event)
        events.append(event)
        return event

    def import_file(self, path, actor):
        before = path.stat()
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        event_id = 'excel-' + digest
        if (self.events / (event_id + '.json')).exists(): return None
        records, errors = read_excel(path)
        after = path.stat()
        if before.st_size != after.st_size or before.st_mtime_ns != after.st_mtime_ns:
            raise OSError('Excel正在变化，等待下次扫描')
        if not records: raise ValueError('未读到有效样本；' + '；'.join(errors[:3]))
        with self.lock():
            events = self.read_events()
            self.append_locked(events, 'import', dict(rows=records, errors=errors, file=path.name), actor, event_id)
        return len(records), errors

    def action(self, kind, bas, actor, expected, value=None):
        with self.lock():
            events = self.read_events()
            state, _ = replay(events)
            sample = state.get(bas)
            if sample is None: raise ValueError('样本已不存在，请刷新')
            revision = sample.get('edit_rev', '') if kind == 'edit' else sample.get('status_rev', '')
            if revision != expected: raise ValueError('其他同事刚修改了此样本，请刷新后核对再操作')
            if kind == 'complete' and sample['done']: raise ValueError('此样本已经完成')
            if kind == 'undo' and not sample['done']: raise ValueError('此样本已经恢复为未完成')
            self.append_locked(events, kind, dict(bas=bas, value=value), actor)

def replay(events):
    samples, reports = {}, []
    for event in events:
        kind, p = event['kind'], event['payload']
        if kind == 'import':
            reports.extend(f"{p['file']}：{s}" for s in p.get('errors', []))
            for row in p['rows']:
                bas = row['bas']
                if bas not in samples:
                    samples[bas] = dict(row, done=False, status_rev='', edit_rev='', conflicts=[], overrides=[], completed_by='', completed_at='')
                else:
                    s = samples[bas]
                    for field in ('project', 'collected', 'date'):
                        new, old = row.get(field, ''), s.get(field, '')
                        if field in s['overrides']: continue
                        if new and old and new != old:
                            conflict = f'{field}：{old} / {new}（{row["source"]}）'
                            if conflict not in s['conflicts']: s['conflicts'].append(conflict)
                        elif new and not old: s[field] = new
                    if row.get('issue') and 'collected' not in s['overrides']: s['issue'] = row['issue']
        elif p.get('bas') in samples:
            s = samples[p['bas']]
            if kind in ('complete', 'undo'):
                s['done'] = kind == 'complete'; s['status_rev'] = event['id']
                s['completed_by'] = event['actor'] if s['done'] else ''
                s['completed_at'] = event['at'] if s['done'] else ''
            elif kind == 'edit':
                for field in ('collected', 'project'):
                    if field in p['value']:
                        s[field] = p['value'][field]
                        if field not in s['overrides']: s['overrides'].append(field)
                s['date'] = s['collected'][:10]
                if 'date' not in s['overrides']: s['overrides'].append('date')
                s['issue'] = ''; s['conflicts'] = []; s['edit_rev'] = event['id']
                s['edited_by'], s['edited_at'] = event['actor'], event['at']
    return samples, reports

def deadline(sample):
    try: return datetime.strptime(sample.get('collected', ''), STAMP) + timedelta(hours=96)
    except ValueError: return None

def category(sample, now):
    if sample['done']: return 'done'
    if sample.get('issue') or sample.get('conflicts') or not deadline(sample): return 'missing'
    return 'expired' if deadline(sample) <= now else 'active'

def remaining(seconds):
    h, rem = divmod(max(0, int(seconds)), 3600)
    m, s = divmod(rem, 60)
    return f'{h:02d}:{m:02d}:{s:02d}'

def color(seconds):
    stops = [(0, (172,38,56)), (6, (207,54,64)), (12, (193,102,29)), (24, (151,123,22)), (96, (27,128,102))]
    h = max(0, min(96, seconds / 3600))
    for (lo,a), (hi,b) in zip(stops, stops[1:]):
        if lo <= h <= hi:
            return '#%02x%02x%02x' % tuple(round(x+(y-x)*(h-lo)/(hi-lo)) for x,y in zip(a,b))
    return '#1b8066'

class App:
    def __init__(self, root, data_dir=None):
        self.root = root
        self.data_dir = Path(data_dir or (Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'SampleStability96'))
        self.config_path = self.data_dir / 'config.json'
        try: self.config = json.loads(self.config_path.read_text(encoding='utf-8'))
        except (OSError, ValueError): self.config = {}
        self.online = False; self.last_sync = ''; self.busy = False; self.mini = None
        self.last_signature = None; self.scan_notes = []; self.messages = queue.Queue(); self.jobs = queue.Queue()
        self.stop = threading.Event(); self.generation = 0
        self.shared = self.config.get('shared', ''); self.actor = self.config.get('actor', '')
        self.load_cache()
        root.title('96h · 样本稳定期助手')
        root.geometry(f'{min(1180, root.winfo_screenwidth()-70)}x{min(680, root.winfo_screenheight()-140)}+30+30')
        root.minsize(900,580)
        root.configure(bg='#f3f6f8')
        style = ttk.Style(); style.theme_use('clam')
        style.configure('.', font=('Microsoft YaHei UI',10))
        style.configure('TFrame', background='#f3f6f8')
        style.configure('TLabel', background='#f3f6f8', foreground='#263f4b')
        style.configure('TNotebook', background='#f3f6f8', borderwidth=0)
        style.configure('Treeview', rowheight=34, background='white', fieldbackground='white', borderwidth=0)
        style.configure('Treeview.Heading', font=('Microsoft YaHei UI',10,'bold'), background='#e9eff3', padding=8)
        style.map('Treeview', background=[('selected','#dbe9f4')],foreground=[('selected','#162f43')])
        style.configure('TButton',padding=(12,7),background='#e5edf2',foreground='#183f52')
        style.configure('TNotebook.Tab',padding=(22,10),background='#e5edf2')
        header=tk.Frame(root,bg='#142f3b',padx=24,pady=18); header.pack(fill='x')
        tk.Label(header,text='96h',fg='#73dcc1',bg='#142f3b',font=('Segoe UI',29,'bold')).pack(side='left')
        tk.Label(header,text='样本稳定期助手',fg='white',bg='#142f3b',font=('Microsoft YaHei UI',17,'bold')).pack(side='left',padx=18)
        ttk.Button(header,text='桌面小窗',command=self.show_mini).pack(side='right')
        ttk.Button(header,text='设置 / 共享文件夹',command=self.settings).pack(side='right',padx=8)
        self.stats=tk.StringVar()
        tk.Label(root,textvariable=self.stats,bg='#f3f6f8',fg='#142f3b',anchor='w',font=('Microsoft YaHei UI',14,'bold'),padx=24,pady=18).pack(fill='x')
        bar=ttk.Frame(root,padding=(20,0,20,10)); bar.pack(fill='x')
        ttk.Label(bar,text='搜索 BAS / 项目：').pack(side='left')
        self.search=tk.StringVar(); self.search.trace_add('write',lambda *_:self.refresh_view(True))
        ttk.Entry(bar,textvariable=self.search,width=24).pack(side='left')
        ttk.Button(bar,text='立即扫描',command=self.scan_now).pack(side='left',padx=10)
        ttk.Button(bar,text='导入 Excel…',command=self.import_dialog).pack(side='left')
        ttk.Button(bar,text='导入 / 核对提示',command=self.show_issues).pack(side='right')
        self.notebook=ttk.Notebook(root); self.notebook.pack(fill='both',expand=True,padx=20)
        self.home=ttk.Frame(self.notebook); forecast=ttk.Frame(self.notebook); completed=ttk.Frame(self.notebook)
        self.notebook.add(self.home,text='待检测'); self.notebook.add(forecast,text='样本预报 · 待做'); self.notebook.add(completed,text='已完成')
        hb=ttk.Frame(self.home,padding=10); hb.pack(fill='x')
        ttk.Label(hb,text='按项目分组 · 临期优先 · 双击样本核对 / 补时间').pack(side='left')
        ttk.Button(hb,text='标记已完成上机',command=lambda:self.action('complete')).pack(side='right')
        self.tree=self.make_tree(self.home,('collected','deadline','remaining'),('采样时间','截止时间（96h）','剩余时间'),(190,190,140))
        self.tree.bind('<Double-1>',self.edit_selected)
        self.drop = tk.PhotoImage(width=13, height=18)
        for y in range(2,16):
            radius = (y-1)*.48 if y<10 else max(0, 5-abs(y-11)*.9)
            for x in range(13):
                if abs(x-6)<=radius: self.drop.put('#bd3045',(x,y))
        self.open_groups=set(); self.known_groups=set()
        ttk.Label(self.home,text='缺时间 / 有冲突的样本计入总数，单独列示；过期自动移出。',padding=10).pack(side='bottom',fill='x',before=self.tree.master)
        tk.Label(forecast,text='样本预报\n待做',font=('Microsoft YaHei UI',24,'bold'),fg='#81919e',pady=100).pack()
        ttk.Label(forecast,text='当前版本不生成预报；待项目采样规则明确后再启用。').pack()
        db=ttk.Frame(completed,padding=10); db.pack(fill='x')
        ttk.Label(db,text='完成代表已完成上机检测；重复导入不会恢复为待检测。').pack(side='left')
        ttk.Button(db,text='撤销完成',command=lambda:self.action('undo')).pack(side='right')
        self.done_tree=self.make_tree(completed,('project','collected','by','at'),('项目号','采样时间','完成人','完成时间'),(180,180,130,180))
        self.status=tk.StringVar(value='尚未读取共享数据')
        tk.Label(root,textvariable=self.status,anchor='w',bg='#e6edf1',fg='#334d5b',padx=20,pady=10).pack(side='bottom',fill='x',before=self.notebook)
        root.protocol('WM_DELETE_WINDOW',self.close)
        threading.Thread(target=self.worker,daemon=True).start()
        self.tick(); self.poll()
        if not self.shared or not self.actor: root.after(400,self.settings)

    def load_cache(self):
        key=hashlib.sha256(self.shared.lower().encode()).hexdigest()[:16]
        self.cache_path=self.data_dir/('cache-'+key+'.json')
        try:
            data=json.loads(self.cache_path.read_text(encoding='utf-8'))
            self.samples,self.reports,self.last_sync=data['samples'],data.get('reports',[]),data.get('at','')
        except (OSError,ValueError,KeyError): self.samples,self.reports={},[]

    def make_tree(self,parent,columns,labels,widths):
        frame=ttk.Frame(parent); frame.pack(fill='both',expand=True)
        tree=ttk.Treeview(frame,columns=columns,show='tree headings',selectmode='extended')
        tree.heading('#0',text='项目 / BAS 号'); tree.column('#0',width=270,minwidth=200)
        for c,label,width in zip(columns,labels,widths):
            tree.heading(c,text=label); tree.column(c,width=width,minwidth=100)
        scroll=ttk.Scrollbar(frame,orient='vertical',command=tree.yview); tree.configure(yscrollcommand=scroll.set)
        tree.pack(side='left',fill='both',expand=True); scroll.pack(side='right',fill='y')
        tree.tag_configure('missing',foreground='#ac6122',background='#fff7e8')
        tree.tag_configure('group',font=('Microsoft YaHei UI',10,'bold'),background='#edf3f5')
        return tree

    def settings(self):
        win=tk.Toplevel(self.root); win.title('设置 / 共享文件夹'); win.geometry('650x280'); win.transient(self.root); win.grab_set()
        frame=ttk.Frame(win,padding=22); frame.pack(fill='both',expand=True)
        ttk.Label(frame,text='姓名 / 工号（请使用能区分同事的名称）').pack(anchor='w')
        actor=tk.StringVar(value=self.actor); ttk.Entry(frame,textvariable=actor).pack(fill='x',pady=(4,12))
        ttk.Label(frame,text='共享文件夹（所有同事选择同一个网络共享位置）').pack(anchor='w')
        folder=tk.StringVar(value=self.shared); line=ttk.Frame(frame); line.pack(fill='x',pady=5)
        ttk.Entry(line,textvariable=folder).pack(side='left',fill='x',expand=True)
        def browse():
            path=filedialog.askdirectory(parent=win,title='选择共享文件夹',mustexist=True)
            if path: folder.set(path)
        ttk.Button(line,text='浏览…',command=browse).pack(side='right',padx=(8,0))
        ttk.Label(frame,text='扫描根目录及“待导入”子文件夹的 .xlsx / .xlsm，原件不移动。').pack(anchor='w',pady=8)
        def save():
            if not actor.get().strip() or not folder.get().strip():
                messagebox.showerror('设置未完成','请填写姓名并选择共享文件夹',parent=win); return
            self.actor,self.shared=actor.get().strip(),folder.get().strip()
            self.generation+=1; self.online=False; self.last_sync=''; self.last_signature=None; self.busy=False
            self.load_cache()
            try: atomic_json(self.config_path,dict(shared=self.shared,actor=self.actor))
            except OSError as exc: messagebox.showerror('无法保存设置',str(exc),parent=win); return
            win.destroy(); self.refresh_view(True); self.scan_now()
        ttk.Button(frame,text='保存并连接',command=save).pack(anchor='e')

    def scan_now(self):
        if self.shared and not self.busy:
            self.busy=True; self.status.set('正在读取共享文件夹…')
            self.jobs.put(('scan',self.generation,self.shared,self.actor,None))

    def import_dialog(self):
        if not self.shared: self.settings(); return
        path=filedialog.askopenfilename(title='导入 Excel（复制到共享待导入目录）',filetypes=[('Excel 文件','*.xlsx *.xlsm')])
        if path: self.jobs.put(('copy',self.generation,self.shared,self.actor,path))

    def worker(self):
        import shutil
        seen_files = {}
        while not self.stop.is_set():
            try: kind,gen,shared,actor,payload=self.jobs.get(timeout=.5)
            except queue.Empty: continue
            try:
                store=SharedStore(shared)
                if kind=='copy':
                    if not store.root.is_dir(): raise OSError('共享文件夹不可访问')
                    source=Path(payload); target_dir=Path(shared)/'待导入'; target_dir.mkdir(parents=True,exist_ok=True)
                    digest=hashlib.sha256(source.read_bytes()).hexdigest()[:12]
                    target=target_dir/(source.stem+'-'+digest+source.suffix.lower())
                    if source.resolve()!=target.resolve() and not target.exists():
                        temp=target.with_suffix('.tmp-'+uuid.uuid4().hex)
                        try: shutil.copyfile(source,temp); os.replace(temp,target)
                        finally:
                            if temp.exists(): temp.unlink()
                elif kind=='action':
                    action,entries=payload
                    for bas,expected,value in entries: store.action(action,bas,actor,expected,value)
                notes=[]
                if kind in ('scan','copy'):
                    if not store.root.is_dir(): raise OSError('共享文件夹不可访问')
                    files=[]
                    for directory in (store.root,store.root/'待导入'):
                        if directory.exists():
                            files.extend(p for p in directory.iterdir() if p.is_file() and p.suffix.lower() in ('.xlsx','.xlsm') and not p.name.startswith('~$'))
                    for path in sorted(files):
                        try:
                            stat = path.stat()
                            if time.time()-stat.st_mtime<3: continue
                            key = (shared, str(path))
                            signature = (stat.st_size, stat.st_mtime_ns)
                            if seen_files.get(key) == signature: continue
                            result=store.import_file(path,actor)
                            seen_files[key] = signature
                            if result: notes.append(f'{path.name}：导入 {result[0]} 行')
                        except Exception as exc: notes.append(f'{path.name}：{exc}')
                samples,reports=replay(store.read_events())
                self.messages.put(('ok',gen,(samples,reports,datetime.now().strftime(STAMP),notes)))
            except Exception as exc: self.messages.put(('error',gen,str(exc)))
            finally: self.jobs.task_done()

    def poll(self):
        try:
            while True:
                kind,gen,data=self.messages.get_nowait()
                if gen!=self.generation: continue
                self.busy=False
                if kind=='ok':
                    self.samples,self.reports,self.last_sync,self.scan_notes=data; self.online=True
                    try: atomic_json(self.cache_path,dict(samples=self.samples,reports=self.reports,at=self.last_sync))
                    except OSError as exc: self.scan_notes.append('本地缓存保存失败：'+str(exc))
                    self.refresh_view(True)
                else:
                    self.online=False; self.scan_notes=[data]
                    self.status.set('未同步 / 操作未确认：'+data+' · 显示缓存，请刷新核对')
        except queue.Empty: pass
        self.root.after(150,self.poll)

    def refresh_view(self,force=False):
        now=datetime.now(); query=self.search.get().strip().lower()
        active=[s for s in self.samples.values() if category(s,now)=='active']
        missing=[s for s in self.samples.values() if category(s,now)=='missing']
        expired=[s for s in self.samples.values() if category(s,now)=='expired']
        near24=sum(0<(deadline(s)-now).total_seconds()<=86400 for s in active)
        near6=sum(0<(deadline(s)-now).total_seconds()<=21600 for s in active)
        self.stats.set(f'待检测 {len(active)+len(missing)}   ·   24h 内到期 {near24}   ·   6h 内到期 {near6}   ·   时间待补 / 核对 {len(missing)}')
        visible=lambda s:not query or query in (s['bas']+' '+s['project']).lower()
        groups={}
        for s in sorted(active,key=deadline):
            if visible(s): groups.setdefault(s['project'],[]).append(s)
        ordered=sorted(groups,key=lambda p:deadline(groups[p][0]))
        signature=(tuple((p,tuple(s['bas'] for s in groups[p])) for p in ordered),tuple(s['bas'] for s in missing if visible(s)),tuple(sorted(s['bas'] for s in self.samples.values() if s['done'] and visible(s))))
        if force or signature!=self.last_signature:
            selection=self.tree.selection(); done_selection=self.done_tree.selection()
            scroll=self.tree.yview(); done_scroll=self.done_tree.yview()
            self.open_groups={i for i in self.tree.get_children('') if self.tree.item(i,'open')}
            self.tree.delete(*self.tree.get_children()); self.done_tree.delete(*self.done_tree.get_children())
            for p in ordered:
                gid='g:'+p; opened=gid in self.open_groups or gid not in self.known_groups
                self.known_groups.add(gid)
                self.tree.insert('','end',iid=gid,text=p,open=opened,tags=('group',))
                for s in groups[p]: self.tree.insert(gid,'end',iid='s:'+s['bas'],text=s['bas'])
            pending=[s for s in missing if visible(s)]
            if pending:
                gid='missing'
                self.tree.insert('','end',iid=gid,text=f'时间待补 / 待核对 · {len(pending)} 个',open=gid in self.open_groups or gid not in self.known_groups,tags=('missing',))
                self.known_groups.add(gid)
                for s in pending: self.tree.insert(gid,'end',iid='s:'+s['bas'],text=s['bas']+' · '+s['project'],values=(s.get('collected') or s.get('date') or '待补','需手动核对','—'),tags=('missing',))
            for s in sorted(self.samples.values(),key=lambda s:s.get('completed_at',''),reverse=True):
                if s['done'] and visible(s): self.done_tree.insert('','end',iid='s:'+s['bas'],text=s['bas'],values=(s['project'],s.get('collected',''),s['completed_by'],s['completed_at']))
            self.tree.selection_set([i for i in selection if self.tree.exists(i)])
            self.done_tree.selection_set([i for i in done_selection if self.done_tree.exists(i)])
            if scroll:self.tree.yview_moveto(scroll[0])
            if done_scroll:self.done_tree.yview_moveto(done_scroll[0])
            self.last_signature=signature
        for p in ordered:
            nearest=(deadline(groups[p][0])-now).total_seconds()
            self.tree.item('g:'+p,text=f'{p} · {len(groups[p])} 个',values=('','',remaining(nearest)))
            for s in groups[p]:
                seconds=(deadline(s)-now).total_seconds(); tag='c'+str(int(seconds//300))
                self.tree.tag_configure(tag,foreground=color(seconds),background='#fff0f0' if seconds<=21600 else '#ffffff')
                self.tree.item('s:'+s['bas'],values=(s['collected'],deadline(s).strftime(STAMP),remaining(seconds)),tags=(tag,),image=self.drop if seconds<=21600 else '')
        if self.online:
            self.status.set(f'共享目录读取于 {self.last_sync} · 已过期未完成 {len(expired)} 个（已移出） · 导入/扫描提示 {len(self.scan_notes)+len(self.reports)} 条')
        self.update_mini(active,missing,now)

    def action(self,kind):
        if not self.online: messagebox.showwarning('未连接','请先成功读取共享文件夹，再提交操作。'); return
        tree=self.done_tree if kind=='undo' else self.tree
        entries=[(i[2:],self.samples[i[2:]].get('status_rev',''),None) for i in tree.selection() if i.startswith('s:')]
        if not entries: messagebox.showinfo('请选择样本','选择样本行，可按 Ctrl / Shift 多选。'); return
        self.jobs.put(('action',self.generation,self.shared,self.actor,(kind,entries)))
        self.status.set('正在提交操作，等待共享文件夹确认…')

    def edit_selected(self,event=None):
        selection=self.tree.selection()
        if len(selection)!=1 or not selection[0].startswith('s:'): return
        if not self.online: messagebox.showwarning('未连接','离线时不能提交修改。'); return
        s=self.samples[selection[0][2:]]
        win=tk.Toplevel(self.root); win.title('核对 / 补采样时间 · '+s['bas']); win.geometry('650x400'); win.transient(self.root); win.grab_set()
        frame=ttk.Frame(win,padding=20); frame.pack(fill='both',expand=True)
        ttk.Label(frame,text='BAS：'+s['bas'],font=('Microsoft YaHei UI',12,'bold')).pack(anchor='w')
        ttk.Label(frame,text='项目号').pack(anchor='w',pady=(12,2)); project=tk.StringVar(value=s['project']); ttk.Entry(frame,textvariable=project).pack(fill='x')
        ttk.Label(frame,text='采样时间：YYYY-MM-DD HH:MM[:SS]（单位当地时间）').pack(anchor='w',pady=(12,2))
        dt=tk.StringVar(value=s.get('collected') or (s.get('date','')+' ')); ttk.Entry(frame,textvariable=dt).pack(fill='x')
        box=tk.Text(frame,height=5,wrap='word',font=('Microsoft YaHei UI',9)); box.pack(fill='both',expand=True,pady=10)
        box.insert('1.0','来源：'+s.get('source','')+'\n'+s.get('issue','')+'\n'+'\n'.join(s.get('conflicts',[]))); box.configure(state='disabled')
        def save():
            try: parsed=datetime.fromisoformat(dt.get().strip())
            except ValueError: messagebox.showerror('格式错误','请输入完整日期和时分，例如 2026-11-08 09:30',parent=win); return
            if ':' not in dt.get() or parsed.tzinfo or not project.get().strip():
                messagebox.showerror('格式错误','请填写项目和完整当地采样时间',parent=win); return
            value=dict(project=project.get().strip(),collected=parsed.strftime(STAMP))
            self.jobs.put(('action',self.generation,self.shared,self.actor,('edit',[(s['bas'],s.get('edit_rev',''),value)])))
            win.destroy(); self.status.set('正在保存人工核对结果…')
        ttk.Button(frame,text='确认采样时间及冲突信息并保存',command=save).pack(anchor='e')

    def show_issues(self):
        win=tk.Toplevel(self.root); win.title('导入与核对提示'); win.geometry('850x500')
        box=tk.Text(win,wrap='word',font=('Microsoft YaHei UI',10)); box.pack(fill='both',expand=True,padx=12,pady=12)
        lines=list(self.scan_notes)+self.reports
        for s in self.samples.values():
            if s.get('issue') or s.get('conflicts'): lines.append(s['bas']+'：'+s.get('issue','')+'；'.join(s.get('conflicts',[])))
        box.insert('1.0','\n\n'.join(lines) or '目前没有导入或核对提示。'); box.configure(state='disabled')

    def show_mini(self):
        if self.mini and self.mini.winfo_exists(): self.mini.lift(); return
        self.mini=tk.Toplevel(self.root); self.mini.title('96h · 桌面小窗'); self.mini.geometry('410x390'); self.mini.minsize(380,340)
        self.mini.configure(bg='#142f3b'); self.mini.attributes('-topmost',True)
        self.mini_title=tk.Label(self.mini,bg='#142f3b',fg='#73dcc1',font=('Microsoft YaHei UI',15,'bold'),pady=12); self.mini_title.pack(fill='x')
        self.mini_rows= []
        rows=tk.Frame(self.mini,bg='#142f3b'); rows.pack(fill='both',expand=True)
        for _ in range(4):
            label=tk.Label(rows,bg='#142f3b',fg='white',justify='left',anchor='nw',font=('Microsoft YaHei UI',10),padx=16,pady=5)
            label.pack(fill='x'); self.mini_rows.append(label)
        self.mini_status=tk.Label(self.mini,bg='#142f3b',fg='#bacbd3',font=('Microsoft YaHei UI',9)); self.mini_status.pack(fill='x')
        buttons=ttk.Frame(self.mini); buttons.pack(fill='x',padx=10,pady=10)
        ttk.Button(buttons,text='展开主窗口',command=self.restore).pack(side='left')
        ttk.Button(buttons,text='收起主窗口',command=self.root.withdraw).pack(side='left',padx=4)
        top=tk.BooleanVar(value=True); ttk.Checkbutton(buttons,text='置顶',variable=top,command=lambda:self.mini.attributes('-topmost',top.get())).pack(side='right')
        self.mini.protocol('WM_DELETE_WINDOW',self.close_mini)
        self.refresh_view()

    def close_mini(self):
        self.restore(); self.mini.destroy(); self.mini=None

    def restore(self): self.root.deiconify(); self.root.lift()

    def update_mini(self,active,missing,now):
        if not self.mini or not self.mini.winfo_exists(): return
        ordered=sorted(active,key=deadline)[:4]
        self.mini_title.configure(text=f'待检测 {len(active)+len(missing)} · 待补/核对 {len(missing)}')
        for index,label in enumerate(self.mini_rows):
            if index<len(ordered):
                s=ordered[index]; seconds=(deadline(s)-now).total_seconds()
                tone=color(seconds)
                rgb=tuple(int(tone[i:i+2],16) for i in (1,3,5))
                light='#%02x%02x%02x'%tuple(round(c*.65+255*.35) for c in rgb)
                label.configure(text=s['project']+'  ·  '+s['bas']+'\n'+remaining(seconds),fg=light)
            else: label.configure(text='暂无可倒计时样本' if index==0 else '')
        self.mini_status.configure(text=('最近读取 '+self.last_sync[11:]) if self.online else '未同步 · 显示本地缓存')

    def tick(self):
        self.refresh_view()
        if not hasattr(self,'last_scan') or time.monotonic()-self.last_scan>=10:
            self.last_scan=time.monotonic(); self.scan_now()
        self.root.after(1000,self.tick)

    def close(self): self.stop.set(); self.root.destroy()

def main():
    root=tk.Tk(); App(root); root.mainloop()

if __name__=='__main__': main()
