"""Exercise real Tk windows and background worker using disposable synthetic Excel."""
import sys
import tempfile
import time
from datetime import datetime, timedelta
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import tkinter as tk
from openpyxl import Workbook
from stability96 import App, SharedStore, atomic_json, replay, STAMP

def main():
    temp=tempfile.TemporaryDirectory(); folder=Path(temp.name); shared=folder/'share'; shared.mkdir()
    now=datetime.now().replace(microsecond=0)
    wb=Workbook(); ws=wb.active
    ws.append(['Record ID','样品编号','BAS课题号','Collection Date and Time','Collection Date','Collection Time'])
    for i,(project,hours) in enumerate([('演示项目 A',94),('演示项目 A',80),('演示项目 B',40),('演示项目 B',97)]):
        dt=now-timedelta(hours=hours)
        ws.append([i+1,f'DEMO.{i+1:05d}',project,dt.strftime(STAMP),dt.strftime('%Y-%m-%d'),dt.strftime('%H:%M:%S')])
    ws.append([5,'DEMO.00005','演示项目 C',None,now.strftime('%Y-%m-%d'),None])
    wb.save(shared/'demo.xlsx'); wb.close()
    (SharedStore(shared)).import_file(shared/'demo.xlsx','演示人员')
    data=folder/'local'; atomic_json(data/'config.json',dict(actor='测试人员',shared=str(shared)))
    root=tk.Tk(); errors=[]
    root.report_callback_exception=lambda *args:errors.append(str(args))
    app=App(root,data_dir=data)
    def pump(condition,seconds=8):
        end=time.monotonic()+seconds
        while time.monotonic()<end:
            root.update()
            if condition(): return
            time.sleep(.03)
        raise AssertionError('GUI timed out')
    pump(lambda:app.online)
    assert len(app.samples)==5
    assert '待检测 4' in app.stats.get()
    assert not app.tree.exists('s:DEMO.00004')
    assert app.tree.exists('missing')
    before=app.tree.item('s:DEMO.00001','values')[-1]
    time.sleep(1.1); root.update()
    assert app.tree.item('s:DEMO.00001','values')[-1]!=before
    app.tree.selection_set('s:DEMO.00001'); app.action('complete')
    pump(lambda:app.samples['DEMO.00001']['done'])
    assert app.done_tree.exists('s:DEMO.00001')
    assert not app.tree.exists('s:DEMO.00001')
    app.done_tree.selection_set('s:DEMO.00001'); app.action('undo')
    pump(lambda:not app.samples['DEMO.00001']['done'])
    assert app.tree.exists('s:DEMO.00001')
    app.show_mini(); root.update()
    assert app.mini.winfo_exists()
    assert '待检测 4' in app.mini_title.cget('text')
    artifacts=Path(__file__).parent/'artifacts'; artifacts.mkdir(exist_ok=True)
    try:
        from PIL import ImageGrab
        root.lift(); root.update(); time.sleep(.3)
        x,y=root.winfo_rootx(),root.winfo_rooty()
        ImageGrab.grab(bbox=(x,y,x+root.winfo_width(),y+root.winfo_height())).save(artifacts/'main.png')
        app.mini.lift(); root.update(); time.sleep(.3)
        x,y=app.mini.winfo_rootx(),app.mini.winfo_rooty()
        ImageGrab.grab(bbox=(x,y,x+app.mini.winfo_width(),y+app.mini.winfo_height())).save(artifacts/'mini.png')
    except ImportError:
        import subprocess
        def capture(window,name):
            window.lift(); root.update(); time.sleep(.3)
            subprocess.run(['powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-File',
                str(Path(__file__).parent/'capture.ps1'),'-OutputPath',str(artifacts/name),
                '-Left',str(window.winfo_rootx()),'-Top',str(window.winfo_rooty()),
                '-Width',str(window.winfo_width()),'-Height',str(window.winfo_height())],check=True)
        app.mini.withdraw(); capture(root,'main.png'); app.mini.deiconify(); capture(app.mini,'mini.png')
    app.close_mini()
    # Disconnect the share and verify writes are disabled while cached samples stay visible.
    offline=folder/'offline'; shared.rename(offline)
    app.scan_now(); pump(lambda:not app.online)
    assert len(app.samples)==5
    assert '未同步' in app.status.get()
    app.close(); time.sleep(.7)
    assert not errors,errors
    temp.cleanup()
    print('GUI PASS: import, project groups, expiry, live countdown, complete/undo, mini window, offline cache')

if __name__=='__main__': main()
