"""Local PhilRice Standard Work Time forecasting UI."""
from email.parser import BytesParser
from email.policy import default
import json, sys, tempfile, threading, uuid, webbrowser, traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
import weekly_engine as weekly
import monthly_engine as monthly
from weekly_engine import display_model
FORECAST_LOCK=threading.Lock()
TEMPLATES={"weekly":"Percentage_Standard_Work_Time_Input_Template.xlsx", "monthly":"Workload_by_Specific_Tag_Input_Template.xlsx"}

BASE=Path(getattr(sys,'_MEIPASS',Path(__file__).resolve().parent))
HTML=(BASE/'templates'/'index.html').read_bytes()
JOBS={}; LOCK=threading.Lock()

def worker(job_id,file_path,forecast_months,output,mode='weekly'):
    def report(done,total,detail):
        with LOCK: JOBS[job_id].update(done=done,total=total,detail=detail,status='running')
    try:
        engine=weekly if mode=='weekly' else monthly
        if mode not in TEMPLATES: raise ValueError('Choose a supported forecast type.')
        loaded=engine.read_inputs(file_path)
        datasets,warnings_list=loaded if mode=='weekly' else (loaded,[])
        warnings_list += [f'{e}: short monthly history; results are exploratory.' for e,v in datasets.items() if mode=='monthly' and len(v)<12]
        # Serialize complete jobs because matplotlib is not thread-safe.
        with FORECAST_LOCK:
            results=engine.run_all(datasets,forecast_months,report)
            with LOCK: JOBS[job_id].update(detail='Creating Excel results and charts...')
            if mode=='weekly': engine.write_workbook(results,warnings_list,output)
            else: engine.write_workbook(results,output)
        summary=[]
        for e,v in results.items():
            if mode=='weekly':
                f=v['fit']; summary.append({'employee':e,'model':display_model(f['model'],f['order'],f['extra']),'rmse':f['selected_rmse'],'points':len(v['dates'])})
            else: summary.append({'employee':e,'model':str(len(v['tags']))+' tag models','points':len(v['dates'])})
        with LOCK: JOBS[job_id].update(status='done',detail='Forecast completed',done=JOBS[job_id].get('total',1),summary=summary,warnings=warnings_list)
    except Exception as exc:
        with LOCK: JOBS[job_id].update(status='error',detail=str(exc))
        traceback.print_exc()

class Handler(BaseHTTPRequestHandler):
    def response(self,code,data,typ):
        self.send_response(code); self.send_header('Content-Type',typ); self.send_header('Cache-Control','no-store'); self.send_header('Content-Length',str(len(data))); self.end_headers(); self.wfile.write(data)
    def do_GET(self):
        p=urlparse(self.path).path
        if p=='/': return self.response(200,HTML,'text/html; charset=utf-8')
        if p.startswith('/template/'):
            mode=p.split('/')[-1]
            if mode not in TEMPLATES: return self.response(404,b'Not found','text/plain')
            name=TEMPLATES[mode]; blob=(BASE/'input_templates'/name).read_bytes()
            self.send_response(200); self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); self.send_header('Content-Disposition',f'attachment; filename="{name}"'); self.send_header('Content-Length',str(len(blob))); self.end_headers(); self.wfile.write(blob); return
        if p.startswith('/api/progress/'):
            jid=p.split('/')[-1]
            with LOCK: job=JOBS.get(jid)
            if not job:return self.response(404,b'{}','application/json')
            return self.response(200,json.dumps({k:v for k,v in job.items() if k!='file'}).encode(),'application/json')
        if p.startswith('/download/'):
            jid=p.split('/')[-1]
            with LOCK: job=JOBS.get(jid)
            if not job or job['status']!='done': return self.response(404,b'Not ready','text/plain')
            blob=Path(job['file']).read_bytes(); self.send_response(200); self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); self.send_header('Content-Disposition','attachment; filename="PhilRice_Forecast_Results.xlsx"'); self.send_header('Content-Length',str(len(blob))); self.end_headers(); self.wfile.write(blob); return
        return self.response(404,b'Not found','text/plain')
    def do_POST(self):
        if self.path!='/api/start': return self.response(404,b'Not found','text/plain')
        try:
            size=int(self.headers.get('Content-Length','0'))
            if size<1 or size>25*1024*1024: raise ValueError('Please upload a file smaller than 25 MB.')
            data=self.rfile.read(size)
            multipart=BytesParser(policy=default).parsebytes(b'Content-Type: '+self.headers['Content-Type'].encode()+b'\r\nMIME-Version: 1.0\r\n\r\n'+data)
            form={part.get_param('name',header='content-disposition'):part for part in multipart.iter_parts()}
            mode=form['mode'].get_content().strip() if 'mode' in form else ''
            if mode not in TEMPLATES: raise ValueError('Choose a forecast type.')
            forecast_months=int(form['months'].get_content()) if 'months' in form else 12
            if not 1<=forecast_months<=24: raise ValueError('Months to forecast must be between 1 and 24.')
            file=form.get('file'); filename=file.get_filename() if file else ''
            if not filename or not filename.lower().endswith(('.xlsx','.xlsm')): raise ValueError('Please select an Excel workbook (.xlsx or .xlsm).')
            jid=uuid.uuid4().hex; output=Path(tempfile.gettempdir())/f'philrice_swt_result_{jid}.xlsx'; input_file=Path(tempfile.gettempdir())/f'philrice_swt_input_{jid}.xlsx'; input_file.write_bytes(file.get_payload(decode=True))
            with LOCK:JOBS[jid]={'status':'queued','done':0,'total':0,'detail':'Validating input data...','file':str(output),'summary':[],'warnings':[]}
            threading.Thread(target=worker,args=(jid,input_file,forecast_months,output,mode),daemon=True).start()
            return self.response(200,json.dumps({'job':jid}).encode(),'application/json')
        except Exception as ex:return self.response(400,json.dumps({'error':str(ex)}).encode(),'application/json')
    def log_message(self,fmt,*args):pass

def start_app():
    srv=ThreadingHTTPServer(('127.0.0.1',0),Handler); port=srv.server_address[1]; url=f'http://127.0.0.1:{port}'; threading.Timer(.8,lambda:webbrowser.open(url)).start(); srv.serve_forever()

if __name__=='__main__':
    try:start_app()
    except Exception:
        error=traceback.format_exc(); logfile=Path(tempfile.gettempdir())/'PhilRice_Standard_Work_Time_Forecast_Error.log'; logfile.write_text(error,encoding='utf-8')
        if sys.platform.startswith('win'):
            import ctypes; ctypes.windll.user32.MessageBoxW(None,f'Unable to start PhilRice Standard Work Time Forecast.\nLog: {logfile}','PhilRice Forecast',0x10)
        raise
