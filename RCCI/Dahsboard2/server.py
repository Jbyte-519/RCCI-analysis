"""Local API for Dahsboard2. It calls the existing RUN_1 analyzer unchanged."""
from __future__ import annotations
import email, email.policy, importlib.util, json, math, os, sys, time
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import pandas as pd

# Report plots are rendered by API worker threads. Force a headless backend so
# Matplotlib never opens a Windows GUI or crashes with a breakpoint exception.
os.environ.setdefault('MPLBACKEND', 'Agg')
import matplotlib
matplotlib.use('Agg', force=True)

ROOT=Path(__file__).resolve().parent
RUN_DIR=ROOT.parent/'RUN_1'
# A local engine copy makes this dashboard portable when the workspace uses
# numbered run folders (1, 2, ...), while retaining the original fallback.
ANALYZER_PATH=ROOT/'rcci_engine.py' if (ROOT/'rcci_engine.py').exists() else RUN_DIR/'rcci_gui.py'
RESULTS=ROOT/'results'; UPLOADS=ROOT/'uploads'
spec=importlib.util.spec_from_file_location('rcci_dahsboard2_analyzer',ANALYZER_PATH)
analyzer=importlib.util.module_from_spec(spec); sys.modules[spec.name]=analyzer; spec.loader.exec_module(analyzer)

def clean(obj):
    if isinstance(obj,dict): return {str(k):clean(v) for k,v in obj.items()}
    if isinstance(obj,(list,tuple)): return [clean(v) for v in obj]
    if hasattr(obj,'item'):
        try: obj=obj.item()
        except Exception: pass
    if isinstance(obj,float) and not math.isfinite(obj): return None
    return obj if isinstance(obj,(str,int,float,bool)) or obj is None else str(obj)

class Handler(BaseHTTPRequestHandler):
    def log_message(self,fmt,*args): print('[dahsboard2]',fmt%args)
    def send(self,code,payload,content='application/json; charset=utf-8'):
        self.send_response(code); self.send_header('Content-Type',content); self.send_header('Access-Control-Allow-Origin','*'); self.send_header('Access-Control-Allow-Headers','Content-Type'); self.send_header('Access-Control-Allow-Methods','GET,POST,OPTIONS'); self.send_header('Content-Length',str(len(payload))); self.end_headers(); self.wfile.write(payload)
    def do_OPTIONS(self): self.send(204,b'')
    def do_GET(self):
        if self.path=='/api/health': return self.send(200,b'{"ok":true}')
        if self.path.startswith('/artifacts/'):
            rel=self.path[len('/artifacts/'):].split('?')[0]; target=(RESULTS/rel).resolve()
            if not target.is_relative_to(RESULTS.resolve()) or not target.is_file(): return self.send(404,b'Not found','text/plain')
            typ={'.html':'text/html; charset=utf-8','.png':'image/png','.xlsx':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','.csv':'text/csv; charset=utf-8'}.get(target.suffix,'application/octet-stream')
            return self.send(200,target.read_bytes(),typ)
        return self.send(404,b'{"error":"Not found"}')
    def do_POST(self):
        try:
            if self.path!='/api/analyze': return self.send(404,b'{"error":"Not found"}')
            size=int(self.headers.get('Content-Length','0'))
            if size<=0 or size>120*1024*1024: raise ValueError('Upload must be below 120 MB.')
            body=self.rfile.read(size); message=BytesParser(policy=email.policy.default).parsebytes(b'Content-Type: '+self.headers['Content-Type'].encode()+b'\r\nMIME-Version: 1.0\r\n\r\n'+body)
            if not message.is_multipart(): raise ValueError('Expected multipart form data.')
            fields={}; files={}
            for part in message.iter_parts():
                name=part.get_param('name',header='content-disposition'); filename=part.get_filename(); data=part.get_payload(decode=True) or b''
                if filename: files[name]=(Path(filename).name,data)
                elif name: fields[name]=data.decode('utf-8','replace')
            if 'trace' not in files: raise ValueError('Select the instrument workbook first.')
            stamp=time.strftime('%Y%m%d_%H%M%S')+'_'+str(time.time_ns())[-6:]; work=UPLOADS/stamp; out=RESULTS/stamp; work.mkdir(parents=True,exist_ok=True); out.mkdir(parents=True,exist_ok=True)
            def save(k):
                if k not in files or not files[k][1]: return None
                p=work/files[k][0]; p.write_bytes(files[k][1]); return p
            trace_path,cycle_path,perf_path=save('trace'),save('cycles'),save('performance')
            keys=['rpm','torque','fuel_kgh','air_kgh','lhv_mjkg','gamma','displacement_cc','compression_ratio','bore_mm','stroke_mm','e20_energy_share','smooth','egt','nox','co','hc','co2','smoke']
            vals={k:fields.get(k,'') for k in keys}
            performance={}
            if perf_path:
                performance=analyzer.load_performance_report(str(perf_path),context_paths=[str(x) for x in (cycle_path,trace_path) if x],rpm_hint=vals.get('rpm',''),torque_hint=vals.get('torque',''))
                for k in ('rpm','torque','fuel_kgh'):
                    if math.isfinite(float(performance.get(k,float('nan')))): vals[k]=str(performance[k])
            trace=analyzer.load_instrument(trace_path)
            metrics=analyzer.analyze(trace,cycle_path,vals,out,performance)
            pd.DataFrame([metrics]).to_csv(out/'summary.csv',index=False)
            plot={'angle':trace.angle.to_numpy(float).tolist(),'pressure':trace.pressure_bar.to_numpy(float).tolist(),'mfb':(trace.mfb.to_numpy(float)/(100 if 'mfb' in trace and trace.mfb.max()>1.5 else 1)).tolist() if 'mfb' in trace else []}
            analysis=pd.read_csv(out/'crank_angle_analysis.csv')
            for key in ('calculated_RPR_bar_per_deg','calculated_HRR_J_per_deg','calculated_CHRR_J'):
                plot[key]=analysis[key].where(pd.notna(analysis[key]),None).tolist() if key in analysis else []
            plot['rpr']=plot['calculated_RPR_bar_per_deg']; plot['hrr']=plot['calculated_HRR_J_per_deg']; plot['chrr']=plot['calculated_CHRR_J']
            result={'metrics':clean(metrics),'plot':clean(plot),'inputs':vals,'run_id':stamp,'artifacts':{k:f'/artifacts/{stamp}/{name}' for k,name in [('report','report.html'),('workbook','analysis.xlsx'),('summary','summary.csv'),('trace','crank_angle_analysis.csv')]}}
            self.send(200,json.dumps(result,allow_nan=False).encode())
        except Exception as exc:
            self.send(400,json.dumps({'error':str(exc)}).encode())

if __name__=='__main__':
    RESULTS.mkdir(exist_ok=True); UPLOADS.mkdir(exist_ok=True); print('Dahsboard2 API: http://127.0.0.1:8766',flush=True); ThreadingHTTPServer(('127.0.0.1',8766),Handler).serve_forever()
