"""Modern CustomTkinter alternative for the RCCI experimental analyzer.

This is intentionally separate from RUN_1/rcci_gui.py. It imports that module's
loader and analysis functions, so the calculation and report workflow stays
consistent with the original application.
"""
from __future__ import annotations
import importlib.util
import math
import os
import queue
import threading
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Analysis runs in a worker thread; plots are saved to files rather than
# displayed, so use a headless Matplotlib backend.
os.environ.setdefault('MPLBACKEND', 'Agg')
import matplotlib
matplotlib.use('Agg', force=True)

ROOT=Path(__file__).resolve().parent
RUN_DIR=ROOT.parent/'RUN_1'
ANALYZER_PATH=ROOT/'rcci_engine.py' if (ROOT/'rcci_engine.py').exists() else RUN_DIR/'rcci_gui.py'
spec=importlib.util.spec_from_file_location('rcci_customtkinter_engine',ANALYZER_PATH)
engine=importlib.util.module_from_spec(spec); spec.loader.exec_module(engine)

FIELDS=[('rpm','Engine speed','rpm','1529'),('torque','Brake torque','Nm','0'),('fuel_kgh','Total fuel flow','kg/h','0.16'),('air_kgh','Air flow','kg/h',''),('lhv_mjkg','Fuel blend LHV','MJ/kg',''),('gamma','Specific heat ratio γ','',''),('displacement_cc','Displacement','cc','670'),('compression_ratio','Compression ratio','','18'),('bore_mm','Bore','mm','93'),('stroke_mm','Stroke','mm','98'),('e20_energy_share','E20 energy share','%','30'),('smooth','Derivative smoothing','samples','5'),('egt','Exhaust temperature','°C',''),('nox','NOx','ppm',''),('co','CO','as recorded',''),('hc','HC','ppm',''),('co2','CO₂','%',''),('smoke','Smoke opacity','%','')]

class App(ctk.CTk):
    def __init__(self):
        super().__init__(); self.title('RCCI Pulse · CustomTkinter'); self.geometry('1240x800'); self.minsize(980,650)
        ctk.set_appearance_mode('dark'); ctk.set_default_color_theme('green'); self.vars={}; self.paths={}; self.messages=queue.Queue(); self._build(); self.after(100,self._poll)
    def _build(self):
        self.grid_columnconfigure(1,weight=1); self.grid_rowconfigure(1,weight=1)
        side=ctk.CTkFrame(self,width=220,corner_radius=0,fg_color='#101914'); side.grid(row=0,column=0,rowspan=2,sticky='nsew'); side.grid_propagate(False)
        ctk.CTkLabel(side,text='RCCI PULSE',font=ctk.CTkFont(size=19,weight='bold'),text_color='#baf091').pack(anchor='w',padx=24,pady=(35,2)); ctk.CTkLabel(side,text='EXPERIMENTAL LAB',font=ctk.CTkFont(size=10),text_color='#829389').pack(anchor='w',padx=25,pady=(0,35))
        for text,active in [('◈  Run analysis',True),('◌  Combustion signals',False),('▦  Reports & exports',False)]: ctk.CTkButton(side,text=text,anchor='w',fg_color='#263a2b' if active else 'transparent',hover_color='#263a2b',text_color='#baf091' if active else '#91a095',height=38,corner_radius=7,command=lambda:None).pack(fill='x',padx=13,pady=3)
        ctk.CTkLabel(side,text='●  LOCAL ANALYSIS\n     Files stay on this machine',justify='left',anchor='w',text_color='#7e9283',font=ctk.CTkFont(size=10)).pack(side='bottom',fill='x',padx=22,pady=24)
        head=ctk.CTkFrame(self,fg_color='transparent'); head.grid(row=0,column=1,sticky='ew',padx=28,pady=(25,14)); head.grid_columnconfigure(0,weight=1)
        ctk.CTkLabel(head,text='SINGLE-CYLINDER · FOUR-STROKE · RCCI',text_color='#8aa08e',font=ctk.CTkFont(size=10)).grid(row=0,column=0,sticky='w'); ctk.CTkLabel(head,text='Experimental analysis cockpit',font=ctk.CTkFont(size=25,weight='bold')).grid(row=1,column=0,sticky='w',pady=(5,0)); ctk.CTkLabel(head,text='A modern shell for the same defensible calculations and report workflow.',text_color='#91a198',font=ctk.CTkFont(size=11)).grid(row=2,column=0,sticky='w',pady=(4,0)); ctk.CTkLabel(head,text='● READY',text_color='#baf091',font=ctk.CTkFont(size=10,weight='bold')).grid(row=1,column=1,rowspan=2,padx=14)
        body=ctk.CTkFrame(self,fg_color='transparent'); body.grid(row=1,column=1,sticky='nsew',padx=28,pady=(0,25)); body.grid_columnconfigure(0,weight=0); body.grid_columnconfigure(1,weight=1); body.grid_rowconfigure(0,weight=1)
        left=ctk.CTkScrollableFrame(body,width=430,label_text='01 · Inputs and experimental files',label_font=ctk.CTkFont(size=14,weight='bold')); left.grid(row=0,column=0,sticky='nsew',padx=(0,15)); left.grid_columnconfigure(0,weight=1)
        for key,label in [('trace','Instrument workbook (.xlsx)'),('cycles','Consecutive-cycle pressure (.xls/.xlsx)'),('performance','Performance report (.xls/.xlsx)')]: self._file_row(left,key,label)
        ctk.CTkLabel(left,text='02 · Operating point & assumptions',anchor='w',font=ctk.CTkFont(size=14,weight='bold')).grid(row=3,column=0,sticky='ew',pady=(23,10))
        form=ctk.CTkFrame(left,fg_color='transparent'); form.grid(row=4,column=0,sticky='ew'); form.grid_columnconfigure(0,weight=1); form.grid_columnconfigure(1,weight=1)
        for i,(key,label,unit,default) in enumerate(FIELDS):
            r,c=divmod(i,2); wrap=ctk.CTkFrame(form,fg_color='transparent'); wrap.grid(row=r,column=c,sticky='ew',padx=4,pady=4); wrap.grid_columnconfigure(0,weight=1); ctk.CTkLabel(wrap,text=f'{label} {f"({unit})" if unit else ""}',anchor='w',font=ctk.CTkFont(size=10),text_color='#aebcb1').grid(row=0,column=0,sticky='w'); var=ctk.StringVar(value=default); self.vars[key]=var; ctk.CTkEntry(wrap,textvariable=var,height=30).grid(row=1,column=0,sticky='ew',pady=(3,0))
        ctk.CTkLabel(left,text='Blank measurements stay unavailable. Air flow → AFR; LHV → independent efficiencies; γ → HRR/CHRR.',wraplength=390,justify='left',text_color='#809086',font=ctk.CTkFont(size=10)).grid(row=5,column=0,sticky='w',pady=(16,4)); self.run_btn=ctk.CTkButton(left,text='Run analysis  →',height=42,fg_color='#baf091',hover_color='#d0f8ae',text_color='#172217',font=ctk.CTkFont(size=12,weight='bold'),command=self.confirm_run); self.run_btn.grid(row=6,column=0,sticky='ew',pady=(14,4)); self.progress=ctk.CTkProgressBar(left,mode='indeterminate',height=4); self.progress.grid(row=7,column=0,sticky='ew',pady=(2,15)); self.progress.set(0)
        right=ctk.CTkFrame(body,fg_color='transparent'); right.grid(row=0,column=1,sticky='nsew'); right.grid_columnconfigure(0,weight=1); right.grid_rowconfigure(2,weight=1)
        self.status=ctk.CTkLabel(right,text='Ready for your first run',anchor='w',font=ctk.CTkFont(size=18,weight='bold')); self.status.grid(row=0,column=0,sticky='ew',pady=(4,4)); self.substatus=ctk.CTkLabel(right,text='Upload the instrument workbook, then add cycle and performance files for the fullest validation.',anchor='w',justify='left',wraplength=650,text_color='#91a198',font=ctk.CTkFont(size=11)); self.substatus.grid(row=1,column=0,sticky='ew',pady=(0,14)); self.output=ctk.CTkTextbox(right,corner_radius=10,font=ctk.CTkFont(family='Consolas',size=11)); self.output.grid(row=2,column=0,sticky='nsew'); self.output.insert('1.0','Analysis output will appear here.\n\nReports and plots are written to the output folder you select after a successful run.')
    def _file_row(self,parent,key,label):
        row=ctk.CTkFrame(parent,fg_color='#18251b',corner_radius=7); row.grid(row=len(self.paths),column=0,sticky='ew',pady=5); row.grid_columnconfigure(0,weight=1); self.paths[key]=ctk.StringVar(value=label); ctk.CTkLabel(row,textvariable=self.paths[key],anchor='w',font=ctk.CTkFont(size=10),text_color='#b7c5b8').grid(row=0,column=0,sticky='ew',padx=11,pady=(8,1)); ctk.CTkButton(row,text='Browse',width=68,height=25,command=lambda k=key:self.browse(k)).grid(row=1,column=0,sticky='e',padx=8,pady=(0,7))
    def browse(self,key):
        path=filedialog.askopenfilename(filetypes=[('Excel data','*.xls *.xlsx'),('All files','*.*')]);
        if path:self.paths[key].set(path)
    def confirm_run(self):
        if not self.paths['trace'].get(): messagebox.showwarning('Instrument workbook required','Select the instrument workbook before starting.'); return
        if messagebox.askyesno('Start analysis?','Read the selected workbooks and create a new report folder? Existing files will not be modified.',parent=self): self.start_run()
    def start_run(self):
        out=filedialog.askdirectory(title='Choose report output folder');
        if not out:return
        self.run_btn.configure(state='disabled',text='Analyzing…'); self.progress.start(); self.status.configure(text='Running defensible analysis…'); self.substatus.configure(text='Reading files · calculating channels · writing report')
        threading.Thread(target=self._worker,args=(Path(out),),daemon=True).start()
    def _worker(self,out):
        try:
            trace_path=Path(self.paths['trace'].get()); cycle_path=Path(self.paths['cycles'].get()) if self.paths['cycles'].get() else None; perf_path=Path(self.paths['performance'].get()) if self.paths['performance'].get() else None; vals={k:v.get() for k,v in self.vars.items()}; performance=engine.load_performance_report(str(perf_path) if perf_path else '',context_paths=[str(p) for p in (cycle_path,trace_path) if p],rpm_hint=vals.get('rpm',''),torque_hint=vals.get('torque',''))
            for k in ('rpm','torque','fuel_kgh'):
                if math.isfinite(engine.number(performance.get(k))): vals[k]=str(performance[k])
            trace=engine.load_instrument(trace_path); metrics=engine.analyze(trace,cycle_path if cycle_path and cycle_path.exists() else None,vals,out,performance); self.messages.put(('ok',metrics,out))
        except Exception as exc:self.messages.put(('error',str(exc),None))
    def _poll(self):
        try:
            kind,payload,out=self.messages.get_nowait(); self.progress.stop(); self.run_btn.configure(state='normal',text='Run analysis  →')
            if kind=='error': self.status.configure(text='Analysis failed'); self.substatus.configure(text=payload); messagebox.showerror('Analysis error',payload,parent=self)
            else:
                self.status.configure(text='Analysis complete ✓'); self.substatus.configure(text=f'Report and plots saved in {out}'); self.output.delete('1.0','end');
                for k,v in payload.items(): self.output.insert('end',f'{k}: {v:.6g}\n' if isinstance(v,(int,float)) and math.isfinite(v) else f'{k}: unavailable\n')
                messagebox.showinfo('Analysis complete',f'Report saved to:\n{Path(out)/"report.html"}',parent=self)
        except queue.Empty: pass
        self.after(100,self._poll)

if __name__=='__main__': App().mainloop()
