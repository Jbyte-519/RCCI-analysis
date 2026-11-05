#!/usr/bin/env python3
"""GUI for defensible RCCI test-run analysis. See README_RCCI_GUI.txt."""
from __future__ import annotations
import math, re, traceback
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


def number(v):
    try:
        x=float(v)
        return x if np.isfinite(x) else np.nan
    except (TypeError,ValueError): return np.nan

def findcol(df, names):
    lut={re.sub(r'[^a-z0-9]','',str(c).lower()):c for c in df.columns}
    for n in names:
        k=re.sub(r'[^a-z0-9]','',n.lower())
        if k in lut:return lut[k]
    return None

def _load_target_from_paths(paths):
    """Read the requested load from a selected run filename when present."""
    for path in paths:
        if not path:
            continue
        match=re.search(r'load\s*([0-9]+(?:\.[0-9]+)?)\s*nm',str(path),re.IGNORECASE)
        if match:
            return float(match.group(1))
    return np.nan

def load_performance_report(path, context_paths=(), rpm_hint=np.nan, torque_hint=np.nan):
    """Extract the performance row belonging to the selected run.

    Performance Report.xls contains a history of load points, so row zero is
    not necessarily the selected run. A LoadXX Nm token in the selected trace
    or cycle filename is used as the measured-load selector. If no such token
    exists, an explicitly entered nonzero torque is used; with one report row
    that row is unambiguous.
    """
    if not path or not Path(path).exists():
        return {}
    report=pd.read_excel(path,sheet_name=0)
    if report.empty:
        return {}
    report.columns=[str(c).strip() for c in report.columns]
    aliases={
        'rpm':['RPM'], 'torque':['Torque (Nm)','Torque'],
        'fuel_kgh':['Fuel (Kg/h)','Fuel kg/h'], 'sfc':['SFC (kg/kWh)'],
        'ip_reported':['IHP(kW)','IHP (kW)'], 'bp_reported':['BHP(kW)','BHP (kW)'],
        'ite_reported':['I.T.Effi. (%)'], 'bte_reported':['B.T.Effi. (%)'],
        'mech_reported':['Mech.Effi. (%)'],
    }
    cols={key:findcol(report,names) for key,names in aliases.items()}
    required=[cols['rpm'],cols['torque']]
    if any(c is None for c in required):
        raise ValueError('Performance report must contain RPM and Torque (Nm) columns.')
    numeric=report.copy()
    for col in set(c for c in cols.values() if c is not None):
        numeric[col]=pd.to_numeric(numeric[col],errors='coerce')
    valid=numeric[cols['rpm']].notna() & numeric[cols['torque']].notna()
    numeric=numeric.loc[valid]
    if numeric.empty:
        return {}
    target=_load_target_from_paths(context_paths)
    if not np.isfinite(target) and np.isfinite(number(torque_hint)) and abs(number(torque_hint))>0:
        target=number(torque_hint)
    if np.isfinite(target):
        distances=(numeric[cols['torque']]-target).abs()
        nearest=distances.min()
        candidates=numeric.loc[distances==nearest]
        # Reports append repeated samples for a load point. Keep the latest
        # row among equal load matches so a later 15 Nm measurement is not
        # replaced by an earlier 15 Nm row from the same report.
        idx=candidates.index[-1]
    elif len(numeric)==1:
        idx=numeric.index[0]
    elif np.isfinite(number(rpm_hint)):
        idx=(numeric[cols['rpm']]-number(rpm_hint)).abs().idxmin()
    else:
        raise ValueError('Performance report has multiple rows; select a run file containing LoadXX Nm or enter its torque.')
    row=numeric.loc[idx]
    result={}
    for key,col in cols.items():
        result[key]=number(row[col]) if col is not None else np.nan
    result['selected_row_index']=int(idx)
    result['selected_load_target_Nm']=target
    return result

def integrate_closed_pv(pressure_bar, volume_cm3):
    """Integrate one complete four-stroke pressure-volume loop in joules."""
    pressure=np.asarray(pressure_bar,dtype=float); volume=np.asarray(volume_cm3,dtype=float)
    if pressure.size!=volume.size or pressure.size<2:
        raise ValueError('Pressure and volume arrays must have the same length.')
    finite=np.isfinite(pressure)&np.isfinite(volume)
    pressure=pressure[finite]; volume=volume[finite]
    if pressure.size<2:
        return np.nan
    # The supplied files contain -360..+359 degCA. Append the first point to
    # close the 720-degree loop instead of silently dropping the final segment.
    return float(np.trapezoid(np.r_[pressure,pressure[0]]*1e5,
                              np.r_[volume,volume[0]]*1e-6))

def geometric_swept_volume_cm3(bore_mm, stroke_mm):
    """Return cylinder displacement from bore and stroke in mm and cm³."""
    bore=number(bore_mm); stroke=number(stroke_mm)
    if not np.isfinite(bore) or not np.isfinite(stroke) or bore<=0 or stroke<=0:
        return np.nan
    return float(np.pi/4*(bore/10.0)**2*(stroke/10.0))

def load_instrument(path, sheet=0):
    raw=pd.read_excel(path,sheet_name=sheet,header=None)
    header=None
    for i in range(min(30,len(raw))):
        vals=[str(x).strip().lower() for x in raw.iloc[i].tolist()]
        if any('angle' in x for x in vals) and any('pressure' in x for x in vals): header=i;break
    if header is None: raise ValueError('Could not locate C. Angle / Pressure header in instrument workbook.')
    d=pd.read_excel(path,sheet_name=sheet,header=header); d.columns=[str(c).strip() for c in d.columns]
    ac=findcol(d,['C. Angle','angle']); pc=findcol(d,['Pressure']); vc=findcol(d,['Volume'])
    if not all([ac,pc,vc]): raise ValueError('Instrument data must include angle, pressure, and volume columns.')
    out=pd.DataFrame({'angle':pd.to_numeric(d[ac],errors='coerce'),'pressure_bar':pd.to_numeric(d[pc],errors='coerce'),'volume':pd.to_numeric(d[vc],errors='coerce')})
    for key,aliases in {'mfb':['MFB'],'nhr':['NHR'],'mgt':['MGT'],'chrr':['CHRR'],'rpr_reported':['R.P.R']}.items():
        c=findcol(d,aliases)
        if c:out[key]=pd.to_numeric(d[c],errors='coerce')
    return out.dropna(subset=['angle','pressure_bar','volume']).reset_index(drop=True)

def xing(x,y,q):
    inds=np.flatnonzero((y[:-1]<q)&(y[1:]>=q))
    if not len(inds):return np.nan
    i=inds[0]; return float(x[i]+(q-y[i])*(x[i+1]-x[i])/(y[i+1]-y[i]))

def _analyze_cycle_primary(trace, cycles, values, outdir, performance=None):
    ca=trace.angle.to_numpy(float); p=trace.pressure_bar.to_numpy(float); V=trace.volume.to_numpy(float)
    if len(ca)<10 or np.any(np.diff(ca)<=0): raise ValueError('Trace needs at least 10 increasing crank-angle samples.')
    validation_warnings=[]
    angle_step=np.diff(ca)
    step_deg=float(np.nanmedian(angle_step)) if len(angle_step) else np.nan
    closed_angle_span=float(ca[-1]-ca[0]+step_deg) if np.isfinite(step_deg) else np.nan
    if not np.isfinite(step_deg) or not np.allclose(angle_step,step_deg,rtol=0,atol=1e-9):
        validation_warnings.append('Crank-angle samples are not uniformly spaced; derivative and p-V integration use the supplied angle values.')
    if not np.isfinite(closed_angle_span) or abs(closed_angle_span-720.0)>max(1.0,abs(step_deg)*1.5):
        validation_warnings.append(f'Crank-angle coverage is {closed_angle_span:.3g} degCA including the closing sample, not a complete 720 degCA loop.')
    if len(ca)>1 and abs(V[-1]-V[0])>0.02*(np.nanmax(V)-np.nanmin(V)):
        validation_warnings.append('First and last volume samples do not close within 2% of the measured swept volume; p-V closure should be reviewed.')
    validation_warnings.append('Pressure is interpreted as bar and volume as cm³ from the instrument labels; the consecutive-cycle file has no explicit unit metadata.')
    measured_swept_cm3=float(np.nanmax(V)-np.nanmin(V))
    geometric_swept_cm3=geometric_swept_volume_cm3(values.get('bore_mm',''),values.get('stroke_mm',''))
    nominal_displacement_cm3=number(values.get('displacement_cc',''))
    if np.isfinite(geometric_swept_cm3) and measured_swept_cm3>0:
        if abs(measured_swept_cm3-geometric_swept_cm3)/geometric_swept_cm3>0.05:
            validation_warnings.append(f'Measured swept volume ({measured_swept_cm3:.2f} cm³) differs from bore-stroke volume ({geometric_swept_cm3:.2f} cm³) by more than 5%; verify volume calibration and crank-angle reference. Measured volume is retained for IMEP.')
    if np.isfinite(nominal_displacement_cm3) and np.isfinite(geometric_swept_cm3) and nominal_displacement_cm3>0:
        if abs(nominal_displacement_cm3-geometric_swept_cm3)/nominal_displacement_cm3>0.05:
            validation_warnings.append(f'Nominal displacement ({nominal_displacement_cm3:.2f} cm³) differs from bore-stroke volume ({geometric_swept_cm3:.2f} cm³) by more than 5%; verify engine specifications.')
    if measured_swept_cm3<=0:
        validation_warnings.append('Measured swept volume is non-positive; IMEP cannot be validated.')
    validation_warnings.append('The main trace and cycle traces are integrated in increasing crank-angle order with the first sample appended to close the 720-degree p-V loop.')
    # Pressure is smoothed only for derivatives. Raw measured maxima remain raw.
    win=int(values['smooth']); win=max(1,win); win+=1 if win>1 and win%2==0 else 0
    if win>=5 and win<len(p):
        from scipy.signal import savgol_filter
        ps=savgol_filter(p,win,min(2,win-1),mode='interp')
    else: ps=p.copy()
    dp=np.gradient(ps,ca); dV=np.gradient(V*1e-6,ca); ppa=ps*1e5
    gamma=number(values['gamma']); hrr=np.full(len(ca),np.nan); chrr=np.full(len(ca),np.nan)
    if np.isfinite(gamma) and gamma>1:
        hrr=gamma/(gamma-1)*ppa*dV + (V*1e-6)*1e5*dp/(gamma-1)
        chrr=np.r_[0,np.cumsum((hrr[1:]+hrr[:-1])*np.diff(ca)/2)]
        validation_warnings.append(f'Independent apparent HRR uses constant gamma={gamma:.6g}, pressure in Pa, volume in m³, and derivatives per degree CA; wall heat transfer and variable properties are not included.')
    elif np.isfinite(gamma):
        validation_warnings.append('Specific heat ratio gamma must be greater than 1; independent HRR and CHRR were not calculated.')
    else:
        validation_warnings.append('Independent apparent HRR and CHRR are unavailable because gamma was not supplied.')
    validation_warnings.append('Instrument MGT convention/units remain unverified.')
    validation_warnings.append('Instrument CHRR units remain ambiguous; reported NHR/CHRR are retained separately from calculated HRR/CHRR.')
    mfb=trace['mfb'].to_numpy(float) if 'mfb' in trace else np.full(len(ca),np.nan)
    # Workbook MFB is percent; accept fraction series too.
    if np.nanmax(mfb)>1.5:mfb=mfb/100
    mfb=np.clip(mfb,0,1)
    combustion=(ca>=0)&(ca<=180)&np.isfinite(mfb)
    cx,cy=ca[combustion],mfb[combustion]
    ca10,ca50,ca90=[xing(cx,cy,q) for q in (.1,.5,.9)] if len(cx)>2 else [np.nan]*3
    imax=int(np.nanargmax(p)); rpr_i=int(np.nanargmax(dp));
    # Use the same closed-loop p-dV integration for the main trace and every
    # consecutive cycle. The measured volume trace supplies swept volume for
    # IMEP; the engine displacement input is reserved for BMEP.
    main_work=integrate_closed_pv(p, V)
    if np.isfinite(main_work) and main_work<0:
        validation_warnings.append('Main pressure-volume loop integrates negative work in the supplied crank-angle direction; sign was retained and not forced positive.')
    swept=measured_swept_cm3
    swept_m3=swept*1e-6
    main_imep=main_work/swept_m3/1e5 if swept_m3>0 else np.nan
    performance=performance or {}
    rpm=number(performance.get('rpm',values['rpm'])); torque=number(performance.get('torque',values['torque'])); disp=number(values['displacement_cc'])*1e-6
    filename_rpm=number(performance.get('filename_rpm_target'))
    if np.isfinite(filename_rpm) and np.isfinite(rpm) and abs(filename_rpm-rpm)>5:
        validation_warnings.append(f'Selected report RPM ({rpm:.4g}) differs from the RPM in the run filename ({filename_rpm:.4g}) by more than 5 rpm; verify that the performance row and pressure trace belong to the same operating point.')
    main_ip=main_work*rpm/120000 if rpm>0 else np.nan

    # Validate and integrate only complete, finite consecutive pressure cycles.
    # The source cycle workbook has no crank-angle column, so row-for-row
    # alignment to the instrument angle/volume trace is explicit and reported.
    cycle_stats={}; cycle_df=None; cycle_work=np.array([],dtype=float); cycle_imep=np.array([],dtype=float)
    if cycles is None:
        validation_warnings.append('No consecutive-cycle pressure file was supplied; main-trace indicated results remain primary.')
    else:
        try:
            cdf=pd.read_excel(cycles,header=0).apply(pd.to_numeric,errors='coerce')
            cdf=cdf.dropna(axis=1,how='all').dropna(axis=0,how='all')
            if len(cdf)!=len(ca):
                validation_warnings.append(f'Cycle alignment is invalid: {len(cdf)} pressure rows were supplied for {len(ca)} crank-angle samples; cycle averages were not used as primary results.')
            else:
                cdf.index=ca; cycle_df=cdf
                valid_columns=[]
                for col in cdf.columns:
                    arr=cdf[col].to_numpy(float)
                    if len(arr)!=len(ca) or not np.isfinite(arr).all():
                        validation_warnings.append(f'Cycle {col} contains missing/non-numeric pressure samples and was excluded from cycle averages.')
                    else:
                        valid_columns.append(col)
                if valid_columns:
                    valid_cycles=cdf[valid_columns]
                    cycle_work=np.array([integrate_closed_pv(valid_cycles[col].to_numpy(float),V) for col in valid_columns])
                    cycle_imep=cycle_work/swept_m3/1e5 if swept_m3>0 else np.full(len(cycle_work),np.nan)
                    if np.any(cycle_work<0):
                        validation_warnings.append('At least one consecutive pressure cycle integrates negative work in the supplied crank-angle direction; signs were retained.')
                    main_peak=float(np.nanmax(np.abs(p))) if np.isfinite(p).any() else np.nan
                    cycle_peak=float(np.nanmax(np.abs(valid_cycles.to_numpy(float)))) if np.isfinite(valid_cycles.to_numpy(float)).any() else np.nan
                    if np.isfinite(main_peak) and np.isfinite(cycle_peak) and main_peak>0 and not 0.5<=cycle_peak/main_peak<=2.0:
                        validation_warnings.append(f'Cycle pressure magnitude ({cycle_peak:.4g}) is inconsistent with the main trace ({main_peak:.4g}); verify pressure units and column selection.')
                    validation_warnings.append(f'{len(valid_columns)} valid consecutive pressure cycles were integrated row-for-row against the {len(ca)}-sample volume trace.')
                else:
                    validation_warnings.append('No complete finite consecutive pressure cycles were available for averaging.')
        except Exception as exc:
            validation_warnings.append(f'Consecutive-cycle file could not be validated or read: {exc}')
    cycle_available=len(cycle_work)>0 and np.isfinite(cycle_work).any()
    cycle_mean_work=float(np.nanmean(cycle_work)) if cycle_available else np.nan
    cycle_mean_imep=float(np.nanmean(cycle_imep)) if cycle_available else np.nan
    cycle_ip=cycle_mean_work*rpm/120000 if cycle_available and rpm>0 else np.nan
    primary_work=cycle_mean_work if cycle_available else main_work
    primary_imep=cycle_mean_imep if cycle_available else main_imep
    ip=cycle_ip if cycle_available else main_ip
    if cycle_available:
        validation_warnings.append('Cycle-averaged indicated work, IMEP, and IP are the primary indicated results; main-trace values are retained separately.')
    brake_work=4*np.pi*torque if np.isfinite(torque) else np.nan
    bp=brake_work*rpm/120000 if rpm>0 and np.isfinite(brake_work) else np.nan
    bmep=brake_work/disp/1e5 if disp>0 and np.isfinite(brake_work) else np.nan
    fuel=number(performance.get('fuel_kgh',values['fuel_kgh'])); lhv=number(values['lhv_mjkg']); air=number(values['air_kgh'])
    fuel_s=fuel/3600 if fuel>0 else np.nan
    ite_calculated=ip/(fuel_s*lhv*1e6)*100 if fuel_s>0 and lhv>0 and np.isfinite(ip) else np.nan
    bte_calculated=bp/(fuel_s*lhv*1e6)*100 if fuel_s>0 and lhv>0 and np.isfinite(bp) and bp>0 else np.nan
    bsfc_calculated=fuel/bp if fuel>0 and bp>0 else np.nan
    ite_reported=number(performance.get('ite_reported'))
    bte_reported=number(performance.get('bte_reported'))
    bsfc_reported=number(performance.get('sfc'))
    # The report values are measured software outputs. Prefer them for the
    # named report metrics while retaining independent calculations separately.
    ite=ite_reported if np.isfinite(ite_reported) else ite_calculated
    bte=bte_reported if np.isfinite(bte_reported) else bte_calculated
    bsfc=bsfc_reported if np.isfinite(bsfc_reported) else bsfc_calculated
    afr=air/fuel if air>0 and fuel>0 else np.nan
    reported_ip=number(performance.get('ip_reported')); reported_bp=number(performance.get('bp_reported'))
    reported_imep=number(performance.get('imep_reported')); reported_mech=number(performance.get('mech_reported'))
    report_mech_from_report=reported_bp/reported_ip*100 if reported_ip>0 and np.isfinite(reported_bp) else np.nan
    if np.isfinite(reported_bp) and np.isfinite(bp) and abs(reported_bp-bp)>max(0.03,0.02*abs(reported_bp)):
        validation_warnings.append(f'Calculated BP ({bp:.4g} kW) differs from reported BP ({reported_bp:.4g} kW); calculated BP uses selected RPM and torque.')
    if np.isfinite(reported_ip) and np.isfinite(ip) and abs(reported_ip-ip)>max(0.03,0.02*abs(reported_ip)):
        validation_warnings.append(f'Instrument IP ({reported_ip:.4g} kW) differs from the selected indicated basis ({ip:.4g} kW); mechanical efficiency values therefore use different averaging bases.')
    if np.isfinite(reported_imep) and np.isfinite(primary_imep) and abs(reported_imep-primary_imep)>max(0.05,0.02*abs(reported_imep)):
        validation_warnings.append(f'Instrument IMEP ({reported_imep:.4g} bar) differs from the selected indicated IMEP ({primary_imep:.4g} bar).')
    if np.isfinite(reported_mech) and np.isfinite(report_mech_from_report) and abs(reported_mech-report_mech_from_report)>0.5:
        validation_warnings.append('Reported mechanical efficiency is inconsistent with the report BP/IP pair after rounding; it was retained as reported.')
    if not np.isfinite(reported_imep):
        validation_warnings.append('Instrument-reported IMEP is unavailable in the selected performance report; no value was inferred from IHP.')
    r=lambda x: float(x) if np.isfinite(x) else np.nan
    metrics={
      'Peak cylinder pressure (bar)':p[imax], 'Peak pressure angle (deg CA)':ca[imax],
      'Maximum calculated RPR (bar/deg CA)':np.nanmax(dp),'RPR peak angle (deg CA)':ca[rpr_i],
      'Instrument reported maximum RPR (bar/deg CA)':float(trace.rpr_reported.max()) if 'rpr_reported' in trace else np.nan,
      'Instrument NHR maximum (as reported, J/deg CA)':float(trace.nhr.max()) if 'nhr' in trace else np.nan,
      'Independent apparent net HRR maximum (J/deg CA)':float(np.nanmax(hrr)) if np.isfinite(hrr).any() else np.nan,
      'Independent cumulative heat release span (J)':float(chrr[-1]-chrr[0]) if np.isfinite(chrr).any() else np.nan,
      'Independent cumulative heat release final (J)':float(chrr[-1]) if np.isfinite(chrr).any() else np.nan,
      'Instrument CHRR maximum (workbook unit label ambiguous)':float(trace.chrr.max()) if 'chrr' in trace else np.nan,
      'CA10 (deg CA)':ca10,'CA50 (deg CA)':ca50,'CA90 (deg CA)':ca90,'Combustion duration CA10-CA90 (deg CA)':ca90-ca10 if np.isfinite(ca10+ca90) else np.nan,
      'Instrument MGT maximum (reported value; likely K)':float(trace.mgt.max()) if 'mgt' in trace else np.nan,
      'Calculated indicated work per cycle (J)':primary_work,'IMEP (bar)':primary_imep,'IP (kW)':ip,'BMEP (bar)':bmep,'BP (kW)':bp,
      'Main trace indicated work per cycle (J)':main_work,'Main trace IMEP (bar)':main_imep,'Main trace IP (kW)':main_ip,
      'Cycle averaged indicated work per cycle (J)':cycle_mean_work,'Cycle averaged IMEP (bar)':cycle_mean_imep,'Cycle averaged IP (kW)':cycle_ip,
      'BTE (%)':bte,'ITE (%)':ite,'BSFC (kg/kWh)':bsfc,'AFR (kg air/kg fuel)':afr,'Mechanical efficiency (%)':bp/ip*100 if ip>0 and np.isfinite(bp) else np.nan,
      'Calculated BTE (%)':bte_calculated,'Calculated ITE (%)':ite_calculated,'Calculated BSFC (kg/kWh)':bsfc_calculated,
      'Reported BTE (%)':bte_reported,'Reported ITE (%)':ite_reported,'Reported BSFC (kg/kWh)':bsfc_reported,
      'Main trace mechanical efficiency (%)':bp/main_ip*100 if main_ip>0 and np.isfinite(bp) else np.nan,
      'Cycle averaged mechanical efficiency (%)':bp/cycle_ip*100 if cycle_ip>0 and np.isfinite(bp) else np.nan,
      'Reported mechanical efficiency (%)':reported_mech,'Mechanical efficiency from reported BP/IP (%)':report_mech_from_report,
      'Reported IP (kW)':reported_ip,'Reported IMEP (bar)':reported_imep,'Reported BP (kW)':reported_bp,
      'Selected report RPM':rpm,'Run filename RPM':filename_rpm,'Selected report torque (Nm)':torque,'Selected report fuel flow (kg/h)':fuel,
      'Measured swept volume (cm3)':swept,'Geometric bore-stroke swept volume (cm3)':geometric_swept_cm3,'BMEP displacement basis (cm3)':number(values['displacement_cc']),
      'Engine displacement input (cc)':number(values['displacement_cc']),'Compression ratio input':number(values['compression_ratio']),
      'Bore input (mm)':number(values['bore_mm']),'Stroke input (mm)':number(values['stroke_mm']),
      'E20 low-reactivity fuel energy share input (%)':number(values['e20_energy_share']),
    }
    # Preserve the existing cycle-variation outputs while using only valid
    # cycles for the averages above.
    if cycle_available:
        maxima=valid_cycles.max(axis=0,skipna=True).to_numpy(float)
        cycle_stats.update({'Cycles analyzed':len(maxima),'Mean cycle peak pressure (same units as raw cycle file)':float(np.nanmean(maxima)),
                            'SD cycle peak pressure (same units as raw cycle file)':float(np.nanstd(maxima,ddof=1)) if len(maxima)>1 else np.nan,
                            'COV of cycle peak pressure (%)':float(np.nanstd(maxima,ddof=1)/np.nanmean(maxima)*100) if len(maxima)>1 and np.nanmean(maxima)!=0 else np.nan,
                            'Mean indicated work from consecutive pressure cycles (J)':cycle_mean_work,
                            'SD indicated work from consecutive pressure cycles (J)':float(np.nanstd(cycle_work,ddof=1)) if len(cycle_work)>1 else np.nan,
                            'COV indicated work from consecutive pressure cycles (%)':float(np.nanstd(cycle_work,ddof=1)/np.nanmean(cycle_work)*100) if len(cycle_work)>1 and np.nanmean(cycle_work)!=0 else np.nan,
                            'Mean IMEP from consecutive pressure cycles (bar)':cycle_mean_imep})
    elif cycles is not None:
        cycle_stats['Cycles analyzed']=0
    metrics.update(cycle_stats)
    metrics['Main-to-cycle work difference (%)']=(cycle_mean_work-main_work)/main_work*100 if cycle_available and np.isfinite(main_work) and main_work!=0 else np.nan
    metrics['Crank-angle closed span (deg CA)']=closed_angle_span
    metrics['HRR gamma used']=gamma
    metrics['HRR derivative smoothing window (samples)']=win
    metrics['Validation warning count']=len(validation_warnings)
    metrics['Validation warnings']=' | '.join(validation_warnings) if validation_warnings else 'None'
    # Measured emissions remain user-supplied; no inferred/fabricated values.
    for label,key,unit in [('NOx','nox','ppm'),('CO','co','% or ppm as entered'),('HC','hc','ppm'),('CO2','co2','%'),('Smoke opacity','smoke','%')]:
        metrics[f'{label} ({unit})']=number(values[key])
    metrics['Exhaust gas temperature (deg C)']=number(values['egt'])
    outdir.mkdir(parents=True,exist_ok=True)
    trace_out=trace.copy(); trace_out['pressure_smoothed_bar']=ps; trace_out['calculated_RPR_bar_per_deg']=dp; trace_out['calculated_HRR_J_per_deg']=hrr; trace_out['calculated_CHRR_J']=chrr; trace_out['MFB_fraction']=mfb
    pd.DataFrame([metrics]).to_csv(outdir/'summary.csv',index=False)
    trace_out.to_csv(outdir/'crank_angle_analysis.csv',index=False)
    with pd.ExcelWriter(outdir/'analysis.xlsx',engine='openpyxl') as xw:
        pd.DataFrame([metrics]).to_excel(xw,index=False,sheet_name='Summary'); trace_out.to_excel(xw,index=False,sheet_name='Crank angle')
        if cycle_df is not None: cycle_df.to_excel(xw,sheet_name='Cycle pressures')
    figs=[]
    for slug,y,title,ylabel in [('pressure',p,'Cylinder pressure','Pressure (bar)'),('rpr',dp,'Pressure rise rate','bar/deg CA'),('hrr',hrr,'Independent apparent net heat release','J/deg CA'),('mfb',mfb*100,'Mass fraction burned','MFB (%)')]:
        if not np.isfinite(y).any():continue
        fig,ax=plt.subplots(figsize=(8,4.5)); ax.plot(ca,y); ax.set(xlabel='Crank angle (deg CA)',ylabel=ylabel,title=title); ax.grid(alpha=.25); fig.tight_layout(); fig.savefig(outdir/f'{slug}.png',dpi=180); plt.close(fig)
    def html_value(value):
        if isinstance(value,(float,int,np.number)):
            return f'{float(value):.6g}' if np.isfinite(value) else 'unavailable'
        return str(value)
    rows=''.join(f'<tr><td>{k}</td><td>{html_value(v)}</td></tr>' for k,v in metrics.items())
    absent=[k for k,v in metrics.items() if isinstance(v,(float,int,np.number)) and not np.isfinite(v)]
    warning_html=''.join(f'<li>{warning}</li>' for warning in validation_warnings)
    html=f'''<!doctype html><meta charset="utf-8"><title>RCCI experimental analysis</title><style>body{{font:15px Arial;max-width:1000px;margin:30px auto;color:#18212b}}table{{border-collapse:collapse;width:100%}}td,th{{border:1px solid #ccd;padding:7px;text-align:left;vertical-align:top}}th{{background:#edf3f8}}img{{max-width:48%;margin:1%}}.warnings{{background:#fff7df;border:1px solid #e3c36f;padding:12px}}</style><h1>RCCI experimental analysis</h1><p>Calculated from the selected experimental files and GUI inputs. Missing measurements are not estimated.</p><h2>Results</h2><table><tr><th>Parameter</th><th>Value</th></tr>{rows}</table><h2>Not available</h2><p>{', '.join(absent) if absent else 'None'}</p><h2>Validation warnings</h2><div class="warnings"><ul>{warning_html}</ul></div><h2>Plots</h2>{''.join(f'<img src="{slug}.png">' for slug in ['pressure','rpr','hrr','mfb'] if (outdir/f'{slug}.png').exists())}<h2>Methods and assumptions</h2><p>Pressure units are bar and volume is cm³ as labelled by the instrument workbook. Four-stroke, one cylinder is assumed. Cycle-averaged indicated work, IMEP, and IP are primary when complete finite consecutive cycles are available; main-trace and instrument-reported values remain separate. Indicated work is the signed integral ∮p dV over a closed 720-degree trace. IMEP uses the validated measured swept volume; BMEP uses the nominal engine displacement only as an engine specification. IP is work per cycle × rpm / 120 for a four-stroke single cylinder. Brake quantities use measured torque and speed. Independent apparent net HRR uses constant gamma: dQ/dθ=γ/(γ−1)p(dV/dθ)+(1/(γ−1))V(dp/dθ); pressure derivatives use the selected Savitzky-Golay smoothing window, with pressure converted from bar to Pa and volume from cm³ to m³. CA10/50/90 are first MFB crossings on 0–180°CA. Cycle peak-pressure COV uses sample SD / mean over valid pressure columns aligned row-for-row to the trace. Instrument-reported and independently calculated signals are retained separately. MGT convention and CHRR units require verification from instrument documentation. Heat-transfer corrections, crevice flow, variable gas properties, and absolute-pressure calibration are not applied.</p>'''
    (outdir/'report.html').write_text(html,encoding='utf-8')
    return metrics

def analyze(trace, cycles, values, outdir, performance=None):
    """Run the original GUI analysis with corrected report extraction."""
    ca=trace.angle.to_numpy(float); p=trace.pressure_bar.to_numpy(float); V=trace.volume.to_numpy(float)
    if len(ca)<10 or np.any(np.diff(ca)<=0): raise ValueError('Trace needs at least 10 increasing crank-angle samples.')
    win=int(values['smooth']); win=max(1,win); win+=1 if win>1 and win%2==0 else 0
    if win>=5 and win<len(p):
        from scipy.signal import savgol_filter
        ps=savgol_filter(p,win,min(2,win-1),mode='interp')
    else: ps=p.copy()
    dp=np.gradient(ps,ca); dV=np.gradient(V*1e-6,ca); ppa=ps*1e5
    gamma=number(values['gamma']); hrr=np.full(len(ca),np.nan); chrr=np.full(len(ca),np.nan)
    if np.isfinite(gamma) and gamma>1:
        hrr=gamma/(gamma-1)*ppa*dV + (V*1e-6)*1e5*dp/(gamma-1)
        chrr=np.r_[0,np.cumsum((hrr[1:]+hrr[:-1])*np.diff(ca)/2)]
    mfb=trace['mfb'].to_numpy(float) if 'mfb' in trace else np.full(len(ca),np.nan)
    if np.nanmax(mfb)>1.5:mfb=mfb/100
    mfb=np.clip(mfb,0,1)
    combustion=(ca>=0)&(ca<=180)&np.isfinite(mfb)
    cx,cy=ca[combustion],mfb[combustion]
    ca10,ca50,ca90=[xing(cx,cy,q) for q in (.1,.5,.9)] if len(cx)>2 else [np.nan]*3
    imax=int(np.nanargmax(p)); rpr_i=int(np.nanargmax(dp))
    work=integrate_closed_pv(p,V)
    swept=(np.nanmax(V)-np.nanmin(V)); swept_m3=swept*1e-6
    imep=work/swept_m3/1e5 if swept_m3>0 else np.nan
    performance=performance or {}
    rpm=number(performance.get('rpm',values['rpm'])); torque=number(performance.get('torque',values['torque'])); disp=number(values['displacement_cc'])*1e-6
    ip=work*rpm/120000 if rpm>0 else np.nan
    brake_work=4*np.pi*torque if np.isfinite(torque) else np.nan
    bp=brake_work*rpm/120000 if rpm>0 and np.isfinite(brake_work) else np.nan
    bmep=brake_work/disp/1e5 if disp>0 and np.isfinite(brake_work) else np.nan
    fuel=number(performance.get('fuel_kgh',values['fuel_kgh'])); lhv=number(values['lhv_mjkg']); air=number(values['air_kgh'])
    fuel_s=fuel/3600 if fuel>0 else np.nan
    ite_calculated=ip/(fuel_s*lhv*1e6)*100 if fuel_s>0 and lhv>0 and np.isfinite(ip) else np.nan
    bte_calculated=bp/(fuel_s*lhv*1e6)*100 if fuel_s>0 and lhv>0 and np.isfinite(bp) and bp>0 else np.nan
    bsfc_calculated=fuel/bp if fuel>0 and bp>0 else np.nan
    ite_reported=number(performance.get('ite_reported')); bte_reported=number(performance.get('bte_reported')); bsfc_reported=number(performance.get('sfc'))
    ite=ite_reported if np.isfinite(ite_reported) else ite_calculated
    bte=bte_reported if np.isfinite(bte_reported) else bte_calculated
    bsfc=bsfc_reported if np.isfinite(bsfc_reported) else bsfc_calculated
    afr=air/fuel if air>0 and fuel>0 else np.nan
    reported_ip=number(performance.get('ip_reported')); reported_bp=number(performance.get('bp_reported')); reported_mech=number(performance.get('mech_reported'))
    cycle_stats={}; cycle_df=None
    if cycles is not None:
        cdf=pd.read_excel(cycles,header=0); cdf=cdf.apply(pd.to_numeric,errors='coerce')
        cdf=cdf.dropna(axis=1,how='all').dropna(axis=0,how='all')
        if len(cdf)==len(ca):
            cdf.index=ca; cycle_df=cdf
            maxima=cdf.max(axis=0,skipna=True).to_numpy(float)
            cycle_work=np.array([integrate_closed_pv(cdf[col].to_numpy(float),V) for col in cdf.columns])
            cycle_imep=cycle_work/swept_m3/1e5 if swept_m3>0 else np.full(len(cycle_work),np.nan)
            cycle_stats={'Cycles analyzed':len(maxima),'Mean cycle peak pressure (same units as raw cycle file)':float(np.nanmean(maxima)),
                         'SD cycle peak pressure (same units as raw cycle file)':float(np.nanstd(maxima,ddof=1)) if len(maxima)>1 else np.nan,
                         'COV of cycle peak pressure (%)':float(np.nanstd(maxima,ddof=1)/np.nanmean(maxima)*100) if len(maxima)>1 and np.nanmean(maxima)!=0 else np.nan,
                         'Mean indicated work from consecutive pressure cycles (J)':float(np.nanmean(cycle_work)),
                         'SD indicated work from consecutive pressure cycles (J)':float(np.nanstd(cycle_work,ddof=1)) if len(cycle_work)>1 else np.nan,
                         'COV indicated work from consecutive pressure cycles (%)':float(np.nanstd(cycle_work,ddof=1)/np.nanmean(cycle_work)*100) if len(cycle_work)>1 and np.nanmean(cycle_work)!=0 else np.nan,
                         'Mean IMEP from consecutive pressure cycles (bar)':float(np.nanmean(cycle_imep))}
        else: cycle_stats={'Cycle file status':f'Not aligned: {len(cdf)} pressure rows vs {len(ca)} crank-angle samples'}
    metrics={
      'Peak cylinder pressure (bar)':p[imax], 'Peak pressure angle (deg CA)':ca[imax],
      'Maximum calculated RPR (bar/deg CA)':np.nanmax(dp),'RPR peak angle (deg CA)':ca[rpr_i],
      'Instrument reported maximum RPR (bar/deg CA)':float(trace.rpr_reported.max()) if 'rpr_reported' in trace else np.nan,
      'Instrument NHR maximum (as reported, J/deg CA)':float(trace.nhr.max()) if 'nhr' in trace else np.nan,
      'Independent apparent net HRR maximum (J/deg CA)':float(np.nanmax(hrr)) if np.isfinite(hrr).any() else np.nan,
      'Independent cumulative heat release span (J)':float(chrr[-1]-chrr[0]) if np.isfinite(chrr).any() else np.nan,
      'Instrument CHRR maximum (workbook unit label ambiguous)':float(trace.chrr.max()) if 'chrr' in trace else np.nan,
      'CA10 (deg CA)':ca10,'CA50 (deg CA)':ca50,'CA90 (deg CA)':ca90,'Combustion duration CA10-CA90 (deg CA)':ca90-ca10 if np.isfinite(ca10+ca90) else np.nan,
      'Instrument MGT maximum (reported value; likely K)':float(trace.mgt.max()) if 'mgt' in trace else np.nan,
      'Calculated indicated work per cycle (J)':work,'IMEP (bar)':imep,'IP (kW)':ip,'BMEP (bar)':bmep,'BP (kW)':bp,
      'BTE (%)':bte,'ITE (%)':ite,'BSFC (kg/kWh)':bsfc,'AFR (kg air/kg fuel)':afr,'Mechanical efficiency (%)':bp/ip*100 if ip>0 and np.isfinite(bp) else np.nan,
      'Calculated BTE (%)':bte_calculated,'Calculated ITE (%)':ite_calculated,'Calculated BSFC (kg/kWh)':bsfc_calculated,
      'Reported BTE (%)':bte_reported,'Reported ITE (%)':ite_reported,'Reported BSFC (kg/kWh)':bsfc_reported,
      'Reported mechanical efficiency (%)':reported_mech,
      'Reported IP (kW)':reported_ip,'Reported BP (kW)':reported_bp,
      'Selected report RPM':rpm,'Selected report torque (Nm)':torque,'Selected report fuel flow (kg/h)':fuel,
      'Measured swept volume (cm3)':swept,'BMEP displacement basis (cm3)':number(values['displacement_cc']),
      'Engine displacement input (cc)':number(values['displacement_cc']),'Compression ratio input':number(values['compression_ratio']),
      'Bore input (mm)':number(values['bore_mm']),'Stroke input (mm)':number(values['stroke_mm']),
      'E20 low-reactivity fuel energy share input (%)':number(values['e20_energy_share']),
    }
    metrics.update(cycle_stats)
    for label,key,unit in [('NOx','nox','ppm'),('CO','co','% or ppm as entered'),('HC','hc','ppm'),('CO2','co2','%'),('Smoke opacity','smoke','%')]:
        metrics[f'{label} ({unit})']=number(values[key])
    metrics['Exhaust gas temperature (deg C)']=number(values['egt'])
    outdir.mkdir(parents=True,exist_ok=True)
    trace_out=trace.copy(); trace_out['pressure_smoothed_bar']=ps; trace_out['calculated_RPR_bar_per_deg']=dp; trace_out['calculated_HRR_J_per_deg']=hrr; trace_out['calculated_CHRR_J']=chrr; trace_out['MFB_fraction']=mfb
    pd.DataFrame([metrics]).to_csv(outdir/'summary.csv',index=False); trace_out.to_csv(outdir/'crank_angle_analysis.csv',index=False)
    with pd.ExcelWriter(outdir/'analysis.xlsx',engine='openpyxl') as xw:
        pd.DataFrame([metrics]).to_excel(xw,index=False,sheet_name='Summary'); trace_out.to_excel(xw,index=False,sheet_name='Crank angle')
        if cycle_df is not None: cycle_df.to_excel(xw,sheet_name='Cycle pressures')
    for slug,y,title,ylabel in [('pressure',p,'Cylinder pressure','Pressure (bar)'),('rpr',dp,'Pressure rise rate','bar/deg CA'),('hrr',hrr,'Independent apparent net heat release','J/deg CA'),('mfb',mfb*100,'Mass fraction burned','MFB (%)')]:
        if not np.isfinite(y).any():continue
        fig,ax=plt.subplots(figsize=(8,4.5)); ax.plot(ca,y); ax.set(xlabel='Crank angle (deg CA)',ylabel=ylabel,title=title); ax.grid(alpha=.25); fig.tight_layout(); fig.savefig(outdir/f'{slug}.png',dpi=180); plt.close(fig)
    rows=''.join(f'<tr><td>{k}</td><td>{v:.6g}</td></tr>' for k,v in metrics.items() if isinstance(v,(float,int,np.number)) and np.isfinite(v))
    absent=[k for k,v in metrics.items() if isinstance(v,(float,int,np.number)) and not np.isfinite(v)]
    html=f'''<!doctype html><meta charset="utf-8"><title>RCCI experimental analysis</title><style>body{{font:15px Arial;max-width:1000px;margin:30px auto;color:#18212b}}table{{border-collapse:collapse;width:100%}}td,th{{border:1px solid #ccd;padding:7px;text-align:left}}th{{background:#edf3f8}}img{{max-width:48%;margin:1%}}</style><h1>RCCI experimental analysis</h1><p>Calculated from the selected experimental files and GUI inputs. Missing measurements are not estimated.</p><h2>Results</h2><table><tr><th>Parameter</th><th>Value</th></tr>{rows}</table><h2>Not available</h2><p>{', '.join(absent) if absent else 'None'}</p><h2>Plots</h2>{''.join(f'<img src="{slug}.png">' for slug in ['pressure','rpr','hrr','mfb'] if (outdir/f'{slug}.png').exists())}<h2>Methods and assumptions</h2><p>Pressure units are bar and volume is cm³ as labelled by the instrument workbook. Four-stroke, one cylinder is assumed. Indicated work is the signed integral ∮p dV over the supplied trace, divided by measured swept volume for IMEP. IP is work per cycle × rpm / 120 for a four-stroke single cylinder. Brake quantities use measured torque and speed. Independent apparent net HRR uses constant gamma: dQ/dθ=γ/(γ−1)p(dV/dθ)+(1/(γ−1))V(dp/dθ); pressure derivatives use the selected Savitzky-Golay smoothing window. CA10/50/90 are first MFB crossings on 0–180°CA. Cycle peak-pressure COV uses sample SD / mean over pressure columns aligned row-for-row to the trace. Instrument-reported and independently calculated signals are retained separately. MGT convention and CHRR units require verification from instrument documentation. Heat-transfer corrections, crevice flow, variable gas properties, and absolute-pressure calibration are not applied.</p>'''
    (outdir/'report.html').write_text(html,encoding='utf-8')
    return metrics

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title('RCCI Engine Experimental Analysis'); self.geometry('1120x760'); self.minsize(900,650)
        self.vars={}; self.paths={}
        self._build()
        base=Path(__file__).resolve().parent
        for k,n in [('trace','Book1.xlsx'),('cycles','Sat, Sep 26, 2026 _ Load0 Nm_ RPM1527 _ _.xls'),('performance','Performance Report.xls')]:
            p=base/n
            if p.exists(): self.paths[k].set(str(p))
        self.load_perf_defaults()
    def _build(self):
        root=ttk.Frame(self,padding=10); root.pack(fill='both',expand=True)
        files=ttk.LabelFrame(root,text='Input files',padding=8); files.pack(fill='x')
        self.paths={k:tk.StringVar() for k in ['trace','cycles','performance']}
        for i,(k,label) in enumerate([('trace','Instrument workbook (.xlsx)'),('cycles','Consecutive-cycle pressure (.xls/.xlsx)'),('performance','Performance report (.xls/.xlsx)')]):
            ttk.Label(files,text=label,width=34).grid(row=i,column=0,sticky='w',pady=3); ttk.Entry(files,textvariable=self.paths[k]).grid(row=i,column=1,sticky='ew',padx=5); ttk.Button(files,text='Browse',command=lambda x=k:self.browse(x)).grid(row=i,column=2)
        files.columnconfigure(1,weight=1)
        body=ttk.Frame(root); body.pack(fill='both',expand=True,pady=8)
        lf=ttk.LabelFrame(body,text='Operating inputs and optional measurements',padding=8); lf.pack(side='left',fill='y')
        fields=[('rpm','Engine speed (rpm)','1529'),('torque','Brake torque (Nm)','0'),('fuel_kgh','Total fuel flow (kg/h)','0.16'),('air_kgh','Air flow (kg/h)',''),('lhv_mjkg','Fuel blend LHV (MJ/kg)',''),('gamma','Specific heat ratio γ (>1)',''),('displacement_cc','Displacement (cc)','670'),('compression_ratio','Compression ratio','18'),('bore_mm','Bore (mm)','93'),('stroke_mm','Stroke (mm)','98'),('e20_energy_share','E20 energy share (%)','30'),('smooth','Derivative smoothing window (samples; 1=off)','5'),('egt','Exhaust temperature (°C)',''),('nox','NOx (ppm)',''),('co','CO (unit as recorded)',''),('hc','HC (ppm)',''),('co2','CO₂ (%)',''),('smoke','Smoke opacity (%)','')]
        for i,(key,label,default) in enumerate(fields):
            ttk.Label(lf,text=label).grid(row=i,column=0,sticky='w',pady=2); v=tk.StringVar(value=default); self.vars[key]=v; ttk.Entry(lf,textvariable=v,width=16).grid(row=i,column=1,padx=5,pady=2)
        ttk.Label(lf,text='Vehicle: Mahindra Jeeto, 670 cc; RCCI E20/diesel; stated E20 energy share 30%.').grid(row=len(fields),column=0,columnspan=2,sticky='w',pady=8)
        side=ttk.Frame(body); side.pack(side='left',fill='both',expand=True,padx=(10,0))
        ttk.Button(side,text='Analyze and generate report',command=self.run).pack(anchor='w')
        self.status=tk.StringVar(value='Select files and review the input values.'); ttk.Label(side,textvariable=self.status).pack(anchor='w',pady=6)
        self.output=tk.Text(side,wrap='none',font=('Consolas',10)); self.output.pack(fill='both',expand=True)
        ttk.Label(root,text='Missing or blank measurements stay blank. BTE/ITE require a documented LHV; AFR requires measured air and fuel flow; brake outputs require nonzero measured torque.').pack(anchor='w')
    def browse(self,k):
        p=filedialog.askopenfilename(filetypes=[('Data files','*.xls *.xlsx'),('All files','*.*')]);
        if p:self.paths[k].set(p)
    def load_perf_defaults(self):
        p=self.paths['performance'].get()
        if not p:return
        try:
            selected=load_performance_report(p,context_paths=[self.paths.get('cycles',tk.StringVar()).get()],rpm_hint=self.vars.get('rpm',tk.StringVar()).get(),torque_hint=self.vars.get('torque',tk.StringVar()).get())
            for key in ('rpm','torque','fuel_kgh'):
                value=selected.get(key,np.nan)
                if np.isfinite(value) and self.vars[key].get() in ('','1529','0','0.16'):
                    self.vars[key].set(str(value))
        except Exception:pass
    def run(self):
        try:
            tracep=Path(self.paths['trace'].get()); cyc=Path(self.paths['cycles'].get()) if self.paths['cycles'].get() else None
            if not tracep.exists():raise ValueError('Select the instrument workbook.')
            trace=load_instrument(tracep)
            vals={k:v.get() for k,v in self.vars.items()}
            performance=load_performance_report(
                self.paths['performance'].get(),
                context_paths=[self.paths['cycles'].get(),self.paths['trace'].get()],
                rpm_hint=vals.get('rpm',''),torque_hint=vals.get('torque',''))
            # Use the measured operating point from the selected report for
            # calculations, without changing the existing GUI controls.
            for key in ('rpm','torque','fuel_kgh'):
                if np.isfinite(number(performance.get(key))):
                    vals[key]=str(performance[key])
            out=Path(filedialog.askdirectory(title='Choose report output folder') or '')
            if not str(out) or str(out)=='.':return
            metrics=analyze(trace,cyc if cyc and cyc.exists() else None,vals,out,performance)
            self.output.delete('1.0','end')
            for k,v in metrics.items():self.output.insert('end',f'{k}: {v:.6g}\n' if isinstance(v,(float,int,np.number)) and np.isfinite(v) else f'{k}: unavailable\n')
            self.status.set(f'Report and analysis files saved in {out}')
            messagebox.showinfo('Analysis complete',f'Report saved to:\n{out / "report.html"}')
        except Exception as e:
            self.status.set('Analysis failed.'); messagebox.showerror('Analysis error',str(e))

if __name__=='__main__': App().mainloop()
