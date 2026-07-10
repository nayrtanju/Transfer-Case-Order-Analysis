import zipfile
import xml.etree.ElementTree as ET
import re
from typing import Dict, Mapping, Optional, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.signal import savgol_filter

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
G_TO_MS2 = 9.80665

TRANSFER_CASE_TARGET_RPM = np.array([1000,1500,2000,2500,3000,3500,4000,4500], dtype=float)
TRANSFER_CASE_ORDERS = {
    63.0: {"label":"63.00 Order - Gear Mesh","harmonic":"1st","target_rpm":TRANSFER_CASE_TARGET_RPM,"target_amp":np.array([5.0,7.5,10.0,12.5,15.0,17.5,20.0,22.5])},
    85.05: {"label":"85.05 Order - Gear Mesh","harmonic":"1st","target_rpm":TRANSFER_CASE_TARGET_RPM,"target_amp":np.array([2.0,3.0,4.0,5.0,6.0,7.0,8.0,9.0])},
    126.0: {"label":"126.00 Order - 2nd Harmonic","harmonic":"2nd","target_rpm":None,"target_amp":None},
    170.10: {"label":"170.10 Order - 2nd Harmonic","harmonic":"2nd","target_rpm":None,"target_amp":None},
}

def load_shared_strings(z):
    strings=[]
    if "xl/sharedStrings.xml" not in z.namelist(): return strings
    with z.open("xl/sharedStrings.xml") as f:
        for _,e in ET.iterparse(f, events=("end",)):
            if e.tag==NS+"si":
                strings.append("".join((t.text or "") for t in e.iter(NS+"t")))
                e.clear()
    return strings

def col_index(cell_ref):
    m=re.match(r"([A-Z]+)", cell_ref)
    if m is None: raise ValueError(f"Invalid Excel cell reference: {cell_ref}")
    n=0
    for ch in m.group(1): n=n*26+ord(ch)-64
    return n-1

def read_xlsx_numeric(path, max_columns=5):
    with zipfile.ZipFile(path) as z:
        sheet="xl/worksheets/sheet1.xml"
        if sheet not in z.namelist(): raise ValueError("The Excel file does not contain the first worksheet.")
        shared=load_shared_strings(z); headers=[None]*max_columns; rows=[]
        with z.open(sheet) as f:
            for _,row in ET.iterparse(f, events=("end",)):
                if row.tag!=NS+"row": continue
                rnum=int(row.attrib.get("r","0")); vals=[np.nan]*max_columns
                for c in row.findall(NS+"c"):
                    try: j=col_index(c.attrib.get("r",""))
                    except ValueError: continue
                    if j>=max_columns: continue
                    v=c.find(NS+"v")
                    if v is None or v.text is None: continue
                    txt=v.text; typ=c.attrib.get("t")
                    if rnum==1:
                        headers[j]=shared[int(txt)] if typ=="s" and int(txt)<len(shared) else txt
                    else:
                        try: vals[j]=float(txt)
                        except (TypeError,ValueError): vals[j]=np.nan
                if rnum>1: rows.append(vals)
                row.clear()
    if not rows: raise ValueError("No numeric data rows found.")
    return headers, np.asarray(rows,dtype=float)

def repair_time_vector_without_dropping_rows(time):
    time=np.asarray(time,dtype=float)
    d=np.diff(time); pos=d[np.isfinite(d)&(d>0)]
    if len(pos)==0: raise ValueError("Time vector contains no positive increments.")
    md=float(np.median(pos)); d=np.where(np.isfinite(d)&(d>0),d,md)
    out=np.empty_like(time); out[0]=time[0]; out[1:]=time[0]+np.cumsum(d)
    return out

def angular_resample(time,rpm,signal,samples_per_rev=512):
    time=np.asarray(time,float); rpm=np.asarray(rpm,float); signal=np.asarray(signal,float)
    if not(len(time)==len(rpm)==len(signal)): raise ValueError("Time, RPM and signal vectors must have the same length.")
    mask=np.isfinite(time)&np.isfinite(rpm)&np.isfinite(signal)&(rpm>0)
    time,rpm,signal=time[mask],rpm[mask],signal[mask]
    if len(time)<3: raise ValueError("Not enough valid samples.")
    if np.any(np.diff(time)<0): raise ValueError("Time values must not decrease.")
    time=repair_time_vector_without_dropping_rows(time)
    dt=np.diff(time,prepend=time[0]); dt[0]=float(np.median(dt[1:]))
    theta=np.cumsum(2*np.pi*rpm/60.0*dt)
    keep=np.r_[True,np.diff(theta)>0]; theta,rpm,signal=theta[keep],rpm[keep],signal[keep]
    dtheta=2*np.pi/float(samples_per_rev); theta_u=np.arange(theta[0],theta[-1],dtheta)
    if len(theta_u)<samples_per_rev*2: raise ValueError("At least two complete revolutions are required.")
    return theta_u,np.interp(theta_u,theta,signal),np.interp(theta_u,theta,rpm)

def order_map(theta_u,x_u,rpm_u,samples_per_rev=512,revs_per_block=20,overlap=0.75,max_order=200.0):
    theta_u=np.asarray(theta_u,float); x_u=np.asarray(x_u,float); rpm_u=np.asarray(rpm_u,float)
    if not(len(theta_u)==len(x_u)==len(rpm_u)): raise ValueError("theta_u, x_u and rpm_u must have same length.")
    if max_order>samples_per_rev/2: raise ValueError("max_order exceeds angular Nyquist limit.")
    available=len(x_u)/float(samples_per_rev)
    if available<revs_per_block: raise ValueError(f"Insufficient duration: {available:.2f} rev available, {revs_per_block} required.")
    nper=int(samples_per_rev*revs_per_block); hop=max(1,int(round(nper*(1-overlap))))
    win=np.hanning(nper); win_sum=float(np.sum(win))
    all_orders=np.fft.rfftfreq(nper,d=1.0/samples_per_rev); keep=all_orders<=max_order; orders=all_orders[keep]
    specs=[]; rpms=[]
    for start in range(0,len(x_u)-nper+1,hop):
        stop=start+nper; block=x_u[start:stop]; rblock=rpm_u[start:stop]
        if len(block)!=nper or not np.all(np.isfinite(block)) or not np.all(np.isfinite(rblock)): continue
        block=block-np.mean(block); X=np.fft.rfft(block*win)
        amp=np.sqrt(2.0)*np.abs(X)/win_sum
        specs.append(amp[keep]); rpms.append(float(np.mean(rblock)))
    if not specs: raise ValueError("No valid FFT blocks generated.")
    return orders,np.asarray(rpms,float),np.vstack(specs)

def smooth_curve(y,window_length=9,polyorder=2):
    y=np.asarray(y,float)
    if len(y)<5: return y
    if window_length%2==0: window_length+=1
    if window_length>=len(y): window_length=len(y)-1
    if window_length%2==0: window_length-=1
    if window_length<=polyorder or window_length<5: return y
    return savgol_filter(y,window_length=window_length,polyorder=polyorder)

def resample_to_rpm_step(rpm,amp,rpm_step=10):
    rpm=np.asarray(rpm,float); amp=np.asarray(amp,float); mask=np.isfinite(rpm)&np.isfinite(amp)
    rpm,amp=rpm[mask],amp[mask]
    if len(rpm)<2: raise ValueError("Not enough valid RPM blocks.")
    idx=np.argsort(rpm,kind="stable"); rpm,amp=rpm[idx],amp[idx]
    urpm,inv=np.unique(rpm,return_inverse=True)
    if len(urpm)!=len(rpm):
        s=np.zeros(len(urpm)); c=np.zeros(len(urpm)); np.add.at(s,inv,amp); np.add.at(c,inv,1); rpm,amp=urpm,s/np.maximum(c,1)
    rmin=np.ceil(rpm[0]/rpm_step)*rpm_step; rmax=np.floor(rpm[-1]/rpm_step)*rpm_step
    if rmax<=rmin: raise ValueError("RPM range too narrow for selected step.")
    grid=np.arange(rmin,rmax+rpm_step,rpm_step)
    return grid,np.interp(grid,rpm,amp)

def extract_order_vs_rpm(orders,rpms,spec,target_order,width=0.15,rpm_step=10,smooth=True):
    orders=np.asarray(orders,float); rpms=np.asarray(rpms,float); spec=np.asarray(spec,float)
    if spec.ndim!=2: raise ValueError(f"Order spectrum must be 2D, got {spec.shape}.")
    if spec.shape!=(len(rpms),len(orders)): raise ValueError("Spectrum dimensions do not match axes.")
    if target_order<orders[0] or target_order>orders[-1]: raise ValueError(f"Target order {target_order:.2f} outside calculated range.")
    band=(orders>=target_order-width/2)&(orders<=target_order+width/2)
    if np.any(band): amp=np.sqrt(np.sum(spec[:,band]**2,axis=1))
    else: amp=np.asarray([np.interp(target_order,orders,row) for row in spec])
    idx=np.argsort(rpms,kind="stable"); rpms,amp=rpms[idx],amp[idx]
    if smooth: amp=smooth_curve(amp)
    return resample_to_rpm_step(rpms,amp,rpm_step)

def integrate_positive_area(rpm,difference):
    pos=np.maximum(np.asarray(difference,float),0.0); rpm=np.asarray(rpm,float)
    return float(np.trapezoid(pos,rpm)) if hasattr(np,"trapezoid") else float(np.trapz(pos,rpm))

def evaluate_curve_against_target(rpm,amplitude,target_rpm,target_amp):
    i=int(np.argmax(amplitude)); peak_rpm=float(rpm[i]); peak_amp=float(amplitude[i])
    if target_rpm is None or target_amp is None:
        return {"Peak RPM":peak_rpm,"Peak Amplitude [m/s²]":peak_amp,"Target at Peak RPM [m/s²]":np.nan,"Max Margin [m/s²]":np.nan,"Max Margin [%]":np.nan,"Exceedance Area [m/s²·RPM]":np.nan,"Status":"INFO"}
    target=np.interp(rpm,target_rpm,target_amp); margin=amplitude-target; j=int(np.argmax(margin)); area=integrate_positive_area(rpm,margin)
    denom=float(target[j]); pct=float(margin[j]/denom*100.0) if denom>0 else np.nan
    return {"Peak RPM":peak_rpm,"Peak Amplitude [m/s²]":peak_amp,"Target at Peak RPM [m/s²]":float(np.interp(peak_rpm,target_rpm,target_amp)),"Max Margin [m/s²]":float(margin[j]),"Max Margin [%]":pct,"Exceedance Area [m/s²·RPM]":area,"Status":"PASS" if area<=1e-9 else "FAIL"}

def analyze_transfer_case_orders(time,rpm,channels,order_definitions=None,samples_per_rev=512,revs_per_block=20,overlap=0.75,max_order=200.0,order_width=0.15,rpm_step=10,calibration_factor=1.0):
    defs=TRANSFER_CASE_ORDERS if order_definitions is None else order_definitions
    if max_order<max(defs): raise ValueError(f"max_order must be at least {max(defs):.2f}.")
    curves={float(o):{} for o in defs}; rows={float(o):[] for o in defs}
    for ch,signal in channels.items():
        th,xu,ru=angular_resample(time,rpm,signal,samples_per_rev)
        orders,brpm,spec=order_map(th,xu,ru,samples_per_rev,revs_per_block,overlap,max_order)
        for o,d in defs.items():
            r,a=extract_order_vs_rpm(orders,brpm,spec,float(o),order_width,rpm_step,True); a=a*calibration_factor
            curves[float(o)][ch]={"rpm":r,"amp":a}
            rows[float(o)].append({"Order":float(o),"Order Label":d["label"],"Harmonic":d["harmonic"],"Channel":ch,**evaluate_curve_against_target(r,a,d.get("target_rpm"),d.get("target_amp"))})
    results={}; raw={}
    for o,d in defs.items():
        o=float(o); results[o]=pd.DataFrame(rows[o]); base=next(iter(curves[o].values()))["rpm"]; df=pd.DataFrame({"RPM":base})
        for ch,c in curves[o].items(): df[ch]=np.interp(base,c["rpm"],c["amp"])
        if d.get("target_rpm") is not None: df["Target"]=np.interp(base,d["target_rpm"],d["target_amp"])
        raw[o]=df
    return curves,results,raw

def plot_order_map(orders,rpms,spec,channel_name="Channel",db_reference=1.0):
    orders=np.asarray(orders,float); rpms=np.asarray(rpms,float); spec=np.asarray(spec,float)
    if spec.ndim!=2 or spec.shape!=(len(rpms),len(orders)): raise ValueError("Invalid order map dimensions.")
    idx=np.argsort(rpms); r=rpms[idx]; s=spec[idx,:]; db=20*np.log10(np.maximum(s,1e-12)/db_reference)
    fig,ax=plt.subplots(figsize=(12,7)); im=ax.imshow(db,aspect="auto",origin="lower",extent=[orders[0],orders[-1],r[0],r[-1]],interpolation="nearest",cmap="jet")
    fig.colorbar(im,ax=ax,label="Amplitude [dB re 1 m/s²]"); ax.set_xlabel("Order"); ax.set_ylabel("RPM"); ax.set_title(f"Transfer Case Order Map - {channel_name}")
    return fig

