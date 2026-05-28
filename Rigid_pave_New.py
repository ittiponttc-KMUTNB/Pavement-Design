"""
Rigid Pavement Design V7
AASHTO 1993 — JPCP/JRCP & CRCP
พัฒนาโดย รศ.ดร.อิทธิพล มีผล
ภาควิชาครุศาสตร์โยธา มจพ.
"""

import streamlit as st
import math, json, io
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from scipy.interpolate import interp1d
from io import BytesIO
from datetime import datetime

# ============================================================
# ENGINE IMPORT
# ============================================================
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

# ── ESAL ────────────────────────────────────────────────────
_TON_TO_KIP = 2.2046
_VEHICLE_AXLES = {
    'MB':  [(4,  1, 1), (11, 1, 1)],
    'HB':  [(5,  1, 1), (20, 2, 1)],
    'MT':  [(4,  1, 1), (11, 1, 1)],
    'HT':  [(5,  1, 1), (20, 2, 1)],
    'TR':  [(5,  1, 1), (20, 2, 1), (11, 1, 1), (11, 1, 1)],
    'STR': [(5,  1, 1), (20, 2, 1), (20, 2, 1)],
}

def _ealf_rigid(L1_ton, L2, D_in, pt):
    L1  = L1_ton * _TON_TO_KIP
    Gt  = math.log10((4.5 - pt) / (4.5 - 1.5))
    Bx  = 1.0 + 3.63*(L1+L2)**5.20/((D_in+1)**8.46*L2**3.52)
    B18 = 1.0 + 3.63*(18+1)**5.20/((D_in+1)**8.46*1.0**3.52)
    return 10**(4.62*math.log10(L1+L2)-3.28*math.log10(L2)
                -4.62*math.log10(19)+Gt*(1/B18-1/Bx))

def compute_esal_for_d(traffic_data, pt, lane_factor, direction_factor, d_cm):
    d_inch = round(d_cm / 2.54)
    tf = {code: sum(_ealf_rigid(L1,L2,d_inch,pt)*cnt for L1,L2,cnt in axles)
          for code, axles in _VEHICLE_AXLES.items()}
    acc = sum(row.get(code,0)*tf[code]*lane_factor*direction_factor*365
              for row in traffic_data for code in tf)
    return (round(acc), d_inch, tf)

# ── Fig.3.3 ─────────────────────────────────────────────────
_F33_X_MIN, _F33_X_MAX = 89.198802, 731.269461
_F33_Y_MIN, _F33_Y_MAX = 13.661078, 658.142515

_F33_RAW = {
    'Turning Line': [(731.269,657.339),(409.029,336.705)],
    'Dsb': [(89.787,337.504),(129.069,337.504),(169.153,337.504),
            (209.237,337.504),(249.320,336.703),(289.404,336.703),
            (329.487,336.703),(369.571,336.703)],
    'Esb 15000':   [(90.002,125.360),(129.378,135.807),(169.558,148.665),(209.738,161.522),
                    (249.114,177.594),(289.293,196.077),(329.473,218.577),(368.849,245.096),(408.225,281.257)],
    'Esb 30000':   [(89.199,109.289),(129.378,121.343),(169.558,134.200),(209.738,148.665),
                    (249.114,165.540),(289.293,184.826),(329.473,208.131),(368.849,236.256),(409.029,274.025)],
    'Esb 50000':   [(89.199,97.235),(129.378,109.289),(169.558,122.950),(209.738,137.414),
                    (249.114,155.897),(289.293,175.987),(329.473,200.095),(368.046,229.828),(409.029,270.811)],
    'Esb 75000':   [(89.199,87.592),(128.575,100.449),(168.754,114.914),(209.738,130.182),
                    (249.114,148.665),(289.293,169.558),(329.473,195.273),(368.849,225.006),(409.832,265.989)],
    'Esb 100,000': [(89.199,81.163),(129.378,94.824),(169.558,109.289),(209.738,125.360),
                    (249.917,143.843),(289.293,165.540),(329.473,190.451),(368.046,220.184),(409.029,262.775)],
    'Esb 200,000': [(88.395,65.091),(129.378,77.949),(169.558,93.217),(209.738,110.092),
                    (249.114,129.378),(289.293,151.879),(329.473,180.005),(368.046,212.952),(409.029,260.364)],
    'Esb 400,000': [(89.199,49.823),(128.575,65.091),(168.754,80.359),(209.738,98.038),
                    (249.114,118.128),(288.490,141.432),(309.383,155.093),(329.473,170.362),
                    (350.366,188.041),(368.046,204.916),(389.743,227.417),(409.029,252.328)],
    'Esb 600,000': [(88.395,39.376),(128.575,54.644),(168.754,71.520),(209.738,90.002),
                    (229.828,99.646),(249.917,110.896),(270.007,122.146),(289.293,135.004),
                    (309.785,149.468),(329.473,163.933),(351.170,182.416),(368.849,199.291),
                    (390.546,223.399),(409.431,248.712)],
    'Esb 1,000,000':[(89.199,27.322),(128.575,44.198),(168.754,61.073),(208.934,79.556),
                     (249.114,101.253),(270.007,114.110),(288.490,126.164),(310.187,141.432),
                     (328.669,156.701),(350.366,175.183),(371.260,196.880),(393.760,222.595),(409.029,243.489)],
    'Mr 1000':  [(89.787,351.934),(129.069,359.149),(169.955,366.364),(209.237,374.381),
                 (249.320,384.001),(270.007,390.546),(289.293,396.573),(310.990,404.207),
                 (330.277,411.258),(350.366,421.083),(368.849,429.697),(389.743,441.976),(409.655,454.548)],
    'Mr 2000':  [(90.002,384.921),(129.378,392.957),(169.558,401.796),(191.255,406.618),
                 (209.738,410.636),(229.024,416.261),(249.114,421.886),(270.007,428.315),
                 (289.293,435.547),(310.990,443.583),(329.473,450.816),(350.366,461.262),
                 (368.849,471.709),(389.743,483.763),(409.029,498.228)],
    'Mr 3000':  [(90.002,405.011),(129.378,412.243),(169.558,421.886),(209.738,431.529),
                 (229.024,437.154),(249.917,442.780),(270.007,449.208),(289.293,456.441),
                 (310.990,464.477),(330.277,472.513),(350.366,482.557),(369.653,493.406),(409.029,523.943)],
    'Mr 5000':  [(90.002,429.119),(129.378,437.154),(151.075,442.378),(169.558,446.798),
                 (190.853,451.619),(208.934,456.441),(229.828,462.066),(249.917,468.495),
                 (270.007,475.727),(289.293,483.763),(310.990,492.602),(329.473,502.246),
                 (350.366,513.496),(368.849,524.746),(389.743,540.014),(409.832,556.086)],
    'Mr 7000':  [(90.002,445.190),(128.977,453.628),(150.272,458.851),(169.558,463.673),
                 (190.451,469.298),(209.738,474.120),(229.828,480.950),(249.516,486.977),
                 (270.007,494.210),(289.293,501.844),(310.990,511.085),(329.473,519.925),
                 (350.768,531.978),(369.653,544.032),(390.546,559.301),(409.832,576.980)],
    'Mr 10,000':[(90.002,462.869),(129.378,470.905),(169.960,480.549),(209.738,490.995),
                 (229.024,496.620),(249.917,503.853),(270.007,511.889),(289.293,520.728),
                 (310.990,530.371),(330.277,540.416),(351.170,552.872),(369.653,564.926),
                 (390.546,580.998),(409.832,598.677)],
    'Mr 12,000':[(90.002,470.102),(129.378,479.745),(152.683,485.370),(178.398,492.602),
                 (220.988,504.656),(238.667,510.683),(278.043,526.353),(296.526,534.389),
                 (320.634,544.836),(338.313,555.283),(358.402,568.140),(379.296,583.408),
                 (400.993,602.695),(410.636,612.338)],
    'Mr 16,000':[(90.002,484.566),(129.780,493.004),(169.558,503.853),(208.934,514.299),
                 (235.453,523.943),(264.382,535.193),(295.722,548.854),(321.437,560.908),
                 (340.723,572.158),(361.617,585.016),(381.707,601.087),(396.171,614.748),(409.832,628.811)],
    'Mr 20,000':[(90.002,493.406),(130.182,503.853),(159.111,511.889),(194.469,522.335),
                 (216.166,528.764),(255.543,542.425),(274.829,550.461),(298.937,562.515),
                 (319.026,572.962),(340.723,585.016),(364.831,601.087),(378.492,611.534),
                 (399.386,630.419),(409.832,643.276)],
    'K(pci) 50 Q2':[(570.551,337.509),(409.832,175.987)],
    'K 100': [(409.832,128.575),(617.561,337.509)],
    'K 200': [(409.029,79.556),(666.178,336.705)],
    'K 300': [(409.029,53.439),(692.697,337.107)],
    'K 400': [(409.029,31.742),(712.787,336.705)],
    'K 500': [(409.832,14.465),(731.269,336.705)],
    'K 600': [(421.886,13.661),(731.269,323.044)],
    'K 800': [(441.172,14.063),(730.466,305.365)],
    'K 1000':[(457.244,13.661),(731.269,289.293)],
    'K 1500':[(485.370,13.661),(731.269,259.560)],
    'K 2000':[(505.460,14.465),(731.269,240.274)],
}

def _f33_norm(pts):
    return [((x-_F33_X_MIN)/(_F33_X_MAX-_F33_X_MIN),
             1-(y-_F33_Y_MIN)/(_F33_Y_MAX-_F33_Y_MIN)) for x,y in pts]
def _f33_sx(pts): return sorted(pts, key=lambda p: p[0])

_F33_DSB_TICKS  = [20,18,16,14,12,10,8,6]
_F33_DSB_NORM_X = sorted([p[0] for p in _f33_norm(_F33_RAW['Dsb'])])
_F33_D_ASC      = np.array(_F33_DSB_TICKS[::-1])
_F33_X_ASC      = np.array(_F33_DSB_NORM_X[::-1])

_F33_ESB = {15000:_f33_sx(_f33_norm(_F33_RAW['Esb 15000'])),
            30000:_f33_sx(_f33_norm(_F33_RAW['Esb 30000'])),
            50000:_f33_sx(_f33_norm(_F33_RAW['Esb 50000'])),
            75000:_f33_sx(_f33_norm(_F33_RAW['Esb 75000'])),
            100000:_f33_sx(_f33_norm(_F33_RAW['Esb 100,000'])),
            200000:_f33_sx(_f33_norm(_F33_RAW['Esb 200,000'])),
            400000:_f33_sx(_f33_norm(_F33_RAW['Esb 400,000'])),
            600000:_f33_sx(_f33_norm(_F33_RAW['Esb 600,000'])),
            1000000:_f33_sx(_f33_norm(_F33_RAW['Esb 1,000,000']))}
_F33_MR  = {1000:_f33_sx(_f33_norm(_F33_RAW['Mr 1000'])),
            2000:_f33_sx(_f33_norm(_F33_RAW['Mr 2000'])),
            3000:_f33_sx(_f33_norm(_F33_RAW['Mr 3000'])),
            5000:_f33_sx(_f33_norm(_F33_RAW['Mr 5000'])),
            7000:_f33_sx(_f33_norm(_F33_RAW['Mr 7000'])),
            10000:_f33_sx(_f33_norm(_F33_RAW['Mr 10,000'])),
            12000:_f33_sx(_f33_norm(_F33_RAW['Mr 12,000'])),
            16000:_f33_sx(_f33_norm(_F33_RAW['Mr 16,000'])),
            20000:_f33_sx(_f33_norm(_F33_RAW['Mr 20,000']))}
_F33_K   = {50:_f33_norm(_F33_RAW['K(pci) 50 Q2']),
            100:_f33_norm(_F33_RAW['K 100']),200:_f33_norm(_F33_RAW['K 200']),
            300:_f33_norm(_F33_RAW['K 300']),400:_f33_norm(_F33_RAW['K 400']),
            500:_f33_norm(_F33_RAW['K 500']),600:_f33_norm(_F33_RAW['K 600']),
            800:_f33_norm(_F33_RAW['K 800']),1000:_f33_norm(_F33_RAW['K 1000']),
            1500:_f33_norm(_F33_RAW['K 1500']),2000:_f33_norm(_F33_RAW['K 2000'])}

_F33_TL   = _f33_norm(_F33_RAW['Turning Line'])
_F33_TLX  = np.array([p[0] for p in _F33_TL])
_F33_TLY  = np.array([p[1] for p in _F33_TL])
_F33_TLX[1], _F33_TLY[1] = 0.5, 0.5
_F33_TLS  = (_F33_TLY[1]-_F33_TLY[0])/(_F33_TLX[1]-_F33_TLX[0])

def _f33_yon(pts, x):
    xs=np.array([p[0] for p in pts]); ys=np.array([p[1] for p in pts])
    return float(np.interp(x,xs,ys))

def _f33_ibw(dd, val, x):
    ks=sorted(dd.keys())
    lo=max([k for k in ks if k<=val],default=ks[0])
    hi=min([k for k in ks if k>=val],default=ks[-1])
    if lo==hi: return _f33_yon(dd[lo],x)
    return _f33_yon(dd[lo],x)+((val-lo)/(hi-lo))*(_f33_yon(dd[hi],x)-_f33_yon(dd[lo],x))

def calc_composite_k(MR_psi, ESB_psi, DSB_in):
    MR_psi  = float(np.clip(MR_psi,  1000, 20000))
    ESB_psi = float(np.clip(ESB_psi, 15000,1000000))
    DSB_in  = float(np.clip(DSB_in,  6.0,  20.0))
    x_dsb = float(np.interp(DSB_in, _F33_D_ASC, _F33_X_ASC))
    y_A   = _f33_ibw(_F33_MR,  MR_psi,  x_dsb)
    y_B   = _f33_ibw(_F33_ESB, ESB_psi, x_dsb)
    x_D   = _F33_TLX[0]+(y_A-_F33_TLY[0])/_F33_TLS
    x_C, y_C = x_D, y_B
    ks    = sorted(_F33_K.keys())
    ky    = []
    for v in ks:
        ps=sorted(_F33_K[v],key=lambda p:p[0])
        x0,y0=ps[0]; x1,y1=ps[-1]
        ky.append(float(np.interp(x_C,[x0,x1],[y0,y1])))
    ky=np.array(ky)
    il=int(np.clip(np.searchsorted(ky,y_C)-1,0,len(ks)-2))
    ih=il+1
    frac=(y_C-ky[il])/(ky[ih]-ky[il]) if ky[ih]!=ky[il] else 0
    k_inf=ks[il]+frac*(ks[ih]-ks[il])
    return {'k_inf_pci':round(float(k_inf),0),
            'MR_psi':MR_psi,'ESB_psi':ESB_psi,'DSB_in':DSB_in,
            'x_dsb':x_dsb,'y_A':y_A,'y_B':y_B,'x_C':x_C,'y_C':y_C,
            'x_D':x_D,'y_D':y_A}

# ── Fig.3.4 LS ──────────────────────────────────────────────
_F34_LS_RAW = {
    0:[(0.0012,0.0),(0.8952,1.0)],
    1:[(0.0265,0.0),(0.9988,0.9017)],
    2:[(0.0422,0.0),(0.9988,0.6821)],
    3:[(0.1036,0.0007),(1.0,0.5330)],
}
_F34_LXMN,_F34_LXMX = np.log10(1),np.log10(2000)
_F34_LYMN,_F34_LYMX = np.log10(1),np.log10(1000)

def _f34_xp(n): return 10**(_F34_LXMN+n*(_F34_LXMX-_F34_LXMN))
def _f34_yp(n): return 10**(_F34_LYMN+n*(_F34_LYMX-_F34_LYMN))

_F34_LS_PCI = {ls:(_f34_xp(p[0][0]),_f34_yp(p[0][1]),
                   _f34_xp(p[1][0]),_f34_yp(p[1][1]))
               for ls,p in _F34_LS_RAW.items()}

def apply_loss_of_support(k_inf_pci, ls):
    if ls <= 0: return float(k_inf_pci)
    k_inf_pci = float(np.clip(k_inf_pci, 1, 2000))
    ks = sorted(_F34_LS_PCI.keys())
    def _keff(ls_i, ki):
        x1,y1,x2,y2 = _F34_LS_PCI[ls_i]
        f = interp1d(np.log10([x1,x2]),np.log10([y1,y2]),fill_value='extrapolate')
        return float(10**f(np.log10(ki)))
    lo=max([k for k in ks if k<=ls],default=0)
    hi=min([k for k in ks if k>=ls],default=3)
    if lo==hi: return _keff(lo, k_inf_pci)
    ylo=_keff(lo,k_inf_pci); yhi=_keff(hi,k_inf_pci)
    frac=(ls-lo)/(hi-lo)
    return float(np.clip(ylo+frac*(yhi-ylo),1,k_inf_pci))

# ── Odemark ─────────────────────────────────────────────────
def calc_odemark(layers):
    """layers: list of (thickness_cm, E_MPa)"""
    valid = [(t,e) for t,e in layers if t>0 and e>0]
    if not valid: return None
    D = [t*0.393701 for t,e in valid]
    E = [e*145.038  for t,e in valid]
    Dt = sum(D)
    Eq = (sum(d*(e**(1/3)) for d,e in zip(D,E))/Dt)**3
    return (round(Dt,2), round(Eq,0))

# ── AASHTO Design ────────────────────────────────────────────
ZR_TABLE = {50:-0.000,60:-0.253,70:-0.524,75:-0.674,80:-0.841,
            85:-1.037,90:-1.282,91:-1.340,92:-1.405,93:-1.476,
            94:-1.555,95:-1.645,96:-1.751,97:-1.881,98:-2.054,99:-2.327}
J_DEFAULT = {'JPCP/JRCP':2.8,'CRCP':2.6}

def convert_cube_to_cyl(fc_cube): return 0.8*fc_cube
def calc_ec(fc_cyl): return 57000*math.sqrt(fc_cyl*14.223)
def calc_sc(fc_cyl): return 10.0*math.sqrt(fc_cyl*14.223)
def get_zr(r): return ZR_TABLE.get(int(r),-1.282)
def mr_from_cbr(cbr): return 1500*cbr if cbr<10 else 1000+555*cbr

def calc_w18(d_in, dpsi, pt, zr, so, sc, cd, j, ec, k):
    t1 = zr*so
    t2 = 7.35*math.log10(d_in+1)-0.06
    t3 = math.log10(dpsi/3.0)/(1+1.624e7/(d_in+1)**8.46)
    dp = d_in**0.75
    num4 = sc*cd*(dp-1.132)
    den4 = 215.63*j*(dp-18.42/(ec/k)**0.25)
    if num4<=0 or den4<=0: return (float('-inf'),0)
    inn = num4/den4
    if inn<=0: return (float('-inf'),0)
    t4 = (4.22-0.32*pt)*math.log10(inn)
    lw = t1+t2+t3+t4
    return (lw, 10**lw)

def check_design(w18_req, w18_cap):
    ratio = w18_cap/w18_req if w18_req>0 else float('inf')
    return (w18_cap>=w18_req, ratio)

def find_optimum_k(w18_req, d_in, dpsi, pt, zr, so, sc, cd, j, ec):
    for k in range(50,1010,10):
        _,w = calc_w18(d_in,dpsi,pt,zr,so,sc,cd,j,ec,k)
        if w>=w18_req: return k
    return None

def compare_d(d_list_cm, dpsi, pt, zr, so, sc, cd, j, ec, k, w18_req):
    rows=[]
    for dcm in d_list_cm:
        din=round(dcm/2.54)
        lw,wc=calc_w18(din,dpsi,pt,zr,so,sc,cd,j,ec,k)
        p,r=check_design(w18_req,wc)
        rows.append({'d_cm':dcm,'d_inch':din,'log_w18':round(lw,4),
                     'w18_cap':round(wc,0),'passed':p,'ratio':round(r,3)})
    return rows

# ── Material Library ─────────────────────────────────────────
MATERIAL_MODULUS = {
    "รองผิวทางคอนกรีตด้วย AC":2500,
    "รองผิวทางคอนกรีตด้วย PMA(AC)":3700,
    "หินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (CTB)":1200,
    "หินคลุกผสมซีเมนต์ UCS 24.5 ksc":850,
    "หินคลุก CBR 80%":350,
    "ดินซีเมนต์ UCS 17.5 ksc":350,
    "วัสดุหมุนเวียน (Recycling)":850,
    "รองพื้นทางวัสดุมวลรวม CBR 25%":150,
    "วัสดุคัดเลือก ก":100,
    "ดินถมคันทาง / ดินเดิม":100,
    "กำหนดเอง...":100,
}
LAYER_COLORS = {
    "รองผิวทางคอนกรีตด้วย AC":"#2C3E50",
    "รองผิวทางคอนกรีตด้วย PMA(AC)":"#1A252F",
    "หินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (CTB)":"#7F8C8D",
    "หินคลุกผสมซีเมนต์ UCS 24.5 ksc":"#95A5A6",
    "หินคลุก CBR 80%":"#BDC3C7",
    "ดินซีเมนต์ UCS 17.5 ksc":"#AAB7B8",
    "วัสดุหมุนเวียน (Recycling)":"#85929E",
    "รองพื้นทางวัสดุมวลรวม CBR 25%":"#FFCC99",
    "วัสดุคัดเลือก ก":"#E8DAEF",
    "ดินถมคันทาง / ดินเดิม":"#F5CBA7",
    "กำหนดเอง...":"#FADBD8",
    "Concrete Slab":"#808080",
}
LAYER_NAMES_EN = {
    "รองผิวทางคอนกรีตด้วย AC":"AC Interlayer",
    "รองผิวทางคอนกรีตด้วย PMA(AC)":"PMA Interlayer",
    "หินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (CTB)":"Cement Treated Base",
    "หินคลุกผสมซีเมนต์ UCS 24.5 ksc":"Mod. Crushed Rock",
    "หินคลุก CBR 80%":"Crushed Rock Base",
    "ดินซีเมนต์ UCS 17.5 ksc":"Soil Cement",
    "วัสดุหมุนเวียน (Recycling)":"Recycled Material",
    "รองพื้นทางวัสดุมวลรวม CBR 25%":"Aggregate Subbase",
    "วัสดุคัดเลือก ก":"Selected Material",
    "ดินถมคันทาง / ดินเดิม":"Subgrade",
    "กำหนดเอง...":"Custom Material",
}
D_LIST_CM = [20,22,25,28,30,32,35,38,40]

# ============================================================
# CSS
# ============================================================
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=Sarabun:wght@300;400;600;700&display=swap');

[data-testid="stAppViewContainer"] { background: #F0F4FA; }
[data-testid="stHeader"] { background: transparent; }

.rp-header {
    background: #1565C0;
    border-radius: 10px;
    padding: 16px 20px 12px;
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.rp-header-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 15px;
    color: #FFFFFF;
    font-weight: 600;
    letter-spacing: 0.05em;
}
.rp-header-sub {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    color: #BBDEFB;
    margin-top: 3px;
}
.rp-badge {
    background: rgba(255,255,255,0.15);
    border: 1px solid rgba(255,255,255,0.3);
    border-radius: 6px;
    padding: 4px 12px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    color: #FFFFFF;
}
.rp-card {
    background: #FFF8E1;
    border: 1px solid #FFECB3;
    border-radius: 8px;
    padding: 8px 12px;
    margin-bottom: 6px;
}
.rp-card-title {
    font-size: 13px;
    font-weight: 600;
    color: #1565C0;
    margin-bottom: 12px;
    padding-bottom: 8px;
    border-bottom: 1px solid #E3F2FD;
    display: flex;
    align-items: center;
    gap: 6px;
}
.rp-section-label {
    font-size: 11px;
    font-weight: 600;
    color: #90A4AE;
    letter-spacing: 0.07em;
    text-transform: uppercase;
    margin-bottom: 8px;
    margin-top: 10px;
}
.rp-metric {
    background: #F5F9FE;
    border: 1px solid #BBDEFB;
    border-radius: 8px;
    padding: 10px 12px;
    text-align: center;
}
.rp-metric-label {
    font-size: 11px;
    color: #90A4AE;
    margin-bottom: 4px;
}
.rp-metric-val {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 20px;
    font-weight: 600;
    color: #1565C0;
}
.rp-metric-val.pass { color: #2E7D32; }
.rp-metric-val.fail { color: #C62828; }
.rp-status-ok {
    background: #E8F5E9;
    border: 1px solid #A5D6A7;
    border-radius: 8px;
    padding: 8px 14px;
    font-size: 13px;
    color: #2E7D32;
    font-weight: 600;
}
.rp-status-warn {
    background: #FFF8E1;
    border: 1px solid #FFD54F;
    border-radius: 8px;
    padding: 8px 14px;
    font-size: 13px;
    color: #E65100;
    font-weight: 600;
}
.rp-status-info {
    background: #E3F2FD;
    border: 1px solid #90CAF9;
    border-radius: 8px;
    padding: 8px 14px;
    font-size: 13px;
    color: #1565C0;
}
.rp-pave-card {
    background: #F5F9FE;
    border: 1px solid #90CAF9;
    border-radius: 12px;
    padding: 14px 16px;
}
.rp-pave-title {
    font-size: 13px;
    font-weight: 600;
    color: #1565C0;
    margin-bottom: 10px;
}
.rp-result-row {
    display: flex;
    justify-content: space-between;
    padding: 5px 0;
    border-bottom: 1px solid #E3F2FD;
    font-size: 13px;
}
.rp-result-label { color: #90A4AE; }
.rp-result-val {
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    color: #1A237E;
}
.rp-layer-row {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 4px 0;
    font-size: 12px;
    color: #37474F;
}
.badge-pass {
    background: #E8F5E9; color: #2E7D32;
    border: 1px solid #A5D6A7;
    border-radius: 5px; padding: 2px 8px;
    font-size: 12px; font-weight: 600;
}
.badge-fail {
    background: #FFEBEE; color: #C62828;
    border: 1px solid #EF9A9A;
    border-radius: 5px; padding: 2px 8px;
    font-size: 12px; font-weight: 600;
}
div[data-testid="stTabs"] button {
    font-weight: 600 !important;
    color: #546E7A !important;
}
div[data-testid="stTabs"] button[aria-selected="true"] {
    color: #1565C0 !important;
    border-bottom: 3px solid #1565C0 !important;
}
</style>
"""

# ============================================================
# FIGURE FUNCTIONS
# ============================================================

def plot_nomograph_f33(MR_psi, ESB_psi, DSB_in, k_inf_pci,
                       x_dsb, y_A, y_B, x_C, y_C, x_D, y_D):
    """วาด Fig.3.3 พร้อมเส้นแดง"""
    fig, ax = plt.subplots(figsize=(9,9))
    fig.patch.set_facecolor('white'); ax.set_facecolor('white')
    ax.set_xlim(0,1); ax.set_ylim(0,1)
    ax.set_xticks([]); ax.set_yticks([])
    for sp in ax.spines.values(): sp.set_visible(False)
    ax.add_patch(plt.Rectangle((0,0),1,1,lw=2.5,ec='black',fc='none'))
    ax.axhline(0.5,color='black',lw=2.0)
    ax.axvline(0.5,color='black',lw=2.0)
    ax.plot([_F33_TLX[0],_F33_TLX[1]],[_F33_TLY[0],_F33_TLY[1]],color='black',lw=1.2)
    ax.text(0.74,0.17,'(Turning Line)',fontsize=6.5,fontstyle='italic',rotation=-47)
    esb_labels={15000:'15,000',30000:'30,000',50000:'50,000',75000:'75,000',
                100000:'100,000',200000:'200,000',400000:'400,000',
                600000:'600,000',1000000:'1,000,000'}
    for val,pts in _F33_ESB.items():
        ax.plot([p[0] for p in pts],[p[1] for p in pts],color='black',lw=0.8)
        ax.text(-0.01,pts[0][1],esb_labels[val],fontsize=5.5,ha='right',va='center')
    mr_labels={1000:'1,000',2000:'2,000',3000:'3,000',5000:'5,000',
               7000:'7,000',10000:'10,000',12000:'12,000',16000:'16,000',20000:'20,000'}
    for val,pts in _F33_MR.items():
        ax.plot([p[0] for p in pts],[p[1] for p in pts],color='black',lw=0.8)
        ax.text(-0.01,pts[0][1],mr_labels[val],fontsize=5.5,ha='right',va='center')
    k_labels={50:'50',100:'100',200:'200',300:'300',400:'400',500:'500',
              600:'600',800:'800',1000:'1000',1500:'1500',2000:'2000'}
    for val,pts in _F33_K.items():
        ps=sorted(pts,key=lambda p:p[0])
        x0,y0=ps[0]; x1,y1=ps[-1]
        ax.plot([x0,x1],[y0,y1],color='black',lw=0.8)
        xm,ym=(x0+x1)/2,(y0+y1)/2
        ang=np.degrees(np.arctan2(y1-y0,x1-x0))
        ax.text(xm,ym,k_labels[val],fontsize=5.5,ha='center',va='center',
                rotation=ang,bbox=dict(fc='white',ec='none',pad=0.8))
    dsb_ticks=[20,18,16,14,12,10,8,6]
    for i,d in enumerate(dsb_ticks):
        x=_F33_DSB_NORM_X[i]
        ax.plot([x,x],[0.5,0.513],color='black',lw=0.8)
        ax.plot([x,x],[0.5,0.487],color='black',lw=0.8)
        ax.text(x,0.517,str(d),ha='center',va='bottom',fontsize=7)
    ax.text(0.20,0.545,'Subbase Thickness, $D_{SB}$ (inches)',ha='center',va='bottom',fontsize=8)
    ax.text(0.47,0.97,'Subbase Elastic\nModulus, $E_{SB}$ (psi)',
            ha='right',va='top',fontsize=8,style='italic',
            bbox=dict(fc='white',ec='none',pad=1))
    ax.text(0.98,0.97,'Composite Modulus of\nSubgrade Reaction,\n$k_\\infty$ (pci)',
            ha='right',va='top',fontsize=7,style='italic',
            bbox=dict(fc='white',ec='none',pad=1))
    ax.text(0.13,0.05,'Roadbed Soil\nResilient Modulus,\n$M_R$ (psi)',
            ha='center',va='bottom',fontsize=8,style='italic',
            bbox=dict(fc='white',ec='none',pad=1))
    # เส้นแดง
    ax.plot([x_dsb,x_dsb],[y_A,y_B],'r-',lw=1.8)
    ax.plot([x_dsb,x_D],[y_A,y_D],'r-',lw=1.8)
    ax.plot([x_dsb,x_C],[y_B,y_C],'r-',lw=1.8)
    ax.plot([x_D,x_C],[y_D,y_C],'r-',lw=1.8)
    ax.plot(x_C,y_C,'ro',markersize=8,zorder=5)
    ax.annotate(f'$k_\\infty$ = {k_inf_pci:.0f} pci',
                xy=(x_C,y_C),xytext=(x_C+0.07,y_C-0.06),
                fontsize=9,color='red',fontweight='bold',
                arrowprops=dict(arrowstyle='->',color='red',lw=1.2))
    ax.set_title(f'$M_R$={MR_psi:,.0f} psi   $D_{{SB}}$={DSB_in} in   '
                 f'$E_{{SB}}$={ESB_psi:,.0f} psi   →   $k_\\infty$ = {k_inf_pci:.0f} pci',
                 fontsize=9,color='red')
    plt.tight_layout()
    return fig

def plot_nomograph_f34(k_inf_pci, ls, k_eff_pci):
    """วาด Fig.3.4 พร้อมเส้นแดง"""
    fig, ax = plt.subplots(figsize=(8,7))
    fig.patch.set_facecolor('white'); ax.set_facecolor('white')
    ax.set_xscale('log'); ax.set_yscale('log')
    ax.set_xlim(1,2000); ax.set_ylim(1,1000)
    xticks=[1,2,5,10,20,50,100,200,500,1000,2000]
    yticks=[1,2,5,10,20,50,100,200,500,1000]
    ax.set_xticks(xticks); ax.set_xticklabels([str(x) for x in xticks],fontsize=8)
    ax.set_yticks(yticks); ax.set_yticklabels([str(y) for y in yticks],fontsize=8)
    ax.grid(True,which='both',color='#cccccc',lw=0.5)
    ax.grid(True,which='minor',color='#eeeeee',lw=0.3)
    ls_colors={0:'#1a237e',1:'#1565C0',2:'#0288d1',3:'#4fc3f7'}
    ls_labels={0:'LS = 0',1:'LS = 1.0',2:'LS = 2.0',3:'LS = 3.0'}
    lx_mid={0:30,1:50,2:80,3:130}
    for lsi,(x1,y1,x2,y2) in _F34_LS_PCI.items():
        xs=np.logspace(np.log10(x1),np.log10(x2),100)
        f=interp1d(np.log10([x1,x2]),np.log10([y1,y2]),fill_value='extrapolate')
        ys=10**f(np.log10(xs))
        ax.plot(xs,ys,color=ls_colors[lsi],lw=2.0)
        xm=lx_mid[lsi]; ym=float(10**f(np.log10(xm)))
        dx=0.3
        pa=ax.transData.transform((_f34_xp(np.log10(xm*10**(-dx))/(_F34_LXMX-_F34_LXMN)),ym))
        pb=ax.transData.transform((_f34_xp(np.log10(xm*10**(dx))/(_F34_LXMX-_F34_LXMN)),ym))
        xa=10**(np.log10(xm)-dx); xb=10**(np.log10(xm)+dx)
        ya=float(10**f(np.log10(xa))); yb=float(10**f(np.log10(xb)))
        pa=ax.transData.transform((xa,ya)); pb=ax.transData.transform((xb,yb))
        angle=np.degrees(np.arctan2(pb[1]-pa[1],pb[0]-pa[0]))
        ax.text(xm,ym,ls_labels[lsi],fontsize=8.5,color=ls_colors[lsi],
                fontweight='bold',ha='center',va='center',rotation=angle,
                bbox=dict(fc='white',ec='none',pad=2.0))
    # เส้นแดง
    ax.plot([k_inf_pci,k_inf_pci],[1,k_eff_pci],'r-',lw=2.0)
    ax.annotate('',xy=(k_inf_pci,k_eff_pci),xytext=(k_inf_pci,1),
                arrowprops=dict(arrowstyle='->',color='red',lw=2.0))
    ax.plot([k_inf_pci,1],[k_eff_pci,k_eff_pci],'r-',lw=2.0)
    ax.annotate('',xy=(1,k_eff_pci),xytext=(k_inf_pci,k_eff_pci),
                arrowprops=dict(arrowstyle='->',color='red',lw=2.0))
    ax.plot(k_inf_pci,k_eff_pci,'ro',markersize=8,zorder=5)
    ax.annotate(f'k_eff = {k_eff_pci:.0f} pci',
                xy=(1,k_eff_pci),xytext=(2.5,k_eff_pci*1.8),
                fontsize=9,color='red',fontweight='bold',
                arrowprops=dict(arrowstyle='->',color='red',lw=1.2))
    ax.annotate(f'k∞ = {k_inf_pci:.0f} pci',
                xy=(k_inf_pci,1),xytext=(k_inf_pci*0.4,1.6),
                fontsize=9,color='red',fontweight='bold',
                arrowprops=dict(arrowstyle='->',color='red',lw=1.2))
    ax.set_xlabel('Effective Modulus of Subgrade Reaction, k∞ (pci)',fontsize=10)
    ax.set_ylabel('k (Corrected for Potential Loss of Support) (pci)',fontsize=9)
    ax.set_title(f'AASHTO 1993 Figure 3.4 — Loss of Support\n'
                 f'k∞={k_inf_pci:.0f} pci, LS={ls} → k_eff={k_eff_pci:.0f} pci',
                 fontsize=10,color='red')
    plt.tight_layout()
    return fig

def plot_structure(layers, concrete_cm=None, title='Pavement Structure'):
    """วาดรูปโครงสร้างชั้นทาง"""
    all_layers = []
    if concrete_cm and concrete_cm > 0:
        all_layers.append({'name':'Concrete Slab','thickness_cm':concrete_cm,'E_MPa':None})
    all_layers.extend([l for l in layers if l.get('thickness_cm',0)>0])
    if not all_layers: return None
    total = sum(l['thickness_cm'] for l in all_layers)
    min_h = 8
    disp  = [max(l['thickness_cm'],min_h) for l in all_layers]
    tot_d = sum(disp)
    fig, ax = plt.subplots(figsize=(10,7))
    w,xc = 3,6
    xs = xc-w/2
    y = tot_d
    for i,layer in enumerate(all_layers):
        t=layer['thickness_cm']; n=layer['name']; e=layer.get('E_MPa')
        dh=disp[i]; yb=y-dh
        col=LAYER_COLORS.get(n,'#CCCCCC')
        ax.add_patch(patches.Rectangle((xs,yb),w,dh,lw=2,ec='black',fc=col))
        yc=yb+dh/2
        en=LAYER_NAMES_EN.get(n,n)
        dark=n in ["รองผิวทางคอนกรีตด้วย AC","รองผิวทางคอนกรีตด้วย PMA(AC)",
                   "Concrete Slab","หินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (CTB)",
                   "หินคลุกผสมซีเมนต์ UCS 24.5 ksc","วัสดุหมุนเวียน (Recycling)"]
        tc='white' if dark else 'black'
        ax.text(xc,yc,f'{t} cm',ha='center',va='center',fontsize=14,fontweight='bold',color=tc)
        ax.text(xs-0.4,yc,en,ha='right',va='center',fontsize=12,fontweight='bold')
        if e: ax.text(xs+w+0.4,yc,f'E = {e:,} MPa',ha='left',va='center',fontsize=11,color='#0066CC')
        y=yb
    ax.annotate('',xy=(xs+w+3.5,tot_d),xytext=(xs+w+3.5,0),
                arrowprops=dict(arrowstyle='<->',color='red',lw=2))
    ax.text(xs+w+4,tot_d/2,f'Total\n{total} cm',ha='left',va='center',fontsize=13,color='red',fontweight='bold')
    mg=10
    ax.set_xlim(0,14); ax.set_ylim(-mg,tot_d+mg)
    ax.axis('off')
    ax.set_title(title,fontsize=16,fontweight='bold',pad=16)
    plt.tight_layout()
    return fig

def fig_to_bytes(fig):
    buf=BytesIO()
    fig.savefig(buf,format='png',dpi=150,bbox_inches='tight',facecolor='white')
    buf.seek(0)
    return buf.read()

# ============================================================
# HELPERS UI
# ============================================================

def section_card(title, icon=''):
    st.markdown(f'<div class="rp-card-title">{icon} {title}</div>', unsafe_allow_html=True)

def result_row(label, value):
    st.markdown(f'''<div class="rp-result-row">
        <span class="rp-result-label">{label}</span>
        <span class="rp-result-val">{value}</span></div>''', unsafe_allow_html=True)

def layer_input_block(prefix, num_layers, default_layers):
    """Block กรอก layers — returns list of dicts"""
    mat_opts = list(MATERIAL_MODULUS.keys())
    layers = []
    for i in range(num_layers):
        with st.container():
            st.markdown(f'<div class="rp-section-label">ชั้นที่ {i+1}</div>', unsafe_allow_html=True)
            ca, cb, cc = st.columns([3,1,1])
            def_name = st.session_state.get(f'{prefix}_name_{i}',
                       default_layers[i]['name'] if i<len(default_layers) else 'หินคลุก CBR 80%')
            def_thick= st.session_state.get(f'{prefix}_thick_{i}',
                       default_layers[i]['thick'] if i<len(default_layers) else 20)
            if def_name not in mat_opts: def_name = mat_opts[-1]
            idx = mat_opts.index(def_name)
            with ca:
                name = st.selectbox('วัสดุ', mat_opts, index=idx,
                                    key=f'{prefix}_name_{i}', label_visibility='collapsed')
            with cb:
                thick= st.number_input('ซม.', 0, 200, def_thick,
                                       key=f'{prefix}_thick_{i}', label_visibility='collapsed')
            rec_e = MATERIAL_MODULUS.get(name,100)
            def_e = st.session_state.get(f'{prefix}_E_{i}_{name}', rec_e)
            with cc:
                e_val= st.number_input('E (MPa)', 10,10000, def_e,
                                       key=f'{prefix}_E_{i}_{name}', label_visibility='collapsed')
            layers.append({'name':name,'thickness_cm':thick,'E_MPa':e_val})
    return layers

DEFAULT_LAYERS_JPCP = [
    {'name':'รองผิวทางคอนกรีตด้วย AC','thick':5},
    {'name':'หินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (CTB)','thick':20},
    {'name':'หินคลุก CBR 80%','thick':15},
    {'name':'รองพื้นทางวัสดุมวลรวม CBR 25%','thick':25},
    {'name':'วัสดุคัดเลือก ก','thick':30},
]
DEFAULT_LAYERS_CRCP = [
    {'name':'หินคลุกปรับปรุงคุณภาพด้วยปูนซีเมนต์ (CTB)','thick':10},
    {'name':'รองพื้นทางวัสดุมวลรวม CBR 25%','thick':15},
    {'name':'วัสดุคัดเลือก ก','thick':20},
]

# ============================================================
# MAIN
# ============================================================

def main():
    st.set_page_config(
        page_title='Rigid Pavement Design V7 — AASHTO 1993',
        page_icon='🛣️', layout='wide'
    )
    st.markdown(CSS, unsafe_allow_html=True)

    # ── Session State Init ───────────────────────────────────
    defaults = {
        'project_name': '',
        'esal_data': None,
        'esal_file_id': None,
        'cbr': 4.0,
        # JPCP k
        'jpcp_k_inf': None, 'jpcp_k_eff': None,
        'jpcp_ls': 1.0, 'jpcp_fig33': None, 'jpcp_fig34': None,
        # CRCP k
        'crcp_k_inf': None, 'crcp_k_eff': None,
        'crcp_ls': 1.0, 'crcp_fig33': None, 'crcp_fig34': None,
        'crcp_copy_jpcp': False,
        # Design
        'jpcp_d_cm': 30, 'crcp_d_cm': 25,
    }
    for k,v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # ── Header ───────────────────────────────────────────────
    proj = st.session_state.get('project_name','') or '— ยังไม่ระบุโครงการ'
    st.markdown(f'''
    <div class="rp-header">
        <div>
            <div class="rp-header-title">🛣️ Rigid Pavement Design — AASHTO 1993</div>
            <div class="rp-header-sub">ภาควิชาครุศาสตร์โยธา · มจพ. · Version 7.0</div>
        </div>
        <div class="rp-badge">📁 {proj}</div>
    </div>''', unsafe_allow_html=True)

    # ── Tabs ─────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs([
        '🚦 Tab 1 — Traffic & ESAL',
        '🪨 Tab 2 — Subgrade & k∞',
        '🏗️ Tab 3 — Design',
    ])

    # ============================================================
    # TAB 1: Traffic & ESAL
    # ============================================================
    with tab1:
        st.markdown('<div class="rp-card">', unsafe_allow_html=True)
        section_card('ข้อมูลโครงการ', '📋')
        proj_name = st.text_input('ชื่อโครงการ', value=st.session_state['project_name'],
                                  key='project_name_input', placeholder='เช่น ทางหลวงหมายเลข 1 ตอน กรุงเทพ-สระบุรี')
        st.session_state['project_name'] = proj_name
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="rp-card">', unsafe_allow_html=True)
        section_card('นำเข้า ESAL จาก ESAL Calculator', '📊')
        st.caption('อัปโหลดไฟล์ .json ที่ Save จาก ESAL Calculator (Rigid)')

        esal_file = st.file_uploader('เลือกไฟล์ ESAL Project (.json)',
                                     type=['json'], key='esal_uploader',
                                     help='ไฟล์ที่ได้จากปุ่ม บันทึก Project ใน ESAL Calculator')

        if esal_file is not None:
            fid = f'{esal_file.name}_{esal_file.size}'
            if st.session_state.get('esal_file_id') != fid:
                st.session_state['esal_file_id'] = fid
                try:
                    raw = json.load(esal_file)
                    if raw.get('pavement_type') != 'rigid':
                        st.error('❌ ไฟล์นี้เป็น Flexible Pavement — กรุณาใช้ไฟล์ Rigid')
                        st.session_state['esal_data'] = None
                    elif 'traffic_data' not in raw:
                        st.error('❌ ไฟล์ JSON ไม่ถูกต้อง (ขาด traffic_data)')
                        st.session_state['esal_data'] = None
                    else:
                        st.session_state['esal_data'] = {
                            'traffic_data':     raw['traffic_data'],
                            'pt':               raw.get('pt', 2.0),
                            'lane_factor':      raw.get('lane_factor', 0.9),
                            'direction_factor': raw.get('direction_factor', 0.5),
                            'filename':         esal_file.name,
                            'num_years':        len(raw['traffic_data']),
                        }
                        st.rerun()
                except Exception as ex:
                    st.error(f'❌ อ่านไฟล์ไม่ได้: {ex}')

        ed = st.session_state.get('esal_data')
        if ed:
            st.markdown(f'''<div class="rp-status-ok">
                ✅ นำเข้าสำเร็จ: {ed["filename"]} |
                ระยะออกแบบ {ed["num_years"]} ปี |
                pt = {ed["pt"]} |
                Lane Factor = {ed["lane_factor"]} |
                Dir. Factor = {ed["direction_factor"]}
            </div>''', unsafe_allow_html=True)
            st.markdown('')
            # Preview traffic table
            with st.expander('📋 ดูข้อมูลจราจร', expanded=False):
                df = pd.DataFrame(ed['traffic_data'])
                st.dataframe(df, use_container_width=True, hide_index=True)
            if st.button('🗑️ ล้างข้อมูล ESAL', key='clear_esal'):
                st.session_state['esal_data'] = None
                st.session_state['esal_file_id'] = None
                st.rerun()
        else:
            st.markdown('<div class="rp-status-info">⚪ ยังไม่ได้นำเข้าข้อมูล ESAL</div>',
                        unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Serviceability & Reliability ────────────────────
        st.markdown('<div class="rp-card">', unsafe_allow_html=True)
        section_card('Serviceability & Reliability', '📉')
        c1,c2,c3 = st.columns(3)
        with c1:
            pt = st.slider('Terminal Serviceability (Pt)', 1.5, 3.0,
                           st.session_state.get('pt',2.0), 0.1, key='pt')
            dpsi = 4.5 - pt
            st.caption(f'ΔPSI = 4.5 − {pt:.1f} = **{dpsi:.1f}**')
            # ── warning pt ไม่สอดคล้อง ──
            ed_check = st.session_state.get('esal_data')
            if ed_check:
                pt_json = ed_check.get('pt', pt)
                if abs(pt - pt_json) > 0.01:
                    st.warning(f'⚠️ pt ที่ใช้ ({pt:.1f}) ต่างจาก ESAL JSON ({pt_json:.1f})\n\nW18 จะถูกคำนวณใหม่ด้วย pt = {pt:.1f}')
                    if st.button(f'↩️ ใช้ค่าจาก ESAL JSON (pt = {pt_json:.1f})', key='reset_pt'):
                        st.session_state['pt'] = pt_json
                        st.rerun()
                else:
                    st.success(f'✅ pt สอดคล้องกับ ESAL JSON')
        with c2:
            reliability = st.select_slider('Reliability (%)',
                          options=[80,85,90,95],
                          value=st.session_state.get('reliability',90), key='reliability')
            zr = get_zr(reliability)
            st.caption(f'ZR = **{zr:.3f}**')
        with c3:
            so = st.number_input('Overall Std. Deviation (So)',
                                 0.30,0.45,
                                 st.session_state.get('so',0.35), 0.01,
                                 format='%.2f', key='so')
        st.markdown('</div>', unsafe_allow_html=True)

        # ── ESAL Summary — 4 cards ตาม D ───────────────────
        if ed:
            st.markdown('<div class="rp-card">', unsafe_allow_html=True)
            section_card('W18 ตามความหนาคอนกรีต D (คำนวณจาก ESAL JSON)', '🔢')
            st.caption('W18 ขึ้นกับ D และ pt — แสดง 4 ค่า D ที่ใช้บ่อย')
            pt_val = st.session_state.get('pt', 2.0)
            D_CARDS_IN = [25, 28, 30, 32]   # หน่วย inch
            cols = st.columns(4)
            for i, d_in in enumerate(D_CARDS_IN):
                d_cm = round(d_in * 2.54)
                w18, _, _ = compute_esal_for_d(
                    ed['traffic_data'], pt_val,
                    ed['lane_factor'], ed['direction_factor'], d_cm)
                with cols[i]:
                    st.markdown(f'''<div class="rp-metric">
                        <div class="rp-metric-label">D = {d_in} in ({d_cm} ซม.)</div>
                        <div class="rp-metric-val" style="font-size:16px">{w18:,}</div>
                        <div style="font-size:11px;color:#90A4AE;margin-top:2px">
                            {w18/1e6:.2f} ล้าน ESALs
                        </div>
                    </div>''', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

    # ============================================================
    # TAB 2 & 3 — placeholder สำหรับ Step ถัดไป
    # ============================================================
    with tab2:
        st.info('🔧 Tab 2 — Subgrade & k∞ (Step 3 — coming next)')

    with tab3:
        st.info('🔧 Tab 3 — Design (Step 4 — coming next)')

    # Footer
    st.markdown('---')
    st.caption('พัฒนาโดย รศ.ดร.อิทธิพล มีผล · ภาควิชาครุศาสตร์โยธา · มจพ.')

if __name__ == '__main__':
    main()
