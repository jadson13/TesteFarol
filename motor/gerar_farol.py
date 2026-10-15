#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera o Farol de Inspecao Pre-Viagem (Amazon) a partir dos arquivos mais recentes
na pasta "Farol Amazon - Entradas". Detecta os arquivos por conteudo.

Regras (v2 - ajustes do cliente):
  - Multidrop: 1 VRID = 1 viagem. Quando o VRID tem varias pernas, considera SEMPRE a
    primeira (perna cuja origem = primeiro ponto da rota da coluna C; desempate = horario
    mais cedo). As demais pernas sao descartadas.
  - Janela:  Aderente = inspecao ate 12h antes da viagem
             Fora do prazo = inspecao entre 12h e 24h antes
             Sem inspecao  = nenhuma inspecao valida (inclui inspecao > 24h antes, inspecao
                             apos a viagem, ou placa/VRID sem inspecao) -> Amazon nao considera
  - Anomalia: para a inspecao casada, "NOK" se algum item respondido = 'N'; senao "OK".
  - Dashboard HTML tem 2 abas: (1) Farol de aderencia  (2) Inspecoes sem VRID preenchido.

Saidas (na raiz da pasta conectada):  Farol_Amazon.xlsx  e  Farol_Amazon.html
"""
import os, re, sys, glob, json
from datetime import datetime, timedelta, date
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import PieChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList

SUP = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.dirname(SUP)
ENTRADAS = os.path.join(BASE, "Farol Amazon - Entradas")
PLACAS_ENT = os.path.join(BASE, "Placas Espelhamento - Entradas")
BANCO = os.path.join(BASE, "Banco")
W12 = timedelta(hours=12); W24 = timedelta(hours=24)
# Auditoria: descartar inspecoes de auditor (Inspecao de Seguranca) anteriores a esta data.
AUD_MIN_DATE = datetime(2026, 8, 1)

ADER="Aderente (ate 12h)"; FORA="Fora do prazo (12-24h)"; SEMI="Sem inspecao"; SDAT="Sem data de viagem"
LBL={0:ADER,1:FORA,2:SEMI,3:SDAT}
MESES={1:"Jan",2:"Fev",3:"Mar",4:"Abr",5:"Mai",6:"Jun",7:"Jul",8:"Ago",9:"Set",10:"Out",11:"Nov",12:"Dez"}
# Placa ficticia usada pelo auditor quando o veiculo NAO tem cadastro no sistema.
# Toda inspecao com essa placa = "veiculo sem cadastro identificado" (Melhoria 1).
SEM_CADASTRO_PLATE="ASDFG"
# Auditor de referencia para a visao de anomalias de auditoria (Melhoria 2).
MIGUEL="Miguel de Sousa Cardoso"

def npl(p): return re.sub(r'[^A-Z0-9]','',str(p).upper()) if p is not None else ""
def nvr(v): return str(v).strip().upper() if v is not None else ""
def log(*a): print(*a, flush=True)
def _split_carriers(t):
    """Separa transportadoras combinadas por virgula (NAO por '/', p/ nao quebrar 'CARGOLIFT S/A').
    Uma placa que atua em varias transportadoras conta em cada uma."""
    parts=[p.strip() for p in str(t or "").split(",")]; seen=set(); out=[]
    for p in parts:
        if not p: continue
        k=p.upper()
        if k not in seen: seen.add(k); out.append(p)
    return out or ["(sem transp.)"]
def _norm_tipo(s):
    return re.sub(r'\s+',' ',str(s or "").replace('-',' ')).strip().upper()

# ---------------- deteccao dos arquivos ----------------
def list_xlsx():
    return sorted([f for f in glob.glob(os.path.join(ENTRADAS,"*.xlsx")) if not os.path.basename(f).startswith("~$")],
                  key=os.path.getmtime, reverse=True)
def sniff(path):
    try: wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception: return None
    kind=None
    for ws in wb.worksheets:
        head=[]
        for r in ws.iter_rows(min_row=1,max_row=6,values_only=True):
            head += [str(c).strip().lower() for c in r if c is not None]
        j=" | ".join(head)
        if "planned_orig_checkin" in j or ("vrid" in j and "placa final" in j) or ("vrid" in j and ("previsão chegada na origem" in j or "previsao chegada na origem" in j)): kind="viagens"; break
        if "data checklist" in j and ("identificador" in j or "vrid" in j): kind="inspecoes"; break
    wb.close(); return kind
def pick_inputs():
    """Retorna (lista_viagens, lista_inspecoes) com TODOS os arquivos de cada tipo
    encontrados na pasta de Entradas (permite acumular bases e deduplicar)."""
    vs=[]; iss=[]
    for f in list_xlsx():
        k=sniff(f)
        if k=="viagens": vs.append(f)
        elif k=="inspecoes": iss.append(f)
    return vs, iss

# ---------------- viagens: normaliza cada arquivo, junta e dedup por VRID ----------------
# registro padrao: (vrid, carrier, route, lane, week_year, time|None, plate)
CI={"vrid":0,"carrier":1,"route":2,"lane":3,"week":4,"time":5,"plate":6}

def _clean(x):
    if x is None: return ""
    s=str(x).strip()
    return "" if s.lower() in ("null","none","nan") else s

def _select_trip_sheet(wb):
    """Escolhe a aba cujo cabecalho tem vrid + coluna de checkin (evita abas de outro relatorio)."""
    for ws in wb.worksheets:
        first=next(ws.iter_rows(min_row=1,max_row=1,values_only=True),())
        hdr=[str(c).strip().lower() for c in first if c is not None]
        j=" ".join(hdr)
        if "vrid" in hdr and ("checkin_time" in j or "planned_orig_checkin_time" in j): return ws
    return max(wb.worksheets, key=lambda w:(w.max_row or 0))

def _to_dt(v):
    if isinstance(v, datetime): return v
    if v is None: return None
    s=str(v).strip().replace('T',' ')
    if not s: return None
    import re as _re
    m=_re.match(r'(\d{4})-(\d{2})-(\d{2})[ ]?(\d{2}:\d{2}:\d{2})?(\d{2}:\d{2})?', s)
    if m:
        try:
            y,mo,d=int(m.group(1)),int(m.group(2)),int(m.group(3))
            if m.group(4):
                hh,mm,ss=[int(x) for x in m.group(4).split(':')]; return datetime(y,mo,d,hh,mm,ss)
            if m.group(5):
                hh,mm=[int(x) for x in m.group(5).split(':')]; return datetime(y,mo,d,hh,mm)
            return datetime(y,mo,d)
        except Exception: pass
    return None

def normalize_trips(path):
    wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws=_select_trip_sheet(wb)
    rows=list(ws.iter_rows(values_only=True)); wb.close()
    if not rows: return [], []
    hdr=[str(c).strip().lower() if c is not None else "" for c in rows[0]]
    def col(*names):
        for n in names:
            if n in hdr: return hdr.index(n)
        return None
    def col_c(*subs):
        for i,h in enumerate(hdr):
            for s in subs:
                if s in h: return i
        return None
    c_vrid=col("vrid") if col("vrid") is not None else col_c("vrid")
    c_car =col("carrier","transportadora")
    c_route=col("route","rota")
    if c_route is None: c_route=col_c("lane (origem e destino","lane")   # RelatorioMaster: a rota completa vem na coluna Lane
    c_lane=col("lane")
    c_week=col("week_year") if col("week_year") is not None else col_c("week_year")
    c_time=col("planned_orig_checkin_time_local","planned_orig_checkin_time")
    if c_time is None: c_time=col_c("checkin_time","checkin")
    if c_time is None: c_time=col("previsão chegada na origem","previsao chegada na origem")  # RelatorioMaster: check-in = Previsão Chegada na Origem
    if c_time is None: c_time=col_c("previsão chegada na origem","previsao chegada na origem")
    c_plate=col("placa final","placa_final","placa","plate_brk","plate")
    if c_plate is None: c_plate=col_c("plate_brk","placa final","placa","plate")
    def g(r,c): return r[c] if (c is not None and c<len(r)) else None
    out=[]; pairs=[]
    for r in rows[1:]:
        if c_vrid is None or c_vrid>=len(r) or r[c_vrid] is None: continue
        t=_to_dt(g(r,c_time))
        wy=_clean(g(r,c_week))
        rec=(_clean(g(r,c_vrid)), _clean(g(r,c_car)) or "(sem transp.)", _clean(g(r,c_route)),
             _clean(g(r,c_lane)), wy, t, _clean(g(r,c_plate)))
        if not rec[0]: continue
        out.append(rec)
        if wy and t: pairs.append((wy, t.date()))
    return out, pairs

def _first_leg(rs):
    if len(rs)==1: return rs[0]
    def orig(s): return re.split(r'\s*-+>\s*', s)[0].strip().upper() if s else ""
    def t(r): return r[5] if isinstance(r[5],datetime) else datetime.max
    cand=[r for r in rs if r[3] and orig(r[3])==orig(r[2])]  # origem da lane == origem da rota
    pool=cand if cand else rs
    return sorted(pool, key=t)[0]

def read_trips(paths):
    if isinstance(paths,str): paths=[paths]
    all_recs=[]; all_pairs=[]; nlines=0
    for p in paths:
        recs,pairs=normalize_trips(p); all_recs+=recs; all_pairs+=pairs; nlines+=len(recs)
    week_of=build_week_of(all_pairs)
    groups=defaultdict(list)
    for rec in all_recs: groups[rec[0]].append(rec)
    trips=[_first_leg(rs) for rs in groups.values()]
    return trips, CI, nlines, week_of

def build_week_of(pairs):
    """pairs = lista de (week_year, date). Semana calendario Amazon (dom-sab), ancorada
    nos sabados observados. Fallback: semana ISO."""
    from datetime import timedelta
    maxdate={}
    for wy,d in pairs:
        if wy and (wy not in maxdate or d>maxdate[wy]): maxdate[wy]=d
    anchor=None
    for wy,d in maxdate.items():
        m=re.match(r'(\d{4})W(\d{1,2})', wy)
        if m and d.weekday()==5:
            n=int(m.group(2)); y=int(m.group(1))
            if anchor is None or d>anchor[1]: anchor=(n,d,y)
    def week_of(d):
        if anchor is None:
            y,w,_=d.isocalendar(); return f"{y}W{w:02d}"
        n0,s0,y0=anchor
        sat=d+timedelta(days=(5-d.weekday())%7)
        delta=(sat-s0).days//7
        return f"{y0}W{n0+delta:02d}"
    return week_of

# ---------------- inspecoes: parser por arquivo (p/ banco) + construtor do insp_info ----------------
def parse_insp_file(path):
    """Devolve (lista de registros, item_group{nome:grupo}). Registro:
       {ident,tipo,transp,user,data(iso),carreta,vrid,transpv,nok:[nomes de itens]}"""
    wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws=wb["Checklist - 4"] if "Checklist - 4" in wb.sheetnames else max(wb.worksheets, key=lambda w:(w.max_row or 0))
    rows=list(ws.iter_rows(values_only=True)); wb.close()
    def find_col(*labels):
        labs=[l.lower() for l in labels]
        for ri in range(min(6,len(rows))):
            for ci,val in enumerate(rows[ri]):
                if val is not None and str(val).strip().lower() in labs: return ci
        return None
    c_ident=find_col("identificador") or 0
    c_tipo =find_col("tipo equipamento")
    c_transp=find_col("transportadora")
    c_user =find_col("usuario responsavel","usuário responsável")
    c_data =find_col("data checklist"); c_data=4 if c_data is None else c_data
    c_carr =find_col("placa carreta")
    c_vrid =find_col("numero do vrid","número do vrid")
    c_transpv=find_col("transportadora responsavel pela viagem","transportadora responsável pela viagem")
    r_item = rows[2] if len(rows)>2 else []
    r_grp  = rows[1] if len(rows)>1 else []
    resp_cols=[]
    for ri in range(min(6,len(rows))):
        rr=[i for i,v in enumerate(rows[ri]) if v is not None and str(v).strip().lower()=="resposta"]
        if rr: resp_cols=rr; break
    loc=[(c,(str(r_item[c]).strip() if c<len(r_item) and r_item[c] else f"col{c}"),
            (str(r_grp[c]).strip() if c<len(r_grp) and r_grp[c] else "")) for c in resp_cols]
    item_group={nm:gp for (_,nm,gp) in loc}
    start=0
    for ri,r in enumerate(rows):
        if c_data<len(r) and isinstance(r[c_data],datetime): start=ri; break
    def g(r,c): return (r[c] if (c is not None and c<len(r)) else None)
    recs=[]
    for r in rows[start:]:
        if c_data>=len(r) or not isinstance(r[c_data],datetime): continue
        dt=r[c_data]
        nk=[nm for (c,nm,gp) in loc if c<len(r) and str(r[c]).strip().upper()=="N"]
        recs.append({"ident":str(g(r,c_ident) or ""),"tipo":(str(g(r,c_tipo)).strip() if g(r,c_tipo) else "(?)"),
                     "transp":(str(g(r,c_transp)).strip() if g(r,c_transp) else "(sem transp.)"),
                     "user":str(g(r,c_user) or ""),"data":dt.isoformat(),
                     "carreta":str(g(r,c_carr) or ""),"vrid":str(g(r,c_vrid) or ""),
                     "transpv":str(g(r,c_transpv) or ""),"nok":nk})
    return recs, item_group

def build_insp_info(records, item_group, week_of=None):
    if week_of is None: week_of=lambda d:(lambda t:f"{t[0]}W{t[1]:02d}")(d.isocalendar())
    carriers=[]; cix={}; months=[]; mix={}; tipos=[]; tix={}; weeks=[]; wix={}
    item_names=[]; item_groups=[]; iname_idx={}
    def _idx(x,lst,d):
        if x not in d: d[x]=len(lst); lst.append(x)
        return d[x]
    def cidx(x): return _idx(x,carriers,cix)
    def midx(x): return _idx(x,months,mix)
    def tidx(x): return _idx(x,tipos,tix)
    def widx(x): return _idx(x,weeks,wix)
    def item_i(nm):
        if nm not in iname_idx:
            iname_idx[nm]=len(item_names); item_names.append(nm); item_groups.append(item_group.get(nm,""))
        return iname_idx[nm]
    iv=defaultdict(list); ip=defaultdict(list)
    total=0; com_vrid=0; nok_total=0; agg=[]; sem_rows=[]; nok_rows=[]
    for rec in records:
        dt=datetime.fromisoformat(rec["data"])
        nk=[item_i(nm) for nm in rec.get("nok",[])]; nok=len(nk)>0
        total+=1
        if nok: nok_total+=1
        month=f"{MESES[dt.month]}/{dt.year}"; wk=week_of(dt.date())
        cI,mI,tI,wI=cidx(rec["transp"]),midx(month),tidx(rec["tipo"]),widx(wk)
        v=nvr(rec["vrid"]); has=1 if v else 0
        if v: com_vrid+=1; iv[v].append((dt,nok))
        agg.append([cI,mI,tI,has,1 if nok else 0,wI])
        ds=dt.strftime("%Y-%m-%d %H:%M")
        if not v:
            sem_rows.append([rec["ident"],tI,cI,rec["user"],ds,rec["carreta"],rec.get("transpv",""),("NOK" if nok else "OK"),mI,wI])
        if nok:
            nok_rows.append([rec["ident"],tI,cI,rec["user"],ds,rec["carreta"],rec["vrid"],mI,nk,wI])
        for p in (npl(rec["ident"]), npl(rec["carreta"])):
            if p: ip[p].append((dt,nok))
    for d in (iv,ip):
        for k in d: d[k].sort(key=lambda x:x[0])
    insp_info={"total":total,"com_vrid":com_vrid,"sem_vrid":len(sem_rows),"nok_total":nok_total,
               "carriers":carriers,"months":months,"tipos":tipos,"weeks":weeks,
               "items":item_names,"item_groups":item_groups,
               "agg":agg,"sem_rows":sem_rows,"nok_rows":nok_rows}
    return iv, ip, insp_info

# ---------------- placas / espelhamento ----------------
def parse_placas_file(path):
    """Devolve lista de tuplas (vrid, carrier, route, time_iso, placaF)."""
    wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws=max(wb.worksheets, key=lambda w:(w.max_row or 0))
    rows=list(ws.iter_rows(values_only=True)); wb.close()
    if not rows: return []
    hdr=[str(c).strip().lower() if c is not None else "" for c in rows[0]]
    def col(*names):
        for n in names:
            if n in hdr: return hdr.index(n)
        return None
    def col_c(*subs):
        for i,h in enumerate(hdr):
            for s in subs:
                if s in h: return i
        return None
    c_v=col("vrid") if col("vrid") is not None else col_c("vrid")
    c_car=col("carrier","transportadora")
    c_route=col("route","rota")
    c_time=col("planned_orig_checkin_time_local_ats","planned_orig_checkin_time_local","planned_orig_checkin_time")
    if c_time is None: c_time=col_c("checkin")
    c_pl=col("placa","placa final") if col("placa","placa final") is not None else col_c("placa","plate")
    def g(r,c): return r[c] if (c is not None and c<len(r)) else None
    out=[]
    for r in rows[1:]:
        if c_v is None or c_v>=len(r) or r[c_v] is None: continue
        t=g(r,c_time); tiso=t.isoformat() if isinstance(t,datetime) else ""
        out.append((_clean(g(r,c_v)), _clean(g(r,c_car)) or "(sem transp.)", _clean(g(r,c_route)), tiso, _clean(g(r,c_pl))))
    return out

def build_placas_info(placas, iv, week_of):
    """Sem espelhamento = placaF == 'Sem Aceite'. 1 por VRID (perna mais cedo).
       Cruza por VRID com inspecoes (iv) aplicando a regua do farol (0 aderente/1 fora/2 sem)."""
    byv={}
    for (v,car,route,tiso,pl) in placas:
        if str(pl).strip().lower()!="sem aceite": continue
        t=datetime.fromisoformat(tiso) if tiso else None
        k=v.upper(); cur=byv.get(k)
        if cur is None or (t and (cur[3] is None or t<cur[3])): byv[k]=(v,car,route,t)
    carriers=[]; cix={}; weeks=[]; wix={}; months=[]; mix={}
    def cidx(x):
        if x not in cix: cix[x]=len(carriers); carriers.append(x)
        return cix[x]
    def widx(x):
        if x not in wix: wix[x]=len(weeks); weeks.append(x)
        return wix[x]
    def midx(x):
        if x not in mix: mix[x]=len(months); months.append(x)
        return mix[x]
    rows=[]; cnt=[0,0,0,0]
    for k,(v,car,route,t) in byv.items():
        if t is None:
            sc=3; ds=""; wI=-1; mI=-1
        else:
            r=pick(iv.get(v.upper(),[]), t); sc=r[0] if r else 2
            ds=t.strftime("%Y-%m-%d %H:%M"); wI=widx(week_of(t.date())); mI=midx(f"{MESES[t.month]}/{t.year}")
        cnt[sc]+=1
        rows.append([v,cidx(car),route,ds,sc,wI,mI])
    return {"carriers":carriers,"weeks":weeks,"months":months,
            "cat":{"aderente":cnt[0],"fora":cnt[1],"sem":cnt[2],"semdata":cnt[3]},
            "total":len(byv),"rows":rows}

# ================= NOVAS FONTES: Auditor, Prontuário (documentos), Equipamentos (combustível) =================
def looks_plate(x):
    s=str(x).strip() if x is not None else ""
    return 5<=len(s)<=10 and " " not in s and any(c.isdigit() for c in s) and any(c.isalpha() for c in s)

def _extract_sem_cad(obs_pairs):
    """Para um checklist com a placa ficticia (ASDFG), o auditor digita a placa/transportadora
    REAL do veiculo sem cadastro na observacao de algum item (texto livre, item varia).
    Convencao observada: 1a linha = placa, 2a linha = transportadora. Retorna (placa, transp, obs_raw)."""
    placa=""; transp=""; raw=""
    for nm,ov in obs_pairs:
        lines=[x.strip() for x in re.split(r'[\n\r]+', str(ov)) if x.strip()]
        if lines and looks_plate(lines[0]):
            placa=lines[0].upper()
            if len(lines)>1: transp=lines[1]
            raw=ov; break
    if not raw and obs_pairs:
        raw=" | ".join(dict.fromkeys(str(ov) for _,ov in obs_pairs))
    return placa, transp, raw

# ---- Inspeção de AUDITOR: abas "Inspeção de Segurança" nos arquivos de inspeção ----
def parse_auditor_file(path):
    wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
    out=[]
    for sh in wb.sheetnames:
        ws=wb[sh]
        first=next(ws.iter_rows(min_row=1,max_row=1,values_only=True),())
        nome=_clean(first[0]) if first else ""
        if "inspeção de segurança" not in nome.lower() and "inspecao de seguranca" not in nome.lower(): continue
        rows=list(ws.iter_rows(values_only=True))
        def findcol(*labs):
            labs=[l.lower() for l in labs]
            for ri in range(min(6,len(rows))):
                for ci,v in enumerate(rows[ri]):
                    if v is not None and str(v).strip().lower() in labs: return ci
            return None
        c_ident=findcol("identificador") or 0
        c_tipo=findcol("tipo equipamento"); c_transp=findcol("transportadora")
        c_user=findcol("usuário responsável","usuario responsavel")
        c_data=findcol("data checklist")
        r_item=rows[2] if len(rows)>2 else []; r_grp=rows[1] if len(rows)>1 else []
        resp_cols=[]; rrow=None
        for ri in range(min(6,len(rows))):
            rr=[i for i,v in enumerate(rows[ri]) if v is not None and str(v).strip().lower()=="resposta"]
            if rr: resp_cols=rr; rrow=ri; break
        def _is_obs(ci):
            if rrow is None or ci is None: return False
            hr=rows[rrow]
            return ci<len(hr) and hr[ci] is not None and "observ" in str(hr[ci]).strip().lower()
        # loc: (resp_col, item_name, grupo, obs_col)  -- obs_col = coluna "Observacao" do item (ou None)
        loc=[]
        for c in resp_cols:
            nm=str(r_item[c]).strip() if c<len(r_item) and r_item[c] else f"col{c}"
            gp=str(r_grp[c]).strip() if c<len(r_grp) and r_grp[c] else ""
            loc.append((c,nm,gp,(c+1 if _is_obs(c+1) else None)))
        start=0
        for ri,r in enumerate(rows):
            if c_data is not None and c_data<len(r) and isinstance(r[c_data],datetime): start=ri; break
        def g(r,c): return (str(r[c]).strip() if (c is not None and c<len(r) and r[c] is not None) else "")
        for r in rows[start:]:
            if c_data is None or c_data>=len(r) or not isinstance(r[c_data],datetime): continue
            dt=r[c_data]
            nk=[nm for (c,nm,gp,oc) in loc if c<len(r) and str(r[c]).strip().upper()=="N"]
            grp={nm:gp for (c,nm,gp,oc) in loc}
            ident=g(r,c_ident).upper()
            rec={"ident":ident,"tipo":g(r,c_tipo) or "(?)","transp":g(r,c_transp) or "(sem transp.)",
                 "user":g(r,c_user),"data":dt.isoformat(),"nok":nk,"itemgrp":grp,"sheet":nome}
            if ident==SEM_CADASTRO_PLATE:
                obs_pairs=[(nm,g(r,oc)) for (c,nm,gp,oc) in loc if oc is not None and g(r,oc)]
                placa,transp_real,obs_raw=_extract_sem_cad(obs_pairs)
                rec["sc"]=True; rec["placa_real"]=placa; rec["transp_real"]=transp_real; rec["obs_raw"]=obs_raw
            out.append(rec)
    wb.close(); return out

# ---- PRONTUÁRIO de Equipamentos (documentos + inspeções por placa) ----
PRONT_DOCS=["CRLV","CRV","Código Identificador da Operação de Transportes","Nota Fiscal","Preventiva Equipamento"]
PRONT_DOC_SHORT={"CRLV":"CRLV","CRV":"CRV","Código Identificador da Operação de Transportes":"Cód. Operação",
                 "Nota Fiscal":"Nota Fiscal","Preventiva Equipamento":"Preventiva"}
def is_prontuario(path):
    try: wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception: return False
    hit=False
    for ws in wb.worksheets:
        for r in ws.iter_rows(min_row=1,max_row=5,values_only=True):
            j=" ".join(str(c).strip().lower() for c in r if c is not None)
            if "inspeção de segurança amazon" in j or "inspecao de seguranca amazon" in j: hit=True; break
        if hit: break
    wb.close(); return hit
def parse_prontuario(path):
    wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws=None
    for w in wb.worksheets:
        if "prontu" in w.title.lower(): ws=w; break
    if ws is None: ws=max(wb.worksheets, key=lambda w:(w.max_row or 0))
    rows=list(ws.iter_rows(values_only=True)); wb.close()
    name_ri=None
    for ri in range(min(8,len(rows))):
        if any(str(v).strip().lower()=="identificador" for v in rows[ri] if v is not None): name_ri=ri; break
    if name_ri is None: return []
    lvl3_ri=name_ri+1
    ncol=max(len(r) for r in rows[:lvl3_ri+1])
    def cell(ri,ci): return str(rows[ri][ci]).strip() if (ci<len(rows[ri]) and rows[ri][ci] is not None) else ""
    name={ci:cell(name_ri,ci) for ci in range(ncol)}; lvl3={ci:cell(lvl3_ri,ci) for ci in range(ncol)}
    def fn(lbl,lvl=None):
        for ci in range(ncol):
            if name[ci].lower()==lbl.lower() and (lvl is None or lvl3[ci].lower()==lvl.lower()): return ci
        return None
    def fnc(sub,lvl=None):
        for ci in range(ncol):
            if sub.lower() in name[ci].lower() and (lvl is None or lvl3[ci].lower()==lvl.lower()): return ci
        return None
    c_ident=fn("Identificador"); c_transp=fn("Transportador"); c_tipo=fn("Tipo de Equipamento")
    c_cad=fn("Data Cadastro") if fn("Data Cadastro") is not None else fnc("Data Cadastro")
    doc_cols={d:fn(d,"Situação") for d in PRONT_DOCS}
    c_seg=fnc("Inspeção de Segurança","Situação"); c_pre=fnc("Pré Uso","Situação")
    # data de geracao do relatorio (posicao do snapshot): 1a celula de data da 1a linha ("Geração:")
    asof=""
    if rows:
        for v in rows[0]:
            if isinstance(v,(datetime,date)): asof=v.strftime("%Y-%m-%d"); break
    def gdate(r,c):
        if c is None or c>=len(r): return ""
        v=r[c]
        if isinstance(v,(datetime,date)): return v.strftime("%Y-%m-%d")
        s=str(v).strip() if v is not None else ""
        m=re.search(r'(\d{4})-(\d{2})-(\d{2})',s)
        if m: return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        m=re.search(r'(\d{2})/(\d{2})/(\d{4})',s)
        if m: return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
        return ""
    start=lvl3_ri+1
    while start<len(rows) and not looks_plate(rows[start][c_ident] if (c_ident is not None and c_ident<len(rows[start])) else None): start+=1
    def g(r,c): return (str(r[c]).strip() if (c is not None and c<len(r) and r[c] is not None) else "")
    recs=[]
    for r in rows[start:]:
        if c_ident is None or c_ident>=len(r) or not looks_plate(r[c_ident]): continue
        docs={PRONT_DOC_SHORT[d]:g(r,doc_cols[d]) for d in PRONT_DOCS}
        recs.append({"ident":g(r,c_ident).upper(),"transp":g(r,c_transp) or "(sem transp.)","tipo":g(r,c_tipo) or "(?)",
                     "docs":docs,"seg":g(r,c_seg),"pre":g(r,c_pre),"cad":gdate(r,c_cad)})
    return recs, asof

# ---- EQUIPAMENTOS (combustível) ----
def is_equipamentos(path):
    try: wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception: return False
    hit=False
    for ws in wb.worksheets:
        for r in ws.iter_rows(min_row=1,max_row=4,values_only=True):
            j=" ".join(str(c).strip().lower() for c in r if c is not None)
            if "veículo a diesel?" in j or "veiculo a diesel?" in j: hit=True; break
        if hit: break
    wb.close(); return hit
def parse_equipamentos(path):
    wb=openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws=wb["Equipamentos"] if "Equipamentos" in wb.sheetnames else max(wb.worksheets,key=lambda w:(w.max_row or 0))
    rows=list(ws.iter_rows(values_only=True)); wb.close()
    H=min(5,len(rows))
    def findcol(pred):
        for ci in range(max(len(r) for r in rows[:H])):
            for ri in range(H):
                if ci<len(rows[ri]) and rows[ri][ci] is not None and pred(str(rows[ri][ci]).strip().lower()): return ci
        return None
    c_ident=findcol(lambda s:s=="identificador"); c_transp=findcol(lambda s:s=="transportador")
    c_tipo=findcol(lambda s:s=="tipo equipamento")
    c_el=findcol(lambda s:"elétric" in s or "eletric" in s); c_bio=findcol(lambda s:"biometano" in s); c_di=findcol(lambda s:"diesel" in s)
    c_sit=findcol(lambda s:"situação equipamento" in s or "situacao equipamento" in s)
    def val(r,c): return str(r[c]).strip() if (c is not None and c<len(r) and r[c] is not None) else ""
    def sim(x): return x.lower()=="sim"
    recs=[]
    for r in rows:
        if c_ident is None or c_ident>=len(r) or not looks_plate(r[c_ident]): continue
        el,bio,di=sim(val(r,c_el)),sim(val(r,c_bio)),sim(val(r,c_di))
        if el: fuel="Elétrico"
        elif bio: fuel="Biometano"
        elif di: fuel="Diesel"
        else: fuel="Sem informação"
        recs.append({"ident":val(r,c_ident).upper(),"transp":val(r,c_transp) or "(sem transp.)","tipo":val(r,c_tipo) or "(?)",
                     "fuel":fuel,"el":el,"bio":bio,"di":di,"sit":val(r,c_sit)})
    return recs

# ---------------- agregações das novas abas ----------------
def build_plc_info(insp, insp_seg, pront):
    """Aba Placas & Documentos. insp=pré-uso(banco), insp_seg=auditor, pront=prontuário."""
    carriers=[]; cix={}
    def cidx(x):
        x=x or "(sem transp.)"
        if x not in cix: cix[x]=len(carriers); carriers.append(x)
        return cix[x]
    # cobertura por placa
    pre_by_car=defaultdict(set); aud_by_car=defaultdict(set)
    pre_set=set(); aud_set=set()
    for r in insp:
        p=str(r.get("ident","")).strip().upper()
        if not p: continue
        pre_set.add(p); pre_by_car[cidx(r.get("transp"))].add(p)
    for r in insp_seg:
        p=str(r.get("ident","")).strip().upper()
        if not p: continue
        aud_set.add(p); aud_by_car[cidx(r.get("transp"))].add(p)
    allc=set(pre_by_car)|set(aud_by_car)
    byCarrier=[[c,len(pre_by_car.get(c,set())),len(aud_by_car.get(c,set()))] for c in allc]
    byCarrier.sort(key=lambda x:-(x[1]+x[2]))
    # documentos (prontuário)
    doc_types=[PRONT_DOC_SHORT[d] for d in PRONT_DOCS]
    doc_ok=[0]*len(doc_types); doc_no=[0]*len(doc_types)
    docByCar=defaultdict(lambda:[0]+[0]*len(doc_types))  # [totalPlacas, okType0, okType1,...]
    docRows=[]; seg_ok=0; pre_ok=0
    def _split_transp(t):
        parts=[p.strip() for p in str(t or "").split(",")]; seen=set(); out=[]
        for p in parts:
            if not p: continue
            k=p.upper()
            if k not in seen: seen.add(k); out.append(p)
        return out or ["(sem transp.)"]
    for r in pront:
        cars=_split_transp(r.get("transp"))
        sits=[]; okf=[]
        for j,dt in enumerate(doc_types):
            s=r["docs"].get(dt,"")
            sits.append(s); ok=(s.upper()=="OK"); okf.append(ok)
            if ok: doc_ok[j]+=1          # totais globais: 1x por placa
            elif s: doc_no[j]+=1
        if r.get("seg","").upper()=="OK": seg_ok+=1
        if r.get("pre","").upper()=="OK": pre_ok+=1
        for c in cars:                    # placa em varias transportadoras -> conta em cada uma (virgula = separador)
            ci=cidx(c); agg=docByCar[ci]; agg[0]+=1
            for j,ok in enumerate(okf):
                if ok: agg[1+j]+=1
            docRows.append([r["ident"],ci,r.get("tipo","")]+sits+[r.get("seg",""),r.get("pre","")])
    docByCarrier=[[c]+v for c,v in docByCar.items()]
    docByCarrier.sort(key=lambda x:-x[1])
    return {"carriers":carriers,
            "cov":{"pre":len(pre_set),"aud":len(aud_set),"both":len(pre_set&aud_set),
                   "only_pre":len(pre_set-aud_set),"only_aud":len(aud_set-pre_set)},
            "byCarrier":byCarrier,
            "doc":{"types":doc_types,"ok":doc_ok,"no":doc_no,"total_placas":len(pront),
                   "seg_ok":seg_ok,"pre_ok":pre_ok,"byCarrier":docByCarrier,"rows":docRows}}

def build_fuel_info(equip):
    carriers=[]; cix={}
    def cidx(x):
        x=x or "(sem transp.)"
        if x not in cix: cix[x]=len(carriers); carriers.append(x)
        return cix[x]
    cats=["Diesel","Biometano","Elétrico","Sem informação"]
    catix={c:i for i,c in enumerate(cats)}
    dist=[0,0,0,0]; flagSim={"Diesel":0,"Biometano":0,"Elétrico":0}
    byCar=defaultdict(lambda:[0,0,0,0]); rows=[]
    for r in equip:
        f=r["fuel"]; k=catix[f]
        dist[k]+=1                       # distribuicao global: 1x por equipamento (placa)
        if r.get("di"): flagSim["Diesel"]+=1
        if r.get("bio"): flagSim["Biometano"]+=1
        if r.get("el"): flagSim["Elétrico"]+=1
        for c in _split_carriers(r.get("transp")):   # placa em varias transportadoras -> conta em cada uma (sem linha combinada)
            ci=cidx(c); byCar[ci][k]+=1; rows.append([r["ident"],ci,r.get("tipo",""),k])
    byCarrier=[[c]+v for c,v in byCar.items()]
    return {"carriers":carriers,"cats":cats,"dist":dist,"total":len(equip),
            "flagSim":flagSim,"byCarrier":byCarrier,"rows":rows}

# ---- PLACAS ATIVAS (FATURAVEIS) — fonte: Checklists (inspecoes) do Frota Legal ----
ATIVAS_EXCLUI_TIPOS={_norm_tipo(x) for x in ["VEÍCULO TESTE","TESTE"]}
def build_ativas_info(insp, insp_seg):
    """Aba 'Placas ativas (faturaveis)'. Criterio da medicao (definicao do cliente):
    Placa ativa no mes de referencia = veiculo DISTINTO (por Identificador) que, NAQUELE mes, teve:
      (A) checklist de pre-viagem ('Pre Uso Motorista') — base de pre-uso (insp); OU
      (B) checklist feito pelo colaborador Miguel (Miguel de Sousa Cardoso) — inspecao de auditor (insp_seg).
    Regras: 1 por placa por mes (varios checklists no mes = 1); so o mes de referencia (checklist com
    data em outro mes nao conta); exclui APENAS os tipos VEICULO TESTE e TESTE (CARRETA e SEMI REBOQUE
    BI-TREM contam — entram pela inspecao do Miguel). Fonte = export de Checklists, NAO o Prontuario.
    origem (bitmask): 1 = pre-viagem, 2 = Miguel; 3 = os dois."""
    agg={}   # (ident, (ano,mes)) -> {"tipo":str,"carriers":set,"origin":int}
    def add(ident, tipo, transp, diso, origin):
        if _norm_tipo(tipo) in ATIVAS_EXCLUI_TIPOS: return
        if _only_demarco(transp): return
        ident=str(ident or "").strip().upper()
        if not ident: return
        try: dt=datetime.fromisoformat(diso)
        except Exception: return
        key=(ident,(dt.year,dt.month)); e=agg.get(key)
        if e is None: e={"tipo":tipo or "(?)","carriers":set(),"origin":0}; agg[key]=e
        e["origin"]|=origin
        if (e["tipo"]=="(?)" or not e["tipo"]) and tipo: e["tipo"]=tipo
        for c in _split_carriers(_clean_carrier(transp) or "(sem transp.)"): e["carriers"].add(c)
    for r in insp:                                   # (A) pre-uso motorista (toda a base de pre-uso)
        add(r.get("ident",""), r.get("tipo","(?)"), r.get("transp",""), r.get("data",""), 1)
    for r in insp_seg:                               # (B) checklist do Miguel (inspecao de auditor)
        if r.get("user","")!=MIGUEL: continue
        add(r.get("ident",""), r.get("tipo","(?)"), r.get("transp",""), r.get("data",""), 2)
    ymset=sorted({ym for (_id,ym) in agg})
    months=[f"{MESES[m]}/{y}" for (y,m) in ymset]; midx={ym:i for i,ym in enumerate(ymset)}
    carriers=[]; cix={}; tipos=[]; tix={}
    def cidx(x):
        if x not in cix: cix[x]=len(carriers); carriers.append(x)
        return cix[x]
    def tidx(x):
        if x not in tix: tix[x]=len(tipos); tipos.append(x)
        return tix[x]
    rows=[]
    for (ident,ym),e in agg.items():
        mi=midx[ym]; ti=tidx(e["tipo"])
        for c in e["carriers"]:                      # placa em varias transportadoras -> 1 linha por transp
            rows.append([ident, cidx(c), ti, mi, e["origin"]])
    return {"months":months,"carriers":carriers,"tipos":tipos,"rows":rows,
            "total":len({k[0] for k in agg})}

def build_resumo(detail, insp_seg, insp_info, placas_info, plc_info, fuel_info, ativas_info=None):
    from collections import Counter as _C
    inv={v:k for k,v in MESES.items()}
    def mkey(m):
        try: ab,yr=m.split("/"); return (int(yr),inv.get(ab,0))
        except: return (0,0)
    meses=sorted({d[12] for d in detail if d[12]}, key=mkey)
    import datetime as _dt; _now=_dt.date.today()
    _past=[m for m in meses if mkey(m)<=(_now.year,_now.month)]   # nao deixa viagens agendadas p/ o futuro definirem o mes corrente
    cm=(_past[-1] if _past else (meses[-1] if meses else ""))
    c=_C(d[11] for d in detail); total=len(detail)
    def mlabel_iso(s):
        try: return f"{MESES[int(s[5:7])]}/{int(s[:4])}"
        except: return ""
    audCM=len({r["ident"] for r in insp_seg if r.get("data") and mlabel_iso(r["data"])==cm})
    placasVridCM=len({npl(d[5]) for d in detail if d[12]==cm and d[5]})
    doc=plc_info["doc"]; nt=len(doc["types"])
    comDep=len({r[0] for r in doc["rows"] if any(str(r[3+j]).upper()=="OK" for j in range(nt))})
    agg=insp_info["agg"]; nokr=insp_info["nok_rows"]; cat=placas_info["cat"]
    pa={"mes":"","total":0,"preuso":0,"miguel":0,"soMiguel":0}
    if ativas_info and ativas_info.get("months"):
        lastIdx=len(ativas_info["months"])-1
        tot=set(); pre=set(); mig=set()
        for r in ativas_info["rows"]:
            if r[3]!=lastIdx: continue
            tot.add(r[0])
            if r[4]&1: pre.add(r[0])
            if r[4]&2: mig.add(r[0])
        pa={"mes":ativas_info["months"][lastIdx],"total":len(tot),"preuso":len(pre),
            "miguel":len(mig),"soMiguel":len(mig-pre)}
    return {"mesCorrente":cm,"placasAtivas":pa,
      "preViagem":{"total":total,"aderente":c[0],"fora":c[1],"sem":c[2],"semdata":c[3]},
      "auditoria":{"mes":cm,"inspecionados":audCM,"placasVrid":placasVridCM,"totalAudit":len({r["ident"] for r in insp_seg}),"totalAuditorias":len(insp_seg),"noSistema":doc["total_placas"]},
      "documentacao":{"noSistema":doc["total_placas"],"comDeposito":comDep},
      "anomalias":{"total":len(agg),"comAnomalia":sum(1 for a in agg if a[4]==1),"itensNok":sum(len(r[8]) for r in nokr),"transp":len({r[2] for r in nokr})},
      "semEsp":{"total":placas_info["total"],"comInsp":cat.get("aderente",0)+cat.get("fora",0),"semInsp":cat.get("sem",0),"fora":cat.get("fora",0)},
      "combustivel":{"total":fuel_info["total"],"sem":fuel_info["dist"][3],"diesel":fuel_info["dist"][0],"biometano":fuel_info["dist"][1],"eletrico":fuel_info["dist"][2]}}

def build_bases(b, insp_info):
    """Extratos compactos das bases para download no dashboard."""
    carriers=[]; cix={}; tipos=[]; tix={}
    def cidx(x):
        x=x or "(sem transp.)"
        if x not in cix: cix[x]=len(carriers); carriers.append(x)
        return cix[x]
    def tidx(x):
        x=x or "(?)"
        if x not in tix: tix[x]=len(tipos); tipos.append(x)
        return tix[x]
    pre=[]
    for r in b.get("insp",[]):
        pre.append([r.get("ident",""), tidx(r.get("tipo")), cidx(r.get("transp")), r.get("user",""),
                    (r.get("data","") or "").replace("T"," ")[:16], r.get("carreta",""), r.get("vrid",""),
                    r.get("transpv",""), len(r.get("nok",[]))])
    aud=[]
    for r in b.get("insp_seg",[]):
        aud.append([r.get("ident",""), r.get("tipo",""), r.get("transp",""), r.get("user",""),
                    (r.get("data","") or "").replace("T"," ")[:16], len(r.get("nok",[])), " | ".join(r.get("nok",[]))])
    esp=[]
    for t in b.get("placas",[]):
        tiso=t[3] if len(t)>3 else ""
        esp.append([t[0], t[1] if len(t)>1 else "", t[2] if len(t)>2 else "",
                    (tiso or "").replace("T"," ")[:16], t[4] if len(t)>4 else ""])
    return {"preuso":{"carriers":carriers,"tipos":tipos,"rows":pre},"auditor":{"rows":aud},"esp":{"rows":esp}}

# ---------------- BANCO DE DADOS (persistente + ingestao incremental) ----------------
import csv as _csv
def _ensure_dirs():
    for d in (BANCO, ENTRADAS, PLACAS_ENT): os.makedirs(d, exist_ok=True)
def _fsig(path):
    st=os.stat(path); return f"{os.path.basename(path)}|{st.st_size}"
def load_banco():
    _ensure_dirs(); b={"viagens":[],"insp":[],"itens":{},"placas":[],"manifest":set(),"insp_seg":[],"prontuario":[],"equip":[],"meta":{}}
    fp=os.path.join(BANCO,"viagens.csv")
    if os.path.exists(fp):
        with open(fp,newline='',encoding='utf-8') as f: b["viagens"]=[tuple(r) for r in _csv.reader(f)]
    fp=os.path.join(BANCO,"inspecoes.jsonl")
    if os.path.exists(fp):
        with open(fp,encoding='utf-8') as f:
            for ln in f:
                ln=ln.strip()
                if ln: b["insp"].append(json.loads(ln))
    fp=os.path.join(BANCO,"itens.json")
    if os.path.exists(fp): b["itens"]=json.load(open(fp,encoding='utf-8'))
    fp=os.path.join(BANCO,"placas.csv")
    if os.path.exists(fp):
        with open(fp,newline='',encoding='utf-8') as f: b["placas"]=[tuple(r) for r in _csv.reader(f)]
    fp=os.path.join(BANCO,"insp_seg.jsonl")
    if os.path.exists(fp):
        with open(fp,encoding='utf-8') as f:
            for ln in f:
                ln=ln.strip()
                if ln: b["insp_seg"].append(json.loads(ln))
    fp=os.path.join(BANCO,"prontuario.json")
    if os.path.exists(fp): b["prontuario"]=json.load(open(fp,encoding='utf-8'))
    fp=os.path.join(BANCO,"equip.json")
    if os.path.exists(fp): b["equip"]=json.load(open(fp,encoding='utf-8'))
    fp=os.path.join(BANCO,"meta.json")
    if os.path.exists(fp): b["meta"]=json.load(open(fp,encoding='utf-8'))
    fp=os.path.join(BANCO,"manifest.json")
    if os.path.exists(fp): b["manifest"]=set(json.load(open(fp,encoding='utf-8')))
    return b
def save_banco(b):
    _ensure_dirs()
    with open(os.path.join(BANCO,"viagens.csv"),"w",newline='',encoding='utf-8') as f: _csv.writer(f).writerows(b["viagens"])
    with open(os.path.join(BANCO,"inspecoes.jsonl"),"w",encoding='utf-8') as f:
        for rec in b["insp"]: f.write(json.dumps(rec,ensure_ascii=False)+"\n")
    json.dump(b["itens"], open(os.path.join(BANCO,"itens.json"),"w",encoding='utf-8'), ensure_ascii=False)
    with open(os.path.join(BANCO,"placas.csv"),"w",newline='',encoding='utf-8') as f: _csv.writer(f).writerows(b["placas"])
    with open(os.path.join(BANCO,"insp_seg.jsonl"),"w",encoding='utf-8') as f:
        for rec in b.get("insp_seg",[]): f.write(json.dumps(rec,ensure_ascii=False)+"\n")
    json.dump(b.get("prontuario",[]), open(os.path.join(BANCO,"prontuario.json"),"w",encoding='utf-8'), ensure_ascii=False)
    json.dump(b.get("equip",[]), open(os.path.join(BANCO,"equip.json"),"w",encoding='utf-8'), ensure_ascii=False)
    json.dump(b.get("meta",{}), open(os.path.join(BANCO,"meta.json"),"w",encoding='utf-8'), ensure_ascii=False)
    json.dump(sorted(b["manifest"]), open(os.path.join(BANCO,"manifest.json"),"w",encoding='utf-8'), ensure_ascii=False)
def ingest(b):
    """Ingesta arquivos novos (fora do manifest). Viagens/Inspecoes em Entradas; Placas em PLACAS_ENT."""
    _ensure_dirs()
    vsig={(t[0],t[5],t[2],t[6]) for t in b["viagens"]}
    isig={(r["ident"].strip().upper(), r["data"], r["vrid"].strip().upper()) for r in b["insp"]}
    psig={(t[0],t[3],t[4]) for t in b["placas"]}
    added={"viagens":0,"inspecoes":0,"placas":0,"arquivos":[]}
    # Entradas: viagens e inspecoes
    for f in sorted(glob.glob(os.path.join(ENTRADAS,"*.xlsx"))):
        if os.path.basename(f).startswith("~$"): continue
        sig=_fsig(f)
        if sig in b["manifest"]: continue
        k=sniff(f)
        if k=="viagens":
            recs,_=normalize_trips(f)
            for rec in recs:
                tiso=rec[5].isoformat() if isinstance(rec[5],datetime) else ""
                row=(rec[0],rec[1],rec[2],rec[3],rec[4],tiso,rec[6])
                key=(row[0],row[5],row[2],row[6])
                if key in vsig: continue
                vsig.add(key); b["viagens"].append(row); added["viagens"]+=1
        elif k=="inspecoes":
            recs,ig=parse_insp_file(f)
            for nm,gp in ig.items():
                if nm not in b["itens"]: b["itens"][nm]=gp
            for rec in recs:
                key=(rec["ident"].strip().upper(), rec["data"], rec["vrid"].strip().upper())
                if key in isig: continue
                isig.add(key); b["insp"].append(rec); added["inspecoes"]+=1
        else:
            continue
        b["manifest"].add(sig); added["arquivos"].append(os.path.basename(f)); save_banco(b)
    # Placas (espelhamento). O export da Amazon alimenta a aba Sem Espelhamento E TAMBEM
    # as viagens (todas as demais abas), junto com o arquivo do Frota Legal.
    for f in sorted(glob.glob(os.path.join(PLACAS_ENT,"*.xlsx"))):
        if os.path.basename(f).startswith("~$"): continue
        sig=_fsig(f)
        if sig in b["manifest"]: continue
        for row in parse_placas_file(f):
            key=(row[0],row[3],row[4])
            if key in psig: continue
            psig.add(key); b["placas"].append(row); added["placas"]+=1
        # tambem ingere como VIAGENS quando o arquivo tem cabecalho de viagens
        if sniff(f)=="viagens":
            recs,_=normalize_trips(f)
            for rec in recs:
                tiso=rec[5].isoformat() if isinstance(rec[5],datetime) else ""
                vrow=(rec[0],rec[1],rec[2],rec[3],rec[4],tiso,rec[6])
                vkey=(vrow[0],vrow[5],vrow[2],vrow[6])
                if vkey in vsig: continue
                vsig.add(vkey); b["viagens"].append(vrow); added["viagens"]+=1
        b["manifest"].add(sig); added["arquivos"].append(os.path.basename(f)); save_banco(b)
    _ingest_novas_fontes(b, added)
    return added

def _newest(paths):
    return sorted(paths, key=os.path.getmtime, reverse=True)[0] if paths else None

def _ingest_novas_fontes(b, added):
    xs=[f for f in glob.glob(os.path.join(ENTRADAS,"*.xlsx")) if not os.path.basename(f).startswith("~$")]
    meta=b.setdefault("meta",{}); kc=meta.setdefault("kinds",{})
    def kind_of(f):
        sig=_fsig(f)
        if sig in kc: return kc[sig]
        k=sniff(f)
        if k is None:
            if is_prontuario(f): k="prontuario"
            elif is_equipamentos(f): k="equipamentos"
        kc[sig]=k; return k
    kinds={f:kind_of(f) for f in xs}
    # AUDITOR (abas Inspeção de Segurança dos arquivos de inspeção)
    insp_files=sorted(os.path.basename(f) for f,k in kinds.items() if k=="inspecoes")
    if insp_files!=meta.get("insp_seg_files") or not b.get("insp_seg"):
        seg=[]; seen=set()
        for f,k in kinds.items():
            if k!="inspecoes": continue
            for rec in parse_auditor_file(f):
                key=(rec["ident"], rec["data"])
                if key in seen: continue
                seen.add(key); seg.append(rec)
        b["insp_seg"]=seg; meta["insp_seg_files"]=insp_files; added["auditor"]=len(seg); save_banco(b)
    # PRONTUÁRIO (snapshot mais recente)
    pr=_newest([f for f,k in kinds.items() if k=="prontuario"])
    if pr:
        sig=_fsig(pr)
        if meta.get("prontuario_sig")!=sig:
            b["prontuario"],_pasof=parse_prontuario(pr); meta["prontuario_sig"]=sig; meta["prontuario_asof"]=_pasof
            added["prontuario"]=len(b["prontuario"]); added["arquivos"].append(os.path.basename(pr)); save_banco(b)
    # EQUIPAMENTOS combustível (snapshot mais recente)
    eq=_newest([f for f,k in kinds.items() if k=="equipamentos"])
    if eq:
        sig=_fsig(eq)
        if meta.get("equip_sig")!=sig:
            b["equip"]=parse_equipamentos(eq); meta["equip_sig"]=sig
            added["equipamentos"]=len(b["equip"]); added["arquivos"].append(os.path.basename(eq)); save_banco(b)
    return added

def trips_from_banco(b):
    """Reconstroi trips (dedup 1 por VRID) e week_of a partir do banco de viagens."""
    pairs=[]; recs=[]
    for row in b["viagens"]:
        v,car,route,lane,wy,tiso,pl=row
        t=datetime.fromisoformat(tiso) if tiso else None
        recs.append((v,car,route,lane,wy,t,pl))
        if wy and t: pairs.append((wy,t.date()))
    week_of=build_week_of(pairs)
    groups=defaultdict(list)
    for rec in recs: groups[rec[0]].append(rec)
    trips=[_first_leg(rs) for rs in groups.values()]
    return trips, week_of

# ---------------- classificacao ----------------
def pick(items, T):
    before=[x for x in items if x[0]<=T]
    if not before: return None
    dt,nok=before[-1]; hb=(T-dt).total_seconds()/3600
    sc = 0 if hb<=12 else (1 if hb<=24 else 2)
    return sc,dt,hb,nok

def build_detail(trips, ci, iv, ip, week_of=None):
    detail=[]
    for t in trips:
        def g(k):
            i=ci[k]; return t[i] if (i is not None and i<len(t)) else None
        vrid=nvr(g("vrid")); carrier=(str(g("carrier") or "").strip()) or "(sem transp.)"
        route=str(g("route") or ""); T=g("time"); plate=str(g("plate") or ""); pl=npl(plate)
        week=(str(g("week") or "").strip())
        if not week and isinstance(T,datetime) and week_of: week=week_of(T.date())  # deriva semana p/ bases sem week_year
        if not week: week="(sem sem.)"
        month=""
        if isinstance(T,datetime): month=f"{MESES[T.month]}/{T.year}"
        if not isinstance(T,datetime):
            sc,method,best,hb,nok=3,"-",None,None,None
        else:
            r=None; method="-"
            if vrid in iv: r=pick(iv[vrid],T); method="VRID"
            if r is None and pl in ip: r=pick(ip[pl],T); method="Placa"
            if r is None: sc,best,hb,nok=2,None,None,None; method="-"
            else: sc,best,hb,nok=r
        anom = ("NOK" if nok else "OK") if (best is not None) else ""
        detail.append([str(g("vrid") or ""), carrier, route, week,
                       T if isinstance(T,datetime) else None, plate, LBL[sc], method,
                       best, (round(hb,1) if hb is not None else None), anom, sc, month])
    return detail

# ================= saida XLSX =================
def _xlsx_novas_abas(wb, plc_info, fuel_info):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    NAVY="232F3E"; FN="Arial"
    hf=PatternFill("solid",fgColor=NAVY); hfont=Font(name=FN,bold=True,color="FFFFFF",size=10)
    thin=Side(style="thin",color="D9D9D9"); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    def hc(c,t): c.value=t; c.fill=hf; c.font=hfont; c.alignment=Alignment(horizontal="center",vertical="center"); c.border=bd
    def title(ws,t,sub=""):
        ws.sheet_view.showGridLines=False
        ws["A1"]=t; ws["A1"].font=Font(name=FN,bold=True,size=15,color=NAVY)
        if sub: ws["A2"]=sub; ws["A2"].font=Font(name=FN,size=10,color="666666")

    # ---------- Placas & Auditoria ----------
    if plc_info:
        ws=wb.create_sheet("Placas & Auditoria")
        title(ws,"Placas — check-list pré-viagem × inspeção de auditor",
              "Auditor = inspeções do tipo 'Inspeção de Segurança'. Contagem de placas distintas.")
        cov=plc_info["cov"]; r=4
        for i,(lbl,v) in enumerate([("Placas com check-list pré-viagem",cov["pre"]),
                                    ("Placas com inspeção de auditor",cov["aud"]),
                                    ("Placas com ambos",cov["both"]),
                                    ("Só pré-viagem",cov["only_pre"]),
                                    ("Só auditor",cov["only_aud"])]):
            a=ws.cell(r+i,1,lbl); b=ws.cell(r+i,2,v)
            a.font=Font(name=FN,size=10,bold=(i<3)); b.font=Font(name=FN,size=10,bold=(i<3)); b.number_format="#,##0"; a.border=bd; b.border=bd
        hr=r+6
        for c,t in [(1,"Transportadora"),(2,"Placas pré-viagem"),(3,"Placas auditor")]: hc(ws.cell(hr,c),t)
        cn=plc_info["carriers"]
        for j,row in enumerate(plc_info["byCarrier"]):
            rr=hr+1+j
            ws.cell(rr,1,cn[row[0]]).font=Font(name=FN,size=9); ws.cell(rr,1).border=bd
            ws.cell(rr,2,row[1]).number_format="#,##0"; ws.cell(rr,3,row[2]).number_format="#,##0"
            ws.cell(rr,2).border=bd; ws.cell(rr,3).border=bd; ws.cell(rr,2).font=Font(name=FN,size=9); ws.cell(rr,3).font=Font(name=FN,size=9)
        for i,w in enumerate([44,18,16],1): ws.column_dimensions[get_column_letter(i)].width=w
        n=len(plc_info["byCarrier"])
        if n:
            ch=BarChart(); ch.type="bar"; ch.title="Placas por transportadora"
            ch.add_data(Reference(ws,min_col=2,max_col=3,min_row=hr,max_row=hr+n),titles_from_data=True)
            ch.set_categories(Reference(ws,min_col=1,min_row=hr+1,max_row=hr+n))
            ch.height=max(8,min(24,n*0.6+4)); ch.width=16; ws.add_chart(ch,"E4")

        # ---------- Documentos ----------
        doc=plc_info["doc"]; ws=wb.create_sheet("Documentos")
        title(ws,"Documentos incluídos por transportador",
              "Situação 'OK' = documento em dia. Fonte: Prontuário de Equipamentos.")
        types=doc["types"]; hr=4
        hc(ws.cell(hr,1),"Documento"); hc(ws.cell(hr,2),"Placas com OK"); hc(ws.cell(hr,3),"Placas pendentes")
        for j,t in enumerate(types):
            rr=hr+1+j
            ws.cell(rr,1,t).font=Font(name=FN,size=10); ws.cell(rr,1).border=bd
            ws.cell(rr,2,doc["ok"][j]).number_format="#,##0"; ws.cell(rr,3,doc["no"][j]).number_format="#,##0"
            ws.cell(rr,2).border=bd; ws.cell(rr,3).border=bd
        base=hr+len(types)+3
        ws.cell(base-1,1,f"Total de placas no prontuário: {doc['total_placas']:,}".replace(",",".")).font=Font(name=FN,size=10,bold=True)
        hc(ws.cell(base,1),"Transportadora"); hc(ws.cell(base,2),"Placas")
        for j,t in enumerate(types): hc(ws.cell(base,3+j),t)
        cn=plc_info["carriers"]
        for j,row in enumerate(doc["byCarrier"]):
            rr=base+1+j
            ws.cell(rr,1,cn[row[0]]).font=Font(name=FN,size=9); ws.cell(rr,1).border=bd
            ws.cell(rr,2,row[1]).number_format="#,##0"; ws.cell(rr,2).border=bd
            for k in range(len(types)):
                ws.cell(rr,3+k,row[2+k]).number_format="#,##0"; ws.cell(rr,3+k).border=bd
        for i,w in enumerate([44,10]+[13]*len(types),1): ws.column_dimensions[get_column_letter(i)].width=w

    # ---------- Combustível ----------
    if fuel_info:
        ws=wb.create_sheet("Combustível")
        title(ws,"Tipo de combustível sinalizado",
              "Fonte: base de Equipamentos. 'Sem informação' = nenhum tipo sinalizado.")
        cats=fuel_info["cats"]; dist=fuel_info["dist"]; r=4
        hc(ws.cell(r,1),"Tipo"); hc(ws.cell(r,2),"Equipamentos"); hc(ws.cell(r,3),"% do total")
        tot=fuel_info["total"] or 1
        for j,c in enumerate(cats):
            rr=r+1+j
            ws.cell(rr,1,c).font=Font(name=FN,size=10); ws.cell(rr,1).border=bd
            ws.cell(rr,2,dist[j]).number_format="#,##0"; ws.cell(rr,2).border=bd
            ws.cell(rr,3,dist[j]/tot).number_format="0.0%"; ws.cell(rr,3).border=bd
        pie=PieChart(); pie.title="Distribuição por combustível"
        pie.add_data(Reference(ws,min_col=2,min_row=r+1,max_row=r+len(cats)))
        pie.set_categories(Reference(ws,min_col=1,min_row=r+1,max_row=r+len(cats)))
        pie.dataLabels=DataLabelList(); pie.dataLabels.showPercent=True; pie.height=7; pie.width=11
        ws.add_chart(pie,"E4")
        # ranking sem informação por transportador
        semidx=cats.index("Sem informação")
        cn=fuel_info["carriers"]
        rank=sorted(fuel_info["byCarrier"], key=lambda x:-x[1+semidx])
        base=r+len(cats)+3
        ws.cell(base-1,1,"Ranking — equipamentos SEM informação de combustível").font=Font(name=FN,bold=True,size=12,color=NAVY)
        hc(ws.cell(base,1),"Transportadora"); hc(ws.cell(base,2),"Sem informação"); hc(ws.cell(base,3),"Total equip.")
        for j,row in enumerate(rank):
            rr=base+1+j; tote=sum(row[1:])
            ws.cell(rr,1,cn[row[0]]).font=Font(name=FN,size=9); ws.cell(rr,1).border=bd
            ws.cell(rr,2,row[1+semidx]).number_format="#,##0"; ws.cell(rr,2).border=bd
            ws.cell(rr,3,tote).number_format="#,##0"; ws.cell(rr,3).border=bd
        for i,w in enumerate([44,16,13],1): ws.column_dimensions[get_column_letter(i)].width=w
        nr=min(len(rank),15)
        if nr:
            ch=BarChart(); ch.type="bar"; ch.title="Top transportadoras sem informação de combustível"
            ch.add_data(Reference(ws,min_col=2,min_row=base,max_row=base+nr),titles_from_data=True)
            ch.set_categories(Reference(ws,min_col=1,min_row=base+1,max_row=base+nr))
            ch.height=max(8,nr*0.6+3); ch.width=16; ws.add_chart(ch,"E20")

def write_xlsx(detail, out, plc_info=None, fuel_info=None):
    weeks=sorted({d[3] for d in detail})
    def mkey(m):
        try:
            ab,yr=m.split("/"); inv={v:k for k,v in MESES.items()}; return (int(yr),inv.get(ab,0))
        except: return (9999,99)
    months=sorted({d[12] for d in detail if d[12]}, key=mkey)
    ctot=defaultdict(int)
    for d in detail: ctot[d[1]]+=1
    carriers=sorted({d[1] for d in detail}, key=lambda c:-ctot[c])
    NAVY="232F3E"; FN="Arial"
    green=PatternFill("solid",fgColor="C6EFCE"); amber=PatternFill("solid",fgColor="FFEB9C")
    red=PatternFill("solid",fgColor="FFC7CE"); gray=PatternFill("solid",fgColor="E2E8F0")
    noknok=PatternFill("solid",fgColor="F8CBAD")
    hf=PatternFill("solid",fgColor=NAVY); hfont=Font(name=FN,bold=True,color="FFFFFF",size=10)
    thin=Side(style="thin",color="D9D9D9"); bd=Border(left=thin,right=thin,top=thin,bottom=thin)
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Detalhe"
    heads=["VRID","Transportadora","Rota","Semana","Data da viagem","Placa","Status","Match","Data da inspecao","Horas antes","Anomalia"]
    ws.append(heads)
    for c in range(1,len(heads)+1):
        x=ws.cell(1,c); x.fill=hf; x.font=hfont; x.alignment=Alignment(horizontal="center",vertical="center"); x.border=bd
    for d in detail: ws.append(d[:11])
    ws.freeze_panes="A2"; ws.auto_filter.ref=f"A1:K{len(detail)+1}"
    for r in range(2,len(detail)+2):
        ws.cell(r,5).number_format="dd/mm/yyyy hh:mm"; ws.cell(r,9).number_format="dd/mm/yyyy hh:mm"; ws.cell(r,10).number_format="0.0"
    for i,w in enumerate([13,20,16,10,17,11,20,8,17,11,10],1): ws.column_dimensions[get_column_letter(i)].width=w
    for row in ws.iter_rows(min_row=2,max_row=len(detail)+1,min_col=1,max_col=11):
        for c in row: c.font=Font(name=FN,size=9)
    rng=f"G2:G{len(detail)+1}"
    for txt,fill in [(ADER,green),(FORA,amber),(SEMI,red),(SDAT,gray)]:
        ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$G2="{txt}"'], fill=fill))
    ws.conditional_formatting.add(f"K2:K{len(detail)+1}", FormulaRule(formula=['$K2="NOK"'], fill=noknok))
    LAST=len(detail)+1  # ultima linha de dados no Detalhe (para limitar ranges e acelerar recalc)
    # ---- Resumo ----
    rs=wb.create_sheet("Resumo",0); rs.sheet_view.showGridLines=False
    rs["A1"]="Farol de Inspecao Pre-Viagem  -  Amazon"; rs["A1"].font=Font(name=FN,bold=True,size=16,color=NAVY); rs.merge_cells("A1:F1")
    rs["A2"]=f"Periodo {weeks[0]}-{weeks[-1]}  |  {len(detail)} viagens (1 por VRID)"; rs["A2"].font=Font(name=FN,size=10,color="555555"); rs.merge_cells("A2:F2")
    rs["A3"]=f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}"; rs["A3"].font=Font(name=FN,size=9,italic=True,color="888888")
    def hc(cell,txt): cell.value=txt; cell.fill=hf; cell.font=hfont; cell.alignment=Alignment(horizontal="center",vertical="center"); cell.border=bd
    DET="Detalhe"; r=5
    for c,t in [(1,"Indicador"),(2,"Qtde"),(3,"% das avaliaveis")]: hc(rs.cell(r,c),t)
    G_="{0}!$G$2:$G${1}".format(DET,LAST)
    kpis=[("Total de viagens",f'=COUNTA({DET}!$A$2:$A${LAST})',None),(ADER,f'=COUNTIF({G_},"{ADER}")',green),
          (FORA,f'=COUNTIF({G_},"{FORA}")',amber),(SEMI,f'=COUNTIF({G_},"{SEMI}")',red),
          (SDAT,f'=COUNTIF({G_},"{SDAT}")',None)]
    for i,(lbl,frm,fill) in enumerate(kpis):
        rr=r+1+i; a=rs.cell(rr,1,lbl); b=rs.cell(rr,2,frm)
        a.font=Font(name=FN,size=10,bold=(i==0)); b.font=Font(name=FN,size=10,bold=(i==0)); b.number_format="#,##0"; a.border=bd; b.border=bd
        if fill:
            cc=rs.cell(rr,3,f"=B{rr}/($B${r+1}-$B${r+5})"); cc.number_format="0.0%"; cc.font=Font(name=FN,size=10); cc.border=bd; a.fill=fill
        else: rs.cell(rr,3).border=bd
    TOT=r+1; SD=r+5
    # anomalia resumo
    ar=r+7
    rs.cell(ar,1,"Anomalia (inspecoes casadas)").font=Font(name=FN,bold=True,size=11,color=NAVY)
    K_="{0}!$K$2:$K${1}".format(DET,LAST)
    rs.cell(ar+1,1,"OK").font=Font(name=FN,size=10); rs.cell(ar+1,2,f'=COUNTIF({K_},"OK")').number_format="#,##0"
    rs.cell(ar+2,1,"NOK").font=Font(name=FN,size=10); rs.cell(ar+2,2,f'=COUNTIF({K_},"NOK")').number_format="#,##0"
    for rr in (ar+1,ar+2):
        rs.cell(rr,2).font=Font(name=FN,size=10)
    # por semana
    def block(title, key_items, startrow, det_col):
        rs.cell(startrow,1,title).font=Font(name=FN,bold=True,size=12,color=NAVY); h=startrow+1
        wh=["",ADER,FORA,SEMI,SDAT,"Total","% Aderente"]
        for i,x in enumerate(wh,1): hc(rs.cell(h,i),x if x else title.split()[-1])
        for j,k in enumerate(key_items):
            rr=h+1+j; rs.cell(rr,1,k).font=Font(name=FN,size=10); rs.cell(rr,1).border=bd
            for c2,st in enumerate([ADER,FORA,SEMI,SDAT],2):
                cc=rs.cell(rr,c2,f'=COUNTIFS({DET}!${det_col}$2:${det_col}${LAST},$A{rr},{DET}!$G$2:$G${LAST},"{st}")'); cc.number_format="#,##0"; cc.font=Font(name=FN,size=10); cc.border=bd
            rs.cell(rr,6,f"=SUM(B{rr}:E{rr})").number_format="#,##0"; rs.cell(rr,6).font=Font(name=FN,size=10,bold=True); rs.cell(rr,6).border=bd
            rs.cell(rr,7,f"=IFERROR(B{rr}/(F{rr}-E{rr}),0)").number_format="0.0%"; rs.cell(rr,7).font=Font(name=FN,size=10); rs.cell(rr,7).border=bd
        return h+1, h+len(key_items)
    wf,wl=block("Aderencia por semana", weeks, 15, "D")
    mf,ml=block("Aderencia por mes", months, wl+3, "M")
    cf,cl=block("Aderencia por transportadora", carriers, ml+3, "B")
    for i,w in enumerate([26,20,16,14,14,10,12],1): rs.column_dimensions[get_column_letter(i)].width=w
    # ---- graficos com rotulos ----
    pie=PieChart(); pie.title="Distribuicao por status"
    pie.add_data(Reference(rs,min_col=2,min_row=TOT+1,max_row=SD)); pie.set_categories(Reference(rs,min_col=1,min_row=TOT+1,max_row=SD))
    pie.dataLabels=DataLabelList(); pie.dataLabels.showPercent=True; pie.height=7; pie.width=11
    rs.add_chart(pie,"I5")
    def stacked(title, f, l, anchor, h, w):
        b=BarChart(); b.type="col"; b.grouping="stacked"; b.overlap=100; b.title=title
        b.add_data(Reference(rs,min_col=2,max_col=5,min_row=f-1,max_row=l),titles_from_data=True)
        b.set_categories(Reference(rs,min_col=1,min_row=f,max_row=l))
        b.dataLabels=DataLabelList(); b.dataLabels.showVal=True; b.height=h; b.width=w
        rs.add_chart(b,anchor)
    stacked("Viagens por semana", wf, wl, "I20", 7, 13)
    stacked("Viagens por mes", mf, ml, "I36", 7, 13)
    bc=BarChart(); bc.type="bar"; bc.grouping="stacked"; bc.overlap=100; bc.title="Viagens por transportadora"
    bc.add_data(Reference(rs,min_col=2,max_col=5,min_row=cf-1,max_row=cl),titles_from_data=True)
    bc.set_categories(Reference(rs,min_col=1,min_row=cf,max_row=cl))
    bc.dataLabels=DataLabelList(); bc.dataLabels.showVal=True; bc.height=13; bc.width=13
    rs.add_chart(bc,"I52")
    _xlsx_novas_abas(wb, plc_info, fuel_info)
    wb.save(out)

# ================= saida HTML =================
def write_html(detail, insp_info, placas_info, plc_info, fuel_info, bases, resumo, sem_cad, anom_aud, ativas, out):
    tpl=open(os.path.join(SUP,"template.html"),encoding="utf-8").read()
    chart=open(os.path.join(SUP,"chart.umd.js"),encoding="utf-8").read()
    dlab=open(os.path.join(SUP,"datalabels.min.js"),encoding="utf-8").read()
    carriers=[]; ci={}; weeks=[]; wi={}; months=[]; mi={}
    def cidx(c):
        if c not in ci: ci[c]=len(carriers); carriers.append(c)
        return ci[c]
    cat=defaultdict(int); rows=[]; MET={"VRID":0,"Placa":1,"-":2}; AN={"OK":1,"NOK":2,"":0}
    for d in detail:
        sc=d[11]; cat[LBL[sc]]+=1
        wk=d[3];
        if wk not in wi: wi[wk]=len(weeks); weeks.append(wk)
        mo=d[12] or "-"
        if mo not in mi: mi[mo]=len(months); months.append(mo)
        rows.append([d[0], cidx(d[1]), wi[wk], mi[mo],
                     d[4].strftime("%Y-%m-%d %H:%M") if isinstance(d[4],datetime) else "",
                     d[5], sc, MET.get(d[7],2),
                     d[8].strftime("%m-%d %H:%M") if isinstance(d[8],datetime) else "",
                     (d[9] if d[9] is not None else ""), AN.get(d[10],0)])
    # ordena semanas e meses e remapeia indices
    def mkey(m):
        try:
            ab,yr=m.split("/"); inv={v:k for k,v in MESES.items()}; return (int(yr),inv.get(ab,0))
        except: return (9999,99)
    wsorted=sorted(weeks); msorted=sorted([m for m in months], key=mkey)
    wmap={w:i for i,w in enumerate(wsorted)}; mmap={m:i for i,m in enumerate(msorted)}
    for r in rows:
        r[2]=wmap[weeks[r[2]]]; r[3]=mmap[months[r[3]]]
    # ---- inspecoes (abas 2 e 3): ordena meses/semanas e remapeia indices ----
    def _proc(ii):
        ms=sorted(ii["months"], key=mkey); mm=[ms.index(m) for m in ii["months"]]
        ws=sorted(ii["weeks"]); wm=[ws.index(w) for w in ii["weeks"]]
        agg=[[a[0], mm[a[1]], a[2], a[3], a[4], wm[a[5]]] for a in ii["agg"]]
        semR=[[r[0],r[1],r[2],r[3],r[4],r[5],r[6],r[7], mm[r[8]], wm[r[9]]] for r in ii["sem_rows"]]
        nokR=[[r[0],r[1],r[2],r[3],r[4],r[5],r[6], mm[r[7]], r[8], wm[r[9]]] for r in ii["nok_rows"]]
        return {"total":ii["total"],"com_vrid":ii["com_vrid"],"sem_vrid":ii["sem_vrid"],"nok_total":ii["nok_total"],
                "carriers":ii["carriers"],"months":ms,"weeks":ws,"tipos":ii["tipos"],
                "items":ii["items"],"item_groups":ii["item_groups"],
                "agg":agg,"semRows":semR,"nokRows":nokR}
    data={"generated":datetime.now().strftime("%d/%m/%Y %H:%M"),"total":len(rows),
          "cat":{"aderente":cat[ADER],"fora_prazo":cat[FORA],"sem_inspecao":cat[SEMI],"sem_data":cat[SDAT]},
          "carriers":carriers,"weeks":wsorted,"months":msorted,"rows":rows,
          "insp":_proc(insp_info)}
    data["anomAud"]=_proc(anom_aud)          # Melhoria 2: anomalias de auditoria (Miguel)
    data["semCad"]=sem_cad                    # Melhoria 1: veiculos sem cadastro (ASDFG)
    # ---- aba 4: Sem Espelhamento ---- (row: [vrid, carrierIdx, route, dataStr, sc, weekIdx, monthIdx])
    P=placas_info
    pwsorted=sorted([w for w in P["weeks"]]); pmsorted=sorted([m for m in P["months"]], key=mkey)
    pwmap={P["weeks"].index(w):i for i,w in enumerate(pwsorted)}
    pmmap={P["months"].index(m):i for i,m in enumerate(pmsorted)}
    prows=[[r[0],r[1],r[2],r[3],r[4],(pwmap[r[5]] if r[5]>=0 else -1),(pmmap[r[6]] if r[6]>=0 else -1)] for r in P["rows"]]
    data["esp"]={"total":P["total"],"cat":P["cat"],"carriers":P["carriers"],
                 "weeks":pwsorted,"months":pmsorted,"rows":prows}
    data["plc"]=plc_info; data["fuel"]=fuel_info; data["bases"]=bases; data["resumo"]=resumo; data["ativas"]=ativas
    html=(tpl.replace("__DATA__",json.dumps(data,ensure_ascii=False,separators=(",",":")))
             .replace("__CHARTJS__",chart).replace("__DATALABELS__",dlab))
    open(out,"w",encoding="utf-8").write(html)

def recalc_xlsx(path):
    import shutil, subprocess, tempfile
    soffice=shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice: log("Aviso: LibreOffice ausente; Excel recalculara ao abrir."); return
    try:
        with tempfile.TemporaryDirectory() as td:
            env=dict(os.environ); env["HOME"]=td
            subprocess.run([soffice,"--headless","--calc","--convert-to","xlsx:Calc MS Excel 2007 XML","--outdir",td,path],
                           check=True,capture_output=True,timeout=150,env=env)
            conv=os.path.join(td, os.path.splitext(os.path.basename(path))[0]+".xlsx")
            if os.path.exists(conv): shutil.copyfile(conv,path); log("Excel recalculado.")
    except Exception as e: log("Aviso: recalc falhou (",e,")")

def _only_demarco(name):
    parts=[p.strip() for p in str(name or "").split(",") if p.strip()]
    return len(parts)>0 and all(p.lower()=="demarco" for p in parts)
def _clean_carrier(name):
    parts=[p.strip() for p in str(name or "").split(",")]
    kept=[p for p in parts if p and p.strip().lower()!="demarco"]
    return ", ".join(kept)
def apply_filters(b):
    """Filtro aplicado na GERACAO (nao-destrutivo no banco):
       - remove veiculos INATIVOS (Equipamentos: Situacao == 'Inativo') pela placa;
       - remove a empresa de teste DEMARCO de TODAS as bases (inclusive limpando de combinacoes)."""
    INACT={npl(r.get("ident","")) for r in b.get("equip",[]) if str(r.get("sit","")).strip().lower()=="inativo"}
    def keep_plate(pl): return npl(pl) not in INACT
    bf=dict(b)
    sem_cad=[]  # Melhoria 1: registros da placa ficticia ASDFG (veiculo sem cadastro)
    def _cln(recs, plate_key, split_sc=False):
        out=[]
        for r in recs:
            # ASDFG e' isento da regra DEMARCO/inativo: coletado a' parte, nao entra no farol normal.
            if split_sc and str(r.get("ident","")).strip().upper()==SEM_CADASTRO_PLATE:
                sem_cad.append(dict(r)); continue
            if _only_demarco(r.get("transp")): continue
            if not keep_plate(r.get(plate_key,"")): continue
            r=dict(r); r["transp"]=_clean_carrier(r.get("transp")) or "(sem transp.)"; out.append(r)
        return out
    bf["insp"]=_cln(b.get("insp",[]),"ident")
    # Auditoria: considera somente inspecoes a partir de AUD_MIN_DATE (01/08/2026); anteriores sao descartadas.
    def _aud_after_cutoff(r):
        try: return datetime.fromisoformat(r.get("data","")) >= AUD_MIN_DATE
        except Exception: return False
    _seg_src=[r for r in b.get("insp_seg",[]) if _aud_after_cutoff(r)]
    bf["insp_seg"]=_cln(_seg_src,"ident",split_sc=True)
    bf["sem_cad"]=sem_cad
    bf["prontuario"]=_cln(b.get("prontuario",[]),"ident")
    eq=[]
    for r in b.get("equip",[]):
        if str(r.get("sit","")).strip().lower()=="inativo": continue
        if _only_demarco(r.get("transp")): continue
        r=dict(r); r["transp"]=_clean_carrier(r.get("transp")) or "(sem transp.)"; eq.append(r)
    bf["equip"]=eq
    vg=[]
    for row in b.get("viagens",[]):
        v,car,route,lane,wy,tiso,pl=row
        if not keep_plate(pl): continue
        if _only_demarco(car): continue
        vg.append((v,_clean_carrier(car) or "(sem transp.)",route,lane,wy,tiso,pl))
    bf["viagens"]=vg
    pls=[]
    for t in b.get("placas",[]):
        car=t[1] if len(t)>1 else ""
        if _only_demarco(car): continue
        t=list(t)
        if len(t)>1: t[1]=_clean_carrier(car) or "(sem transp.)"
        pls.append(tuple(t))
    bf["placas"]=pls
    return bf
def build_sem_cad(recs):
    """Melhoria 1 (aba Auditoria): veiculos SEM CADASTRO identificados pela placa ficticia ASDFG.
    Cada inspecao = uma identificacao. placa_real/transp_real vem do texto da observacao (heuristico);
    obs_raw preserva o texto cru para conferencia."""
    rows=[]
    for r in recs:
        rows.append([
            (r.get("data","") or "").replace("T"," ")[:16],   # 0 data
            r.get("user",""),                                  # 1 auditor
            r.get("tipo",""),                                  # 2 tipo (declarado)
            r.get("placa_real",""),                            # 3 placa real (extraida)
            (r.get("transp_real","") or "(não informada)"),    # 4 transportadora (extraida)
            r.get("nok",[]),                                   # 5 itens NOK (lista de nomes)
            r.get("obs_raw",""),                               # 6 observacao crua
        ])
    rows.sort(key=lambda x:x[0], reverse=True)
    placas=len({x[3] for x in rows if x[3]})
    nok=sum(1 for x in rows if x[5])
    return {"total":len(rows),"placas":placas,"nok":nok,"rows":rows}

def build_anom_aud(insp_seg, week_of):
    """Melhoria 2 (aba Anomalias, origem Auditoria): anomalias das inspecoes de auditoria do MIGUEL.
    Reaproveita build_insp_info montando registros compativeis (sem vrid/carreta)."""
    recs=[]; itemgrp={}
    for r in insp_seg:
        if r.get("user","")!=MIGUEL: continue
        itemgrp.update(r.get("itemgrp",{}))
        recs.append({"data":r["data"],"nok":r.get("nok",[]),"transp":r.get("transp","(sem transp.)"),
                     "tipo":r.get("tipo","(?)"),"vrid":"","ident":r.get("ident",""),
                     "user":r.get("user",""),"carreta":"","transpv":""})
    _iv,_ip,info=build_insp_info(recs, itemgrp, week_of)
    return info

def build_all(b):
    """A partir do banco carregado, constroi detail, insp_info e placas_info."""
    b=apply_filters(b)
    trips, week_of = trips_from_banco(b)
    iv, ip, insp_info = build_insp_info(b["insp"], b["itens"], week_of)
    detail = build_detail(trips, CI, iv, ip, week_of)
    placas_info = build_placas_info(b["placas"], iv, week_of)
    plc_info = build_plc_info(b.get("insp",[]), b.get("insp_seg",[]), b.get("prontuario",[]))
    fuel_info = build_fuel_info(b.get("equip",[]))
    ativas_info = build_ativas_info(b.get("insp",[]), b.get("insp_seg",[]))   # fonte: checklists (pre-uso OU Miguel)
    bases = build_bases(b, insp_info)
    resumo = build_resumo(detail, b.get("insp_seg",[]), insp_info, placas_info, plc_info, fuel_info, ativas_info)
    sem_cad = build_sem_cad(b.get("sem_cad",[]))
    anom_aud = build_anom_aud(b.get("insp_seg",[]), week_of)
    return detail, insp_info, placas_info, plc_info, fuel_info, bases, resumo, sem_cad, anom_aud, ativas_info

def main():
    b=load_banco()
    add=ingest(b)
    save_banco(b)
    if not b["viagens"] or not b["insp"]:
        log("ERRO: banco vazio (sem viagens ou inspecoes)."); sys.exit(2)
    log("Ingerido:", add)
    detail, insp_info, placas_info, plc_info, fuel_info, bases, resumo, sem_cad, anom_aud, ativas_info = build_all(b)
    ho=os.path.join(BASE,"Farol_Amazon.html")
    write_html(detail, insp_info, placas_info, plc_info, fuel_info, bases, resumo, sem_cad, anom_aud, ativas_info, ho)
    c=defaultdict(int)
    for d in detail: c[LBL[d[11]]]+=1
    tot=len(detail); ev=tot-c[SDAT] or 1
    log(f"BANCO: viagens(legs)={len(b['viagens'])} insp={len(b['insp'])} placas={len(b['placas'])}")
    log(f"Viagens (1/VRID)={tot} | Aderente={c[ADER]} ({c[ADER]/ev*100:.1f}%) Fora={c[FORA]} SemInsp={c[SEMI]} SemData={c[SDAT]}")
    log(f"Inspecoes sem VRID={insp_info['sem_vrid']}/{insp_info['total']} | NOK={insp_info['nok_total']}")
    log(f"Sem espelhamento={placas_info['total']} cat={placas_info['cat']}")
    log(f"PLACAS: pre-uso={plc_info['cov']['pre']} auditor={plc_info['cov']['aud']} ambos={plc_info['cov']['both']} | prontuario placas={plc_info['doc']['total_placas']}")
    log(f"COMBUSTIVEL: equip={fuel_info['total']} dist={dict(zip(fuel_info['cats'],fuel_info['dist']))} flagsSim={fuel_info['flagSim']}")
    log(f"BASES extraiveis: preuso={len(bases['preuso']['rows'])} auditor={len(bases['auditor']['rows'])} esp={len(bases['esp']['rows'])}")
    log(f"RESUMO: mes={resumo['mesCorrente']} auditCM={resumo['auditoria']['inspecionados']}/{resumo['auditoria']['placasVrid']} docComDep={resumo['documentacao']['comDeposito']}/{resumo['documentacao']['noSistema']}")
    log(f"SEM CADASTRO (ASDFG): registros={sem_cad['total']} placas={sem_cad['placas']} comNOK={sem_cad['nok']}")
    log(f"ANOMALIAS AUDITORIA (Miguel): checks={anom_aud['total']} comNOK={anom_aud['nok_total']}")
    _pa=resumo['placasAtivas']
    log(f"PLACAS ATIVAS (faturaveis, checklists): mes={_pa['mes']} total={_pa['total']} preuso={_pa['preuso']} miguel={_pa['miguel']} soMiguel={_pa['soMiguel']} | meses={len(ativas_info.get('months',[]))}")
    log("Gerado:", ho)

if __name__=="__main__":
    main()
