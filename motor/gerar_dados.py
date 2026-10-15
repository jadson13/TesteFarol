#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Runner da automacao do Farol (Amazon) — GitHub Actions.

Le a pasta `dados/` (bases importadas pelo botao do dashboard) e o HTML-base
congelado (`motor/base_farol.html` — snapshot 21/09, historico W11..W39), aplica
EXATAMENTE as regras do site e grava `index.html` na raiz do repo.

Arquitetura (igual ao processo manual diario, "congela a base + sobrepoe o novo"):
  - As VIAGENS da base (D.rows) ficam CONGELADAS (zero drift). So viagens com VRID
    novo (dados/viagens.xlsx) sao classificadas e anexadas.
  - A aba "Sem Espelhamento" (esp) nao faz parte das 4 bases diarias -> fica congelada.
  - Pre-uso / auditoria / prontuario / combustivel / placas ativas sao recompostos
    com as funcoes REAIS do motor (gerar_farol.py) sobre (base reconstruida + novo),
    deduplicando sobreposicoes.

Modos (pelo conteudo da pasta dados/):
  - COMPLETA  = viagens + inspecoes + equipamentos + prontuario (4 bases)
  - AUDITORIA = inspecoes + equipamentos + prontuario (3 bases; sem viagens novas)

Uso:
  python3 motor/gerar_dados.py [--repo <dir>] [--base <base_farol.html>] [--out <index.html>]
  (defaults: repo = pasta acima de motor/, base = motor/base_farol.html, out = <repo>/index.html)
"""
import os, re, sys, json, glob, argparse
from datetime import datetime
from collections import defaultdict, deque

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import gerar_farol as E   # motor real (mesma pasta)

def log(*a): print(*a, flush=True)

# ----------------------------------------------------------------------------
# 1. Carrega a base congelada (payload D do HTML)
# ----------------------------------------------------------------------------
def load_base(base_html_path):
    html = open(base_html_path, encoding="utf-8").read()
    m = re.search(r'<script id="payload" type="application/json">(.*?)</script>', html, re.S)
    if not m:
        raise SystemExit("ERRO: payload nao encontrado no HTML-base.")
    D = json.loads(m.group(1))
    return html, D

# ----------------------------------------------------------------------------
# 2. Reconstroi registros crus da BASE (ja filtrados) a partir de D
# ----------------------------------------------------------------------------
def _d2t(data):  # 'YYYY-MM-DD HH:MM' -> 'YYYY-MM-DDTHH:MM' (iso aceito por fromisoformat)
    return data.replace(" ", "T") if data else ""

def reconstruct_base(D):
    items = D["insp"]["items"]
    # join nokRows -> nomes de item reprovado por inspecao de pre-uso
    nokmap = defaultdict(deque)
    for r in D["insp"]["nokRows"]:
        # r = [ident,tI,cI,user,data,carreta,vrid,mI,nk,wI]
        nokmap[(r[0], r[4], r[5], r[6])].append(r[8])
    pcar = D["bases"]["preuso"]["carriers"]; ptip = D["bases"]["preuso"]["tipos"]
    base_insp = []
    for r in D["bases"]["preuso"]["rows"]:
        ident, tI, cI, user, data, carreta, vrid, transpv, nc = r
        nknames = []
        if nc > 0:
            dq = nokmap.get((ident, data, carreta, vrid))
            if dq:
                nknames = [items[i] for i in dq.popleft()]
            else:
                nknames = ["(item)"] * nc
        base_insp.append({"ident": ident,
                          "tipo": ptip[tI] if tI < len(ptip) else "(?)",
                          "transp": pcar[cI] if cI < len(pcar) else "(sem transp.)",
                          "user": user, "data": _d2t(data), "carreta": carreta,
                          "vrid": vrid, "transpv": transpv, "nok": nknames})
    # auditoria (nao-ASDFG, >= AUD_MIN_DATE, ja filtrada)
    aud_grp = dict(zip(D["anomAud"]["items"], D["anomAud"]["item_groups"]))
    base_seg = []
    for r in D["bases"]["auditor"]["rows"]:
        ident, tipo, transp, user, data, ncount, noknames = r
        nks = [x for x in (noknames.split(" | ") if noknames else []) if x]
        base_seg.append({"ident": ident, "tipo": tipo, "transp": transp, "user": user,
                         "data": _d2t(data), "nok": nks, "itemgrp": aud_grp, "sheet": ""})
    # ASDFG (veiculo sem cadastro) -> registros para build_sem_cad
    base_semcad = []
    for r in D["semCad"]["rows"]:
        data, user, tipo, placa_real, transp_real, noklist, obs_raw = r
        base_semcad.append({"ident": E.SEM_CADASTRO_PLATE, "tipo": tipo, "transp": "(sem transp.)",
                            "user": user, "data": _d2t(data), "nok": noklist, "itemgrp": aud_grp,
                            "sheet": "", "sc": True, "placa_real": placa_real,
                            "transp_real": "" if transp_real == "(não informada)" else transp_real,
                            "obs_raw": obs_raw})
    # viagens congeladas (D.rows -> formato detail de build_detail)
    MET_INV = {0: "VRID", 1: "Placa", 2: "-"}; AN_INV = {0: "", 1: "OK", 2: "NOK"}
    cs, ws, ms = D["carriers"], D["weeks"], D["months"]
    base_detail = []
    for r in D["rows"]:
        vrid, cI, wI, mI, tstr, plate, sc, metc, bests, hb, anc = r
        T = datetime.strptime(tstr, "%Y-%m-%d %H:%M") if tstr else None
        best = datetime.strptime("2000-" + bests, "%Y-%m-%d %H:%M") if bests else None
        base_detail.append([vrid, cs[cI], "", ws[wI], T, plate, E.LBL[sc], MET_INV.get(metc, "-"),
                            best, (hb if hb != "" else None), AN_INV.get(anc, ""), sc, ms[mI]])
    base_itens = dict(zip(D["insp"]["items"], D["insp"]["item_groups"]))
    return base_insp, base_seg, base_semcad, base_detail, base_itens

# ----------------------------------------------------------------------------
# 3. Le bases novas da pasta dados/
# ----------------------------------------------------------------------------
def _sniff_kind(path):
    k = E.sniff(path)
    if k: return k
    try:
        if E.is_prontuario(path): return "prontuario"
        if E.is_equipamentos(path): return "equipamentos"
    except Exception: pass
    return None

def read_dados(dados_dir):
    files = [f for f in glob.glob(os.path.join(dados_dir, "*.xlsx"))
             if not os.path.basename(f).startswith("~$")]
    kinds = {f: _sniff_kind(f) for f in files}
    viagens = [f for f, k in kinds.items() if k == "viagens"]
    inspec  = sorted(f for f, k in kinds.items() if k == "inspecoes")
    equips  = [f for f, k in kinds.items() if k == "equipamentos"]
    pronts  = [f for f, k in kinds.items() if k == "prontuario"]
    newest = lambda xs: sorted(xs, key=os.path.getmtime, reverse=True)[0] if xs else None
    return {"viagens": newest(viagens), "inspecoes": inspec,
            "equip": newest(equips), "pront": newest(pronts),
            "n_viagens": len(viagens), "n_inspecoes": len(inspec),
            "n_equip": len(equips), "n_pront": len(pronts)}

def parse_novo(dd):
    # pre-uso + itens
    new_insp = []; new_itens = {}
    for f in dd["inspecoes"]:
        recs, ig = E.parse_insp_file(f)
        new_insp += recs
        for nm, gp in ig.items(): new_itens.setdefault(nm, gp)
    # auditor
    new_seg = []
    for f in dd["inspecoes"]:
        new_seg += E.parse_auditor_file(f)
    # prontuario / equipamentos (snapshots mais recentes)
    pront = E.parse_prontuario(dd["pront"])[0] if dd["pront"] else []
    equip = E.parse_equipamentos(dd["equip"]) if dd["equip"] else []
    # viagens -> linhas no formato banco (vrid,car,route,lane,wy,tiso,plate)
    new_vrows = []
    if dd["viagens"]:
        recs, _ = E.normalize_trips(dd["viagens"])
        for rec in recs:
            tiso = rec[5].isoformat() if isinstance(rec[5], datetime) else ""
            new_vrows.append((rec[0], rec[1], rec[2], rec[3], rec[4], tiso, rec[6]))
    return new_insp, new_itens, new_seg, pront, equip, new_vrows

# ----------------------------------------------------------------------------
# 4. Monta o payload (replica write_html, com esp congelado)
# ----------------------------------------------------------------------------
def build_payload(D, detail, insp_info, plc_info, fuel_info, bases, resumo, sem_cad, anom_aud, ativas):
    MESES = E.MESES; LBL = E.LBL; ADER, FORA, SEMI, SDAT = E.ADER, E.FORA, E.SEMI, E.SDAT
    carriers = []; ci = {}; weeks = []; wi = {}; months = []; mi = {}
    def cidx(c):
        if c not in ci: ci[c] = len(carriers); carriers.append(c)
        return ci[c]
    cat = defaultdict(int); rows = []
    MET = {"VRID": 0, "Placa": 1, "-": 2}; AN = {"OK": 1, "NOK": 2, "": 0}
    for d in detail:
        sc = d[11]; cat[LBL[sc]] += 1
        wk = d[3]
        if wk not in wi: wi[wk] = len(weeks); weeks.append(wk)
        mo = d[12] or "-"
        if mo not in mi: mi[mo] = len(months); months.append(mo)
        rows.append([d[0], cidx(d[1]), wi[wk], mi[mo],
                     d[4].strftime("%Y-%m-%d %H:%M") if isinstance(d[4], datetime) else "",
                     d[5], sc, MET.get(d[7], 2),
                     d[8].strftime("%m-%d %H:%M") if isinstance(d[8], datetime) else "",
                     (d[9] if d[9] is not None else ""), AN.get(d[10], 0)])
    def mkey(m):
        try:
            ab, yr = m.split("/"); inv = {v: k for k, v in MESES.items()}; return (int(yr), inv.get(ab, 0))
        except: return (9999, 99)
    wsorted = sorted(weeks); msorted = sorted(months, key=mkey)
    wmap = {w: i for i, w in enumerate(wsorted)}; mmap = {m: i for i, m in enumerate(msorted)}
    for r in rows:
        r[2] = wmap[weeks[r[2]]]; r[3] = mmap[months[r[3]]]
    def _proc(ii):
        ms = sorted(ii["months"], key=mkey); mm = [ms.index(m) for m in ii["months"]]
        ws = sorted(ii["weeks"]); wm = [ws.index(w) for w in ii["weeks"]]
        agg = [[a[0], mm[a[1]], a[2], a[3], a[4], wm[a[5]]] for a in ii["agg"]]
        semR = [[r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], mm[r[8]], wm[r[9]]] for r in ii["sem_rows"]]
        nokR = [[r[0], r[1], r[2], r[3], r[4], r[5], r[6], mm[r[7]], r[8], wm[r[9]]] for r in ii["nok_rows"]]
        return {"total": ii["total"], "com_vrid": ii["com_vrid"], "sem_vrid": ii["sem_vrid"], "nok_total": ii["nok_total"],
                "carriers": ii["carriers"], "months": ms, "weeks": ws, "tipos": ii["tipos"],
                "items": ii["items"], "item_groups": ii["item_groups"],
                "agg": agg, "semRows": semR, "nokRows": nokR}
    data = {"generated": datetime.now().strftime("%d/%m/%Y %H:%M"), "total": len(rows),
            "cat": {"aderente": cat[ADER], "fora_prazo": cat[FORA], "sem_inspecao": cat[SEMI], "sem_data": cat[SDAT]},
            "carriers": carriers, "weeks": wsorted, "months": msorted, "rows": rows,
            "insp": _proc(insp_info)}
    data["anomAud"] = _proc(anom_aud)
    data["semCad"] = sem_cad
    data["esp"] = D["esp"]                      # aba Sem Espelhamento: congelada da base
    data["plc"] = plc_info; data["fuel"] = fuel_info; data["bases"] = bases
    data["resumo"] = resumo; data["ativas"] = ativas
    return data

# ----------------------------------------------------------------------------
# 5. Orquestracao
# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", default=os.path.dirname(HERE))
    ap.add_argument("--base", default=os.path.join(HERE, "base_farol.html"))
    ap.add_argument("--out", default=None)
    ap.add_argument("--dados", default=None)
    a = ap.parse_args()
    repo = a.repo
    dados_dir = a.dados or os.path.join(repo, "dados")
    out = a.out or os.path.join(repo, "index.html")

    base_html, D = load_base(a.base)
    base_insp, base_seg, base_semcad, base_detail, base_itens = reconstruct_base(D)
    log(f"BASE: viagens(detail)={len(base_detail)} preuso={len(base_insp)} auditor={len(base_seg)} semCad={len(base_semcad)}")

    dd = read_dados(dados_dir)
    modo = "COMPLETA" if dd["viagens"] else "AUDITORIA"
    log(f"DADOS: modo={modo} | viagens={dd['n_viagens']} inspecoes={dd['n_inspecoes']} equip={dd['n_equip']} pront={dd['n_pront']}")
    if not dd["inspecoes"]:
        raise SystemExit("ERRO: nenhum arquivo de inspecoes em dados/ (minimo: inspecoes + equip + pront).")
    new_insp, new_itens, new_seg, new_pront, new_equip, new_vrows = parse_novo(dd)
    log(f"NOVO(cru): preuso={len(new_insp)} auditor={len(new_seg)} pront={len(new_pront)} equip={len(new_equip)} viagens(legs)={len(new_vrows)}")

    # filtra o NOVO com a mesma logica do motor (DEMARCO, inativos, split ASDFG, limpa transp)
    b_new = {"insp": new_insp, "insp_seg": new_seg, "prontuario": new_pront, "equip": new_equip,
             "viagens": new_vrows, "placas": [], "itens": new_itens, "meta": {}}
    bf = E.apply_filters(b_new)
    log(f"NOVO(filtrado): preuso={len(bf['insp'])} auditor={len(bf['insp_seg'])} asdfg={len(bf['sem_cad'])} "
        f"pront={len(bf['prontuario'])} equip={len(bf['equip'])} viagens(legs)={len(bf['viagens'])}")

    # ---- combina base + novo (dedup) ----
    def dedup(seq, keyf):
        out = []; seen = set()
        for r in seq:
            k = keyf(r)
            if k in seen: continue
            seen.add(k); out.append(r)
        return out
    ikey = lambda r: (str(r.get("ident", "")).strip().upper(),
                      (r.get("data", "") or "").replace("T", " ")[:16], E.nvr(r.get("vrid", "")))
    skey = lambda r: (str(r.get("ident", "")).strip().upper(),
                      (r.get("data", "") or "").replace("T", " ")[:16])
    sckey = lambda r: (r.get("user", ""), (r.get("data", "") or "").replace("T", " ")[:16])
    insp = dedup(base_insp + bf["insp"], ikey)
    insp_seg = dedup(base_seg + bf["insp_seg"], skey)
    sem_cad_recs = dedup(base_semcad + bf["sem_cad"], sckey)
    itens = dict(base_itens); itens.update(new_itens)
    log(f"COMBINADO: preuso={len(insp)} auditor={len(insp_seg)} semCad={len(sem_cad_recs)}")

    # ---- week_of (base + novo) ----
    pairs = []
    for d in base_detail:
        if isinstance(d[4], datetime) and re.match(r'\d{4}W\d', str(d[3])):
            pairs.append((d[3], d[4].date()))
    for (v, car, route, lane, wy, tiso, pl) in bf["viagens"]:
        if wy and tiso: pairs.append((wy, datetime.fromisoformat(tiso).date()))
    week_of = E.build_week_of(pairs)

    # ---- pre-uso / classificacao ----
    iv, ip, insp_info = E.build_insp_info(insp, itens, week_of)

    # viagens novas: dedup por VRID contra a base, 1 perna por VRID (first leg)
    base_vrids = {E.nvr(d[0]) for d in base_detail}
    groups = defaultdict(list)
    for (v, car, route, lane, wy, tiso, pl) in bf["viagens"]:
        t = datetime.fromisoformat(tiso) if tiso else None
        groups[v].append((v, car, route, lane, wy, t, pl))
    new_trips = [E._first_leg(rs) for vr, rs in groups.items() if E.nvr(vr) not in base_vrids]
    new_detail = E.build_detail(new_trips, E.CI, iv, ip, week_of)
    detail = base_detail + new_detail
    log(f"VIAGENS: base(congeladas)={len(base_detail)} novas={len(new_detail)} total={len(detail)}")

    # ---- demais tabelas (recompostas) ----
    plc_info = E.build_plc_info(insp, insp_seg, bf["prontuario"])
    fuel_info = E.build_fuel_info(bf["equip"])
    ativas_info = E.build_ativas_info(insp, insp_seg)
    bases = E.build_bases({"insp": insp, "insp_seg": insp_seg, "placas": []}, insp_info)
    bases["esp"] = D["bases"]["esp"]            # extrato esp: congelado
    placas_stub = {"total": D["esp"]["total"], "cat": D["esp"]["cat"]}
    resumo = E.build_resumo(detail, insp_seg, insp_info, placas_stub, plc_info, fuel_info, ativas_info)
    sem_cad = E.build_sem_cad(sem_cad_recs)
    anom_aud = E.build_anom_aud(insp_seg, week_of)

    # ---- payload -> index.html ----
    data = build_payload(D, detail, insp_info, plc_info, fuel_info, bases, resumo, sem_cad, anom_aud, ativas_info)
    pj = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    new_html = re.sub(r'(<script id="payload" type="application/json">).*?(</script>)',
                      lambda m: m.group(1) + pj + m.group(2), base_html, count=1, flags=re.S)
    open(out, "w", encoding="utf-8").write(new_html)

    c = defaultdict(int)
    for d in detail: c[d[11]] += 1
    tot = len(detail); ev = (tot - c[3]) or 1
    log("-" * 60)
    log(f"OK modo={modo}  ->  {out}  ({len(new_html)/1e6:.1f} MB)")
    log(f"Viagens={tot} | Aderente={c[0]} ({c[0]/ev*100:.1f}%) Fora={c[1]} SemInsp={c[2]} SemData={c[3]}")
    log(f"Inspecoes pre-uso={insp_info['total']} semVRID={insp_info['sem_vrid']} NOK={insp_info['nok_total']}")
    log(f"Auditoria: checks={resumo['auditoria']['totalAuditorias']} placas={resumo['auditoria']['totalAudit']} "
        f"mesCorrente={resumo['mesCorrente']} inspMes={resumo['auditoria']['inspecionados']}/{resumo['auditoria']['placasVrid']}")
    log(f"Prontuario placas={plc_info['doc']['total_placas']} | Combustivel equip={fuel_info['total']} | "
        f"PlacasAtivas mes={resumo['placasAtivas']['mes']} total={resumo['placasAtivas']['total']}")
    log(f"SemCadastro(ASDFG)={sem_cad['total']} | AnomaliasAuditoria(Miguel) checks={anom_aud['total']} NOK={anom_aud['nok_total']}")

if __name__ == "__main__":
    main()
