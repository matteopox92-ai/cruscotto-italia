"""Prepara i file dati del «Cruscotto Economia Italia».

Uso: python costruisci_cruscotto.py <DEF.xlsx> <cartella Statistiche Italia> <cartella Cruscotto>
Legge lo storico del Documento di Economia e Finanza (fogli «Governo ...»), l'ultima fotografia
delle province, le classifiche e il gettito letti dal gioco, e scrive in <cartella Cruscotto>:
oggi.json, storico.json, prezzi.json, classifiche.json, costi.json, diario.json, donazioni.json, donazioni_totali.json.
"""
import json
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl

PROV = ["Venezia", "Lombardia", "Marsiglia", "Napoli", "Roma", "Firenze", "Palermo"]
NOMI_GIOCO = {"VENEZIA": "Venezia", "LOMBARDIA": "Lombardia", "MARSIGLIA": "Marsiglia", "NAPOLI": "Napoli",
              "ROMA": "Roma", "FIRENZE": "Firenze", "PALERMO": "Palermo"}


def num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("$", "").replace("€", "").replace("(", "-").replace(")", "").strip()
    s = s.replace(".", "").replace(",", ".") if s.count(".") > 1 or "," in s else s
    try:
        return float(s)
    except ValueError:
        return None


def gioco_num(v):
    s = str(v)
    if "." in s and len(s.split(".")[-1]) == 3:
        return float(s.replace(".", ""))
    return float(s)


def storico_def(xlsx):
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    righe, prezzi, governi = [], {}, {}
    for ws in wb:
        if not ws.title[:1].isdigit():
            continue
        giorno = None
        for row in ws.iter_rows(values_only=True):
            row = list(row) + [None] * 40
            b = row[1]
            if any(isinstance(x, str) and "RIEPILOGO" in x for x in row[:4]):
                giorno = None  # il riepilogo di fine mandato non è un giorno
                continue
            if isinstance(b, datetime):
                giorno = b.date().isoformat()
                continue
            if giorno is None or b is None:
                continue
            nome = str(b).strip()
            if nome == "Lombardia" and num(row[31]) and num(row[32]):
                prezzi[giorno] = {"data": giorno, "pelle": num(row[31]), "oro": num(row[32]),
                                  "petrolio": num(row[33]), "nte": num(row[34]), "fonte": "DEF (media settimanale)"}
            governo = ws.title.split(". ", 1)[-1]
            governi.setdefault(governo, set()).add(giorno)
            # colonne: 2-6 manutenzione, 7 confini, 8-12 gettito, 13 tassazioni, 14-18 donazioni, 24-28 tesoro
            # (ognuna in denaro, pelle, oro, petrolio, NTE); i costi nel DEF sono negativi
            gruppo = lambda prefisso, c, segno=1: {f"{prefisso}_{k}": (None if num(row[c + i]) is None else segno * num(row[c + i]))
                                                  for i, k in enumerate(SUFF)}
            if nome == "Stato":
                tes = num(row[24])
                if tes:
                    righe.append({"data": giorno, "provincia": "Stato", **gruppo("donazioni", 14), **gruppo("tesoro", 24),
                                  "governo": governo, "fonte": f"DEF · {ws.title}"})
                continue
            if nome not in PROV and nome != "Bahrein":  # il Bahrein fu italiano nei primi due governi
                continue
            manut = num(row[2])
            if manut is None:
                continue
            righe.append({
                "data": giorno, "provincia": nome,
                **gruppo("manutenzione", 2, -1), "confini_eur": -(num(row[7]) or 0),
                **gruppo("gettito", 8), "tassazioni": num(row[13]),
                **{k: v or 0 for k, v in gruppo("donazioni", 14).items()}, **gruppo("tesoro", 24),
                "governo": governo, "fonte": f"DEF · {ws.title}",
            })
    elenco = sorted(({"governo": g, "inizio": min(d), "fine": max(d)} for g, d in governi.items()), key=lambda r: r["inizio"])
    return righe, prezzi, elenco


SUFF = ["eur", "pelle", "oro", "petrolio", "nte"]


def governo_del(giorno, governi):
    """Il governo in carica in quel giorno: l'ultimo iniziato entro quella data."""
    g = [r["governo"] for r in governi if r["inizio"] <= giorno]
    return g[-1] if g else None


NOMI_SLUG = {"Stato": "Stato", "Venedik": "Venezia", "Lombardiya": "Lombardia", "Marsilya": "Marsiglia", "Napoli": "Napoli",
             "Roma": "Roma", "Floransa": "Firenze", "Palermo": "Palermo"}
RIS5 = ["denaro", "pelle", "oro", "petrolio", "nte"]


def donazioni(out_dir):
    """Da tutti i file donazioni-*.json: righe per giorno e destinatario, totali di sempre, tesoro dello Stato letto."""
    per_giorno, totali, stato = {}, {}, []
    for f in sorted(Path(out_dir).glob("donazioni-*.json")):  # i file più recenti sovrascrivono i giorni già visti
        for l in json.loads(f.read_text(encoding="utf-8"))["letture"]:
            dest = NOMI_SLUG.get(l["d"], l["d"])
            for giorno, v in l["g"].items():
                per_giorno[(giorno, dest)] = {"chiave": f"{giorno}-{dest}", "data": giorno, "destinatario": dest,
                                              **dict(zip(RIS5, v[:5])), "numero": v[5]}
            if l.get("t"):
                totali[dest] = {"destinatario": dest, "letto": l["l"][:10], "elenco_dal": l.get("da"),
                                **{"totale_" + k: x for k, x in zip(RIS5, l["t"])}}
            if l.get("s"):
                stato.append({"data": l["l"][:10], **{"tesoro_" + k: x for k, x in zip(RIS5, l["s"])}})
    righe = sorted(per_giorno.values(), key=lambda r: (r["data"], r["destinatario"]))
    ordine = list(NOMI_SLUG.values())
    return righe, sorted(totali.values(), key=lambda r: ordine.index(r["destinatario"])), stato


def aggiungi_tesoro_stato(storico, stato):
    """Il tesoro dello Stato letto dal gioco vale più di quello del DEF per lo stesso giorno."""
    for s in stato:
        storico[:] = [r for r in storico if not (r["provincia"] == "Stato" and r["data"] == s["data"])]
        storico.append({"data": s["data"], "provincia": "Stato", **{"tesoro_" + ("eur" if k == "denaro" else k): s["tesoro_" + k] for k in RIS5},
                        "fonte": "gioco · pagina del Paese"})
    storico.sort(key=lambda r: (r["data"], r["provincia"]))


def main(xlsx, stat_dir, out_dir):
    stat_dir, out_dir = Path(stat_dir), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    foto = json.loads(sorted(stat_dir.glob("20*.json"))[-1].read_text(encoding="utf-8"))
    clas = json.loads(sorted(stat_dir.glob("classifiche-*.json"))[-1].read_text(encoding="utf-8"))
    get = json.loads(sorted(out_dir.glob("gettito-*.json"))[-1].read_text(encoding="utf-8"))
    oggi_data = foto["data"][:10]
    d_oggi = date.fromisoformat(oggi_data)

    storico, prezzi, governi = storico_def(xlsx)
    gia = {(r["data"], r["provincia"]) for r in storico if r.get("gettito_eur") is not None}
    # giorni completi dal gioco (i 6 prima della lettura) che il DEF non ha, da tutti i file di gettito salvati
    for gf in sorted(out_dir.glob("gettito-*.json")):
      gg = json.loads(gf.read_text(encoding="utf-8"))
      d_lett = date.fromisoformat(gg["data"][:10])
      for nome, g in gg["province"].items():
        for k, v in enumerate(g["giorni"][:-1]):
            giorno = (d_lett - timedelta(days=6 - k)).isoformat()
            if (giorno, nome) in gia:
                continue
            storico = [r for r in storico if not (r["data"] == giorno and r["provincia"] == nome)]
            conf = gg.get("confini_addebiti", {}).get(giorno)
            pr = next(p for p in foto["province"] if NOMI_GIOCO[p["p"]] == nome)
            storico.append({
                "data": giorno, "provincia": nome,
                "manutenzione_eur": gioco_num(pr["man"][0]),
                "confini_eur": conf / 7 if conf else get["confini_per_provincia"],
                "gettito_eur": v, "tassazioni": None, "donazioni_eur": None, "tesoro_eur": None,
                "fonte": "gioco · Gettito fiscale (arrotondato)",
            })
    # tesori delle province da ogni fotografia salvata (una al giorno)
    for f in sorted(stat_dir.glob("20*.json")):
        ft = json.loads(f.read_text(encoding="utf-8"))
        giorno = ft["data"][:10]
        for pr in ft["province"]:
            nome = NOMI_GIOCO[pr["p"]]
            riga = next((r for r in storico if r["data"] == giorno and r["provincia"] == nome), None)
            if riga is None:
                riga = {"data": giorno, "provincia": nome, "manutenzione_eur": gioco_num(pr["man"][0]),
                        "confini_eur": get["confini_per_provincia"], "gettito_eur": None, "tassazioni": None,
                        "donazioni_eur": None, "fonte": "gioco · fotografia"}
                storico.append(riga)
            riga["tesoro_eur"] = gioco_num(pr["tes"][0])
    storico.sort(key=lambda r: (r["data"], r["provincia"]))

    oggi = []
    for pr in foto["province"]:
        nome = NOMI_GIOCO[pr["p"]]
        g = get["province"][nome]
        tes = [gioco_num(x) for x in pr["tes"]]
        man = [gioco_num(x) for x in pr["man"]]
        sei = g["giorni"][:-1]
        oggi.append({
            "provincia": nome, "giocatori": pr["gioc"], "fabbriche": pr["fab"], "governatore": pr["gov"],
            "indice_istruzione": pr["ind"][0], "indice_militare": pr["ind"][1],
            "indice_sviluppo": pr["ind"][2], "indice_salute": pr["ind"][3],
            "scuole": pr["sc"]["n"], "basi": pr["ba"]["n"], "strade": pr["st"]["n"], "ospedali": pr["os"]["n"],
            "tesoro_eur": tes[0], "tesoro_pelle": tes[1], "tesoro_oro": tes[2], "tesoro_petrolio": tes[3], "tesoro_nte": tes[4],
            "manutenzione_eur": man[0], "manutenzione_pelle": man[1], "manutenzione_oro": man[2],
            "manutenzione_petrolio": man[3], "manutenzione_nte": man[4],
            "confini_eur": get["confini_per_provincia"],
            "gettito_medio_eur": round(sum(sei) / len(sei)),
            "gettito_7g_eur": g["tot7"]["denaro"], "gettito_7g_pelle": g["tot7"]["pelle"], "gettito_7g_oro": g["tot7"]["oro"],
            "gettito_7g_petrolio": g["tot7"]["petrolio"], "gettito_7g_nte": g["tot7"]["nte"],
            "tassazioni_oggi": g["tassazioniOggi"],
            "bonus_diamante": int(pr["bon"][0]), "bonus_pelle": int(pr["bon"][1]), "bonus_oro": int(pr["bon"][2]),
            "bonus_petrolio": int(pr["bon"][3]), "bonus_nte": int(pr["bon"][4]),
            "data": oggi_data,
        })

    righe_clas = []
    for edificio in ("scuole", "basi", "strade", "ospedali"):
        for pos, n in enumerate(clas[edificio], start=1):
            righe_clas.append({"edificio": edificio, "posizione": pos, "edifici": n})

    costi = [
        {"edificio": "scuole", "denaro": 100000, "pelle": 400, "oro": 440, "petrolio": 0, "nte": 300},
        {"edificio": "basi", "denaro": 120000, "pelle": 300, "oro": 0, "petrolio": 400, "nte": 400},
        {"edificio": "strade", "denaro": 100000, "pelle": 500, "oro": 0, "petrolio": 400, "nte": 400},
        {"edificio": "ospedali", "denaro": 110000, "pelle": 0, "oro": 440, "petrolio": 300, "nte": 500},
    ]

    # diario del piano: Napoli e Palermo, una riga per fotografia salvata
    diario = []
    for f in sorted(stat_dir.glob("20*.json")):
        ft = json.loads(f.read_text(encoding="utf-8"))
        giorno = ft["data"][:10]
        gf = out_dir / f"gettito-{giorno}.json"
        gt = json.loads(gf.read_text(encoding="utf-8")) if gf.exists() else {"province": {}}
        for pr in ft["province"]:
            nome = NOMI_GIOCO[pr["p"]]
            if nome not in ("Napoli", "Palermo"):
                continue
            g = gt["province"].get(nome, {})
            diario.append({
                "chiave": f"{giorno}-{nome}", "data": giorno, "momento": ft["data"][11:] or "", "provincia": nome,
                "scuole": pr["sc"]["n"], "caserme": pr["ba"]["n"], "strade": pr["st"]["n"], "ospedali": pr["os"]["n"],
                "istruzione": pr["ind"][0], "militare": pr["ind"][1], "sviluppo": pr["ind"][2], "salute": pr["ind"][3],
                "tesoro_eur": gioco_num(pr["tes"][0]), "manutenzione_eur": gioco_num(pr["man"][0]),
                "fabbriche": pr["fab"], "giocatori": pr["gioc"], "lavori_oggi": g.get("tassazioniOggi"),
            })

    don, don_tot, tes_stato = donazioni(out_dir)
    aggiungi_tesoro_stato(storico, tes_stato)
    # giorni senza donazioni nel DEF: si usano quelle lette dagli storici del gioco (solo quelle dei giocatori)
    for d in don:
        riga = next((r for r in storico if r["data"] == d["data"] and r["provincia"] == d["destinatario"]), None)
        if riga is None:
            riga = {"data": d["data"], "provincia": d["destinatario"], "fonte": "gioco · donazioni"}
            storico.append(riga)
        if riga.get("donazioni_eur") is None:
            riga.update({"donazioni_" + ("eur" if k == "denaro" else k): d[k] for k in RIS5})
    for r in storico:
        r.setdefault("governo", governo_del(r["data"], governi))
    storico.sort(key=lambda r: (r["data"], r["provincia"]))

    for nome, dati in (("governi", governi), ("donazioni", don), ("donazioni_totali", don_tot), ("diario", diario), ("oggi", oggi), ("storico", storico), ("prezzi", sorted(prezzi.values(), key=lambda r: r["data"])),
                       ("classifiche", righe_clas), ("costi", costi)):
        (out_dir / f"{nome}.json").write_text(json.dumps(dati, ensure_ascii=False, indent=0), encoding="utf-8")
        print(nome, len(dati))


if __name__ == "__main__":
    main(*sys.argv[1:4])
