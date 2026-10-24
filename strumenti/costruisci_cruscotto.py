"""Prepara i file dati del «Cruscotto Economia Italia».

Uso: python costruisci_cruscotto.py <DEF.xlsx> <cartella Statistiche Italia> <cartella Cruscotto>
Legge lo storico del Documento di Economia e Finanza (fogli «Governo ...»), l'ultima fotografia
delle province, le classifiche e il gettito letti dal gioco, e scrive in <cartella Cruscotto>:
oggi.json, storico.json, prezzi.json, classifiche.json, costi.json, diario.json, donazioni.json, donazioni_totali.json.

Giornata italiana (00:00–24:00, regola dell'utente del 22/10/2026): dal 21/10 i giorni dello storico vengono da
DEF/giornate.json, scritto da aggiorna_def.py, così cruscotto, DEF e recap danno le stesse cifre. Le letture orarie
hanno anche i totali dalla mezzanotte italiana (gettito_it_*, tassazioni_it), usati dalla scheda Storico.
"""
import json
import re
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
from giornata_it import di_mezzanotte, gettito_tra, giorno_gioco, istante  # noqa: E402

PROV = ["Venezia", "Lombardia", "Marsiglia", "Napoli", "Roma", "Firenze", "Palermo"]
LIM5 = ("diamante", "pelle", "oro", "petrolio", "nte")  # ordine del riquadro «Limiti delle risorse provinciali»
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
                **gruppo("donazioni", 14), **gruppo("tesoro", 24),
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


def righe_orarie(foto, gettito, don_lett, quando):
    """Una «fotografia» per lettura: tesoro, gettito e tassazioni di oggi, donazioni di oggi, per provincia e Stato.
    quando = ora della lettura (datetime con fuso). Il gettito di oggi è della giornata del gioco (00:00 GMT+3),
    le donazioni della data italiana (come nel DEF)."""
    ts = quando.isoformat(timespec="minutes")
    data_it = quando.date().isoformat()
    giorno_gioco = quando.astimezone(timezone(timedelta(hours=3))).date().isoformat()
    don = {NOMI_SLUG.get(l["d"], l["d"]): l for l in (don_lett or [])}
    righe = []
    for pr in foto["province"]:
        nome = NOMI_GIOCO[pr["p"]]
        g = (gettito or {}).get("province", {}).get(nome, {})
        oggi = g.get("oggi") or {}
        riga = {"ts": ts, "provincia": nome, "data_it": data_it, "giorno_gioco": giorno_gioco,
                **{"tesoro_" + ("eur" if k == "denaro" else k): gioco_num(x) for k, x in zip(RIS5, pr["tes"])},
                "tassazioni_oggi": g.get("tassazioniOggi")}
        if pr.get("lim"):  # quanto resta dei limiti giornalieri della provincia (giornata del gioco, dal 24/10/2026)
            riga.update({"limite_" + k: v for k, v in zip(LIM5, pr["lim"])})
        if oggi:
            riga.update({"gettito_oggi_" + ("eur" if k == "denaro" else k): oggi.get(k) for k in RIS5})
        elif g.get("denaroOggi") is not None or g.get("giorni"):
            riga["gettito_oggi_eur"] = g.get("denaroOggi") if g.get("denaroOggi") is not None else g["giorni"][-1]
        d = don.get(nome)
        if d is not None:
            v = d.get("g", {}).get(data_it, [0] * 6)
            riga.update({"donazioni_oggi_" + ("eur" if k == "denaro" else k): x for k, x in zip(RIS5, v[:5])})
        righe.append(riga)
    st = don.get("Stato")
    if st is not None and st.get("s"):
        v = st.get("g", {}).get(data_it, [0] * 6)
        righe.append({"ts": ts, "provincia": "Stato", "data_it": data_it, "giorno_gioco": giorno_gioco,
                      **{"tesoro_" + ("eur" if k == "denaro" else k): x for k, x in zip(RIS5, st["s"])},
                      **{"donazioni_oggi_" + ("eur" if k == "denaro" else k): x for k, x in zip(RIS5, v[:5])}})
    return righe


def ora_lettura(foto, file_foto):
    """Ora della fotografia: «AAAA-MM-GG ora HH:MM» se c'è, altrimenti l'ora del file (letture «sera»)."""
    testo = foto.get("data", "")
    if " ora " in testo:
        return datetime.strptime(testo[:10] + " " + testo.split(" ora ")[1][:5], "%Y-%m-%d %H:%M").astimezone()
    return datetime.fromtimestamp(Path(file_foto).stat().st_mtime).astimezone().replace(second=0, microsecond=0)


def aggiungi_orario(vecchie, nuove):
    """Aggiunge le righe di una lettura se non ci sono già (stessa ora, o stessi numeri nello stesso giorno)."""
    firma = lambda r: (r["provincia"], r.get("data_it"), r.get("tesoro_eur"), r.get("gettito_oggi_eur"),
                       r.get("tassazioni_oggi"), r.get("donazioni_oggi_eur"), r.get("limite_diamante"), r.get("limite_oro"))
    gia = {(r["provincia"], r["ts"]) for r in vecchie} | {firma(r) for r in vecchie}
    for r in nuove:
        if (r["provincia"], r["ts"]) not in gia and firma(r) not in gia:
            vecchie.append(r)
    vecchie.sort(key=lambda r: (r["ts"], r["provincia"]))
    return vecchie


def aggiungi_tesoro_stato(storico, stato):
    """Il tesoro dello Stato letto dal gioco vale più di quello del DEF per lo stesso giorno."""
    for s in stato:
        storico[:] = [r for r in storico if not (r["provincia"] == "Stato" and r["data"] == s["data"])]
        storico.append({"data": s["data"], "provincia": "Stato", **{"tesoro_" + ("eur" if k == "denaro" else k): s["tesoro_" + k] for k in RIS5},
                        "fonte": "gioco · pagina del Paese"})
    storico.sort(key=lambda r: (r["data"], r["provincia"]))


def carica_giornate(out_dir):
    """Giornate italiane calcolate da aggiorna_def.py (DEF/giornate.json accanto alla cartella del cruscotto)."""
    f = Path(out_dir).parent / "DEF" / "giornate.json"
    return json.loads(f.read_text(encoding="utf-8")) if f.exists() else []


def applica_giornate(storico, giornate, governi):
    """Mette nello storico i numeri delle giornate italiane (dove il DEF del governo non ha già un valore)."""
    for g in giornate:
        giorno = g["giorno"]
        conf = (g["voci"].get("TOTALE") or {}).get("confini")
        for chi, voci in g["voci"].items():
            if chi not in PROV and chi != "Stato":
                continue
            riga = next((r for r in storico if r["data"] == giorno and r["provincia"] == chi), None)
            if riga is None:
                riga = {"data": giorno, "provincia": chi, "governo": governo_del(giorno, governi)}
                storico.append(riga)
            nuovi = {}
            for voce in ("manutenzione", "gettito", "donazioni", "tesoro"):
                if voci.get(voce) is not None:
                    nuovi.update({f"{voce}_{k}": v for k, v in zip(SUFF, voci[voce])})
            if voci.get("tassazioni") is not None:
                nuovi["tassazioni"] = voci["tassazioni"]
            if chi != "Stato" and conf is not None:
                nuovi["confini_eur"] = conf / 7  # nel DEF ogni provincia ha 1/7 dei confini
            for k, v in nuovi.items():
                # una formula del DEF senza dati (confini =H/7 con H vuoto) vale 0: in quel caso conta la nostra lettura
                if riga.get(k) is None or (k == "confini_eur" and not riga.get(k)):
                    riga[k] = v
            riga["provvisorio"] = not g["chiusa"]
            riga["fonte"] = "DEF aggiornato · giornata italiana 00–24" + ("" if g["chiusa"] else " (in corso)")
    storico.sort(key=lambda r: (r["data"], r["provincia"]))


def aggiungi_dalla_mezzanotte(orario, gioco):
    """Per ogni lettura oraria: gettito e tassazioni dalla mezzanotte italiana (la stessa regola del DEF).
    Serve la lettura di mezzanotte di quel giorno; senza, il valore resta vuoto."""
    per_p = {}
    for r in orario:
        if r["provincia"] != "Stato":
            per_p.setdefault(r["provincia"], []).append(r)
    for p, righe in per_p.items():
        righe.sort(key=lambda r: istante(r["ts"]))
        chiusi = {date.fromisoformat(g): v for g, v in (gioco.get(p) or {}).items()}
        ult_t = {}
        for r in righe:  # tassazioni: l'ultimo valore letto in ogni giornata del gioco vale come totale
            if r.get("tassazioni_oggi") is not None:
                ult_t[giorno_gioco(istante(r["ts"]))] = r["tassazioni_oggi"]
        for r in righe:
            t = istante(r["ts"])
            d = date.fromisoformat(r["data_it"])
            m = next((x for x in righe if di_mezzanotte(istante(x["ts"]), d) and x.get("gettito_oggi_eur") is not None), None)
            for k in SUFF:
                r.pop("gettito_it_" + k, None)
            r.pop("tassazioni_it", None)
            if m is None or istante(m["ts"]) > t:
                continue
            tm = istante(m["ts"])
            for i, k in enumerate(SUFF):
                a, b = m.get("gettito_oggi_" + k), r.get("gettito_oggi_" + k)
                if a is None or b is None:
                    continue
                v, manca = gettito_tra((tm, [a]), (t, [b]), {g: [x[i]] for g, x in chiusi.items()})
                if not manca and v[0] >= 0:
                    r["gettito_it_" + k] = v[0]
            a, b = m.get("tassazioni_oggi"), r.get("tassazioni_oggi")
            if a is not None and b is not None:
                v, manca = gettito_tra((tm, [a]), (t, [b]), {g: [n] for g, n in ult_t.items() if g < giorno_gioco(t)})
                if not manca and v[0] >= 0:
                    r["tassazioni_it"] = v[0]
    return orario


def main(xlsx, stat_dir, out_dir):
    stat_dir, out_dir = Path(stat_dir), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    foto = json.loads(sorted(stat_dir.glob("20*.json"))[-1].read_text(encoding="utf-8"))
    clas = json.loads(sorted(stat_dir.glob("classifiche-*.json"))[-1].read_text(encoding="utf-8"))
    get = json.loads(sorted(out_dir.glob("gettito-*.json"))[-1].read_text(encoding="utf-8"))
    oggi_data = foto["data"][:10]
    d_oggi = date.fromisoformat(oggi_data)

    storico, prezzi, governi = storico_def(xlsx)
    giornate = carica_giornate(out_dir)
    coperti = {g["giorno"] for g in giornate}  # giorni che vengono dal DEF aggiornato (giornata italiana)
    gia = {(r["data"], r["provincia"]) for r in storico if r.get("gettito_eur") is not None}
    # giorni completi dal gioco (i 6 prima della lettura) che il DEF non ha, da tutti i file di gettito salvati
    for gf in sorted(out_dir.glob("gettito-*.json")):
      gg = json.loads(gf.read_text(encoding="utf-8"))
      # le barre finiscono con la giornata del gioco della lettura (dopo le 23 italiane è già il giorno dopo)
      d_lett = (giorno_gioco(istante(gg["data"][:10] + "T" + gg["data"].split(" ora ")[1][:5])) if " ora " in gg["data"]
                else date.fromisoformat(gg["data"][:10]))
      for nome, g in gg["province"].items():
        for k, v in enumerate(g["giorni"][:-1]):
            giorno = (d_lett - timedelta(days=6 - k)).isoformat()
            if (giorno, nome) in gia or giorno in coperti:
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
        if giorno in coperti:
            continue
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
            **({"limite_" + k: v for k, v in zip(LIM5, pr["lim"])} if pr.get("lim") else {}),
        })
        # edifici che mancano all'indice successivo secondo il gioco («+9 edifici → Indice 6»): più preciso del modello a fasce.
        # Vale solo se porta all'indice subito sopra: per gli ospedali il gioco indica il prossimo bonus («+82 edifici → Indice 10»
        # da salute 7), che salta indici; allora si usa il modello a fasce (errore scoperto il 25/10/2026 su Palermo: +82 invece di +2)
        for ed, k in (("scuole", "sc"), ("basi", "ba"), ("strade", "st"), ("ospedali", "os")):
            dato = pr.get(k) or {}
            m = re.match(r"\s*\+\s*([\d.]+)\s*edific[^→]*→\s*Indice\s*(\d+)", dato.get("prossimo") or "")
            ok = m and dato.get("indice") is not None and int(m.group(2)) == int(dato["indice"]) + 1
            oggi[-1]["per_salire_" + ed] = int(m.group(1).replace(".", "")) if ok else None

    # classifica precedente (giorno diverso): serve a misurare quanto salgono le soglie («inflazione» degli indici)
    file_clas = sorted(stat_dir.glob("classifiche-*.json"))
    d_clas = date.fromisoformat(file_clas[-1].stem.split("-", 1)[1])
    prec, giorni_prec = None, None
    for f in reversed(file_clas[:-1]):
        d_f = date.fromisoformat(f.stem.split("-", 1)[1])
        if d_f < d_clas:
            prec, giorni_prec = json.loads(f.read_text(encoding="utf-8")), (d_clas - d_f).days
            break
    # lettura incompleta (lo scorrimento non ha caricato tutta la classifica, come le strade del 24/10/2026:
    # 193 province su 230): le soglie verrebbero troppo alte, quindi si tiene la classifica precedente
    letti, sostituite = {e: len(clas[e]) for e in ("scuole", "basi", "strade", "ospedali")}, set()
    if prec:
        for edificio in ("scuole", "basi", "strade", "ospedali"):
            n, n_prec = len(clas[edificio]), len(prec[edificio])
            if n < 0.9 * n_prec:
                print(f"AVVISO: classifica {edificio} incompleta ({n} province contro {n_prec} della precedente): "
                      f"uso quella di {giorni_prec} giorni fa")
                clas[edificio] = prec[edificio]
                sostituite.add(edificio)
    righe_clas = []
    for edificio in ("scuole", "basi", "strade", "ospedali"):
        for pos, n in enumerate(clas[edificio], start=1):
            riga = {"edificio": edificio, "posizione": pos, "edifici": n}
            if pos == 1:  # per la riga «Salute dei dati» del cruscotto: province lette e classifica sostituita
                riga.update({"letti": letti[edificio], "sostituita": edificio in sostituite})
            if prec and pos <= len(prec[edificio]):
                riga.update({"edifici_prec": prec[edificio][pos - 1], "giorni_prec": giorni_prec})
            righe_clas.append(riga)

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
    aggiungi_tesoro_stato(storico, [s for s in tes_stato if s["data"] not in coperti])
    # giorni senza donazioni nel DEF: si usano quelle lette dagli storici del gioco (solo quelle dei giocatori)
    for d in don:
        if d["data"] in coperti:
            continue
        riga = next((r for r in storico if r["data"] == d["data"] and r["provincia"] == d["destinatario"]), None)
        if riga is None:
            riga = {"data": d["data"], "provincia": d["destinatario"], "fonte": "gioco · donazioni"}
            storico.append(riga)
        if riga.get("donazioni_eur") is None:
            riga.update({"donazioni_" + ("eur" if k == "denaro" else k): d[k] for k in RIS5})
    applica_giornate(storico, giornate, governi)
    for r in storico:
        r.setdefault("governo", governo_del(r["data"], governi))
    storico.sort(key=lambda r: (r["data"], r["provincia"]))
    # gettito medio per l'oggi: ultime 6 giornate chiuse dello storico (giornata italiana dove c'è)
    for o in oggi:
        chiuse = [r["gettito_eur"] for r in storico if r["provincia"] == o["provincia"] and r["data"] < oggi_data
                  and r.get("gettito_eur") is not None and not r.get("provvisorio")][-6:]
        if len(chiuse) == 6:
            o["gettito_medio_eur"] = round(sum(chiuse) / 6)

    # fotografie orarie (scheda Storico, dettaglio «Ore»): si aggiunge la lettura di adesso a quelle già salvate
    f_orario = out_dir / "orario.json"
    orario = json.loads(f_orario.read_text(encoding="utf-8")) if f_orario.exists() else []
    file_foto = sorted(stat_dir.glob("20*.json"))[-1]
    f_don = out_dir / f"donazioni-{oggi_data}.json"
    don_lett = json.loads(f_don.read_text(encoding="utf-8"))["letture"] if f_don.exists() else []
    orario = aggiungi_orario(orario, righe_orarie(foto, get if get["data"][:10] == oggi_data else None, don_lett,
                                                   ora_lettura(foto, file_foto)))
    f_gioco = out_dir.parent / "DEF" / "gettito_gioco.json"
    orario = aggiungi_dalla_mezzanotte(orario, json.loads(f_gioco.read_text(encoding="utf-8")) if f_gioco.exists() else {})

    for nome, dati in (("governi", governi), ("donazioni", don), ("donazioni_totali", don_tot), ("diario", diario), ("oggi", oggi), ("storico", storico), ("prezzi", sorted(prezzi.values(), key=lambda r: r["data"])),
                       ("classifiche", righe_clas), ("costi", costi), ("orario", orario)):
        (out_dir / f"{nome}.json").write_text(json.dumps(dati, ensure_ascii=False, indent=0), encoding="utf-8")
        print(nome, len(dati))


if __name__ == "__main__":
    main(*sys.argv[1:4])
