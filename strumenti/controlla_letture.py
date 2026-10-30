"""Controlla le letture del gioco prima di costruire il cruscotto (usato dalla GitHub Action del sito).

Uso: python controlla_letture.py <cartella dati>   (con dentro statistiche/ e cruscotto/)
Le letture possono arrivare dagli aiutanti: si accettano solo file con il formato giusto e valori plausibili.
Esce con 1 se trova errori, e il sito resta com'era.
"""
import json
import re
import sys
from pathlib import Path

PROVINCE = {"VENEZIA", "LOMBARDIA", "MARSIGLIA", "NAPOLI", "ROMA", "FIRENZE", "PALERMO"}
NOMI = {"Venezia", "Lombardia", "Marsiglia", "Napoli", "Roma", "Firenze", "Palermo"}
DESTINATARI = {"Stato", "Venedik", "Lombardiya", "Marsilya", "Napoli", "Roma", "Floransa", "Palermo"}
GIORNO = re.compile(r"^20\d\d-\d\d-\d\d")
ERRORI = []


def errore(f, msg):
    ERRORI.append(f"{f.name}: {msg}")


def numero_gioco(v, massimo=1e13):
    """I numeri del gioco arrivano come testo con i punti delle migliaia («4.362.747») o con un decimale («4804.8»)."""
    try:
        s = str(v)
        x = float(s.replace(".", "")) if "." in s and len(s.split(".")[-1]) == 3 else float(s)
    except ValueError:
        return False
    return 0 <= x <= massimo


def numero(v, massimo=1e13):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and 0 <= v <= massimo


def data_del_file(f, d):
    giorno = f.stem[-10:]
    if not isinstance(d.get("data"), str) or not d["data"].startswith(giorno):
        errore(f, f"il campo data ({d.get('data')}) non è del giorno {giorno}")


def statistiche(f, d):
    data_del_file(f, d)
    prov = d.get("province") or []
    if {p.get("p") for p in prov} != PROVINCE or len(prov) != 7:
        return errore(f, "servono le 7 province italiane")
    for p in prov:
        nome = p["p"]
        if not all(numero(p.get(k), 100000) for k in ("gioc", "fab")):
            errore(f, f"{nome}: giocatori o fabbriche non validi")
        if not (isinstance(p.get("ind"), list) and len(p["ind"]) == 4 and all(numero(x, 10) for x in p["ind"])):
            errore(f, f"{nome}: indici non validi")
        for k in ("tes", "man"):
            # dal 30/10/2026 il gioco non ha più il mantenimento delle province: «man» può mancare
            if k == "man" and p.get(k) is None:
                continue
            if not (isinstance(p.get(k), list) and len(p[k]) == 5 and all(numero_gioco(x) for x in p[k])):
                errore(f, f"{nome}: {k} non valido")
        if not (isinstance(p.get("bon"), list) and len(p["bon"]) == 5 and all(str(x).isdigit() for x in p["bon"])):
            errore(f, f"{nome}: bonus non validi")
        for k in ("sc", "ba", "st", "os"):
            e = p.get(k)
            if not (isinstance(e, dict) and numero(e.get("n"), 1e6) and numero(e.get("indice"), 10)):
                errore(f, f"{nome}: edificio {k} non valido")


def classifiche(f, d):
    data_del_file(f, d)
    for k in ("scuole", "basi", "strade", "ospedali"):
        v = d.get(k)
        if not (isinstance(v, list) and 10 <= len(v) <= 2000 and all(numero(x, 1e6) for x in v)):
            errore(f, f"classifica {k} non valida")


def gettito(f, d):
    data_del_file(f, d)
    if not numero(d.get("confini_per_provincia"), 1e10):
        errore(f, "costo dei confini non valido")
    for g, v in (d.get("confini_addebiti") or {}).items():
        if not (GIORNO.match(g) and numero(v, 1e11)):
            errore(f, f"addebito dei confini {g} non valido")
    prov = d.get("province") or {}
    if set(prov) != NOMI:
        return errore(f, "servono le 7 province italiane")
    for nome, g in prov.items():
        if not (isinstance(g.get("giorni"), list) and len(g["giorni"]) == 7 and all(numero(x, 1e11) for x in g["giorni"])):
            errore(f, f"{nome}: gettito dei 7 giorni non valido")
        t = g.get("tot7") or {}
        if not all(numero(t.get(k), 1e12) for k in ("denaro", "pelle", "oro", "petrolio", "nte")):
            errore(f, f"{nome}: totale di 7 giorni non valido")
        if not numero(g.get("tassazioniOggi"), 1e7):
            errore(f, f"{nome}: tassazioni di oggi non valide")


def donazioni(f, d):
    data_del_file(f, d)
    letture = d.get("letture") or []
    if not letture or any(l.get("d") not in DESTINATARI for l in letture):
        return errore(f, "letture delle donazioni non valide")
    for l in letture:
        for giorno, v in (l.get("g") or {}).items():
            if not (GIORNO.match(giorno) and isinstance(v, list) and len(v) == 6 and all(numero(x, 1e12) for x in v)):
                errore(f, f"{l['d']}: donazioni del {giorno} non valide")
        for k in ("s", "t"):
            if l.get(k) is not None and not (len(l[k]) == 5 and all(numero(x, 1e13) for x in l[k])):
                errore(f, f"{l['d']}: campo {k} non valido")


def main(cartella):
    c = Path(cartella)
    controlli = [(c / "statistiche", "20*.json", statistiche), (c / "statistiche", "classifiche-*.json", classifiche),
                 (c / "cruscotto", "gettito-*.json", gettito), (c / "cruscotto", "donazioni-*.json", donazioni)]
    contati = 0
    for cart, schema, fn in controlli:
        for f in sorted(cart.glob(schema)):
            contati += 1
            try:
                fn(f, json.loads(f.read_text(encoding="utf-8")))
            except (ValueError, AttributeError, KeyError, TypeError) as e:
                errore(f, f"file non leggibile ({e})")
    for e in ERRORI:
        print("ERRORE", e)
    print(f"{contati} file controllati, {len(ERRORI)} errori")
    sys.exit(1 if ERRORI or not contati else 0)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "dati")
