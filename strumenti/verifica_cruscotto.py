"""Verifica indipendente dei dati del «Cruscotto Economia Italia».

Uso: python -I verifica_cruscotto.py <cartella Cruscotto>
Rifà i conti senza usare il codice della pagina e stampa OK / ERRORE per ogni controllo:
manutenzione = 3% dei costi di costruzione, indici dalla classifica mondiale, conti del piano
(spese, risparmi, saldi, pareggi), tesori dello storico, più alcune prove su casi limite.
Esce con codice 1 se un controllo fallisce.
"""
import json
import math
import sys
from collections import defaultdict
from pathlib import Path

ERRORI = []
sys.stdout.reconfigure(encoding="utf-8")


def controlla(cond, msg):
    print(("OK      " if cond else "ERRORE  ") + msg)
    if not cond:
        ERRORI.append(msg)


def carica(cart, nome):
    return json.loads((cart / f"{nome}.json").read_text(encoding="utf-8"))


def fasce(v):
    """Soglie come nella pagina: v = edifici in classifica, dal primo all'ultimo."""
    b = len(v) // 10
    return b, (lambda k: 1 + (11 - k) * b)  # ultima posizione (da 1) della fascia k


def indice_da_posizione(pos, b):
    if pos == 1:
        return 11  # 10+
    return 10 - (pos - 2) // b if b else None


def main(cart):
    cart = Path(cart)
    oggi, costi, clas = carica(cart, "oggi"), carica(cart, "costi"), carica(cart, "classifiche")
    piano, ob, prezzi = carica(cart, "piano"), carica(cart, "obiettivi")[0], carica(cart, "prezzi")
    storico = carica(cart, "storico")
    c = {r["edificio"]: r for r in costi}
    ris = ["pelle", "oro", "petrolio", "nte"]

    print("== Manutenzione = 3% del costo di costruzione")
    for r in oggi:
        for k in ["denaro"] + ris:
            atteso = sum(r[e] * c[e][k] * 0.03 for e in c)
            letto = r["manutenzione_eur" if k == "denaro" else "manutenzione_" + k]
            controlla(abs(atteso - letto) < 1, f"{r['provincia']} {k}: calcolata {atteso:,.1f}, letta {letto:,.1f}")

    print("== Indici ricavati dalla classifica mondiale")
    campo = {"scuole": "indice_istruzione", "basi": "indice_militare", "strade": "indice_sviluppo", "ospedali": "indice_salute"}
    for e in campo:
        v = [x["edifici"] for x in sorted((x for x in clas if x["edificio"] == e), key=lambda x: x["posizione"])]
        b, _ = fasce(v)
        for r in oggi:
            n = r[e]
            if n == 0:
                continue
            pos_alta = 1 + sum(1 for x in v if x > n)  # posizione migliore a parità
            pos_bassa = sum(1 for x in v if x >= n)   # posizione peggiore a parità
            if pos_bassa < pos_alta:
                # numero di edifici cambiato dopo la lettura della classifica (che si rilegge una volta al giorno):
                # la provincia è ancora nella classifica con il valore vecchio, sopra o sotto
                pos_bassa, pos_alta = pos_alta, max(1, pos_alta - 1)
            ammessi = {indice_da_posizione(p, b) for p in range(pos_alta, pos_bassa + 1)}
            k = r[campo[e]]
            controlla(k in ammessi or (k == 10 and 11 in ammessi),
                      f"{r['provincia']} {e}: {n} edifici, posizione {pos_alta}-{pos_bassa} -> indice {sorted(ammessi)}, gioco {k}")

    print("== Piano: costi, risparmi, saldi")
    m_ped = {k: c["ospedali"][k] * 0.03 for k in ["denaro"] + ris}
    na = next(r for r in oggi if r["provincia"] == "Napoli")
    # risparmio fissato nel piano (dal 22/10/2026); prima era la manutenzione di Napoli dell'ultima lettura
    risp = {"denaro": ob.get("risparmio_notte", na["manutenzione_eur"]), **{k: ob.get("risparmio_notte_" + k, na["manutenzione_" + k]) for k in ris}}
    n_osp = ob.get("palermo_ospedali_nuovi", ob["palermo_ospedali_meta"] - next(r for r in oggi if r["provincia"] == "Palermo")["ospedali"])

    def spesa_giorno(g, k):
        # edifici ricostruiti a Napoli quel giorno (dal 22/10/2026 scritti in ogni riga del piano)
        st = g.get("napoli_strade", 3 if g["giorno"] == 1 else ob["napoli_strade_giorno"])
        osp = g.get("napoli_ospedali", 3 if g["giorno"] == 1 else ob["napoli_ospedali_giorno"])
        return st * c["strade"][k] + osp * c["ospedali"][k] + (n_osp * c["ospedali"][k] if g["giorno"] == 1 else 0) + n_osp * m_ped[k]

    acc = defaultdict(float)
    for g in piano:
        for k in ["denaro"] + ris:
            spesa = spesa_giorno(g, k)
            saldo = risp[k] - spesa
            acc[k] += saldo
            nome_s, nome_a = ("saldo", "accumulato") if k == "denaro" else ("saldo_" + k, "accumulato_" + k)
            if k == "denaro":
                controlla(abs(g["speso"] - spesa) < 1, f"giorno {g['giorno']} speso {g['speso']:,.0f} = {spesa:,.0f}")
            controlla(abs(g[nome_s] - saldo) < 1 and abs(g[nome_a] - acc[k]) < 1,
                      f"giorno {g['giorno']} {k}: saldo {g[nome_s]:,.0f} accumulato {g[nome_a]:,.0f}")
    a, pareggio = 0.0, "mai"  # oltre le due settimane si ripete il saldo dell'ultimo giorno
    for n in range(1, 400):
        a += piano[min(n, len(piano)) - 1]["saldo"]
        if a >= 0:
            pareggio = n
            break
    controlla(pareggio == ob["giorno_pareggio"], f"pareggio in soldi al giorno {pareggio} (obiettivi: {ob['giorno_pareggio']})")
    for k in ris:
        # dopo le due settimane si ripete il saldo dell'ultimo giorno; «mai» se il saldo a regime è negativo
        a, g = 0.0, "mai"
        for n in range(1, 400):
            a += piano[min(n, len(piano)) - 1]["saldo_" + k]
            if a >= 0:
                g = n
                break
        controlla(g == ob["pareggio_" + k], f"pareggio {k} al giorno {g} (obiettivi: {ob['pareggio_' + k]})")
    p = ob.get("prezzi_piano") or prezzi[-1]  # i prezzi con cui è stato calcolato il piano
    val = sum(ob["saldo_notte_" + k] * p[k] for k in ris)
    controlla(abs(val - ob["valore_risorse_notte"]) < 2, f"valore risorse a notte {val:,.0f} (obiettivi: {ob['valore_risorse_notte']:,})")

    print("== Tesori nello storico: cambiano solo per manutenzione e trasferimenti")
    by = defaultdict(dict)
    for r in storico:
        by[r["provincia"]][r["data"]] = r
    spiegati = tot = 0
    for prov, giorni in by.items():
        if prov == "Stato":
            continue
        ds = sorted(giorni)
        for a, b in zip(ds, ds[1:]):
            ra, rb = giorni[a], giorni[b]
            if None in (ra.get("tesoro_eur"), rb.get("tesoro_eur"), rb.get("manutenzione_eur")):
                continue
            resto = rb["tesoro_eur"] - ra["tesoro_eur"] + rb["manutenzione_eur"]
            tot += 1
            spiegati += abs(resto) < 1 or abs(resto) % 10000 < 1
    controlla(tot == 0 or spiegati / tot > 0.7,
              f"{spiegati} giorni su {tot}: tesoro = ieri − manutenzione + trasferimento a cifra tonda (il gettito non entra nel tesoro della provincia)")

    print("== Casi limite delle soglie")
    v = list(range(300, 0, -1))  # 300 province, fasce da 30
    b, ultimo = fasce(v)
    controlla(b == 30 and ultimo(10) == 31 and ultimo(1) == 301, "300 province: fasce da 30, la 10 finisce al 31º posto")
    controlla(indice_da_posizione(2, b) == 10 and indice_da_posizione(31, b) == 10 and indice_da_posizione(32, b) == 9,
              "posizioni 2 e 31 indice 10, posizione 32 indice 9")
    controlla(indice_da_posizione(1, b) == 11, "primo al mondo = 10+")
    controlla(fasce(list(range(9)))[0] == 0, "meno di 10 province: fasce vuote (la pagina mostra «—»)")

    print()
    print("TUTTO OK" if not ERRORI else f"{len(ERRORI)} ERRORI")
    sys.exit(1 if ERRORI else 0)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else ".")
