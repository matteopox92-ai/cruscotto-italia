# La giornata italiana (00:00–24:00, ora di Roma): regola unica per DEF, recap serale, donazioni e cruscotto
# (decisa dall'utente il 22/10/2026: «tutte le analisi dalle 24 alle 24»).
#
# Il gioco ragiona invece in giornate GMT+3: il riquadro «Oggi» del gettito riparte alle 00:00 GMT+3
# (23:00 italiane con l'ora legale, 22:00 con quella solare) e a quell'ora si pagano manutenzione e confini.
# Per passare da una regola all'altra serve una lettura subito dopo la mezzanotte italiana: il gettito della
# giornata italiana D è quello entrato tra la lettura di mezzanotte di D e quella di D+1 (vedi gettito_tra).
#
# Niente zoneinfo: sul PC manca il pacchetto tzdata e gli script girano con python -I. L'ora legale italiana
# va dall'ultima domenica di marzo all'ultima domenica di ottobre, con il cambio alle 01:00 UTC.
import datetime as dt

UTC = dt.timezone.utc
GMT3 = dt.timezone(dt.timedelta(hours=3))
# una lettura vale «di mezzanotte» se arriva entro le 06:00: se quella delle 00:05 salta (app chiusa) vale la prima
# lettura oraria dopo; il gettito tra la mezzanotte e quella lettura finisce nel giorno prima (gli script lo annotano)
FINESTRA_MEZZANOTTE = dt.timedelta(hours=6)
RITARDO_NOTEVOLE = dt.timedelta(hours=1)


def _ultima_domenica(anno, mese):
    d = dt.date(anno, mese + 1, 1) - dt.timedelta(days=1) if mese < 12 else dt.date(anno, 12, 31)
    return d - dt.timedelta(days=(d.weekday() + 1) % 7)


def scarto_roma(t_utc):
    """Scarto dell'ora di Roma rispetto a UTC nell'istante t (datetime con fuso)."""
    t = t_utc.astimezone(UTC)
    inizio = dt.datetime.combine(_ultima_domenica(t.year, 3), dt.time(1), UTC)
    fine = dt.datetime.combine(_ultima_domenica(t.year, 10), dt.time(1), UTC)
    return dt.timedelta(hours=2 if inizio <= t < fine else 1)


def roma(t):
    """Lo stesso istante con il fuso di Roma."""
    return t.astimezone(dt.timezone(scarto_roma(t)))


def inizio_giornata(d):
    """Mezzanotte italiana del giorno d, come istante con fuso."""
    if isinstance(d, str):
        d = dt.date.fromisoformat(d)
    locale = dt.datetime.combine(d, dt.time())
    # a mezzanotte non c'è mai il cambio d'ora (avviene alle 02:00/03:00): basta lo scarto di quell'ora
    prova = locale.replace(tzinfo=UTC) - dt.timedelta(hours=1)
    return locale.replace(tzinfo=dt.timezone(scarto_roma(prova)))


def fine_giornata(d):
    if isinstance(d, str):
        d = dt.date.fromisoformat(d)
    return inizio_giornata(d + dt.timedelta(days=1))


def giorno_it(t):
    return roma(t).date()


def giorno_gioco(t):
    return t.astimezone(GMT3).date()


def fine_giornata_gioco(g):
    """Fine della giornata del gioco g (24:00 GMT+3): l'ora dell'addebito di manutenzione e confini."""
    if isinstance(g, str):
        g = dt.date.fromisoformat(g)
    return dt.datetime.combine(g + dt.timedelta(days=1), dt.time(), GMT3)


def istante(x):
    """Da stringa ISO (con fuso) a datetime; le stringhe senza fuso sono ora di Roma."""
    if isinstance(x, dt.datetime):
        t = x
    else:
        t = dt.datetime.fromisoformat(str(x))
    if t.tzinfo is None:
        t = t.replace(tzinfo=dt.timezone(scarto_roma(t.replace(tzinfo=UTC) - dt.timedelta(hours=1))))
    return t


def di_mezzanotte(t, d):
    """True se l'istante t è una lettura «di mezzanotte» del giorno d (tra le 00:00 e le 06:00 italiane)."""
    a = inizio_giornata(d)
    return a <= t < a + FINESTRA_MEZZANOTTE


def ora_breve(t):
    return roma(istante(t)).strftime("%d/%m %H:%M")


def gettito_tra(a, b, chiusi):
    """Gettito entrato tra due letture a e b (a prima di b), per una provincia.
    a, b = (istante, [oggi denaro, pelle, oro, petrolio, nte]) con «oggi» = riquadro del gioco (dall'inizio della
    giornata GMT+3); chiusi = {giornata del gioco: [totale esatto]}.
    Conto: tutto quello che entra è oggi(b) − oggi(a), più le giornate del gioco finite tra le due letture.
    Restituisce (valori, giornate mancanti)."""
    ga, gb = giorno_gioco(a[0]), giorno_gioco(b[0])
    if ga == gb:
        return [y - x for x, y in zip(a[1], b[1])], []
    mancano, somma = [], [0] * len(a[1])
    g = ga
    while g < gb:
        if g not in chiusi:
            mancano.append(g)
        else:
            somma = [s + v for s, v in zip(somma, chiusi[g])]
        g += dt.timedelta(days=1)
    return [s + y - x for s, x, y in zip(somma, a[1], b[1])], mancano
