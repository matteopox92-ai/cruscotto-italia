"""Crea una copia autonoma del Cruscotto Economia Italia: un solo file HTML con pagina, dati e d3 dentro.

Uso: python cruscotto_offline.py <cartella Cruscotto> <file .html di uscita>
Funziona aperto con doppio clic, anche senza internet. Rifallo dopo ogni aggiornamento dei dati.
"""
import json
import sys
from datetime import datetime
from pathlib import Path

DATASET = ["oggi", "storico", "prezzi", "classifiche", "costi", "piano", "obiettivi", "diario", "donazioni", "donazioni_totali", "governi"]

STILE = """
:root { color-scheme: dark; /* sempre scuro: preferenza dell'utente (22/10/2026) */
  --color-bg: #12161c; --color-panel: #1a2029; --color-fg: #e6eaf0; --color-fg-muted: #9aa5b3; --color-border-line: #2c3440;
  --color-ok: #5cc98e; --color-warn: #e0a83a; --color-bad: #f07167;
  --cds-chart-grid: #262e39; --cds-chart-axis: #4a5462; --cds-chart-reference: #7c8796; --cds-chart-status-critical: #f07167;
  --cds-surface-popover: #222a35; --cds-border: #3a4350; --cds-shadow-popover: 0 4px 16px rgba(0,0,0,.4);
  --font-anthropic-sans: "Segoe UI", Arial, sans-serif; --font-anthropic-serif: Georgia, "Times New Roman", serif;
  --cds-font-size-caption: 12px; --cds-font-size-body: 14px; --cds-font-size-heading: 16px; --cds-font-size-title: 26px;
  --cds-font-weight-medium: 600; --cds-radius: 8px;
  --cds-gap-xs: 4px; --cds-gap-sm: 8px; --cds-gap-md: 12px; --cds-gap-lg: 20px; --cds-gap-xl: 28px;
  --cds-pad-xs: 4px; --cds-pad-sm: 8px; --cds-pad-md: 12px; --cds-pad-lg: 16px; --cds-pad-xl: 24px;
  --cds-dur-slow: .3s; --cds-ease-out: ease-out; }
html, body { margin: 0; background: var(--color-bg); color: var(--color-fg); font-family: var(--font-anthropic-sans); font-size: 14px; }
#dash-root { max-width: 1280px; margin: 0 auto; padding: 20px max(16px, 3vw) 40px; }
.offline-nota { font-size: 12px; color: var(--color-fg-muted); margin: 0 0 12px; }
[hidden] { display: none !important; }
"""

SHIM = """
(function () {
  const DATI = window.__DATI;
  const calcs = {}, ordine = [], ascoltatori = [], cache = {};
  const colonne = (rows) => rows && rows.length ? Object.keys(rows[0]) : [];
  const risolvi = (id) => {
    if (id in cache) return cache[id];
    if (id in DATI) return (cache[id] = DATI[id]);
    const c = calcs[id];
    if (!c) return (cache[id] = null);
    const ingressi = c.inputs.map((i) => risolvi(i) || []);
    try { cache[id] = c.fn(...ingressi); } catch (e) { console.error('calcolo', id, e); cache[id] = null; }
    return cache[id];
  };
  window.dash = {
    data(id) {
      const d = risolvi(id);
      return d == null ? { status: 'missing', data: [], columns: [], rows: [] } : { status: 'ok', data: d, rows: d, columns: colonne(d), meta: [], refreshing: false };
    },
    calc(id, def) { calcs[id] = def; ordine.push(id); },
    loader() {},
    onData(fn) { ascoltatori.push(fn); if (document.readyState !== 'loading') try { fn(); } catch (e) { console.error(e); } },
    colors: ['#3b6fd6', '#d9822b', '#2e9e6a', '#9b59b6', '#c0392b', '#16a2b8', '#7f8c8d', '#b8860b'],
    setLink(id) { try { history.replaceState(null, '', '#' + id); } catch (e) {} },
    // filtri della pagina: nessuna query qui, quindi cambiano solo il disegno
    params() { return { ...parametri }; },
    setParams(o) { Object.assign(parametri, o); tutti(); },
    resetParams() { for (const k in parametri) delete parametri[k]; tutti(); },
    refresh() {},
  };
  const parametri = {};
  const tutti = () => ascoltatori.forEach((f) => { try { f(); } catch (e) { console.error(e); } });
  document.addEventListener('DOMContentLoaded', tutti);
  let t; window.addEventListener('resize', () => { clearTimeout(t); t = setTimeout(tutti, 150); });
})();
"""


def main(cartella, uscita):
    c = Path(cartella)
    dati = {k: json.loads((c / f"{k}.json").read_text(encoding="utf-8")) for k in DATASET if (c / f"{k}.json").exists()}
    pagina = (c / "index.html").read_text(encoding="utf-8")
    d3 = (Path(__file__).parent / "lib" / "d3-7.9.0.min.js").read_text(encoding="utf-8")
    ultima = max((r.get("data", "") for r in dati.get("oggi", [])), default="")
    json_dati = json.dumps(dati, ensure_ascii=False).replace("</", "<\\/")
    html = f"""<!doctype html>
<html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Cruscotto Economia Italia</title>
<style>{STILE}</style>
<script>{d3}</script>
<script>window.__DATI = {json_dati};</script>
<script>{SHIM}</script>
</head><body><div id="dash-root">
<p class="offline-nota">Copia autonoma del cruscotto, dati del gioco aggiornati al {ultima}. File creato il {datetime.now().strftime('%d/%m/%Y %H:%M')}.</p>
{pagina}
</div></body></html>
"""
    Path(uscita).write_text(html, encoding="utf-8")
    print("salvato", uscita, round(len(html.encode()) / 1024), "KB", list(dati))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
