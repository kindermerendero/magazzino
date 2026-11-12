// ponytail: tracciati generici, da sostituire con quelli del gestionale per ciascuna funzione
const TRACCIATI = {
  carico:     ["Data", "Operatore", "Codice", "Descrizione", "Qta bolla", "Qta contata", "Differenza", "Esito", "Qta ordinata"],
  scarico:    ["Data", "Reparto", "Operatore", "Codice", "Descrizione", "Qta richiesta", "Qta prelevata", "Mancante", "Esito", "Motivo"],
  inventario: ["Data", "Ubicazione", "Operatore", "Codice", "Descrizione", "Qta contata"],
  ordini:     ["Data", "Reparto", "Operatore", "Codice", "Descrizione", "Qta richiesta"],
  acquisti:   ["Data", "Operatore", "Codice", "Descrizione", "Giacenza", "Urgente", "Scorta min", "Scorta max"],
};

function rigaExcel(tipo, data, rif, operatore, r) {
  // carico e acquisti senza riferimento; nello scarico è il reparto dell'ordine evaso (vuoto se scarico libero)
  const qta = Number(r.qta), base = [data, ...(tipo === "carico" || tipo === "acquisti" ? [] : [rif]), operatore, r.codice, r.descrizione || ""];
  if (tipo === "scarico") {
    if (r.richiesta === undefined || r.richiesta === "") return [...base, "", qta, "", "", r.causale || ""];
    const rich = Number(r.richiesta);
    return [...base, rich, qta, Math.max(0, rich - qta), qta >= rich ? "EVASO" : qta === 0 ? "NON DISPONIBILE" : "PARZIALE", r.causale || ""];
  }
  if (tipo === "acquisti") return [...base, Number(r.giacenza), r.urgente ? "SI" : "", ...(r.scorta || ["", ""])];
  if (tipo !== "carico") return [...base, qta];
  const dich = Number(r.dichiarata), diff = qta - dich;
  return [...base, dich, qta, diff, diff === 0 ? "OK" : "NON CONFORME", r.richiesta ?? ""];
}

function nomeFile(tipo, d, rif) {
  const p = n => String(n).padStart(2, "0");
  const ts = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
  const s = rif.replace(/[^A-Za-z0-9_-]+/g, "-").slice(0, 40);
  return `${tipo}_${ts}${s ? "_" + s : ""}.xlsx`;
}

if (typeof module !== "undefined") module.exports = { TRACCIATI, rigaExcel, nomeFile };

// "12/11/26, 14:30" (data dei movimenti, scritta dal telefono) → Date; null se non si legge
function quando(t) {
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(?:,?\s+(\d{1,2})[:.](\d{2}))?/.exec(String(t ?? "").trim());
  return m ? new Date(+m[3] + (m[3].length === 2 ? 2000 : 0), m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0)) : null;
}

// indicatori della dashboard sugli ultimi `giorni`. mov: righe della scheda movimenti
// (Doc, Data, Tipo, Operatore, Rif, Codice, Descrizione, Qta, Attesa, Ordine, Urgente, Motivo), dalla più vecchia
function kpi(mov, aperti, giacenze, scorte, nProdotti, ora, giorni) {
  const G = 864e5, n = v => Number(v) || 0, r3 = v => Math.round(v * 1000) / 1000, pct = (a, b) => b ? Math.round(a / b * 100) : null;
  mov = mov.filter(m => { const d = quando(m[1]); return !d || d >= ora - giorni * G; });
  // ordini reparto aperti da più di un giorno
  const eta = aperti.filter(o => o.tipo === "ordini").map(o => quando(o.data)).filter(Boolean).map(d => (ora - d) / G).filter(g => g > 1);
  // righe d'ordine passate dallo scarico: prelevato (Qta) contro richiesto (Attesa)
  const prel = mov.filter(m => m[2] === "scarico" && m[9] && n(m[8]) > 0), manc = {};
  prel.filter(m => n(m[7]) < n(m[8])).forEach(m => {
    const x = manc[m[5]] = manc[m[5]] || { codice: m[5], descrizione: m[6], volte: 0, qta: 0 };
    x.volte++; x.qta = r3(x.qta + n(m[8]) - n(m[7]));
  });
  // sotto la scorta minima e in nessuna richiesta d'acquisto aperta
  const richiesti = new Set(aperti.filter(o => o.tipo === "acquisti").flatMap(o => o.righe.map(r => r.codice)));
  const sotto = Object.keys(giacenze).filter(c => scorte[c] && scorte[c][0] !== "" && giacenze[c] < scorte[c][0]);
  // carichi con la bolla compilata: contato diverso dalla bolla
  const car = mov.filter(m => m[2] === "carico" && m[8] !== ""), nc = car.filter(m => n(m[7]) !== n(m[8]));
  // tempo di evasione: dall'ordine all'ultimo scarico (o chiusura) che lo riguarda, solo per gli ordini non più aperti
  const nati = {}, chiusi = {}, idAperti = new Set(aperti.map(o => o.id));
  mov.forEach(m => {
    const d = quando(m[1]); if (!d) return;
    if (m[2] === "ordini") nati[m[0]] = d; else if ((m[2] === "scarico" || m[2] === "chiusura") && m[9]) chiusi[m[9]] = d;
  });
  const tempi = Object.keys(nati).filter(id => chiusi[id] && !idAperti.has(id)).map(id => (chiusi[id] - nati[id]) / 36e5);
  // consumo: le quantità si sommano solo per codice (unità di misura diverse), per reparto si contano prelievi e codici
  const rep = {}, art = {};
  mov.filter(m => m[2] === "scarico" && n(m[7]) > 0).forEach(m => {
    const k = String(m[4]).trim().toUpperCase() || "SENZA REPARTO", x = rep[k] = rep[k] || { reparto: k, prelievi: 0, codici: new Set() };
    x.prelievi++; x.codici.add(m[5]);
    const a = art[m[5]] = art[m[5]] || { codice: m[5], descrizione: m[6], qta: 0, prelievi: 0 };
    a.qta = r3(a.qta + n(m[7])); a.prelievi++;
  });
  return {
    vecchi: eta.length, piuVecchio: eta.length ? Math.floor(Math.max(...eta)) : 0,
    evasione: pct(prel.length - prel.filter(m => n(m[7]) < n(m[8])).length, prel.length), righeOrdine: prel.length,
    mancanti: Object.values(manc).sort((a, b) => b.volte - a.volte || b.qta - a.qta).slice(0, 10),
    sotto: sotto.length, daRichiedere: sotto.filter(c => !richiesti.has(c)),
    nonConformi: pct(nc.length, car.length), carichi: car.length, righeNC: nc.slice().reverse(),
    copertura: pct(Object.keys(giacenze).length, nProdotti), contati: Object.keys(giacenze).length,
    tempo: tempi.length ? tempi.reduce((a, b) => a + b, 0) / tempi.length : null, evasi: tempi.length,
    reparti: Object.values(rep).map(x => ({ ...x, codici: x.codici.size })).sort((a, b) => b.prelievi - a.prelievi),
    articoli: Object.values(art).sort((a, b) => b.prelievi - a.prelievi || b.qta - a.qta).slice(0, 10),
  };
}
if (typeof module !== "undefined") Object.assign(module.exports, { quando, kpi });
