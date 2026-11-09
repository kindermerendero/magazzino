// ponytail: tracciati generici, da sostituire con quelli del gestionale per ciascuna funzione
const TRACCIATI = {
  carico:     ["Data", "Operatore", "Codice", "Descrizione", "Qta bolla", "Qta contata", "Differenza", "Esito", "Qta ordinata"],
  scarico:    ["Data", "Reparto", "Operatore", "Codice", "Descrizione", "Qta richiesta", "Qta prelevata", "Mancante", "Esito"],
  inventario: ["Data", "Ubicazione", "Operatore", "Codice", "Descrizione", "Qta contata"],
  ordini:     ["Data", "Reparto", "Operatore", "Codice", "Descrizione", "Qta richiesta"],
  acquisti:   ["Data", "Operatore", "Codice", "Descrizione", "Giacenza", "Urgente"],
};

function rigaExcel(tipo, data, rif, operatore, r) {
  // carico e acquisti senza riferimento; nello scarico è il reparto dell'ordine evaso (vuoto se scarico libero)
  const qta = Number(r.qta), base = [data, ...(tipo === "carico" || tipo === "acquisti" ? [] : [rif]), operatore, r.codice, r.descrizione || ""];
  if (tipo === "scarico") {
    if (r.richiesta === undefined || r.richiesta === "") return [...base, "", qta, "", ""];
    const rich = Number(r.richiesta);
    return [...base, rich, qta, Math.max(0, rich - qta), qta >= rich ? "EVASO" : qta === 0 ? "NON DISPONIBILE" : "PARZIALE"];
  }
  if (tipo === "acquisti") return [...base, Number(r.giacenza), r.urgente ? "SI" : ""];
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
