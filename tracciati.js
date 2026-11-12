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

// ordini aperti dello stesso reparto e dello stesso giorno (per le vasche anche dello stesso operatore) → uno solo: stesse righe sommate, ogni riga ricorda da quali ordini viene (parti)
function unisci(ordini) {
  const g = {};
  ordini.forEach(o => {
    // vasche: più operatori sullo stesso reparto, ognuno col suo prelievo → si uniscono solo gli ordini dello stesso operatore
    const rep = String(o.reparto || "").trim().toUpperCase(), op = /VASC?HE/.test(rep) ? "|" + String(o.operatore || "").trim().toUpperCase() : "";
    const k = rep + "|" + String(o.data).split(",")[0] + op;
    const u = g[k] = g[k] || { ...o, ids: [], urgente: false, motivi: [], righe: [] };
    u.ids.push(o.id); u.urgente = u.urgente || !!o.urgente;
    if (o.urgente && o.motivo && !u.motivi.includes(o.motivo)) u.motivi.push(o.motivo);
    o.righe.forEach(r => {
      let x = u.righe.find(x => x.codice === r.codice);
      if (x) x.qta += r.qta; else u.righe.push(x = { codice: r.codice, descrizione: r.descrizione, qta: r.qta, parti: [] });
      x.parti.push({ ordine: o.id, qta: r.qta });
    });
  });
  return Object.values(g);
}

// il prelevato di una riga unita torna sugli ordini di partenza, dal più vecchio; quello in più va sull'ultimo
function ripartisci(qta, parti) {
  let resto = Number(qta) || 0;
  return parti.map((p, i) => {
    const q = i === parti.length - 1 ? resto : Math.min(resto, p.qta);
    resto = Math.round((resto - q) * 1000) / 1000;
    return { ordine: p.ordine, qta: q, attesa: p.qta };
  });
}

if (typeof module !== "undefined") module.exports = { TRACCIATI, rigaExcel, nomeFile, unisci, ripartisci };
