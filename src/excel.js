import * as XLSX from 'xlsx';

const EUR = '#,##0.00 "€"';

function s(ws, r, c, val, opts = {}) {
  const cell = XLSX.utils.encode_cell({ r, c });
  ws[cell] = { v: val, t: typeof val === 'number' ? 'n' : 's' };
  if (opts.numFmt) ws[cell].z = opts.numFmt;
  return ws;
}

function nightsFrom(a, b) {
  if (!a || !b) return 0;
  return Math.max(0, Math.round((new Date(b) - new Date(a)) / 86400000));
}

export function generateInternalExcel(trip, calcResult) {
  const wb = XLSX.utils.book_new();
  const c = calcResult;
  const pax = trip.totalParticipants || 1;

  // Sheet 1: Riepilogo
  const data1 = [
    ['IL SIPARIO MUSICALE — Preventivo Interno'],
    [''],
    ['Nome viaggio', trip.name || ''],
    ['Tipo', trip.type === 'catalogo' ? 'Catalogo' : 'Su misura'],
    ['Date', (trip.dateFrom || '') + ' — ' + (trip.dateTo || '')],
    ['Partecipanti', pax],
    [''],
    ['BREAKDOWN COSTI A PERSONA'],
  ];

  c.legBreakdowns.forEach((lb, i) => {
    data1.push(['Hotel ' + (lb.leg.destination || 'Tappa ' + (i + 1)), lb.hotelBase]);
  });
  data1.push(['Complessivi pro-capite', c.groupPP]);
  data1.push(['Individuali', c.totalIndAll]);
  data1.push(['']);
  data1.push(['Subtotale pp', c.subtotalPP]);
  data1.push(['Utile pp', trip.markupPerPerson || 0]);
  c.commDetails.forEach(cd => {
    data1.push([cd.name + ' (' + cd.pct + '%)', cd.value]);
  });
  data1.push(['TOTALE PACCHETTO PP', c.total]);
  data1.push(['ARROTONDATO PP', c.rounded]);
  data1.push(['']);
  data1.push(['TOTALI VIAGGIO']);
  data1.push(['Costo totale viaggio', c.subtotalPP * pax]);
  data1.push(['Ricavo totale', c.rounded * pax]);
  data1.push(['Margine totale', c.rounded * pax - c.subtotalPP * pax]);
  data1.push(['']);

  if (c.supplements.length > 0) {
    data1.push(['SUPPLEMENTI CAMERA']);
    data1.push(['Camera', 'Diff/notte', 'Totale', 'Venduto a']);
    c.supplements.forEach(s => {
      data1.push([s.room.name, s.diffPerNight, s.total, s.rounded]);
    });
  }

  const ws1 = XLSX.utils.aoa_to_sheet(data1);
  ws1['!cols'] = [{ wch: 35 }, { wch: 18 }, { wch: 14 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, ws1, 'Riepilogo');

  // Sheet 2: Programma per tappa
  trip.legs.forEach((leg, li) => {
    const rows = [
      [(leg.destination || 'Tappa ' + (li + 1)) + ' — ' + (leg.hotelName || '')],
      ['Data', 'Attività', 'Tipo', 'Importo', 'Staff'],
    ];
    (leg.days || []).forEach(day => {
      rows.push([day.date || '', day.title || '', '', '', '']);
      (day.activities || []).forEach(act => {
        const staffNames = (act.staffMirror || []).map(sid => {
          const st = trip.staff.find(s => s.id === sid);
          return st ? st.role : '';
        }).join(', ');
        rows.push(['', act.description || '', act.costType === 'group' ? 'Complessivo' : 'Individuale', act.amount || 0, staffNames]);
      });
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 14 }, { wch: 35 }, { wch: 14 }, { wch: 12 }, { wch: 20 }];
    XLSX.utils.book_append_sheet(wb, ws, (leg.destination || 'Tappa ' + (li + 1)).slice(0, 28));
  });

  // Sheet 3: Staff
  const staffRows = [
    ['STAFF & COSTI'],
    ['Ruolo', 'Nome', 'Onorario', 'Hotel', 'Volo', 'Totale'],
  ];
  trip.staff.forEach(s => {
    const tot = (s.totalFee || 0) + (s.totalHotel || 0) + (s.flight || 0);
    staffRows.push([s.role, s.name, s.totalFee || 0, s.totalHotel || 0, s.flight || 0, tot]);
  });
  staffRows.push(['']);
  staffRows.push(['COSTI VARI']);
  trip.miscCosts.forEach(m => {
    staffRows.push([m.description || '', m.amount || 0]);
  });
  const ws3 = XLSX.utils.aoa_to_sheet(staffRows);
  ws3['!cols'] = [{ wch: 18 }, { wch: 18 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];
  XLSX.utils.book_append_sheet(wb, ws3, 'Staff & Costi');

  // Sheet 4: Partecipanti
  if (trip.participants && trip.participants.length > 0) {
    const partRows = [['Nome', 'Cognome', 'Camera']];
    trip.participants.forEach(p => {
      let roomName = '';
      trip.legs.forEach(leg => {
        const chosen = leg.roomTypes.find(r => r.id === p.roomChoices?.[leg.id]) || leg.roomTypes[leg.baseRoomIndex];
        if (chosen) roomName = chosen.name;
      });
      partRows.push([p.firstName, p.lastName, roomName]);
    });
    const ws4 = XLSX.utils.aoa_to_sheet(partRows);
    ws4['!cols'] = [{ wch: 18 }, { wch: 18 }, { wch: 22 }];
    XLSX.utils.book_append_sheet(wb, ws4, 'Partecipanti');
  }

  XLSX.writeFile(wb, (trip.name || 'preventivo') + '_interno.xlsx');
}

export function generateClientExcelMisura(trip, allItems, basePPfinal, extraItems) {
  const wb = XLSX.utils.book_new();
  const pax = trip.totalParticipants || 1;
  const round5 = v => Math.ceil(v / 5) * 5;

  const data = [
    ['IL SIPARIO MUSICALE'],
    ['Itinerari di Musica ed Arte'],
    [''],
    [trip.name || 'Viaggio'],
    [(trip.dateFrom || '') + ' — ' + (trip.dateTo || '')],
    [''],
    ['PACCHETTO BASE'],
    ['Prezzo per persona', round5(basePPfinal)],
    [''],
  ];

  if (extraItems.length > 0) {
    data.push(['SERVIZI EXTRA (opzionali)']);
    data.push(['Servizio', 'Prezzo']);
    extraItems.forEach(item => {
      data.push([item.label, item.clientPrice || 'Da definire']);
    });
    data.push(['']);
  }

  // Programma
  data.push(['PROGRAMMA']);
  trip.legs.forEach(leg => {
    data.push([leg.destination + (leg.hotelName ? ' — ' + leg.hotelName : '')]);
    (leg.days || []).forEach(day => {
      data.push([day.date || '', day.title || '']);
      (day.activities || []).forEach(act => {
        data.push(['', act.description || '']);
      });
    });
    data.push(['']);
  });

  data.push(['']);
  data.push(['Il Sipario Musicale S.r.l.']);
  data.push(['Via Molino delle Armi 11 — 20123 Milano']);
  data.push(['Tel. +39 02 5834941']);

  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{ wch: 40 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Offerta');
  XLSX.writeFile(wb, (trip.name || 'offerta') + '_cliente.xlsx');
}

export function generateClientPdfCatalogo(trip, calcResult) {
  const c = calcResult;
  const round5 = v => Math.ceil(v / 5) * 5;
  const fmt = v => new Intl.NumberFormat('it-IT', { style: 'currency', currency: 'EUR' }).format(v);
  const fmtDate = d => { if (!d) return ''; try { return new Date(d + 'T00:00:00').toLocaleDateString('it-IT', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }); } catch { return d; } };

  const w = window.open('', '_blank');
  if (!w) { alert('Abilita i popup per generare il PDF'); return; }

  let programHtml = '';
  trip.legs.forEach((leg, li) => {
    programHtml += `<h2 style="color:#8B1A1A;font-size:16px;margin:20px 0 8px;border-bottom:2px solid #E6DFD4;padding-bottom:4px">${leg.destination || 'Tappa ' + (li + 1)}${leg.hotelName ? ' — ' + leg.hotelName : ''}</h2>`;
    (leg.days || []).forEach((day, di) => {
      programHtml += `<div style="margin-bottom:12px"><div style="font-weight:700;font-size:13px;color:#8B1A1A;margin-bottom:4px">${day.date ? fmtDate(day.date) : 'Giorno ' + (di + 1)}${day.title ? ' — ' + day.title : ''}</div>`;
      (day.activities || []).forEach(act => {
        programHtml += `<div style="padding-left:16px;font-size:12px;margin:2px 0">&bull; ${act.description || ''}</div>`;
      });
      programHtml += `</div>`;
    });
  });

  let suppHtml = '';
  if (c.supplements.length > 0) {
    suppHtml = `<div style="margin-top:16px"><h3 style="font-size:13px;color:#8B1A1A;margin-bottom:6px">Supplementi camera</h3>`;
    c.supplements.forEach(s => {
      suppHtml += `<div style="display:flex;justify-content:space-between;font-size:12px;padding:3px 0;border-bottom:1px solid #eee"><span>${s.room.name}</span><span style="font-weight:700;color:#C49B2B">${fmt(s.rounded)}</span></div>`;
    });
    suppHtml += `</div>`;
  }

  w.document.write(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>${trip.name || 'Preventivo'}</title>
  <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;700&family=Nunito+Sans:wght@400;600;700&display=swap" rel="stylesheet">
  <style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Nunito Sans',sans-serif;font-size:12px;color:#2D2218;padding:32px;max-width:700px;margin:0 auto}
  h1{font-family:'Cormorant Garamond',serif;font-size:26px;color:#8B1A1A}
  .hdr{text-align:center;padding-bottom:24px;margin-bottom:24px;border-bottom:3px double #C49B2B}
  .price-box{background:#FDF8EC;border:1.5px solid #C49B2B;border-radius:10px;padding:20px;text-align:center;margin:24px 0}
  .footer{margin-top:32px;padding-top:16px;border-top:1px solid #E6DFD4;text-align:center;font-size:10px;color:#8A7E70}
  @media print{body{padding:16px}}</style></head><body>
  <div class="hdr">
    <div style="font-size:10px;font-weight:700;letter-spacing:0.15em;text-transform:uppercase;color:#C49B2B;margin-bottom:8px">Il Sipario Musicale</div>
    <h1>${trip.name || 'Viaggio Culturale'}</h1>
    <div style="color:#8A7E70;margin-top:6px">${trip.dateFrom ? fmtDate(trip.dateFrom) : ''} ${trip.dateTo ? '— ' + fmtDate(trip.dateTo) : ''}</div>
  </div>
  ${programHtml}
  <div class="price-box">
    <div style="font-size:11px;font-weight:700;color:#C49B2B;text-transform:uppercase;letter-spacing:0.08em;margin-bottom:8px">Quota di partecipazione</div>
    <div style="font-family:'Cormorant Garamond',serif;font-size:32px;font-weight:700;color:#8B1A1A">${fmt(c.rounded)}</div>
    <div style="font-size:11px;color:#8A7E70;margin-top:4px">per persona in camera base</div>
  </div>
  ${suppHtml}
  <div class="footer">
    <div>Il Sipario Musicale Itinerari di Musica ed Arte S.r.l.</div>
    <div>Via Molino delle Armi 11 — 20123 Milano &middot; Tel. +39 02 5834941</div>
    <div>P.IVA 01044750675 &middot; Licenza 91075/06</div>
  </div>
  <script>setTimeout(()=>window.print(),500)<\/script>
  </body></html>`);
  w.document.close();
}
