function buscar(consulta, limite) {
  limite = limite || 10;
  const idx = cargarIndice_();
  const tokens = tokenizar_(consulta);
  if (!tokens.length) return { tokens: [], resultados: [] };

  const puntajes = {};
  tokens.forEach(t => {
    const p = idx.post[t];
    if (!p) return;
    const idfT = idx.idf[t];
    for (const cid in p) {
      const f = p[cid], dl = idx.len[cid] || idx.avgdl;
      const num = f * (idx.k1 + 1);
      const den = f + idx.k1 * (1 - idx.b + idx.b * dl / idx.avgdl);
      puntajes[cid] = (puntajes[cid] || 0) + idfT * num / den;
    }
  });

  const orden = Object.keys(puntajes)
    .sort((a, b) => puntajes[b] - puntajes[a])
    .slice(0, limite);
  if (!orden.length) return { tokens: tokens, resultados: [] };

  const hoja = SpreadsheetApp.openById(SHEET_ID).getSheetByName(HOJA_FRAGMENTOS);
  const rangos = orden.map(cid => 'A' + idx.filas[cid] + ':K' + idx.filas[cid]);
  const valores = hoja.getRangeList(rangos).getRanges().map(r => r.getValues()[0]);

  const resultados = orden.map((cid, i) => {
    const v = valores[i];
    return {
      chunk_id: cid, puntaje: Math.round(puntajes[cid] * 100) / 100,
      archivo: v[1], file_id: v[2], pagina: v[3],
      radicado: v[4], fecha: v[5], anio: v[6],
      entidad: v[7], tema: v[8], subtema: v[9], texto: v[10],
      url: 'https://drive.google.com/file/d/' + v[2] + '/view#page=' + v[3]
    };
  });
  return { tokens: tokens, resultados: resultados };
}