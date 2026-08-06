function guardarCache_(clave, texto) {
  const cache = CacheService.getScriptCache();
  const T = 90000, n = Math.ceil(texto.length / T), partes = {};
  for (let i = 0; i < n; i++) partes[clave + '_' + i] = texto.substr(i * T, T);
  partes[clave + '_n'] = String(n);
  cache.putAll(partes, TTL);
}

function leerCache_(clave) {
  const cache = CacheService.getScriptCache();
  const n = cache.get(clave + '_n');
  if (!n) return null;
  const claves = [];
  for (let i = 0; i < Number(n); i++) claves.push(clave + '_' + i);
  const partes = cache.getAll(claves);
  let out = '';
  for (const k of claves) {
    if (partes[k] === undefined) return null;   // Cuando expira un fragmento, invalida todo el caché
    out += partes[k];
  }
  return out;
}

function limpiarCache_() {
  CacheService.getScriptCache().remove('idx_n');
  Logger.log('Caché invalidada — la próxima búsqueda recarga desde Sheets');
}

function cargarIndice_() {
  const enCache = leerCache_('idx');
  if (enCache) return JSON.parse(enCache);

  const ss = SpreadsheetApp.openById(SHEET_ID);

  const stats = {};
  ss.getSheetByName(HOJA_STATS).getDataRange().getValues().slice(1)
    .forEach(([k, v]) => stats[k] = v);

  const post = {}, idf = {};
  ss.getSheetByName(HOJA_INDICE).getDataRange().getValues().slice(1)
    .forEach(([termino, df, valorIdf, postings]) => {
      if (!termino) return;
      idf[termino] = Number(valorIdf);
      const m = {};
      String(postings).split(',').forEach(par => {
        const i = par.lastIndexOf(':');
        if (i > 0) m[par.slice(0, i)] = Number(par.slice(i + 1));
      });
      post[termino] = m;
    });

  const filas = {};
  ss.getSheetByName(HOJA_FRAGMENTOS).getDataRange().getValues()
    .forEach((r, i) => { if (i > 0 && r[0]) filas[r[0]] = i + 1; });

  const longitudes = {};
  Object.keys(stats).forEach(k => {
    if (k.indexOf('len:') === 0) longitudes[k.slice(4)] = Number(stats[k]);
  });

  const idx = {
    N: Number(stats['N']), avgdl: Number(stats['avgdl']),
    k1: Number(stats['k1']), b: Number(stats['b']),
    post: post, idf: idf, filas: filas, len: longitudes
  };
  guardarCache_('idx', JSON.stringify(idx));
  return idx;
}