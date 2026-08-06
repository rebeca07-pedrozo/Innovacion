function probarParidad() {
  const casos = [
    ['Impuesto de Industria y Comercio',       'impuesto,industria,comercio'],
    ['DIFERENCIA EN CAMBIO no está gravada',   'diferencia,cambio,gravada'],
    ['estampillas distritales, hecho generador','estampilla,distrital,hecho,generador'],
    ['Declaraciones y obligaciones tributarias','declaracion,obligacion,tributaria'],
    ['actividades no sujetas del año 2023',    'actividad,sujeta,ano,2023'],
    ['Señor, la Dirección Distrital',          'senor,direccion,distrital'],
    ['Compañía ATM cajeros automáticos',       'compania,atm,cajero,automatico'],
    ['retención ICA sector financiero',        'retencion,ica,sector,financiero'],
    ['Bogotá D.C., 16 de septiembre de 2016',  'bogota,septiembre,2016'],
    ['Radicado 2023ER404739O1',                'radicado,2023er404739o1']
  ];
  let fallos = 0;
  casos.forEach(([entrada, esperado]) => {
    const obtenido = tokenizar_(entrada).join(',');
    if (obtenido !== esperado) {
      fallos++;
      Logger.log('FALLA: "%s"\n  esperado: %s\n  obtenido: %s', entrada, esperado, obtenido);
    }
  });
  Logger.log(fallos === 0 ? 'PARIDAD OK — los ' + casos.length + ' casos coinciden'
                          : fallos + ' casos fallaron');
}

function probarBusqueda() {
  ['diferencia en cambio', 'estampillas distritales', 'ICA sector financiero']
    .forEach(q => {
      const r = buscar(q, 3);
      Logger.log('>>> %s  tokens=%s', q, r.tokens.join(','));
      if (!r.resultados.length) { Logger.log('   (sin resultados)'); return; }
      r.resultados.forEach(x =>
        Logger.log('   %s  p%s  %s', x.puntaje, x.pagina, String(x.archivo).slice(0, 45)));
    });
}

function probarIndice() {
  const idx = cargarIndice_();
  Logger.log('N=%s  avgdl=%s  k1=%s  b=%s', idx.N, idx.avgdl, idx.k1, idx.b);
  Logger.log('términos=%s  fragmentos mapeados=%s',
             Object.keys(idx.post).length, Object.keys(idx.filas).length);
}