
const MESES_ES = {
  'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4, 'mayo': 5, 'junio': 6,
  'julio': 7, 'agosto': 8, 'septiembre': 9, 'setiembre': 9, 'octubre': 10,
  'noviembre': 11, 'diciembre': 12
};

function parsearFecha(valor, anioReferencia) {
  if (valor === null || valor === undefined || valor === '') {
    return { fecha: null, valida: false, motivo: 'Fecha vacía' };
  }

  if (valor instanceof Date) {
    if (isNaN(valor.getTime())) return { fecha: null, valida: false, motivo: 'Fecha inválida (objeto Date corrupto)' };
    return validarRangoFecha(valor, anioReferencia);
  }

  const texto = String(valor).trim().toLowerCase();
  if (!texto) return { fecha: null, valida: false, motivo: 'Fecha vacía' };

  let match = texto.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) {
    const fecha = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
    return validarRangoFecha(fecha, anioReferencia);
  }

  match = texto.match(/^(\d{1,2})\s+de\s+([a-záéíóúñ]+)\.?$/i);
  if (match) {
    const dia = Number(match[1]);
    const mes = MESES_ES[match[2]];
    if (!mes) return { fecha: null, valida: false, motivo: 'Mes no reconocido: "' + match[2] + '"' };
    const fecha = new Date(anioReferencia, mes - 1, dia);
    return validarRangoFecha(fecha, anioReferencia);
  }

  return { fecha: null, valida: false, motivo: 'Formato de fecha no reconocido: "' + valor + '"' };
}

function validarRangoFecha(fecha, anioReferencia) {
  const anio = fecha.getFullYear();
  if (anio < anioReferencia || anio > anioReferencia + 1) {
    return { fecha: fecha, valida: false, motivo: 'Año sospechoso (' + anio + '), revisar manualmente' };
  }
  return { fecha: fecha, valida: true, motivo: '' };
}

function extraerContactos(celda) {
  if (celda === null || celda === undefined) return [];
  const texto = String(celda).trim();
  if (!texto || texto.toUpperCase() === 'NA') return [];

  const resultados = [];
  const regexConNombre = /([^<>,;]+)<\s*([^<>\s]+@[^<>\s]+)\s*>/g;
  let match;
  let textoRestante = texto;
  while ((match = regexConNombre.exec(texto)) !== null) {
    resultados.push({ nombre: match[1].trim(), email: match[2].trim().toLowerCase() });
    textoRestante = textoRestante.replace(match[0], '');
  }

  const regexEmailSuelto = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
  const emailsYaCapturados = resultados.map(r => r.email);
  let matchEmail;
  while ((matchEmail = regexEmailSuelto.exec(textoRestante)) !== null) {
    const email = matchEmail[0].trim().toLowerCase();
    if (!emailsYaCapturados.includes(email)) {
      resultados.push({ nombre: derivarNombreDeEmail(email), email: email });
      emailsYaCapturados.push(email);
    }
  }

  return resultados;
}

function derivarNombreDeEmail(email) {
  const usuario = email.split('@')[0];
  return usuario.split(/[._]/).map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
}

function contactosAString(contactos) {
  return {
    nombres: contactos.map(c => c.nombre).join('; '),
    emails: contactos.map(c => c.email).join('; ')
  };
}


function limpiarNombreCompania(nombre) {
  return String(nombre).trim().replace(/^[-\s]+/, '').replace(/\s+/g, ' ').toUpperCase();
}


function construirMapaCanonicoNIT(datosExtract, mapaExtract) {
  const conteo = {};
  datosExtract.forEach(fila => {
    const nit = String(valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.NIT) || '').trim();
    const nombreCrudo = valorPorColumna(fila, mapaExtract, CONFIG.COL_EXTRACT.COMPANIA);
    if (!nit || !nombreCrudo) return;
    const nombreLimpio = limpiarNombreCompania(nombreCrudo);
    if (!conteo[nit]) conteo[nit] = {};
    conteo[nit][nombreLimpio] = (conteo[nit][nombreLimpio] || 0) + 1;
  });

  const canonico = {};
  Object.keys(conteo).forEach(nit => {
    let mejorNombre = null, mejorConteo = -1;
    Object.entries(conteo[nit]).forEach(([nombre, veces]) => {
      if (veces > mejorConteo || (veces === mejorConteo && nombre.length > (mejorNombre || '').length)) {
        mejorNombre = nombre;
        mejorConteo = veces;
      }
    });
    canonico[nit] = mejorNombre;
  });
  return canonico;
}