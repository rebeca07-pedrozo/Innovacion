const STOPWORDS = new Set(('a al algo algun alguna algunas alguno algunos ante antes aquel aquella ' +
  'aquellas aquello aquellos aqui asi aun aunque cada como con contra cual cuales cuando cuanto de ' +
  'del desde donde dos e el ella ellas ello ellos en entre era eran es esa esas ese eso esos esta ' +
  'estan estas este esto estos fue fueron ha haber habia han hasta hay la las le les lo los mas me ' +
  'mi mientras mismo mucho muy nada ni no nos nuestra nuestro o os otra otras otro otros para pero ' +
  'poco por porque que quien quienes se segun ser si sin sobre son su sus tal tambien tanto te tiene ' +
  'tienen toda todas todo todos tras un una unas uno unos y ya').split(' '));

function sinTildes_(s) {
  return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function normalizar_(s) {
  s = sinTildes_(String(s).toLowerCase());
  s = s.replace(/[^a-z0-9\s]/g, ' ');
  return s.replace(/\s+/g, ' ').trim();
}

function raiz_(p) {
  if (p.length > 4 && p.slice(-2) === 'es') return p.slice(0, -2);
  if (p.length > 3 && p.slice(-1) === 's')  return p.slice(0, -1);
  return p;
}

function tokenizar_(texto) {
  return normalizar_(texto).split(' ')
    .filter(p => p.length > 2 && !STOPWORDS.has(p))
    .map(raiz_);
}