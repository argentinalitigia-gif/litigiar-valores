// ═════════════════════════════════════════════════════════════════════
//  LitigiAR — Scraper de valores oficiales (GitHub Actions)
//  Corre cada 6hs, produce valores.json en raíz del repo.
//
//  AUTO-SCRAPEADOS:
//    - datos.gob.ar (IPC Nacional, UVA, CER, UVI, RIPTE, SMVM, BADLAR)
//    - BCRA diar_icl.xls (ICL)
//    - ColProBA / La Matanza / Morón (JUS, Bonos)
//    - INDEC Canasta Crianza XLSX
//    - IDECBA IPC CABA/GBA (HTML scrape)
//
//  MANUAL (requieren edición periódica de valores.json):
//    - tasas.acta2601, acta2630, acta2658  (CNAT publica planillas PDF)
//    - tasas.bnaActiva, bnaPasiva, bnaLibre36, bnaLibre72 (BNA sitio dinámico)
//    - tasas.bpActiva, bpPasiva (BPBA sitio dinámico)
//
//  Política tolerante a fallos:
//    - Si un scraper falla, mantiene el valor anterior con status="stale"
//    - Si nunca hubo valor, cae a fallback con status="fallback"
//    - Los valores "manual" mantienen status="manual" hasta próxima edición
// ═════════════════════════════════════════════════════════════════════

const fs      = require('fs');
const path    = require('path');
const axios   = require('axios');
const cheerio = require('cheerio');
const XLSX    = require('xlsx');

const VALORES_PATH = path.join(__dirname, '..', 'valores.json');

// ── FALLBACK (congelado al 19/04/2026) ──
const FALLBACK = {
  pba: {
    jusLey14967:    {valor:49750, vigenteDesde:"2026-04-01", fuenteNombre:"SCBA Acuerdo 4222", fuenteUrl:"https://www.scba.gov.ar/", status:"fallback"},
    jusDL8904:      {valor:33998, vigenteDesde:"2026-04-01", fuenteNombre:"SCBA Acuerdo 4222", fuenteUrl:"https://www.scba.gov.ar/", status:"fallback"},
    jusPrevisional: {valor:24000, vigenteDesde:"2025-08-01", fuenteNombre:"Caja de la Abogacía PBA", fuenteUrl:"https://www.cajaabogados.org.ar/", status:"fallback"},
    bonoVerde8480:  {valor:23100, fuenteNombre:"Colegio Departamental / ColProBA", fuenteUrl:"https://colproba.org.ar/", status:"fallback"},
    bonoAzul8480:   {valor:11550, fuenteNombre:"Colegio Departamental / ColProBA", fuenteUrl:"https://colproba.org.ar/", status:"fallback"}
  },
  ipc:  {periodo:"2026-03", mensual:3.4, interanual:32.6, acumulada:9.4, fuenteNombre:"INDEC", fuenteUrl:"https://apis.datos.gob.ar/series/api/series/?ids=148.3_INIVELNAL_DICI_M_26", status:"fallback"},
  bcra: {
    uva: {fecha:"2026-04-19", valor:1886.96, fuenteNombre:"BCRA", fuenteUrl:"https://www.bcra.gob.ar/", status:"fallback"},
    cer: {fecha:"2026-04-19", valor:748.8806, fuenteNombre:"BCRA", fuenteUrl:"https://www.bcra.gob.ar/", status:"fallback"},
    icl: {fecha:"2026-04-19", valor:31.62,   fuenteNombre:"BCRA diar_icl.xls", fuenteUrl:"https://www.bcra.gob.ar/Pdfs/PublicacionesEstadisticas/diar_icl.xls", status:"fallback"},
    uvi: {fecha:"2026-04-19", valor:1722.45, fuenteNombre:"BCRA", fuenteUrl:"https://www.bcra.gob.ar/", status:"fallback"}
  },
  ripte:          {periodo:"2026-02", valor:1128450, fuenteNombre:"MTSS (datos.gob.ar)", fuenteUrl:"https://apis.datos.gob.ar/series/api/series/?ids=158.1_REPTE_0_0_5", status:"fallback"},
  smvm:           {periodo:"2026-02", valor:391000, vigenteDesde:"2026-02-01", fuenteNombre:"Res. 3/2026 CNEPySMVyM", fuenteUrl:"https://apis.datos.gob.ar/series/api/series/?ids=57.1_SMVMM_0_M_34", status:"fallback"},
  canastaCrianza: {periodo:"2026-03", menor1:515236, edad1a3:616046, edad4a5:538587, edad6a12:676431, fuenteNombre:"INDEC Canasta de Crianza", fuenteUrl:"https://www.indec.gob.ar/ftp/cuadros/sociedad/serie_canasta_crianza.xlsx", status:"fallback"},
  // ── IPC CABA/GBA (IDECBA) — se intenta scrape, sino fallback ──
  ipcGba: {periodo:"2026-03", mensual:3.2, interanual:31.8, acumulada:9.1, fuenteNombre:"IDECBA Dir. Estadística CABA", fuenteUrl:"https://www.estadisticaciudad.gob.ar/eyc/?p=indice_de_precios_al_consumidor", status:"fallback"},
  // ── Tasas (MANUAL) — editar valores.json directamente cuando cambien.
  //     Fuente: CNAT Planilla actualización (PDF mensual) + BNA.com.ar + BPBA.com.ar
  tasas: {
    // Valores fallback actualizados a abril 2026 (post-desinflación). Verificar
    // mensualmente contra https://www.cnat.gob.ar/ y https://www.bna.com.ar/
    // Cuando Playwright corra en GitHub Actions, estas se reescriben con scrape directo.
    // Status="fallback" → estos valores SÍ se actualizan en cada release del scraper.
    acta2601:   {tasaAnual:42,  vigenteDesde:"2026-04-01", fuenteNombre:"CNAT Planilla Acta 2601 (Activa Cart. Gral. BNA 30 días) — fallback editorial", fuenteUrl:"https://www.cnat.gob.ar/", status:"fallback", nota:"Estimación — verificar contra planilla CNAT mensual"},
    acta2630:   {tasaAnual:45,  vigenteDesde:"2026-04-01", fuenteNombre:"CNAT Planilla Acta 2630 (Nominal 36 meses BNA) — fallback editorial", fuenteUrl:"https://www.cnat.gob.ar/", status:"fallback", nota:"Estimación — verificar contra planilla CNAT mensual"},
    acta2658:   {tasaAnual:50,  vigenteDesde:"2026-04-01", fuenteNombre:"CNAT Planilla Acta 2658 (Efectiva anual 49-60m BNA) — fallback editorial", fuenteUrl:"https://www.cnat.gob.ar/", status:"fallback", nota:"Estimación — verificar contra planilla CNAT mensual"},
    bnaActiva:  {tasaAnual:42,  vigenteDesde:"2026-04-01", fuenteNombre:"BNA Tasa Activa Cartera General 30 días — fallback editorial", fuenteUrl:"https://www.bna.com.ar/", status:"fallback", nota:"Estimación — verificar contra www.bna.com.ar"},
    bnaPasiva:  {tasaAnual:21,  vigenteDesde:"2026-04-01", fuenteNombre:"BNA Tasa Pasiva — fallback editorial", fuenteUrl:"https://www.bna.com.ar/", status:"fallback"},
    bnaLibre36: {tasaAnual:48,  vigenteDesde:"2026-04-01", fuenteNombre:"BNA Préstamo Libre 36 meses — fallback editorial", fuenteUrl:"https://www.bna.com.ar/", status:"fallback", nota:"Estimación — verificar contra www.bna.com.ar"},
    bnaLibre72: {tasaAnual:52,  vigenteDesde:"2026-04-01", fuenteNombre:"BNA Préstamo Libre 72 meses — fallback editorial", fuenteUrl:"https://www.bna.com.ar/", status:"fallback", nota:"Estimación — verificar contra www.bna.com.ar"},
    bpActiva:   {tasaAnual:75,  vigenteDesde:"2026-04-01", fuenteNombre:"BPBA Activa Cartera General", fuenteUrl:"https://www.bancoprovincia.com.ar/", status:"manual"},
    bpPasiva:   {tasaAnual:45,  vigenteDesde:"2026-04-01", fuenteNombre:"BPBA Pasiva", fuenteUrl:"https://www.bancoprovincia.com.ar/", status:"manual"},
    badlar:     {tasaAnual:35,  vigenteDesde:"2026-04-01", fuenteNombre:"BADLAR Bancos Privados (BCRA)", fuenteUrl:"https://www.bcra.gob.ar/", status:"fallback"},
    tamar:      {tasaAnual:40,  vigenteDesde:"2026-04-01", fuenteNombre:"TAMAR Bancos Privados (BCRA)", fuenteUrl:"https://www.bcra.gob.ar/", status:"fallback"},
    // Acta 2783 / Plenario Barrios: tasa pura del 6% sumada a actualización CER del BCRA.
    // El factor CER se aplica luego en el cliente (al-calc-bcra.js) usando el dato CER vigente.
    acta2783:   {tasaAnual:6,   vigenteDesde:"2026-04-01", fuenteNombre:"CNAT Acta 2783 — Tasa pura 6% sobre CER (BCRA id=30)", fuenteUrl:"https://www.bcra.gob.ar/", status:"fallback", nota:"Aplicar factor CER + 6% anual"},
    plenarioBarrios: {tasaAnual:6, vigenteDesde:"2026-04-01", fuenteNombre:"Plenario Barrios SCBA — Tasa pura 6% sobre CER", fuenteUrl:"https://www.bcra.gob.ar/", status:"fallback", nota:"Aplicar factor CER + 6% anual (similar a Acta 2783)"}
  }
};

const UA = {'User-Agent':'Mozilla/5.0 (LitigiAR-bot/1.0; +https://argentinalitigia.com)'};

function parsearNumArg(txt){
  if(!txt) return null;
  const m = String(txt).replace(/\$/g,'').replace(/\s/g,'').replace(/\./g,'').replace(',', '.');
  const n = Number(m);
  return isFinite(n) ? n : null;
}

function loadPrevious(){
  try {
    if (fs.existsSync(VALORES_PATH)) return JSON.parse(fs.readFileSync(VALORES_PATH,'utf8'));
  } catch(e){ console.warn('[LOAD] no se pudo leer valores.json previo:', e.message); }
  return null;
}

function logStatus(seccion, obj){
  console.log(`[VALORES] ${seccion}:`, Object.keys(obj).map(k => {
    const v = obj[k]; return `${k}=${v.status||'?'}`;
  }).join(' | '));
}

// ═══════ datos.gob.ar ═══════
async function fetchSerieGob(url){
  try {
    const r = await axios.get(url, {timeout:20000, headers:UA});
    if (!r.data || !r.data.data || !r.data.data.length) return null;
    return r.data.data[0];
  } catch (e) { return null; }
}

async function fetchIPC(){
  const rVar = await fetchSerieGob('https://apis.datos.gob.ar/series/api/series/?ids=148.3_INIVELNAL_DICI_M_26:percent_change&sort=desc&limit=1&metadata=none');
  if (!rVar) return null;
  const periodo = String(rVar[0]).substring(0,7);
  const raw = Number(rVar[1]);
  // datos.gob.ar :percent_change devuelve proporción — multiplicar x100
  const mensual = Number((Math.abs(raw) < 1 ? raw*100 : raw).toFixed(2));
  let interanual=null, acumulada=null;
  try {
    const r2 = await axios.get('https://apis.datos.gob.ar/series/api/series/?ids=148.3_INIVELNAL_DICI_M_26&sort=desc&limit=13&metadata=none', {timeout:20000, headers:UA});
    if (r2.data && r2.data.data && r2.data.data.length >= 13) {
      const actual = r2.data.data[0][1];
      const hace12 = r2.data.data[12][1];
      interanual = Number((((actual/hace12)-1)*100).toFixed(2));
      const anio = periodo.substring(0,4);
      const dicAnt = r2.data.data.find(r => String(r[0]).substring(0,7) === (parseInt(anio)-1)+"-12");
      if (dicAnt) acumulada = Number((((actual/dicAnt[1])-1)*100).toFixed(2));
    }
  } catch(e){}
  return {periodo, mensual, interanual, acumulada, fuenteNombre:"INDEC / datos.gob.ar", fuenteUrl:"https://apis.datos.gob.ar/series/api/series/?ids=148.3_INIVELNAL_DICI_M_26", status:"ok", fechaConsulta:new Date().toISOString()};
}

// Mapa de series datos.gob.ar → idVariable BCRA v4 (fuente más actual)
// datos.gob.ar suele atrasarse 4-8 semanas; BCRA v4 publica diario.
const BCRA_V4_FALLBACK = {
  '94.2_UVAD_D_0_0_10': { id: 31, nombre: 'UVA' },        // BCRA v4 id 31 = UVA
  '94.2_CD_D_0_0_10':   { id: 30, nombre: 'CER' },        // BCRA v4 id 30 = CER
  '94.2_UVID_D_0_0_10': { id: 32, nombre: 'UVI' }         // BCRA v4 id 32 = UVI
};

async function fetchBCRAdiario(id){
  // 1) Intentar BCRA v4 directo (más actualizado)
  const fb = BCRA_V4_FALLBACK[id];
  if (fb) {
    try {
      const https = require('https');
      const hoy = new Date();
      const desde = new Date(hoy.getTime() - 7*86400000).toISOString().substring(0,10);
      const hasta = hoy.toISOString().substring(0,10);
      const r = await axios.get(`https://api.bcra.gob.ar/estadisticas/v4.0/monetarias/${fb.id}?desde=${desde}&hasta=${hasta}`, {
        timeout: 20000,
        headers: UA,
        httpsAgent: new https.Agent({ rejectUnauthorized: false })
      });
      const det = r.data && r.data.results && r.data.results[0] && r.data.results[0].detalle;
      if (det && det.length) {
        // El primer elemento es el más reciente
        const ultimo = det[0];
        return {
          fecha: ultimo.fecha,
          valor: Number(Number(ultimo.valor).toFixed(4)),
          fuenteNombre: `BCRA v4 (${fb.nombre} id=${fb.id})`,
          fuenteUrl: `https://api.bcra.gob.ar/estadisticas/v4.0/monetarias/${fb.id}`,
          status: 'ok',
          fechaConsulta: new Date().toISOString()
        };
      }
    } catch(e) {
      console.warn(`[VALORES] BCRA v4 ${fb.nombre} fail:`, e.message);
    }
  }
  // 2) Fallback a datos.gob.ar
  const r = await fetchSerieGob('https://apis.datos.gob.ar/series/api/series/?ids='+id+'&sort=desc&limit=1&metadata=none');
  if (!r) return null;
  return {
    fecha: String(r[0]).substring(0,10),
    valor: Number(Number(r[1]).toFixed(4)),
    fuenteNombre: 'BCRA / datos.gob.ar (puede tener atraso)',
    fuenteUrl: 'https://apis.datos.gob.ar/series/api/series/?ids='+id,
    status: 'ok',
    fechaConsulta: new Date().toISOString()
  };
}

async function fetchSMVM(){
  const r = await fetchSerieGob('https://apis.datos.gob.ar/series/api/series/?ids=57.1_SMVMM_0_M_34&sort=desc&limit=1&metadata=none');
  if (!r) return null;
  return {
    periodo: String(r[0]).substring(0,7),
    valor: Math.round(Number(r[1])),
    fuenteNombre: "MTEySS / datos.gob.ar",
    fuenteUrl: "https://apis.datos.gob.ar/series/api/series/?ids=57.1_SMVMM_0_M_34",
    status: "ok",
    fechaConsulta: new Date().toISOString()
  };
}

async function fetchRIPTE(){
  const r = await fetchSerieGob('https://apis.datos.gob.ar/series/api/series/?ids=158.1_REPTE_0_0_5&sort=desc&limit=1&metadata=none');
  if (!r) return null;
  return {
    periodo: String(r[0]).substring(0,7),
    valor: Number(Number(r[1]).toFixed(2)),
    fuenteNombre: "MTSS / datos.gob.ar",
    fuenteUrl: "https://apis.datos.gob.ar/series/api/series/?ids=158.1_REPTE_0_0_5",
    status: "ok",
    fechaConsulta: new Date().toISOString()
  };
}

// ═══════ BCRA XLS ICL ═══════
async function scrapearICLBCRA(){
  try {
    const url = 'https://www.bcra.gob.ar/Pdfs/PublicacionesEstadisticas/diar_icl.xls';
    const r = await axios.get(url, {responseType:'arraybuffer', timeout:30000, headers:UA});
    const wb = XLSX.read(r.data, {type:'buffer'});
    const sh = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sh, {header:1});
    let ultima = null;
    for (let i = rows.length - 1; i >= 0; i--) {
      const row = rows[i];
      if (row && row[0] && typeof row[1] === 'number' && row[1] > 0) { ultima = row; break; }
    }
    if (!ultima) return null;
    let fecha = ultima[0];
    if (typeof fecha === 'number') {
      const date = new Date(Math.round((fecha - 25569) * 86400 * 1000));
      fecha = date.toISOString().substring(0,10);
    } else {
      fecha = String(fecha).substring(0,10);
    }
    return {
      fecha,
      valor: Number(Number(ultima[1]).toFixed(4)),
      fuenteNombre: 'BCRA diar_icl.xls',
      fuenteUrl: url,
      status: 'ok',
      fechaConsulta: new Date().toISOString()
    };
  } catch (e) { console.warn('[VALORES] ICL BCRA XLS fail:', e.message); return null; }
}

// ═══════ INDEC Canasta Crianza ═══════
async function scrapearCanastaCrianza(){
  try {
    const url = 'https://www.indec.gob.ar/ftp/cuadros/sociedad/serie_canasta_crianza.xlsx';
    const r = await axios.get(url, {responseType:'arraybuffer', timeout:30000, headers:UA});
    const wb = XLSX.read(r.data, {type:'buffer'});
    const sh = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sh, {header:1});
    let ultima = null;
    for (let i = rows.length - 1; i >= 0; i--) {
      const row = rows[i];
      if (row && row[0] && typeof row[1] === 'number' && row[1] > 0) { ultima = row; break; }
    }
    if (!ultima) return null;
    return {
      periodo: String(ultima[0]).substring(0,7),
      menor1:  Math.round(ultima[1]),
      edad1a3: Math.round(ultima[2]),
      edad4a5: Math.round(ultima[3]),
      edad6a12:Math.round(ultima[4]),
      fuenteNombre: 'INDEC Canasta de Crianza',
      fuenteUrl: url,
      status: 'ok',
      fechaConsulta: new Date().toISOString()
    };
  } catch (e) { console.warn('[VALORES] Canasta Crianza fail:', e.message); return null; }
}

// ═══════ PBA: La Matanza + Morón + ColProBA ═══════
function scrapearHTMLcolegio(url, label){
  return axios.get(url, {timeout:20000, headers:UA}).then(({data}) => {
    const texto = cheerio.load(data)('body').text().replace(/\s+/g,' ');
    const cercano = (keyword) => {
      const patterns = [
        new RegExp('\\$\\s*([\\d.,]+)[^$\\n]{0,40}?' + keyword, 'i'),
        new RegExp(keyword + '[^$\\n]{0,40}?\\$\\s*([\\d.,]+)', 'i')
      ];
      for (const re of patterns) { const m = texto.match(re); if (m) return parsearNumArg(m[1]); }
      return null;
    };
    return {
      jusLey14967:    cercano('JUS\\s*(?:Arancelario|Ley|14\\.?967)'),
      jusDL8904:      cercano('(?:DL\\s*8904|8904\\/77|8\\.?904)'),
      jusPrevisional: cercano('JUS\\s*Previsional'),
      bonoVerde8480:  cercano('Bono[^$]*?Verde'),
      bonoAzul8480:   cercano('Bono[^$]*?Azul')
    };
  }).catch(err => {
    console.warn(`[VALORES] ${label} fail:`, err.message);
    return null;
  });
}

async function refreshPBA(prev){
  const [lm, mor, cp] = await Promise.all([
    scrapearHTMLcolegio('https://colegioabogadoslamatanza.com.ar/', 'La Matanza'),
    scrapearHTMLcolegio('https://camoron.org.ar/', 'Morón'),
    scrapearHTMLcolegio('https://colproba.org.ar/', 'ColProBA')
  ]);
  const out = {};
  const claves = Object.keys(FALLBACK.pba);
  const fuenteElegida = (k) => {
    if (lm && lm[k]) return {valor: lm[k], fuenteNombre:'Colegio La Matanza'};
    if (mor && mor[k]) return {valor: mor[k], fuenteNombre:'Colegio Morón'};
    if (cp && cp[k]) return {valor: cp[k], fuenteNombre:'ColProBA'};
    return null;
  };
  for (const k of claves){
    const res = fuenteElegida(k);
    if (res && typeof res.valor === 'number' && res.valor > 0) {
      out[k] = {
        valor: res.valor,
        fuenteNombre: res.fuenteNombre,
        fuenteUrl: FALLBACK.pba[k].fuenteUrl,
        status: 'ok',
        fechaConsulta: new Date().toISOString()
      };
    } else if (prev && prev.pba && prev.pba[k] && prev.pba[k].status === 'ok') {
      // Mantener último valor válido con status stale
      out[k] = Object.assign({}, prev.pba[k], {status:'stale', nota:'Se mantiene último valor válido — scraper no respondió'});
    } else {
      out[k] = Object.assign({}, FALLBACK.pba[k], {status:'fallback'});
    }
  }
  return out;
}

// ═══════ IPC CABA / GBA (IDECBA) ═══════
async function scrapearIDECBA(){
  try {
    // IDECBA publica tabla en su portal. Buscamos la serie más reciente.
    const url = 'https://www.estadisticaciudad.gob.ar/eyc/?p=indice_de_precios_al_consumidor';
    const { data } = await axios.get(url, { timeout: 20000, headers: UA });
    const texto = cheerio.load(data)('body').text().replace(/\s+/g, ' ');
    // Patrón: busca "variación mensual" + un porcentaje cerca
    const matMensual = texto.match(/variaci[oó]n\s+mensual[^%]{0,80}?(\d+[,.]\d+)\s*%/i);
    const matInteranual = texto.match(/interanual[^%]{0,80}?(\d+[,.]\d+)\s*%/i);
    const matAcumulada = texto.match(/acumulad[ao][^%]{0,80}?(\d+[,.]\d+)\s*%/i);
    const parseNum = (s) => s ? parseFloat(String(s[1]).replace(',', '.')) : null;
    const mensual = parseNum(matMensual);
    if (!mensual) return null;
    // Período: generalmente el mes anterior al actual
    const ahora = new Date();
    ahora.setMonth(ahora.getMonth() - 1);
    const periodo = ahora.toISOString().substring(0, 7);
    return {
      periodo,
      mensual,
      interanual: parseNum(matInteranual),
      acumulada: parseNum(matAcumulada),
      fuenteNombre: 'IDECBA Dir. Estadística CABA',
      fuenteUrl: url,
      status: 'ok',
      fechaConsulta: new Date().toISOString()
    };
  } catch (e) {
    console.warn('[VALORES] IDECBA fail:', e.message);
    return null;
  }
}

// ═══════ Tasas CNAT / BNA / BPBA (scraping multi-fuente) ═══════
// Estrategia: probamos 3 fuentes que publican tasas actualizadas:
//   1) CNAT oficial (tasas-de-interes)
//   2) Enlaces Jurídicos (tablas agregadas, públicas)
//   3) TodosXDerecho (alternativa)
// Extraemos con regex flexible cada valor por Acta.
async function scrapearTasasPublicadas(){
  var fuentes = [
    // CNAT / PJN oficial (pueden responder desde IPs US de GitHub Actions)
    'https://old.pjn.gov.ar/04_cnat/index.asp',
    'https://www.pjn.gov.ar/',
    // Enlaces Jurídicos (aggregator público)
    'https://www.enlacesjuridicos.com.ar/',
    'https://www.enlacesjuridicos.com.ar/index',
    // Colegios profesionales que publican tasas (redundancia)
    'https://camoron.org.ar/',
    'https://colproba.org.ar/',
    // BNA directo (puede tener bot-block)
    'https://www.bna.com.ar/Personas/TasasdeInteres',
    // Google-cached fallback
    'https://www.google.com/search?q=tasa+acta+cnat+2601+vigente+bna'
  ];
  var tasas = {};
  var fuenteExitosa = null;

  for (var i = 0; i < fuentes.length; i++) {
    var url = fuentes[i];
    try {
      var res = await axios.get(url, { timeout: 15000, headers: UA });
      var $ = cheerio.load(res.data);
      var texto = $('body').text().replace(/\s+/g, ' ');

      // Patrones flexibles — buscan "Acta NNNN" cercana a un porcentaje
      var patrones = {
        acta2601: /Acta\s*2601[^%]{0,120}?(\d{1,4}[,.]\d{1,4})\s*%/i,
        acta2630: /Acta\s*2630[^%]{0,120}?(\d{1,4}[,.]\d{1,4})\s*%/i,
        acta2658: /Acta\s*2658[^%]{0,120}?(\d{1,4}[,.]\d{1,4})\s*%/i,
        acta2764: /Acta\s*2764[^%]{0,120}?(\d{1,4}[,.]\d{1,4})\s*%/i,
        acta2783: /Acta\s*2783[^%]{0,120}?(\d{1,4}[,.]\d{1,4})\s*%/i,
        // BNA: activa / pasiva / libre 36m / libre 72m
        bnaActiva:  /BNA\s*(?:Tasa\s*)?Activa[^%]{0,80}?(\d{1,4}[,.]\d{1,4})\s*%/i,
        bnaPasiva:  /BNA\s*(?:Tasa\s*)?Pasiva[^%]{0,80}?(\d{1,4}[,.]\d{1,4})\s*%/i,
        bnaLibre36: /(?:BNA\s*)?Libre\s*(?:destino\s*)?36\s*meses[^%]{0,80}?(\d{1,4}[,.]\d{1,4})\s*%/i,
        bnaLibre72: /(?:BNA\s*)?Libre\s*(?:destino\s*)?72\s*meses[^%]{0,80}?(\d{1,4}[,.]\d{1,4})\s*%/i,
        // BPBA
        bpActiva:   /(?:BP|Banco\s*Provincia)[^%]{0,30}?Activa[^%]{0,80}?(\d{1,4}[,.]\d{1,4})\s*%/i,
        bpPasiva:   /(?:BP|Banco\s*Provincia)[^%]{0,30}?Pasiva[^%]{0,80}?(\d{1,4}[,.]\d{1,4})\s*%/i
      };

      Object.keys(patrones).forEach(function(key){
        if (tasas[key]) return; // ya lo tenemos de otra fuente
        var m = texto.match(patrones[key]);
        if (m) {
          var val = parseFloat(String(m[1]).replace(',', '.'));
          // Filtro razonable: tasa entre 1% y 500% anual
          if (isFinite(val) && val >= 1 && val <= 500) {
            tasas[key] = {
              tasaAnual: val,
              fuenteNombre: url.replace(/^https?:\/\/(www\.)?/, '').split('/')[0],
              fuenteUrl: url,
              status: 'ok',
              fechaConsulta: new Date().toISOString()
            };
            if (!fuenteExitosa) fuenteExitosa = url;
          }
        }
      });

      // Si ya tenemos las 3 principales, salimos
      if (tasas.acta2601 && tasas.acta2630 && tasas.acta2658) break;
    } catch(e){
      console.warn('[VALORES] Tasas fuente '+url+' fail:', e.message);
    }
  }

  if (Object.keys(tasas).length === 0) return null;
  console.log('[VALORES] Tasas OK: '+Object.keys(tasas).join(', ')+' (fuente: '+fuenteExitosa+')');
  return tasas;
}

// ═══════ BCRA API v4.0 — tasas oficiales diarias ═══════
// Variables relevantes del BCRA (1220 total):
//   7  = BADLAR nominal · 12 = Depósitos 30d (BNA Pasiva aprox)
//   13 = Adelantos cta cte · 14 = Préstamos personales (BNA Activa aprox)
//   35 = BADLAR efectiva · 43 = Com. P 14.290 (Uso de Justicia)
//   44 = TAMAR nominal · 45 = TAMAR efectiva · 160 = Política monetaria
async function fetchBCRAv4(){
  try {
    const https = require('https');
    const r = await axios.get('https://api.bcra.gob.ar/estadisticas/v4.0/monetarias', {
      timeout: 25000,
      headers: UA,
      httpsAgent: new https.Agent({ rejectUnauthorized: false })
    });
    if (!r.data || !r.data.results) return null;
    const mapa = {};
    r.data.results.forEach(v => {
      mapa[v.idVariable] = {
        valor: Number(v.ultValorInformado),
        fecha: v.ultFechaInformada,
        descripcion: (v.descripcion || '').trim(),
        unidad: v.unidadExpresion
      };
    });
    return mapa;
  } catch(e) {
    console.warn('[VALORES] BCRA v4 fail:', e.message);
    return null;
  }
}

// Mapeo de nuestras tasas → IDs variables BCRA
//
// IMPORTANTE: el BCRA NO publica como serie individual la "Tasa Activa Cartera
// General BNA 30 días" ni las TEAs específicas del BNA (Libre Destino 36/72m,
// préstamo personal 49-60m, etc.). Esas son tasas DE UN BANCO concreto que solo
// se obtienen del sitio bna.com.ar (vía Playwright) o de la planilla mensual CNAT.
//
// Por eso ACÁ solo mapeamos a BCRA las tasas que SÍ son sistema-promedio o
// productos publicados por BCRA. Las tasas BNA específicas (Acta 2601/2630/2658,
// bnaActiva, bnaLibre36, bnaLibre72) NO están en este mapeo — se obtienen
// exclusivamente desde scraperPlaywright(BNA) o, si falla, quedan en fallback.
const MAPEO_TASAS = {
  // Pasiva (depósitos 30d) — id=12 — sí es la fuente correcta para tasa pasiva
  bnaPasiva:  { idVar: 12, fuenteNombre: "BCRA v4 — Tasa de depósitos 30 días (Pasiva sistema)" },
  bpPasiva:   { idVar: 12, fuenteNombre: "BCRA v4 — Tasa de depósitos 30 días (Pasiva sistema)" },
  // Activa BPBA — usamos BADLAR (id=7) como aproximación más razonable que
  // adelantos en cuenta corriente (id=13). De todos modos lo ideal es scraper directo.
  bpActiva:   { idVar: 7,  fuenteNombre: "BCRA v4 — BADLAR (aprox. Activa BPBA cuando no hay scrape directo)" },
  // Series sistema — sí están en BCRA
  badlar:     { idVar: 7,  fuenteNombre: "BCRA v4 — BADLAR bancos privados (nominal anual)" },
  tamar:      { idVar: 44, fuenteNombre: "BCRA v4 — TAMAR bancos privados (nominal anual)" }
};

// Tasas que NUNCA deben venir del BCRA sistema (son específicas de un banco
// concreto). Si el scraper Playwright BNA falla, quedan en status='fallback'
// con el último valor manual conocido y nota explícita al usuario.
const TASAS_SOLO_BNA = ["acta2601","acta2630","acta2658","bnaActiva","bnaLibre36","bnaLibre72"];

// Devuelve tasas scrapeadas en formato listo para mergear
function aplicarBCRAaTasas(bcraMap){
  if (!bcraMap) return null;
  const out = {};
  Object.keys(MAPEO_TASAS).forEach(k => {
    const m = MAPEO_TASAS[k];
    const v = bcraMap[m.idVar];
    if (v && isFinite(v.valor) && v.valor > 0 && v.valor < 500) {
      out[k] = {
        tasaAnual: Number(v.valor.toFixed(2)),
        vigenteDesde: v.fecha,
        fuenteNombre: m.fuenteNombre,
        fuenteUrl: 'https://api.bcra.gob.ar/estadisticas/v4.0/monetarias',
        status: 'ok',
        fechaConsulta: new Date().toISOString()
      };
    }
  });
  return Object.keys(out).length > 0 ? out : null;
}

// ═══════════════════════════════════════════════════════════
//  PLAYWRIGHT — scraping de sitios JS-rendered (BNA, BPBA, CNAT)
// ═══════════════════════════════════════════════════════════
const PLAYWRIGHT_ENABLED = process.env.PLAYWRIGHT_ENABLED === '1';

async function getBrowser(){
  if (!PLAYWRIGHT_ENABLED) return null;
  try {
    const { chromium } = require('playwright');
    return await chromium.launch({ headless: true, args: ['--no-sandbox', '--disable-dev-shm-usage'] });
  } catch(e) {
    console.warn('[VALORES] Playwright launch fail:', e.message);
    return null;
  }
}

async function extraerTasasBNA(browser){
  if (!browser) return null;
  const page = await browser.newPage({ userAgent: UA['User-Agent'] });
  // Lista de URLs candidatas (BNA cambia rutas con frecuencia)
  const candidatos = [
    'https://www.bna.com.ar/Personas/PrestamosVigentes',
    'https://www.bna.com.ar/Personas/Prestamos',
    'https://www.bna.com.ar/Personas/TasasdeInteres',
    'https://www.bna.com.ar/Personas',
    'https://www.bna.com.ar/'
  ];
  let resultado = { activa: null, pasiva: null, libre36: null, libre72: null, fuente: null };
  try {
    for (const url of candidatos) {
      try {
        await page.goto(url, { waitUntil: 'networkidle', timeout: 30000 });
        const texto = await page.textContent('body');
        if (!texto || texto.length < 200) continue;
        const match = (regex) => {
          const m = texto.match(regex);
          return m ? parseFloat(String(m[1]).replace(',', '.')) : null;
        };
        // Tasa Activa Cartera General BNA 30 días — la del Acta 2601 CNAT
        if (!resultado.activa) {
          resultado.activa = match(/Cartera\s*General[^%]{0,400}?(?:TNA|TEA|tasa\s*nominal|tasa\s*efectiva)[^%]{0,100}?(\d{1,3}[,.]\d{1,4})\s*%/i)
                          || match(/Activa\s*Cartera\s*General[^%]{0,400}?(\d{1,3}[,.]\d{1,4})\s*%/i)
                          || match(/(?:TNA|TEA)[^%]{0,40}?Cartera[^%]{0,80}?(\d{1,3}[,.]\d{1,4})\s*%/i);
        }
        // Préstamos personales libre destino 36/72 meses — Actas 2658/2764
        if (!resultado.libre36) {
          resultado.libre36 = match(/Libre\s*Destino[^%]{0,300}?36\s*(?:meses|cuotas)[^%]{0,200}?(?:TNA|TEA)[^%]{0,80}?(\d{1,3}[,.]\d{1,4})\s*%/i)
                          || match(/36\s*(?:meses|cuotas)[^%]{0,400}?(?:TNA|TEA|tasa)[^%]{0,100}?(\d{1,3}[,.]\d{1,4})\s*%/i);
        }
        if (!resultado.libre72) {
          resultado.libre72 = match(/Libre\s*Destino[^%]{0,300}?72\s*(?:meses|cuotas)[^%]{0,200}?(?:TNA|TEA)[^%]{0,80}?(\d{1,3}[,.]\d{1,4})\s*%/i)
                          || match(/72\s*(?:meses|cuotas)[^%]{0,400}?(?:TNA|TEA|tasa)[^%]{0,100}?(\d{1,3}[,.]\d{1,4})\s*%/i);
        }
        if (!resultado.pasiva) {
          resultado.pasiva = match(/(?:plazo\s*fijo\s*tradicional|depósito\s*a\s*plazo)[^%]{0,300}?(?:TNA|tasa)[^%]{0,80}?(\d{1,3}[,.]\d{1,4})\s*%/i)
                          || match(/(?:plazo\s*fijo|depósito)[^%]{0,300}?(\d{1,3}[,.]\d{1,4})\s*%/i);
        }
        if (resultado.activa || resultado.libre36 || resultado.libre72) {
          resultado.fuente = url;
          // Si tenemos al menos las 3 principales, salimos
          if (resultado.activa && resultado.libre36 && resultado.libre72) break;
        }
      } catch(e) {
        console.warn('[VALORES] BNA candidato '+url+' fail:', e.message);
      }
    }
    if (resultado.activa || resultado.libre36 || resultado.libre72 || resultado.pasiva) {
      console.log('[VALORES] BNA Playwright OK: activa='+resultado.activa+' L36='+resultado.libre36+' L72='+resultado.libre72+' pasiva='+resultado.pasiva+' (fuente: '+resultado.fuente+')');
      return resultado;
    }
    console.warn('[VALORES] BNA Playwright: ninguna URL devolvió tasas extraibles');
    return null;
  } finally { await page.close().catch(()=>{}); }
}

async function extraerTasasBPBA(browser){
  if (!browser) return null;
  const page = await browser.newPage({ userAgent: UA['User-Agent'] });
  try {
    await page.goto('https://www.bancoprovincia.com.ar/cuentas_personales/tasas', { waitUntil: 'networkidle', timeout: 40000 });
    const texto = await page.textContent('body');
    const match = (regex) => {
      const m = texto.match(regex);
      return m ? parseFloat(String(m[1]).replace(',', '.')) : null;
    };
    const activa = match(/Cartera\s*General[^%]{0,400}?(\d{1,3}[,.]\d{1,4})\s*%/i);
    const pasiva = match(/(?:plazo\s*fijo|depósito)[^%]{0,300}?(\d{1,3}[,.]\d{1,4})\s*%/i);
    return { activa, pasiva, fuente: 'https://www.bancoprovincia.com.ar/' };
  } catch(e) {
    console.warn('[VALORES] BPBA Playwright fail:', e.message);
    return null;
  } finally { await page.close().catch(()=>{}); }
}

async function extraerTasasCNAT(browser){
  if (!browser) return null;
  const page = await browser.newPage({ userAgent: UA['User-Agent'] });
  try {
    // CNAT publica Planilla de Actualización en diversos URLs de PJN
    const candidatos = [
      'https://www.pjn.gov.ar/',
      'https://www.cnat.gob.ar/'
    ];
    const actas = {};
    for (const url of candidatos) {
      try {
        await page.goto(url, { waitUntil: 'networkidle', timeout: 30000 });
        const texto = await page.textContent('body');
        ['2601','2630','2658','2764','2783'].forEach(nro => {
          if (actas['acta'+nro]) return;
          const re = new RegExp('Acta\\s*'+nro+'[^%]{0,150}?(\\d{1,3}[,.]\\d{1,4})\\s*%', 'i');
          const m = texto.match(re);
          if (m) {
            const val = parseFloat(String(m[1]).replace(',', '.'));
            if (isFinite(val) && val > 1 && val < 500) {
              actas['acta'+nro] = val;
            }
          }
        });
        if (Object.keys(actas).length >= 3) break;
      } catch(e){}
    }
    return Object.keys(actas).length ? Object.assign(actas, {fuente: 'CNAT/PJN'}) : null;
  } catch(e) {
    console.warn('[VALORES] CNAT Playwright fail:', e.message);
    return null;
  } finally { await page.close().catch(()=>{}); }
}

// Cruza BNA + BCRA v4 + colegios. Si ≥2 fuentes coinciden (±5%), status=ok consensus.
function consensuarTasa(fuentes, etiquetaKey){
  const vals = fuentes.filter(v => v != null && isFinite(v) && v > 0).map(Number);
  if (vals.length === 0) return null;
  if (vals.length === 1) return { valor: vals[0], consenso: 1 };
  // Chequeo de consenso: si los valores están dentro del 10% entre sí, promediamos
  const min = Math.min(...vals), max = Math.max(...vals);
  const margen = max === 0 ? 0 : (max - min) / max;
  if (margen <= 0.10) {
    const prom = vals.reduce((s,x)=>s+x,0) / vals.length;
    return { valor: Number(prom.toFixed(2)), consenso: vals.length };
  }
  // Sin consenso → devolver el primer valor (BCRA suele ser el primero por prioridad)
  return { valor: vals[0], consenso: 1, disenso: vals };
}

// ═══════ Orquestador ═══════
async function main(){
  console.log('[VALORES] Iniciando scrape…', new Date().toISOString());
  const prev = loadPrevious();

  const [ipc, uva, cer, uvi, icl, ripte, smvm, canasta, pba, ipcGba, badlar, tasasPub] = await Promise.all([
    fetchIPC().catch(()=>null),
    fetchBCRAdiario('94.2_UVAD_D_0_0_10').catch(()=>null),
    fetchBCRAdiario('94.2_CD_D_0_0_10').catch(()=>null),
    fetchBCRAdiario('94.2_UVID_D_0_0_10').catch(()=>null),
    scrapearICLBCRA().catch(()=>null),
    fetchRIPTE().catch(()=>null),
    fetchSMVM().catch(()=>null),
    scrapearCanastaCrianza().catch(()=>null),
    refreshPBA(prev),
    scrapearIDECBA().catch(()=>null),
    fetchBCRAv4().catch(()=>null),
    scrapearTasasPublicadas().catch(()=>null)
  ]);
  // Transformamos el map BCRA v4 a nuestro formato de tasas
  const tasasBCRA = aplicarBCRAaTasas(badlar); // 'badlar' ahora contiene mapa completo BCRA v4

  // ── Playwright: scraping JS-rendered real (BNA / BPBA / CNAT) ──
  const browser = await getBrowser();
  let bnaReal = null, bpbaReal = null, cnatReal = null;
  if (browser) {
    [bnaReal, bpbaReal, cnatReal] = await Promise.all([
      extraerTasasBNA(browser).catch(()=>null),
      extraerTasasBPBA(browser).catch(()=>null),
      extraerTasasCNAT(browser).catch(()=>null)
    ]);
    await browser.close().catch(()=>{});
    console.log('[VALORES] Playwright: BNA='+(bnaReal?'ok':'fail')+' BPBA='+(bpbaReal?'ok':'fail')+' CNAT='+(cnatReal?'ok':'fail'));
  } else {
    console.log('[VALORES] Playwright deshabilitado (PLAYWRIGHT_ENABLED!=1) — usando solo BCRA v4');
  }

  // ── Construcción de tasasFinales ────────────────────────────────────────
  // Reglas:
  //  · Tasas BNA específicas (Acta 2601/2630/2658, bnaActiva, bnaLibre36/72)
  //    → SOLO si Playwright BNA o CNAT scrape devolvieron valor real.
  //      No usar BCRA v4 sistema-promedio como sustituto (id=14 Préstamos
  //      personales NO es la TACG ni la TEA libre destino BNA).
  //  · Tasas pasivas (bnaPasiva, bpPasiva) → BCRA v4 id=12 es válido (es la
  //    tasa promedio sistema, comparable a la BNA pasiva).
  //  · BADLAR / TAMAR → BCRA v4 directo (mismo origen).
  //  · bpActiva → si Playwright BPBA falla, BADLAR como aproximación con
  //    nota explícita.
  const tasasFinales = {};
  const TODAS_LAS_TASAS = Object.keys(FALLBACK.tasas);

  TODAS_LAS_TASAS.forEach(k => {
    const fuentes = [];
    let notaConsenso = '';
    let fuenteUrlFinal = null;
    let fuenteNombreFinal = null;
    let vigenteDesde = null;

    // 1) BCRA v4 — solo si está en MAPEO_TASAS (= NO es tasa BNA específica)
    if (tasasBCRA && tasasBCRA[k]) {
      fuentes.push(tasasBCRA[k].tasaAnual);
      notaConsenso = MAPEO_TASAS[k] ? MAPEO_TASAS[k].fuenteNombre : tasasBCRA[k].fuenteNombre;
      fuenteUrlFinal = tasasBCRA[k].fuenteUrl;
      fuenteNombreFinal = notaConsenso;
      vigenteDesde = tasasBCRA[k].vigenteDesde;
    }

    // 2) Playwright BNA (PRIMARIO para tasas BNA específicas)
    if (bnaReal) {
      const mapBNA = {
        acta2601: bnaReal.activa, acta2630: bnaReal.activa, acta2658: bnaReal.libre36,
        bnaActiva: bnaReal.activa, bnaPasiva: bnaReal.pasiva,
        bnaLibre36: bnaReal.libre36, bnaLibre72: bnaReal.libre72
      };
      const valBNA = mapBNA[k];
      if (valBNA && isFinite(valBNA) && valBNA > 0 && valBNA < 500) {
        fuentes.push(valBNA);
        if (TASAS_SOLO_BNA.includes(k)) {
          // Esta tasa solo viene confiable de BNA — sobrescribimos fuente
          fuenteNombreFinal = 'BNA oficial (scrape directo www.bna.com.ar)';
          fuenteUrlFinal = bnaReal.fuente || 'https://www.bna.com.ar/';
          vigenteDesde = new Date().toISOString().substring(0,10);
        } else {
          notaConsenso += ' + BNA';
        }
      }
    }

    // 3) Playwright BPBA
    if (bpbaReal) {
      const mapBPBA = { bpActiva: bpbaReal.activa, bpPasiva: bpbaReal.pasiva };
      const v = mapBPBA[k];
      if (v && isFinite(v) && v > 0 && v < 500) {
        fuentes.push(v);
        if (k === 'bpActiva') {
          fuenteNombreFinal = 'BPBA oficial (scrape directo bancoprovincia.com.ar)';
          fuenteUrlFinal = bpbaReal.fuente || 'https://www.bancoprovincia.com.ar/';
          vigenteDesde = new Date().toISOString().substring(0,10);
        } else {
          notaConsenso += ' + BPBA';
        }
      }
    }

    // 4) Scrape HTML CNAT/PJN — para Actas
    if (cnatReal && cnatReal[k] && isFinite(cnatReal[k])) {
      fuentes.push(cnatReal[k]);
      if (TASAS_SOLO_BNA.includes(k)) {
        fuenteNombreFinal = 'CNAT/PJN (planilla oficial)';
        fuenteUrlFinal = 'https://www.cnat.gob.ar/';
        vigenteDesde = new Date().toISOString().substring(0,10);
      } else {
        notaConsenso += ' + CNAT';
      }
    }

    // 5) tasasPub (scrapearTasasPublicadas — regex sobre HTML público)
    if (tasasPub && tasasPub[k] && tasasPub[k].tasaAnual) {
      fuentes.push(tasasPub[k].tasaAnual);
      if (TASAS_SOLO_BNA.includes(k) && !fuenteNombreFinal) {
        fuenteNombreFinal = tasasPub[k].fuenteNombre;
        fuenteUrlFinal = tasasPub[k].fuenteUrl;
      } else {
        notaConsenso += ' + scrape';
      }
    }

    // Decisión final
    if (fuentes.length > 0) {
      const c = consensuarTasa(fuentes, k);
      if (c) {
        tasasFinales[k] = {
          tasaAnual: c.valor,
          vigenteDesde: vigenteDesde || new Date().toISOString().substring(0,10),
          fuenteNombre: fuenteNombreFinal || notaConsenso,
          fuenteUrl: fuenteUrlFinal || (tasasBCRA && tasasBCRA[k] ? tasasBCRA[k].fuenteUrl : null),
          status: 'ok',
          fechaConsulta: new Date().toISOString(),
          consenso: c.consenso,
          fuentes: fuentes.length > 1 ? fuentes : undefined,
          disenso: c.disenso
        };
      }
    }
  });

  const reemplazarBCRA = tasasFinales;

  // Helper para elegir entre nuevo / stale / fallback
  const elegir = (seccion, campo, nuevo, fallback) => {
    if (nuevo) return nuevo;
    if (prev && prev[seccion] && prev[seccion][campo] && prev[seccion][campo].status === 'ok') {
      return Object.assign({}, prev[seccion][campo], {status:'stale', nota:'Se mantiene último valor válido'});
    }
    if (prev && seccion === campo && prev[seccion] && prev[seccion].status === 'ok') {
      return Object.assign({}, prev[seccion], {status:'stale', nota:'Se mantiene último valor válido'});
    }
    return Object.assign({}, fallback, {status:'fallback'});
  };

  // Merge tasas: prioridad → 1) BCRA v4 + Playwright (consenso) · 2) scrape HTML · 3) prev ok/manual · 4) fallback
  const tasasPrev = (prev && prev.tasas) || {};
  const tasasSalida = {};

  // Heurística: si el valor previo viene del bug viejo del scraper (que mapeaba
  // todas las TASAS_SOLO_BNA a id=14 "Préstamos personales"), descartar y usar
  // fallback manual hasta que Playwright/CNAT scrape devuelvan valor real.
  const esValorLegadoBuggy = (k, v) => {
    if (!v || !TASAS_SOLO_BNA.includes(k)) return false;
    const fn = String(v.fuenteNombre || '');
    return /Pr[eé]stamos\s*personales/i.test(fn) || /aprox/i.test(fn);
  };

  Object.keys(FALLBACK.tasas).forEach(k => {
    if (reemplazarBCRA && reemplazarBCRA[k]) {
      // BCRA + Playwright consensus — prioridad máxima
      tasasSalida[k] = Object.assign({}, FALLBACK.tasas[k], reemplazarBCRA[k]);
    } else if (tasasPub && tasasPub[k]) {
      tasasSalida[k] = Object.assign({}, FALLBACK.tasas[k], tasasPub[k]);
    } else if (tasasPrev[k] && tasasPrev[k].status === 'ok' && !esValorLegadoBuggy(k, tasasPrev[k])) {
      tasasSalida[k] = Object.assign({}, tasasPrev[k], {status:'stale', nota:'Último valor válido · scraper no pudo refrescar este ciclo'});
    } else if (tasasPrev[k] && tasasPrev[k].status === 'manual') {
      tasasSalida[k] = tasasPrev[k];
    } else {
      // Cae a fallback manual (las TASAS_SOLO_BNA tienen valores razonables editables)
      tasasSalida[k] = Object.assign({}, FALLBACK.tasas[k]);
    }
  });

  // ── Acta 2783 / Plenario Barrios: tasa pura 6% + factor CER en runtime ──
  // El cliente (al-calc-bcra.js) calculará el monto = capital × (CER_hasta/CER_desde)
  //   + capital × 0.06 × años. El scraper aporta CER vigente y el 6% como tasaPura.
  if (cer && cer.valor) {
    ['acta2783', 'plenarioBarrios'].forEach(k => {
      tasasSalida[k] = {
        tasaAnual: 6,
        tasaPura: 6,
        cerReferencia: { fecha: cer.fecha, valor: cer.valor },
        formula: "Monto = C × (CER_hasta/CER_desde) + C × 0.06 × años",
        vigenteDesde: cer.fecha,
        fuenteNombre: k === 'acta2783'
          ? 'CNAT Acta 2783 — CER (BCRA id=30) + 6% pura anual'
          : 'Plenario Barrios SCBA — CER (BCRA id=30) + 6% pura anual',
        fuenteUrl: 'https://api.bcra.gob.ar/estadisticas/v4.0/monetarias/30',
        status: 'ok',
        fechaConsulta: new Date().toISOString(),
        nota: 'El factor CER se calcula dinámicamente con CER inicial y final del período. La tasa pura del 6% es jurisprudencial.'
      };
    });
  }

  const cntOK = Object.values(tasasSalida).filter(t => t.status === 'ok').length;
  console.log('[VALORES] Tasas OK (BCRA v4 + scrape): '+cntOK+'/'+Object.keys(tasasSalida).length);

  const salida = {
    ok: true,
    actualizadoAl: new Date().toISOString(),
    pba,
    ipc:    ipc    || elegir('ipc', 'ipc', null, FALLBACK.ipc),
    ipcGba: ipcGba || elegir('ipcGba', 'ipcGba', null, FALLBACK.ipcGba),
    ripte:  ripte  || elegir('ripte', 'ripte', null, FALLBACK.ripte),
    smvm:   smvm   || elegir('smvm', 'smvm', null, FALLBACK.smvm),
    bcra: {
      uva: uva || elegir('bcra','uva', null, FALLBACK.bcra.uva),
      cer: cer || elegir('bcra','cer', null, FALLBACK.bcra.cer),
      uvi: uvi || elegir('bcra','uvi', null, FALLBACK.bcra.uvi),
      icl: icl || elegir('bcra','icl', null, FALLBACK.bcra.icl)
    },
    canastaCrianza: canasta || elegir('canastaCrianza','canastaCrianza', null, FALLBACK.canastaCrianza),
    tasas: tasasSalida
  };

  logStatus('PBA', salida.pba);
  logStatus('BCRA', salida.bcra);
  logStatus('TASAS', salida.tasas);
  console.log(`[VALORES] IPC Nac: ${salida.ipc.status} (${salida.ipc.periodo}) ${salida.ipc.mensual}%`);
  console.log(`[VALORES] IPC GBA (IDECBA): ${salida.ipcGba.status} (${salida.ipcGba.periodo}) ${salida.ipcGba.mensual}%`);
  console.log(`[VALORES] RIPTE: ${salida.ripte.status} (${salida.ripte.periodo}) $${salida.ripte.valor}`);
  console.log(`[VALORES] SMVM: ${salida.smvm.status} (${salida.smvm.periodo}) $${salida.smvm.valor}`);
  console.log(`[VALORES] BADLAR: ${salida.tasas.badlar.status} ${salida.tasas.badlar.tasaAnual}%`);
  console.log(`[VALORES] Canasta: ${salida.canastaCrianza.status} (${salida.canastaCrianza.periodo})`);

  // Contar statuses
  const contar = (obj) => {
    const c = {ok:0, stale:0, fallback:0};
    const recorrer = (o) => {
      for (const k in o) {
        if (o[k] && typeof o[k] === 'object') {
          if (o[k].status) c[o[k].status] = (c[o[k].status]||0)+1;
          else recorrer(o[k]);
        }
      }
    };
    recorrer(obj);
    return c;
  };
  const stats = contar(salida);
  console.log(`[VALORES] Resumen: OK=${stats.ok} STALE=${stats.stale} FALLBACK=${stats.fallback}`);

  fs.writeFileSync(VALORES_PATH, JSON.stringify(salida, null, 2) + '\n', 'utf8');
  console.log('[VALORES] Escrito en', VALORES_PATH);

  // Nunca hacemos exit 1 por fallbacks — esos son esperados.
  // Sólo fallamos si hay un error no capturado (main().catch más abajo).
  // Notamos en stdout si hay muchos fallbacks para monitoreo:
  if (stats.fallback > 8) {
    console.warn('[VALORES] ATENCIÓN: '+stats.fallback+' valores en fallback — revisar fuentes');
  }
}

main().catch(err => {
  console.error('[VALORES] Error fatal:', err);
  process.exit(2);
});
