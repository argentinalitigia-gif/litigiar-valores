// ═════════════════════════════════════════════════════════════════════
//  LitigiAR — Scraper de valores oficiales (GitHub Actions)
//  Corre cada 6hs, produce valores.json en raíz del repo.
//  Fuentes:
//    - datos.gob.ar (IPC, UVA, CER, UVI, RIPTE, SMVM)
//    - BCRA diar_icl.xls (ICL)
//    - ColProBA / La Matanza / Morón (JUS, Bonos)
//    - INDEC Canasta Crianza XLSX
//
//  Política tolerante a fallos:
//    - Si un scraper falla, mantiene el valor anterior con status="stale"
//    - Si nunca hubo valor, cae a fallback con status="fallback"
//    - Nunca pierde un valor exitoso
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
  canastaCrianza: {periodo:"2026-03", menor1:515236, edad1a3:616046, edad4a5:538587, edad6a12:676431, fuenteNombre:"INDEC Canasta de Crianza", fuenteUrl:"https://www.indec.gob.ar/ftp/cuadros/sociedad/serie_canasta_crianza.xlsx", status:"fallback"}
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

async function fetchBCRAdiario(id){
  const r = await fetchSerieGob('https://apis.datos.gob.ar/series/api/series/?ids='+id+'&sort=desc&limit=1&metadata=none');
  if (!r) return null;
  return {
    fecha: String(r[0]).substring(0,10),
    valor: Number(Number(r[1]).toFixed(4)),
    fuenteNombre: 'BCRA / datos.gob.ar',
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

// ═══════ Orquestador ═══════
async function main(){
  console.log('[VALORES] Iniciando scrape…', new Date().toISOString());
  const prev = loadPrevious();

  const [ipc, uva, cer, uvi, icl, ripte, smvm, canasta, pba] = await Promise.all([
    fetchIPC().catch(()=>null),
    fetchBCRAdiario('94.2_UVAD_D_0_0_10').catch(()=>null),
    fetchBCRAdiario('94.2_CD_D_0_0_10').catch(()=>null),
    fetchBCRAdiario('94.2_UVID_D_0_0_10').catch(()=>null),
    scrapearICLBCRA().catch(()=>null),
    fetchRIPTE().catch(()=>null),
    fetchSMVM().catch(()=>null),
    scrapearCanastaCrianza().catch(()=>null),
    refreshPBA(prev)
  ]);

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

  const salida = {
    ok: true,
    actualizadoAl: new Date().toISOString(),
    pba,
    ipc:   ipc   || elegir('ipc', 'ipc', null, FALLBACK.ipc),
    ripte: ripte || elegir('ripte', 'ripte', null, FALLBACK.ripte),
    smvm:  smvm  || elegir('smvm', 'smvm', null, FALLBACK.smvm),
    bcra: {
      uva: uva || elegir('bcra','uva', null, FALLBACK.bcra.uva),
      cer: cer || elegir('bcra','cer', null, FALLBACK.bcra.cer),
      uvi: uvi || elegir('bcra','uvi', null, FALLBACK.bcra.uvi),
      icl: icl || elegir('bcra','icl', null, FALLBACK.bcra.icl)
    },
    canastaCrianza: canasta || elegir('canastaCrianza','canastaCrianza', null, FALLBACK.canastaCrianza)
  };

  logStatus('PBA', salida.pba);
  logStatus('BCRA', salida.bcra);
  console.log(`[VALORES] IPC: ${salida.ipc.status} (${salida.ipc.periodo}) ${salida.ipc.mensual}%`);
  console.log(`[VALORES] RIPTE: ${salida.ripte.status} (${salida.ripte.periodo}) $${salida.ripte.valor}`);
  console.log(`[VALORES] SMVM: ${salida.smvm.status} (${salida.smvm.periodo}) $${salida.smvm.valor}`);
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

  // Exit code según salud: si hay >5 fallback, algo está mal (excepto primera corrida)
  if (prev && stats.fallback > 5) {
    console.error('[VALORES] ALERTA: demasiados valores en fallback — revisar scrapers');
    process.exit(1);
  }
}

main().catch(err => {
  console.error('[VALORES] Error fatal:', err);
  process.exit(2);
});
