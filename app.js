'use strict';

const MASTER_HEADERS = [
  'ID ','FECHA','SEM','Provincia ','SOLICITANTE','MOTIVO','SOLPED','RUC',
  'NOMBRE DEL PROVEEDOR ','TIPO DOC','N° Doc','VALORIZADO','DETALLE SERVICIO',
  'N° OC','N° HES','Lib Solped','¿Solped Liberada','Lib OS','¿OS Liberada?',
  'OBSERVACIONES','MIRO','OBS. MIRO','ESTATUS'
];

const MASTER_FIELDS = [
  'id','fecha','sem','provincia','solicitante','motivo','solped','ruc','proveedor',
  'tipo_doc','documento','valorizado','detalle','oc','hes','lib_solped',
  'solped_liberada','lib_os','os_liberada','observaciones','miro','obs_miro','estatus'
];

const INCIDENT_HEADERS = ['SEM','PROVINCIA','DOCUMENTO','FUENTE','CAMPO','GRAVEDAD','MOTIVO'];
const YELLOW_HEADER_COLS = new Set([2,4,6,12,13,14,16,18,19]);
const BLOCKING = 'BLOQUEANTE';
const WARNING = 'ADVERTENCIA';

const FIELD_ALIASES = {
  id:['ID'], fecha:['FECHA','FECHADOCUMENTO','FECHASOLICITUD'], sem:['SEM','SEMANA'],
  provincia:['PROVINCIA','SEDE'], solicitante:['SOLICITANTE','AREA'], motivo:['MOTIVO','SERVICIO'],
  solped:['SOLPED','SOLICITUDDEPEDIDO','SOLICITUDPEDIDO'], ruc:['RUC','CODIGO'],
  proveedor:['NOMBREDELPROVEEDOR','NONBREDELPROVEEDOR','MONBREDELPROVEEDOR','PROVEEDOR','PROVEEDOR2'],
  tipo_doc:['TIPODOC','TIPODOCUMENTO'], documento:['NDOC','RXHN','DOCUMENTO','NUMERODOCUMENTO'],
  valorizado:['VALORIZADO','VALOR','VALORTOTAL','PRECIODEVALORACION'],
  detalle:['DETALLESERVICIO','TEXTOBREVE','DETALLE'], oc:['NOC','PEDIDO','ORDENDECOMPRA','OS'],
  hes:['NHES','HOJADEENTRADA','HE'], lib_solped:['LIBSOLPED','INDICADORLIBERACION'],
  solped_liberada:['SOLPEDLIBERADA'], lib_os:['LIBOS','INDLIBERACION'], os_liberada:['OSLIBERADA'],
  observaciones:['OBSERVACIONES'], miro:['MIRO','NMIRO'], obs_miro:['OBSMIRO'], estatus:['ESTATUS','STATUS'],
  centro:['CENTRO']
};

const state = {
  selectionCounter: 0,
  selections: new Map(),
  files: new Map(),
  fingerprints: new Set(),
  filePaths: new WeakMap(),
  previousFile: null,
  result: null,
  pdfInfo: [],
};

let pdfjsModulePromise = null;

async function pdfjsLibrary(){
  if(globalThis.pdfjsLib) return globalThis.pdfjsLib;
  if(!pdfjsModulePromise){
    pdfjsModulePromise=import('./vendor/pdf.min.mjs').then(pdfjs=>{
      pdfjs.GlobalWorkerOptions.workerSrc=new URL('./vendor/pdf.worker.min.mjs',window.location.href).href;
      return pdfjs;
    });
  }
  return pdfjsModulePromise;
}

const $ = id => document.getElementById(id);

window.addEventListener('load', () => {
  if (typeof XLSX === 'undefined') $('alert-nolib').classList.remove('hidden');
  bindUi();
  renderSelections();
});

function bindUi(){
  $('btn-root').addEventListener('click', event => { event.stopPropagation(); $('fi-root').click(); });
  $('btn-files').addEventListener('click', event => { event.stopPropagation(); $('fi-files').click(); });
  $('btn-previous').addEventListener('click', () => $('fi-previous').click());
  $('btn-clear').addEventListener('click', clearSelection);
  $('btn-process').addEventListener('click', processSelection);
  $('btn-download-master').addEventListener('click', downloadMaster);
  $('btn-download-sap').addEventListener('click', downloadSap);
  if($('btn-download-renamed')) $('btn-download-renamed').addEventListener('click', downloadRenamedZip);
  if($('btn-rename-local')) $('btn-rename-local').addEventListener('click', renameLocalFiles);
  $('fi-root').addEventListener('change', async event => {
    const added=await addSelection([...event.target.files], selectionLabel([...event.target.files], 'Carpeta'));
    event.target.value = '';
    if(added) await processSelection();
  });
  $('fi-files').addEventListener('change', async event => {
    const added=await addSelection([...event.target.files], `Archivos ${state.selectionCounter + 1}`);
    event.target.value = '';
    if(added) await processSelection();
  });
  $('fi-previous').addEventListener('change', async event => {
    state.previousFile = event.target.files[0] || null;
    event.target.value = '';
    renderSelections();
    if(state.result && state.previousFile){
      await applyManualToCurrentResult();
    }
  });
  ['filter-province','filter-type','filter-status'].forEach(id => $(id).addEventListener('change', renderPreview));
  $('filter-search').addEventListener('input', renderPreview);

  const drop = $('drop-zone');
  drop.addEventListener('dragover', event => { event.preventDefault(); drop.classList.add('drag-over'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
  drop.addEventListener('drop', async event => {
    event.preventDefault(); drop.classList.remove('drag-over');
    const dropped = await filesFromDrop(event.dataTransfer);
    const added=await addSelection(dropped, `Arrastre ${state.selectionCounter + 1}`);
    if(added) await processSelection();
  });
}

function selectionLabel(files, fallback){
  const path = files[0] ? relativePath(files[0]) : '';
  return path.includes('/') ? path.split('/')[0] : fallback;
}

function relativePath(file){
  return state.filePaths.get(file) || file.webkitRelativePath || file.name;
}

async function filesFromDrop(dataTransfer){
  const output = [];
  const items = [...(dataTransfer.items || [])];
  for(const item of items){
    if(item.getAsFileSystemHandle){
      try{
        const handle = await item.getAsFileSystemHandle();
        if(handle) await walkHandle(handle, '', output);
        continue;
      }catch(_){ /* fallback */ }
    }
    const entry = item.webkitGetAsEntry ? item.webkitGetAsEntry() : null;
    if(entry){ await walkEntry(entry, '', output); continue; }
    const file = item.getAsFile();
    if(file) output.push(file);
  }
  if(!output.length) output.push(...[...(dataTransfer.files || [])]);
  return output;
}

async function walkHandle(handle, prefix, output){
  const path = prefix ? `${prefix}/${handle.name}` : handle.name;
  if(handle.kind === 'file'){
    const file = await handle.getFile(); state.filePaths.set(file, path); output.push(file); return;
  }
  for await(const child of handle.values()) await walkHandle(child, path, output);
}

function readEntries(reader){
  return new Promise((resolve,reject) => reader.readEntries(resolve,reject));
}

async function walkEntry(entry, prefix, output){
  const path = prefix ? `${prefix}/${entry.name}` : entry.name;
  if(entry.isFile){
    const file = await new Promise((resolve,reject) => entry.file(resolve,reject));
    state.filePaths.set(file, path); output.push(file); return;
  }
  if(entry.isDirectory){
    const reader = entry.createReader();
    for(;;){
      const entries = await readEntries(reader);
      if(!entries.length) break;
      for(const child of entries) await walkEntry(child, path, output);
    }
  }
}

async function addSelection(files, label){
  const supported = files.filter(file => /\.pdf$/i.test(file.name) && !file.name.startsWith('~$'));
  if(!supported.length){ alert('No se encontraron archivos PDF compatibles en la selección.'); return false; }
  const id = `sel-${++state.selectionCounter}`;
  const keys = [];
  let duplicates = 0;
  for(const file of supported){
    const path = relativePath(file).replace(/\\/g,'/');
    const fingerprint = `${path.toUpperCase()}|${file.size}|${file.lastModified}`;
    if(state.fingerprints.has(fingerprint)){ duplicates++; continue; }
    state.fingerprints.add(fingerprint);
    const key = `${id}:${path}`;
    state.files.set(key,{file,path,selectionId:id,fingerprint}); keys.push(key);
  }
  if(keys.length) state.selections.set(id,{id,label:label || id,keys,duplicates});
  renderSelections();
  return keys.length>0;
}

function removeSelection(id){
  const selection = state.selections.get(id);
  if(!selection) return;
  for(const key of selection.keys){
    const entry = state.files.get(key);
    if(entry) state.fingerprints.delete(entry.fingerprint);
    state.files.delete(key);
  }
  state.selections.delete(id);
  state.result = null;
  renderSelections(); hideResults();
}

function clearSelection(){
  state.selections.clear(); state.files.clear(); state.fingerprints.clear();
  state.previousFile = null; state.result = null; state.pdfInfo = [];
  renderSelections(); hideResults();
}

function hideResults(){
  ['card-stats','card-preview','card-download'].forEach(id => $(id).classList.add('hidden'));
}

function renderSelections(){
  const list = $('selection-list'); list.innerHTML = '';
  $('btn-root').textContent = state.selections.size ? 'Agregar otra carpeta' : 'Agregar carpeta raíz';
  for(const selection of state.selections.values()){
    const row = document.createElement('div'); row.className = 'selection-row';
    const sites=selectionSites(selection),siteHtml=sites.length?`<div class="selection-sites">${sites.map(site=>`<span class="selection-site">📂 ${escapeHtml(site.code?`${site.code} ${site.province}`:site.province)} · ${site.files} archivos</span>`).join('')}</div>`:'';
    row.innerHTML = `<span class="folder-icon">📁</span><div class="selection-main"><div class="selection-name">${escapeHtml(selection.label)}</div><div class="selection-meta">${selection.keys.length} archivos${selection.duplicates ? ` · ${selection.duplicates} duplicados omitidos` : ''}</div>${siteHtml}</div><button class="btn btn-danger" type="button">Quitar</button>`;
    row.querySelector('button').addEventListener('click', () => removeSelection(selection.id)); list.appendChild(row);
  }
  if(state.previousFile) list.appendChild(specialSelection('✍️','Maestro editado',state.previousFile.name,() => { state.previousFile=null; renderSelections(); }));
  if(!list.children.length) list.innerHTML = '<div class="muted">Aún no hay carpetas ni archivos agregados.</div>';
}

function selectionSites(selection){
  const sites=new Map();
  for(const key of selection.keys){const entry=state.files.get(key);if(!entry)continue;const context=pathContext(entry.path);if(!context.province)continue;const siteKey=normProvince(context.province);if(!sites.has(siteKey))sites.set(siteKey,{code:context.code,province:context.province,files:0});sites.get(siteKey).files++;}
  return [...sites.values()].sort((a,b)=>(a.code||a.province).localeCompare(b.code||b.province));
}

function specialSelection(icon,title,name,onRemove){
  const row = document.createElement('div'); row.className='selection-row';
  row.innerHTML=`<span class="folder-icon">${icon}</span><div class="selection-main"><div class="selection-name">${title}</div><div class="selection-meta">${escapeHtml(name)}</div></div><button class="btn btn-danger" type="button">Quitar</button>`;
  row.querySelector('button').addEventListener('click',onRemove); return row;
}

function normalizeText(value){
  return String(value ?? '').replace(/�/g,'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim().toUpperCase();
}

function headerKey(value){ return normalizeText(value).replace(/[^A-Z0-9]+/g,''); }
function cleanId(value){
  if(value == null) return '';
  if(typeof value === 'number' && Number.isFinite(value)) return Number.isInteger(value) ? String(value) : String(value);
  return String(value).replace(/\u00a0/g,'').trim().replace(/\.0$/,'');
}
function numeric(value){
  if(value == null || value === '') return null;
  if(typeof value === 'number') return Number.isFinite(value) ? value : null;
  const match=String(value).toUpperCase().replace(/S\//g,'').replace(/PEN/g,'').replace(/,/g,'').match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : null;
}
function normProvince(value){ return normalizeText(value).replace(/^\d{3,5}[\s_-]*/,'').replace(/[^A-Z0-9]+/g,''); }
function displayProvince(value){ return normalizeText(value).replace(/^\d{3,5}[\s_-]*/,'').replace(/[_-]+/g,' ').trim(); }
function normDoc(value){
  const text=normalizeText(value).replace(/^0[12]-0/,'').replace(/\b(RXH|RHE|FACTURA|FACT)\b/g,'').replace(/[^A-Z0-9]+/g,'');
  const match=text.match(/^([A-Z]\d{3})(\d+)$/); return match ? match[1]+String(Number(match[2])) : text;
}
function displayDoc(value){
  const key=normDoc(value), match=key.match(/^([A-Z]\d{3})(\d+)$/); return match ? `${match[1]}-${Number(match[2])}` : cleanId(value);
}
function docType(value,ruc=''){
  const text=normalizeText(value); if(text.includes('FACT') || ['01','FAC'].includes(text)) return 'FACT';
  if(text.includes('RXH') || ['RH','RHE','02'].includes(text) || text.includes('HONOR')) return 'RXH';
  return cleanId(ruc).startsWith('20') ? 'FACT' : 'RXH';
}
function weekValue(value){ const match=cleanId(value).match(/\d+/); return match ? String(Number(match[0])) : cleanId(value); }
function asDate(value){
  if(value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if(typeof value === 'number'){
    const parsed=XLSX.SSF.parse_date_code(value); if(parsed) return new Date(Date.UTC(parsed.y,parsed.m-1,parsed.d));
  }
  const text=String(value || '').trim(); let m=text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if(m) return new Date(Date.UTC(Number(m[3]),Number(m[2])-1,Number(m[1])));
  m=text.match(/^(\d{4})-(\d{2})-(\d{2})/); if(m) return new Date(Date.UTC(Number(m[1]),Number(m[2])-1,Number(m[3])));
  return null;
}
function formatDate(value,separator='/'){
  const date=asDate(value); if(!date) return cleanId(value);
  return `${String(date.getUTCDate()).padStart(2,'0')}${separator}${String(date.getUTCMonth()+1).padStart(2,'0')}${separator}${date.getUTCFullYear()}`;
}
function excelDateSerial(value){
  const date=asDate(value);if(!date)return null;
  return Math.floor((Date.UTC(date.getUTCFullYear(),date.getUTCMonth(),date.getUTCDate())-Date.UTC(1899,11,30))/86400000);
}
function findField(value){
  const key=headerKey(value); return Object.keys(FIELD_ALIASES).find(field => FIELD_ALIASES[field].includes(key)) || null;
}
function detectHeader(rows){
  let best=null;
  for(let r=0;r<Math.min(rows.length,12);r++){
    const columns={}; (rows[r]||[]).forEach((value,c)=>{ const field=findField(value); if(field && columns[field]==null) columns[field]=c; });
    const score=Object.keys(columns).length; if(!best || score>best.score) best={score,row:r,columns};
  }
  if(!best || best.score<4) throw new Error('No se reconocieron los encabezados de una hoja Excel.');
  return best;
}
function getValue(row,columns,field){ const c=columns[field]; return c==null ? null : row[c]; }
function rowRecord(row,columns,source){
  const ruc=cleanId(getValue(row,columns,'ruc'));
  return {
    id:cleanId(getValue(row,columns,'id')), fecha:asDate(getValue(row,columns,'fecha')), sem:weekValue(getValue(row,columns,'sem')),
    provincia:displayProvince(getValue(row,columns,'provincia')), solicitante:cleanId(getValue(row,columns,'solicitante')),
    motivo:cleanId(getValue(row,columns,'motivo')), solped:cleanId(getValue(row,columns,'solped')), ruc,
    proveedor:cleanId(getValue(row,columns,'proveedor')), tipo_doc:docType(getValue(row,columns,'tipo_doc'),ruc),
    documento:displayDoc(getValue(row,columns,'documento')), valorizado:numeric(getValue(row,columns,'valorizado')),
    detalle:cleanId(getValue(row,columns,'detalle')), oc:cleanId(getValue(row,columns,'oc')), hes:cleanId(getValue(row,columns,'hes')),
    lib_solped:cleanId(getValue(row,columns,'lib_solped')), solped_liberada:numeric(getValue(row,columns,'solped_liberada')),
    lib_os:cleanId(getValue(row,columns,'lib_os')), os_liberada:numeric(getValue(row,columns,'os_liberada')),
    observaciones:cleanId(getValue(row,columns,'observaciones')), miro:cleanId(getValue(row,columns,'miro')),
    obs_miro:cleanId(getValue(row,columns,'obs_miro')), estatus:cleanId(getValue(row,columns,'estatus')),
    _source:source,_blocking:[],_warnings:[],_physicalSupported:false,_evidence:{},
  };
}

function groupKey(record){
  return [weekValue(record.sem),normProvince(record.provincia),cleanId(record.ruc).replace(/\D/g,''),docType(record.tipo_doc,record.ruc),normDoc(record.documento),cleanId(record.oc).replace(/\D/g,''),cleanId(record.hes).replace(/\D/g,'')].join('|');
}
function recordKey(record){ return `${groupKey(record)}|${headerKey(record.motivo)}|${record.valorizado==null?'':Number(record.valorizado).toFixed(2)}`; }

function readWorkbook(file){
  if(!file) return Promise.resolve(createDefaultReferenceWorkbook());
  if(file.SheetNames && file.Sheets) return Promise.resolve(file);
  return file.arrayBuffer().then(buffer => XLSX.read(new Uint8Array(buffer),{type:'array',cellDates:true,cellFormula:true}));
}
function sheetRows(workbook,name){
  const actual=workbook.SheetNames.find(item => normalizeText(item)===normalizeText(name));
  if(!actual) return [];
  return XLSX.utils.sheet_to_json(workbook.Sheets[actual],{header:1,defval:null,raw:true});
}
function parseRows(rows,source){
  const header=detectHeader(rows), output=[];
  for(const row of rows.slice(header.row+1)){
    if(!row || !row.some(value=>value!=null && value!=='')) continue;
    const record=rowRecord(row,header.columns,source);
    if(record.fecha || record.documento || record.solped) output.push(record);
  }
  return {records:output,columns:header.columns};
}

function pathContext(path){
  const parts=path.replace(/\\/g,'/').split('/'); let week='',code='',province='';
  for(const part of parts){
    const wm=part.match(/SEM(?:ANA)?[\s_-]*(\d+)/i); if(wm) week=String(Number(wm[1]));
    const sm=part.match(/^(\d{3,5})[\s_-]+(.+)$/); if(sm){code=sm[1];province=displayProvince(sm[2]);}
  }
  return {week,code,province};
}

function filenameInfo(entry){
  const context=pathContext(entry.path), normalized=normalizeText(entry.file.name.replace(/\.pdf$/i,''));
  let base=normalized, prefix=''; const pm=base.match(/^(\d{10})\s*-\s*/); if(pm){prefix=pm[1];base=base.slice(pm[0].length);}
  let type='OTRO',number='';
  for(const [label,pattern] of [['HES',/\bHES[\s_-]*(\d{8,12})/],['SOLPED',/\bSOLPED[\s_-]*(\d{7,12})/],['OC',/\bOC[\s_-]*(\d{8,12})/]]){
    const m=base.match(pattern); if(m){type=label;number=m[1];break;}
  }
  const documentMatches=[...base.matchAll(/\b(RXH|FACTURA|FACT)[\s_-]*([A-Z]\d{3})[\s_-]*(\d+)/g)];
  const documents=[...new Set(documentMatches.map(match=>`${match[2]}${Number(match[3])}`))],document=documentMatches[0]?`${documentMatches[0][2]}${Number(documentMatches[0][3])}`:'';
  const documentType=documentMatches[0]&&documentMatches[0][1].includes('FACT')?'FACT':document?'RXH':'';
  if(type==='OTRO' && document) type=documentType;
  const segments=base.split(/\s+-\s+/),providerHint=segments.length>1?segments[segments.length-1].trim():'';
  return {entry,province:context.province,siteCode:context.code,week:context.week,type,number,document,documents,documentType,providerHint,prefix};
}

function physicalIndex(entries){
  const index=new Map();
  for(const entry of entries.filter(item=>/\.pdf$/i.test(item.file.name))){
    const info=filenameInfo(entry), key=normProvince(info.province); if(!index.has(key)) index.set(key,[]); index.get(key).push(info);
  }
  return index;
}

function siteInventory(index){
  return [...index.entries()].filter(([key])=>key).map(([key,files])=>{
    const first=files[0]||{},hesCount=new Set(files.filter(item=>item.type==='HES'&&item.number).map(item=>item.number)).size;
    return {key,code:first.siteCode||'',province:first.province||key,pdfCount:files.length,hesCount};
  }).sort((a,b)=>(a.code||a.province).localeCompare(b.code||b.province));
}

function addIncident(incidents,record,source,field,severity,reason){
  const incident={week:record.sem,province:record.provincia,document:record.documento,source,field,severity,reason};
  if(!incidents.some(item=>JSON.stringify(item)===JSON.stringify(incident))) incidents.push(incident);
  const target=severity===BLOCKING ? '_blocking' : '_warnings'; if(!record[target].includes(reason)) record[target].push(reason);
}

function evidenceFor(record,index){
  const files=index.get(normProvince(record.provincia))||[], doc=normDoc(record.documento), oc=cleanId(record.oc).replace(/\D/g,''),hes=cleanId(record.hes).replace(/\D/g,''),solped=cleanId(record.solped).replace(/\D/g,'');
  const evidence={documento:[],oc:[],hes:[],solped:[]};
  for(const info of files){
    if(doc && info.document===doc && ['RXH','FACT'].includes(info.type)) evidence.documento.push(info.entry.path);
    if(oc && info.type==='OC' && info.number===oc) evidence.oc.push(info.entry.path);
    if(hes && info.type==='HES' && info.number===hes) evidence.hes.push(info.entry.path);
    if(solped && info.type==='SOLPED' && info.number===solped) evidence.solped.push(info.entry.path);
  }
  return evidence;
}

function validatePhysical(record,index,incidents){
  const evidence=evidenceFor(record,index); record._evidence=evidence;
  for(const field of ['oc','hes']){
    if(!evidence[field].length) addIncident(incidents,record,'Expediente PDF',field.toUpperCase(),BLOCKING,`No se encontró el PDF de ${field.toUpperCase()} en la carpeta de la provincia.`);
    else if(evidence[field].length>1) addIncident(incidents,record,'Expediente PDF',field.toUpperCase(),BLOCKING,`Se encontraron varios PDF candidatos para ${field.toUpperCase()}.`);
  }
  if(evidence.documento.length>1){
    addIncident(incidents,record,'Expediente PDF','DOCUMENTO',BLOCKING,'Se encontraron varios PDF candidatos para DOCUMENTO.');
  }else if(evidence.documento.length===0 && (!record.documento || record.valorizado==null)){
    addIncident(incidents,record,'Expediente PDF','DOCUMENTO',BLOCKING,'No se encontró el PDF de DOCUMENTO ni respaldo del importe en la carpeta de la provincia.');
  }
  record._physicalSupported=evidence.oc.length===1 && evidence.hes.length===1 && (evidence.documento.length===1 || (evidence.documento.length===0 && Boolean(record.documento) && record.valorizado!=null));
}

function validateRequired(record,incidents){
  for(const field of ['fecha','sem','provincia','ruc','tipo_doc','documento','valorizado','oc','hes']){
    if(record[field]==null || record[field]===''){
      addIncident(incidents,record,record._source,field.toUpperCase(),BLOCKING,`Falta el campo esencial ${field.toUpperCase()}.`);
    }
  }
}

function deduplicate(records,incidents){
  const seen=new Map(),output=[];
  for(const record of records){
    const key=recordKey(record); if(!seen.has(key)){seen.set(key,record);output.push(record);} else addIncident(incidents,seen.get(key),'Consolidación','DUPLICADO',WARNING,'El mismo registro apareció en varias selecciones y se conservó una sola vez.');
  }
  return output;
}

function finalizeRecord(record){
  const notes=[];
  if(record._blocking.length) notes.push(`BLOQUEANTE: ${record._blocking.join(' | ')}`);
  if(record._warnings.length) notes.push(`ADVERTENCIA: ${record._warnings.join(' | ')}`);
  if(record._manualObservation) notes.push(record._manualObservation);
  record.observaciones=notes.join(' | ');

  const miroClean=cleanId(record.miro||'');
  const isValidMiro=/^\d{10}$/.test(miroClean);

  if(record._blocking.length){
    record.estatus='OBSERVADO';
  } else if(!miroClean){
    record.estatus='PENDIENTE MIRO';
  } else if(isValidMiro && (record._physicalSupported || record._physical_supported)){
    record.estatus='CERRADO';
  } else {
    record.estatus='OBSERVADO';
  }
}

const MIRO_CACHE_KEY = 'calera_miro_cache_v1';

function getMiroCache(){
  try{
    return JSON.parse(localStorage.getItem(MIRO_CACHE_KEY) || '{}');
  }catch(e){
    return {};
  }
}

function saveMiroToCache(key, miro){
  if(!key) return;
  try{
    const cache = getMiroCache();
    if(miro && /^\d{10}$/.test(miro)){
      cache[key] = miro;
    } else if(!miro){
      delete cache[key];
    }
    localStorage.setItem(MIRO_CACHE_KEY, JSON.stringify(cache));
  }catch(e){}
}

function clearMiroCache(){
  try{
    localStorage.removeItem(MIRO_CACHE_KEY);
  }catch(e){}
}

async function parseManual(file){
  if(!file) return new Map(); const wb=await readWorkbook(file), records=parseRows(sheetRows(wb,'BD'),`${file.name}:BD`).records, grouped=new Map();
  for(const record of records){const key=groupKey(record);if(!grouped.has(key))grouped.set(key,[]);grouped.get(key).push(record);}
  const output=new Map();
  for(const [key,items] of grouped){
    const miros=new Set(items.map(item=>item.miro).filter(Boolean)); output.set(key,{conflict:miros.size>1,miro:miros.size===1?[...miros][0]:'',obs_miro:items.find(item=>item.obs_miro)?.obs_miro||'',observaciones:items.find(item=>item.observaciones)?.observaciones||''});
  }
  return output;
}

function includesDocument(info,document){return Boolean(document)&&((info.documents||[]).includes(document)||info.document===document);}
function uniqueValue(values){const unique=[...new Set(values.filter(Boolean))];return unique.length===1?unique[0]:'';}

function cleanProviderName(rawProvider,wordHints=[]){
  let provider=String(rawProvider||'').replace(/\s+/g,' ').trim();
  if(!provider)return '';
  const words=[...new Set(wordHints.flatMap(hint=>normalizeText(hint).split(/[^A-Z0-9]+/)).filter(word=>word.length>=3))].sort((a,b)=>b.length-a.length);
  for(const word of words){
    const escaped=word.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
    provider=provider.replace(new RegExp(`([A-Z]{2,})(${escaped})\\b`,'g'),'$1 $2');
    provider=provider.replace(new RegExp(`\\b(${escaped})([A-Z]{2,})`,'g'),'$1 $2');
  }
  return provider.replace(/\s+/g,' ').trim();
}

function inferServiceWeek(documentInfo,ocInfo,hesInfo){
  if(documentInfo&&documentInfo.date){
    const docDate=asDate(documentInfo.date);
    if(docDate){
      const iso=Number(isoWeekFromDate(docDate));
      return String(docDate.getUTCDay()===0?iso:Math.max(1,iso-1));
    }
  }
  const sapDate=(ocInfo&&(ocInfo.deliveryDate||ocInfo.date))||(hesInfo&&hesInfo.date)||null;
  if(sapDate){
    const iso=Number(isoWeekFromDate(sapDate));
    if(iso)return String(Math.max(1,iso-1));
  }
  return '';
}

function recordsFromPdfExpedients(pdfIndex,incidents){
  const records=[];
  for(const [provinceKey] of pdfIndex){
    const parsedFiles=state.pdfInfo.filter(info=>normProvince(info.province)===provinceKey);
    const hesFiles=parsedFiles.filter(info=>info.type==='HES'&&info.number);
    for(const hesInfo of hesFiles){
      const document=hesInfo.document||'',documentType=hesInfo.documentType||docType('',hesInfo.ruc),byDocument=parsedFiles.filter(info=>includesDocument(info,document));
      const ocCandidates=parsedFiles.filter(info=>info.type==='OC'&&((hesInfo.oc&&info.number===hesInfo.oc)||(!hesInfo.oc&&includesDocument(info,document))));
      const ocNumber=hesInfo.oc||uniqueValue(ocCandidates.map(info=>info.number)),ocInfo=ocCandidates.find(info=>info.number===ocNumber)||ocCandidates[0]||null;
      const solpedCandidates=byDocument.filter(info=>info.type==='SOLPED'),solped=uniqueValue(solpedCandidates.map(info=>info.number)),solpedInfo=solpedCandidates.find(info=>info.number===solped)||null;
      const documentCandidates=byDocument.filter(info=>['FACT','RXH'].includes(info.type)),documentInfo=documentCandidates[0]||null;
      const sourceInfos=[documentInfo,ocInfo,hesInfo,solpedInfo].filter(Boolean);
      const date=(documentInfo&&documentInfo.date)||(ocInfo&&(ocInfo.deliveryDate||ocInfo.date))||(hesInfo&&hesInfo.date)||(solpedInfo&&solpedInfo.date)||null;
      let week=sourceInfos.map(info=>info.week).find(Boolean)||hesInfo.week||'';
      if(!week)week=inferServiceWeek(documentInfo,ocInfo,hesInfo);
      const rawProvider=sourceInfos.map(info=>info.provider).find(Boolean)||hesInfo.providerHint||'';
      const provider=cleanProviderName(rawProvider,[hesInfo.providerHint,hesInfo.hesPersonHint,ocInfo&&ocInfo.providerHint,documentInfo&&documentInfo.providerHint]);
      const valorizado=(documentInfo&&documentInfo.total!=null)?documentInfo.total:(hesInfo.serviceTotal!=null?hesInfo.serviceTotal:(ocInfo&&ocInfo.serviceTotal!=null?ocInfo.serviceTotal:null));
      const hasSolpedOrOc=Boolean(solped||ocNumber),hasOcAndHes=Boolean(ocNumber&&hesInfo.number);
      const fileMiro=sourceInfos.map(info=>info.prefix).find(Boolean)||'';
      const record={
        id:'',fecha:asDate(date),sem:week,provincia:hesInfo.province,solicitante:sourceInfos.map(info=>info.solicitante).find(Boolean)||'ALMACEN',
        motivo:sourceInfos.map(info=>info.motivo).find(Boolean)||'',solped,ruc:sourceInfos.map(info=>info.ruc).find(Boolean)||'',
        proveedor:provider,tipo_doc:documentType,
        documento:displayDoc(document),valorizado,
        detalle:sourceInfos.map(info=>info.detalle).find(Boolean)||'',oc:ocNumber,hes:hesInfo.number,
        lib_solped:hasSolpedOrOc?'2':'',solped_liberada:hasSolpedOrOc?2:0,lib_os:hasOcAndHes?'G':'',os_liberada:hasOcAndHes?2:0,observaciones:'',miro:fileMiro,obs_miro:'',estatus:'',
        _source:hesInfo.entry.path,_sitePath:hesInfo.entry.path,_context:pathContext(hesInfo.entry.path),_blocking:[],_warnings:[],_physicalSupported:false,_evidence:{},
      };
      if(!hesInfo.readable)addIncident(incidents,record,hesInfo.entry.path,'LECTURA PDF',WARNING,'El PDF HES no permitió extraer texto; se conservaron los datos disponibles en el nombre del archivo.');
      if(ocCandidates.length>1)addIncident(incidents,record,'Expediente PDF','N° OC',BLOCKING,'Se encontraron varios PDF de OC candidatos para el mismo expediente.');
      if(solpedCandidates.length>1)addIncident(incidents,record,'Expediente PDF','SOLPED',BLOCKING,'Se encontraron varios PDF de SOLPED candidatos para el mismo documento.');
      if(documentCandidates.length>1)addIncident(incidents,record,'Expediente PDF','DOCUMENTO',BLOCKING,'Se encontraron varios comprobantes PDF para el mismo documento.');
      records.push(record);
    }
  }
  return records;
}

async function processSelection(){
  if(typeof XLSX==='undefined'){alert('SheetJS no está disponible.');return;}
  const entries=[...state.files.values()];
  if(!entries.length){alert('Agrega una carpeta con expedientes para comenzar.');return;}
  state.result=null;
  hideResults();
  $('card-stats').classList.remove('hidden');
  $('stats-grid').innerHTML='';
  $('folder-list').innerHTML='';
  $('incidence-list').innerHTML='';
  setProgress(0,'Preparando archivos…');
  try{
    const pIndex=physicalIndex(entries),sites=siteInventory(pIndex);
    const incidents=[];
    setProgress(5,'Leyendo todos los PDF…');
    state.pdfInfo=[]; const pdfEntries=entries.filter(entry=>/\.pdf$/i.test(entry.file.name));
    for(let i=0;i<pdfEntries.length;i++){
      state.pdfInfo.push(await parsePdf(pdfEntries[i]));
      setProgress(5+70*(i+1)/Math.max(pdfEntries.length,1),`Leyendo PDF ${i+1}/${pdfEntries.length}`);
    }
    let records=recordsFromPdfExpedients(pIndex,incidents);
    if(!records.length) throw new Error('No se encontraron PDF HES para generar registros. Cada expediente del maestro se crea a partir de una HES.');
    applyPdfAmounts(records,incidents);
    for(const record of records){validatePhysical(record,pIndex,incidents);validateRequired(record,incidents);}
    const manual=await parseManual(state.previousFile);
    for(const record of records){
      const values=manual.get(groupKey(record)); if(values){if(values.conflict)addIncident(incidents,record,'Maestro editado','MIRO',BLOCKING,'El maestro editado contiene MIRO distintos para la misma clave documental.');else if(values.miro&&!/^\d{10}$/.test(values.miro))addIncident(incidents,record,'Maestro editado','MIRO',BLOCKING,`El MIRO ${values.miro} no tiene 10 dígitos.`);else {if(values.miro){record.miro=values.miro;saveMiroToCache(groupKey(record),values.miro);if(record.hes)saveMiroToCache('HES:'+cleanId(record.hes),values.miro);}record.obs_miro=values.obs_miro;record._manualObservation=values.observaciones;}}
    }
    const cachedMiros=getMiroCache();
    for(const record of records){
      if(!record.miro){
        const gk=groupKey(record), hk=record.hes?'HES:'+cleanId(record.hes):null;
        const cached=cachedMiros[gk]||(hk?cachedMiros[hk]:null);
        if(cached&&/^\d{10}$/.test(cached)){
          record.miro=cached;
        }
      }
    }
    records=deduplicate(records,incidents);
    records.sort((a,b)=>(a._context.code||a.provincia).localeCompare(b._context.code||b.provincia)||Number(a.hes||0)-Number(b.hes||0));
    records.forEach((record,index)=>{record.id=index+1;finalizeRecord(record);});
    state.result={week:'Todas',mode:'expedientes',records,incidents,selectedRecords:records,files:pdfEntries.length,sites};
    setProgress(100,'Procesamiento completo'); renderResults();
  }catch(error){console.error(error);alert(error.message);setProgress(0,'Error de procesamiento');}
}

function setProgress(percent,label){
  $('prog-fill').style.width=`${Math.max(0,Math.min(100,percent))}%`;$('prog-pct').textContent=`${Math.round(percent)}%`;$('prog-label').textContent=label;
}

function renderStatsGrid(){
  if(!state.result)return;
  const {records,incidents}=state.result, supported=records.filter(r=>r._physicalSupported).length, provinces=[...new Set(records.map(r=>r.provincia))].sort(),sites=state.result.sites||[];
  const totalValorizado=records.reduce((sum,r)=>sum+(Number(r.valorizado)||0),0);
  const valorizadoFormatted='S/ '+totalValorizado.toLocaleString('es-PE',{minimumFractionDigits:2,maximumFractionDigits:2});
  const matchRate=records.length?((supported/records.length)*100).toFixed(1)+'%':'0.0%';
  const cerradosCount=records.filter(r=>r.estatus==='CERRADO').length;
  const pendientesMiroCount=records.filter(r=>r.estatus==='PENDIENTE MIRO').length;
  const observadosCount=records.filter(r=>r.estatus==='OBSERVADO').length;
  const blockingCount=incidents.filter(i=>i.severity===BLOCKING).length;
  const warnCount=incidents.filter(i=>i.severity!==BLOCKING).length;

  $('stats-grid').innerHTML=[
    { label:'Tasa Conciliación', value:matchRate, cls:'kpi-accent', valCls:'kpi-blue', sub:`${supported} de ${records.length} soportados` },
    { label:'Total Valorizado', value:valorizadoFormatted, cls:'kpi-gold', valCls: 'money', sub:'Importe consolidado' },
    { label:'Cerrados Conforme', value:cerradosCount, cls:'', valCls:'', sub:'Listos con MIRO 10 dígitos' },
    { label:'Pendiente MIRO', value:pendientesMiroCount, cls:'kpi-warn', valCls:'kpi-warn', sub:'Esperando registro SAP' },
    { label:'Observados', value:observadosCount, cls:observadosCount?'kpi-alert':'', valCls:observadosCount?'kpi-red':'', sub:'Requieren subsanación' },
    { label:'Bloqueantes / Warnings', value:`${blockingCount} / ${warnCount}`, cls:blockingCount?'kpi-alert':'', valCls:blockingCount?'kpi-red':'', sub:`${incidents.length} incidencias totales` },
    { label:'HES Procesadas', value:records.filter(r=>r.hes).length, cls:'', valCls:'', sub:`${records.filter(r=>r.solped).length} con SOLPED vinculada` },
    { label:'Sedes / Carpetas', value:`${provinces.length} / ${sites.length}`, cls:'', valCls:'', sub:`${state.result.files||0} archivos PDF leídos` },
  ].map(card=>`
    <div class="stat-box ${card.cls}">
      <div class="stat-value ${card.valCls}">${card.value}</div>
      <div class="stat-label">${card.label}</div>
      <div class="stat-sub">${card.sub}</div>
    </div>
  `).join('');
}

function renderResults(){
  const {records,incidents}=state.result, provinces=[...new Set(records.map(r=>r.provincia))].sort(),sites=state.result.sites||[];
  renderStatsGrid();
  const counts=new Map();records.forEach(r=>counts.set(r.provincia,(counts.get(r.provincia)||0)+1));
  $('folder-list').innerHTML=sites.map(site=>{const matched=[...counts].find(([province])=>normProvince(province)===site.key),count=matched?matched[1]:0;return `<div class="folder-item ${count?'done':'warning'}"><span class="folder-icon">📁</span><div class="folder-info"><div class="folder-name">${escapeHtml(site.code?`${site.code}_${site.province}`:site.province)}</div><div class="folder-meta">${site.pdfCount} PDF detectados · ${site.hesCount} expedientes HES · ${count} registros generados</div></div><span class="folder-badge ${count?'':'warn'}">${count?`${count} registros`:'Sin HES'}</span></div>`;}).join('');
  $('incidence-list').innerHTML=incidents.slice(0,80).map(item=>`<div class="incidence ${item.severity===BLOCKING?'blocking':''}"><strong>${escapeHtml(item.severity)}</strong> · ${escapeHtml(item.province)} · ${escapeHtml(item.document||item.field)} — ${escapeHtml(item.reason)}</div>`).join('') || '<div class="muted">No se encontraron incidencias.</div>';
  $('filter-province').innerHTML='<option value="">Todas</option>'+sites.map(site=>`<option>${escapeHtml(site.province)}</option>`).join('');
  ['card-stats','card-preview','card-download'].forEach(id=>$(id).classList.remove('hidden'));renderPreview();
}

function previewValue(record,field){
  const value=record[field]; if(field==='fecha')return formatDate(value);if(field==='documento')return sapRef(record.documento,record.tipo_doc);if(field==='valorizado')return value==null?'':`S/ ${Number(value).toLocaleString('es-PE',{minimumFractionDigits:2,maximumFractionDigits:2})}`;return cleanId(value);
}
function trafficLightHtml(val){
  const num=Number(val);
  if(val===''||val==null||isNaN(num))return '<span class="muted">—</span>';
  if(num>=2)return '<span class="traffic-light traffic-green" title="Liberado (2)"></span>';
  if(num>=1)return '<span class="traffic-light traffic-yellow" title="Pendiente (1)"></span>';
  return '<span class="traffic-light traffic-red" title="No liberado (0)"></span>';
}
function renderPreview(){
  if(!state.result)return;const province=$('filter-province').value,type=$('filter-type').value,status=$('filter-status').value,query=normalizeText($('filter-search').value);
  const rows=state.result.records.filter(r=>(!province||r.provincia===province)&&(!type||r.tipo_doc===type)&&(!status||r.estatus===status)&&(!query||normalizeText([r.solped,r.ruc,r.oc,r.hes,r.documento,r.proveedor].join(' ')).includes(query)));
  $('preview-thead').innerHTML='<tr>'+MASTER_HEADERS.map((h,i)=>`<th class="${YELLOW_HEADER_COLS.has(i)?'th-teal':''} ${i===16||i===18?'th-center':''}">${escapeHtml(h)}</th>`).join('')+'</tr>';
  $('preview-tbody').innerHTML=rows.length?rows.slice(0,100).map(record=>'<tr>'+MASTER_FIELDS.map(field=>{
    if(field==='estatus'){const cls=record.estatus==='OBSERVADO'?'status-observado':record.estatus==='CERRADO'?'status-cerrado':'status-pendiente';return `<td class="td-status-${record.id}"><span class="status-pill ${cls}">${escapeHtml(record.estatus)}</span></td>`;}
    if(field==='solped_liberada'||field==='os_liberada')return `<td class="td-center">${trafficLightHtml(record[field])}</td>`;
    if(field==='miro'){
      const val=cleanId(record.miro||'');
      return `<td class="td-center"><input type="text" class="miro-cell-input ${val?'has-value':''}" data-id="${record.id}" value="${escapeHtml(val)}" placeholder="10 dígitos" maxlength="10"></td>`;
    }
    return `<td class="${field==='valorizado'?'td-amt':''}">${escapeHtml(previewValue(record,field))}</td>`;
  }).join('')+'</tr>').join(''):`<tr class="empty-preview"><td colspan="${MASTER_FIELDS.length}">No hay registros PDF que cumplan los filtros seleccionados.</td></tr>`;
  document.querySelectorAll('.miro-cell-input').forEach(input=>{
    input.addEventListener('change', event=>{
      const id=Number(event.target.getAttribute('data-id')), rec=state.result.records.find(r=>r.id===id);
      if(!rec)return;
      const raw=event.target.value.replace(/\D/g,'').slice(0,10);
      event.target.value=raw; rec.miro=raw;
      saveMiroToCache(groupKey(rec), raw);
      if(rec.hes) saveMiroToCache('HES:'+cleanId(rec.hes), raw);
      if(raw)event.target.classList.add('has-value'); else event.target.classList.remove('has-value');
      finalizeRecord(rec);
      const statusTd=document.querySelector(`.td-status-${rec.id}`);
      if(statusTd){
        const cls=rec.estatus==='OBSERVADO'?'status-observado':rec.estatus==='CERRADO'?'status-cerrado':'status-pendiente';
        statusTd.innerHTML=`<span class="status-pill ${cls}">${escapeHtml(rec.estatus)}</span>`;
      }
      renderStatsGrid();
    });
  });
}

function pdfLiteralStrings(source){
  const values=[];
  for(let i=0;i<source.length;i++){
    if(source[i]!=='(')continue;
    let value='',depth=1;
    while(++i<source.length&&depth){
      const ch=source[i];
      if(ch==='\\'){
        const escaped=source[++i];
        if(escaped==='n')value+='\n';
        else if(escaped==='r')value+='\r';
        else if(escaped==='t')value+='\t';
        else if(escaped==='b')value+='\b';
        else if(escaped==='f')value+='\f';
        else if(escaped==='('||escaped===')'||escaped==='\\')value+=escaped;
        else if(/[0-7]/.test(escaped||'')){
          let octal=escaped;
          for(let count=0;count<2&&/[0-7]/.test(source[i+1]||'');count++)octal+=source[++i];
          value+=String.fromCharCode(parseInt(octal,8));
        }else if(escaped!=='\n'&&escaped!=='\r')value+=escaped||'';
      }else if(ch==='('){depth++;value+=ch;}
      else if(ch===')'){if(--depth)value+=ch;}
      else value+=ch;
    }
    if(!depth)values.push(value);
  }
  return values.join(' ');
}
const MONTHS={enero:'01',febrero:'02',marzo:'03',abril:'04',mayo:'05',junio:'06',julio:'07',agosto:'08',setiembre:'09',septiembre:'09',octubre:'10',noviembre:'11',diciembre:'12'};

function isoWeekFromDate(value){
  const date=asDate(value);if(!date)return '';
  const target=new Date(Date.UTC(date.getUTCFullYear(),date.getUTCMonth(),date.getUTCDate()));
  target.setUTCDate(target.getUTCDate()+4-(target.getUTCDay()||7));
  const yearStart=new Date(Date.UTC(target.getUTCFullYear(),0,1));
  return String(Math.ceil((((target-yearStart)/86400000)+1)/7));
}

function pdfDate(text){
  let match=text.match(/FECHA\s+DE\s+EMISI[OÓ]N\s*:?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{4})/i)
    ||text.match(/N[UÚ]MERO\s*\/?\s*FECHA\s+\d{8,12}\s*\/\s*(\d{1,2}[./-]\d{1,2}[./-]\d{4})/i);
  if(match)return match[1].replace(/[.-]/g,'/');
  match=text.match(/(\d{1,2})\s+(?:DE\s+)?([A-ZÁÉÍÓÚÑ]+)\s+(?:DEL?\s+)?(\d{4})/i);
  if(match){const month=MONTHS[normalizeText(match[2]).toLowerCase()];if(month)return `${String(match[1]).padStart(2,'0')}/${month}/${match[3]}`;}
  return null;
}

function pdfDeliveryDate(text){
  const match=text.match(/\b\d{4}\s*-\s*CD-[A-ZÁÉÍÓÚÑ_-]+\s+(\d{1,2}[./-]\d{1,2}[./-]\d{4})/i);
  return match?match[1].replace(/[.-]/g,'/'):null;
}

function pdfMoney(value){return value?Number(value.replace(/,/g,'')):null;}

function serviceMotive(text){
  const value=normalizeText(text);
  if(/MONTACARGA|MAQUINARIA PESADA/.test(value))return 'SERV. OP. MONTACARGA';
  if(/LIMPIEZA/.test(value))return 'SERV. LIMPIEZA';
  if(/ESTIBA|DESESTIBA/.test(value))return 'SERV. ESTIBA';
  if(/SANEO|RESANEO|SELECCION/.test(value))return 'SERV. SANEO';
  const compact=value.replace(/_{3,}[\s\S]*/,'').trim();
  return compact.length<=80?compact:compact.slice(0,80).trim();
}

function pdfFields(text,base){
  const clean=text.replace(/[\x00-\x1f\ufffd\xa0]/g,' ').replace(/\s+/g,' ').trim(),upper=normalizeText(clean);
  const isFact=base.documentType==='FACT'||base.type==='FACT'||upper.includes('FACTURA ELECTRONICA');
  const isRxh=base.documentType==='RXH'||base.type==='RXH'||upper.includes('RECIBO POR HONORARIOS');
  const kind=['HES','OC','SOLPED'].includes(base.type)?base.type:isFact?'FACT':isRxh?'RXH':base.type;
  let total=null,serviceTotal=null;
  if(isFact){const match=upper.match(/IMPORTE\s+TOTAL\s*:?\s*(?:S\/?\s*)?([0-9,]+\.\d{2})/);if(match)total=pdfMoney(match[1]);}
  if(isRxh){const match=upper.match(/TOTAL\s+NETO\s+RECIBIDO[\s\S]{0,180}?([0-9,]+\.\d{2})\s+SOLES/)||upper.match(/TOTAL\s+POR\s+HONORARIOS\s*:?\s*(?:S\/?\s*)?([0-9,]+\.\d{2})/)||upper.match(/\b20\d{2}\s+([0-9,]+\.\d{2})\b/);if(match)total=pdfMoney(match[1]);}
  const serviceMatch=upper.match(/VALOR\s+TOTAL\s+SIN\s+IVA\s+PEN\s+([0-9,]+\.\d{2})/);
  if(serviceMatch)serviceTotal=pdfMoney(serviceMatch[1]);
  if(serviceTotal==null && kind==='OC'){
    const netMatch=upper.match(/\bNETO\s+DESCUENTO\s+SUB-TOTAL\s+I\.G\.V\.\s+TOTAL\s+([0-9,]+\.\d{2})/);
    if(netMatch)serviceTotal=pdfMoney(netMatch[1]);
  }
  let ruc='',provider='',hesPersonHint='';
  const vendor=upper.match(/PROVEEDOR\s*:\s*(.+?)\s+RUC\s*:\s*(\d{11})/);
  if(vendor){provider=vendor[1].trim();ruc=vendor[2];}
  if(!ruc&&isFact){const match=upper.match(/FACTURA\s+ELECTRONICA\s+RUC\s*:\s*(\d{11})/);if(match)ruc=match[1];}
  if(!ruc&&isRxh){const match=upper.match(/NRO\s*:\s*(\d{11})\s*-\s*[A-Z]\d{3}/);if(match)ruc=match[1];}
  const hesPersonMatch=upper.match(/VALOR\s+TOTAL\s+SIN\s+IVA\s+PEN\s+[0-9,]+\.\d{2}\s+(.{3,65}?)\s+(?:JR\b|CALLE\b|AV\b|URB\b|AH\b|AA\s*HH\b|ASOC\b|ASC\b|FUNDO\b|GRAN\b|LAS\b|P\.?J\.?\b|MZA?\b|LOTE\b|LT\b|\d{5})/);
  if(hesPersonMatch)hesPersonHint=hesPersonMatch[1].trim();
  const ocMatch=upper.match(/POSICION\s+PEDIDO\s*\/?\s*FECHA\s+(43\d{7,10})/)||upper.match(/ORDEN\s+DE\s+SERVICIO\s+N\s*(43\d{7,10})/);
  const hesMatch=upper.match(/N[UÚ]MERO\s*\/?\s*FECHA\s+(10\d{8,10})/);
  const weekMatch=upper.match(/\bSEM(?:ANA)?\s*[-:]?\s*(\d{1,2})\b/);
  let detail='';
  const detailMatch=upper.match(/\b\d{7}\s+(.{3,120}?)(?:L.{0,8}PLANIFICACI|LINEA\s+PLANIFICACION)/);
  if(detailMatch)detail=detailMatch[1].trim();
  const lead=clean.split(/_{5,}/)[0].trim();
  return {text:clean,kind,total,serviceTotal,date:pdfDate(clean),deliveryDate:pdfDeliveryDate(clean),week:weekMatch?String(Number(weekMatch[1])):'',ruc,provider,hesPersonHint,oc:ocMatch?ocMatch[1]:'',hes:hesMatch?hesMatch[1]:'',solicitante:upper.includes('ALM_LIMA_PRO')?'ALMACEN':'',motivo:base.type==='HES'?serviceMotive(lead):serviceMotive(detail),detalle:detail||lead};
}

async function parsePdf(entry){
  try{
    const base=filenameInfo(entry),pdfjs=await pdfjsLibrary(),bytes=new Uint8Array(await entry.file.arrayBuffer());
    const document=await pdfjs.getDocument({data:bytes,verbosity:0,useSystemFonts:true,disableFontFace:true}).promise;let text='';
    for(let pageNumber=1;pageNumber<=document.numPages;pageNumber++){
      const page=await document.getPage(pageNumber),content=await page.getTextContent();
      text+=' '+content.items.map(item=>item.str||'').join(' ');
    }
    await document.destroy();
    return {...base,...pdfFields(text,base),readable:true};
  }catch(error){const base=filenameInfo(entry);console.warn('PDF no legible',entry.file.name,error);return {...base,kind:base.type,total:null,serviceTotal:null,date:null,deliveryDate:null,ruc:'',provider:'',hesPersonHint:'',oc:'',hes:'',solicitante:'',motivo:'',detalle:'',week:'',readable:false,error:error.message};}
}

function applyPdfAmounts(records,incidents){
  for(const record of records){
    const documents=record._evidence?.documento||[];
    if(documents.length===1){
      const info=state.pdfInfo.find(item=>item.entry.path===documents[0]&&item.document===normDoc(record.documento)&&['FACT','RXH'].includes(item.kind));
      if(info&&info.total!=null) record.valorizado=info.total;
    }
    if(record.valorizado==null){
      addIncident(incidents,record,record._source,'VALORIZADO',BLOCKING,'No se pudo extraer el importe total del expediente PDF.');
    }
  }
  records.forEach(finalizeRecord);
}

function statusFormula(row){return `IF(OR(B${row}="",C${row}="",D${row}="",H${row}="",J${row}="",K${row}="",L${row}="",N${row}="",O${row}="",LEFT(T${row},10)="BLOQUEANTE"),"OBSERVADO",IF(U${row}="","PENDIENTE MIRO","CERRADO"))`;}
function workbookBase(){const wb=XLSX.utils.book_new();wb.Workbook={CalcPr:{calcMode:'auto',fullCalcOnLoad:true,forceFullCalc:true}};return wb;}
function headerStyle(yellow=false){return {fill:{fgColor:{rgb:yellow?'FFF200':'00B050'}},font:{bold:true,color:{rgb:yellow?'000000':'FFFFFF'},name:'Calibri',sz:10},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:{bottom:{style:'thin',color:{rgb:'808080'}}}};}
function bodyStyle(fill='FFFFFF'){return {fill:{fgColor:{rgb:fill}},font:{name:'Calibri',sz:10,color:{rgb:'1F1F1F'}},alignment:{vertical:'center'},border:{bottom:{style:'thin',color:{rgb:'D9E1D5'}}}};}
function styleSheet(ws,rows,cols){for(let r=0;r<rows;r++)for(let c=0;c<cols;c++){const ref=XLSX.utils.encode_cell({r,c});if(!ws[ref])ws[ref]={t:'z',v:null};ws[ref].s=r===0?headerStyle(YELLOW_HEADER_COLS.has(c)):bodyStyle(c===14?'F4CCCC':(c===20||c===21)?'FFF2CC':r%2?'FFFFFF':'F7F9F6');}}

function saveBlob(blob,filename){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);}
async function writeWorkbookWithFrozenRows(wb,filename,sheetIndexes,options={}){
  if(typeof JSZip==='undefined')throw new Error('No se cargó el componente local de exportación ZIP.');
  const bytes=XLSX.write(wb,{bookType:'xlsx',type:'array'}),zip=await JSZip.loadAsync(bytes);
  for(const index of sheetIndexes){
    const path=`xl/worksheets/sheet${index}.xml`,part=zip.file(path);
    if(!part)continue;
    let xml=await part.async('string');
    const frozen='<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>';
    if(/<sheetViews>[\s\S]*?<\/sheetViews>/.test(xml))xml=xml.replace(/<sheetViews>[\s\S]*?<\/sheetViews>/,frozen);
    else xml=xml.replace(/(<worksheet[^>]*>)/,`$1${frozen}`);
    if(options.conditionalFormatting&&options.conditionalFormatting.sheetIndex===index){
      const lastRow=options.conditionalFormatting.lastRow||100;
      const cfXml=`<conditionalFormatting sqref="Q2:Q${lastRow}"><cfRule type="iconSet" priority="1"><iconSet iconSet="3TrafficLights1" showValue="0"><cfvo type="num" val="0"/><cfvo type="num" val="1"/><cfvo type="num" val="2"/></iconSet></cfRule></conditionalFormatting><conditionalFormatting sqref="S2:S${lastRow}"><cfRule type="iconSet" priority="2"><iconSet iconSet="3TrafficLights1" showValue="0"><cfvo type="num" val="0"/><cfvo type="num" val="1"/><cfvo type="num" val="2"/></iconSet></cfRule></conditionalFormatting>`;
      if(xml.includes('<ignoredErrors'))xml=xml.replace('<ignoredErrors',cfXml+'<ignoredErrors');
      else if(xml.includes('<pageMargins'))xml=xml.replace('<pageMargins',cfXml+'<pageMargins');
      else xml=xml.replace('</worksheet>',cfXml+'</worksheet>');
    }
    zip.file(path,xml);
  }
  saveBlob(await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:6}}),filename);
}
async function downloadMaster(){
  if(!state.result)return;try{
    const wb=workbookBase(),aoa=[MASTER_HEADERS,...state.result.records.map(record=>MASTER_FIELDS.map(field=>field==='fecha'?excelDateSerial(record.fecha):field==='documento'?sapRef(record.documento,record.tipo_doc):record[field]??null))],ws=XLSX.utils.aoa_to_sheet(aoa);
    for(let r=1;r<aoa.length;r++){const row=r+1,record=state.result.records[r-1],ref=`W${row}`;ws[ref]={t:'s',v:record.estatus,f:statusFormula(row),s:bodyStyle('FFFFFF')};ws[`Q${row}`]={t:'n',v:record.lib_solped==='2'?2:0,f:`IFERROR(IF(VALUE(P${row})=2,2,0),0)`,s:bodyStyle(row%2?'FFFFFF':'F7F9F6')};ws[`S${row}`]={t:'n',v:record.lib_os==='G'?2:record.lib_os==='B'?0:1,f:`IF(R${row}="G",2,IF(R${row}="B",0,1))`,s:bodyStyle(row%2?'FFFFFF':'F7F9F6')};for(const col of ['G','H','K','N','O','U']){const cell=ws[`${col}${row}`];if(cell){cell.t='s';cell.z='@';}}const d=ws[`B${row}`];if(d)d.z='dd/mm/yyyy';const amount=ws[`L${row}`];if(amount)amount.z='"S/ "#,##0.00';}
    styleSheet(ws,aoa.length,23);ws['!cols']=[8,12,7,15,16,24,14,14,38,12,14,16,38,14,14,14,17,12,16,46,16,28,18].map(wch=>({wch}));ws['!rows']=[{hpt:30}];ws['!autofilter']={ref:`A1:W${aoa.length}`};ws['!freeze']={xSplit:0,ySplit:1,topLeftCell:'A2',activePane:'bottomLeft'};ws['!cols'][12].hidden=true;ws['!cols'][15].hidden=true;ws['!cols'][17].hidden=true;XLSX.utils.book_append_sheet(wb,ws,'BD');
    const summaryRows=[['Control de órdenes de servicio','Todas las semanas cargadas'],[],['Indicador','Valor'],['Registros',state.result.records.length],['Importe total',state.result.records.reduce((s,r)=>s+(r.valorizado||0),0)],['Pendientes de MIRO',state.result.records.filter(r=>!r.miro).length],['Observados',state.result.records.filter(r=>r.estatus==='OBSERVADO').length],['Cerrados',state.result.records.filter(r=>r.estatus==='CERRADO').length],['Incidencias',state.result.incidents.length],[],['Provincia','Registros','RXH','FACT','Importe','Pendiente MIRO','Observado','Cerrado']];
    for(const province of [...new Set(state.result.records.map(r=>r.provincia))].sort()){const rows=state.result.records.filter(r=>r.provincia===province);summaryRows.push([province,rows.length,rows.filter(r=>r.tipo_doc==='RXH').length,rows.filter(r=>r.tipo_doc==='FACT').length,rows.reduce((s,r)=>s+(r.valorizado||0),0),rows.filter(r=>!r.miro).length,rows.filter(r=>r.estatus==='OBSERVADO').length,rows.filter(r=>r.estatus==='CERRADO').length]);}
    const sws=XLSX.utils.aoa_to_sheet(summaryRows);sws['!cols']=[28,18,12,12,18,18,15,15].map(wch=>({wch}));['A3','B3'].forEach(ref=>sws[ref].s=headerStyle(false));for(let c=0;c<8;c++){const ref=XLSX.utils.encode_cell({r:10,c});if(sws[ref])sws[ref].s=headerStyle(false);}XLSX.utils.book_append_sheet(wb,sws,'RESUMEN');
    const ia=[INCIDENT_HEADERS,...state.result.incidents.map(i=>[i.week,i.province,i.document,i.source,i.field,i.severity,i.reason])],iws=XLSX.utils.aoa_to_sheet(ia);for(let c=0;c<7;c++)iws[XLSX.utils.encode_cell({r:0,c})].s=headerStyle(false);iws['!cols']=[9,18,18,28,22,16,90].map(wch=>({wch}));iws['!autofilter']={ref:`A1:G${ia.length}`};XLSX.utils.book_append_sheet(wb,iws,'INCIDENCIAS');
    await writeWorkbookWithFrozenRows(wb,'Control_Ordenes_Servicio_Todas_las_semanas_Generado.xlsx',[1],{conditionalFormatting:{sheetIndex:1,lastRow:aoa.length}});
  }catch(error){console.error(error);alert(`No se pudo generar el maestro: ${error.message}`);}
}

const SAP_COLS=[['id','ID CABC(1)',8],['fecha','Fecha Doc',13],['texto_cab','Texto Cab Doc',18],['referencia','REFERENCIA',24],['id_acr','ID POSC(2)Acreedor',12],['ruc','RUC',14],['cond','Cond Pago',10],['impto','Ind Impto',10],['tipo_detr','Tipo Detr',10],['ind_detr','Ind Detra',10],['asig','Asignación',14],['texto','Texto Posición',42],['id_gasto','ID POSC(2) Cuenta GASTO',10],['cuenta','Cuenta Gasto',13],['detalle','DETALLE',20],['centro','Centro Costo',14],['importe','Importe MD',14],['impto2','Ind Impto',10],['asig2','Asignación',14],['texto2','Texto Posición',42],['oc','N° OC',14]];
function sapText(motivo,week){let t=normalizeText(motivo).replace(/\bSERV\.\s*/g,'SERVICIO DE ').replace(/\bTRANSP\.\s*/g,'TRANSPORTE ').replace(/\bMTTO\b/g,'MANTENIMIENTO').replace(/\bAYUDANTE\s+REP\.\s*/g,'AYUDANTE DE REPARTO ').replace(/\bREP\.\s*/g,'REPARTO ').replace(/\s*SEMANA\s*\d+\s*$/,'').replace(/\s+/g,' ').trim();return `${t} SEMANA ${week}`;}
function sapRef(document,type){const key=normDoc(document),m=key.match(/^([A-Z]\d{3})(\d+)$/),prefix=type==='FACT'?'01-0':'02-0';return !key?'':m?`${prefix}${m[1]}-${m[2].padStart(7,'0')}`:prefix+cleanId(document);}
function findPdfInfo(record,kind){return state.pdfInfo.find(info=>normProvince(info.province)===normProvince(record.provincia)&&(kind==='DOC'?info.document===normDoc(record.documento)&&['FACT','RXH'].includes(info.kind):info.type===kind&&info.number===cleanId(record[kind.toLowerCase()])));}
function sapRow(record,index){const doc=findPdfInfo(record,'DOC'),ocPdf=findPdfInfo(record,'OC'),isFact=record.tipo_doc==='FACT',amount=record.valorizado,date=(doc&&doc.date)||(ocPdf&&ocPdf.date)||formatDate(record.fecha,'.'),text=sapText(record.motivo,record.sem),tax=isFact?'C1':'C0';return {id:index,fecha:date,texto_cab:null,referencia:sapRef(record.documento,record.tipo_doc),id_acr:2,ruc:record.ruc,cond:'CP02',impto:tax,tipo_detr:null,ind_detr:null,asig:null,texto:text,id_gasto:2,cuenta:null,detalle:null,centro:null,importe:amount,impto2:tax,asig2:null,texto2:text,oc:record.oc};}
async function downloadSap(){
  if(!state.result)return;try{
    const wb=workbookBase(),groups=new Map();const source=state.result.selectedRecords.length?state.result.selectedRecords:state.result.records.filter(r=>r._physicalSupported);
    source.forEach(record=>{const key=record.provincia;if(!groups.has(key))groups.set(key,[]);groups.get(key).push(record);});
    for(const [province,records] of groups){const rows=records.map((record,index)=>sapRow(record,index+1)),aoa=[SAP_COLS.map(c=>c[1]),...rows.map(row=>SAP_COLS.map(c=>row[c[0]]??null)),SAP_COLS.map(c=>c[0]==='id'?'TOTAL':c[0]==='importe'?rows.reduce((s,r)=>s+r.importe,0):null)],ws=XLSX.utils.aoa_to_sheet(aoa);for(let r=0;r<aoa.length;r++)for(let c=0;c<SAP_COLS.length;c++){const ref=XLSX.utils.encode_cell({r,c});if(!ws[ref])ws[ref]={t:'z',v:null};ws[ref].s=r===0?headerStyle(false):r===aoa.length-1?headerStyle(false):bodyStyle(c===16?'FFF2CC':c===20?'DEEAF1':r%2?'FFFFFF':'F2F2F2');if(c===16&&r>0){ws[ref].z='"S/ "#,##0.00';ws[ref].t='n';}}ws['!cols']=SAP_COLS.map(c=>({wch:c[2]}));ws['!autofilter']={ref:`A1:U${aoa.length-1}`};ws['!freeze']={xSplit:0,ySplit:1,topLeftCell:'A2',activePane:'bottomLeft'};XLSX.utils.book_append_sheet(wb,ws,`${province.titleCase?province.titleCase():province} Todas las semanas`.slice(0,31));}
    await writeWorkbookWithFrozenRows(wb,'Masivo Reporte.xlsx',Array.from({length:groups.size},(_,index)=>index+1));
  }catch(error){console.error(error);alert(`No se pudo generar el masivo SAP: ${error.message}`);}
}

async function applyManualToCurrentResult(){
  if(!state.result || !state.previousFile) return;
  try {
    const manual = await parseManual(state.previousFile);
    let updatedCount = 0;
    for(const record of state.result.records){
      const values = manual.get(groupKey(record));
      if(values){
        if(values.miro && values.miro !== record.miro){
          record.miro = values.miro;
          saveMiroToCache(groupKey(record), values.miro);
          if(record.hes) saveMiroToCache('HES:'+cleanId(record.hes), values.miro);
          updatedCount++;
        }
        if(values.obs_miro) record.obs_miro = values.obs_miro;
        if(values.observaciones) record._manualObservation = values.observaciones;
        finalizeRecord(record);
      }
    }
    renderResults();
    alert(`Se importaron ${updatedCount} códigos MIRO del maestro editado.`);
  } catch(err){
    console.error(err);
    alert('Error al leer el maestro editado: ' + err.message);
  }
}

async function downloadRenamedZip(){
  if(!state.result || !state.result.records.length){
    alert('Primero debes procesar los expedientes PDF.');
    return;
  }
  const scope = $('rename-scope')?.value || 'hes';
  const readyRecords = state.result.records.filter(r => /^\d{10}$/.test(cleanId(r.miro)));
  if(!readyRecords.length){
    alert('No hay registros con código MIRO de 10 dígitos.\nPuedes ingresar los códigos MIRO en la tabla de vista previa o hacer clic en "Importar maestro editado".');
    return;
  }
  const btn = $('btn-download-renamed');
  const originalText = btn.innerHTML;
  btn.disabled = true;
  btn.innerHTML = '<span class="spin"></span> Generando ZIP…';
  try {
    const zip = new JSZip();
    const pathToMiro = new Map();
    for(const r of readyRecords){
      const miro = cleanId(r.miro);
      if(r._source) pathToMiro.set(r._source, miro);
      if(r._evidence?.hes) r._evidence.hes.forEach(p => pathToMiro.set(p, miro));
      if(scope === 'todos'){
        if(r._evidence?.oc) r._evidence.oc.forEach(p => pathToMiro.set(p, miro));
        if(r._evidence?.documento) r._evidence.documento.forEach(p => pathToMiro.set(p, miro));
        if(r._evidence?.solped) r._evidence.solped.forEach(p => pathToMiro.set(p, miro));
      }
    }

    let renamedCount = 0;
    for(const [path, entry] of state.files.entries()){
      const miro = pathToMiro.get(path);
      let targetName = entry.file.name;
      if(miro){
        const baseName = targetName.replace(/^\d{10}\s*-\s*/, '');
        targetName = `${miro} - ${baseName}`;
        renamedCount++;
      }
      const parts = path.split('/');
      parts[parts.length - 1] = targetName;
      const zipPath = parts.join('/');
      const buffer = await entry.file.arrayBuffer();
      zip.file(zipPath, buffer);
    }

    const blob = await zip.generateAsync({type:'blob', compression:'DEFLATE', compressionOptions:{level:5}});
    saveBlob(blob, `Expedientes_Renombrados_MIRO_${readyRecords.length}_expedientes.zip`);
  } catch(err){
    console.error(err);
    alert('Error al generar el ZIP de expedientes: ' + err.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = originalText;
  }
}

async function renameLocalFiles(){
  if(!state.result || !state.result.records.length){
    alert('Primero debes procesar los expedientes PDF.');
    return;
  }
  const scope = $('rename-scope')?.value || 'hes';
  const readyRecords = state.result.records.filter(r => /^\d{10}$/.test(cleanId(r.miro)));
  if(!readyRecords.length){
    alert('No hay registros con código MIRO de 10 dígitos.\nIngresa los códigos MIRO en la tabla o importa el maestro editado.');
    return;
  }
  if(!window.showDirectoryPicker){
    alert('Tu navegador no permite renombrado directo en disco.\nUsa el botón "Descargar ZIP con nombres MIRO" o el archivo "Renombrar_Documentos_MIRO.bat".');
    return;
  }
  try {
    const dirHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
    const btn = $('btn-rename-local');
    const originalText = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = '<span class="spin"></span> Renombrando…';
    
    const pathToMiro = new Map();
    for(const r of readyRecords){
      const miro = cleanId(r.miro);
      if(r._source) pathToMiro.set(r._source, miro);
      if(r._evidence?.hes) r._evidence.hes.forEach(p => pathToMiro.set(p, miro));
      if(scope === 'todos'){
        if(r._evidence?.oc) r._evidence.oc.forEach(p => pathToMiro.set(p, miro));
        if(r._evidence?.documento) r._evidence.documento.forEach(p => pathToMiro.set(p, miro));
        if(r._evidence?.solped) r._evidence.solped.forEach(p => pathToMiro.set(p, miro));
      }
    }

    let renamed = 0;
    async function walk(handle, currentRelPath=''){
      for await(const entry of handle.values()){
        const relPath = currentRelPath ? `${currentRelPath}/${entry.name}` : entry.name;
        if(entry.kind === 'directory'){
          await walk(entry, relPath);
        } else if(entry.kind === 'file' && /\.pdf$/i.test(entry.name)){
          let miro = null;
          for(const [storedPath, m] of pathToMiro.entries()){
            if(storedPath.endsWith(entry.name) || relPath.endsWith(storedPath)){
              miro = m;
              break;
            }
          }
          if(miro && !entry.name.startsWith(`${miro} - `)){
            const cleanBase = entry.name.replace(/^\d{10}\s*-\s*/, '');
            const newName = `${miro} - ${cleanBase}`;
            if(entry.move){
              await entry.move(newName);
              renamed++;
            } else {
              const file = await entry.getFile();
              const newHandle = await handle.getFileHandle(newName, { create: true });
              const writable = await newHandle.createWritable();
              await writable.write(await file.arrayBuffer());
              await writable.close();
              await handle.removeEntry(entry.name);
              renamed++;
            }
          }
        }
      }
    }
    await walk(dirHandle);
    btn.disabled = false;
    btn.innerHTML = originalText;
    alert(`¡Completado con éxito! Se renombraron ${renamed} archivos en la carpeta seleccionada.`);
  } catch(err){
    if(err.name === 'AbortError') return;
    console.error(err);
    alert('No se pudo renombrar en disco: ' + err.message + '\nPuedes usar "Descargar ZIP con nombres MIRO" en su lugar.');
  }
}

function escapeHtml(value){return String(value??'').replace(/[&<>'"]/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]));}
