/* Asistente de importación desde un ERP: elegir hojas y columnas, rellenar lo que el ERP no trae y revisar */
(function(){
const M=window.EIOSMap, RW=window.EIOSRaw, IC=window.EIOSImport, $=s=>document.querySelector(s);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const W={sheets:{},map:{},erp:'',P:{},cp:{},files:[]};
const SUPV=[['vig','Vigencia de la previsión de consumo (días)',90],['periodo','Días de ventas para calcular el consumo',92],['obj','Cobertura objetivo (días)',30],['mult','Múltiplo de compra',1],['grave','Retraso de cobro que se considera grave (días)',15],['desvio','Retraso de entrega que cuenta como incidencia (días)',3],['pPerd','Parte de la venta que se pierde sin stock (0 a 1)',0.5],['recargo','Recargo de un pedido urgente (0 a 1)',0.05]];
const LSg=()=>{ try{ return JSON.parse(localStorage.getItem('eios-mapeos')||'{}'); }catch(e){ return {}; } };
async function mapeos(){ if(ST.db){ try{ const s=await ST.db.collection('mapeos').get(); const o={}; s.docs.forEach(d=>o[d.id]=d.data()); return o; }catch(e){} } return LSg(); }
async function guardaMapeo(id,doc){ if(ST.db){ try{ await ST.db.collection('mapeos').doc(id).set(doc); return true; }catch(e){ return false; } } try{ const o=LSg(); o[id]=doc; localStorage.setItem('eios-mapeos',JSON.stringify(o)); return true; }catch(e){ return false; } }
const slug=s=>M.norm(s).replace(/ /g,'-')||'erp';
function emsg(c,t){ $('#emsg').innerHTML=`<div class="banner ${c==='ok'?'ok':c==='warn'?'warn':'bad'} small">${esc(t)}</div>`; }
async function leerFicheros(files){ emsg('warn','Leyendo…'); W.sheets={}; W.files=[];
  try{ for(const f of files){ const wb=XLSX.read(await f.arrayBuffer(),{type:'array'}); W.files.push(f.name);
      for(const n of wb.SheetNames){ const rows=XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,defval:null}); if(rows.length>1) W.sheets[(files.length>1?f.name.replace(/\.[^.]+$/,'')+' › ':'')+n]=rows; } }
    if(!Object.keys(W.sheets).length) throw new Error('No he encontrado ninguna hoja con datos en el archivo.');
    W.map=M.suggestAll(W.sheets); W.erp=W.erp||W.files[0].replace(/\.[^.]+$/,''); await aplicaGuardado(); emsg('ok','Leídas '+Object.keys(W.sheets).length+' hojas. Te he propuesto las columnas; comprueba cada bloque.'); render();
  }catch(e){ emsg('bad',(e&&e.message)||'No he podido leer el archivo.'); } }
async function aplicaGuardado(){ const all=await mapeos(), g=all[slug(W.erp)]; if(!g) return; let n=0;
  for(const t of M.TABLES){ const gt=g.tablas&&g.tablas[t.id]; if(gt===null&&!t.req){ W.map[t.id]=null; continue; } if(!gt||!W.sheets[gt.sheet]) continue; const rows=W.sheets[gt.sheet], hd=gt.hdr||0, hs=(rows[hd]||[]).map(M.norm), f={};
    for(const fd of t.f){ const q=(gt.f||{})[fd.k]||{h:'',v:''}; const j=q.h?hs.indexOf(M.norm(q.h)):-1; f[fd.k]={c:j,v:q.v||''}; }
    W.map[t.id]={sheet:gt.sheet,hdr:hd,f,completo:t.f.filter(x=>x.req).every(x=>f[x.k].c>=0||f[x.k].v!=='')}; n++; }
  if(g.params){ Object.assign(W.P,g.params); }
  if(n) emsg('ok','Mapeo guardado de «'+W.erp+'» aplicado a '+n+' tablas.'); }
const optSheets=(sel,opcional)=>(opcional?`<option value="">— no tengo esta tabla —</option>`:'')+Object.keys(W.sheets).map(n=>`<option value="${esc(n)}" ${n===sel?'selected':''}>${esc(n)}</option>`).join('');
function estado(t){ const m=W.map[t.id]; if(!m) return t.req?['bad','falta']:['muted','sin tabla']; const f=t.f.filter(x=>x.req&&!((m.f[x.k]||{}).c>=0)&&!((m.f[x.k]||{}).v!==''&&(m.f[x.k]||{}).v!=null)); return f.length?['bad','faltan: '+f.map(x=>x.l).join(', ')]:['ok','listo']; }
function tablaHTML(t){ const m=W.map[t.id], st=estado(t); let h=`<details class="tb" data-t="${t.id}" ${st[0]==='bad'?'open':''}><summary><b>${esc(t.l)}</b> ${t.req?'':'<span class="muted small">(opcional)</span>'} <span class="chip ${st[0]}" data-chip="${t.id}">${esc(st[1])}</span></summary><div class="stack" style="margin-top:8px">
    <label class="small muted">Hoja del Excel<br><select data-sh="${t.id}">${optSheets(m&&m.sheet,!t.req)}${t.req&&!m?'<option value="" selected>— elige una hoja —</option>':''}</select></label>`;
  if(m){ const rows=W.sheets[m.sheet], hs=(rows[m.hdr]||[]);
    h+=`<label class="small muted">Fila de encabezados (número de fila del Excel)<br><input type="number" min="1" max="30" value="${m.hdr+1}" data-hd="${t.id}" style="width:90px"></label>
    <div class="tablewrap"><table style="min-width:520px"><thead><tr><th>Dato de EIOS</th><th>Columna de tu Excel</th><th>Valor fijo (si no hay columna)</th></tr></thead><tbody>${t.f.map(f=>{ const q=m.f[f.k]||{c:-1,v:''}; return `<tr><td>${esc(f.l)}${f.req?' <b class="ko-t">*</b>':''}</td><td><select data-c="${t.id}.${f.k}"><option value="-1">— ninguna —</option>${hs.map((x,i)=>`<option value="${i}" ${q.c===i?'selected':''}>${esc(x==null?'(columna '+(i+1)+')':x)}</option>`).join('')}</select></td><td><input type="text" data-v="${t.id}.${f.k}" value="${esc(q.v)}" style="width:100%;min-width:80px" ${q.c>=0?'disabled':''}></td></tr>`; }).join('')}</tbody></table></div>
    <div class="small muted">Vista previa (3 primeras filas):</div><div class="tablewrap"><table style="min-width:520px"><tbody>${rows.slice(m.hdr+1,m.hdr+4).map(r=>`<tr>${hs.map((_,i)=>`<td>${esc(r[i]==null?'':r[i])}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`; }
  return h+'</div></details>'; }
function paramsHTML(){ const P=W.P, v=(k,d)=>P[k]!=null&&P[k]!==''?P[k]:(d==null?'':d), S=P.supuestos||{}; const sv=(k,d)=>S[k]!=null&&S[k]!==''?S[k]:d;
  return `<h3>Datos de la empresa y parámetros</h3><div class="grid2">
   <label class="small muted">Empresa<br><input type="text" id="p_empresa" value="${esc(v('empresa'))}"></label>
   <label class="small muted">Fecha de los datos (corte) <b class="ko-t">*</b><br><input type="date" id="p_corte" value="${esc(v('corte'))}"></label>
   <label class="small muted">Saldo mínimo de tesorería (EUR) <b class="ko-t">*</b><br><input type="number" id="p_saldoMin" value="${esc(v('saldoMin'))}"></label>
   <label class="small muted">Margen ámbar sobre el mínimo (EUR) <b class="ko-t">*</b><br><input type="number" id="p_margenAmbar" value="${esc(v('margenAmbar'))}"></label>
   <label class="small muted">IVA de las compras (%)<br><input type="number" id="p_iva" value="${esc(v('iva',21))}"></label>
   <label class="small muted">Horizonte de tesorería (fecha; vacío = corte + 60 días)<br><input type="date" id="p_horizonte" value="${esc(v('horizonte'))}"></label></div>
  <details><summary><b>Supuestos</b> <span class="chip warn">SUPUESTO</span> <span class="muted small">cosas que tu ERP no suele traer; puedes cambiarlas</span></summary><div class="grid2" style="margin-top:8px">${SUPV.map(([k,l,d])=>`<label class="small muted">${esc(l)}<br><input type="number" step="any" id="s_${k}" value="${esc(sv(k,d))}"></label>`).join('')}</div></details>`; }
function condHTML(){ if(W.map.condProv) return ''; const c=cur(); if(!c) return ''; const cp=W.cp;
  return `<h3>Condiciones de cada proveedor <span class="chip warn">SUPUESTO</span></h3><p class="small muted" style="margin:0">Tu Excel no trae condiciones de proveedor. Escríbelas aquí: el plazo de pago y el plazo de entrega son obligatorios; lo demás en blanco cuenta como «sin descuento / sin portes / sin mínimo».</p>
  <div class="tablewrap"><table style="min-width:760px"><thead><tr><th>Proveedor</th><th>Plazo pago (d) *</th><th>Entrega (d) *</th><th>Forma de pago</th><th>Cuotas máx.</th><th>Dto. volumen %</th><th>Umbral EUR</th><th>Portes EUR</th><th>Franco EUR</th><th>Mínimo EUR</th></tr></thead><tbody>${c.canon.proveedores.map(p=>{ const q=cp[p.cod]||{}; const inp=(k,w)=>`<td><input type="${k==='forma'?'text':'number'}" step="any" data-cp="${esc(p.cod)}.${k}" value="${esc(q[k]==null?'':q[k])}" style="width:${w||70}px"></td>`; return `<tr><td>${esc(p.nombre)}</td>${inp('plazo')}${inp('entrega')}${inp('forma',90)}${inp('cuotas',60)}${inp('dtoVol',60)}${inp('umbral')}${inp('portes')}${inp('franco')}${inp('minimo')}</tr>`; }).join('')}</tbody></table></div>`; }
function leeParams(){ const g=id=>{ const e=$('#'+id); return e?e.value:''; }, n=id=>{ const v=g(id); return v===''?null:parseFloat(v); };
  if(!$('#p_corte')) return; W.P.empresa=g('p_empresa'); W.P.corte=g('p_corte'); W.P.saldoMin=n('p_saldoMin'); W.P.margenAmbar=n('p_margenAmbar'); W.P.iva=n('p_iva'); W.P.horizonte=g('p_horizonte')||null; W.P.supuestos={}; for(const [k] of SUPV){ const v=n('s_'+k); if(v!=null) W.P.supuestos[k]=v; }
  document.querySelectorAll('[data-cp]').forEach(e=>{ const [c,k]=e.dataset.cp.split('.'); (W.cp[c]=W.cp[c]||{})[k]=k==='forma'?e.value:(e.value===''?null:parseFloat(e.value)); }); }
function cur(){ try{ const P=Object.assign({},W.P,{condProv:W.cp}); return M.buildCanon(W.sheets,W.map,P); }catch(e){ return null; } }
function render(){ const el=$('#ewiz'); el.hidden=false;
  el.innerHTML=`<label class="small muted">Nombre de tu ERP (sirve para recordar este mapeo)<br><input type="text" id="erpn" value="${esc(W.erp)}" style="width:100%;max-width:360px"></label>
   <div class="stack" id="etab">${M.TABLES.map(tablaHTML).join('')}</div><div class="stack" id="eparams">${paramsHTML()}<div id="econd" class="stack">${condHTML()}</div></div>
   <div id="eerr"></div><div style="display:flex;gap:10px;flex-wrap:wrap"><button class="btn primary" id="ecalc" type="button">Calcular y revisar</button><button class="btn" id="esave" type="button">Guardar este mapeo</button></div>`;
  $('#erpn').onchange=async e=>{ W.erp=e.target.value; await aplicaGuardado(); render(); };
  el.querySelectorAll('[data-sh]').forEach(s=>s.onchange=()=>{ const id=s.dataset.sh, t=M.TABLES.find(x=>x.id===id); if(!s.value){ W.map[id]=null; } else { const hd=M.headerRow(W.sheets[s.value]), c=M.suggestCols(W.sheets[s.value][hd]||[],t); W.map[id]={sheet:s.value,hdr:hd,f:Object.fromEntries(Object.entries(c).map(([k,v])=>[k,{c:v,v:''}]))}; } leeParams(); refrescaTabla(id); });
  el.querySelectorAll('[data-hd]').forEach(s=>s.onchange=()=>{ const id=s.dataset.hd, t=M.TABLES.find(x=>x.id===id), m=W.map[id]; m.hdr=Math.max(0,(parseInt(s.value)||1)-1); const c=M.suggestCols(W.sheets[m.sheet][m.hdr]||[],t); m.f=Object.fromEntries(Object.entries(c).map(([k,v])=>[k,{c:v,v:''}])); refrescaTabla(id); });
  enlaza(el); }
function refrescaTabla(id){ leeParams(); const t=M.TABLES.find(x=>x.id===id), d=document.querySelector(`details[data-t="${id}"]`), tmp=document.createElement('div'); tmp.innerHTML=tablaHTML(t); const nd=tmp.firstChild; nd.open=true; d.replaceWith(nd); render2(); }
function render2(){ const k=$('#econd'); if(k){ k.innerHTML=condHTML(); } enlaza(document); }
function enlaza(el){ el.querySelectorAll('[data-c]').forEach(s=>s.onchange=()=>{ const [t,k]=s.dataset.c.split('.'), m=W.map[t]; m.f[k]=Object.assign(m.f[k]||{v:''},{c:parseInt(s.value)}); const inp=el.querySelector(`[data-v="${t}.${k}"]`); if(inp) inp.disabled=m.f[k].c>=0; chip(t); });
  el.querySelectorAll('[data-v]').forEach(s=>s.oninput=()=>{ const [t,k]=s.dataset.v.split('.'), m=W.map[t]; m.f[k]=Object.assign(m.f[k]||{c:-1},{v:s.value}); chip(t); });
  el.querySelectorAll('[data-sh],[data-hd]').forEach(()=>{}); $('#ecalc').onclick=calcular; $('#esave').onclick=guardar; }
function chip(id){ const t=M.TABLES.find(x=>x.id===id), st=estado(t), c=document.querySelector(`[data-chip="${id}"]`); if(c){ c.className='chip '+st[0]; c.textContent=st[1]; } }
async function guardar(){ leeParams(); W.erp=$('#erpn').value||W.erp; const tablas={}; for(const t of M.TABLES){ const m=W.map[t.id]; if(!m){ tablas[t.id]=null; continue; } const hs=(W.sheets[m.sheet][m.hdr]||[]); tablas[t.id]={sheet:m.sheet,hdr:m.hdr,f:Object.fromEntries(t.f.map(f=>{ const q=m.f[f.k]||{c:-1,v:''}; return [f.k,{h:q.c>=0?String(hs[q.c]==null?'':hs[q.c]):'',v:q.v||''}]; }))}; }
  const ok=await guardaMapeo(slug(W.erp),{nombre:W.erp,fecha:new Date().toISOString(),tablas,params:{empresa:W.P.empresa,iva:W.P.iva,supuestos:W.P.supuestos,saldoMin:W.P.saldoMin,margenAmbar:W.P.margenAmbar}}); emsg(ok?'ok':'bad',ok?'Mapeo de «'+W.erp+'» guardado. La próxima vez que subas un Excel de este ERP se aplicará solo.':'No he podido guardar el mapeo.'); }
function calcular(){ leeParams(); const errs=[]; const e=$('#eerr'); e.innerHTML='';
  for(const t of M.TABLES){ const st=estado(t); if(st[0]==='bad') errs.push(t.l+': '+st[1]); }
  const P=W.P; if(!P.corte) errs.push('Indica la fecha de los datos (corte).'); if(P.saldoMin==null) errs.push('Indica el saldo mínimo de tesorería.'); if(P.margenAmbar==null) errs.push('Indica el margen ámbar.');
  const r=errs.length?null:cur(); if(r){ const sp=r.canon.articulos.filter(a=>!a.prov); if(sp.length) errs.push(sp.length+' artículo(s) sin proveedor (por ejemplo '+sp.slice(0,3).map(a=>a.sku).join(', ')+'). Indica la columna del proveedor o añade la tabla «Proveedor de cada artículo».'); }
  if(r&&!errs.length&&!W.map.condProv){ const f=r.canon.condProv.filter(x=>x.plazo==null||x.entrega==null); if(f.length) errs.push('Faltan el plazo de pago o el de entrega de '+f.length+' proveedor(es): '+f.slice(0,5).map(x=>x.cod).join(', ')+(f.length>5?'…':'')+'.'); }
  if(errs.length){ e.innerHTML=`<div class="banner bad small"><b>Falta esto antes de calcular:</b><ul style="margin:6px 0 0;padding-left:18px">${errs.map(x=>`<li>${esc(x)}</li>`).join('')}</ul></div>`; return; }
  try{ const c=r.canon, wb=RW.toTemplate(XLSX,c), warns=[], o=IC.buildD(XLSX,wb,w=>warns.push(w)), ex=r.avisos.slice();
    const sup=Object.assign({},RW.SUP,c.supuestos);
    ex.push('Calculados por la app con tus datos: consumo diario ('+sup.periodo+' días de ventas), entregas y fiabilidad de proveedores, fiabilidad de clientes y datos de rotura.');
    if(!c.compras.some(x=>x.pedido)) ex.push('Tu Excel no trae la fecha del pedido: no se puede medir el retraso de las entregas y los proveedores quedan «sin historial de entregas».');
    if(!c.incidencias.length) ex.push('Sin datos de incidencias de proveedor: se cuentan como incidencias las entregas con retraso de '+sup.desvio+' días o más; no se conocen las abiertas ni los tiempos de respuesta.');
    if(!c.cobros.length) ex.push('Sin historial de cobros reales: no se puede medir la fiabilidad de los clientes y aparecerán «sin historial».');
    ex.push('Supuestos usados (editables arriba): cobertura objetivo '+sup.obj+' días, múltiplo '+sup.mult+', pérdida de venta sin stock '+Math.round(sup.pPerd*100)+' %, recargo urgente '+Math.round(sup.recargo*100)+' %, vigencia '+sup.vig+' días.');
    mostrar(o,(W.erp||'ERP')+' (asistente)',warns,ex); $('#prev').scrollIntoView({behavior:'smooth'});
  }catch(err){ e.innerHTML=`<div class="banner bad small">${esc((err&&err.message)||'No he podido calcular.')}</div>`; } }
const inp=$('#efile'); if(inp) inp.onchange=e=>{ if(e.target.files.length) leerFicheros([...e.target.files]); };
window.EIOSWiz={W,leerFicheros,render,calcular};
})();
