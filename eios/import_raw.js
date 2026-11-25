/* Datos "crudos" de un ERP -> libro con el diseño de la plantilla EIOS (con los datos derivados ya calculados) -> buildD */
(function(root){
const DAY=864e5;
const iso=v=>{ if(v==null||v==='') return null;
  if(typeof v==='number') return new Date(Math.round((v-25569)*DAY)).toISOString().slice(0,10);
  if(v instanceof Date) return new Date(Math.round(v.getTime()/DAY)*DAY).toISOString().slice(0,10);
  const s=String(v).trim(); let m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/); if(m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  m=s.match(/^(\d{4})-(\d{2})-(\d{2})/); return m?m[0]:null; };
const T=v=>v==null?'':String(v).trim();
const N=v=>{ if(v==null||v==='') return 0; if(typeof v==='number') return v; let t=String(v).trim().replace(/\s/g,''); if(/,/.test(t)&&/\./.test(t)) t=t.replace(/\./g,'').replace(',','.'); else t=t.replace(',','.'); const n=parseFloat(t); return isNaN(n)?0:n; };
const ser=d=>d?Date.parse(d)/DAY+25569:null;
const dd=(a,b)=>Math.round((Date.parse(a)-Date.parse(b))/DAY);
const addD=(d,n)=>new Date(Date.parse(d)+n*DAY).toISOString().slice(0,10);
const mean=a=>a.reduce((x,y)=>x+y,0)/a.length;
const SUP={pPerd:0.5,recargo:0.05,obj:30,mult:1,vig:90,grave:15,desvio:3,periodo:92,horizonte:60};
/* ---- canon -> libro plantilla ---- */
function toTemplate(XLSX,c){
  const S=Object.assign({},SUP,c.supuestos||{}), corte=c.corte, wb=XLSX.utils.book_new();
  const add=(n,rows)=>XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),n);
  const cp={}; for(const x of c.condProv||[]) cp[x.cod]=x;
  const art=c.articulos||[], arts={}; art.forEach(a=>arts[a.sku]=a);
  add('016_Parametros_App',[['Parámetro','Valor','Unidad','Estado'],
    ['Saldo mínimo de seguridad de tesorería',c.saldoMin,'EUR',''],['Margen ámbar sobre el mínimo',c.margenAmbar,'EUR',''],['IVA de compras',c.iva,'%',''],
    ['Fecha de corte',ser(corte),'fecha',''],['Horizonte de la previsión de tesorería',ser(c.horizonte||addD(corte,S.horizonte)),'fecha',''],['Vigencia de la previsión de consumo',S.vig,'días',''],
    ['Retraso de cobro considerado grave',S.grave,'días',''],['Desvío de entrega que genera incidencia',S.desvio,'días',''],['Empresa',c.empresa||'','','']]);
  add('001_Proveedores',[['Proveedor','Cod.Proveedor']].concat((c.proveedores||[]).map(p=>[p.nombre,p.cod])));
  add('002_Clientes',[['Cod.Cliente','Razón Social']].concat((c.clientes||[]).map(p=>[p.cod,p.nombre])));
  const pn={}; for(const p of c.proveedores||[]) pn[p.cod]=p.nombre;
  add('003_ Articulos',[['Sku','Descripción','Proveedor','Familia','Cod.Proveedor','Moneda','Unidad']].concat(art.map(a=>[a.sku,a.desc,pn[a.prov]||'',a.fam,a.prov,'EUR',a.uni])));
  add('004_Inventario',[['Sku','Descripción','Existencias físicas','Mínimo','Seguridad','Reservado','Unidad','Fecha']].concat((c.inventario||[]).map(i=>[i.sku,(arts[i.sku]||{}).desc,i.fisico,i.min,i.seg,i.reservado,(arts[i.sku]||{}).uni,ser(corte)])));
  add('008_Referencias_Precios',[['Sku','Descripción','Coste','PVP']].concat((c.precios||[]).map(i=>[i.sku,(arts[i.sku]||{}).desc,i.coste,i.pvp])));
  add('006_Compras_historicas',[['Numero de Factura','Fecha Factura','Cod.Proveedor','Sku','Cantidad','Unidad','Precio unitario neto de línea','Descuento (%)','Iva (%)','Importe Neto','Moneda','Descripción','Fecha entrega']].concat((c.compras||[]).map(x=>[x.fac,ser(x.fecha),x.prov,x.sku,x.cant,(arts[x.sku]||{}).uni,x.pu,0,c.iva*100,x.cant*x.pu,'EUR',(arts[x.sku]||{}).desc,ser(x.entrega)])));
  add('007_Ventas_historicas',[['Número de Factura','Fecha de Factura','Cod. Cliente','Sku','Cliente','Descripción','Cantidad','Unidad','Precio unitario neto','IVA (%)','Descuento (%)','Importe neto de línea']].concat((c.ventas||[]).map(x=>[x.fac,ser(x.fecha),x.cli,x.sku,'',(arts[x.sku]||{}).desc,x.cant,(arts[x.sku]||{}).uni,x.pu,0,0,x.imp!=null?x.imp:x.cant*x.pu])));
  add('010_Flujos_Tesoreria',[['Referencia flujo','Tipo de flujo','Importe','Moneda','Fecha de vencimiento','Contraparte']].concat((c.flujos||[]).map(x=>[x.ref,x.tipo,x.importe,'EUR',ser(x.fecha),x.who])));
  add('011_Balance',[['Código de cuenta','Descripción','Saldo']].concat((c.balance||[]).map(x=>[x.cuenta,x.desc,x.saldo])));
  add('012_Cartera_Cobros',[['Referencia cuota','Cod.Cliente','Número de factura','Número de cuota','Fecha de vencimiento','Importe','Moneda','Forma de pago o cobro','Importe pendiente al corte']].concat((c.cartera||[]).map(x=>[x.ref,x.cli,x.fac,x.cuota,ser(x.vto),x.importe,'EUR',x.forma,x.pend])));
  if((c.recepciones||[]).length) add('014_Recepciones_Pendientes',[['Referencia','SKU','Unidad','Cantidad pendiente','Fecha prevista']].concat(c.recepciones.map(x=>[x.ref,x.sku,(arts[x.sku]||{}).uni,x.cant,ser(x.fecha)])));
  else add('014_Recepciones_Pendientes',[['Referencia','SKU','Unidad','Cantidad pendiente','Fecha prevista']]);
  add('015_Condiciones_Proveedor',[[],[],[],['Cod.Proveedor','Proveedor','Plazo de pago (días)','Forma de pago','Cuotas máximas','Descuento por volumen (%)','Umbral de volumen (EUR netos)','Portes por pedido (EUR)','Franco de portes (EUR netos)','Pedido mínimo (EUR netos)','Plazo de entrega (días)']].concat((c.condProv||[]).map(x=>[x.cod,pn[x.cod]||'',x.plazo,x.forma,x.cuotas,x.dtoVol,x.umbral,x.portes,x.franco,x.minimo,x.entrega])));
  /* 013: referencia histórica y supuestos */
  const desde=c.periodo&&c.periodo.desde||addD(corte,-(S.periodo-1)), hasta=c.periodo&&c.periodo.hasta||corte, dias=dd(hasta,desde)+1;
  const vs={}; for(const v of c.ventas||[]){ if(!v.fecha||v.fecha<desde||v.fecha>hasta) continue; const o=vs[v.sku]=vs[v.sku]||{q:0,n:0}; o.q+=v.cant; o.n++; }
  const inv={}; for(const i of c.inventario||[]) inv[i.sku]=i;
  const decl=c.stock||{}, cons={};
  const b1=[], b2=[];
  for(const a of art){ const v=vs[a.sku]||{q:0,n:0}, i=inv[a.sku]||{fisico:0,reservado:0,min:0,seg:0}, med=v.n>0?v.q/dias:null, d=decl[a.sku]||{};
    const cs=d.consumo!=null&&d.consumo!==''?N(d.consumo):(med!=null?med:null); cons[a.sku]=cs;
    b1.push([a.sku,a.uni,ser(desde),ser(hasta),dias,v.q,v.n,med!=null?med:'Sin referencia',i.fisico-i.reservado,i.min,i.seg,cs>0?(i.fisico-i.reservado)/cs:'Sin consumo válido']);
    const cpa=cp[a.prov]||{}; b2.push([a.sku,a.uni,cs!=null?cs:'',ser(corte),ser(addD(corte,S.vig)),d.lead!=null&&d.lead!==''?N(d.lead):(cpa.entrega!=null?cpa.entrega:''),d.obj!=null&&d.obj!==''?N(d.obj):S.obj,d.mult!=null&&d.mult!==''?N(d.mult):S.mult]); }
  add('013_Planificacion_Stock',[['PLANIFICACIÓN DE STOCK'],[],['SKU','Unidad','Desde','Hasta','Días naturales','Ventas del periodo','Líneas de venta','Media por día','Disponible al corte','Mínimo','Seguridad','Cobertura teórica (días)']].concat(b1,[[],['SKU','Unidad','Consumo / día','Vigente desde','Vigente hasta','Entrega (días)','Objetivo (días)','Múltiplo compra']],b2));
  /* 018 y 020 proveedores */
  const e18=[], pv={};
  (c.compras||[]).forEach((x,k)=>{ if(!x.pedido||!x.entrega) return; const plz=(cp[x.prov]||{}).entrega; if(plz==null) return; const prom=addD(x.pedido,plz), des=dd(x.entrega,prom);
    e18.push([x.fac,ser(x.fecha),x.prov,x.sku,x.cant,ser(x.pedido),plz,ser(prom),ser(x.entrega),des,'',dd(x.entrega,x.pedido)]); (pv[x.prov]=pv[x.prov]||[]).push({des,real:dd(x.entrega,x.pedido)}); });
  add('018_Entregas_Compras',[['ENTREGAS DE COMPRAS'],[],[],['Número de factura','Fecha factura','Cod.Proveedor','Sku','Cantidad','Fecha del pedido','Plazo prometido (días)','Fecha de entrega prometida','Fecha de entrega real (006)','Desvío (días)','Estado de la entrega','Plazo real pedido→entrega (días)']].concat(e18));
  const cl={}; for(const x of c.cobros||[]){ if(!x.vto||!x.cobro) continue; (cl[x.cli]=cl[x.cli]||[]).push({d:dd(x.cobro,x.vto),imp:x.importe||0,rec:!!x.reclam}); }
  const pend={}; for(const x of c.cartera||[]) pend[x.cli]=(pend[x.cli]||0)+(x.pend||0);
  const cn={}; for(const x of c.clientes||[]) cn[x.cod]=x.nombre;
  const rowsC=Object.keys(cl).sort().map(k=>{ const a=cl[k]; return [k,cn[k]||k,a.length,a.reduce((s,x)=>s+x.imp,0),mean(a.map(x=>x.d)),a.filter(x=>x.d>S.grave).length/a.length,a.filter(x=>x.rec).length,pend[k]||0]; });
  const inc={}; for(const x of c.incidencias||[]) (inc[x.prov]=inc[x.prov]||[]).push(x);
  const rowsP=Object.keys(pv).sort().map(k=>{ const a=pv[k], I=inc[k]; let nInc,ab=0,resp='',resol='';
    if(I){ nInc=I.length; ab=I.filter(x=>!x.fecRes).length; const r1=I.filter(x=>x.fecResp&&x.fecCom).map(x=>dd(x.fecResp,x.fecCom)), r2=I.filter(x=>x.fecRes&&x.fecCom).map(x=>dd(x.fecRes,x.fecCom)); if(r1.length) resp=mean(r1); if(r2.length) resol=mean(r2); }
    else nInc=a.filter(x=>x.des>=S.desvio).length;
    return [k,pn[k]||k,a.length,mean(a.map(x=>x.des)),a.filter(x=>x.des<=0).length/a.length,mean(a.map(x=>x.real)),(cp[k]||{}).entrega,nInc,ab,resp,resol]; });
  add('020_Fiabilidad_Resumen',[['INDICADORES'],[],[],['CLIENTES'],['Cod.Cliente','Cliente','Cuotas cobradas','Importe cobrado (EUR)','Retraso medio (días)','% cuotas con retraso grave','Reclamaciones enviadas','Pendiente de cobro al corte (EUR)']].concat(rowsC,[[null,'Total'],[],['PROVEEDORES'],['Cod.Proveedor','Proveedor','Entregas','Desvío medio (días)','% entregas en plazo o antes','Plazo real pedido→entrega (días, media)','Plazo de 015 (días)','Incidencias','Incidencias abiertas','Días medios hasta la respuesta','Días medios hasta la resolución']],rowsP));
  const ro=c.rotura||{}; add('021_Parametros_Rotura',[[],[],[],['Sku','Descripción','Familia','Coste (EUR)','PVP (EUR)','Margen unitario (EUR)','Consumo / día','% de la venta que se pierde si no hay stock','','','','Recargo de un pedido urgente (% sobre el coste)']].concat(art.map(a=>{ const p=((c.precios||[]).find(x=>x.sku===a.sku))||{coste:0,pvp:0}, r=ro[a.sku]||{}; return [a.sku,a.desc,a.fam,p.coste,p.pvp,p.pvp-p.coste,cons[a.sku],r.pPerd!=null?r.pPerd:S.pPerd,'','','',r.recargo!=null?r.recargo:S.recargo]; })));
  for(const k of ['022_Estacionalidad','023_Caducidad','024_Alternativas_Proveedor']) if(c.extra&&c.extra[k]) add(k,c.extra[k]);
  return wb;
}
/* ---- libro plantilla -> canon (sirve para el round trip y para empezar el asistente desde el libro actual) ---- */
function fromTemplate(XLSX,wb){
  const R=n=>wb.Sheets[n]?XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,defval:null}):[];
  const body=(rows,from)=>{ const o=[]; for(let i=from;i<rows.length;i++){ if(rows[i][0]==null||rows[i][0]==='') break; o.push(rows[i]); } return o; };
  const hdr=(rows,f)=>rows.findIndex(f);
  const c={supuestos:{}}, P={}; for(const r of R('016_Parametros_App')){ const k=T(r[0]); if(!k) continue;
    if(k.startsWith('Saldo')) c.saldoMin=N(r[1]); else if(k.startsWith('Margen ámbar')) c.margenAmbar=N(r[1]); else if(k.startsWith('IVA')) c.iva=N(r[1]); else if(k.startsWith('Fecha de corte')) c.corte=iso(r[1]); else if(k.startsWith('Horizonte')) c.horizonte=iso(r[1]);
    else if(k.startsWith('Vigencia')) c.supuestos.vig=N(r[1]); else if(k.startsWith('Retraso')) c.supuestos.grave=N(r[1]); else if(k.startsWith('Desvío')) c.supuestos.desvio=N(r[1]); else if(k.startsWith('Empresa')) c.empresa=T(r[1]); }
  c.proveedores=R('001_Proveedores').slice(1).filter(r=>r[1]).map(r=>({nombre:T(r[0]),cod:T(r[1])}));
  c.clientes=R('002_Clientes').slice(1).filter(r=>r[0]).map(r=>({cod:T(r[0]),nombre:T(r[1])}));
  c.articulos=R('003_ Articulos').slice(1).filter(r=>r[0]).map(r=>({sku:T(r[0]),desc:T(r[1]),fam:T(r[3]),prov:T(r[4]),uni:T(r[6])}));
  c.inventario=R('004_Inventario').slice(1).filter(r=>r[0]).map(r=>({sku:T(r[0]),fisico:N(r[2]),min:N(r[3]),seg:N(r[4]),reservado:N(r[5])}));
  c.precios=R('008_Referencias_Precios').slice(1).filter(r=>r[0]).map(r=>({sku:T(r[0]),coste:N(r[2]),pvp:N(r[3])}));
  const r6=R('006_Compras_historicas').slice(1).filter(r=>r[3]); const r18=R('018_Entregas_Compras'), h18=hdr(r18,r=>T(r[0]).startsWith('Número de factura')), e18=h18<0?[]:body(r18,h18+1);
  c.compras=r6.map((r,k)=>({fac:T(r[0]),fecha:iso(r[1]),prov:T(r[2]),sku:T(r[3]),cant:N(r[4]),pu:N(r[6]),entrega:iso(r[12]),pedido:e18[k]?iso(e18[k][5]):null}));
  c.ventas=R('007_Ventas_historicas').slice(1).filter(r=>r[3]).map(r=>({fac:T(r[0]),fecha:iso(r[1]),cli:T(r[2]),sku:T(r[3]),cant:N(r[6]),pu:N(r[8]),imp:N(r[11])}));
  c.flujos=R('010_Flujos_Tesoreria').slice(1).filter(r=>r[0]).map(r=>({ref:T(r[0]),tipo:T(r[1]),importe:r[2],fecha:iso(r[4]),who:T(r[5])}));
  c.balance=R('011_Balance').slice(1).filter(r=>r[0]).map(r=>({cuenta:T(r[0]),desc:T(r[1]),saldo:N(r[2])}));
  c.cartera=R('012_Cartera_Cobros').slice(1).filter(r=>r[1]).map(r=>({ref:T(r[0]),cli:T(r[1]),fac:T(r[2]),cuota:T(r[3]),vto:iso(r[4]),importe:N(r[5]),forma:T(r[7]),pend:N(r[8])}));
  c.recepciones=R('014_Recepciones_Pendientes').slice(1).filter(r=>r[0]&&r[1]).map(r=>({ref:T(r[0]),sku:T(r[1]),cant:N(r[3]),fecha:iso(r[4])}));
  const r15=R('015_Condiciones_Proveedor'), h15=hdr(r15,r=>T(r[0])==='Cod.Proveedor');
  c.condProv=body(r15,h15+1).map(r=>({cod:T(r[0]),plazo:N(r[2]),forma:T(r[3]),cuotas:N(r[4]),dtoVol:N(r[5]),umbral:N(r[6]),portes:N(r[7]),franco:N(r[8]),minimo:N(r[9]),entrega:N(r[10])}));
  const r13=R('013_Planificacion_Stock'), h1=hdr(r13,r=>T(r[0])==='SKU'&&T(r[2])==='Desde'), h2=hdr(r13,r=>T(r[0])==='SKU'&&T(r[2]).startsWith('Consumo'));
  const b1=body(r13,h1+1); if(b1.length) c.periodo={desde:iso(b1[0][2]),hasta:iso(b1[0][3])};
  c.stock={}; for(const r of body(r13,h2+1)) c.stock[T(r[0])]={consumo:r[2],lead:r[5],obj:r[6],mult:r[7]};
  const r17=R('017_Cobros_Reales'); c.cobros=r17.slice(4).filter(r=>r[1]&&r[7]).map(r=>({cli:T(r[1]),vto:iso(r[4]),cobro:iso(r[7]),importe:N(r[5]),reclam:/^s/i.test(T(r[11]))}));
  const r19=R('019_Incidencias_Proveedor'); c.incidencias=r19.slice(4).filter(r=>r[1]&&/^\d/.test(T(r[1]))).map(r=>({prov:T(r[1]),fecCom:iso(r[7]),fecResp:iso(r[8]),fecRes:iso(r[9])}));
  const r21=R('021_Parametros_Rotura'), h21=hdr(r21,r=>T(r[0])==='Sku'); c.rotura={}; for(const r of body(r21,h21+1)) c.rotura[T(r[0])]={pPerd:N(r[7]),recargo:N(r[11])};
  c.extra={}; for(const k of ['022_Estacionalidad','023_Caducidad','024_Alternativas_Proveedor']) if(wb.Sheets[k]) c.extra[k]=XLSX.utils.sheet_to_json(wb.Sheets[k],{header:1,raw:true,defval:null});
  return c;
}
const api={toTemplate,fromTemplate,SUP}; if(typeof module!=='undefined') module.exports=api; else root.EIOSRaw=api;
})(typeof window!=='undefined'?window:this);
