/* Lector del libro de datos EIOS: convierte las hojas en el conjunto de datos que usa la app. */
(function(root){
const DAY=864e5;
/* origen de cada dato: hoja y celda del libro (traza del cálculo) */
const CL=i=>{ let t='',n=i+1; while(n>0){ const m=(n-1)%26; t=String.fromCharCode(65+m)+t; n=Math.floor((n-1)/26); } return t; };
const tagRows=(XLSX,ws,rows,name)=>{ let r0=0,c0=0; try{ const g=XLSX.utils.decode_range(ws['!ref']); r0=g.s.r; c0=g.s.c; }catch(e){} rows.forEach((r,i)=>{ if(r){ r.__s=name; r.__i=i+r0; r.__c=c0; } }); return rows; };
const cell=(r,c)=>(r&&r.__s)?r.__s+'!'+CL(c+(r.__c||0))+(r.__i+1):'';
const iso=v=>{ if(v==null||v==='') return null;
  if(typeof v==='number') return new Date(Math.round((v-25569)*DAY)).toISOString().slice(0,10);
  if(v instanceof Date) return new Date(Math.round(v.getTime()/DAY)*DAY).toISOString().slice(0,10);
  const s=String(v).trim(); let m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if(m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  m=s.match(/^(\d{4})-(\d{2})-(\d{2})/); return m?m[0]:null; };
const T=v=>v==null?'':String(v).trim();
const N=v=>{ if(v==null||v==='') return 0; const n=typeof v==='number'?v:parseFloat(String(v).replace(',','.')); return isNaN(n)?0:n; };
const r2=x=>Math.round((x+Number.EPSILON)*100)/100;
const rn=(x,k)=>{ const f=Math.pow(10,k); return Math.round(x*f)/f; };
function buildD(XLSX,wb,warn){
  warn=warn||(()=>{}); const miss=[], supu=[], mk=[], gl=[];
  const vac=v=>v==null||v===''||(typeof v==='string'&&isNaN(parseFloat(String(v).replace(',','.'))));
  const Nn=(v,ctx,key)=>{ if(vac(v)){ miss.push(ctx); mk.push({id:ctx.split(':')[0],key:key||'',t:ctx.split(': ').slice(1).join(': ')}); return 0; } return N(v); };
  let LS=''; const R=name=>{ LS=name; const ws=wb.Sheets[name]; if(!ws) throw new Error('Falta la hoja «'+name+'». Los nombres de las hojas deben ser los de la plantilla (mira «00_Leeme»).'); return tagRows(XLSX,ws,XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null}),name); };
  const have=name=>!!wb.Sheets[name];
  const body=(rows,from)=>{ const out=[]; for(let i=from;i<rows.length;i++){ if(rows[i][0]==null||rows[i][0]==='') break; out.push(rows[i]); } return out; };
  const hdr=(rows,pred)=>{ const i=rows.findIndex(pred); if(i<0) throw new Error('No encuentro la fila de encabezados en la hoja «'+LS+'». Comprueba que la primera columna se llama como en la plantilla y que no has movido ni borrado los encabezados.'); return i; };
  /* parámetros */
  const P={}; for(const r of R('016_Parametros_App')){ const k=T(r[0]); if(!k) continue;
    if(k.startsWith('Saldo mínimo')){ P.saldoMin=N(r[1]); P.oSaldoMin=cell(r,1); } else if(k.startsWith('Margen ámbar')) P.margenAmbar=N(r[1]); else if(k.startsWith('IVA')) P.iva=N(r[1]);
    else if(k.startsWith('Fecha de corte')) P.corte=iso(r[1]); else if(k.startsWith('Horizonte')) P.horizonte=iso(r[1]); else if(k.startsWith('Vigencia')) P.vigencia=N(r[1]); else if(k.startsWith('Empresa')) P.empresa=T(r[1]); }
  for(const k of ['saldoMin','margenAmbar','iva','corte','horizonte','vigencia']) if(P[k]==null||P[k]===''||P[k]===0&&k!=='margenAmbar') throw new Error('Falta el parámetro «'+k+'» en 016_Parametros_App.');
  /* balance */
  let banco=0, ac=0, pc=0, b430=0, oBanco=''; for(const r of R('011_Balance').slice(1)){ const c=T(r[0]), s=N(r[2]); if(!c) continue; if(c==='430') b430=s;
    if(c==='572'){ banco=s; oBanco=cell(r,2); } if(s>0&&/^[345]/.test(c)) ac+=s; if(s<0&&/^[45]/.test(c)) pc-=s; }
  ac=r2(ac); pc=r2(pc);
  /* proveedores */
  const nomP={}; for(const r of R('001_Proveedores').slice(1)) if(r[1]) nomP[T(r[1])]=T(r[0]);
  const r15=R('015_Condiciones_Proveedor'), h15=hdr(r15,r=>T(r[0])==='Cod.Proveedor'), suppliers={};
  for(const r of body(r15,h15+1)){ const c=T(r[0]); suppliers[c]={nombre:nomP[c]||T(r[1]),plazo:Nn(r[2],c+': plazo de pago (015)','plazoPago'),forma:T(r[3]),cuotas:N(r[4]),dtoVol:N(r[5]),umbral:N(r[6]),portes:N(r[7]),franco:N(r[8]),minimo:N(r[9]),entrega:N(r[10]),org:{plazo:cell(r,2),dtoVol:cell(r,5),umbral:cell(r,6),portes:cell(r,7),franco:cell(r,8),minimo:cell(r,9),entrega:cell(r,10)}}; }
  /* artículos */
  const inv={}; for(const r of R('004_Inventario').slice(1)) if(r[0]) { const k=T(r[0]); inv[k]={fisico:Nn(r[2],k+': stock físico (004)','fisico'),min:Nn(r[3],k+': stock mínimo (004)','min'),seg:Nn(r[4],k+': stock de seguridad (004)','seg'),reservado:Nn(r[5],k+': unidades reservadas (004)','reservado'),org:{fisico:cell(r,2),min:cell(r,3),seg:cell(r,4),reservado:cell(r,5)}}; }
  const pr={}; for(const r of R('008_Referencias_Precios').slice(1)) if(r[0]) { const k=T(r[0]); pr[k]={coste:Nn(r[2],k+': coste (008)','coste'),pvp:Nn(r[3],k+': PVP (008)','pvp'),org:{coste:cell(r,2),pvp:cell(r,3)}}; }
  const r13=R('013_Planificacion_Stock'), h1=hdr(r13,r=>T(r[0])==='SKU'&&T(r[2])==='Desde'), h2=hdr(r13,r=>T(r[0])==='SKU'&&T(r[2]).startsWith('Consumo'));
  const per={}; for(const r of body(r13,h1+1)) per[T(r[0])]={v92:N(r[5]),lv:N(r[6]),dd:N(r[4]),cons:N(r[7]),cob:N(r[11]),d:iso(r[2]),h:iso(r[3]),org:{v92:cell(r,5),dd:cell(r,4),lv:cell(r,6),consP:cell(r,7),cob:cell(r,11)}};
  const sup2={}; for(const r of body(r13,h2+1)) { const k=T(r[0]); sup2[k]={cons:Nn(r[2],k+': consumo diario (013)','cons'),lead:Nn(r[5],k+': plazo de entrega (013)','lead'),obj:Nn(r[6],k+': cobertura objetivo (013)','obj'),mult:N(r[7])||1,org:{cons:cell(r,2),lead:cell(r,5),obj:cell(r,6),mult:cell(r,7)}}; }
  const pend={}; if(have('014_Recepciones_Pendientes')) for(const r of body(R('014_Recepciones_Pendientes'),1)){ (pend[T(r[1])]=pend[T(r[1])]||[]).push({ref:T(r[0]),q:N(r[3]),d:iso(r[4])}); }
  const compras={}; for(const r of R('006_Compras_historicas').slice(1)){ const s=T(r[3]); if(!s) continue; (compras[s]=compras[s]||[]).push({fac:T(r[0]),fecha:iso(r[1]),pu:N(r[6]),cant:N(r[4]),org:cell(r,6)}); }
  const ventas={}; for(const r of R('007_Ventas_historicas').slice(1)){ const s=T(r[3]), f=iso(r[1]), p=per[s]; if(!s||!f||!p||f<p.d||f>p.h) continue; const v=ventas[s]=ventas[s]||{imp:0,q:0}; v.imp+=N(r[11]); v.q+=N(r[6]); }
  const items=[]; for(const r of R('003_ Articulos').slice(1)){ const sku=T(r[0]); if(!sku) continue;
    const iv=inv[sku], p=pr[sku], s2=sup2[sku], pe=per[sku]; if(!iv||!p||!s2||!pe) throw new Error('Al artículo '+sku+' le faltan datos en 004, 008 o 013.');
    const cs=(compras[sku]||[]).slice().sort((a,b)=>String(a.fecha).localeCompare(String(b.fecha))), last=cs[cs.length-1], v=ventas[sku];
    items.push({sku,desc:T(r[1]),fam:T(r[3]),uni:T(r[6]),prov:T(r[4]),fisico:iv.fisico,reservado:iv.reservado,disp:iv.fisico-iv.reservado,min:iv.min,seg:iv.seg,coste:p.coste,pvp:p.pvp,
      consumo:s2.cons||pe.cons,cob:pe.cob,v92:pe.v92,lv:pe.lv,dd:pe.dd,lead:s2.lead,obj:s2.obj,mult:s2.mult,pend:pend[sku]||[],
      ult:last?{fac:last.fac,fecha:last.fecha,pu:last.pu,cant:last.cant}:{fac:'',fecha:'',pu:0,cant:0},org:Object.assign({art:cell(r,0),ult:last?last.org:''},iv.org,p.org,{cons:(s2.cons?s2.org.cons:pe.org.consP),lead:s2.org.lead,obj:s2.org.obj,mult:s2.org.mult,v92:pe.org.v92,dd:pe.org.dd,cob:pe.org.cob}),ncomp:cs.length,falta:[],pmin:cs.length?Math.min(...cs.map(x=>x.pu)):0,pmax:cs.length?Math.max(...cs.map(x=>x.pu)):0,pvm:v&&v.q>0?v.imp/v.q:0}); }
  for(const it of items){ const f=[]; for(const m of mk){ if(m.id!==it.sku&&m.id!==it.prov) continue; if(m.key==='cons'&&it.consumo>0) continue; f.push({k:m.key,t:(m.id===it.prov?'proveedor '+m.id+': ':'')+m.t}); } it.falta=f; }
  for(const it of items) if(!suppliers[it.prov]) throw new Error('El proveedor '+it.prov+' del artículo '+it.sku+' no está en 015_Condiciones_Proveedor.');
  /* flujos de tesorería */
  const flows=[]; for(const r of R('010_Flujos_Tesoreria').slice(1)){ if(!r[0]) continue; const t=T(r[1]); if(t!=='COBRO'&&t!=='PAGO') continue; const a=Math.abs(N(r[2])); if(!iso(r[4])||r[2]==null||r[2]==='') { const m='Flujo de tesorería sin fecha o sin importe ('+T(r[5])+')'; miss.push(m); gl.push(m); } flows.push({d:iso(r[4]),amt:t==='PAGO'?-a:a,t,who:T(r[5])}); }
  /* clientes */
  const r20=R('020_Fiabilidad_Resumen'), hc=hdr(r20,r=>T(r[0])==='Cod.Cliente'), hp=hdr(r20,r=>T(r[0])==='Cod.Proveedor');
  const f20={}; for(const r of body(r20,hc+1)) f20[T(r[0])]={cobradas:N(r[2]),retraso:rn(N(r[4]),2),grave:rn(N(r[5]),4),reclam:N(r[6])};
  const nomC={}; for(const r of R('002_Clientes').slice(1)) if(r[0]) nomC[T(r[0])]=T(r[1]);
  const car={}; for(const r of R('012_Cartera_Cobros').slice(1)){ const c=T(r[1]); if(!c) continue; const o=car[c]=car[c]||{pend:0,cuotas:0,cuotasPend:0,fac:new Set(),rec:0}; o.cuotas++; const pe=N(r[8]); o.pend+=pe; if(pe>0) o.cuotasPend++; o.fac.add(T(r[2])); if(/recibo/i.test(T(r[7]))) o.rec++; }
  const vf={}; for(const r of R('007_Ventas_historicas').slice(1)){ const c=T(r[2]), f=iso(r[1]); if(c&&f&&(!vf[c]||f<vf[c])) vf[c]=f; }
  const cli=Object.keys(car).sort().map(c=>{ const o=car[c], f=f20[c]||{cobradas:0,retraso:0,grave:0,reclam:0}; return Object.assign({cod:c,nombre:nomC[c]||c,pend:r2(o.pend),cuotas:o.cuotas,cuotasPend:o.cuotasPend,desde:vf[c]||'',facturas:o.fac.size,recibo:rn(o.rec/o.cuotas,4)},f); });
  /* proveedores: historial */
  const r18=R('018_Entregas_Compras'), h18=hdr(r18,r=>T(r[0]).startsWith('Número de factura')), ent={}; for(const r of body(r18,h18+1)){ const c=T(r[2]); if(!c||!iso(r[1])||!iso(r[8])||typeof r[1]==='string'&&!/^\d/.test(r[1])) continue; const f=iso(r[1]), e=iso(r[8]); (ent[c]=ent[c]||[]).push({f,dd:Math.round((Date.parse(e)-Date.parse(f))/DAY)}); }
  const provH={}; for(const r of body(r20,hp+1)){ const c=T(r[0]), e=ent[c]||[], ds=e.map(x=>x.dd);
    provH[c]={n:N(r[2]),min:ds.length?Math.min(...ds):0,media:ds.length?rn(ds.reduce((a,b)=>a+b,0)/ds.length,1):0,max:ds.length?Math.max(...ds):0,desde:e.length?e.map(x=>x.f).sort()[0]:'',desvio:rn(N(r[3]),2),pctPlazo:rn(N(r[4]),4),plazoReal:rn(N(r[5]),2),nInc:N(r[7]),abiertas:N(r[8]),resp:typeof r[9]==='number'?rn(r[9],2):null,resol:typeof r[10]==='number'?rn(r[10],2):null}; }
  /* rotura */
  const r21=R('021_Parametros_Rotura'), h21=hdr(r21,r=>T(r[0])==='Sku'), rot={};
  for(const r of body(r21,h21+1)) rot[T(r[0])]={margen:rn(N(r[5]),2),pvp:N(r[4]),coste:N(r[3]),pPerd:N(r[7]),recargo:N(r[11])};
  for(const it of items) if(!rot[it.sku]) throw new Error('Falta el artículo '+it.sku+' en 021_Parametros_Rotura.');
  /* hojas opcionales: estacionalidad y caducidad (exploración) */
  let estac, estacMeta, cad; const skuSet=new Set(items.map(i=>i.sku));
  if(have('022_Estacionalidad')){ const r=R('022_Estacionalidad'), h=hdr(r,x=>T(x[0])==='Sku'); const e={}, em={}; const col=(hd,rx)=>hd.findIndex(c=>rx.test(T(c).toLowerCase())); const meta=(hd,x)=>{ const cf=col(hd,/fuente/), ca=col(hd,/a(ñ|n)os/), cb=col(hd,/autorizado/), cc=col(hd,/fecha de autoriz/); const fu=cf>=0?T(x[cf]).toLowerCase():''; return {fuente:/real/.test(fu)?'real':(/declar/.test(fu)?'declarado':'prueba'),anios:ca>=0?N(x[ca]):0,aut:cb>=0?T(x[cb]):'',fecha:cc>=0?iso(x[cc]):''}; }; for(const x of body(r,h+1)){ const k=T(x[0]); const f=[]; for(let m=2;m<14;m++) f.push(N(x[m])); if(!skuSet.has(k)){ supu.push('022: el artículo '+k+' no existe en 003'); continue; } if(f.some(v=>v<0)||!f.some(v=>v>0)){ supu.push('022: los factores de estacionalidad de '+k+' no son válidos y se ignoran'); continue; } e[k]=f; em[k]=meta(r[h],x); } if(Object.keys(e).length){ estac=e; estacMeta=em; } }
  if(have('023_Caducidad')){ const r=R('023_Caducidad'), h=hdr(r,x=>T(x[0])==='Sku'); const c={}; for(const x of body(r,h+1)){ const k=T(x[0]), v=N(x[3]); if(!skuSet.has(k)){ supu.push('023: el artículo '+k+' no existe en 003'); continue; } if(v>0){ const hd=r[h], cv=hd.findIndex(q=>/validad/.test(T(q).toLowerCase())), cf=hd.findIndex(q=>/fecha de validaci/.test(T(q).toLowerCase())); c[k]={vida:v,trat:T(x[4]),val:cv>=0?T(x[cv]):'',valF:cf>=0?iso(x[cf]):''}; } } if(Object.keys(c).length) cad=c; }
  /* hoja opcional: proveedores alternativos (024) */
  let alt;
  if(have('024_Alternativas_Proveedor')){ const r=R('024_Alternativas_Proveedor'), h=hdr(r,x=>T(x[0])==='Sku'); const al={};
    for(const x of body(r,h+1)){ const k=T(x[0]); if(!skuSet.has(k)){ supu.push('024: el artículo '+k+' no existe en 003'); continue; } al[k]=al[k]||[];
      const cod=T(x[2]); if(!cod) continue;
      const pr=N(x[4]); if(!(pr>0)){ supu.push('024: la alternativa '+cod+' de '+k+' no tiene precio válido y se ignora'); continue; }
      al[k].push({prov:cod,nombre:T(x[3])||cod,precio:pr,dto:N(x[5]),portes:N(x[6]),franco:N(x[7]),minimo:N(x[8]),plazo:N(x[9]),entrega:N(x[10]),equiv:/^s[ií]$/i.test(T(x[11])),hasta:iso(x[12])}); }
    if(Object.keys(al).length) alt=al; }
  const expl=(estac||cad)?'Se han leído '+[estac?'estacionalidad (022, '+Object.keys(estac).length+' artículos)':'',cad?'caducidad (023, '+Object.keys(cad).length+' artículos)':''].filter(Boolean).join(' y ')+'. Son supuestos y solo se usan en la tarjeta de Exploración «Temporada y caducidad», fuera del motor gobernado de EIOS.':'';
  const nAlt=alt?Object.values(alt).reduce((n,l)=>n+l.length,0):0;
  const explAlt=alt?'Se han leído '+nAlt+' ofertas de proveedores alternativos (024, '+Object.keys(alt).length+' artículos). Son supuestos de prueba y alimentan las reglas R-PROV-001 y R-PROV-002.':'';
  const f0=items.length?per[items[0].sku]:null, mes=f0&&f0.d&&f0.h?Math.max(1,Math.round(((Date.parse(f0.h)-Date.parse(f0.d))/DAY+1)/30.4)):0;
  return {info:{b430:r2(b430),miss,supu,expl:[expl,explAlt].filter(Boolean).join(' '),gl},D:{consumoBase:{fuente:'ventas',desde:f0?f0.d:'',hasta:f0?f0.h:'',meses:mes},empresa:P.empresa||'',org:{banco:oBanco,saldoMin:P.oSaldoMin||''},falta:gl.slice(),corte:P.corte,horizonte:P.horizonte,banco,saldoMin:P.saldoMin,margenAmbar:P.margenAmbar,iva:P.iva,vigencia:P.vigencia,suppliers,items,flows,ac,pc,...(estac?{estac}:{}),...(estacMeta?{estacMeta}:{}),...(cad?{cad}:{}),...(alt?{alt}:{})},cli,provH,rot};
}

/* Actualización diaria: aplica sobre una versión ya cargada solo las hojas que cambian cada día. */
const DAILY=['004_Inventario','014_Recepciones_Pendientes','011_Balance','010_Flujos_Tesoreria','012_Cartera_Cobros','016_Parametros_App'];
function applyDaily(XLSX,wb,base,opts){
  opts=opts||{}; const B=JSON.parse(JSON.stringify(base)), D=B.D, rep={hojas:[],avisos:[],cambios:[]}, DAYMS=864e5;
  const have=n=>!!wb.Sheets[n], R=n=>tagRows(XLSX,wb.Sheets[n],XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,defval:null}),n);
  const vac=v=>v==null||v===''||(typeof v==='string'&&isNaN(parseFloat(String(v).replace(',','.'))));
  const found=DAILY.filter(have); if(!found.length) throw new Error('El libro no trae ninguna de las hojas del día: '+DAILY.join(', ')+'. Los nombres deben ser los de la plantilla.');
  const oldCorte=D.corte; let corte=null, hor=null;
  if(have('016_Parametros_App')){ rep.hojas.push('016'); for(const r of R('016_Parametros_App')){ const k=T(r[0]); if(!k) continue;
    if(k.startsWith('Fecha de corte')&&iso(r[1])) corte=iso(r[1]); else if(k.startsWith('Horizonte')&&iso(r[1])) hor=iso(r[1]);
    else if(k.startsWith('Saldo mínimo')&&!vac(r[1])) D.saldoMin=N(r[1]); else if(k.startsWith('Margen ámbar')&&!vac(r[1])) D.margenAmbar=N(r[1]); else if(k.startsWith('IVA')&&!vac(r[1])) D.iva=N(r[1]); else if(k.startsWith('Vigencia')&&!vac(r[1])) D.vigencia=N(r[1]); } }
  if(opts.corte) corte=opts.corte;
  if(!corte) throw new Error('Indica la fecha de los datos (fecha de corte) en la página o en la hoja 016_Parametros_App.');
  if(corte<oldCorte) rep.avisos.push('La fecha de los datos ('+corte+') es anterior a la de la versión en uso ('+oldCorte+'). Revisa que sea el libro correcto.');
  if(!hor){ const off=Math.round((Date.parse(D.horizonte)-Date.parse(oldCorte))/DAYMS); hor=new Date(Date.parse(corte)+off*DAYMS).toISOString().slice(0,10); rep.avisos.push('El horizonte de tesorería no venía en el libro; se mantiene la misma amplitud que antes ('+off+' días).'); }
  D.corte=corte; D.horizonte=hor; rep.cambios.push('Fecha de los datos: '+corte);
  const bySku={}; for(const it of D.items) bySku[it.sku]=it;
  if(have('004_Inventario')){ rep.hojas.push('004'); const vistos=new Set(); let n=0;
    for(const r of R('004_Inventario').slice(1)){ const k=T(r[0]); if(!k) continue; const it=bySku[k]; if(!it){ rep.avisos.push('004: el artículo '+k+' no existe en la versión en uso y se ignora.'); continue; } vistos.add(k); n++;
      const set=(key,idx,txt)=>{ it.falta=(it.falta||[]).filter(f=>f.k!==key); if(vac(r[idx])){ it.falta.push({k:key,t:k+': '+txt+' (004)'}); return 0; } return N(r[idx]); };
      it.org=Object.assign({},it.org||{},{fisico:cell(r,2),min:cell(r,3),seg:cell(r,4),reservado:cell(r,5)}); it.fisico=set('fisico',2,'stock físico'); it.min=set('min',3,'stock mínimo'); it.seg=set('seg',4,'stock de seguridad'); it.reservado=set('reservado',5,'unidades reservadas'); it.disp=it.fisico-it.reservado; }
    for(const it of D.items) if(!vistos.has(it.sku)) rep.avisos.push('004: no hay existencias de '+it.sku+'; se mantiene el dato anterior.');
    rep.cambios.push('Existencias de '+n+' artículos'); }
  if(have('014_Recepciones_Pendientes')){ rep.hojas.push('014'); for(const it of D.items) it.pend=[]; let n=0;
    for(const r of R('014_Recepciones_Pendientes').slice(1)){ const k=T(r[1]); if(!k||!r[0]) continue; const it=bySku[k]; if(!it){ if(/^[A-Z]{2,4}-\d+/.test(k)) rep.avisos.push('014: el artículo '+k+' no existe en la versión en uso y se ignora.'); continue; } it.pend.push({ref:T(r[0]),q:N(r[3]),d:iso(r[4])}); n++; }
    rep.cambios.push('Recepciones pendientes: '+n); }
  if(have('011_Balance')){ rep.hojas.push('011'); let banco=0, ac=0, pc=0, hay=false, oB='';
    for(const r of R('011_Balance').slice(1)){ const cc=T(r[0]), s=N(r[2]); if(!cc) continue; hay=true; if(cc==='572'){ banco=s; oB=cell(r,2); } if(s>0&&/^[345]/.test(cc)) ac+=s; if(s<0&&/^[45]/.test(cc)) pc-=s; }
    if(!hay) rep.avisos.push('011: la hoja está vacía; se mantiene el saldo anterior.'); else { D.banco=banco; D.org=Object.assign({},D.org||{},{banco:oB}); D.ac=r2(ac); D.pc=r2(pc); rep.cambios.push('Banco al corte: '+r2(banco)); } }
  if(have('010_Flujos_Tesoreria')){ rep.hojas.push('010'); const flows=[]; D.falta=(D.falta||[]).filter(x=>!/^Flujo de tesorería/.test(x));
    for(const r of R('010_Flujos_Tesoreria').slice(1)){ if(!r[0]) continue; const t=T(r[1]); if(t!=='COBRO'&&t!=='PAGO') continue; const a=Math.abs(N(r[2])); if(!iso(r[4])||vac(r[2])) D.falta.push('Flujo de tesorería sin fecha o sin importe ('+T(r[5])+')'); flows.push({d:iso(r[4]),amt:t==='PAGO'?-a:a,t,who:T(r[5])}); }
    D.flows=flows; rep.cambios.push('Cobros y pagos previstos: '+flows.length); }
  if(have('012_Cartera_Cobros')){ rep.hojas.push('012'); const car={};
    for(const r of R('012_Cartera_Cobros').slice(1)){ const cc=T(r[1]); if(!cc) continue; const o=car[cc]=car[cc]||{pend:0,cuotas:0,cuotasPend:0,fac:new Set(),rec:0}; o.cuotas++; const pe=N(r[8]); o.pend+=pe; if(pe>0) o.cuotasPend++; o.fac.add(T(r[2])); if(/recibo/i.test(T(r[7]))) o.rec++; }
    for(const c of B.cli){ const o=car[c.cod]; if(!o){ c.pend=0; c.cuotasPend=0; continue; } c.pend=r2(o.pend); c.cuotas=o.cuotas; c.cuotasPend=o.cuotasPend; c.facturas=o.fac.size; c.recibo=rn(o.rec/o.cuotas,4); }
    for(const k of Object.keys(car)) if(!B.cli.some(c=>c.cod===k)) rep.avisos.push('012: el cliente '+k+' no existe en la versión en uso y se ignora (para añadirlo, importa el libro completo).');
    rep.cambios.push('Cartera de cobros de '+Object.keys(car).length+' clientes'); }
  D.actualizado={fecha:corte,hojas:rep.hojas.slice()};
  return {B,rep};
}
const api={buildD,applyDaily,DAILY}; if(typeof module!=='undefined') module.exports=api; else root.EIOSImport=api;
})(typeof window!=='undefined'?window:this);
