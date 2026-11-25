/* Asistente de columnas: sugiere qué columna del Excel del ERP es cada dato y construye el conjunto "canon" */
(function(root){
const norm=s=>String(s==null?'':s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/[^a-z0-9]+/g,' ').trim();
const PROV=[/(cod|codigo|id|ref|num|n).*proveedor|proveedor.*(cod|codigo|id|ref|num)/,/^proveedor( habitual| principal)?$/,/^prov$/,/registrar factura de proveedor|^factura de proveedor$/,/^socio$/];
const CLI=[/(cod|codigo|id|ref|num|n).*cliente|cliente.*(cod|codigo|id|ref|num)/,/^cliente$/,/^(socio|contacto|empresa)$/,/registrar factura de cliente|^factura de cliente$/];
const SKU=[/^sku$/,/referencia interna/,/(codigo|cod|ref|referencia) (del )?(articulo|producto)/,/(articulo|producto|item) (codigo|cod|ref)/,/^(codigo|cod|referencia|ref|articulo|producto|item)$/,/(^| )(producto|articulo)$/];
const FAC=[/(numero|num|n|no) (de )?factura/,/^factura/,/albaran|documento/,/^numero$/];
const TABLES=[
 {id:'articulos',l:'Artículos',req:1,sheet:/articul|producto|item|catalogo|maestro/,f:[
  {k:'sku',l:'Código del artículo (SKU)',req:1,syn:SKU,t:'t'},{k:'desc',l:'Descripción',syn:[/^(?!.*moneda).*(descripcion|denominacion|nombre)/],t:'t'},{k:'fam',l:'Familia',syn:[/familia|categoria|grupo|seccion/],t:'t'},
  {k:'prov',l:'Código del proveedor habitual (si no lo tienes aquí, usa la tabla «Proveedor de cada artículo»)',syn:PROV,t:'t'},{k:'uni',l:'Unidad',syn:[/unidad|medida|^ud$/],t:'t'}]},
 {id:'inventario',l:'Existencias',req:1,sheet:/inventario|stock|existencias|almacen/,f:[
  {k:'sku',l:'Código del artículo (SKU)',req:1,syn:SKU,t:'t'},{k:'fisico',l:'Existencias físicas',req:1,syn:[/existencias|stock (fisico|actual|real)|cantidad (real|a mano|en (stock|almacen))|^stock$|unidades en|fisico|disponible/],t:'n'},
  {k:'min',l:'Stock mínimo',syn:[/minimo/],t:'n'},{k:'seg',l:'Stock de seguridad',syn:[/seguridad/],t:'n'},{k:'reservado',l:'Unidades reservadas',syn:[/reserva|comprometid/],t:'n'}]},
 {id:'precios',l:'Costes y precios de venta',req:1,sheet:/precio|tarifa|coste|referencia/,f:[
  {k:'sku',l:'Código del artículo (SKU)',req:1,syn:SKU,t:'t'},{k:'coste',l:'Coste',req:1,syn:[/^(coste|costo)( unitario| medio)?$/,/^(?!.*moneda).*(coste|costo|precio (de )?compra)/],t:'n'},{k:'pvp',l:'Precio de venta (PVP)',req:1,syn:[/pvp|precio (de )?venta|tarifa/],t:'n'}]},
 {id:'compras',l:'Facturas de compra',req:1,fill:['fac','fecha','prov','entrega','pedido'],sheet:/compra|proveedor.*factura|factura.*proveedor/,f:[
  {k:'fac',l:'Número de factura',req:1,syn:FAC,t:'t'},{k:'fecha',l:'Fecha de la factura',req:1,syn:[/fecha (de )?(la )?factura/,/^fecha$/],t:'d'},{k:'prov',l:'Código del proveedor',req:1,syn:PROV,t:'t'},{k:'sku',l:'Código del artículo (SKU)',req:1,syn:SKU,t:'t'},
  {k:'cant',l:'Cantidad',req:1,syn:[/cantidad|unidades/],t:'n'},{k:'pu',l:'Precio unitario neto',req:1,syn:[/precio (unitario|neto|unit)/,/^precio/,/coste unitario/],t:'n'},
  {k:'entrega',l:'Fecha de entrega real',syn:[/fecha (de )?(entrega|recepcion)/,/entrega real|recibido/],t:'d'},{k:'pedido',l:'Fecha del pedido',syn:[/fecha (de )?pedido/],t:'d'}]},
 {id:'ventas',l:'Facturas de venta',req:1,fill:['fac','fecha','cli'],sheet:/venta|factura.*cliente|cliente.*factura/,f:[
  {k:'fac',l:'Número de factura',req:1,syn:FAC,t:'t'},{k:'fecha',l:'Fecha de la factura',req:1,syn:[/fecha (de )?(la )?factura/,/^fecha$/],t:'d'},{k:'cli',l:'Código del cliente',req:1,syn:CLI,t:'t'},{k:'sku',l:'Código del artículo (SKU)',req:1,syn:SKU,t:'t'},
  {k:'cant',l:'Cantidad',req:1,syn:[/cantidad|unidades/],t:'n'},{k:'pu',l:'Precio unitario neto',syn:[/precio (unitario|neto|unit)/,/^precio/],t:'n'},{k:'imp',l:'Importe neto de la línea',req:1,syn:[/importe (neto|linea)|total (linea|neto)?|^importe$/],t:'n'}]},
 {id:'flujos',l:'Cobros y pagos previstos (tesorería)',req:1,nn:1,sheet:/tesoreria|flujo|prevision|vencimiento/,f:[
  {k:'ref',l:'Referencia',syn:[/referencia|^id$|numero|concepto/],t:'t'},{k:'tipo',l:'Tipo (cobro o pago)',req:1,syn:[/tipo|signo|clase|naturaleza/],t:'t'},{k:'importe',l:'Importe',req:1,syn:[/importe|cantidad|valor|total/],t:'n'},
  {k:'fecha',l:'Fecha de vencimiento',req:1,syn:[/vencimiento/,/fecha/],t:'d'},{k:'who',l:'Cliente o proveedor',syn:[/contraparte|cliente|proveedor|nombre|tercero|socio/],t:'t'}]},
 {id:'balance',l:'Balance (saldos al corte)',req:1,nn:1,sheet:/balance|saldo|cuentas|contab/,f:[
  {k:'cuenta',l:'Código de cuenta',req:1,syn:[/cuenta|codigo/],t:'t'},{k:'desc',l:'Descripción',syn:[/descripcion|nombre|concepto/],t:'t'},{k:'saldo',l:'Saldo',req:1,syn:[/saldo|importe/],t:'n'}]},
 {id:'cartera',l:'Cartera de cobros (cuotas de clientes)',req:1,sheet:/cartera|cobro|recibo|vencimiento/,f:[
  {k:'ref',l:'Referencia de la cuota',syn:[/referencia|recibo|cuota/],t:'t'},{k:'cli',l:'Código del cliente',req:1,syn:CLI,t:'t'},{k:'fac',l:'Número de factura',req:1,syn:FAC,t:'t'},{k:'cuota',l:'Número de cuota',syn:[/(numero|num|n) (de )?cuota|^cuota$/],t:'t'},
  {k:'vto',l:'Fecha de vencimiento',req:1,syn:[/vencimiento/,/fecha/],t:'d'},{k:'importe',l:'Importe de la cuota',req:1,syn:[/^importe$/,/importe (total|cuota)/,/^total/],t:'n'},{k:'forma',l:'Forma de cobro',syn:[/forma|medio|metodo/],t:'t'},{k:'pend',l:'Importe pendiente al corte',req:1,syn:[/pendiente|adeudado/,/saldo/],t:'n'}]},
 {id:'condProv',l:'Condiciones de los proveedores',req:0,sheet:/condicion|proveedor/,f:[
  {k:'cod',l:'Código del proveedor',req:1,syn:PROV.concat([/^codigo$/,/^cod$/]),t:'t'},{k:'plazo',l:'Plazo de pago (días)',req:1,syn:[/plazo (de )?pago|dias (de )?pago/],t:'n'},{k:'forma',l:'Forma de pago',syn:[/forma/],t:'t'},{k:'cuotas',l:'Cuotas máximas',syn:[/cuotas/],t:'n'},
  {k:'dtoVol',l:'Descuento por volumen (%)',syn:[/descuento|dto/],t:'n'},{k:'umbral',l:'Umbral de volumen (EUR)',syn:[/umbral|volumen/],t:'n'},{k:'portes',l:'Portes por pedido (EUR)',syn:[/^(?!.*franco).*(portes|transporte)/],t:'n'},{k:'franco',l:'Franco de portes (EUR)',syn:[/franco/],t:'n'},
  {k:'minimo',l:'Pedido mínimo (EUR)',syn:[/minimo/],t:'n'},{k:'entrega',l:'Plazo de entrega (días)',req:1,syn:[/plazo (de )?entrega|dias (de )?entrega|lead ?time/],t:'n'}]},
 {id:'recepciones',l:'Recepciones pendientes',req:0,sheet:/recepcion|pendiente|pedido/,f:[
  {k:'ref',l:'Referencia',req:1,syn:[/referencia|pedido|^id$/],t:'t'},{k:'sku',l:'Código del artículo (SKU)',req:1,syn:SKU,t:'t'},{k:'cant',l:'Cantidad pendiente',req:1,syn:[/cantidad|pendiente/],t:'n'},{k:'fecha',l:'Fecha prevista',req:1,syn:[/prevista/,/fecha/],t:'d'}]},
 {id:'cobros',l:'Historial de cobros reales (para medir la fiabilidad de clientes)',req:0,sheet:/cobros? real|historico|cobrad/,f:[
  {k:'cli',l:'Código del cliente',req:1,syn:CLI,t:'t'},{k:'vto',l:'Fecha de vencimiento',req:1,syn:[/vencimiento/],t:'d'},{k:'cobro',l:'Fecha real de cobro',req:1,syn:[/fecha (real )?(de )?cobro/,/fecha (real|de pago)/],t:'d'},{k:'importe',l:'Importe cobrado',syn:[/importe|cobrado/],t:'n'},{k:'reclam',l:'Reclamación enviada (Sí/No)',syn:[/reclam/],t:'t'}]},
 {id:'incidencias',l:'Incidencias de proveedores',req:0,sheet:/incidencia|reclamaci/,f:[
  {k:'prov',l:'Código del proveedor',req:1,syn:PROV,t:'t'},{k:'fecCom',l:'Fecha de comunicación al proveedor',req:1,syn:[/comunica/,/fecha (de la )?incidencia/,/^fecha$/],t:'d'},{k:'fecResp',l:'Fecha de primera respuesta',syn:[/respuesta/],t:'d'},{k:'fecRes',l:'Fecha de resolución',syn:[/resolucion|cierre/],t:'d'}]},
{id:'provArt',l:'Proveedor de cada artículo (precio y plazo de entrega)',req:0,sheet:/proveedor|producto/,f:[
  {k:'sku',l:'Código del artículo (SKU)',req:1,syn:SKU.concat([/plantilla de producto|^producto$/]),t:'t'},{k:'prov',l:'Proveedor',req:1,syn:PROV,t:'t'},{k:'precio',l:'Precio unitario de compra',syn:[/precio/],t:'n'},{k:'entrega',l:'Plazo de entrega (días)',syn:[/plazo (de )?entrega|dias (de )?entrega|lead ?time/],t:'n'}]},
 {id:'proveedores',l:'Proveedores (nombres)',req:0,sheet:/proveedor/,f:[{k:'cod',l:'Código del proveedor',req:1,syn:PROV.concat([/^codigo$/,/^cod$/]),t:'t'},{k:'nombre',l:'Nombre',req:1,syn:[/nombre|razon|proveedor/],t:'t'}]},
 {id:'clientes',l:'Clientes (nombres)',req:0,sheet:/cliente/,f:[{k:'cod',l:'Código del cliente',req:1,syn:CLI.concat([/^codigo$/,/^cod$/]),t:'t'},{k:'nombre',l:'Nombre',req:1,syn:[/nombre|razon|cliente/],t:'t'}]}
];
/* fila de encabezados: la primera de las 15 primeras con >=2 celdas de texto y mayoría de texto */
function headerRow(rows){ for(let i=0;i<Math.min(rows.length,15);i++){ const r=rows[i]||[], ne=r.filter(v=>v!=null&&v!==''); if(ne.length>=2&&ne.filter(v=>typeof v==='string').length/ne.length>=0.7&&r.filter(v=>v!=null&&v!=='').length>=Math.min(3,ne.length)) { const nx=(rows[i+1]||[]).filter(v=>v!=null&&v!==''); if(nx.length) return i; } } return 0; }
function suggestCols(headers,table){ const used=new Set(), out={}; const hs=headers.map(norm);
  for(const f of table.f){ out[f.k]=-1; for(const rx of f.syn){ const j=hs.findIndex((h,ix)=>!used.has(ix)&&h&&rx.test(h)); if(j>=0){ out[f.k]=j; used.add(j); break; } } }
  return out; }
/* sheets = {nombre: filas[][]}; devuelve la propuesta de mapeo por tabla */
function suggestAll(sheets){ const res={}, names=Object.keys(sheets), hd={}; for(const n of names) hd[n]=headerRow(sheets[n]);
  for(const t of TABLES){ let best=null;
    for(const n of names){ const h=(sheets[n][hd[n]]||[]), c=suggestCols(h,t); const reqN=t.f.filter(f=>f.req).length, hitReq=t.f.filter(f=>f.req&&c[f.k]>=0).length, hit=t.f.filter(f=>c[f.k]>=0).length, nameB=t.sheet.test(norm(n))?1.5:0;
      const sc=hitReq*2+(hit-hitReq)+nameB*2; if((hitReq>=reqN||(t.req&&nameB>0&&hitReq>=Math.max(1,reqN-1)))&&((t.req&&!t.nn)||nameB>0)&&(!best||sc>best.sc)) best={sheet:n,hdr:hd[n],c,sc,hitReq,reqN}; }
    res[t.id]=best?{sheet:best.sheet,hdr:best.hdr,f:Object.fromEntries(Object.entries(best.c).map(([k,v])=>[k,{c:v,v:''}])),completo:best.hitReq===best.reqN}:null; }
  return res; }
const N=v=>{ if(v==null||v==='') return null; if(typeof v==='number') return v; let t=String(v).trim().replace(/\s|€/g,''); if(/,/.test(t)&&/\./.test(t)) t=t.replace(/\./g,'').replace(',','.'); else t=t.replace(',','.'); const n=parseFloat(t); return isNaN(n)?null:n; };
const DAY=864e5;
const D=v=>{ if(v==null||v==='') return null; if(typeof v==='number') return new Date(Math.round((v-25569)*DAY)).toISOString().slice(0,10); if(v instanceof Date) return v.toISOString().slice(0,10);
  const s=String(v).trim(); let m=s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/); if(m) return `${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`; m=s.match(/^(\d{4})-(\d{2})-(\d{2})/); return m?m[0]:null; };
const Tt=v=>v==null?'':(typeof v==='number'?String(v):String(v).trim());
/* lee una tabla con su mapeo */
function readTable(sheets,tb,m,avisos){ const t=TABLES.find(x=>x.id===tb); if(!m||!sheets[m.sheet]) return null; const rows=sheets[m.sheet].slice((m.hdr||0)+1), out=[]; let malos=0;
  let prev={}; for(const r of rows){ if(!r||!r.some(v=>v!=null&&v!=='')) continue; const o={}; let ok=true;
    for(const f of t.f){ const mm=(m.f||{})[f.k]||{c:-1,v:''}; let raw=mm.c>=0?r[mm.c]:(mm.v!==''&&mm.v!=null?mm.v:null);
      if(t.fill&&t.fill.includes(f.k)){ if(mm.c>=0&&(raw==null||raw==='')&&prev[f.k]!=null) raw=prev[f.k]; else if(mm.c>=0) prev[f.k]=raw; }
      let val=f.t==='n'?N(raw):f.t==='d'?D(raw):Tt(raw); if(f.k==='sku'&&typeof val==='string'){ const mm=val.match(/^\[([^\]]+)\]/); if(mm) val=mm[1].trim(); } if(f.t==='t'&&val==='') val=''; if(f.req&&(val==null||val==='')){ ok=false; } o[f.k]=val; }
    if(ok) out.push(o); else malos++; }
  if(malos&&avisos) avisos.push(t.l+': '+malos+' filas se han ignorado por tener vacío algún dato obligatorio.');
  return out; }
function tipoFlujo(v,imp){ const s=norm(v); if(/cobr|ingres|entrada|cliente|^c$|^i$|haber/.test(s)) return 'COBRO'; if(/pag|gast|salida|proveedor|^p$|debe/.test(s)) return 'PAGO'; return imp<0?'PAGO':'COBRO'; }
/* canon a partir de las hojas, el mapeo y el formulario de parámetros */
function buildCanon(sheets,map,P){ const av=[], T={}; for(const t of TABLES) T[t.id]=readTable(sheets,t.id,map[t.id],av);
  const c={empresa:P.empresa||'',corte:P.corte,horizonte:P.horizonte||null,saldoMin:P.saldoMin,margenAmbar:P.margenAmbar,iva:P.iva>1?P.iva/100:P.iva,supuestos:Object.assign({},P.supuestos||{}),stock:P.stock||{},rotura:P.rotura||{}};
  for(const k of Object.keys(c.supuestos)) if(c.supuestos[k]===''||c.supuestos[k]==null||isNaN(c.supuestos[k])) delete c.supuestos[k];
  const pa={}, le={}; for(const x of T.provArt||[]){ if(!pa[x.sku]) pa[x.sku]=x.prov; if(x.entrega!=null) (le[x.prov]=le[x.prov]||[]).push(x.entrega); }
  c.articulos=(T.articulos||[]).map(x=>({sku:x.sku,desc:x.desc,fam:x.fam,prov:x.prov||pa[x.sku]||'',uni:x.uni}));
  c.inventario=(T.inventario||[]).map(x=>({sku:x.sku,fisico:x.fisico,min:x.min,seg:x.seg,reservado:x.reservado}));
  c.precios=(T.precios||[]).map(x=>({sku:x.sku,coste:x.coste,pvp:x.pvp}));
  c.compras=(T.compras||[]).map(x=>({fac:x.fac,fecha:x.fecha,prov:x.prov,sku:x.sku,cant:x.cant,pu:x.pu,entrega:x.entrega||null,pedido:x.pedido||null}));
  c.ventas=(T.ventas||[]).map(x=>({fac:x.fac,fecha:x.fecha,cli:x.cli,sku:x.sku,cant:x.cant,pu:x.pu||0,imp:x.imp}));
  c.flujos=(T.flujos||[]).map((x,i)=>({ref:x.ref||('F'+(i+1)),tipo:tipoFlujo(x.tipo,x.importe),importe:x.importe,fecha:x.fecha,who:x.who||''}));
  c.balance=(T.balance||[]).map(x=>({cuenta:x.cuenta,desc:x.desc,saldo:x.saldo}));
  c.cartera=(T.cartera||[]).map((x,i)=>({ref:x.ref||(x.fac+'#'+(x.cuota||i+1)),cli:x.cli,fac:x.fac,cuota:x.cuota||'1',vto:x.vto,importe:x.importe,forma:x.forma||'',pend:x.pend}));
  c.recepciones=T.recepciones||[]; c.cobros=(T.cobros||[]).map(x=>({cli:x.cli,vto:x.vto,cobro:x.cobro,importe:x.importe||0,reclam:/^(s|si|1|true|x)$/i.test(norm(x.reclam))})); c.incidencias=T.incidencias||[];
  const cod=(a,k)=>[...new Set(a.map(x=>x[k]).filter(Boolean))].sort();
  const pn={}; for(const x of T.proveedores||[]) pn[x.cod]=x.nombre; const cn={}; for(const x of T.clientes||[]) cn[x.cod]=x.nombre;
  c.proveedores=cod(c.articulos.concat(c.compras),'prov').map(k=>({cod:k,nombre:pn[k]||k})); c.clientes=cod(c.ventas.concat(c.cartera),'cli').map(k=>({cod:k,nombre:cn[k]||k}));
  if(T.condProv) c.condProv=T.condProv.map(x=>({cod:x.cod,plazo:x.plazo,forma:x.forma||'',cuotas:x.cuotas||0,dtoVol:x.dtoVol||0,umbral:x.umbral||0,portes:x.portes||0,franco:x.franco||0,minimo:x.minimo||0,entrega:x.entrega}));
  else c.condProv=c.proveedores.map(p=>{ const q=(P.condProv||{})[p.cod]||{}; return {cod:p.cod,plazo:q.plazo,forma:q.forma||'',cuotas:q.cuotas||0,dtoVol:q.dtoVol||0,umbral:q.umbral||0,portes:q.portes||0,franco:q.franco||0,minimo:q.minimo||0,entrega:q.entrega!=null?q.entrega:(le[p.cod]?Math.round(le[p.cod].reduce((a,b)=>a+b,0)/le[p.cod].length):undefined)}; });
  if(!c.horizonte&&c.corte) c.horizonte=null;
  return {canon:c,avisos:av}; }
const api={TABLES,norm,headerRow,suggestCols,suggestAll,buildCanon,readTable}; if(typeof module!=='undefined') module.exports=api; else root.EIOSMap=api;
})(typeof window!=='undefined'?window:this);
