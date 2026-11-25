/* Almacenamiento local para demo estática, sin servidor. */
(()=>{
const prefix='andres-ragel/eios-demo/v1/', listeners=new Map();
const clone=v=>JSON.parse(JSON.stringify(v)), read=n=>JSON.parse(localStorage.getItem(prefix+n)||'{}');
const snap=(id,v)=>({id,exists:v!==undefined,data:()=>v===undefined?undefined:clone(v)});
const query=n=>({docs:Object.entries(read(n)).map(([id,v])=>snap(id,v))});
const notify=n=>{for(const x of listeners.get(n)||[]){try{x.cb(query(n));}catch(e){x.err?.(e);}}};
const mutate=(n,id,fn)=>{const d=read(n);fn(d,id);localStorage.setItem(prefix+n,JSON.stringify(d));notify(n);};
const db={collection(n){return {get:async()=>query(n),onSnapshot(cb,err){const s=listeners.get(n)||new Set(),x={cb,err};s.add(x);listeners.set(n,s);try{cb(query(n));}catch(e){err?.(e);}return()=>s.delete(x);},doc(id){return {get:async()=>snap(id,read(n)[id]),set:async v=>mutate(n,id,(d,k)=>{d[k]=clone(v);}),update:async v=>mutate(n,id,(d,k)=>{if(!Object.hasOwn(d,k))throw Error('Documento inexistente');d[k]={...d[k],...clone(v)};}),delete:async()=>mutate(n,id,(d,k)=>{delete d[k];})};}};}};
try{const k=prefix+'test';localStorage.setItem(k,'1');localStorage.removeItem(k);window.claude={use:async kind=>kind==='db'?db:null};}catch(e){console.warn('Almacenamiento local no disponible',e);}
window.addEventListener('storage',e=>{if(e.key?.startsWith(prefix))notify(e.key.slice(prefix.length));});
})();
