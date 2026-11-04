'use strict';
/* ---------- helpers ---------- */
const ICONS={plus:'M12 5v14M5 12h14',back:'M19 12H5M12 19l-7-7 7-7',copy:'M9 9h11v11H9zM5 15H4V4h11v1',check:'M20 6L9 17l-5-5',trash:'M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6',edit:'M12 20h9M16.5 3.5l4 4L8 20H4v-4z',lib:'M4 4v16M9 4v16M14 5l5 15',search:'M11 19a8 8 0 100-16 8 8 0 000 16zM21 21l-4.3-4.3',file:'M14 2H6v20h12V8zM14 2v6h6M8 13h8M8 17h8',menu:'M4 6h16M4 12h16M4 18h16',x:'M18 6L6 18M6 6l12 12',gear:'M12 15a3 3 0 100-6 3 3 0 000 6zM19.4 15a1.7 1.7 0 00.3 1.8l.1.1a2 2 0 11-2.8 2.8l-.1-.1a1.7 1.7 0 00-2.9 1.2V21a2 2 0 11-4 0v-.1a1.7 1.7 0 00-2.9-1.2l-.1.1a2 2 0 11-2.8-2.8l.1-.1a1.7 1.7 0 00-1.2-2.9H3a2 2 0 110-4h.1a1.7 1.7 0 001.2-2.9l-.1-.1a2 2 0 112.8-2.8l.1.1a1.7 1.7 0 002.9-1.2V3a2 2 0 114 0v.1a1.7 1.7 0 002.9 1.2l.1-.1a2 2 0 112.8 2.8l-.1.1a1.7 1.7 0 001.2 2.9h.1a2 2 0 110 4h-.1a1.7 1.7 0 00-1.5 1z',grip:'M9 5h.01M9 12h.01M9 19h.01M15 5h.01M15 12h.01M15 19h.01',split:'M6 3v18M18 3v18M6 12h12',import:'M12 3v12M7 10l5 5 5-5M4 21h16',export:'M12 15V3M7 8l5-5 5 5M4 21h16',down:'M6 9l6 6 6-6',up:'M18 15l-6-6-6 6',eye:'M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8zM12 15a3 3 0 100-6 3 3 0 000 6z',save:'M5 3h11l5 5v13H3V3zM7 3v6h8M7 21v-8h10v8',folder:'M3 6h6l2 3h10v11H3z',move:'M5 9l-3 3 3 3M9 5l3-3 3 3M15 19l-3 3-3-3M19 9l3 3-3 3M2 12h20M12 2v20',book:'M4 19V5a2 2 0 012-2h14v16H6a2 2 0 00-2 2 2 2 0 002 2h14v-4',layers:'M12 2L2 7l10 5 10-5zM2 17l10 5 10-5M2 12l10 5 10-5'};
const icon=n=>{const s=document.createElementNS('http://www.w3.org/2000/svg','svg');s.setAttribute('viewBox','0 0 24 24');s.setAttribute('class','i');s.setAttribute('aria-hidden','true');const p=document.createElementNS('http://www.w3.org/2000/svg','path');p.setAttribute('d',ICONS[n]);s.append(p);return s};
function h(t,p,...c){const e=document.createElement(t);for(const k in p||{}){const v=p[k];if(v==null||v===false)continue;if(k.startsWith('on'))e.addEventListener(k.slice(2),v);else if(k==='class')e.className=v;else if(k==='html')e.innerHTML=v;else if(k in e&&k!=='list'&&k!=='style')e[k]=v;else e.setAttribute(k,v===true?'':v)}
 for(const x of c.flat(9)){if(x==null||x===false)continue;e.append(x)}return e}
const btn=(label,on,cls='',ic,extra={})=>h('button',{type:'button',class:'btn '+cls,onclick:on,...extra},ic&&icon(ic),label);
const ibtn=(ic,on,title,cls='')=>h('button',{type:'button',class:'btn ghost icon '+cls,onclick:on,title,'aria-label':title},icon(ic));
const esc=s=>String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const unhtml=s=>{const d=document.createElement('div');d.innerHTML=s;return d.textContent};
const uid=()=>Date.now().toString(36)+Math.random().toString(36).slice(2,8);
const toast=(m,e)=>{const t=h('div',{class:'toast'+(e?' e':'')},m);document.getElementById('toasts').append(t);setTimeout(()=>t.remove(),e?5000:2600)};
const dateStr=d=>d?new Date(d).toLocaleDateString():'';
const FOOT="font-family: 'Times New Roman', Times, serif; font-size: 10pt; margin: 0; line-height: 1.5; text-align: justify;";
async function copyRich(html,text){
 try{await navigator.clipboard.write([new ClipboardItem({'text/html':new Blob([html],{type:'text/html'}),'text/plain':new Blob([text],{type:'text/plain'})})]);return true}catch(e){}
 try{const d=h('div',{html,contentEditable:'true',style:'position:fixed;left:-9999px;top:0'});document.body.append(d);const r=document.createRange();r.selectNodeContents(d);const s=getSelection();s.removeAllRanges();s.addRange(r);const ok=document.execCommand('copy');s.removeAllRanges();d.remove();if(ok)return true}catch(e){}
 try{await navigator.clipboard.writeText(text);return true}catch(e){toast('Copying was refused here. Select the text and copy it by hand.',1);return false}
}
/* ---------- data store: one row per record in Supabase, private to each account ---------- */
const sb=supabase.createClient(LEXCITR_CONFIG.SUPABASE_URL,LEXCITR_CONFIG.SUPABASE_KEY);
const KINDS=['style','type','citation','library','pinpoint'];
const S={style:[],type:[],citation:[],library:[],pinpoint:[]};
let storeMode='loading',storeNote='',currentUser=null;
const sortAll=()=>{KINDS.forEach(k=>S[k].sort((a,b)=>String(a.created_date).localeCompare(String(b.created_date))));S.citation.reverse()};
const safeId=s=>String(s).replace(/[^A-Za-z0-9_\-]/g,'-').slice(0,120);
async function getUser(){const{data}=await sb.auth.getSession();currentUser=data.session?data.session.user:null;return currentUser}
async function loadAll(){const next={style:[],type:[],citation:[],library:[],pinpoint:[]};const size=1000;
 for(let from=0;;from+=size){const{data,error}=await sb.from('records').select('kind,data').order('id').range(from,from+size-1);if(error)throw error;data.forEach(r=>{if(next[r.kind])next[r.kind].push(r.data)});if(data.length<size)break}
 KINDS.forEach(k=>S[k]=next[k]);sortAll();storeMode='db';storeNote=''}
async function put(kind,item){
 if(!item.id)item.id=uid();item.id=safeId(item.id);if(!item.created_date)item.created_date=new Date().toISOString();
 const i=S[kind].findIndex(x=>x.id===item.id);if(i>=0)S[kind][i]=item;else S[kind].push(item);sortAll();
 const{error}=await sb.from('records').upsert({user_id:currentUser.id,id:kind+'_'+item.id,kind,data:JSON.parse(JSON.stringify(item)),updated_at:new Date().toISOString()},{onConflict:'user_id,id'});
 if(error){toast('Could not save. Check your connection and try again.',1);throw error}
 return item}
async function del(kind,id){S[kind]=S[kind].filter(x=>x.id!==id);const{error}=await sb.from('records').delete().eq('user_id',currentUser.id).eq('id',kind+'_'+id);if(error)toast('Could not delete. Check your connection and try again.',1)}
async function signOut(){await sb.auth.signOut();location.reload()}
function authView(onDone,compact){let mode='in',email='',pw='',msg='',bad=false;const root=h('div',{class:compact?'authwrap small-auth':'authwrap'});
 const draw=()=>{const go2=async e=>{e.preventDefault();msg='';const r=mode==='in'?await sb.auth.signInWithPassword({email:email.trim(),password:pw}):await sb.auth.signUp({email:email.trim(),password:pw});
    if(r.error){bad=true;msg=mode==='in'?'That email and password do not match. Check them and try again.':r.error.message;draw();return}
    if(!r.data.session){bad=false;msg='Check your email for a confirmation link, then sign in here.';mode='in';draw();return}
    currentUser=r.data.session.user;onDone()};
  root.replaceChildren(h('form',{class:'card pad stack authcard',onsubmit:go2},h('div',{class:'row'},h('div',{class:'logo'},'📚'),h('div',null,h('h1',{style:'font-size:20px'},'LexCitr'),h('p',{class:'small muted'},mode==='in'?'Sign in to your citations':'Create your account'))),
   msg&&h('div',{class:bad?'err':'banner'},msg),
   h('div',null,h('label',{class:'l',for:'auth-email'},'Email'),h('input',{class:'in',id:'auth-email',type:'email',required:true,autocomplete:'username',value:email,oninput:e=>email=e.target.value})),
   h('div',null,h('label',{class:'l',for:'auth-pw'},'Password'),h('input',{class:'in',id:'auth-pw',type:'password',required:true,minLength:8,autocomplete:mode==='in'?'current-password':'new-password',value:pw,oninput:e=>pw=e.target.value})),
   h('button',{type:'submit',class:'btn pri w'},mode==='in'?'Sign In':'Create Account'),
   h('button',{type:'button',class:'btn ghost w',onclick:()=>{mode=mode==='in'?'up':'in';msg='';draw()}},mode==='in'?'Create an account':'I already have an account')))};
 draw();return root}
/* ---------- DOI lookup through Crossref ---------- */
async function crossrefLookup(doi,type){const clean=doi.trim().replace(/^https?:\/\/(dx\.)?doi\.org\//i,'').replace(/^doi:\s*/i,'');
 const res=await fetch('https://api.crossref.org/works/'+encodeURIComponent(clean));if(!res.ok)throw new Error('notfound');const m=(await res.json()).message;
 const people=a=>(a||[]).map(p=>[p.given,p.family].filter(Boolean).join(' ')||p.name||'').filter(Boolean).join('|||');
 const dp=m.issued&&m.issued['date-parts']&&m.issued['date-parts'][0];const year=dp&&dp[0]?String(dp[0]):'';const values={};
 (type.fields||[]).forEach(f=>{const n=f.name.toLowerCase();let v='';
  if(n.includes('short title')||n.includes('additional information'))return;
  if(n.includes('author'))v=people(m.author);else if(n.includes('editor'))v=people(m.editor);
  else if(n.includes('journal')||n.includes('publication'))v=(m['container-title']||[])[0]||'';
  else if(n.includes('title'))v=(m.title||[])[0]||'';else if(n.includes('year'))v=year;else if(n.includes('volume'))v=m.volume||'';else if(n.includes('issue'))v=m.issue||'';
  else if(n.includes('page'))v=(m.page||'').split('-')[0];else if(n==='doi')v=m.DOI||'';else if(n==='url')v=m.URL||'';else if(n.includes('publisher'))v=m.publisher||'';
  if(v)values[f.name]=String(v)});
 const info={};if((type.name||'').toLowerCase().includes('book')){if(m.publisher)info.publisher=m.publisher;if(year)info.year=year}
 return{values,info}}
/* ---------- citation formatting ---------- */
const isExtra=f=>f.name.toLowerCase().includes('additional information');
const byOrder=(a,b)=>(a.order||0)-(b.order||0);
const INFO=['translator','editor','other','edition','publisher','year'];
const wrap=(t,f)=>f.italic&&f.bold?`<b><i>${t}</i></b>`:f.italic?`<i>${t}</i>`:f.bold?`<b>${t}</b>`:t;
function joinMulti(v){if(!v.includes('|||'))return v;const a=v.split('|||').filter(x=>x.trim());return a.length===1?a[0]:a.length===2?`${a[0]} and ${a[1]}`:a.length===3?`${a[0]}, ${a[1]} and ${a[2]}`:`${a[0]} and others`}
function format(type,values,fmt,info,infoFmt,asHtml,volumeLast=true){
 if(!type||!type.fields)return'';
 const E=asHtml?esc:(x=>x),W=asHtml?wrap:(x=>x);
 let fs=type.fields.filter(f=>!isExtra(f)).slice().sort(byOrder);
 if(volumeLast){const v=fs.find(f=>f.name.toLowerCase()==='volume');if(v)fs=fs.filter(f=>f!==v).concat(v)}
 let out='';
 for(const f of fs){let v=(values||{})[f.name]||'';if(!v)continue;const F={...f,...((fmt||{})[f.id]||{})};if(/author|editor/i.test(f.name))v=v.replace(/ +/g,' ');v=joinMulti(v);
  if(v.includes('|')&&F.infix){const[a,b]=v.split('|');out+=`${E(F.prefix||'')}${W(E(a),F)}${E(F.infix)}${W(E(b),F)}${E(F.suffix||'')}`}
  else out+=`${E(F.prefix||'')}${W(E(v),F)}${E(F.suffix||'')}`}
 const add=[];INFO.forEach(k=>{if(info&&info[k])add.push(W(E(info[k]),(infoFmt||{})[k]||{}))});
 if(add.length)out+=` (${add.join(', ')}).`;
 return out}
const typeOf=c=>S.type.find(t=>t.id===c.source_type_id);
const typeName=id=>{const t=S.type.find(t=>t.id===id);return t?t.name:'Unknown'};
const libName=id=>{const l=S.library.find(l=>l.id===id);return l?l.name:'Uncategorized'};
const citeHtml=c=>{const t=typeOf(c);return t&&c.field_values?format(t,c.field_values,c.custom_formatting,c.additional_info,c.additional_info_formatting,true):(c.formatted_citation_html||esc(c.formatted_citation||''))};
/* ---------- bibliography (Primary / Institutional / Secondary grouping) ---------- */
function nameParts(n){const o=n.trim().split(' ').filter(Boolean);if(!o.length)return{last:'',initials:''};if(o.length===1)return{last:o[0],initials:''};const last=o.length>=3?o.slice(2).join(' '):o[o.length-1];const ini=(o.length>=3?o.slice(0,2):o.slice(0,1)).map(x=>x.charAt(0).toUpperCase()).join('');return{last,initials:ini}}
function authorList(v,comma=true){const a=(v.includes('|||')?v.split('|||'):[v]).map(x=>x.trim()).filter(Boolean).map(nameParts);a.sort((x,y)=>x.last.localeCompare(y.last));const T=({last,initials})=>`${last} ${initials}`.trim(),B=comma?',':'';return a.length===1?`${T(a[0])}${B}`:a.length===2?`${T(a[0])} and ${T(a[1])}${B}`:a.length===3?`${T(a[0])}, ${T(a[1])} and ${T(a[2])}${B}`:`${T(a[0])} and others${B}`}
function bibEntry(c){
 const n=typeName(c.source_type_id).toLowerCase(),t=typeOf(c),vals=c.field_values||{};
 const isShort=x=>x.toLowerCase().includes('short title');
 const fs=(t&&t.fields||[]).slice().sort(byOrder);
 const piece=(f,v)=>{const F={...f,...((c.custom_formatting||{})[f.id]||{})};return`${esc(F.prefix||'')}${wrap(esc(v),F)}${esc(F.suffix||'')}`};
 const info=keys=>{if(!c.additional_info)return'';const a=keys.filter(k=>c.additional_info[k]).map(k=>esc(c.additional_info[k]));return a.length?` (${a.join(', ')}).`:''};
 const authored=n.includes('authored book'),edited=n.includes('edited book')&&!n.includes('contribution'),contrib=n.includes('contribution'),journal=n.includes('journal article'),icsid=n.includes('icsid arbitration');
 let out='';
 if(authored||contrib||journal){let seen=false;for(const f of fs){if(isShort(f.name))continue;const v=vals[f.name];if(!v)continue;const l=f.name.toLowerCase();
   if(l.includes('author')&&!seen){seen=true;out+=esc(authorList(v,true))+' '}
   else if(l.includes('editor')&&contrib){const F={...f,...((c.custom_formatting||{})[f.id]||{})};out+=esc(F.prefix||'')+esc(authorList(v,false))+esc(F.suffix||'')}
   else out+=piece(f,v)}
  return out+info(INFO)}
 if(edited){let seen=false;for(const f of fs){if(isShort(f.name))continue;const v=vals[f.name];if(!v)continue;if(f.name.toLowerCase().includes('editor')&&!seen){seen=true;out+=esc(authorList(v,true))+' '}else out+=piece(f,v)}
  return out+info(['translator','other','edition','publisher','year'])}
 if(t&&t.fields){for(const f of fs){if(isShort(f.name))continue;if(icsid&&f.name.toLowerCase()==='reference')continue;const v=vals[f.name];if(v)out+=piece(f,v)}return out}
 return esc(c.formatted_citation||'')}
function bibSortKey(c){const n=typeName(c.source_type_id).toLowerCase(),t=typeOf(c),v=c.field_values||{};const edited=n.includes('edited book')&&!n.includes('contribution');
 const key=x=>{const a=(x.includes('|||')?x.split('|||'):[x]).map(y=>y.trim()).filter(Boolean).map(y=>nameParts(y).last.toLowerCase()).sort();return a[0]||''};
 const f=t&&t.fields&&t.fields.find(f=>f.name.toLowerCase().includes(edited?'editor':'author'));
 if(f&&v[f.name])return key(v[f.name]);return String(Object.values(v)[0]||'').toLowerCase()}
function groupName(name){const l=name.toLowerCase();return{'eu treaty':'EU Primary Law','eu case':'EU Case Law','south african case':'South African Case Law','icsid arbitration':'ICSID Arbitration','icj case':'ICJ Cases','icj cases':'ICJ Cases','icj advisory opinion':'ICJ Advisory Opinions','unga resolution':'UNGA Resolutions'}[l]||name}
function buildBibliography(list){
 const seenBib=new Set();list=list.filter(x=>{const k=String(x.formatted_citation||x.id).replace(/\s+/g,' ').trim().toLowerCase();if(seenBib.has(k))return false;seenBib.add(k);return true});
 const groups={};list.forEach(c=>{const g=groupName(typeName(c.source_type_id));(groups[g]=groups[g]||[]).push(c)});
 Object.values(groups).forEach(a=>a.sort((x,y)=>bibSortKey(x).localeCompare(bibSortKey(y))));
 const has=(s,arr)=>arr.some(x=>s.includes(x));const eu={},sa={},intl={},inst={},sec={},books={};
 Object.entries(groups).forEach(([g,a])=>{const l=g.toLowerCase();
  if(has(l,['eu primary law','eu treaty','eu secondary law','eu case law']))eu[g]=a;
  else if(has(l,['south african legislation','south african case law']))sa[g]=a;
  else if(has(l,['unga','icj advisory','icj cases','icsid arbitration']))intl[g]=a;
  else if(has(l,['institutional document','commission document','commission decision']))inst[g]=a;
  else if(has(l,['authored book','contribution to edited book','edited book']))books[g]=a;
  else sec[g]=a});
 const ordered=(o,r)=>Object.entries(o).sort(([a],[b])=>(r[a]||999)-(r[b]||999));
 const alpha=o=>Object.entries(o).sort(([a],[b])=>a.localeCompare(b));
 const E=ordered(eu,{'EU Primary Law':1,'EU Secondary Law':2,'EU Case Law':3}),A=ordered(sa,{'South African Legislation':1,'South African Case Law':2}),I=ordered(intl,{'UNGA Resolutions':1,'ICJ Advisory Opinions':2,'ICJ Cases':3,'ICSID Arbitration':4}),B=ordered(books,{'Authored Book':1,'Contribution to Edited Book':2,'Edited Book':3}),Sx=alpha(sec),Nx=alpha(inst);
 const htmlG=g=>g.map(([n,a])=>`<div style="margin-bottom: 20px;"><h4 style="font-weight: bold; margin-bottom: 12px; font-size: 12pt;">${esc(n)}</h4>${a.map(c=>`<p style="margin-bottom: 12px;">${bibEntry(c)}</p>`).join('')}</div>`).join('');
 const textG=g=>g.map(([n,a])=>`${n}\n\n${a.map(c=>unhtml(bibEntry(c))).join('\n\n')}\n\n`).join('');
 const H3=t=>`<h3 style="font-weight: bold; margin-bottom: 16px; font-size: 12pt;">${t}</h3>`,H4=t=>`<h4 style="font-weight: bold; margin-bottom: 12px; font-size: 12pt;">${t}</h4>`,sec3=x=>`<div style="margin-bottom: 32px;">${x}</div>`;
 let html=`<div style="font-family: 'Times New Roman', Times, serif; font-size: 12pt; line-height: 1.5; text-align: justify;"><h2 style="font-weight: bold; margin-bottom: 24px; text-align: left; font-size: 12pt;">Bibliography</h2>`,text='Bibliography\n\n';
 if(E.length+A.length+I.length){html+=sec3(H3('Primary Sources')+htmlG(E)+htmlG(A)+(I.length?`<div style="margin-bottom: 20px;">${H4('International Law')}${htmlG(I)}</div>`:''));text+='Primary Sources\n\n'+textG(E)+textG(A)+(I.length?'International Law\n\n'+textG(I):'')}
 if(Nx.length){html+=sec3(H3('Institutional Documents')+htmlG(Nx));text+='Institutional Documents\n\n'+textG(Nx)}
 if(Sx.length+B.length){html+=sec3(H3('Secondary Sources')+htmlG(Sx)+(B.length?`<div style="margin-bottom: 20px;">${H4('Books')}${htmlG(B)}</div>`:''));text+='Secondary Sources\n\n'+textG(Sx)+(B.length?'Books\n\n'+textG(B):'')}
 return{html:html+'</div>',text}}
/* ---------- templates ---------- */
const T=(name,style,fs)=>({name,style,fields:fs.map((f,i)=>({id:'t'+(i+1),name:f[0],prefix:f[1]||'',suffix:f[2]||'',italic:!!f[3],bold:false,order:i}))});
const TEMPLATES=[
T('Journal Article','OSCOLA',[['Author','',', '],['Title',"'","' "],['Year','(',') '],['Volume Number','',' '],['Journal Name or Abbreviation','',' '],['First Page of Article','','.']]),
T('Authored Book','OSCOLA',[['Author','',', '],['Title','','',1],['Volume','vol ',''],['Additional Information','(',').']]),
T('Edited Book (No Author)','OSCOLA',[['Editor','(','), '],['Title','','',1],['Volume','vol ',''],['Additional Information','(',').']]),
T('EU Secondary Law','OSCOLA',[['Legislation Type','',' '],['Legislation Number','',' '],['Legislation Title','',' '],['Year','[','] '],['OJ Series (Eg. C, L, etc.)','OJ ',' '],['OJ Issue','','/'],['OJ First Page','',' '],['Short Title','(',').']]),
T('EU Case','OSCOLA',[['Case Type (C or T)','',' '],['Case Number','',' '],['Case Name','',' ',1],['Year','[','] '],['Report Abbreviation','',' '],['Volume Number (I or II)','','-'],['First Page','',' '],['Short Title','(',').']]),
T('Commission Decision','OSCOLA',[['Case Name (Parties)','',' '],['Case Number','(',') '],['Decision Number','Commission Decision',' '],['Year','[','] '],['OJ','',' '],['OJ Letter','',''],['Issue','','/'],['First Page','','.']]),
T('South African Case','OSCOLA',[['Parties','',' ',1],['Year Reported','',' '],['Report Volume','(',') '],['Name of Law Report Abbreviated','',' '],['First Page','',' '],['Court of Instance','(',')']]),
T('South African Legislation','OSCOLA',[['Act Name','',' ',1],['Act Year',' of ','.']]),
T('ICSID Arbitration','OSCOLA',[['Case Name','',', ',1],['Case Number','ICSID Case No. ',', '],['Document Type','',', '],['Year','',', '],['Reference','ICSID Rep ','.']]),
T('ICJ Cases','OSCOLA',[['Title','',' ',1],['Parties','(','), ',1],['Type of Hearing','',', '],['Year of Reporter','I.C.J. Rep. ',' '],['Date','(',')']]),
T('ICJ Advisory Opinion','OSCOLA',[['Title','',', ',1],['Advisory Opinion','',', '],['Volume Year','I.C.J. Rep ',', '],['Page','p. ','.']]),
T('UNGA Resolution','OSCOLA',[['Resolution Number','G.A. Res. ',', '],['Session','U.N. GAOR, ','., '],['UN DOC','U.N. Doc. ',' '],['Year','(',').']]),
T('Blog','OSCOLA',[['Author','',', '],['Title',"'","' "],['Blog Name','(',', ',1],['Published Date','',') '],['URL','<','>']]),
T('Commission Documents','OSCOLA',[['Body Producing Document','',', '],['Title',"'","' "],['Description','(',')'],['COM Number','COM','']]),
T('SA Comp Commission Publication','OSCOLA',[['Author','',', '],['Title','',' ',1],['Additional Information','(',', '],['Publisher','',' '],['Year','',').']]),
T('Institutional Document (e.g. OECD)','OSCOLA',[['Author','',', '],['Title',"'","' "],['Additional Info','(',')']]),
T('Institutional Report (e.g. OECD Recommendation)','MLA',[['Full Name of Institution','',', '],['Title of Publication','',' ',1],['Document Number','',' '],['Date','(adopted ',') '],['URL','<','>']])];
const stamp=fields=>fields.map((f,i)=>({...f,id:Date.now().toString(36)+i+Math.random().toString(36).slice(2,5),infix:f.infix||'',order:i}));
