'use strict';
/* LexCitr panel inside Word */
const P={tab:'saved',q:'',lib:'all',sel:null,pin:'',style:null,type:null,saveLib:'uncategorized',values:{},info:{},inWord:false,busy:false};
const root=()=>document.getElementById('app');
/* turn formatted citation HTML into text runs so italics and bold survive in Word */
function runsFromHtml(html){const d=document.createElement('div');d.innerHTML=html;const out=[];
 const walk=(n,i,b)=>{n.childNodes.forEach(c=>{if(c.nodeType===3){if(c.nodeValue)out.push({text:c.nodeValue,italic:i,bold:b})}else{const t=c.tagName;walk(c,i||t==='I'||t==='EM',b||t==='B'||t==='STRONG')}})};walk(d,false,false);return out}
function withPin(html,text,pin){const p=pin.trim();if(!p)return{html,text};return{html:html.replace(/\.\s*$/,'')+', '+esc(p)+'.',text:text.replace(/\.\s*$/,'')+', '+p+'.'}}
async function insertFootnote(html){await Word.run(async ctx=>{const at=ctx.document.getSelection().getRange('End');const note=at.insertFootnote('');
  for(const r of runsFromHtml(html)){const t=note.body.insertText(r.text,'End');t.font.italic=r.italic;t.font.bold=r.bold}
  await ctx.sync()})}
async function placeCitation(c,pin){if(!P.inWord){toast('Open this panel from the LexCitr button in Word to insert footnotes.',1);return false}
 const x=withPin(citeHtml(c),c.formatted_citation||'',pin);P.busy=true;
 try{await insertFootnote(x.html);put('pinpoint',{citation_id:c.id,pinpoint:pin.trim(),formatted_footnote:x.text}).catch(()=>{});toast('Footnote inserted');return true}
 catch(e){console.error(e);const ok=await copyRich(`<p>${x.html}</p>`,x.text);toast(ok?'Word refused the footnote, so it was copied instead. Paste it into a footnote.':'Word refused the footnote. Try again with your cursor in the document body.',1);return false}
 finally{P.busy=false}}
/* views */
function draw(){const top=document.querySelector('.tp-body'),s=top?top.scrollTop:0;root().replaceChildren(h('div',{class:'tp'},
  h('div',{class:'tp-head'},h('b',null,'📚 LexCitr'),h('div',{class:'row',style:'gap:4px'},btn('Refresh',refresh,'xs'),btn('Log Out',signOut,'xs'))),
  storeNote&&h('div',{class:'banner',style:'margin:12px 12px 0'},storeNote),
  !P.inWord&&h('div',{class:'banner',style:'margin:12px 12px 0'},'Open this panel from the LexCitr button in Word to insert footnotes.'),
  h('div',{style:'padding:12px 12px 0'},h('div',{class:'tabs'},[['saved','Saved citations'],['new','New citation']].map(([k,l])=>h('button',{type:'button',class:P.tab===k?'on':'',onclick:()=>{P.tab=k;draw()}},l)))),
  P.tab==='saved'?savedView():newView()));const b=document.querySelector('.tp-body');if(b)b.scrollTop=s}
function matches(c){if(P.lib==='uncategorized'&&c.library_id)return false;if(P.lib!=='all'&&P.lib!=='uncategorized'&&c.library_id!==P.lib)return false;const q=P.q.toLowerCase();return!q||(c.formatted_citation||'').toLowerCase().includes(q)||typeName(c.source_type_id).toLowerCase().includes(q)}
function savedView(){const list=h('div',{class:'stack t'});const fill=()=>{const m=S.citation.filter(matches);list.replaceChildren(...(m.length?m.map(c=>h('button',{type:'button',class:'pick'+(P.sel===c.id?' on':''),onclick:()=>{P.sel=c.id;P.pin='';draw()}},h('span',{class:'small',style:'display:block;color:var(--accent-fg);font-weight:500;margin-bottom:2px'},typeName(c.source_type_id)),h('span',{class:'cite',html:citeHtml(c)}))):[h('p',{class:'muted',style:'text-align:center;padding:20px 8px'},S.citation.length?'No citations match your search.':'No citations yet. Use New citation to add your first one.')]))};fill();
 const sel=P.sel&&S.citation.find(c=>c.id===P.sel);const prev=h('div',{class:'cite',style:'padding:8px;background:var(--surface2);border-radius:6px'});const updatePrev=()=>{if(sel)prev.innerHTML=withPin(citeHtml(sel),'',P.pin).html};updatePrev();
 const go2=async()=>{if(await placeCitation(sel,P.pin)){P.pin='';draw()}};
 return[h('div',{class:'tp-body'},h('select',{class:'in',id:'tp-lib','aria-label':'Library',onchange:e=>{P.lib=e.target.value;P.sel=null;draw()}},h('option',{value:'all',selected:P.lib==='all'},'All libraries'),h('option',{value:'uncategorized',selected:P.lib==='uncategorized'},'Uncategorized'),S.library.map(l=>h('option',{value:l.id,selected:P.lib===l.id},l.name))),
   h('input',{class:'in',id:'tp-search',type:'search',placeholder:'Search citations...','aria-label':'Search citations',value:P.q,oninput:e=>{P.q=e.target.value;fill()}}),list),
  sel&&h('div',{class:'tp-foot'},prev,h('input',{class:'in',id:'tp-pin',placeholder:'Pinpoint (optional), e.g. para 12','aria-label':'Pinpoint',value:P.pin,oninput:e=>{P.pin=e.target.value;updatePrev()},onkeydown:e=>{if(e.key==='Enter')go2()}}),btn('Insert Footnote',go2,'pri w','file'))]}
function nameInputs(f){const raw=P.values[f.name]||'',parts=raw.split('|||');const set=()=>{P.values[f.name]=parts.join('|||');upd()};
 return h('div',{class:'stack t'},parts.map((p,i)=>{const w=p.split(' ');const cur=[w[0]||'',w[1]||'',w.slice(2).join(' ')];return h('div',{class:'names'},['First name','Middle name/initial','Last name'].map((ph,k)=>h('input',{class:'in',placeholder:ph,'aria-label':`${f.name} ${ph}`,value:cur[k],oninput:e=>{cur[k]=e.target.value;parts[i]=cur.join(' ').trim();set()}})),parts.length>1&&btn('Remove',()=>{parts.splice(i,1);P.values[f.name]=parts.join('|||');draw()},'xs red'))}),h('div',null,btn('Add another person',()=>{P.values[f.name]=raw+'|||';draw()},'xs','plus')))}
let upd=()=>{};
function newView(){const styles=S.style;if(!styles.length)return[h('div',{class:'tp-body'},h('div',{class:'empty'},h('p',null,'Create a citation style on the LexCitr website first, then press Refresh.')))];
 if(!P.style||!styles.find(s=>s.id===P.style))P.style=styles[0].id;const types=S.type.filter(t=>t.citation_style_id===P.style);if(!types.find(t=>t.id===P.type))P.type=null;const t=types.find(x=>x.id===P.type);
 const prev=h('div',{class:'cite',style:'padding:8px;background:var(--surface2);border-radius:6px;min-height:36px'});const html=()=>t?format(t,P.values,{},P.info,{},true):'';upd=()=>{const x=html();if(x)prev.innerHTML=x;else{prev.textContent='Fill in the fields to see your citation.'}};
 const saveAnd=async ins=>{const x=format(t,P.values,{},P.info,{},false);if(!x){toast('Please fill in at least one field',1);return}
   const c=await put('citation',{source_type_id:t.id,citation_style_id:P.style,library_id:P.saveLib==='uncategorized'?null:P.saveLib,field_values:P.values,custom_formatting:{},additional_info:P.info,additional_info_formatting:{},formatted_citation:x,formatted_citation_html:html(),notes:''});
   if(ins)await placeCitation(c,P.pin);else toast('Citation saved');P.values={};P.info={};P.pin='';draw()};
 const doiImport=async()=>{const d=document.getElementById('tp-doi').value;if(!d.trim())return;try{const r=await crossrefLookup(d,t);Object.assign(P.values,r.values);Object.assign(P.info,r.info);draw();toast('Fields filled. Check them against your source.')}catch(e){toast('No record was found for that DOI. Check it, or fill in the fields by hand.',1)}};
 const body=h('div',{class:'tp-body'},
  styles.length>1&&field('Citation style',h('select',{class:'in',id:'tp-style',onchange:e=>{P.style=e.target.value;P.type=null;P.values={};P.info={};draw()}},styles.map(s=>h('option',{value:s.id,selected:s.id===P.style},s.name)))),
  field('Source type',h('select',{class:'in',id:'tp-type',onchange:e=>{P.type=e.target.value;P.values={};P.info={};draw()}},h('option',{value:'',selected:!t},'Choose a source type...'),types.map(x=>h('option',{value:x.id,selected:x.id===P.type},x.name)))),
  t&&field('Save to library',h('select',{class:'in',id:'tp-savelib',onchange:e=>P.saveLib=e.target.value},h('option',{value:'uncategorized',selected:P.saveLib==='uncategorized'},'Uncategorized'),S.library.map(l=>h('option',{value:l.id,selected:P.saveLib===l.id},l.name)))),
  t&&h('div',{class:'row'},h('input',{class:'in grow',id:'tp-doi',placeholder:'DOI (optional)','aria-label':'DOI'}),btn('Import',doiImport,'sm soft')),
  t&&(t.fields||[]).filter(f=>!isExtra(f)).slice().sort(byOrder).map(f=>field(f.name,/author|editor/i.test(f.name)?nameInputs(f):h('input',{class:'in','aria-label':f.name,value:P.values[f.name]||'',oninput:e=>{P.values[f.name]=e.target.value;upd()}}))),
  t&&(t.name||'').toLowerCase().includes('book')&&[['translator','Translator'],['editor','Editor'],['other','Other information'],['edition','Edition'],['publisher','Publisher'],['year','Year']].filter(([k])=>!(k==='editor'&&t.name.toLowerCase().includes('edited'))).map(([k,l])=>field(l,h('input',{class:'in','aria-label':'Additional '+l,value:P.info[k]||'',oninput:e=>{P.info[k]=e.target.value;upd()}}))));
 upd();
 return[body,t&&h('div',{class:'tp-foot'},prev,h('input',{class:'in',id:'tp-newpin',placeholder:'Pinpoint (optional), e.g. para 12','aria-label':'Pinpoint',value:P.pin,oninput:e=>P.pin=e.target.value}),btn('Save and Insert Footnote',()=>saveAnd(true),'pri w','file'),btn('Save Only',()=>saveAnd(false),'w','save'))]}
async function refresh(){try{await loadAll();draw();toast('Citations refreshed')}catch(e){toast('Could not refresh. Check your connection.',1)}}
function field(label,el){return h('div',null,h('label',{class:'l s'},label),el)}
async function start(){root().replaceChildren(h('p',{class:'muted',style:'padding:16px'},'Loading your citations…'));try{await loadAll()}catch(e){storeNote='Your citations could not be loaded. Check your connection and press Refresh.'}draw()}
window.addEventListener('focus',async()=>{if(!currentUser)return;const a=document.activeElement;if(a&&a.matches&&a.matches('input,select,textarea'))return;try{await loadAll();draw()}catch(e){}});
function detectWord(){return new Promise(res=>{if(!window.Office||!Office.onReady){res(false);return}let done=false;const fin=v=>{if(!done){done=true;res(v)}};Office.onReady(i=>fin(!!(i&&i.host===Office.HostType.Word&&Office.context.requirements.isSetSupported('WordApi','1.5'))));setTimeout(()=>fin(false),4000)})}
(async()=>{P.inWord=await detectWord();const u=await getUser();if(u)start();else root().replaceChildren(authView(start,true))})();
