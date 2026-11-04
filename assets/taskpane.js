'use strict';
/* LexCitr panel inside Word: one list, with new citations, source types and projects added in place */
const P={screen:'list',project:'all',q:'',sel:null,pin:'',style:null,type:null,typeQ:'',values:{},info:{},showMore:false,newProject:false,inWord:false};
const E={id:null,name:'',fields:[],styleId:null,templates:true};
const root=()=>document.getElementById('app');
const go=(s,o)=>{P.screen=s;if(o)Object.assign(P,o);draw()};
/* formatted citation HTML becomes text runs, so italics and bold survive in Word */
function runsFromHtml(html){const d=document.createElement('div');d.innerHTML=html;const out=[];
 const walk=(n,i,b)=>{n.childNodes.forEach(c=>{if(c.nodeType===3){if(c.nodeValue)out.push({text:c.nodeValue,italic:i,bold:b})}else{const t=c.tagName;walk(c,i||t==='I'||t==='EM',b||t==='B'||t==='STRONG')}})};walk(d,false,false);return out}
function withPin(html,text,pin){const p=pin.trim();if(!p)return{html,text};return{html:html.replace(/[.\s]*$/,'')+', '+esc(p)+'.',text:text.replace(/[.\s]*$/,'')+', '+p+'.'}}
/* ---------- footnote tracking ----------
   Each LexCitr footnote holds a hidden content control tagged with a key. Saved in the document:
   key -> citation id, pinpoint and the text last written, so repeats can be recalculated in order. */
const TAG='lexcitr:',SET='lexcitrNotes';
const norm=x=>String(x||'').replace(/\s+/g,' ').trim();
const plainOf=html=>unhtml(html);
function surnames(v){const a=v.split('|||').map(x=>x.trim()).filter(Boolean).map(x=>nameParts(x).last).filter(Boolean);return a.length===1?a[0]:a.length===2?`${a[0]} and ${a[1]}`:a.length===3?`${a[0]}, ${a[1]} and ${a[2]}`:a.length?`${a[0]} and others`:''}
function shortForm(c){const t=typeOf(c),vals=c.field_values||{},fs=(t&&t.fields||[]).slice().sort(byOrder),fmt=c.custom_formatting||{};const F=f=>({...f,...(fmt[f.id]||{})});
 const who=fs.find(f=>/author/i.test(f.name)&&vals[f.name])||fs.find(f=>/editor/i.test(f.name)&&vals[f.name]);if(who){const n=surnames(vals[who.name]);if(n)return esc(n)}
 const pick=fs.find(f=>/short title/i.test(f.name)&&vals[f.name])||fs.find(f=>/case name|parties|act name/i.test(f.name)&&vals[f.name])||fs.find(f=>/title/i.test(f.name)&&vals[f.name])||fs.find(f=>vals[f.name]);
 return pick?wrap(esc(joinMulti(vals[pick.name])),F(pick)):esc(c.formatted_citation||'')}
function noteHtml(c,pin,prev,first){const p=(pin||'').trim();
 if(prev&&prev.c===c.id)return p&&p!==(prev.p||'').trim()?`ibid, ${esc(p)}.`:'ibid.';
 if(first)return`${shortForm(c)} (fn ${first})`+(p?`, ${esc(p)}.`:'.');
 return withPin(citeHtml(c),'',p).html}
async function readNotes(ctx){const notes=ctx.document.body.footnotes;notes.load('items');await ctx.sync();
 notes.items.forEach(n=>n.body.contentControls.load('items/tag,items/text'));const st=ctx.document.settings.getItemOrNullObject(SET);st.load('value');await ctx.sync();
 const raw=st.isNullObject?null:st.value;let map={};try{map=raw?(typeof raw==='string'?JSON.parse(raw):raw):{}}catch(e){map={}}return{notes:notes.items,map}}
const tracked=n=>n.body.contentControls.items.find(c=>(c.tag||'').startsWith(TAG));
function writeRuns(cc,html){const rs=runsFromHtml(html);rs.forEach((r,i)=>{const t=cc.insertText(r.text,i===0?'Replace':'End');t.font.italic=r.italic;t.font.bold=r.bold})}
/* recalculate every LexCitr footnote in document order; hand-edited ones are left alone and counted */
async function syncNotes(ctx){const{notes,map}=await readNotes(ctx);const first={},used={},live={};let prev=null,edited=0,changed=0;
 notes.forEach((n,idx)=>{const cc=tracked(n),num=idx+1;if(!cc){prev=null;return}const key=cc.tag.slice(TAG.length),rec=map[key];if(!rec){prev=null;return}live[key]=rec;(used[rec.c]=used[rec.c]||[]).push(num);
  const c=S.citation.find(x=>x.id===rec.c);if(c){const html=noteHtml(c,rec.p,prev,first[rec.c]),text=plainOf(html);
    if(norm(text)!==norm(cc.text)){if(rec.t==null||norm(cc.text)===norm(rec.t)){writeRuns(cc,html);rec.t=text;changed++}else edited++}else rec.t=text}
  if(!first[rec.c])first[rec.c]=num;prev={c:rec.c,p:rec.p}});
 ctx.document.settings.add(SET,JSON.stringify(live));await ctx.sync();P.used=used;return{changed,edited,used}}
async function scanDoc(){if(!P.inWord)return;try{await Word.run(async ctx=>{const{notes,map}=await readNotes(ctx);const used={};notes.forEach((n,i)=>{const cc=tracked(n),rec=cc&&map[cc.tag.slice(TAG.length)];if(rec)(used[rec.c]=used[rec.c]||[]).push(i+1)});P.used=used})}catch(e){console.error(e)}}
async function placeCitation(c,pin){if(!P.inWord){toast('Open this panel from the LexCitr add-in in Word to insert footnotes.',1);return false}
 const p=(pin||'').trim();
 try{await Word.run(async ctx=>{const{map}=await readNotes(ctx);const key=uid();map[key]={c:c.id,p};
    const note=ctx.document.getSelection().getRange('End').insertFootnote('');const r=note.body.insertText('…','End');const cc=r.insertContentControl();cc.tag=TAG+key;cc.appearance='Hidden';
    ctx.document.settings.add(SET,JSON.stringify(map));await ctx.sync();await syncNotes(ctx)});
  const x=withPin(citeHtml(c),c.formatted_citation||'',p);put('pinpoint',{citation_id:c.id,pinpoint:p,formatted_footnote:x.text}).catch(()=>{});toast('Footnote inserted');return true}
 catch(e){console.error(e);const x=withPin(citeHtml(c),c.formatted_citation||'',p);const ok=await copyRich(`<p>${x.html}</p>`,x.text);toast(ok?'Word refused the footnote, so the full citation was copied instead. Paste it into a footnote.':'Word refused the footnote. Click in the document body and try again.',1);return false}}
async function updateNotes(){if(!P.inWord){toast('Open this panel from the LexCitr add-in in Word to update footnotes.',1);return}
 try{await loadAll();let r;await Word.run(async ctx=>{r=await syncNotes(ctx)});toast(`${r.changed} footnote${r.changed===1?'':'s'} updated`+(r.edited?`. ${r.edited} edited by hand left unchanged.`:'.'),r.edited>0);draw()}catch(e){console.error(e);toast('Footnotes could not be updated. Try again.',1)}}
async function insertBibliography(){if(!P.inWord){toast('Open this panel from the LexCitr add-in in Word to insert a bibliography.',1);return}
 try{await scanDoc();const list=S.citation.filter(c=>(P.used||{})[c.id]);if(!list.length){toast('No LexCitr footnotes in this document yet.',1);return}const r=buildBibliography(list);await Word.run(async ctx=>{ctx.document.getSelection().getRange('End').insertHtml(r.html,'After');await ctx.sync()});toast('Bibliography inserted at your cursor');draw()}catch(e){console.error(e);toast('The bibliography could not be inserted. Try again.',1)}}
const field=(label,el)=>h('div',null,h('label',{class:'l s'},label),el);
const backBar=(title,on)=>h('div',{class:'tp-back'},h('button',{type:'button',class:'btn ghost sm',onclick:on},icon('back'),'Back'),h('b',null,title));
/* frame */
function draw(){const top=document.querySelector('.tp-body'),s=top?top.scrollTop:0;
 const view=P.screen==='pick'?pickView():P.screen==='form'?formView():P.screen==='type'?typeView():listView();
 root().replaceChildren(h('div',{class:'tp'},
  h('div',{class:'tp-head'},h('b',null,'📚 LexCitr'),h('div',{class:'row',style:'gap:4px'},btn('Refresh',refresh,'xs'),btn('Log Out',signOut,'xs'))),
  projectBar(),
  storeNote&&h('div',{class:'banner tp-note'},storeNote),
  !P.inWord&&h('div',{class:'banner tp-note'},'Open this panel from the LexCitr add-in in Word to insert footnotes.'),
  view));const b=document.querySelector('.tp-body');if(b&&P.keepScroll!==false)b.scrollTop=s;P.keepScroll=true}
function projectBar(){
 if(P.newProject){let name='';const create=async()=>{if(!name.trim())return;const l=await put('library',{name:name.trim(),color:'indigo'});P.project=l.id;P.newProject=false;toast('Project created');draw()};
  const inp=h('input',{class:'in grow',id:'tp-newproject',placeholder:'Project name','aria-label':'Project name',oninput:e=>{name=e.target.value},onkeydown:e=>{if(e.key==='Enter')create();if(e.key==='Escape'){P.newProject=false;draw()}}});setTimeout(()=>inp.focus(),0);
  return h('div',{class:'tp-project'},inp,btn('Create',create,'sm pri'),btn('Cancel',()=>{P.newProject=false;draw()},'sm ghost'))}
 return h('div',{class:'tp-project'},h('select',{class:'in grow',id:'tp-lib','aria-label':'Project',onchange:e=>{P.project=e.target.value;P.sel=null;draw()}},h('option',{value:'all',selected:P.project==='all'},'All projects'),S.library.map(l=>h('option',{value:l.id,selected:P.project===l.id},l.name)),h('option',{value:'uncategorized',selected:P.project==='uncategorized'},'No project')),btn('New Project',()=>{P.newProject=true;draw()},'sm','plus'))}
/* list */
function matches(c){if(P.project==='uncategorized'&&c.library_id)return false;if(P.project!=='all'&&P.project!=='uncategorized'&&c.library_id!==P.project)return false;const q=P.q.toLowerCase();return!q||(c.formatted_citation||'').toLowerCase().includes(q)||typeName(c.source_type_id).toLowerCase().includes(q)}
function listView(){const list=h('div',{class:'stack t'});
 const fill=()=>{const m=S.citation.filter(matches);list.replaceChildren(...(m.length?m.map(c=>h('button',{type:'button',class:'pick'+(P.sel===c.id?' on':''),onclick:()=>{P.sel=P.sel===c.id?null:c.id;P.pin='';draw()}},h('span',{class:'tp-kind'},typeName(c.source_type_id),(P.used||{})[c.id]&&h('span',{class:'tp-used'},'fn '+P.used[c.id].join(', '))),h('span',{class:'cite',html:citeHtml(c)}))):[h('p',{class:'muted',style:'text-align:center;padding:20px 8px'},S.citation.length&&(P.q||P.project!=='all')?'No citations here.':'No citations yet. Add your first one above.')]))};fill();
 const sel=P.sel&&S.citation.find(c=>c.id===P.sel);const prev=h('div',{class:'cite tp-prev'});const up=()=>{if(sel)prev.innerHTML=withPin(citeHtml(sel),'',P.pin).html};up();
 const ins=async()=>{if(await placeCitation(sel,P.pin)){P.pin='';P.sel=null;draw()}};
 return[h('div',{class:'tp-body'},btn('New Citation',()=>go('pick',{typeQ:''}),'pri w','plus'),h('div',{class:'row'},h('span',{class:'grow'},btn('Update Footnotes',updateNotes,'sm w')),h('span',{class:'grow'},btn('Bibliography',insertBibliography,'sm w'))),h('input',{class:'in',id:'tp-search',type:'search',placeholder:'Search citations','aria-label':'Search citations',value:P.q,oninput:e=>{P.q=e.target.value;fill()}}),list),
  sel&&h('div',{class:'tp-foot'},prev,h('input',{class:'in',id:'tp-pin',placeholder:'Pinpoint (optional), e.g. para 12','aria-label':'Pinpoint',value:P.pin,oninput:e=>{P.pin=e.target.value;up()},onkeydown:e=>{if(e.key==='Enter')ins()}}),h('p',{class:'small muted'},(P.used||{})[sel.id]?'Already cited. It will go in as ibid or a (fn) reference, depending on where your cursor is.':'First use goes in as the full citation.'),btn('Cite',ins,'pri w','file'))]}
/* choose source type */
function openEditor(t){Object.assign(E,{id:t?t.id:null,name:t?t.name:'',fields:t?clone2(t.fields||[]).sort(byOrder):[],styleId:t?t.citation_style_id:P.style,templates:!t});go('type')}
const clone2=o=>JSON.parse(JSON.stringify(o));
async function addOscolaPanel(){const st=await put('style',{name:'OSCOLA',description:'For legal citations'});for(const t of TEMPLATES.filter(t=>t.style==='OSCOLA'))await put('type',{name:t.name,citation_style_id:st.id,fields:stamp(t.fields)});P.style=st.id;toast('OSCOLA added');draw()}
function pickView(){const styles=S.style,back=backBar('Choose source type',()=>go('list'));
 if(!styles.length){let name='';const mk=async()=>{if(!name.trim())return;const s=await put('style',{name:name.trim(),description:''});P.style=s.id;draw()};return[back,h('div',{class:'tp-body'},h('p',{class:'muted'},'A citation style holds your source types. Add OSCOLA, or name your own style.'),btn('Add OSCOLA with its source types',addOscolaPanel,'pri w','plus'),h('div',{class:'row'},h('input',{class:'in grow',id:'tp-stylename',placeholder:'Style name','aria-label':'Style name',oninput:e=>name=e.target.value}),btn('Create',mk,'sm')))]}
 if(!styles.find(s=>s.id===P.style))P.style=styles[0].id;
 const all=S.type.filter(t=>t.citation_style_id===P.style),list=h('div',{class:'stack t'});
 const fill=()=>{const q=P.typeQ.toLowerCase(),m=all.filter(t=>t.name.toLowerCase().includes(q));list.replaceChildren(...(m.length?m.map(t=>h('div',{class:'tp-row'},h('button',{type:'button',class:'tp-rowmain',onclick:()=>go('form',{type:t.id,values:{},info:{},pin:'',showMore:false,keepScroll:false})},t.name),ibtn('edit',()=>openEditor(t),'Edit '+t.name))):[h('p',{class:'muted',style:'text-align:center;padding:16px'},'No source types found.')]))};fill();
 return[back,h('div',{class:'tp-body'},styles.length>1&&field('Citation style',h('select',{class:'in',id:'tp-style',onchange:e=>{P.style=e.target.value;draw()}},styles.map(s=>h('option',{value:s.id,selected:s.id===P.style},s.name)))),btn('New Source Type',()=>openEditor(null),'w','plus'),h('input',{class:'in',id:'tp-typesearch',type:'search',placeholder:'Search source types','aria-label':'Search source types',value:P.typeQ,oninput:e=>{P.typeQ=e.target.value;fill()}}),list)]}
/* source type editor */
function typeView(){const titleBack=backBar(E.id?'Edit source type':'New source type',()=>go('pick'));
 if(E.templates&&!E.id){return[titleBack,h('div',{class:'tp-body'},btn('Start from blank',()=>{E.templates=false;E.fields=[{id:uid(),name:'Author',prefix:'',infix:'',suffix:', ',italic:false,bold:false,order:0},{id:uid(),name:'Title',prefix:'',infix:'',suffix:'.',italic:false,bold:false,order:1}];draw()},'w','plus'),h('h3',{class:'tp-sub'},'Or start from a template'),h('div',{class:'stack t'},TEMPLATES.map(t=>h('div',{class:'tp-row'},h('button',{type:'button',class:'tp-rowmain',onclick:()=>{E.name=t.name;E.fields=stamp(t.fields);E.templates=false;draw()}},t.name)))))]}
 const prev=h('div',{class:'cite tp-prev'});const sample=()=>{const v={};E.fields.forEach(f=>v[f.name]=f.name);return format({fields:E.fields.map((f,i)=>({...f,order:i}))},v,{},{}, {},true)};const up=()=>{prev.innerHTML=sample()||'Add a field to see the format.'};
 const move=(i,d)=>{const j=i+d;if(j<0||j>=E.fields.length)return;[E.fields[i],E.fields[j]]=[E.fields[j],E.fields[i]];draw()};
 const row=(f,i)=>h('div',{class:'tp-fieldbox'},h('div',{class:'row',style:'gap:4px'},h('input',{class:'in grow','aria-label':'Field name',placeholder:'Field name',value:f.name,oninput:e=>{f.name=e.target.value;up()}}),ibtn('up',()=>move(i,-1),'Move up'),ibtn('down',()=>move(i,1),'Move down'),ibtn('trash',()=>{E.fields=E.fields.filter(x=>x!==f);draw()},'Remove field','red')),
   h('div',{class:'row',style:'gap:4px'},h('input',{class:'in mono grow','aria-label':f.name+' text before',placeholder:'Before',value:f.prefix||'',oninput:e=>{f.prefix=e.target.value;up()}}),h('input',{class:'in mono grow','aria-label':f.name+' text after',placeholder:'After',value:f.suffix||'',oninput:e=>{f.suffix=e.target.value;up()}}),h('button',{type:'button',class:'tog s'+(f.italic?' on':''),title:'Italic',onclick:()=>{f.italic=!f.italic;draw()}},h('i',null,'I')),h('button',{type:'button',class:'tog s'+(f.bold?' on':''),title:'Bold',onclick:()=>{f.bold=!f.bold;draw()}},h('b',null,'B'))));
 const save=async()=>{if(!E.name.trim()){toast('Please name this source type',1);return}const existing=E.id&&S.type.find(t=>t.id===E.id);const t=await put('type',{...(existing||{}),name:E.name.trim(),citation_style_id:E.styleId,fields:E.fields.map((f,i)=>({...f,order:i}))});
   if(existing)for(const c of S.citation.filter(c=>c.source_type_id===t.id))await put('citation',{...c,formatted_citation:format(t,c.field_values,c.custom_formatting,c.additional_info,c.additional_info_formatting,false),formatted_citation_html:format(t,c.field_values,c.custom_formatting,c.additional_info,c.additional_info_formatting,true)});
   toast('Source type saved');if(existing)go('pick');else go('form',{type:t.id,values:{},info:{},pin:'',showMore:false,keepScroll:false})};
 up();
 return[titleBack,h('div',{class:'tp-body'},field('Name',h('input',{class:'in',id:'tp-typename',placeholder:'e.g., EU Regulation',value:E.name,oninput:e=>E.name=e.target.value})),h('h3',{class:'tp-sub'},'Fields, in citation order'),E.fields.map(row),btn('Add Field',()=>{E.fields.push({id:uid(),name:'',prefix:'',infix:'',suffix:'',italic:false,bold:false,order:E.fields.length});draw()},'w','plus')),
  h('div',{class:'tp-foot'},h('span',{class:'small muted'},'Format'),prev,btn('Save Source Type',save,'pri w','save'))]}
/* new citation */
function nameInputs(f,upd){const raw=P.values[f.name]||'',parts=raw.split('|||');
 return h('div',{class:'stack t'},parts.map((p,i)=>{const w=p.split(' ');const cur=[w[0]||'',w[1]||'',w.slice(2).join(' ')];return h('div',{class:'row',style:'gap:4px'},h('div',{class:'tp-names grow'},['First','Middle','Last'].map((ph,k)=>h('input',{class:'in',placeholder:ph,'aria-label':`${f.name} ${ph}`,value:cur[k],oninput:e=>{cur[k]=e.target.value;parts[i]=cur.join(' ').trim();P.values[f.name]=parts.join('|||');upd()}}))),parts.length>1&&ibtn('x',()=>{parts.splice(i,1);P.values[f.name]=parts.join('|||');draw()},'Remove person','red'))}),h('div',null,btn('Add another person',()=>{P.values[f.name]=raw+'|||';draw()},'xs ghost','plus')))}
function formView(){const t=S.type.find(x=>x.id===P.type);if(!t)return pickView();
 const prev=h('div',{class:'cite tp-prev'});const html=()=>format(t,P.values,{},P.info,{},true);const upd=()=>{const x=html();if(x)prev.innerHTML=x;else prev.textContent='Your citation appears here.'};upd();
 const dest=()=>P.project==='all'||P.project==='uncategorized'?null:P.project;
 const saveAnd=async ins=>{const x=format(t,P.values,{},P.info,{},false);if(!x){toast('Please fill in at least one field',1);return}
   const same=S.citation.find(z=>z.source_type_id===t.id&&norm(z.formatted_citation)===norm(x));if(same){if(ins)await placeCitation(same,P.pin);else toast('This citation is already saved');go('list',{values:{},info:{},pin:'',keepScroll:false});return}
   const c=await put('citation',{source_type_id:t.id,citation_style_id:t.citation_style_id,library_id:dest(),field_values:P.values,custom_formatting:{},additional_info:P.info,additional_info_formatting:{},formatted_citation:x,formatted_citation_html:html(),notes:''});
   if(ins)await placeCitation(c,P.pin);else toast('Citation saved');go('list',{values:{},info:{},pin:'',keepScroll:false})};
 const doiImport=async()=>{const d=document.getElementById('tp-doi').value;if(!d.trim())return;try{const r=await crossrefLookup(d,t);Object.assign(P.values,r.values);Object.assign(P.info,r.info);draw();toast('Fields filled. Check them against your source.')}catch(e){toast('No record was found for that DOI. Check it, or fill in the fields by hand.',1)}};
 const isBook=(t.name||'').toLowerCase().includes('book');
 const extras=isBook&&[['publisher','Publisher'],['year','Year'],['edition','Edition'],['editor','Editor'],['translator','Translator'],['other','Other information']].filter(([k])=>!(k==='editor'&&t.name.toLowerCase().includes('edited')));
 const main=extras?extras.slice(0,2):[],more=extras?extras.slice(2):[];
 const infoIn=([k,l])=>field(l,h('input',{class:'in','aria-label':l,value:P.info[k]||'',oninput:e=>{P.info[k]=e.target.value;upd()}}));
 return[backBar(t.name,()=>go('pick')),h('div',{class:'tp-body'},
   h('div',{class:'row'},h('input',{class:'in grow',id:'tp-doi',placeholder:'Paste a DOI to fill fields (optional)','aria-label':'DOI'}),btn('Fill',doiImport,'sm soft')),
   (t.fields||[]).filter(f=>!isExtra(f)).slice().sort(byOrder).map(f=>field(f.name,/author|editor/i.test(f.name)?nameInputs(f,upd):h('input',{class:'in','aria-label':f.name,value:P.values[f.name]||'',oninput:e=>{P.values[f.name]=e.target.value;upd()}}))),
   main.map(infoIn),more.length>0&&(P.showMore?more.map(infoIn):btn('More book details',()=>{P.showMore=true;draw()},'sm ghost','down')),
   h('p',{class:'small muted'},dest()?'Saves to '+libName(dest()):'Saves with no project')),
  h('div',{class:'tp-foot'},prev,h('input',{class:'in',id:'tp-newpin',placeholder:'Pinpoint (optional), e.g. para 12','aria-label':'Pinpoint',value:P.pin,oninput:e=>P.pin=e.target.value}),h('div',{class:'row'},btn('Save Only',()=>saveAnd(false)),h('span',{class:'grow'},btn('Save and Insert',()=>saveAnd(true),'pri w','file'))))]}
async function refresh(){try{await loadAll();await scanDoc();draw();toast('Citations refreshed')}catch(e){toast('Could not refresh. Check your connection.',1)}}
async function start(){root().replaceChildren(h('p',{class:'muted',style:'padding:16px'},'Loading your citations…'));try{await loadAll()}catch(e){storeNote='Your citations could not be loaded. Check your connection and press Refresh.'}await scanDoc();draw()}
window.addEventListener('focus',async()=>{if(!currentUser||P.screen==='type'||P.screen==='form')return;const a=document.activeElement;if(a&&a.matches&&a.matches('input,select,textarea'))return;try{await loadAll();await scanDoc();draw()}catch(e){}});
function detectWord(){return new Promise(res=>{if(!window.Office||!Office.onReady){res(false);return}let done=false;const fin=v=>{if(!done){done=true;res(v)}};Office.onReady(i=>fin(!!(i&&i.host===Office.HostType.Word&&Office.context.requirements.isSetSupported('WordApi','1.5'))));setTimeout(()=>fin(false),4000)})}
(async()=>{P.inWord=await detectWord();const u=await getUser();if(u)start();else root().replaceChildren(authView(start,true))})();
