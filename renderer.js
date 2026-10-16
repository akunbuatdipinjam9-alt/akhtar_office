const $ = (s)=>document.querySelector(s);
const $$ = (s)=>[...document.querySelectorAll(s)];
const state = {
  mode:'home', filePath:null, title:'Untitled',
  doc:{html:'<h1>Untitled Document</h1><p>Start writing here.</p>',pageSize:'a4'},
  code:{content:'',language:'javascript'}, slides:[{id:crypto.randomUUID(),bg:'#ffffff',transition:'none',notes:'',elements:[
    {id:crypto.randomUUID(),type:'text',x:90,y:105,w:720,h:90,text:'Untitled Presentation',fontSize:42,color:'#10131a',fill:'transparent',fontWeight:'700'},
    {id:crypto.randomUUID(),type:'text',x:92,y:210,w:600,h:60,text:'Click any text to edit it.',fontSize:22,color:'#596174',fill:'transparent',fontWeight:'400'}
  ]}], currentSlide:0, selectedElement:null
};
let dirty=false, drag=null, presentIndex=0;
let slideHistory=[], slideFuture=[], historyLock=false, objectClipboard=null;
const slideView={showGrid:false,snap:true,gridSize:10,zoom:null};

function toast(msg){const t=$('#toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),1800)}
function markDirty(){dirty=true;scheduleRecovery();$('#titleStatus').textContent=`${state.mode==='docs'?'Docs':state.mode==='slides'?'Slides':'Home'} · Unsaved`;}
function setClean(){dirty=false;$('#titleStatus').textContent=state.mode==='home'?'Home':state.mode==='docs'?'Docs':'Slides'}
function setMode(mode){
  state.mode=mode; $('#home').style.display=mode==='home'?'block':'none';
  $('#codeScreen').classList.toggle('active',mode==='code'); $('#codeBtn').classList.toggle('active',mode==='code');
  $('#docsScreen').classList.toggle('active',mode==='docs'); $('#slidesScreen').classList.toggle('active',mode==='slides');
  $('#homeBtn').classList.toggle('active',mode==='home'); $('#docsBtn').classList.toggle('active',mode==='docs'); $('#slidesBtn').classList.toggle('active',mode==='slides');
  $('#projectTitle').value=state.title; setClean();
  if(mode==='code')renderCode(); if(mode==='docs') renderDoc(); if(mode==='slides') renderSlides();
}
function newDoc(){state.filePath=null;state.title='Untitled Document';state.doc={html:'<h1>Untitled Document</h1><p>Start writing here.</p>',pageSize:'a4'};setMode('docs');markDirty()}
function newSlides(){state.filePath=null;state.title='Untitled Presentation';state.slides=[{id:crypto.randomUUID(),bg:'#ffffff',transition:'none',notes:'',elements:[]}];state.currentSlide=0;state.selectedElement=null;slideHistory=[];slideFuture=[];setMode('slides');markDirty()}

$('#homeBtn').onclick=()=>setMode('home'); $('#docsBtn').onclick=()=>setMode('docs'); $('#slidesBtn').onclick=()=>setMode('slides');
$('#newDocCard').onclick=newDoc; $('#newSlideCard').onclick=newSlides;
$('#newBtn').onclick=()=>state.mode==='slides'?newSlides():newDoc();
$('#projectTitle').oninput=e=>{state.title=e.target.value;markDirty()};

function serialize(){
  if(state.mode==='code')state.code.content=$('#codeText').value;
  if(state.mode==='docs') state.doc.html=$('#docEditor').innerHTML;
  return {type:state.mode,title:state.title,doc:state.doc,slides:state.slides,currentSlide:state.currentSlide,version:2};
}
async function save(){
  const path=await window.akhtar.saveProject({filePath:state.filePath,data:serialize(),suggestedName:(state.title||'Untitled').replace(/[\\/:*?"<>|]/g,'_')+'.akoffice'});
  if(path){state.filePath=path;setClean();toast('Project saved')}
}
async function open(){
  const res=await window.akhtar.openProject(); if(!res)return;
  const d=res.data; state.filePath=res.filePath;state.title=d.title||'Untitled';state.doc=d.doc||state.doc;state.slides=(d.slides||state.slides).map(sl=>({transition:'none',notes:'',...sl,elements:sl.elements||[]}));state.currentSlide=d.currentSlide||0;state.selectedElement=null;setMode(d.type==='slides'?'slides':'docs');$('#projectTitle').value=state.title;slideHistory=[];slideFuture=[];toast('Project opened')
}
$('#saveBtn').onclick=save; $('#openBtn').onclick=open; $('#homeOpen').onclick=open;
document.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='s'){e.preventDefault();save()}if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='o'){e.preventDefault();open()}})
$('#exportBtn').onclick=async()=>{if(state.mode==='home'){toast('Open a document or presentation first');return}const old=document.title;document.title=state.title;const p=await window.akhtar.printPdf((state.title||'export')+'.pdf');document.title=old;if(p)toast('PDF exported')}
$('#aboutBtn').onclick=async()=>toast('Akhtar Office v'+await window.akhtar.getVersion())

function renderDoc(){
  $('#docEditor').innerHTML=state.doc.html||''; $('#pageSize').value=state.doc.pageSize||'a4'; applyPageSize(); updateDocStats();
}
function applyPageSize(){const p=$('#docEditor');if($('#pageSize').value==='letter'){p.style.width='816px';p.style.minHeight='1056px'}else{p.style.width='850px';p.style.minHeight='1120px'}}
$('#pageSize').onchange=e=>{state.doc.pageSize=e.target.value;applyPageSize();markDirty()}
$('#docEditor').addEventListener('input',()=>{state.doc.html=$('#docEditor').innerHTML;updateDocStats();markDirty()});
function updateDocStats(){const txt=$('#docEditor').innerText.trim();const words=txt?txt.split(/\s+/).length:0;$('#docStats').textContent=`${words} words · ${txt.length} chars`}
$$('[data-cmd]').forEach(b=>b.onclick=()=>{document.execCommand(b.dataset.cmd,false,null);$('#docEditor').focus();markDirty()});
$('#blockType').onchange=e=>{document.execCommand('formatBlock',false,e.target.value);$('#docEditor').focus();markDirty()};
$('#fontFamily').onchange=e=>{document.execCommand('fontName',false,e.target.value);$('#docEditor').focus();markDirty()};
$('#fontSize').onchange=e=>{document.execCommand('fontSize',false,7);$$('#docEditor font[size="7"]').forEach(el=>{el.removeAttribute('size');el.style.fontSize=e.target.value+'px'});markDirty()};
$('#fontColor').oninput=e=>{document.execCommand('foreColor',false,e.target.value);markDirty()};
$('#insertDivider').onclick=()=>{document.execCommand('insertHTML',false,'<hr><p><br></p>');markDirty()};
$('#insertTable').onclick=()=>{let h='<table><tbody>';for(let r=0;r<3;r++){h+='<tr>';for(let c=0;c<3;c++)h+='<td>Cell</td>';h+='</tr>'}h+='</tbody></table><p><br></p>';document.execCommand('insertHTML',false,h);markDirty()};

function slideSnapshot(){return JSON.stringify({slides:state.slides,currentSlide:state.currentSlide,selectedElement:state.selectedElement})}
function restoreSlideSnapshot(raw){historyLock=true;const x=JSON.parse(raw);state.slides=x.slides;state.currentSlide=x.currentSlide;state.selectedElement=x.selectedElement||null;historyLock=false;renderSlides();markDirty()}
function pushSlideHistory(){if(historyLock)return;const snap=slideSnapshot();if(slideHistory.at(-1)!==snap)slideHistory.push(snap);if(slideHistory.length>60)slideHistory.shift();slideFuture=[]}
function undoSlides(){if(!slideHistory.length)return toast('Nothing to undo');slideFuture.push(slideSnapshot());restoreSlideSnapshot(slideHistory.pop())}
function redoSlides(){if(!slideFuture.length)return toast('Nothing to redo');slideHistory.push(slideSnapshot());restoreSlideSnapshot(slideFuture.pop())}
function snapValue(v){return slideView.snap?Math.round(v/slideView.gridSize)*slideView.gridSize:v}
function current(){return state.slides[state.currentSlide]}
function renderSlides(){
  if(!state.slides.length) state.slides=[{id:crypto.randomUUID(),bg:'#fff',transition:'none',notes:'',elements:[]}];
  state.currentSlide=Math.max(0,Math.min(state.currentSlide,state.slides.length-1));
  const s=current(); s.transition=s.transition||'none';s.notes=s.notes||'';
  const wrap=$('.slide-stage-wrap');const fit=Math.min(1,Math.max(.35,(wrap.clientWidth-72)/960));const zoom=slideView.zoom??fit;
  const stage=$('#slideStage');stage.style.zoom=zoom;stage.style.backgroundColor=s.bg;stage.classList.toggle('grid',slideView.showGrid);stage.style.backgroundSize=slideView.gridSize+'px '+slideView.gridSize+'px';
  $('#slideBg').value=normalizeColor(s.bg,'#ffffff');$('#slideTransition').value=s.transition;$('#speakerNotes').value=s.notes; stage.innerHTML='';
  s.elements.forEach(el=>stage.appendChild(makeElement(el,false)));
  $('#slideZoom').value=Math.round(zoom*100);$('#zoomLabel').textContent=Math.round(zoom*100)+'%';
  $('#toggleGrid').classList.toggle('active',slideView.showGrid);$('#toggleSnap').classList.toggle('active',slideView.snap);$('#gridSize').value=String(slideView.gridSize);
  renderSlideList();renderLayerList(); updateSlideStatus(); updateProps();
}
function normalizeColor(c,fb){return /^#[0-9a-f]{6}$/i.test(c||'')?c:fb}
function makeElement(el,thumb,interactive=true){
  const node=document.createElement('div');node.className='slide-element '+(el.type==='shape'?'slide-shape':'slide-text');node.dataset.id=el.id;
  const scale=thumb?.12:1; Object.assign(node.style,{left:el.x*scale+'px',top:el.y*scale+'px',width:el.w*scale+'px',height:el.h*scale+'px',fontSize:(el.fontSize||20)*scale+'px',color:el.color||'#111',background:el.fill==='transparent'?'transparent':el.fill||'transparent',fontWeight:el.fontWeight||'400',borderRadius:el.shape==='circle'?'50%':(el.type==='shape'?'14px':'0')});
  if(el.type==='image'){const img=document.createElement('img');img.src=el.src||'';img.alt=el.name||'Image';Object.assign(img.style,{width:'100%',height:'100%',objectFit:'contain',pointerEvents:'none'});node.appendChild(img)}
  if(el.type==='text'){node.textContent=el.text; if(!thumb && interactive){node.contentEditable=true;node.spellcheck=true;let textHistoryTaken=false;node.addEventListener('focus',()=>{if(!textHistoryTaken){pushSlideHistory();textHistoryTaken=true}});node.addEventListener('blur',()=>{textHistoryTaken=false});node.addEventListener('input',()=>{el.text=node.innerText;markDirty();renderSlideList();renderLayerList()})}}
  if(!thumb && interactive){if(el.id===state.selectedElement)node.classList.add('selected');node.addEventListener('mousedown',ev=>startDrag(ev,el,node));node.addEventListener('click',ev=>{ev.stopPropagation();selectElement(el.id)});const h=document.createElement('div');h.className='resize-handle';h.addEventListener('mousedown',ev=>startResize(ev,el,node));node.appendChild(h)}
  return node
}
function renderSlideList(){const list=$('#slideList');list.innerHTML='';state.slides.forEach((s,i)=>{const wrap=document.createElement('div');wrap.className='slide-thumb'+(i===state.currentSlide?' active':'');const v=document.createElement('div');v.className='slide-thumb-view';v.style.background=s.bg;s.elements.forEach(el=>v.appendChild(makeElement(el,true)));const lab=document.createElement('label');lab.textContent='Slide '+(i+1)+(s.transition&&s.transition!=='none'?' · '+s.transition:'');wrap.append(v,lab);wrap.onclick=()=>{state.currentSlide=i;state.selectedElement=null;renderSlides()};list.appendChild(wrap)})}
function updateSlideStatus(){$('#slideStatus').textContent=`Slide ${state.currentSlide+1} of ${state.slides.length}`}
$('#slideStage').onclick=()=>{state.selectedElement=null;renderSlides()}
function selectElement(id){state.selectedElement=id;renderSlides()}
function getSelected(){return current().elements.find(e=>e.id===state.selectedElement)}
function addElement(el){pushSlideHistory();current().elements.push({id:crypto.randomUUID(),x:120,y:120,w:360,h:70,fontSize:28,color:'#111111',fill:'transparent',fontWeight:'400',...el});state.selectedElement=current().elements.at(-1).id;renderSlides();markDirty()}
$('#addText').onclick=()=>addElement({type:'text',text:'Text'});$('#addTitle').onclick=()=>addElement({type:'text',text:'Title',fontSize:44,fontWeight:'700',w:620,h:90});$('#addRect').onclick=()=>addElement({type:'shape',shape:'rect',w:220,h:120,fill:'#7c8cff'});$('#addCircle').onclick=()=>addElement({type:'shape',shape:'circle',w:140,h:140,fill:'#9d72ff'});
$('#addSlide').onclick=()=>{pushSlideHistory();state.slides.push({id:crypto.randomUUID(),bg:'#ffffff',transition:'none',notes:'',elements:[]});state.currentSlide=state.slides.length-1;state.selectedElement=null;renderSlides();markDirty()};
$('#duplicateSlide').onclick=()=>{pushSlideHistory();const clone=JSON.parse(JSON.stringify(current()));clone.id=crypto.randomUUID();clone.elements=clone.elements.map(e=>({...e,id:crypto.randomUUID()}));state.slides.splice(state.currentSlide+1,0,clone);state.currentSlide++;state.selectedElement=null;renderSlides();markDirty()};
$('#deleteSlide').onclick=()=>{if(state.slides.length===1){toast('At least one slide is required');return}pushSlideHistory();state.slides.splice(state.currentSlide,1);state.currentSlide=Math.max(0,state.currentSlide-1);state.selectedElement=null;renderSlides();markDirty()};
$('#slideBg').onpointerdown=()=>pushSlideHistory();$('#slideBg').oninput=e=>{current().bg=e.target.value;$('#slideStage').style.background=e.target.value;renderSlideList();markDirty()};
function deleteSelected(){const id=state.selectedElement;if(!id)return;pushSlideHistory();current().elements=current().elements.filter(e=>e.id!==id);state.selectedElement=null;renderSlides();markDirty()}
$('#deleteElement').onclick=deleteSelected;

document.addEventListener('keydown',e=>{if(state.mode==='slides'&&['Delete','Backspace'].includes(e.key)&&state.selectedElement&&document.activeElement?.contentEditable!=='true'){e.preventDefault();deleteSelected()}})
function startDrag(ev,el,node){if(ev.target.classList.contains('resize-handle'))return;if(ev.target.contentEditable==='true'&&ev.detail===2)return;ev.preventDefault();ev.stopPropagation();state.selectedElement=el.id;pushSlideHistory();const stage=$('#slideStage').getBoundingClientRect();drag={type:'move',el,node,startX:ev.clientX,startY:ev.clientY,ox:el.x,oy:el.y,stage};document.addEventListener('mousemove',onDrag);document.addEventListener('mouseup',stopDrag,{once:true});renderSelectionOnly()}
function startResize(ev,el,node){ev.preventDefault();ev.stopPropagation();pushSlideHistory();drag={type:'resize',el,node,startX:ev.clientX,startY:ev.clientY,ow:el.w,oh:el.h};document.addEventListener('mousemove',onDrag);document.addEventListener('mouseup',stopDrag,{once:true})}
function onDrag(ev){if(!drag)return;if(drag.type==='move'){drag.el.x=Math.max(0,snapValue(drag.ox+(ev.clientX-drag.startX)/(Number($('#slideStage').style.zoom)||1)));drag.el.y=Math.max(0,snapValue(drag.oy+(ev.clientY-drag.startY)/(Number($('#slideStage').style.zoom)||1)));drag.node.style.left=drag.el.x+'px';drag.node.style.top=drag.el.y+'px'}else{drag.el.w=Math.max(30,snapValue(drag.ow+(ev.clientX-drag.startX)/(Number($('#slideStage').style.zoom)||1)));drag.el.h=Math.max(20,snapValue(drag.oh+(ev.clientY-drag.startY)/(Number($('#slideStage').style.zoom)||1)));drag.node.style.width=drag.el.w+'px';drag.node.style.height=drag.el.h+'px'}updatePropsValues();markDirty()}
function stopDrag(){document.removeEventListener('mousemove',onDrag);drag=null;renderSlideList()}
function renderSelectionOnly(){$$('.slide-element').forEach(n=>n.classList.toggle('selected',n.dataset.id===state.selectedElement));updateProps()}
function updateProps(){const e=getSelected();$('#elementProps').style.display=e?'block':'none';if(!e)return;updatePropsValues()}
function updatePropsValues(){const e=getSelected();if(!e)return;$('#propX').value=Math.round(e.x);$('#propY').value=Math.round(e.y);$('#propW').value=Math.round(e.w);$('#propH').value=Math.round(e.h);$('#propFont').value=e.fontSize||20;$('#propColor').value=normalizeColor(e.color,'#111111');$('#propFill').value=normalizeColor(e.fill,'#ffffff')}
['propX','propY','propW','propH','propFont'].forEach(id=>$('#'+id).onchange=ev=>{const e=getSelected();if(!e)return;pushSlideHistory();const map={propX:'x',propY:'y',propW:'w',propH:'h',propFont:'fontSize'};e[map[id]]=Number(ev.target.value)||0;renderSlides();markDirty()});
$('#propColor').onpointerdown=()=>{if(getSelected())pushSlideHistory()};$('#propColor').oninput=ev=>{const e=getSelected();if(e){e.color=ev.target.value;renderSlides();markDirty()}};$('#propFill').onpointerdown=()=>{if(getSelected())pushSlideHistory()};$('#propFill').oninput=ev=>{const e=getSelected();if(e){e.fill=ev.target.value;renderSlides();markDirty()}};

function renderLayerList(){const host=$('#layerList');host.innerHTML='';const arr=[...current().elements].reverse();if(!arr.length){host.innerHTML='<div style="font-size:12px;color:var(--muted)">No objects</div>';return}arr.forEach(el=>{const row=document.createElement('div');row.className='layer-row'+(el.id===state.selectedElement?' active':'');const name=document.createElement('div');name.className='layer-name';name.textContent=el.type==='text'?(el.text||'Text').slice(0,26):el.type==='image'?(el.name||'Image'):(el.shape||'Shape');const down=document.createElement('button');down.textContent='↓';down.title='Send backward';down.onclick=ev=>{ev.stopPropagation();moveLayerStep(el.id,-1)};const up=document.createElement('button');up.textContent='↑';up.title='Bring forward';up.onclick=ev=>{ev.stopPropagation();moveLayerStep(el.id,1)};row.append(name,down,up);row.onclick=()=>selectElement(el.id);host.appendChild(row)})}
function moveLayerStep(id,dir){const arr=current().elements,idx=arr.findIndex(x=>x.id===id),next=idx+dir;if(idx<0||next<0||next>=arr.length)return;pushSlideHistory();[arr[idx],arr[next]]=[arr[next],arr[idx]];state.selectedElement=id;renderSlides();markDirty()}
$('#slideUndo').onclick=undoSlides;$('#slideRedo').onclick=redoSlides;
$('#toggleGrid').onclick=()=>{slideView.showGrid=!slideView.showGrid;renderSlides()};
$('#toggleSnap').onclick=()=>{slideView.snap=!slideView.snap;renderSlides();toast(slideView.snap?'Snap on':'Snap off')};
$('#gridSize').onchange=e=>{slideView.gridSize=Number(e.target.value)||10;renderSlides()};
$('#fitSlide').onclick=()=>{slideView.zoom=null;renderSlides()};
$('#slideZoom').oninput=e=>{slideView.zoom=Number(e.target.value)/100;renderSlides()};
$('#centerX').onclick=()=>{const e=getSelected();if(!e)return toast('Select an object');pushSlideHistory();e.x=(960-e.w)/2;renderSlides();markDirty()};
$('#centerY').onclick=()=>{const e=getSelected();if(!e)return toast('Select an object');pushSlideHistory();e.y=(540-e.h)/2;renderSlides();markDirty()};
$('#slideTransition').onchange=e=>{pushSlideHistory();current().transition=e.target.value;markDirty()};
let notesHistoryTaken=false;$('#speakerNotes').onfocus=()=>{if(!notesHistoryTaken){pushSlideHistory();notesHistoryTaken=true}};$('#speakerNotes').onblur=()=>{notesHistoryTaken=false};$('#speakerNotes').oninput=e=>{current().notes=e.target.value;markDirty()};
window.addEventListener('resize',()=>{if(state.mode==='slides'&&slideView.zoom===null)renderSlides()});
document.addEventListener('keydown',e=>{if(state.mode!=='slides'||e.target.closest('input,textarea,[contenteditable="true"]'))return;const ctrl=e.ctrlKey||e.metaKey;if(ctrl&&e.key.toLowerCase()==='z'){e.preventDefault();e.shiftKey?redoSlides():undoSlides();return}if(ctrl&&e.key.toLowerCase()==='y'){e.preventDefault();redoSlides();return}const el=getSelected();if(!el)return;const step=e.shiftKey?10:1;if(e.key==='ArrowLeft'){pushSlideHistory();el.x=Math.max(0,el.x-step)}else if(e.key==='ArrowRight'){pushSlideHistory();el.x=Math.min(960-el.w,el.x+step)}else if(e.key==='ArrowUp'){pushSlideHistory();el.y=Math.max(0,el.y-step)}else if(e.key==='ArrowDown'){pushSlideHistory();el.y=Math.min(540-el.h,el.y+step)}else return;e.preventDefault();renderSlides();markDirty()});

$('#presentBtn').onclick=()=>{presentIndex=state.currentSlide;$('#presentOverlay').classList.add('show');renderPresent()}
function renderPresent(){const s=state.slides[presentIndex], host=$('#presentSlide');host.className='present-slide';void host.offsetWidth;host.classList.add(s.transition==='fade'?'t-fade':s.transition==='left'?'t-left':s.transition==='zoom'?'t-zoom':'');host.innerHTML='';host.style.background=s.bg;const sx=innerWidth/960, sy=innerHeight/540;s.elements.forEach(el=>{const n=makeElement({...el,x:el.x*sx,y:el.y*sy,w:el.w*sx,h:el.h*sy,fontSize:(el.fontSize||20)*Math.min(sx,sy)},false,false);n.style.left=el.x*sx+'px';n.style.top=el.y*sy+'px';n.style.width=el.w*sx+'px';n.style.height=el.h*sy+'px';n.style.fontSize=(el.fontSize||20)*Math.min(sx,sy)+'px';host.appendChild(n)})}
document.addEventListener('keydown',e=>{if(!$('#presentOverlay').classList.contains('show'))return;if(e.key==='Escape')$('#presentOverlay').classList.remove('show');if(e.key==='ArrowRight'){presentIndex=Math.min(state.slides.length-1,presentIndex+1);renderPresent()}if(e.key==='ArrowLeft'){presentIndex=Math.max(0,presentIndex-1);renderPresent()}})
$('#presentOverlay').onclick=e=>{if(e.target===$('#presentOverlay'))$('#presentOverlay').classList.remove('show')}
window.addEventListener('resize',()=>{if($('#presentOverlay').classList.contains('show'))renderPresent()})
window.addEventListener('beforeunload',e=>{if(dirty){e.preventDefault();e.returnValue=''}})

$('#addImage').onclick=async()=>{const picked=await window.akhtar.pickImage();if(!picked)return;addElement({type:'image',src:picked.dataUrl,name:picked.name,x:160,y:110,w:480,h:300,fill:'transparent'});toast('Image inserted')};
$('#copyObject').onclick=()=>{const e=getSelected();if(!e)return toast('Select an object');objectClipboard=JSON.parse(JSON.stringify(e));toast('Copied')};
$('#pasteObject').onclick=()=>{if(!objectClipboard)return toast('Copy an object first');const e=JSON.parse(JSON.stringify(objectClipboard));addElement({...e,id:crypto.randomUUID(),x:e.x+24,y:e.y+24})};
function changeLayer(front){const arr=current().elements,i=arr.findIndex(e=>e.id===state.selectedElement);if(i<0)return toast('Select an object');pushSlideHistory();const [el]=arr.splice(i,1);front?arr.push(el):arr.unshift(el);renderSlides();markDirty()}
$('#bringFront').onclick=()=>changeLayer(true);$('#sendBack').onclick=()=>changeLayer(false);
$('#exportPptx').onclick=async()=>{try{const p=await window.akhtar.exportPptx({title:state.title,slides:state.slides});if(p)toast('PPTX exported')}catch(e){toast('Export failed: '+e.message)}};
document.addEventListener('keydown',e=>{if(state.mode!=='slides'||e.target.closest('input,textarea,[contenteditable="true"]'))return;if(!(e.ctrlKey||e.metaKey))return;if(e.key.toLowerCase()==='c'&&getSelected()){e.preventDefault();$('#copyObject').click()}if(e.key.toLowerCase()==='v'&&objectClipboard){e.preventDefault();$('#pasteObject').click()}if(e.key.toLowerCase()==='d'&&getSelected()){e.preventDefault();$('#copyObject').click();$('#pasteObject').click()}});


let recoveryTimer=null;
function scheduleRecovery(){
  clearTimeout(recoveryTimer);
  recoveryTimer=setTimeout(async()=>{
    if(!dirty||state.mode==='home')return;
    try{await window.akhtar.writeRecovery(serialize());$('#titleStatus').textContent+=' · Recovered locally'}catch(e){console.warn('Autosave failed',e)}
  },1800);
}
async function recoverAutosave(){
  try{
    const d=await window.akhtar.readRecovery();if(!d)return toast('No recovery found');
    if(d.type!=='docs'&&d.type!=='slides')return toast('Invalid recovery');
    state.filePath=null;state.title=d.title||'Recovered project';state.doc=d.doc||state.doc;
    state.slides=Array.isArray(d.slides)&&d.slides.length?d.slides:state.slides;
    state.currentSlide=Math.min(Math.max(0,d.currentSlide||0),state.slides.length-1);
    state.selectedElement=null;setMode(d.type);markDirty();toast('Autosave recovered. Save to keep it.');
  }catch(e){toast('Recovery failed: '+e.message)}
}
$('#recoverBtn').onclick=recoverAutosave;
$('#docMargin').onchange=e=>{state.doc.margin=e.target.value;applyDocLayout();markDirty()};
function applyDocLayout(){
  const el=$('#docEditor');el.classList.remove('margin-narrow','margin-wide','a4','letter');
  el.classList.add(state.doc.margin==='narrow'?'margin-narrow':state.doc.margin==='wide'?'margin-wide':'a4');
  if(state.doc.pageSize==='letter'){el.classList.remove('a4');el.classList.add('letter')}
  $('#docMargin').value=state.doc.margin||'normal';
}
const originalRenderDoc=renderDoc;
renderDoc=function(){originalRenderDoc();applyDocLayout()};
$('#pageBreak').onclick=()=>{document.execCommand('insertHTML',false,'<div class="doc-pagebreak" contenteditable="false">Page break</div><p></p>');markDirty()};
$('#docImage').onclick=async()=>{
 const file=await window.akhtar.pickImage();if(!file)return;
 $('#docEditor').focus();document.execCommand('insertHTML',false,`<img src="${file.dataUrl}" alt="Inserted image" style="max-width:100%;height:auto">`);markDirty()
};
$('#exportDocx').onclick=async()=>{
 try{const file=await window.akhtar.exportDocx({title:state.title,html:$('#docEditor').innerHTML,pageSize:state.doc.pageSize,margin:state.doc.margin});if(file)toast('DOCX exported')}catch(e){toast('DOCX export error: '+e.message)}
};
$('#exportHtml').onclick=async()=>{
 const html='<!doctype html><html><head><meta charset="utf-8"><title>'+escapeHTML(state.title)+'</title><style>body{font:16px/1.65 Arial;max-width:850px;margin:50px auto;padding:20px}img{max-width:100%}table{border-collapse:collapse}td,th{border:1px solid #aaa;padding:8px}@media print{.doc-pagebreak{break-after:page}}</style></head><body>'+$('#docEditor').innerHTML+'</body></html>';
 const file=await window.akhtar.saveText({content:html,suggestedName:state.title+'.html',filters:[{name:'HTML',extensions:['html']}]});if(file)toast('HTML exported');
};
function escapeHTML(str){return String(str).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
let findPosition=0;
$('#findNext').onclick=()=>{
 const query=$('#findText').value;if(!query)return;
 const root=$('#docEditor'),walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);
 const nodes=[];let node;while(node=walker.nextNode())nodes.push(node);
 const all=nodes.map(n=>n.textContent).join(''),at=all.toLowerCase().indexOf(query.toLowerCase(),findPosition);
 const idx=at<0?all.toLowerCase().indexOf(query.toLowerCase()):at;
 if(idx<0)return toast('No match');findPosition=idx+query.length;
 let pos=0,startNode=null,endNode=null,start=0,end=0;
 for(const n of nodes){const next=pos+n.textContent.length;if(!startNode&&idx<next){startNode=n;start=idx-pos}if(startNode&&idx+query.length<=next){endNode=n;end=idx+query.length-pos;break}pos=next}
 if(startNode&&endNode){const range=document.createRange();range.setStart(startNode,start);range.setEnd(endNode,end);const sel=window.getSelection();sel.removeAllRanges();sel.addRange(range);startNode.parentElement?.scrollIntoView({block:'nearest'});toast('Match found')}
};
$('#replaceAll').onclick=()=>{
 const query=$('#findText').value;if(!query)return;
 const replacement=$('#replaceText').value,root=$('#docEditor'),walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);
 const nodes=[];let n;while(n=walker.nextNode())nodes.push(n);
 let count=0;for(const node of nodes){const source=node.textContent;const pattern=new RegExp(query.split('').map(c=>'\\^$.*+?()[]{}|'.includes(c)?'\\'+c:c).join(''),'gi');const hits=source.match(pattern);if(hits){count+=hits.length;node.textContent=source.replace(pattern,()=>replacement)}}
 if(count){markDirty();updateDocStats()}toast(count+' replacement(s)');findPosition=0
};
const themeColors={midnight:{bg:'#141b35',text:'#f2f5ff',accent:'#a8b4ff'},paper:{bg:'#faf9f6',text:'#222b39',accent:'#4659bb'},ocean:{bg:'#07536a',text:'#ffffff',accent:'#83e5ea'},sunset:{bg:'#78365b',text:'#ffffff',accent:'#ffbba0'}};
let activeTheme='midnight';
function makeSlideLayout(layout,theme){
 const c=themeColors[theme]||themeColors.midnight,uid=()=>crypto.randomUUID();
 const text=(value,x,y,w,h,size,bold=false,color=c.text)=>({id:uid(),type:'text',x,y,w,h,text:value,fontSize:size,color,fill:'transparent',fontWeight:bold?'700':'400'});
 const shape=(x,y,w,h,fill)=>({id:uid(),type:'shape',shape:'rect',x,y,w,h,fill,border:false});
 const elements=layout==='title'?[shape(72,85,12,355,c.accent),text('Your presentation title',112,150,745,105,48,true),text('Subtitle · Add your name',115,278,700,60,23)]:
 layout==='section'?[text('01',80,85,250,95,65,true,c.accent),text('Section heading',80,235,800,105,47,true)]:
 layout==='two'?[text('Two-column layout',70,45,820,65,37,true),shape(72,132,7,340,c.accent),text('First idea',95,145,350,55,26,true),text('Add supporting information here.',95,218,345,200,20),shape(495,132,7,340,c.accent),text('Second idea',522,145,350,55,26,true),text('Add supporting information here.',522,218,340,200,20)]:
 layout==='quote'?[text('“',70,40,180,170,110,true,c.accent),text('Your most important quote goes here.',115,160,740,185,40,true),text('— Source',120,380,650,55,21)]:[];
 return {id:uid(),bg:c.bg,transition:'fade',notes:'',elements};
}
$('#insertLayout').onclick=()=>{
 pushSlideHistory();const layout=$('#slideLayout').value;state.slides.splice(state.currentSlide+1,0,makeSlideLayout(layout,activeTheme));state.currentSlide++;state.selectedElement=null;renderSlides();markDirty();toast('Layout inserted')
};
$$('[data-theme]').forEach(b=>b.onclick=()=>{activeTheme=b.dataset.theme;const c=themeColors[activeTheme];pushSlideHistory();current().bg=c.bg;renderSlides();markDirty();toast(activeTheme+' theme selected')});
$('#applyThemeAll').onclick=()=>{
 const c=themeColors[activeTheme];pushSlideHistory();for(const slide of state.slides){slide.bg=c.bg;for(const el of slide.elements){if(el.type==='text'&&(el.color==='#10131a'||el.color==='#111111'||el.color==='#ffffff'||el.color==='#f2f5ff'))el.color=c.text}}
 renderSlides();markDirty();toast('Theme applied to all slides')
};


const workspaceTabs=[];let activeTab=null;let tabSerial=0;let paletteSelection=0;
function snapshotTab(){
 if(!activeTab||state.mode==='home')return;
 const t=workspaceTabs.find(t=>t.id===activeTab);if(!t)return;
 if(state.mode==='docs')state.doc.html=$('#docEditor').innerHTML;
 t.data={mode:state.mode,filePath:state.filePath,title:state.title,code:structuredClone(state.code),doc:structuredClone(state.doc),slides:structuredClone(state.slides),currentSlide:state.currentSlide};t.dirty=dirty;
}
function refreshWorkspaceTabs(){
 const host=$('#editorTabs');host.innerHTML='';const explorer=$('#explorerEditors');explorer.innerHTML='';
 for(const t of workspaceTabs){
  const tab=document.createElement('button');tab.className='editor-tab'+(t.id===activeTab?' active':'');tab.title=t.data.title;
  const symbol=document.createElement('span');symbol.textContent=t.data.mode==='code'?'<>':t.data.mode==='slides'?'▣':'▤';symbol.style.color=t.data.mode==='slides'?'#d6c29b':'#b5c7b2';
  const name=document.createElement('span');name.className='tabname';name.textContent=t.data.title;
  const close=document.createElement('span');close.className='tabclose';close.textContent=t.dirty?'●':'×';if(t.dirty)close.classList.add('tabdirty');
  close.onclick=e=>{e.stopPropagation();closeWorkspaceTab(t.id)};
  tab.append(symbol,name,close);tab.onclick=()=>activateWorkspaceTab(t.id);host.append(tab);
  const link=document.createElement('button');link.className='explorer-link'+(t.id===activeTab?' current':'');link.textContent=(t.data.mode==='code'?'< >  ':t.data.mode==='slides'?'▣  ':'▤  ')+t.data.title+(t.dirty?'  ●':'');link.onclick=()=>activateWorkspaceTab(t.id);explorer.append(link);
 }
}
function createWorkspaceTab(){
 if(state.mode==='home')return;
 const t={id:++tabSerial,data:null,dirty};workspaceTabs.push(t);activeTab=t.id;snapshotTab();refreshWorkspaceTabs();
}
function activateWorkspaceTab(id){
 if(id===activeTab)return;
 snapshotTab();const t=workspaceTabs.find(t=>t.id===id);if(!t)return;
 activeTab=id;const d=t.data;
 state.filePath=d.filePath;state.title=d.title;state.code=structuredClone(d.code||{content:'',language:'javascript'});state.doc=structuredClone(d.doc);state.slides=structuredClone(d.slides);state.currentSlide=d.currentSlide;state.selectedElement=null;
 slideHistory=[];slideFuture=[];setMode(d.mode);dirty=t.dirty;
 $('#titleStatus').textContent=d.mode+(dirty?' · Unsaved':'');refreshWorkspaceTabs();
}
function closeWorkspaceTab(id){
 const t=workspaceTabs.find(t=>t.id===id);if(!t)return;
 if(id===activeTab)snapshotTab();
 if(t.dirty&&!confirm('Close "'+t.data.title+'" without saving?'))return;
 const i=workspaceTabs.indexOf(t);workspaceTabs.splice(i,1);
 if(activeTab===id){activeTab=null;if(workspaceTabs.length)activateWorkspaceTab(workspaceTabs[Math.max(0,i-1)].id);else setMode('home')}
 refreshWorkspaceTabs();
}
const originalNewDoc=newDoc,originalNewSlides=newSlides,originalOpen=open,originalSave=save,originalMarkDirty=markDirty;
newDoc=function(){snapshotTab();originalNewDoc();createWorkspaceTab()};
newSlides=function(){snapshotTab();originalNewSlides();createWorkspaceTab()};
open=async function(){
 const before=state.filePath;
 snapshotTab();await originalOpen();
 if(state.filePath&&state.filePath!==before||state.filePath&&state.mode!=='home'&&(!activeTab||workspaceTabs.find(t=>t.id===activeTab)?.data.filePath!==state.filePath))createWorkspaceTab();
};
save=async function(){await originalSave();if(activeTab){snapshotTab();refreshWorkspaceTabs()}};
markDirty=function(){originalMarkDirty();if(activeTab){const t=workspaceTabs.find(t=>t.id===activeTab);if(t){t.dirty=true;t.data.title=state.title;refreshWorkspaceTabs()}}};
$('#newDocCard').onclick=newDoc;$('#newSlideCard').onclick=newSlides;
$('#docsBtn').onclick=newDoc;$('#slidesBtn').onclick=newSlides;
$('#newBtn').onclick=()=>state.mode==='slides'?newSlides():newDoc();
$('#saveBtn').onclick=()=>save();$('#openBtn').onclick=()=>open();$('#homeOpen').onclick=()=>open();
$('#codeBtn').onclick=newCode;$('#newCodeCard').onclick=newCode;$('#explorerNewCode').onclick=newCode;
$('#explorerNewDoc').onclick=newDoc;$('#explorerNewSlides').onclick=newSlides;$('#explorerOpen').onclick=()=>open();
$('#homeBtn').onclick=()=>{snapshotTab();setMode('home');refreshWorkspaceTabs()};
$('#quickNew').onclick=newDoc;
function toggleExplorer(){const el=$('#explorer');el.classList.toggle('hidden')}
$('#explorerToggle').onclick=toggleExplorer;$('#hideExplorer').onclick=toggleExplorer;
const commands=[
 ['New Code File','File',()=>newCode()],['Open File (Code / DOCX / PPTX)','File',()=>openAny()],
 ['New Document','File',()=>newDoc()],['New Presentation','File',()=>newSlides()],
 ['Open Project','File',()=>open()],['Save Project','File',()=>save()],
 ['Export PDF','Export',()=>$('#exportBtn').click()],['Export DOCX','Export',()=>state.mode==='docs'?$('#exportDocx').click():toast('Open Docs first')],
 ['Export PPTX','Export',()=>state.mode==='slides'?$('#exportPptx').click():toast('Open Slides first')],
 ['Toggle Explorer','View',toggleExplorer],['Show Start Page','View',()=>$('#homeBtn').click()],
 ['Insert Table','Document',()=>state.mode==='docs'?$('#insertTable').click():toast('Open Docs first')],
 ['New Slide','Presentation',()=>state.mode==='slides'?$('#addSlide').click():toast('Open Slides first')],
 ['Present Slides','Presentation',()=>state.mode==='slides'?$('#presentBtn').click():toast('Open Slides first')]
];
function renderCommands(){const q=$('#commandInput').value.trim().toLowerCase();const list=commands.filter(c=>c[0].toLowerCase().includes(q)||c[1].toLowerCase().includes(q));paletteSelection=Math.min(paletteSelection,Math.max(0,list.length-1));const host=$('#commandResults');host.innerHTML='';list.forEach((c,i)=>{const b=document.createElement('button');b.className='command-item'+(i===paletteSelection?' selected':'');const name=document.createElement('span');name.textContent=c[0];const group=document.createElement('small');group.textContent=c[1];b.append(name,group);b.onmouseenter=()=>{paletteSelection=i;host.querySelectorAll('.command-item').forEach((n,j)=>n.classList.toggle('selected',i===j))};b.onclick=()=>{closePalette();c[2]()};host.append(b)});if(!list.length)host.textContent='No matching commands'}
function openPalette(){paletteSelection=0;$('#commandInput').value='';$('#commandBackdrop').classList.add('show');renderCommands();$('#commandInput').focus()}
function closePalette(){$('#commandBackdrop').classList.remove('show')}
$('#paletteTop').onclick=openPalette;$('#commandButton').onclick=openPalette;
$('#commandBackdrop').onclick=e=>{if(e.target.id==='commandBackdrop')closePalette()};
$('#commandInput').oninput=()=>{paletteSelection=0;renderCommands()};
$('#commandInput').onkeydown=e=>{const visible=[...$('#commandResults').querySelectorAll('button')];if(e.key==='ArrowDown'){e.preventDefault();paletteSelection=Math.min(visible.length-1,paletteSelection+1);renderCommands()}if(e.key==='ArrowUp'){e.preventDefault();paletteSelection=Math.max(0,paletteSelection-1);renderCommands()}if(e.key==='Enter'){e.preventDefault();visible[paletteSelection]?.click()}if(e.key==='Escape')closePalette()};
document.addEventListener('keydown',e=>{
 const ctrl=e.ctrlKey||e.metaKey;
 if(ctrl&&e.shiftKey&&e.key.toLowerCase()==='p'){e.preventDefault();openPalette()}
 if(ctrl&&!e.shiftKey&&e.key.toLowerCase()==='b'){e.preventDefault();toggleExplorer()}
 if(ctrl&&!e.shiftKey&&e.key.toLowerCase()==='n'){e.preventDefault();newDoc()}
 if(ctrl&&!e.shiftKey&&e.key.toLowerCase()==='w'&&activeTab){e.preventDefault();closeWorkspaceTab(activeTab)}
 if(ctrl&&!e.shiftKey&&e.key.toLowerCase()==='s'){e.preventDefault();save()}
 if(ctrl&&!e.shiftKey&&e.key.toLowerCase()==='o'){e.preventDefault();open()}
 if(e.key==='Escape'&&$('#commandBackdrop').classList.contains('show'))closePalette();
},true);


function codeLanguage(ext){return ({js:'javascript',jsx:'javascript',ts:'typescript',tsx:'typescript',py:'python',md:'markdown',txt:'plaintext',htm:'html',jsonc:'json',yml:'plaintext',yaml:'plaintext'})[ext]||ext||'plaintext'}
function newCode(){snapshotTab();state.filePath=null;state.title='untitled.js';state.code={content:'// Write code here\n',language:'javascript'};setMode('code');markDirty();createWorkspaceTab()}
function renderCode(){const area=$('#codeText');area.value=state.code.content||'';$('#codeFileLabel').textContent=state.title;$('#codeLanguage').value=[...$('#codeLanguage').options].some(o=>o.value===state.code.language)?state.code.language:'plaintext';updateCodeStats()}
function updateCodeStats(){const area=$('#codeText'),text=area.value;const before=text.slice(0,area.selectionStart),line=before.split('\n').length,col=before.length-before.lastIndexOf('\n');$('#codeCursor').textContent='Ln '+line+', Col '+col;$('#codeStats').textContent=text.split('\n').length+' lines · '+text.length+' chars';$('#codeStatusLang').textContent=state.code.language||'Plain text';const n=text.split('\n').length;$('#codeGutter').textContent=Array.from({length:n},(_,i)=>i+1).join('\n')}
$('#codeText').addEventListener('input',()=>{state.code.content=$('#codeText').value;markDirty();updateCodeStats()});
$('#codeText').addEventListener('click',updateCodeStats);$('#codeText').addEventListener('keyup',updateCodeStats);
$('#codeText').addEventListener('scroll',()=>{$('#codeGutter').scrollTop=$('#codeText').scrollTop});
$('#codeText').addEventListener('keydown',e=>{if(e.key==='Tab'){e.preventDefault();const el=e.target,a=el.selectionStart,b=el.selectionEnd;el.setRangeText('  ',a,b,'end');el.dispatchEvent(new Event('input'))}});
$('#codeLanguage').onchange=e=>{state.code.language=e.target.value;markDirty();updateCodeStats()};
async function saveCode(saveAs=false){state.code.content=$('#codeText').value;const filePath=await window.akhtar.saveCode({filePath:state.filePath,content:state.code.content,suggestedName:state.title,saveAs});if(filePath){state.filePath=filePath;state.title=filePath.replace(/^.*[\\/]/,'');$('#projectTitle').value=state.title;setClean();snapshotTab();refreshWorkspaceTabs();renderCode();toast('Source file saved')}}
async function loadImported(data){if(!data)return;snapshotTab();if(data.kind==='project'){const d=data.data;state.filePath=data.filePath;state.title=d.title||'Untitled';state.doc=d.doc||state.doc;state.slides=d.slides||state.slides;state.currentSlide=d.currentSlide||0;setMode(d.type==='slides'?'slides':'docs');createWorkspaceTab()}else if(data.kind==='code'){state.filePath=data.filePath;state.title=data.title;state.code={content:data.content,language:codeLanguage(data.language)};setMode('code');createWorkspaceTab()}else if(data.kind==='docs'){state.filePath=null;state.title=data.title;state.doc=data.doc;setMode('docs');createWorkspaceTab()}else if(data.kind==='slides'){state.filePath=null;state.title=data.title;state.slides=data.slides;state.currentSlide=0;setMode('slides');createWorkspaceTab()}if(data.importWarnings?.length)toast('Imported with limitations — see README');}
async function openAny(){try{await loadImported(await window.akhtar.openAny())}catch(e){toast('Open failed: '+e.message)}}
$('#codeOpen').onclick=openAny;$('#codeSaveAs').onclick=()=>saveCode(true);
const v7save=save;save=async function(){if(state.mode==='code')return saveCode();return v7save()};
$('#saveBtn').onclick=()=>save();
$('#explorerOpen').onclick=openAny;$('#openBtn').onclick=openAny;$('#homeOpen').onclick=openAny;
commands.unshift(['Open Any File','File',openAny],['Save Code As...','Code',()=>state.mode==='code'?saveCode(true):toast('Open code first')]);
window.akhtar.onExternalFiles(async paths=>{for(const file of paths){try{await loadImported(await window.akhtar.readExternal(file))}catch(e){toast('Could not open file: '+e.message)}}});
document.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='s'&&state.mode==='code'){e.preventDefault();e.stopImmediatePropagation();saveCode()}if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='o'){e.preventDefault();e.stopImmediatePropagation();openAny()}},true);

const v11Area=$('#codeText');
function v11Sync(){state.code.content=v11Area.value;markDirty();updateCodeStats()}
function v11FindNext(){
 const term=$('#v11Find').value;if(!term)return toast('Enter text to find');
 const hay=v11Area.value.toLowerCase(),needle=term.toLowerCase();
 let pos=hay.indexOf(needle,v11Area.selectionEnd);
 if(pos<0)pos=hay.indexOf(needle,0);
 if(pos<0)return toast('No matches');
 v11Area.focus();v11Area.setSelectionRange(pos,pos+term.length);updateCodeStats();
}
$('#v11Next').onclick=v11FindNext;
$('#v11ReplaceOne').onclick=()=>{
 const term=$('#v11Find').value;if(!term)return;
 if(v11Area.value.slice(v11Area.selectionStart,v11Area.selectionEnd).toLowerCase()!==term.toLowerCase())v11FindNext();
 if(v11Area.value.slice(v11Area.selectionStart,v11Area.selectionEnd).toLowerCase()!==term.toLowerCase())return;
 v11Area.setRangeText($('#v11Replace').value,v11Area.selectionStart,v11Area.selectionEnd,'end');v11Sync();
};
$('#v11ReplaceAll').onclick=()=>{
 const term=$('#v11Find').value;if(!term)return;
 const original=v11Area.value, replacement=$('#v11Replace').value;
 const lower=original.toLowerCase(),needle=term.toLowerCase();let at=0,last=0,count=0,result='';
 while((at=lower.indexOf(needle,last))!==-1){result+=original.slice(last,at)+replacement;last=at+term.length;count++}
 if(!count)return toast('No matches');
 v11Area.value=result+original.slice(last);v11Sync();toast(count+' replacement(s)');
};
$('#v11Format').onclick=()=>{
 try{const formatted=JSON.stringify(JSON.parse(v11Area.value),null,2)+'\n';v11Area.value=formatted;v11Sync();toast('JSON formatted')}
 catch(e){toast('Invalid JSON: '+e.message)}
};
$('#v11Trim').onclick=()=>{
 const value=v11Area.value.replace(/[ \t]+$/gm,'');if(value!==v11Area.value){v11Area.value=value;v11Sync();toast('Trailing spaces removed')}else toast('No trailing spaces');
};
let v11Wrapped=false;
$('#v11Wrap').onclick=()=>{
 v11Wrapped=!v11Wrapped;v11Area.wrap=v11Wrapped?'soft':'off';
 $('#v11Wrap').textContent='Wrap: '+(v11Wrapped?'on':'off');
 $('#codeGutter').style.visibility=v11Wrapped?'hidden':'visible';
};
$('#v11Goto').onclick=()=>{
 const raw=prompt('Go to line number:');if(raw===null)return;
 const line=Number(raw);if(!Number.isInteger(line)||line<1)return toast('Enter a positive line number');
 const lines=v11Area.value.split('\n');if(line>lines.length)return toast('Line does not exist');
 let offset=0;for(let i=0;i<line-1;i++)offset+=lines[i].length+1;
 v11Area.focus();v11Area.setSelectionRange(offset,offset);updateCodeStats();
 const lh=parseFloat(getComputedStyle(v11Area).lineHeight)||20;v11Area.scrollTop=Math.max(0,(line-4)*lh);
};
function v11MoveSlide(delta){
 if(state.mode!=='slides')return toast('Open a presentation first');
 const i=state.currentSlide,j=i+delta;if(j<0||j>=state.slides.length)return;
 pushSlideHistory();[state.slides[i],state.slides[j]]=[state.slides[j],state.slides[i]];
 state.currentSlide=j;renderSlides();markDirty();
}
$('#v11SlideUp').onclick=()=>v11MoveSlide(-1);
$('#v11SlideDown').onclick=()=>v11MoveSlide(1);
v11Area.addEventListener('keydown',e=>{
 if(e.key!=='Enter'||e.ctrlKey||e.altKey||e.metaKey)return;
 const pos=v11Area.selectionStart,prior=v11Area.value.slice(0,pos);
 const line=prior.slice(prior.lastIndexOf('\n')+1),indent=(line.match(/^[ \t]*/)||[''])[0];
 const extra=/[{[(]\s*$/.test(line)?'  ':'';
 e.preventDefault();v11Area.setRangeText('\n'+indent+extra,pos,v11Area.selectionEnd,'end');v11Sync();
});
commands.push(['Find in Code','Code',()=>{if(state.mode==='code')$('#v11Find').focus()}],
 ['Format JSON','Code',()=>state.mode==='code'&&$('#v11Format').click()],
 ['Go to Line','Code',()=>state.mode==='code'&&$('#v11Goto').click()],
 ['Move Slide Up','Presentation',()=>v11MoveSlide(-1)],
 ['Move Slide Down','Presentation',()=>v11MoveSlide(1)]);
document.addEventListener('keydown',e=>{
 if(state.mode==='code'&&(e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='f'){
 e.preventDefault();e.stopImmediatePropagation();$('#v11Find').focus();$('#v11Find').select()}
},true);


$('#v11FocusDoc').onclick=()=>{
 document.body.classList.toggle('v11-doc-focus');
 toast(document.body.classList.contains('v11-doc-focus')?'Focus mode on — press Ctrl+Shift+F to exit':'Focus mode off');
};
$('#v11InsertDate').onclick=()=>{
 if(state.mode!=='docs')return;
 $('#docEditor').focus();
 document.execCommand('insertText',false,new Date().toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'}));
 state.doc.html=$('#docEditor').innerHTML;markDirty();updateDocStats();
};
$('#v11PlainPaste').onclick=()=>{
 if(state.mode!=='docs')return;
 toast('Paste plain text with Ctrl+Shift+V (supported by Chromium)');
};
document.addEventListener('keydown',e=>{
 if(state.mode==='docs'&&e.ctrlKey&&e.shiftKey&&e.key.toLowerCase()==='f'){
 e.preventDefault();document.body.classList.toggle('v11-doc-focus')}
},true);

setMode('home');
