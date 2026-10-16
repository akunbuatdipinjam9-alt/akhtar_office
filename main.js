const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs/promises');
const PptxGenJS = require('pptxgenjs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, ImageRun, Table, TableRow, TableCell, WidthType, PageBreak } = require('docx');
const cheerio = require('cheerio');
const mammoth = require('mammoth');
const AdmZip = require('adm-zip');

let mainWindow;
let pendingExternalPaths=[];
const supportedExts=new Set(['.akoffice','.json','.docx','.pptx','.txt','.md','.html','.htm','.js','.jsx','.ts','.tsx','.css','.scss','.jsonc','.py','.java','.c','.cpp','.h','.hpp','.rs','.go','.sh','.bat','.cmd','.xml','.yaml','.yml','.sql','.php','.rb','.vue','.svelte']);
function incomingPaths(argv){return argv.filter(x=>typeof x==='string'&&supportedExts.has(path.extname(x).toLowerCase())&&path.isAbsolute(x));}
function queueExternal(paths){for(const item of paths){if(!pendingExternalPaths.includes(item))pendingExternalPaths.push(item);}if(mainWindow&&!mainWindow.isDestroyed()&&pendingExternalPaths.length){mainWindow.webContents.send('external:files',pendingExternalPaths.splice(0));mainWindow.show();}}
const gotLock=app.requestSingleInstanceLock();
if(!gotLock){app.quit();}else{
 app.on('second-instance',(_e,argv)=>queueExternal(incomingPaths(argv)));
 app.on('open-file',(event,file)=>{event.preventDefault();queueExternal([file])});
 pendingExternalPaths.push(...incomingPaths(process.argv));
}


function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1500,
    height: 920,
    minWidth: 1080,
    minHeight: 700,
    backgroundColor: '#0b1020',
    titleBarStyle: 'hidden',
    titleBarOverlay: {
      color: '#0b1020',
      symbolColor: '#e8edf7',
      height: 42
    },
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });

  mainWindow.loadFile('index.html');
  mainWindow.webContents.on('did-finish-load',()=>queueExternal([]));
}

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

ipcMain.handle('file:openProject', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [{ name: 'Akhtar Office Project', extensions: ['akoffice', 'json'] }]
  });
  if (result.canceled || !result.filePaths[0]) return null;
  const filePath = result.filePaths[0];
  const text = await fs.readFile(filePath, 'utf8');
  return { filePath, data: JSON.parse(text) };
});

ipcMain.handle('file:saveProject', async (_event, { filePath, data, suggestedName }) => {
  let target = filePath;
  if (!target) {
    const result = await dialog.showSaveDialog(mainWindow, {
      defaultPath: suggestedName || 'Untitled.akoffice',
      filters: [{ name: 'Akhtar Office Project', extensions: ['akoffice'] }]
    });
    if (result.canceled || !result.filePath) return null;
    target = result.filePath;
  }
  await fs.writeFile(target, JSON.stringify(data, null, 2), 'utf8');
  return target;
});

ipcMain.handle('file:saveText', async (_event, { content, suggestedName, filters }) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: suggestedName,
    filters: filters || [{ name: 'Text', extensions: ['txt'] }]
  });
  if (result.canceled || !result.filePath) return null;
  await fs.writeFile(result.filePath, content, 'utf8');
  return result.filePath;
});

ipcMain.handle('file:pickImage', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [{ name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'webp', 'gif'] }]
  });
  if (result.canceled || !result.filePaths[0]) return null;

  const filePath = result.filePaths[0];
  const ext = path.extname(filePath).toLowerCase();
  const mime = {
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.webp': 'image/webp',
    '.gif': 'image/gif'
  }[ext] || 'application/octet-stream';
  const buffer = await fs.readFile(filePath);
  return {
    name: path.basename(filePath),
    dataUrl: `data:${mime};base64,${buffer.toString('base64')}`
  };
});

ipcMain.handle('window:printPdf', async (_event, suggestedName) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: suggestedName || 'export.pdf',
    filters: [{ name: 'PDF', extensions: ['pdf'] }]
  });
  if (result.canceled || !result.filePath) return null;
  const pdf = await mainWindow.webContents.printToPDF({
    printBackground: true,
    preferCSSPageSize: true
  });
  await fs.writeFile(result.filePath, pdf);
  return result.filePath;
});

ipcMain.handle('slides:exportPptx', async (_event, { title, slides }) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: `${sanitizeFileName(title || 'Presentation')}.pptx`,
    filters: [{ name: 'PowerPoint Presentation', extensions: ['pptx'] }]
  });
  if (result.canceled || !result.filePath) return null;

  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Akhtar Office';
  pptx.subject = title || 'Presentation';
  pptx.title = title || 'Presentation';
  pptx.company = 'Akhtar Office';
  pptx.lang = 'id-ID';
  pptx.theme = {
    headFontFace: 'Aptos Display',
    bodyFontFace: 'Aptos',
    lang: 'id-ID'
  };

  const pxToIn = 1 / 72;

  for (const source of slides || []) {
    const slide = pptx.addSlide();
    slide.background = { color: normalizeHex(source.bg, 'FFFFFF') };

    for (const el of source.elements || []) {
      const x = Math.max(0, Number(el.x) || 0) * pxToIn;
      const y = Math.max(0, Number(el.y) || 0) * pxToIn;
      const w = Math.max(1, Number(el.w) || 1) * pxToIn;
      const h = Math.max(1, Number(el.h) || 1) * pxToIn;

      if (el.type === 'text') {
        slide.addText(String(el.text || ''), {
          x, y, w, h,
          fontFace: el.fontFamily || 'Aptos',
          fontSize: Math.max(6, Number(el.fontSize) || 20),
          bold: String(el.fontWeight || '') === '700' || String(el.fontWeight || '') === 'bold',
          italic: !!el.italic,
          underline: el.underline ? { style: 'sng' } : undefined,
          color: normalizeHex(el.color, '111111'),
          margin: 0.04,
          breakLine: false,
          valign: el.valign || 'mid',
          align: el.align || 'left',
          fit: 'shrink',
          transparency: 0
        });
      } else if (el.type === 'shape') {
        const shapeType = el.shape === 'circle' ? pptx.ShapeType.ellipse : pptx.ShapeType.rect;
        slide.addShape(shapeType, {
          x, y, w, h,
          fill: { color: normalizeHex(el.fill, '7C8CFF'), transparency: el.fill === 'transparent' ? 100 : 0 },
          line: { color: normalizeHex(el.borderColor, normalizeHex(el.fill, '7C8CFF')), transparency: el.border === false ? 100 : 0, width: Number(el.borderWidth) || 1 },
          radius: el.shape === 'circle' ? undefined : 0.12
        });
      } else if (el.type === 'image' && typeof el.src === 'string' && el.src.startsWith('data:image/')) {
        slide.addImage({ data: el.src, x, y, w, h, transparency: Math.max(0, Math.min(100, Number(el.transparency) || 0)) });
      }
    }

    if (source.notes) {
      try {
        if (typeof slide.addNotes === 'function') slide.addNotes(String(source.notes));
      } catch (_) {}
    }
  }

  await pptx.writeFile({ fileName: result.filePath });
  return result.filePath;
});

ipcMain.handle('app:getVersion', () => app.getVersion());

function sanitizeFileName(value) {
  return String(value).replace(/[\\/:*?"<>|]/g, '_').trim() || 'Untitled';
}

function normalizeHex(value, fallback = 'FFFFFF') {
  const s = String(value || '').replace('#', '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(s) ? s : fallback.replace('#', '').toUpperCase();
}

ipcMain.handle('file:writeRecovery',async(_event,data)=>{
  if(!data||!['docs','slides'].includes(data.type))throw new Error('Invalid recovery data');
  const file=path.join(app.getPath('userData'),'akhtar-office-recovery.json');
  const tmp=file+'.tmp';await fs.writeFile(tmp,JSON.stringify(data),'utf8');await fs.rename(tmp,file);return true;
});
ipcMain.handle('file:readRecovery',async()=>{
 try{return JSON.parse(await fs.readFile(path.join(app.getPath('userData'),'akhtar-office-recovery.json'),'utf8'))}catch(e){if(e.code==='ENOENT')return null;throw e}
});
function toRuns($,node,opts={}){
 if(node.type==='text')return [new TextRun({text:node.data||'',...opts})];
 if(node.type!=='tag')return [];
 const name=node.name||'', next={...opts};
 if(['b','strong'].includes(name))next.bold=true;
 if(['i','em'].includes(name))next.italics=true;
 if(name==='u')next.underline={};
 if(name==='br')return [new TextRun({break:1})];
 if(name==='img')return [];
 return (node.children||[]).flatMap(n=>toRuns($,n,next));
}
ipcMain.handle('docs:exportDocx',async(_event,{title,html,pageSize,margin})=>{
 const save=await dialog.showSaveDialog(mainWindow,{defaultPath:sanitizeFileName(title||'Document')+'.docx',filters:[{name:'Word document',extensions:['docx']}]});
 if(save.canceled||!save.filePath)return null;
 const $=cheerio.load('<main>'+String(html||'')+'</main>');const children=[];
 function visit(node){
  if(node.type!=='tag')return;
  const tag=node.name;
  if(tag==='img'){
   const data=$(node).attr('src')||'';const match=/^data:image\/(png|jpeg|jpg);base64,([A-Za-z0-9+/=]+)$/.exec(data);
   if(match){const buf=Buffer.from(match[2],'base64');children.push(new Paragraph({children:[new ImageRun({data:buf,transformation:{width:450,height:280},type:match[1]==='png'?'png':'jpg'})]}))}return;
  }
  if($(node).hasClass('doc-pagebreak')){children.push(new Paragraph({children:[new PageBreak()]}));return}
  if(['p','h1','h2','h3','blockquote','li'].includes(tag)){
   const runs=(node.children||[]).flatMap(n=>toRuns($,n));
   children.push(new Paragraph({children:runs.length?runs:[new TextRun('')],heading:tag==='h1'?HeadingLevel.HEADING_1:tag==='h2'?HeadingLevel.HEADING_2:tag==='h3'?HeadingLevel.HEADING_3:undefined,bullet:tag==='li'?{level:0}:undefined,indent:tag==='blockquote'?{left:700}:undefined,spacing:{after:150}}));return;
  }
  if(tag==='table'){
   const raw=$(node).find('tr').toArray().map(tr=>$(tr).children('td,th').toArray()).filter(cells=>cells.length);
   if(raw.length&&raw.every(cells=>cells.length===raw[0].length)){const rows=raw.map(cells=>new TableRow({children:cells.map(cell=>new TableCell({children:[new Paragraph($(cell).text())]}))}));children.push(new Table({rows,width:{size:100,type:WidthType.PERCENTAGE}}))}return;
  }
  if(tag==='ul'||tag==='ol'){for(const c of node.children||[])visit(c);return}
  for(const child of node.children||[])visit(child);
 }
 for(const node of $('main').get(0).children)visit(node);
 const margins={narrow:850,normal:1417,wide:1984};const m=margins[margin]||margins.normal;
 const doc=new Document({creator:'Akhtar Office',title:String(title||'Document'),sections:[{properties:{page:{size:pageSize==='letter'?{width:12240,height:15840}:{width:11906,height:16838},margin:{top:m,bottom:m,left:m,right:m}}},children:children.length?children:[new Paragraph('')]}]});
 await fs.writeFile(save.filePath,await Packer.toBuffer(doc));return save.filePath;
});

const sourceExts=new Set(['.txt','.md','.html','.htm','.js','.jsx','.ts','.tsx','.css','.scss','.json','.jsonc','.py','.java','.c','.cpp','.h','.hpp','.rs','.go','.sh','.bat','.cmd','.xml','.yaml','.yml','.sql','.php','.rb','.vue','.svelte']);
function htmlEscape(value){return String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
async function importFile(filePath){
 const ext=path.extname(filePath).toLowerCase();
 if(!supportedExts.has(ext))throw new Error('Unsupported file extension');
 const st=await fs.stat(filePath);if(!st.isFile()||st.size>30*1024*1024)throw new Error('File must be under 30 MB');
 if(ext==='.akoffice'||ext==='.json'){
   const content=await fs.readFile(filePath,'utf8');
   if(ext==='.akoffice')return {kind:'project',filePath,data:JSON.parse(content)};
   return {kind:'code',filePath,title:path.basename(filePath),content,language:ext.slice(1)};
 }
 if(sourceExts.has(ext))return {kind:'code',filePath,title:path.basename(filePath),content:await fs.readFile(filePath,'utf8'),language:ext.slice(1)};
 if(ext==='.docx'){
   const result=await mammoth.convertToHtml({path:filePath});
   return {kind:'docs',filePath:null,sourcePath:filePath,title:path.basename(filePath,ext),doc:{html:result.value,pageSize:'a4'},importWarnings:result.messages.map(x=>x.message).slice(0,6)};
 }
 if(ext==='.pptx'){
   const zip=new AdmZip(filePath);const files=zip.getEntries().filter(e=>/^ppt\/slides\/slide\d+\.xml$/.test(e.entryName)).sort((a,b)=>Number(a.entryName.match(/slide(\d+)/)[1])-Number(b.entryName.match(/slide(\d+)/)[1]));
   if(!files.length)throw new Error('No slides found');
   const slides=files.map((entry,i)=>{
     const xml=entry.getData().toString('utf8');const texts=[...xml.matchAll(/<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g)].map(m=>cheerio.load('<x>'+m[1]+'</x>',{xmlMode:true})('x').text()).filter(Boolean);
     return {id:'import-slide-'+i,bg:'#ffffff',transition:'none',notes:'',elements:texts.slice(0,14).map((t,j)=>({id:'import-text-'+i+'-'+j,type:'text',x:70,y:45+j*34,w:820,h:33,text:t,fontSize:j===0?28:19,color:'#222222',fontWeight:j===0?'700':'400'}))};
   });
   return {kind:'slides',filePath:null,sourcePath:filePath,title:path.basename(filePath,ext),slides,importWarnings:['PPTX imported as editable text only. Original images, geometry, fonts, transitions and animations are not preserved.']};
 }
 throw new Error('Unsupported file type');
}
ipcMain.handle('file:openAny',async()=>{
 const result=await dialog.showOpenDialog(mainWindow,{properties:['openFile'],filters:[{name:'Office and code files',extensions:[...supportedExts].map(x=>x.slice(1))},{name:'All files',extensions:['*']}]});
 if(result.canceled||!result.filePaths[0])return null;
 return importFile(result.filePaths[0]);
});
ipcMain.handle('file:readExternal',async(_event,filePath)=>{
 if(typeof filePath!=='string'||!path.isAbsolute(filePath))throw new Error('Invalid path');
 return importFile(filePath);
});
ipcMain.handle('code:save',async(_event,{filePath,content,suggestedName,saveAs})=>{
 if(typeof content!=='string'||content.length>15*1024*1024)throw new Error('Source too large');
 let target=!saveAs&&typeof filePath==='string'?filePath:null;
 if(!target){const result=await dialog.showSaveDialog(mainWindow,{defaultPath:suggestedName||'untitled.js',filters:[{name:'Source files',extensions:['js','ts','py','html','css','json','md','txt']},{name:'All files',extensions:['*']}]});if(result.canceled)return null;target=result.filePath;}
 if(!path.isAbsolute(target))throw new Error('Invalid path');
 await fs.writeFile(target,content,'utf8');return target;
});
