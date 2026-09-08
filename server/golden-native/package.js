"use strict";
// OOXML is the intermediate representation. Office.js inserts the compiled slide
// in one operation, retaining native charts, tables, connectors and the master.
const AdmZip=require('adm-zip');
const {DOMParser,XMLSerializer}=require('@xmldom/xmldom');
const fs=require('fs'),path=require('path'),crypto=require('crypto');
const NS={p:'http://schemas.openxmlformats.org/presentationml/2006/main',a:'http://schemas.openxmlformats.org/drawingml/2006/main',r:'http://schemas.openxmlformats.org/officeDocument/2006/relationships',rel:'http://schemas.openxmlformats.org/package/2006/relationships',c:'http://schemas.openxmlformats.org/drawingml/2006/chart'};
const parse=s=>new DOMParser({errorHandler:{warning(){},error(e){throw Error(e)},fatalError(e){throw Error(e)}}}).parseFromString(String(s),'application/xml');
const xml=d=>new XMLSerializer().serializeToString(d);
const els=(n,ns,tag)=>Array.from(n.getElementsByTagNameNS(NS[ns]||ns,tag));
const read=(zip,name)=>{const e=zip.getEntry(name);if(!e)throw Error('Missing OPC part: '+name);return e.getData().toString('utf8');};
const relfile=p=>path.posix.join(path.posix.dirname(p),'_rels',path.posix.basename(p)+'.rels');
const target=(from,t)=>path.posix.normalize(t.startsWith('/')?t.slice(1):path.posix.join(path.posix.dirname(from),t));
const relationships=(zip,part)=>zip.getEntry(relfile(part))?els(parse(read(zip,relfile(part))),'rel','Relationship'):[];
const hash=b=>crypto.createHash('sha256').update(b).digest('hex');
function setText(node,value){while(node.firstChild)node.removeChild(node.firstChild);node.appendChild(node.ownerDocument.createTextNode(String(value)));}
function paragraphText(p){return Array.from(p.childNodes).map(n=>n.nodeType!==1?'':n.localName==='br'?'\n':els(n,'a','t').map(t=>t.textContent).join('')).join('');}
function ancestor(n,local){for(let p=n.parentNode;p;p=p.parentNode)if(p.localName===local)return p;return null;}
function weightedLength(s){return Array.from(s).reduce((n,ch)=>n+(/[\u0000-\u007f]/.test(ch)?.55:1),0);}
function textSlots(doc){return els(doc,'a','p').map((p,i)=>{
 const shape=ancestor(p,'sp')||ancestor(p,'graphicFrame');const text=paragraphText(p);
 const off=shape&&els(shape,'a','off')[0],ext=shape&&els(shape,'a','ext')[0];
 const x=Number(off?.getAttribute('x')||0)/12700,y=Number(off?.getAttribute('y')||0)/12700,w=Number(ext?.getAttribute('cx')||0)/12700,h=Number(ext?.getAttribute('cy')||0)/12700;
 const rp=els(p,'a','rPr')[0]||els(p,'a','endParaRPr')[0];const fs=Number(rp?.getAttribute('sz')||1150)/100;
 const body=ancestor(p,'txBody'),n=body?Array.from(body.childNodes).filter(n=>n.localName==='p').length:1;
 const capacity=Math.floor(Math.max(weightedLength(text),w/Math.max(8,fs)*Math.max(1,h/Math.max(1,n)/(fs*1.15))*.88));
 return {id:'p'+i,text,role:y>=510?'page':y>=490?'source':y<89?'title':y<125?'lead':'body',fontSize:fs,rect:{x,y,w,h},capacity,paragraphsInShape:n};
 });}
function chartInfo(zip,part){const d=parse(read(zip,part));return {part,type:els(d,'c','plotArea')[0]&&Array.from(els(d,'c','plotArea')[0].childNodes).find(n=>/Chart$/.test(n.localName||''))?.localName,series:els(d,'c','ser').map(s=>({name:els(els(s,'c','tx')[0]||s,'c','v')[0]?.textContent||'',categories:els(els(s,'c','cat')[0]||parse('<x/>'),'c','pt').map(p=>els(p,'c','v')[0]?.textContent||''),values:els(els(s,'c','val')[0]||parse('<x/>'),'c','pt').map(p=>Number(els(p,'c','v')[0]?.textContent))}))};}
function compileSource(filename){
 const buffer=fs.readFileSync(filename),zip=new AdmZip(buffer),pres=parse(read(zip,'ppt/presentation.xml'));
 const refs=new Map(relationships(zip,'ppt/presentation.xml').map(r=>[r.getAttribute('Id'),target('ppt/presentation.xml',r.getAttribute('Target'))]));
 const size=els(pres,'p','sldSz')[0];
 const slides=els(pres,'p','sldId').map((s,i)=>{const part=refs.get(s.getAttributeNS(NS.r,'id'));const d=parse(read(zip,part));const slots=textSlots(d);const chartParts=relationships(zip,part).filter(r=>r.getAttribute('Type').endsWith('/chart')).map(r=>target(part,r.getAttribute('Target')));return {id:String(i+1).padStart(2,'0'),index:i,slideId:s.getAttribute('id'),part,title:slots.filter(s=>s.role==='title').map(s=>s.text).join(' '),slots,charts:chartParts.map(p=>chartInfo(zip,p)),shapeCount:els(d,'p','sp').length,tableCount:els(d,'a','tbl').length};});
 return {version:1,sourceSha256:hash(buffer),width:Number(size.getAttribute('cx'))/12700,height:Number(size.getAttribute('cy'))/12700,slides};
}
function singleSlide(buffer,slide){
 const zip=new AdmZip(buffer),pres=parse(read(zip,'ppt/presentation.xml'));
 for(const s of els(pres,'p','sldId'))if(s.getAttribute('id')!==slide.slideId)s.parentNode.removeChild(s);
 // Cross-slide navigation is not a reusable layout constraint.
 for(const tag of ['custShowLst','sectionLst'])for(const n of els(pres,'p',tag))n.parentNode.removeChild(n);
 zip.updateFile('ppt/presentation.xml',Buffer.from(xml(pres)));
 const rd=parse(read(zip,'ppt/_rels/presentation.xml.rels'));
 for(const r of els(rd,'rel','Relationship'))if(r.getAttribute('Type').endsWith('/slide')&&target('ppt/presentation.xml',r.getAttribute('Target'))!==slide.part)r.parentNode.removeChild(r);
 zip.updateFile('ppt/_rels/presentation.xml.rels',Buffer.from(xml(rd)));
 const root=parse(read(zip,'_rels/.rels'));
 for(const r of els(root,'rel','Relationship'))if(r.getAttribute('Type').endsWith('/thumbnail'))r.parentNode.removeChild(r);
 zip.updateFile('_rels/.rels',Buffer.from(xml(root)));
 const keep=new Set(['[Content_Types].xml','_rels/.rels']);
 function visit(part){if(keep.has(part))return;if(!zip.getEntry(part))throw Error('Broken OPC relationship '+part);keep.add(part);if(zip.getEntry(relfile(part))){keep.add(relfile(part));for(const r of relationships(zip,part))if(r.getAttribute('TargetMode')!=='External')visit(target(part,r.getAttribute('Target')));}}
 for(const r of els(root,'rel','Relationship'))if(r.getAttribute('TargetMode')!=='External')visit(target('',r.getAttribute('Target')));
 for(const e of [...zip.getEntries()])if(!keep.has(e.entryName))zip.deleteFile(e.entryName);
 const types=parse(read(zip,'[Content_Types].xml'));
 for(const n of Array.from(types.documentElement.childNodes))if(n.localName==='Override'&&!keep.has(n.getAttribute('PartName').replace(/^\//,'')))n.parentNode.removeChild(n);
 zip.updateFile('[Content_Types].xml',Buffer.from(xml(types)));
 if(zip.getEntry('docProps/app.xml')){const app=parse(read(zip,'docProps/app.xml'));for(const n of Array.from(app.getElementsByTagName('Slides')))setText(n,1);zip.updateFile('docProps/app.xml',Buffer.from(xml(app)));}
 return zip.toBuffer();
}
function replaceParagraph(p,value){
 if(String(value)===paragraphText(p))return; // Preserve exact run styling for unchanged content.
 const oldRuns=els(p,'a','r'),old=oldRuns[0];const props=old&&els(old,'a','rPr')[0];
 for(const n of Array.from(p.childNodes))if(['r','br','fld'].includes(n.localName))p.removeChild(n);
 const fragments=String(value).split('**');
 fragments.forEach((fragment,i)=>{if(!fragment&&i>0)return;fragment.split('\n').forEach((line,j)=>{if(j)p.appendChild(p.ownerDocument.createElementNS(NS.a,'a:br'));const r=p.ownerDocument.createElementNS(NS.a,'a:r');const rp=props?props.cloneNode(true):p.ownerDocument.createElementNS(NS.a,'a:rPr');if(fragments.length>1)rp.setAttribute('b',i%2?'1':'0');r.appendChild(rp);const t=p.ownerDocument.createElementNS(NS.a,'a:t');setText(t,line);r.appendChild(t);const end=Array.from(p.childNodes).find(n=>n.localName==='endParaRPr');p.insertBefore(r,end||null);});});
}
class GoldenCompiler {
 constructor(filename=path.resolve(__dirname,'../../docs/golden/golden.pptx')){this.filename=filename;this.buffer=fs.readFileSync(filename);this.catalog=compileSource(filename);this.cache=new Map();}
 get(id){const slide=this.catalog.slides.find(s=>s.id===String(id).padStart(2,'0'));if(!slide)throw Error('Unknown golden layout '+id);return slide;}
 reference(id){const slide=this.get(id);if(!this.cache.has(slide.id))this.cache.set(slide.id,singleSlide(this.buffer,slide));return Buffer.from(this.cache.get(slide.id));}
 fill(id,plan){
  // Research path only until all geometric/semantic annotations are bound.
  // A text-only update can otherwise leave a circle area or arrow pointing at
  // the old value. The production generation endpoint must not call this yet.
  if(plan?.experimental!==true)throw Error('Semantic geometry binding is incomplete; experimental:true is required for isolated verification');
  const slide=this.get(id);if(!plan||!plan.texts||Array.isArray(plan.texts))throw Error('texts must explicitly account for every paragraph slot');
  const expected=new Set(slide.slots.map(s=>s.id));for(const k of Object.keys(plan.texts))if(!expected.has(k))throw Error('Unknown text slot '+k);
  const missing=slide.slots.filter(s=>!Object.hasOwn(plan.texts,s.id));if(missing.length)throw Error('Missing text slots: '+missing.map(s=>s.id).join(','));
  for(const s of slide.slots){const t=plan.texts[s.id];if(typeof t!=='string')throw Error('Expected text string for '+s.id);if(weightedLength(t.replace(/\*\*/g,''))>s.capacity+1)throw Error('Text exceeds capacity: '+s.id);}
  if(slide.charts.length&&(!Array.isArray(plan.charts)||plan.charts.length!==slide.charts.length))throw Error('Every native chart requires an explicit data binding');
  const zip=new AdmZip(this.reference(id)),d=parse(read(zip,slide.part));
  els(d,'a','p').forEach((p,i)=>replaceParagraph(p,plan.texts['p'+i]));
  if(plan.geometry)require('./bind-geometry').bindGeometry(d,slide.id,plan.geometry,this.catalog.sourceSha256);
  zip.updateFile(slide.part,Buffer.from(xml(d)));
  slide.charts.forEach((chart,i)=>{
    const compiled=require('./charts').bindChart(read(zip,chart.part),plan.charts[i]);
    const workbookRel=relationships(zip,chart.part).find(r=>r.getAttribute('Type').endsWith('/package'));
    if(!workbookRel)throw Error('Native chart has no editable workbook');
    zip.updateFile(chart.part,Buffer.from(compiled.xml));
    zip.updateFile(target(chart.part,workbookRel.getAttribute('Target')),compiled.workbook);
  });
  // Notes are source content, not a design asset. They must not survive adaptation.
  for(const rel of relationships(zip,slide.part).filter(r=>r.getAttribute('Type').endsWith('/notesSlide'))){const part=target(slide.part,rel.getAttribute('Target'));const note=parse(read(zip,part));for(const t of els(note,'a','t'))setText(t,'');zip.updateFile(part,Buffer.from(xml(note)));}
  return zip.toBuffer();
 }
}
module.exports={GoldenCompiler,compileSource,read,parse,els,xml,relationships,target,relfile,hash,weightedLength,replaceParagraph};
