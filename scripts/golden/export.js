"use strict";
// Artifact exporter only. Production placement remains Office.js + layout.js.
const fs=require('fs'),path=require('path');
const PptxGenJS=require(process.env.PPTXGENJS_PATH || 'pptxgenjs');
const icons=require('../../server/icons');
const rgb=x=>(x||'#252525').replace('#','');
function runs(p){if(!p.boldRanges?.length)return p.text;const rs=[];let i=0;for(const [s,n] of p.boldRanges){if(s>i)rs.push({text:p.text.slice(i,s)});rs.push({text:p.text.slice(s,s+n),options:{bold:true}});i=s+n;}if(i<p.text.length)rs.push({text:p.text.slice(i)});return rs;}
async function exportScenes(entries,file){
 const deck=new PptxGenJS();deck.defineLayout({name:'GOLDEN',width:960/72,height:540/72});deck.layout='GOLDEN';deck.author='Astra';deck.subject='Synthetic slide design evaluation';deck.title='Golden Slides v1';deck.lang='ja-JP';deck.theme={headFontFace:'Yu Gothic',bodyFontFace:'Yu Gothic',lang:'ja-JP'};
 for(const e of entries){const slide=deck.addSlide();slide.background={color:'FFFFFF'};const sc=e.golden||e.scene;
  for(const p of sc.prims){
   if(p.kind==='line'){slide.addShape(deck.ShapeType.line,{x:Math.min(p.x1,p.x2)/72,y:Math.min(p.y1,p.y2)/72,w:Math.max(.1,Math.abs(p.x2-p.x1))/72,h:Math.max(.1,Math.abs(p.y2-p.y1))/72,flipV:(p.x2-p.x1)*(p.y2-p.y1)<0,line:{color:rgb(p.color),width:p.weight||1}});continue;}
   if(p.kind==='image'){const png=icons.renderPng(p.name,p.color,512);if(!png)throw Error('Unknown icon '+p.name);slide.addImage({data:'image/png;base64,'+png.toString('base64'),x:p.x/72,y:p.y/72,w:p.w/72,h:p.h/72});continue;}
   if(p.kind==='table'){slide.addTable(p.cells.map(row=>row.map(c=>({text:String(c.text||''),options:{bold:!!c.bold,color:rgb(c.color),fill:rgb(c.fill||'#FFFFFF'),fontSize:c.fontSize||p.fontSize||18,align:c.align||'left'}}))),{x:p.x/72,y:p.y/72,w:p.w/72,h:p.h/72,colW:p.colWidths.map(x=>x/72),rowH:p.rowHeights.map(x=>x/72),fontFace:'Yu Gothic',margin:4,border:{type:'solid',color:'DDDDDD',pt:.5},autoPage:false,verbose:false});continue;}
   if(p.kind!=='rect')continue;
   if(p.fill||p.line||p.shape&&p.shape!=='rect'){slide.addShape(deck.ShapeType[p.shape]||deck.ShapeType.rect,{x:p.x/72,y:p.y/72,w:p.w/72,h:p.h/72,rotate:p.rotation||0,fill:{color:rgb(p.fill||'#FFFFFF'),transparency:p.fill?0:100},line:{color:rgb(p.line||p.fill||'#FFFFFF'),width:p.line?p.lineWeight||1:0,transparency:p.line?0:100}});}
   if(p.text)slide.addText(runs(p),{x:p.x/72,y:p.y/72,w:p.w/72,h:p.h/72,fontFace:p.fontName||'Yu Gothic',fontSize:p.fontSize||18,color:rgb(p.color),bold:!!p.bold,align:p.align||'left',valign:p.valign==='middle'?'mid':p.valign||'top',margin:p.pad??0,breakLine:false,paraSpaceAfter:0,wrap:p.wrap!==false,vertAnchor:'ctr'});
  }
  slide.addNotes(`Case: ${e.id}\nSynthetic evaluation data.\nDesign: ${e.design?.reason||''}`);
 }
 await deck.writeFile({fileName:path.resolve(file)});
}
if(require.main===module){const input=process.argv[2]||'test/golden/cases.json',out=process.argv[3]||'debug/golden-v1/golden-slides.pptx';const data=JSON.parse(fs.readFileSync(input));exportScenes(data.cases||data,out).then(()=>console.log(path.resolve(out))).catch(e=>{console.error(e);process.exitCode=1;});}
module.exports={exportScenes};
