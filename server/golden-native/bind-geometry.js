"use strict";
const solver=require('./geometry'),bindings=require('./bindings.json');
function bindGeometry(doc,slideId,data,sourceHash){
 const {els}=require('./package');
 if(sourceHash!==bindings.sourceSha256)throw Error('Golden source changed; recalibrate geometry bindings');
 const b=bindings.slides[slideId];if(!b)throw Error('No semantic geometry binding for this layout');
 if(b.kind==='gantt')return require('./gantt').bindGantt(doc,data,b);
 if(!Array.isArray(data?.values)||data.values.length!==b.shapes.length)throw Error('Quantitative values must match the bound shapes');
 const shapes=new Map([...els(doc,'p','sp'),...els(doc,'p','cxnSp')].map(s=>[els(s,'p','cNvPr')[0]?.getAttribute('id'),s]));
 function setRect(id,r){const s=shapes.get(id);if(!s)throw Error('Bound shape missing: '+id);const o=els(s,'a','off')[0],e=els(s,'a','ext')[0];for(const [k,a] of [['x','x'],['y','y']])if(r[k]!=null)o.setAttribute(a,String(Math.round(r[k]*12700)));for(const [k,a] of [['w','cx'],['h','cy']])if(r[k]!=null)e.setAttribute(a,String(Math.max(1,Math.round(r[k]*12700))));if(r.visible!=null){const c=els(s,'p','cNvPr')[0];if(r.visible)c.removeAttribute('hidden');else c.setAttribute('hidden','1');}}
 function label(id,value){const s=shapes.get(id);if(!s)throw Error('Bound label missing');const ps=els(s,'a','p');require('./package').replaceParagraph(ps[0],value);for(const p of ps.slice(1))require('./package').replaceParagraph(p,'');}
 const rects=b.kind==='areas'?solver.areas(data.values,b.centers,b.maxDiameter):solver.bars(data.values,b.plot,...b.domain);
 if(b.targetShape){const x=b.plot.x+solver.domain(...b.domain)(data.target)*b.plot.w;setRect(b.targetShape,{x});setRect(b.targetLabel,{x:x+b.targetLabelOffset});label(b.targetLabel,'目標 '+data.target+b.unit);}
 rects.forEach((r,i)=>{setRect(b.shapes[i],r);label(b.labels[i],b.decimals!=null?data.values[i].toFixed(b.decimals):String(data.values[i])+(b.unit||''));if(b.kind==='bars')setRect(b.labels[i],{x:r.end+b.labelGap});if(b.kind==='areas'&&r.w<44)for(const color of els(shapes.get(b.labels[i]),'a','srgbClr'))color.setAttribute('val','1A1A1A');});
 return {kind:b.kind,rects};
}
module.exports={bindGeometry};
