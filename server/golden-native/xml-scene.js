"use strict";
// Native scene operations shared by semantic layout adapters.
const P=require('./package');
function scene(doc){
 const nodes=new Map([...P.els(doc,'p','sp'),...P.els(doc,'p','cxnSp'),...P.els(doc,'p','graphicFrame')].map(n=>[P.els(n,'p','cNvPr')[0]?.getAttribute('id'),n]));
 let next=Math.max(...P.els(doc,'p','cNvPr').map(n=>Number(n.getAttribute('id'))))+1;
 const get=id=>{const n=nodes.get(String(id));if(!n)throw Error('Missing calibrated shape '+id);return n;};
 function rect(n,r){const o=P.els(n,'a','off')[0],e=P.els(n,'a','ext')[0];if(!o||!e)throw Error('Shape has no transform');for(const k of ['x','y'])if(r[k]!=null)o.setAttribute(k,String(Math.round(r[k]*12700)));for(const [k,key]of [['w','cx'],['h','cy']])if(r[k]!=null)e.setAttribute(key,String(Math.max(0,Math.round(r[k]*12700))));return n;}
 function text(n,t){const ps=P.els(n,'a','p');if(!ps.length)throw Error('Shape has no paragraph');P.replaceParagraph(ps[0],String(t));ps.slice(1).forEach(p=>P.replaceParagraph(p,''));return n;}
 function fill(n,color){const sp=Array.from(n.childNodes).find(x=>x.localName==='spPr');const solid=sp&&Array.from(sp.childNodes).find(x=>x.localName==='solidFill');const rgb=solid&&P.els(solid,'a','srgbClr')[0];if(!rgb)throw Error('Shape has no calibrated fill');rgb.setAttribute('val',color);return n;}
 function repeat(ids,items,build){const originals=ids.map(get),anchor=originals[0],parent=anchor.parentNode;const added=[];items.forEach((it,i)=>{const clones=build(it,i,(templateId)=>{const n=get(templateId).cloneNode(true),prop=P.els(n,'p','cNvPr')[0];prop.setAttribute('id',String(next++));prop.setAttribute('name','FS_native_'+prop.getAttribute('id'));return n;});for(const n of clones){parent.insertBefore(n,anchor);added.push(n);}});originals.forEach(n=>n.parentNode.removeChild(n));return added;}
 function remove(id){const n=get(id);if(n.parentNode)n.parentNode.removeChild(n);}
 return{get,rect,text,fill,repeat,remove};
}
module.exports={scene};
