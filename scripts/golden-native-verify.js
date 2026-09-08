"use strict";
const fs=require('fs'),path=require('path'),assert=require('assert'),{performance}=require('perf_hooks');
const AdmZip=require('adm-zip');
const {GoldenCompiler,read,parse,els,relationships,target,hash}=require('../server/golden-native/package');
const out=path.resolve(process.argv[2]||'debug/golden-native');fs.mkdirSync(path.join(out,'compiled'),{recursive:true});
const g=new GoldenCompiler(),source=new AdmZip(g.buffer),results=[];
for(const slide of g.catalog.slides){
 const t=performance.now(),buf=g.reference(slide.id),ms=performance.now()-t,zip=new AdmZip(buf);
 assert.equal(els(parse(read(zip,'ppt/presentation.xml')),'p','sldId').length,1);
 assert.equal(read(zip,slide.part),read(source,slide.part),'slide XML must be identical before data binding');
 let related=0;for(const entry of zip.getEntries()){
  if(!entry.entryName.endsWith('.rels'))continue;
  const part=entry.entryName==='_rels/.rels'?'':entry.entryName.replace(/(^|\/)\_rels\//,'$1').replace(/\.rels$/,'');
  for(const r of els(parse(entry.getData()),'rel','Relationship'))if(r.getAttribute('TargetMode')!=='External'){assert(zip.getEntry(target(part,r.getAttribute('Target'))),'Missing relationship target: '+entry.entryName+' -> '+target(part,r.getAttribute('Target')));related++;}
 }
 for(const chart of slide.charts)assert.equal(read(zip,chart.part),read(source,chart.part),'Native chart must be retained byte-for-byte');
 fs.writeFileSync(path.join(out,'compiled',slide.id+'.pptx'),buf);
 results.push({id:slide.id,shapeCount:slide.shapeCount,tableCount:slide.tableCount,chartCount:slide.charts.length,textSlots:slide.slots.length,compileMs:+ms.toFixed(3),bytes:buf.length,relationships:related,sourceSlideXmlIdentical:true});
}
const warm=[];for(let i=0;i<256;i++){const t=performance.now();g.reference(g.catalog.slides[i%32].id);warm.push(performance.now()-t);}warm.sort((a,b)=>a-b);
const report={source:'docs/golden/golden.pptx',sourceSha256:g.catalog.sourceSha256,total:results.length,nativeCharts:results.reduce((n,r)=>n+r.chartCount,0),nativeTables:results.reduce((n,r)=>n+r.tableCount,0),warmCacheP50Ms:+warm[127].toFixed(4),warmCacheP95Ms:+warm[243].toFixed(4),coldTotalMs:+results.reduce((n,r)=>n+r.compileMs,0).toFixed(3),results};
fs.writeFileSync(path.join(out,'package-report.json'),JSON.stringify(report,null,2));fs.writeFileSync(path.join(out,'catalog.json'),JSON.stringify(g.catalog,null,2));console.log(JSON.stringify({...report,results:undefined},null,2));
