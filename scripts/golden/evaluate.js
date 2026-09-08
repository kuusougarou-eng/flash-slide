"use strict";
const fs=require('fs'),path=require('path'),{performance}=require('perf_hooks');
const canonical=s=>String(s??'').normalize('NFKC').replace(/[\s,*%％]/g,'').toLowerCase();
const flatten=x=>typeof x==='string'||typeof x==='number'?String(x):Array.isArray(x)?x.map(flatten).join(' '):x&&typeof x==='object'?Object.values(x).map(flatten).join(' '):'';
function family(s){return s.panelCount>1?'panels:'+s.panelCount:s.body?.type||'unknown';}
const header=s=>canonical(s).replace(/^進捗$/,'状態');
function valueMatches(actual,expected){const a=canonical(actual),e=canonical(expected);const nums=x=>x.match(/\d+(?:\.\d+)?/g)||[];return a.includes(e)&&JSON.stringify(nums(a))===JSON.stringify(nums(e));}
function tableBinding(actual,row,headers){
 if(!actual?.rows)return false;
 const direct=actual.rows.find(r=>header(r.head)===header(row.head)||header(flatten(r.cells?.[0]))===header(row.head));
 if(direct){const inlineId=!direct.head&&header(flatten(direct.cells?.[0]))===header(row.head);return row.cells.every((v,i)=>{const j=actual.colHeaders?.findIndex(h=>header(h)===header(headers[i]));return valueMatches(flatten(direct.cells[j>=0?j:i+(inlineId?1:0)]),v);});}
 const col=actual.colHeaders?.findIndex(h=>header(h)===header(row.head));
 if(!(col>=0))return false;
 return row.cells.every((v,i)=>{const r=actual.rows.find(r=>header(r.head)===header(headers[i])||canonical(flatten(r.cells?.[col])).startsWith(header(headers[i])+':'));return !!r&&valueMatches(flatten(r.cells[col]),v);});
}
function assess(c,spec,lay){
 const expected=c.design.checks,all=canonical(flatten(spec)),text=canonical(lay.prims.map(p=>p.kind==='table'?flatten(p.cells):p.text||'').join(' '));
 const missing=c.required.filter(t=>!all.includes(canonical(t))),notRendered=c.required.filter(t=>!text.includes(canonical(t)));
 const textPrims=lay.prims.filter(p=>p.text&&p.role!=='title'&&p.role!=='lead'&&p.role!=='footnote'&&p.role!=='source');
 const tableFonts=lay.prims.filter(p=>p.kind==='table').flatMap(p=>p.cells.flat().map(c=>c.fontSize||p.fontSize||18));
 const minFont=Math.min(...textPrims.map(p=>p.fontSize||18),...tableFonts);
 const outside=lay.prims.filter(p=>p.kind!=='line'&&[p.x,p.y,p.w,p.h].every(Number.isFinite)&&(p.x<-.5||p.y<-.5||p.x+p.w>960.5||p.y+p.h>540.5)).length;
 const bindings=[];const eb=c.expected.body,ab=spec.body;
 if(eb?.type==='bars')for(const it of eb.items){const hit=(ab?.items||[]).find(x=>canonical(x.label)===canonical(it.label));bindings.push({label:it.label,ok:!!hit&&Number(hit.value)===it.value});}
 if(eb?.type==='line')for(let i=0;i<eb.labels.length;i++){const j=ab?.labels?.findIndex(x=>canonical(x)===canonical(eb.labels[i]));bindings.push({label:eb.labels[i],ok:j>=0&&ab.series?.[0]?.values?.[j]===eb.series[0].values[i]});}
 if(['table','ntable'].includes(eb?.type))for(const row of eb.rows)bindings.push({label:row.head,ok:tableBinding(ab,row,eb.colHeaders)});
 const duplicates=spec.note?.text&&[spec.title,spec.lead].filter(Boolean).some(t=>canonical(t)===canonical(spec.note.text));
 const icons=lay.prims.filter(p=>p.kind==='image').length;
 const shapeOK=expected.family!== 'sequence'||expected.horizontal===false||lay.prims.filter(p=>p.shape==='homePlate'&&!p.rotation).length===(spec.body?.steps?.length||0);
 const geometricErrors=(lay.warnings||[]).filter(w=>/overflow|failed|dropped|too small/.test(w));
 const gates={family:family(spec)===expected.family,content:!missing.length,renderedContent:!notRendered.length,bindings:bindings.every(x=>x.ok),bounds:outside===0,readable:minFont>=expected.minBodyFont-.1,noOverflow:!geometricErrors.length,noDuplicateNote:!duplicates,nativeTable:!expected.nativeTable||lay.prims.some(p=>p.kind==='table'),icons:expected.icons==null||icons===expected.icons,direction:shapeOK,bold:!expected.bold||lay.prims.some(p=>p.boldRanges?.length)};
 return {id:c.id,split:c.split,family:family(spec),gates,passed:Object.values(gates).every(Boolean),missing,notRendered,bindings,minFont,icons,outside,warnings:lay.warnings||[],primitiveCount:lay.prims.length};
}
async function run(opts={}){
 require('dotenv').config({quiet:true});
 const L=require(path.resolve(opts.layout||'public/layout.js'));const P=require(path.resolve(opts.prompt||'server/prompt.js'));
 const dataset=JSON.parse(fs.readFileSync('test/golden/cases.json'));const cases=dataset.cases.filter(c=>!opts.split||c.split===opts.split);
 const out=path.resolve(opts.out||'debug/golden-v1/oracle');fs.mkdirSync(out,{recursive:true});
 let cursor=0;const results=new Array(cases.length),scenes=new Array(cases.length);
 async function worker(){while(cursor<cases.length){const i=cursor++,c=cases[i];const start=performance.now();let calls=0;try{
   let raw=c.expected,usage=null,model=null;
   if(opts.live){calls++;const r=await require('../../server/llm').chat({messages:P.buildMessages({prompt:c.input,maxSlides:1}),maxTokens:2400,temperature:.2,jsonMode:true});raw=P.extractJson(r.content);usage=r.usage;model=r.model;}
   const inferenceMs=performance.now()-start;
   if(raw.sections||raw.slides?.length>1)throw Error('Expected one slide for this case');
   const spec=L.normalizeSpec(raw.slides?.[0]||raw),t=performance.now();const lay=L.layout(spec,{width:960,height:540,palette:{accent:'none'}});const layoutMs=performance.now()-t;
   results[i]={...assess(c,spec,lay),inferenceCalls:calls,inferenceMs:opts.live?Math.round(inferenceMs):0,layoutMs:+layoutMs.toFixed(3),usage,model};
   scenes[i]={id:c.id,design:c.design,scene:{width:960,height:540,prims:lay.prims}};
   fs.writeFileSync(path.join(out,c.id+'.json'),JSON.stringify({raw,spec,layout:lay,result:results[i]},null,2));
   console.log(c.id,results[i].passed?'PASS':'FAIL',Object.entries(results[i].gates).filter(([,v])=>!v).map(([k])=>k).join(','));
 }catch(e){results[i]={id:c.id,split:c.split,passed:false,error:e.message,inferenceCalls:calls};console.log(c.id,'ERROR',e.message);}}}
 await Promise.all(Array.from({length:opts.live?2:1},worker));
 const latency=results.filter(x=>x.inferenceMs!=null).map(x=>x.inferenceMs).sort((a,b)=>a-b),quant=q=>latency[Math.max(0,Math.ceil(latency.length*q)-1)]??null;
 const summary={mode:opts.live?'single-inference':'oracle-structure',datasetVersion:dataset.version,model:results.find(x=>x.model)?.model||null,total:results.length,passed:results.filter(x=>x.passed).length,bySplit:Object.fromEntries(['dev','holdout'].map(s=>[s,{total:results.filter(r=>r.split===s).length,passed:results.filter(r=>r.split===s&&r.passed).length}])),p50Ms:quant(.5),p95Ms:quant(.95),notes:['Gates are machine checks, not a visual quality score.','Lexical content checks do not prove semantic equivalence. Numeric row/series binding is checked separately.','One successful generation per case; no density/coverage repair. Transport compatibility retries may occur.'],results};
 fs.writeFileSync(path.join(out,'report.json'),JSON.stringify(summary,null,2));fs.writeFileSync(path.join(out,'scenes.json'),JSON.stringify(scenes.filter(Boolean),null,2));console.log(JSON.stringify({...summary,results:undefined},null,2));return summary;
}
if(require.main===module){const a=process.argv.slice(2),v=k=>a[a.indexOf(k)+1];run({live:a.includes('--live'),out:a.includes('--out')?v('--out'):undefined,split:a.includes('--split')?v('--split'):undefined,layout:a.includes('--layout')?v('--layout'):undefined,prompt:a.includes('--prompt')?v('--prompt'):undefined}).catch(e=>{console.error(e);process.exitCode=1;});}
module.exports={assess,run};
