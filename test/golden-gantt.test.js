"use strict";
const assert=require('assert'),fs=require('fs'),AdmZip=require('adm-zip');
const P=require('../server/golden-native/package'),{layout}=require('../server/golden-native/gantt'),adapter=require('../server/golden-native/generate');
const b=require('../server/golden-native/bindings.json').slides['13'];
const data={title:'工程数と期間数を変えても対応関係を保つ',lead:'ガントの計算を検証する',note:'',source:'架空の検証用データ',periods:Array.from({length:8},(_,i)=>({label:(i+1)+'月',group:'2029年'})),tasks:Array.from({length:6},(_,i)=>({label:'工程'+i,owner:'担当'+i,start:i,end:i+1,critical:i===0})),milestones:[{label:'判定A',at:5},{label:'判定B',at:6},{label:'判定C',at:7}],today:null};
assert.equal(layout(data,b).msRows,2);
for(const n of [1,6,11,16]){const d={...data,tasks:Array.from({length:n},(_,i)=>({label:'工程'+i,owner:'担当',start:0,end:7,critical:false})),milestones:[]};const r=layout(d,b);assert.equal(r.rows.length,n);assert(Math.abs(r.rows.at(-1).y+r.rows.at(-1).h-468)<1e-6);}
assert.throws(()=>layout({...data,tasks:[{...data.tasks[0],label:'説明'.repeat(1000)}]},b),/capacity/);
assert.throws(()=>layout({...data,milestones:Array.from({length:4},(_,i)=>({label:'判定'+i,at:3}))},b),/three readable tracks/);
assert.throws(()=>layout({...data,tasks:[{...data.tasks[0],end:8}]},b),/inclusive/);
const result=adapter.compile(data),z=new AdmZip(result.buffer),d=P.parse(P.read(z,'ppt/slides/slide13.xml'));
const ids=P.els(d,'p','cNvPr').map(n=>n.getAttribute('id'));assert.equal(new Set(ids).size,ids.length);
const all=P.els(d,'a','t').map(t=>t.textContent).join(' ');
for(const t of data.tasks){assert(all.includes(t.label));assert(all.includes(t.owner));}
assert(!/ISO|SKU|SCM|2026|9\/7/.test(all),'no old task/period/today facts may survive');
const sizes=P.els(d,'a','rPr').map(n=>n.getAttribute('sz')).filter(Boolean);assert(sizes.includes('850'),'empty-then-refilled slots retain reference typography');
let calls=0;(async()=>{const r=await adapter.generate('検証入力',{chat:async()=>{calls++;return{content:JSON.stringify(data),model:'fake',usage:{}};}});assert.equal(calls,1);assert.equal(r.inferenceCalls,1);assert(r.buffer.length>1000);console.log('golden gantt: variable tracks/rows, stale-data removal, typography, one inference verified');})().catch(e=>{console.error(e);process.exitCode=1});
