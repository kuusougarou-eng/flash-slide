"use strict";
const fs=require('fs'),AdmZip=require('adm-zip'),P=require('../server/golden-native/package'),b=require('../server/golden-native/bindings.json').slides['13'];
const g=new P.GoldenCompiler(),slide=g.get('13'),zip=new AdmZip(g.reference('13')),d=P.parse(P.read(zip,slide.part));
const ns=new Map([...P.els(d,'p','sp'),...P.els(d,'p','cxnSp')].map(s=>[P.els(s,'p','cNvPr')[0].getAttribute('id'),s]));
const text=id=>P.els(ns.get(id),'a','t').map(t=>t.textContent).join('');
const rect=id=>{const s=ns.get(id),o=P.els(s,'a','off')[0],e=P.els(s,'a','ext')[0];return{x:Number(o.getAttribute('x'))/12700,y:Number(o.getAttribute('y'))/12700,w:Number(e.getAttribute('cx'))/12700,h:Number(e.getAttribute('cy'))/12700};};
const pw=b.plot.w/b.periodIds.length;
const reference={periods:b.periodIds.map(id=>({label:text(id),group:text(b.groupIds.find(g=>{const r=rect(g),x=rect(id).x+pw/2;return r.x<=x&&x<=r.x+r.w;}))})),tasks:[],milestones:[],today:{at:(rect(b.todayLine).x-b.plot.x)/pw,label:text(b.todayLabel)}};
for(let i=0;i<b.taskIds.length;i+=4){const [label,bar,owner]=b.taskIds.slice(i,i+4),r=rect(bar),start=Math.round((r.x-b.plot.x-2)/pw);reference.tasks.push({label:text(label),owner:text(owner),start,end:start+Math.round((r.w+4)/pw)-1,critical:P.els(ns.get(bar),'a','srgbClr')[0].getAttribute('val')===b.accent});}
for(let i=0;i<b.milestoneIds.length;i+=2){const [shape,label]=b.milestoneIds.slice(i,i+2),r=rect(shape);reference.milestones.push({label:text(label),at:Math.floor((r.x+r.w/2-b.plot.x)/pw)});}
fs.writeFileSync('debug/golden-native/gantt-reference-data.json',JSON.stringify(reference,null,2));
fs.writeFileSync('debug/golden-native/bound-13.pptx',g.fill('13',{experimental:true,texts:Object.fromEntries(slide.slots.map(s=>[s.id,s.text])),geometry:reference}));
const changed={periods:Array.from({length:8},(_,i)=>({label:(i+4)+'月',group:'2028 年度'})),tasks:[{label:'要件合意',owner:'企画部',start:0,end:1,critical:true},{label:'設計',owner:'設計部',start:1,end:3,critical:false},{label:'実装',owner:'開発部',start:2,end:5,critical:true},{label:'検証',owner:'品質部',start:4,end:6,critical:false},{label:'移行',owner:'運用部',start:6,end:6,critical:true},{label:'効果確認',owner:'企画部',start:7,end:7,critical:false}],milestones:[{label:'要件合意',at:1},{label:'受入判定',at:5},{label:'移行完了',at:6},{label:'効果確認',at:7}],today:null};
const texts=Object.fromEntries(slide.slots.map(s=>[s.id,s.role==='title'?'基幹システム刷新を8か月で進める':s.role==='lead'?'6工程で進め、要件合意・受入判定・移行完了・効果確認を節目にする':s.role==='source'?'架空の検証用データ':s.role==='page'?'1':'']));
fs.writeFileSync('debug/golden-native/gantt-changed-data.json',JSON.stringify(changed,null,2));
fs.writeFileSync('debug/golden-native/changed-gantt.pptx',g.fill('13',{experimental:true,texts,geometry:changed}));
console.log('Built reference and changed-count gantts');
