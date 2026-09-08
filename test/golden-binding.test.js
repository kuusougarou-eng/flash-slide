"use strict";
const assert=require('assert'),AdmZip=require('adm-zip');
const {GoldenCompiler,parse,read,els}=require('../server/golden-native/package');
const g=new GoldenCompiler();
function fill(id,geometry){const s=g.get(id);return new AdmZip(g.fill(id,{experimental:true,texts:Object.fromEntries(s.slots.map(x=>[x.id,x.text])),geometry}));}
function shape(z,id,shapeId){const d=parse(read(z,g.get(id).part));const s=[...els(d,'p','sp'),...els(d,'p','cxnSp')].find(s=>els(s,'p','cNvPr')[0]?.getAttribute('id')===shapeId);return{node:s,x:Number(els(s,'a','off')[0].getAttribute('x'))/12700,w:Number(els(s,'a','ext')[0].getAttribute('cx'))/12700,text:els(s,'a','t').map(t=>t.textContent).join('')};}
for(const values of [[36,9],[9,36],[0,4]]){
 const z=fill('30',{values}),a=shape(z,'30','7'),b=shape(z,'30','11');
 if(values[0])assert(Math.abs((a.w/b.w)**2-values[0]/values[1])<1e-5);
 else assert.equal(els(a.node,'p','cNvPr')[0].getAttribute('hidden'),'1');
 assert.equal(shape(z,'30','9').text,values[0].toFixed(1));assert.equal(shape(z,'30','13').text,values[1].toFixed(1));
}
const z=fill('31',{values:[90,80,70,60,50,40],target:85});
assert.equal(shape(z,'31','8').w,324);assert.equal(shape(z,'31','9').x,426);assert.equal(shape(z,'31','9').text,'90%');
assert.equal(shape(z,'31','29').x,402);assert.equal(shape(z,'31','30').text,'目標 85%');
assert.throws(()=>fill('31',{values:[1,2,3,4,5,6]}),/finite/);
assert.throws(()=>fill('31',{values:[101,2,3,4,5,6],target:80}),/outside/);
console.log('golden bindings: circle geometry/labels and bar geometry/labels/target stay synchronized');
