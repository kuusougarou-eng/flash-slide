"use strict";
const assert=require('assert'),G=require('../server/golden-native/geometry');
for(const values of [[36,9],[9,36],[0,4],[0,0],[.01,100000]]){
 const r=G.areas(values,[{x:100,y:100},{x:300,y:100}],130);
 assert(r.every(x=>x.w===x.h&&x.w<=130));
 for(let i=0;i<2;i++)assert.equal(r[i].visible,values[i]>0);
 if(values.every(v=>v>0))assert(Math.abs((r[0].w/r[1].w)**2-values[0]/values[1])<1e-7*Math.max(1,values[0]/values[1]));
}
const bs=G.bars([-20,0,30],{x:100,w:500},-50,50);assert.deepEqual(bs.map(b=>b.w),[100,0,150]);assert.equal(bs[0].x,250);assert.equal(bs[2].end,500);
assert.throws(()=>G.areas([-1,2],[{},{}],130),/negative/);
assert.throws(()=>G.bars([101],{x:0,w:100},0,100),/outside/);
for(const n of [4,8,15,24]){const r=G.intervals([{start:0,end:0},{start:0,end:n-1},{start:n-1,end:n-1}],{x:202,w:630},n);assert(Math.abs(r[1].w-626)<1e-7);assert(Math.abs(r[2].x+r[2].w-830)<1e-7);}
assert.throws(()=>G.intervals([{start:2,end:1}],{x:0,w:100},3),/valid inclusive/);
for(const n of [3,6,11,16]){const r=G.rows({x:32,y:164,w:896,h:280},n);assert(Math.abs(r.at(-1).y+r.at(-1).h-444)<1e-7);}
assert.throws(()=>G.rows({x:0,y:0,w:100,h:20},3),/available/);
assert.deepEqual(G.point({x:5,y:10},{x:100,y:100,w:300,h:200},[0,10],[0,20]),{x:250,y:200});
console.log('golden geometry: area ratio, signed bars, inclusive intervals, variable rows, domains verified');
