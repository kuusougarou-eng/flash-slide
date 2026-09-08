"use strict";
// Data -> geometry. No slide IDs, company names or golden values in the solver.
function finite(v,name){if(!Number.isFinite(v))throw Error(name+' must be finite');return v;}
function domain(lo,hi){finite(lo,'minimum');finite(hi,'maximum');if(!(lo<hi))throw Error('Domain must be increasing');return v=>{finite(v,'value');if(v<lo||v>hi)throw Error('Value outside domain');return(v-lo)/(hi-lo);};}
function areas(values,centers,maxDiameter){
 if(!values.length||values.length!==centers.length)throw Error('Area values and centers must correspond');finite(maxDiameter,'diameter');if(maxDiameter<=0)throw Error('Diameter must be positive');
 values.forEach(v=>{finite(v,'area value');if(v<0)throw Error('Area cannot be negative');});const max=Math.max(...values);
 return values.map((v,i)=>{const d=max?maxDiameter*Math.sqrt(v/max):0;return{x:finite(centers[i].x,'center x')-d/2,y:finite(centers[i].y,'center y')-d/2,w:d,h:d,visible:v>0};});
}
function bars(values,plot,lo=0,hi=Math.max(...values,1)){
 const scale=domain(lo,hi);if(lo>0||hi<0)throw Error('Bar domain must contain zero');const zero=plot.x+scale(0)*plot.w;
 return values.map(v=>{const end=plot.x+scale(v)*plot.w;return{x:Math.min(zero,end),w:Math.abs(end-zero),end,zero,visible:v!==0};});
}
function intervals(tasks,plot,periods,padding=2){
 if(!Number.isInteger(periods)||periods<1||periods>120)throw Error('Invalid period count');
 return tasks.map(t=>{if(!Number.isInteger(t.start)||!Number.isInteger(t.end)||t.start<0||t.end<t.start||t.end>=periods)throw Error('Task must have a valid inclusive interval');const w=(t.end-t.start+1)*plot.w/periods;return{x:plot.x+t.start*plot.w/periods+Math.min(padding,w/4),w:w-2*Math.min(padding,w/4)};});
}
function rows(rect,count,gap=0,minHeight=12){
 if(!Number.isInteger(count)||count<1||count>100)throw Error('Invalid row count');const h=(rect.h-gap*(count-1))/count;if(h<minHeight)throw Error('Rows exceed available height');return Array.from({length:count},(_,i)=>({x:rect.x,y:rect.y+i*(h+gap),w:rect.w,h}));
}
function point(value,plot,xd,yd){return{x:plot.x+domain(...xd)(value.x)*plot.w,y:plot.y+(1-domain(...yd)(value.y))*plot.h};}
module.exports={areas,bars,intervals,rows,point,domain};
