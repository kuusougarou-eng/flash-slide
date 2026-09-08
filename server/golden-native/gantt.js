"use strict";
const G=require('./geometry'),{scene}=require('./xml-scene');
function validate(data){
 if(!Array.isArray(data?.periods)||data.periods.length<2||data.periods.length>24||!data.periods.every(p=>typeof p.label==='string'&&p.label&&typeof p.group==='string'&&p.group))throw Error('Gantt requires 2–24 labeled and grouped periods');
 if(!Array.isArray(data.tasks)||!data.tasks.length||data.tasks.length>18||!data.tasks.every(t=>typeof t.label==='string'&&t.label&&typeof t.owner==='string'&&typeof t.critical==='boolean'))throw Error('Gantt requires 1–18 tasks with label, owner and critical flag');
 G.intervals(data.tasks,{x:0,w:100},data.periods.length);
 if(!Array.isArray(data.milestones)||data.milestones.length>12||!data.milestones.every(m=>typeof m.label==='string'&&m.label&&Number.isInteger(m.at)&&m.at>=0&&m.at<data.periods.length))throw Error('Invalid milestones');
 if(data.today!==null&&(!data.today||typeof data.today.label!=='string'||!Number.isFinite(data.today.at)||data.today.at<0||data.today.at>data.periods.length))throw Error('Today must be null or a labeled position');
}
function layout(data,b){
 validate(data);const pw=b.plot.w/data.periods.length;
 const tracks=[];const milestones=data.milestones.map((m,index)=>{const cx=b.plot.x+(m.at+.5)*pw;const w=80;const x=Math.min(b.right-w,cx+6);return{...m,index,cx,x,w};}).sort((a,b)=>a.x-b.x);
 for(const m of milestones){let track=tracks.findIndex(end=>end+6<=m.x);if(track<0)track=tracks.length;if(track>=3)throw Error('Milestones require more than three readable tracks');tracks[track]=m.x+m.w;m.track=track;}
 const msRows=tracks.length,extra=(msRows-1)*b.milestones.h;
 const msY=b.milestones.y-extra;
 const rs=G.rows({...b.rows,h:b.rows.h-extra},data.tasks.length,0,16);
 // Validate readable label/owner capacity before touching the document.
 const fit=(text,w,h,size)=>{const n=require('./package').weightedLength(text);const lines=Math.max(1,Math.ceil(n*size/(w-8)));if(lines*size*1.15>h+1)throw Error('Gantt text exceeds row capacity: '+text);};
 data.tasks.forEach((t,i)=>{fit(t.label,b.label.w,rs[i].h,10);fit(t.owner,b.owner.w,rs[i].h,8.5);});
 const spans=G.intervals(data.tasks,b.plot,data.periods.length);
 const groups=[];data.periods.forEach((p,i)=>{const last=groups.at(-1);if(last&&last.label===p.group)last.count++;else groups.push({label:p.group,start:i,count:1});});
 return{pw,groups,rows:rs,spans,milestones:milestones.sort((a,b)=>a.index-b.index),msRows,msY};
}
function bindGantt(doc,data,b){
 const l=layout(data,b),s=scene(doc);
 s.repeat(b.groupIds,l.groups,(g,i,clone)=>[s.text(s.rect(clone(b.groupIds[0]),{x:b.plot.x+g.start*l.pw,w:g.count*l.pw-2}),g.label)]);
 s.repeat(b.periodIds,data.periods,(p,i,clone)=>[s.text(s.rect(clone(b.periodIds[0]),{x:b.plot.x+i*l.pw,w:l.pw}),p.label)]);
 s.repeat(b.taskIds,data.tasks,(t,i,clone)=>{
  const row=l.rows[i],span=l.spans[i],h=Math.min(18,row.h-10);
  return[s.text(s.rect(clone(b.taskTemplate.label),{y:row.y,h:row.h}),t.label),s.fill(s.rect(clone(b.taskTemplate.bar),{x:span.x,y:row.y+(row.h-h)/2,w:span.w,h}),t.critical?b.accent:b.neutral),s.text(s.rect(clone(b.taskTemplate.owner),{y:row.y,h:row.h}),t.owner),s.rect(clone(b.taskTemplate.rule),{y:row.y+row.h})];
 });
 s.repeat(b.milestoneIds,l.milestones,(m,i,clone)=>{const y=l.msY+m.track*b.milestones.h;return[s.rect(clone(b.milestoneTemplate.shape),{x:m.cx-5,y:y+7}),s.text(s.rect(clone(b.milestoneTemplate.label),{x:m.x,y,w:m.w,h:b.milestones.h}),m.label)];});
 if(l.msRows)s.text(s.rect(s.get(b.milestoneHeading),{y:l.msY,h:l.msRows*b.milestones.h}),'節目');else s.remove(b.milestoneHeading);
 s.repeat(b.gridIds,Array.from({length:data.periods.length-1},(_,i)=>i+1),(i,_,clone)=>[s.rect(clone(b.gridIds[0]),{x:b.plot.x+i*l.pw})]);
 if(data.today){const x=b.plot.x+data.today.at*l.pw;s.rect(s.get(b.todayLine),{x});s.text(s.rect(s.get(b.todayLabel),{x:Math.max(b.plot.x,Math.min(b.plot.x+b.plot.w-44,x-22))}),data.today.label);}else {s.remove(b.todayLine);s.remove(b.todayLabel);}
 s.text(s.get(b.ownerHeading),'担当');
 return l;
}
module.exports={layout,bindGantt};
