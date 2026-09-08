"use strict";
const AdmZip=require('adm-zip');
const C='http://schemas.openxmlformats.org/drawingml/2006/chart';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const {DOMParser,XMLSerializer}=require('@xmldom/xmldom');
const parse=s=>new DOMParser().parseFromString(String(s),'application/xml');
const xml=d=>new XMLSerializer().serializeToString(d);
const all=(n,t)=>Array.from(n.getElementsByTagNameNS(C,t));
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const column=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
function workbook(data){
 const z=new AdmZip();const files={
  '[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
  '_rels/.rels':'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+R+'/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  'xl/workbook.xml':'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="'+R+'"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels':'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+R+'/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
 };
 const rows=[['',...data.series.map(s=>s.name)],...data.categories.map((c,i)=>[c,...data.series.map(s=>s.values[i])])];
 files['xl/worksheets/sheet1.xml']='<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'+rows.map((row,i)=>'<row r="'+(i+1)+'">'+row.map((v,j)=>typeof v==='number'?`<c r="${column(j)}${i+1}"><v>${v}</v></c>`:`<c r="${column(j)}${i+1}" t="inlineStr"><is><t>${esc(v)}</t></is></c>`).join('')+'</row>').join('')+'</sheetData></worksheet>';
 for(const [p,s]of Object.entries(files))z.addFile(p,Buffer.from(s));return z.toBuffer();
}
function validate(data){
 if(!data||!Array.isArray(data.categories)||data.categories.length<1||data.categories.length>48)throw Error('Chart requires 1–48 categories');
 if(!data.categories.every(x=>typeof x==='string'))throw Error('Chart categories must be strings');
 if(!Array.isArray(data.series)||!data.series.length||data.series.length>8)throw Error('Chart requires 1–8 series');
 for(const s of data.series)if(typeof s.name!=='string'||!Array.isArray(s.values)||s.values.length!==data.categories.length||!s.values.every(Number.isFinite))throw Error('Chart series must bind one finite value to every category');
 if(data.domain&&(!Number.isFinite(data.domain.min)||!Number.isFinite(data.domain.max)||data.domain.min>=data.domain.max))throw Error('Chart domain must be an increasing range');
}
function bindChart(xmlText,data){
 validate(data);const d=parse(xmlText);const old=all(d,'ser');if(!old.length)throw Error('Chart has no series');
 const parent=old[0].parentNode;if(old.some(s=>s.parentNode!==parent))throw Error('Mixed chart series require an explicit combination mapping');
 const ref=(name,values,formula,numeric)=>{const kind=numeric?'num':'str';return parse(`<c:${name} xmlns:c="${C}"><c:${kind}Ref><c:f>${esc(formula)}</c:f><c:${kind}Cache>${numeric?'<c:formatCode>General</c:formatCode>':''}<c:ptCount val="${values.length}"/>${values.map((v,i)=>`<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:${kind}Cache></c:${kind}Ref></c:${name}>`).documentElement;};
 data.series.forEach((s,i)=>{
  const clone=old[Math.min(i,old.length-1)].cloneNode(true);
  all(clone,'idx')[0]?.setAttribute('val',String(i));all(clone,'order')[0]?.setAttribute('val',String(i));
  const letter=column(i+1);for(const [name,values,formula,numeric] of [['tx',[s.name],`Sheet1!$${letter}$1`,false],['cat',data.categories,`Sheet1!$A$2:$A$${data.categories.length+1}`,false],['val',s.values,`Sheet1!$${letter}$2:$${letter}$${data.categories.length+1}`,true]]){
   const original=all(clone,name)[0];if(!original)throw Error('Unsupported native chart series: '+name);
   original.parentNode.replaceChild(d.importNode(ref(name,values,formula,numeric),true),original);
  }
  if(s.labelPoints){const labels=all(clone,'dLbl');const last=labels.at(-1);s.labelPoints.forEach((idx,k)=>{let label=labels[k];if(!label){if(!last)throw Error('Chart series has no per-point label to reuse');label=last.cloneNode(true);last.parentNode.insertBefore(label,last.nextSibling);}all(label,'idx')[0].setAttribute('val',String(idx));});labels.slice(s.labelPoints.length).forEach(n=>n.parentNode.removeChild(n));}
  for(const label of all(clone,'dLbl'))if(Number(all(label,'idx')[0]?.getAttribute('val'))>=data.categories.length)label.parentNode.removeChild(label);
  if(s.points){for(const n of all(clone,'dPt'))n.parentNode.removeChild(n);const anchor=['dLbls','trendline','errBars','cat','val'].map(t=>all(clone,t)[0]).find(Boolean);for(const pt of s.points){const n=parse(`<c:dPt xmlns:c="${C}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:idx val="${pt.idx}"/><c:spPr><a:solidFill><a:srgbClr val="${pt.fill}"/></a:solidFill></c:spPr></c:dPt>`).documentElement;clone.insertBefore(d.importNode(n,true),anchor);}}
  parent.insertBefore(clone,old[0]);
 });old.forEach(s=>s.parentNode.removeChild(s));
 const pyFloat=v=>Number.isInteger(v)?v+'.0':String(v);
 for(const ax of all(d,'valAx')){for(const s of all(ax,'scaling')){for(const n of [...all(s,'min'),...all(s,'max')])n.parentNode.removeChild(n);if(data.domain)for(const key of ['max','min']){const n=d.createElementNS(C,'c:'+key);n.setAttribute('val',pyFloat(data.domain[key]));s.appendChild(n);}}
  const unit=data.domain?.majorUnit;const existing=all(ax,'majorUnit');if(unit!=null){if(existing.length)existing[0].setAttribute('val',pyFloat(unit));else{const n=d.createElementNS(C,'c:majorUnit');n.setAttribute('val',pyFloat(unit));ax.appendChild(n);}}else for(const n of existing)n.parentNode.removeChild(n);}
 return {xml:xml(d),workbook:workbook(data)};
}
module.exports={bindChart,validate,workbook};
