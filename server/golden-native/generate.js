"use strict";
// First fully semantic adapter. The model returns schedule data, never slot IDs,
// coordinates, JavaScript, shape IDs or source-company text.
const {GoldenCompiler}=require('./package');
let compiler;
const getCompiler=()=>compiler||(compiler=new GoldenCompiler());
function messages(input){return[{role:'system',content:`工程計画の素材を一枚のコンサル資料に構造化する。入力の事実・数値・条件を残し、入力にない工程・日付・担当・効果を作らない。座標・色・図形IDは出力しない。
JSON形式: {"title":"結論の常体文、70字以内","lead":"タイトルを繰り返さない根拠、100字以内","note":"元の素材にある条件や注意点、なければ空文字","source":"入力にある出典のみ、なければ空文字","periods":[{"label":"4月","group":"2028年度"}],"tasks":[{"label":"要件整理","owner":"担当名、未指定なら空文字","start":0,"end":1,"critical":false}],"milestones":[{"label":"要件合意","at":1}],"today":null}
periodsは時間順で抜けなく2〜24個。上位ラベルgroupと下位labelで年度/月、月/週などを区別する。tasksは1〜18件、start/endはperiodsの0始まり添字、endは終了期間を含む。criticalは入力で重要工程と明示された場合だけtrue。milestonesは入力で節目と明示されたものだけ、なければ空配列。todayは日付と位置が素材から明らかな場合だけ{"at":期間の0始まり位置,"label":"今日の日付"}、それ以外はnull。期間が判断できない入力には{"error":"不足している情報"}を返す。noteは既定で空。tasksやmilestonesに載せた担当・期間・節目を繰り返してはいけない。独立した条件だけをnoteに保持する。JSONのみ返す。`},{role:'user',content:String(input)}];}
function stripRepeatedNote(plan){
 const key=s=>String(s).normalize('NFKC').replace(/\s/g,'');
 const known=new Set();
 for(const t of plan.tasks||[]){const a=plan.periods?.[t.start]?.label,b=plan.periods?.[t.end]?.label;if(!a||!b)continue;for(const span of t.start===t.end?[a]:[a+'〜'+b,a.replace(/[月週]$/,'')+'〜'+b])known.add(key(t.label+'は'+span+'で'+t.owner));}
 return plan.note.split(/[、。]/).filter(part=>!known.has(key(part))).filter(Boolean).join('、');
}
function compile(plan){
 if(plan?.error)throw Error(plan.error);
 if(!plan||!['title','lead','note','source'].every(k=>typeof plan[k]==='string')||!plan.title)throw Error('Title, lead, note and source must be strings');
 const g=getCompiler(),slide=g.get('13');
 const texts=Object.fromEntries(slide.slots.map(s=>[s.id,s.role==='title'?plan.title:s.role==='lead'?plan.lead:s.role==='source'?plan.source:s.role==='page'?'1':'']));
 // Footnote/legend slot is defined by the checked-in reference, not by the model.
 const noteSlot=slide.slots.find(s=>s.rect.y===474);if(!noteSlot)throw Error('Missing calibrated note slot');texts[noteSlot.id]=plan.note;
 const geometry={periods:plan.periods,tasks:plan.tasks,milestones:plan.milestones,today:plan.today};
 const started=performance.now();const buffer=g.fill('13',{texts,geometry,experimental:true});
 return{buffer,compileMs:+(performance.now()-started).toFixed(3),layout:'13',sourceSha256:g.catalog.sourceSha256};
}
async function generate(input,{model,chat}={}){
 if(typeof input!=='string'||!input.trim())throw Error('Input is empty');
 const started=performance.now();const response=await(chat||require('../llm').chat)({messages:messages(input),model,maxTokens:2600,temperature:.2,jsonMode:true});
 const inferenceMs=performance.now()-started;const plan=require('../prompt').extractJson(response.content);
 plan.note=typeof plan.note==='string'?stripRepeatedNote(plan):plan.note;
 const compiled=compile(plan);
 return{...compiled,plan,inferenceCalls:1,inferenceMs:Math.round(inferenceMs),model:response.model,usage:response.usage};
}
module.exports={messages,compile,generate,stripRepeatedNote};
