import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Workbook, SpreadsheetFile, Presentation, PresentationFile } from '@oai/artifact-tool';
process.env.RUNTIME_NODE_MODULES='/Users/zhenshuai/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';

const ROOT='/Users/zhenshuai/Desktop/小禾学伴/市场调研-20261029';
const OUT=path.join(ROOT,'交付');
const BUILD=path.join(ROOT,'.build');
const PRE=path.join(BUILD,'previews');
const FONT='Heiti SC';
const C={ink:'#173B35',green:'#267060',muted:'#61736D',light:'#EAF2EC',paper:'#FAFAF6',amber:'#A76725',white:'#FFFFFF',line:'#D4E1D8'};
const read=async n=>JSON.parse(await fs.readFile(path.join(ROOT,'research',n),'utf8'));
const apps=(await read('apps.json')).products;
const hwData=await read('hardware.json');
const hard=hwData.products || hwData;
const reviews=await read('reviews.json');
const protocol=await read('local_and_protocol.json');
const products=[...apps,...hard];
const shortNames=['作业帮 App','小猿 AI','洋葱学园','讯飞学习机','学而思学习机','作业帮学习机','小猿学练机','作业帮学习桌'];
const pricingText=p=>typeof p.pricing==='string'?p.pricing: [p.pricing?.download, ...(p.pricing?.examples||[]).map(e=>`${e.item}：¥${e.price_cny}`),p.pricing?.paid_rights,p.pricing?.caveat,p.pricing?.hardware,p.pricing?.conditions,p.pricing?.content_membership].filter(Boolean).join('\n');
const versionText=p=>[p.version||p.model||'型号未确认',p.version_date].filter(Boolean).join(' / ');
const sources=products.flatMap(p=>(p.sources||[]).map(s=>({...s,product:p.name})));
const reviewRows=reviews.products.flatMap(p=>p.reviews);
const improvements=[
 ['P0','补足可信内容与知识点映射','30 道样题难以支撑持续复习。扩充包仍待教研确认。','APP-实施计划.md；content-trial-notes.md','选定四年级数学一个单元，完成教材映射、答案复核和不同结构题族。','教研审核通过；每个复查任务能找到未曝光题。','教研人员与内容授权，投入较高','待实施'],
 ['P1','做好提示、独立练习与延迟复查','已有本地规则，尚无实际儿童使用和掌握判定验证。','APP-实施计划.md；竞品官方已有学练测与核验主张','先尝试独立作答，再逐级提示；明确展示帮助状态；安排次日与后续未见题复查。','记录无帮助完成、提示后完成和次日表现；保留原答，规则待教学验证。','交互调整与教研验证，投入中等','待验证'],
 ['P1','给家长可解释的下一步','App 家长协作和跨设备服务尚未完成，现阶段报告需验证可理解性。','APP-实施计划.md；小猿官方角色分工；公开评论','先做单机可分享摘要：已独立完成、仍需帮助、下次复查内容。说明覆盖和付费边界。','邀请家长说明孩子困难和下一步；记录理解错误，不把满意度当学习效果。','报告交互与家长测试，投入中等','待验证'],
 ['探索','受题库约束的 AI 追问','当前学伴引用固定提示，不能完成真实连续语义追问。','APP-实施计划.md；3 个 App 官方介绍','先接入一个单元的受控答疑，设计错误答案、更换问法和不会回答的样例。','核验解释正确性、是否提前给答案、出错后能否纠正。','模型接入、评测与内容治理，投入较高','待研究'],
 ['探索','小规模教师协作试点','正式教师绑定、送达、回复和原生授权未完成。','APP-实施计划.md；小猿班级功能；洋葱教师端','找到试点教师后，验证布置练习、查看问题、回复求助的最小流程。','教师能完成任务；确认持续使用意愿和授权要求。','身份与云端权限、协作流程，投入较高','待研究']
];

const matrix=[
 ['作业帮 App','拍题／批改','AI解析与追问','同步练／单元测','归档细节待核','独立判据未披露','家长工具定位'],
 ['小猿 AI','拍题／批改','视频／AI讲解','同步练／举一反三','分类／错因','独立判据未披露','班级批改统计'],
 ['洋葱学园','拍题精学','动画课／AI私教','互动练／学习计划','跨日规则待核','思考核验主张','教师端独立'],
 ['讯飞学习机','产品线宣称','互动答疑主张','精准学主张','详情见来源','具体规则待核','具体权限待核'],
 ['学而思学习机','拍题诊断主张','AI讲题主张','个性规划主张','错题巩固主张','具体判据待核','具体权益待核'],
 ['作业帮学习机','辅导主张','AI辅导主张','同步课程（产品线）','型号权益待核','具体规则待核','具体权限待核'],
 ['小猿学练机','答案管控*','型号详情待核','练习安排待核','型号详情待核','具体规则待核','家长管理*'],
 ['作业帮学习桌','纸面伴学','实时辅导主张','资源详情待核','未查到','未查到','家长减负主张']
];

function col(n){let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;}
const wb=Workbook.create();
const specs=[];
function sheet(name,title,headers,rows,widths){
 const s=wb.worksheets.add(name);s.showGridLines=false;s.tabColor=C.green;
 const end=col(headers.length-1),last=rows.length+6;
 s.getRange(`A1:${end}${last}`).format.font={name:FONT,size:11,color:C.ink};
 s.getRange('A2').values=[[title]];s.getRange('A2').format.font={name:FONT,size:16,bold:true,color:C.ink};
 s.getRange('A4').values=[['资料截至 2026-10-29。官方介绍与用户反馈均未替代 App 实测。']];
 s.getRange('A4').format.font={name:FONT,size:11,color:C.muted};
 s.getRange(`A6:${end}6`).values=[headers];
 s.getRange(`A6:${end}6`).format={fill:C.ink,font:{name:FONT,size:11,bold:true,color:C.white},rowHeight:34,wrapText:true,verticalAlignment:'center',horizontalAlignment:'center'};
 if(rows.length)s.getRange(`A7:${end}${last}`).values=rows;
 for(let j=0;j<headers.length;j++)s.getRange(`${col(j)}1:${col(j)}${last}`).format.columnWidthPx=widths[j]||180;
 for(let i=0;i<rows.length;i++){
  const r=s.getRange(`A${i+7}:${end}${i+7}`);r.format.wrapText=true;r.format.verticalAlignment='top';
  let lines=1;rows[i].forEach((v,j)=>{const t=String(v??'');let l=t.split('\n').reduce((a,x)=>a+Math.max(1,Math.ceil([...x].reduce((s,c)=>s+(c.charCodeAt(0)>255?1:.54),0)/Math.max(5,((widths[j]||180)-16)/15))),0);lines=Math.max(lines,l);});
  r.format.rowHeightPx=Math.max(64,lines*20+18);if(i%2===1)r.format.fill='#F1F5F1';
 }
 s.freezePanes.freezeRows(6);s.freezePanes.freezeColumns(2);
 const tab=s.tables.add(`A6:${end}${last}`,true,`ResearchTable${specs.length+1}`);tab.style='TableStyleLight1';tab.showFilterButton=true;
 s.getRange(`A6:${end}6`).format.fill=C.ink;s.getRange(`A6:${end}6`).format.font={name:FONT,size:11,bold:true,color:C.white};
 for(let i=0;i<rows.length;i++)s.getRange(`A${i+7}:${end}${i+7}`).format.fill=i%2?'#F1F5F1':'#FFFFFF';
 specs.push({name,last,cols:headers.length});return s;
}
const sm=sheet('功能对比','8 个样本的公开功能对比',['产品','作业输入','讲解方式','练习安排','错题与复习','掌握判定','家长／教师'],matrix,[180,180,180,195,195,195,195]);
sm.getRange('A4').values=[['全表为官方主张或未核实项。* 为配套 App 产品线能力，P60／S6 适配待核。型号、范围和来源见后续工作表。']];
const overview=products.map((p,i)=>[p.id||`P${i+1}`,p.name,p.form,versionText(p),p.audience,p.positioning||p.implications?.join('；')||'',pricingText(p)||(typeof p.pricing==='object'?JSON.stringify(p.pricing):'未核实'),(p.unknowns||[]).join('\n'),new Date('2026-10-29T00:00:00Z'),(p.sources||[]).slice(0,2).map(s=>s.url).join('\n')]);
const so=sheet('产品概况','产品身份、定位与收费口径',['编号','产品','形态','版本／型号','服务对象','主要主张／启发','收费与限制','待核实','检索日期','来源链接'],overview,[100,220,150,220,240,300,350,340,120,430]);
so.getRange(`I7:I${overview.length+6}`).setNumberFormat('yyyy-mm-dd');
let evid=[];
for(const p of products)for(const [j,f] of (p.features||[]).entries())evid.push([`${p.id||p.name}-F${j+1}`,p.name,f.evidence_type||'官方介绍','案头核验',f.dimension,f.claim,new Date('2026-10-29T00:00:00Z'),versionText(p),f.url||'',[f.scope,f.caveat||'未实测'].filter(Boolean).join('：'),'']);
for(const r of reviewRows)evid.push([`R-${r.review_id}`,r.product_name,'用户反馈','未复现',(r.tags||[]).join('／'),`${r.summary}（${r.stars}星；正文${r.sentiment_by_text}）`,new Date(r.date_china.replace(' ','T')+':00+08:00'),r.app_version,r.source_url+`\n评论ID：${r.review_id}`,'抽样线索，不能估计总体发生率','']);
const testTasks=[['T1','不会一道题','输入统一题，记录提示、答案出现时机、可否追问与收费'],['T2','做错一道题','上传典型错解，记录是否识别错误、能否订正和再练'],['T3','独立练习与复查','保存帮助状态，当日及次日用未见变式检查独立完成'],['T4','家长查看结果','记录报告入口、可解释性、下一步及分享权限']];
const firstTestRow=evid.length+7;
for(const p of apps)for(const t of testTasks)evid.push([`${p.id}-${t[0]}`,p.name,'尚未核实','待实测',t[1],t[2],null,'待填实际版本','','记录步骤、用时、付费拦截、答案正确性和截图路径','']);
const se=sheet('体验证据','官方能力、60 条评论与 12 项待实测任务',['证据编号','产品','证据类型','核验状态','维度／任务','发现或待验证问题','日期','版本／型号','来源或评论索引','局限／操作要求','个人实测记录'],evid,[200,200,120,120,210,420,135,180,440,350,400]);
se.getRange(`G7:G${evid.length+6}`).setNumberFormat('yyyy-mm-dd');
se.getRange(`D${firstTestRow}:D${evid.length+6}`).dataValidation={rule:{type:'list',values:['待实测','进行中','已实测','受阻']}};
se.getRange(`K${firstTestRow}:K${evid.length+6}`).format.fill='#FFF1D5';
se.getRange('A4').values=[['每 App 选入20条中国区公开评论。样本窗口不同、年龄与身份不明。实测行在表格末尾，黄色单元格供填写。']];
const si=sheet('改进机会','3 项优先改进与 2 项后续探索',['优先级','建议','问题与依据','证据定位','具体改法','验证方式','资源投入','状态'],improvements,[105,260,400,340,420,420,300,120]);
si.getRange('A4').values=[['建议优先验证家庭端四年级数学。该定位来自现有能力与投入判断，付费意愿和学习效果仍待验证。']];
si.getRange('H7:H11').dataValidation={rule:{type:'list',values:['待实施','待验证','待研究','进行中','完成']}};
wb.recalculate();
await fs.writeFile(path.join(BUILD,'workbook-inspect.txt'),(await wb.inspect({kind:'workbook,sheet,table',maxChars:5000,tableMaxRows:3,tableMaxCols:6})).ndjson);
const errors=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!|#SPILL!|#CALC!',options:{useRegex:true,maxResults:30},maxChars:2000});
await fs.writeFile(path.join(BUILD,'workbook-errors.txt'),errors.ndjson);
for(const s of specs){const blob=await wb.render({sheetName:s.name,range:`A1:${col(Math.min(5,s.cols-1))}${Math.min(s.last,10)}`,scale:1.3,format:'png'});await fs.writeFile(path.join(PRE,`sheet-${s.name}.png`),new Uint8Array(await blob.arrayBuffer()));}
await(await SpreadsheetFile.exportXlsx(wb)).save(path.join(OUT,'竞品对比表.xlsx'));

// Presentation uses native editable text and tables. Screenshots remain sourced images.
const pres=Presentation.create({slideSize:{width:1280,height:720}});
function text(s,txt,x,y,w,h,size=28,color=C.ink,bold=false){const z=s.shapes.add({geometry:'textbox',name:txt.slice(0,24),position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});z.text=txt;z.text.style={typeface:FONT,fontSize:size,color,bold,wrap:'square',autoFit:'none',insets:0,verticalAlignment:'top'};return z;}
function slide(title,note='',refs=[]){const s=pres.slides.add();s.background.fill=C.paper;const cover=title.includes('\n');text(s,title,64,cover?66:45,1152,cover?170:95,cover?56:44,C.ink,true);text(s,String(pres.slides.items.length).padStart(2,'0'),1170,668,46,24,18,C.muted);s.speakerNotes.text=[note,'资料截至2026-10-29。官方主张、用户反馈、项目文档核验与推断分别标识。',...refs].join('\n\n');return s;}
function table(s,rows,widths,{x=64,y=170,w=1152,h=400,size=24}={}){const t=s.tables.add({rows:rows.length,columns:rows[0].length,left:x,top:y,width:w,height:h,columnWidths:widths,values:rows});t.borders.assign({fill:C.line,width:.6,style:'solid'});for(let i=0;i<rows.length;i++){t.rows[i].height=h/rows.length;for(let j=0;j<rows[0].length;j++){const cell=t.getCell(i,j);cell.fill=i===0?C.ink:(i%2?C.white:C.light);cell.text.style={typeface:FONT,fontSize:size,color:i===0?C.white:C.ink,bold:i===0,autoFit:'none',wrap:'square',insets:{left:12,right:12,top:9,bottom:9},verticalAlignment:'middle'};}}return t;}
async function pic(s,file,x,y,w,h){const b=await fs.readFile(path.join(ROOT,'来源与截图',file));s.images.add({blob:b.buffer.slice(b.byteOffset,b.byteOffset+b.byteLength),contentType:file.endsWith('.png')?'image/png':'image/jpeg',alt:`官方页面或开发者展示：${file}`,position:{left:x,top:y,width:w,height:h},fit:'contain'});}
const appURLs=apps.flatMap(a=>a.sources.map(s=>`${s.title}：${s.url}`));
const hwURLs=hard.flatMap(a=>a.sources.map(s=>`${s.title}：${s.url}`));
let s=slide('小学智能学习产品调研\n与小禾学伴定位建议','开场约45秒。此次汇报希望确定目标用户、首个场景和下一轮验证投入。本版已经完成公开资料核验和评论整理，手机实测和真实用户验证尚未完成。');
text(s,'小学阶段 · 以 App 为主',64,280,1100,50,32,C.green);
text(s,'案头调研版',64,422,1100,60,36,C.ink,true);
text(s,'2026 年 10 月 29 日',64,510,1100,40,26,C.muted);
s=slide('建议优先验证家庭端的数学巩固场景','约90秒。推荐的是第一轮试点方向，不是已经证明的市场定位。先面向小学四年级，围绕一个数学单元把独立练习、逐级提示与后续复查做完整。支持依据是现有能力匹配、相对可控的内容范围和教师协作仍待补齐。竞品已有核验和精准学习概念，不能把概念本身作为独家优势。',appURLs.concat(['/Users/zhenshuai/Desktop/小禾学伴/APP-实施计划.md']));
text(s,'目标：帮助孩子看清“哪些能独立完成，哪些还需要练习”',64,174,1152,92,35,C.green,true);
table(s,[['判断依据','对小禾的含义'],['现有基础','练习、帮助记录、原答保存和间隔复查已经形成试用流程'],['进入范围','从四年级数学一个单元开始，便于补内容、做验证'],['竞争现实','竞品已有学练测与核验主张，需要比具体体验和可信度']],[245,907],{y:302,h:252,size:25});
text(s,'待验证：家长是否理解并愿意持续使用，孩子是否能独立完成，是否存在付费意愿。',64,595,1125,60,24,C.amber);
s=slide('样本与证据范围','约60秒。本次8个样本覆盖3个App、4类学习平板及1个纸面伴学设备。评论来自中国区App Store公开RSS，各选20条，样本期和人群不可比。硬件没有实体操作。本版中没有声称已完成竞品App实测。',appURLs.concat(hwURLs));
table(s,[['研究部分','已取得的材料','结论边界'],['8 个产品样本','官网、商店介绍、型号／版本和来源','代表性选样，不构成销量排名'],['60 条公开评论','3 个 App 各20条，保留评论ID','用户陈述未复现，不能计算满意度排名'],['小禾现状','2026-10-27项目文档与交接记录','文档核验，没有重跑本次设备验收'],['待补的验证','3个 App × 4任务，统一6题已备','手机实测、次日复查和真实用户访谈']],[240,490,422],{h:370,size:25});
text(s,'iOS版本：作业帮 14.55.0，小猿 AI 3.145.0，洋葱学园 7.118.0。',64,600,1130,48,24,C.muted);
s=slide('用户任务与需要验证的问题','约75秒。这里是依据产品场景和公开评论整理的需求假设，没有把访谈尚未发生的结论写成用户事实。正面评论肯定解析与课程，负面评论提示收费、准确性、功能使用等问题。年龄和年级不明，应在后续小学家庭样本中验证。\n个案索引：作业帮14717577839认可解题思路，14718060879反馈解析与收费；小猿14717720398对同题多个答案有疑问；洋葱14677388550肯定课程并反馈横竖屏与AI符号问题。均未复现。',appURLs.concat(reviews.products.map(p=>`评论样本：${p.product_name}；${p.method}；${p.reviews[0].source_url}`)));
table(s,[['角色','学习任务','当前线索','下一轮验证'],['学生','不会题时获得帮助\n学后独立再做','解析与课程获得肯定\nAI答案可靠性有投诉','孩子能否解释思路\n减少帮助后能否做对'],['家长','减少重复讲题\n看懂学习结果','会员边界有争议\n对辅导便利性有肯定','能否判断哪里不会\n是否愿意持续使用'],['教师','安排练习\n跟进共性错误','竞品有班级与教师端\n小禾协作仍待补齐','是否愿意参与试点\n接入成本是否可接受']],[135,305,350,362],{h:385,size:25});
text(s,'证据属性：公开反馈与研究假设，尚无本轮家长、学生或教师访谈。',64,605,1130,45,24,C.amber);
s=slide('8 个样本的产品主张','约90秒。按产品形态和用户任务比较。学习机的资源与设备条件无法直接等价为App能力。各硬件的具体型号和取证限制在Excel产品概况中，品牌级能力不外推为全型号权益。',appURLs.concat(hwURLs));
table(s,[['类别','样本','主要主张'],['作业辅导 App','作业帮、小猿 AI','拍题、批改、讲解与同步练习'],['知识学习 App','洋葱学园','动画课程、互动练习与 AI 学伴'],['综合学习机','讯飞、学而思、作业帮','课程、互动答疑与练习等产品线能力'],['学练设备','小猿学练机','配套 App 记录学习轨迹，S6适配待核'],['纸面伴学设备','作业帮 AI 学习桌','固定桌面环境中的伴学与辅导']],[245,345,562],{h:360,size:25});
text(s,'硬件样本按可核实官方页面选取，型号、版本和收费权益分别记录。',64,592,1140,60,24,C.muted);
s=slide('功能覆盖广，真正的体验仍需逐项核验','约90秒。这张表压缩展示核心能力。全部是官方资料或待核实项，不是性能打分。不把未查到解释为没有，也不把产品线宣传解释为所选机型的全部权益。',appURLs.concat(hwURLs));
table(s,[['产品','输入／讲解','练习与复习','掌握规则'],['作业帮 App','拍题、批改、AI追问','同步练，归档待核','独立判据未披露'],['小猿 AI','批改、视频与AI讲解','同步与变式练习','独立判据未披露'],['洋葱学园','拍题精学、动画与AI','互动练习与计划','思考核验主张'],['讯飞学习机','互动答疑（产品线）','精准练习（产品线）','具体规则待核'],['学而思学习机','诊断与AI讲题','规划与错题处理','具体规则待核'],['作业帮学习机','AI家教（商品页）','配套App错题同步*','具体规则待核'],['小猿学练机','答案查看管控*','学习轨迹*','具体规则待核'],['作业帮学习桌','指读／语音伴学','教材同步练／计时','未查到']],[220,320,340,272],{y:160,h:430,size:22});
text(s,'*为配套 App 产品线能力，P60／S6 适配待核。硬件能力需按型号和学段复核。',64,618,1120,47,23,C.amber);
s=slide('App 官方展示与待实测问题','约90秒。图片为开发者商店展示，不能当作我们操作过App的截图。把公开反馈转成明确操作问题：收费边界、独立作答判据和课程后的核验方式。本轮手机实测仍待补齐。',appURLs);
const files=await fs.readdir(path.join(ROOT,'来源与截图'));
for(let i=0;i<3;i++){const p=apps[i],fn=files.find(f=>f.startsWith(`官方展示-${p.id}-1.`));if(fn)await pic(s,fn,75+i*398,166,180,316);text(s,shortNames[i],275+i*398,184,178,55,28,C.green,true);const qs=['免费解析、视频与追问\n分别何时收费？','提示后答对，是否\n标记为独立掌握？','动画课结束后，如何\n检查理解和再练？'];text(s,qs[i],70+i*398,517,360,92,26,C.ink);}
text(s,'开发者宣传截图。用户评论中的问题未复现，12 项实测任务已列入对比表。',64,640,1110,40,22,C.amber);
s=slide('硬件提供的价值与 App 的承接范围','约75秒。这里讨论硬件形态带来的能力和使用条件，不判断护眼、准确率或提分效果。App可以承接拍照核对、打印练习和摘要反馈，固定摄像头、专用屏幕、桌椅等条件则属于设备。',hwURLs);
await pic(s,'05-学而思T6Max-官网.jpg',64,158,548,308);await pic(s,'08-作业帮AI学习桌-官网.jpg',660,158,548,308);
text(s,'学而思 T6 Max 官方页面',64,481,548,40,25,C.green,true);text(s,'作业帮 AI 学习桌官方页面',660,481,548,40,25,C.green,true);
text(s,'可借鉴：学习入口集中、学习过程连续、纸面作业连接。',64,550,1145,45,27,C.ink);
text(s,'小禾先承接拍照核对、练习与反馈；本轮不建议新增硬件研发。',64,611,1145,45,27,C.amber);
s=slide('小禾已有试用闭环，内容与真实使用仍是缺口','约90秒。仅依据当前项目文件。Android已做模拟器独立包验收，但没有多机型真机验证；iOS也未完成原生运行验收。不能把本地固定题库学伴宣传成生成式AI。现有五星和复查是试用规则，不能承诺测量出真实掌握度。',['/Users/zhenshuai/Desktop/小禾学伴/APP-实施计划.md','/Users/zhenshuai/Desktop/小禾学伴/xiaohe-app/content-trial-notes.md']);
table(s,[['已有本机试用','仍需补齐','对定位的影响'],['组卷、作答、提示记录\n原答与订正分开','足够题族、教材映射\n人工教研审核','先限定一个单元\n确保内容可持续'],['拍照保存、基础 OCR\n纸面打印与回传','复杂题型与手写识别\n真机稳定性','保留核对步骤\n避免把识别当评分'],['掌握规则、间隔复查\n固定题库提示','教学效果验证\n真实语义追问','先验证独立练习流程\n谨慎表达掌握结果'],['本地档案与报告','云同步、家长教师协作\niOS原生验收','家庭小规模试用优先\n机构方案后续验证']],[380,390,382],{y:164,h:410,size:24});
text(s,'当前核心样题 30 道，另有 9 道可选扩充题，扩充内容仍待人工教研确认。',64,612,1140,48,24,C.amber);
s=slide('定位比较：优先验证独立练习与复习','约90秒。排序依据是项目现状与投入匹配，不是对市场规模或付费率的测算。家庭作业助手覆盖面广，需要较强识别和AI能力。掌握巩固方向与现有流程最接近，但同样有竞争。教师助手需要更多身份、授权和协作能力。先做一个单元试点，以真实使用决定是否继续。',appURLs.concat(['/Users/zhenshuai/Desktop/小禾学伴/APP-实施计划.md']));
table(s,[['方向','已有匹配','主要投入／竞争','当前建议'],['家庭作业辅导','拍照与固定提示','识别、AI及内容覆盖\n同类入口成熟','保留为基础能力'],['独立练习与复习','帮助记录、原答\n间隔复查','教研题族与规则验证\n竞品已有相近主张','优先做家庭试点'],['教师练习与学情','教研管理基础','真实身份、班级权限\n送达与回复','找到试点教师后推进']],[265,285,390,212],{h:315,size:25});
text(s,'核心主张假设：让孩子知道还需要练什么，让家长看懂为什么。',64,538,1145,70,32,C.green,true);
text(s,'购买方尚未确定。家长付费、教师采用与机构采购均需要直接证据。',64,623,1140,42,23,C.amber);
s=slide('下一版的三项优先改进','约100秒。首先补内容，因为没有足够未见题就无法持续复查。其次做清楚提示与独立作答的流程，保留原答和帮助状态。第三把报告变成家长可以理解和行动的摘要。AI追问与教师协作作为探索，不在首轮同时铺开。',improvements.map(r=>`${r[1]}：${r[3]}`));
table(s,[['优先项','具体改动','验证方式'],['P0 可信内容','一个单元的教材映射\n审核答案与不同结构题族','内容审核通过\n复查任务有未曝光题'],['P1 独立练习','先作答、再逐级提示\n帮助状态与次日复查','记录帮助前后表现\n复查原答可追溯'],['P1 家长反馈','说明当前困难\n给出下一次练习安排','家长能复述问题与行动\n检验持续使用意愿']],[250,440,462],{h:330,size:26});
text(s,'后续探索：受题库约束的 AI 追问；小规模教师协作。',64,550,1140,50,28,C.green,true);
text(s,'资源重点是教研、真实使用验证和可靠交付；本轮不估算尚无依据的工期与费用。',64,622,1140,44,23,C.muted);
s=slide('接下来完成实测，再决定试点投入','约75秒。今天交付的是可追溯的案头版。建议接下来用相同题目完成三个App的四个任务，次日复查；邀请少量家长和教师访谈，在取得相应同意后观察儿童使用。最后更新定位结论和PPT，老板确认试点用户、内容范围和资源即可。');
table(s,[['安排','具体工作','产出'],['第1—2天','核对资料，补手机端免费／付费边界','3份操作记录与截图'],['第3—4天','统一6题、4任务实测，次日复查','12项任务结果与复测记录'],['第5天','访谈家长与试点教师，核验定位','需求与付费假设的支持／反证'],['第6—7天','补证据、调整建议、15分钟试讲','汇报更新版与试点范围']],[210,570,372],{h:335,size:25});
text(s,'请老板确定：先验证家庭场景，先做哪个数学单元，谁负责教研与试用组织。',64,556,1140,85,31,C.green,true);

await fs.writeFile(path.join(BUILD,'deck.json'),JSON.stringify(pres.toProto()));
for(const [i,sl]of pres.slides.items.entries()){const blob=await pres.export({slide:sl,format:'png',scale:1});await fs.writeFile(path.join(PRE,`slide-${String(i+1).padStart(2,'0')}.png`),new Uint8Array(await blob.arrayBuffer()));}
await fs.writeFile(path.join(BUILD,'deck-inspect.txt'),(await pres.inspect({kind:'slide,textbox,table,notes',maxChars:20000})).ndjson);
const candidate=path.join(BUILD,'candidate.pptx');await(await PresentationFile.exportPptx(pres)).save(candidate);
const SKILL='/Users/zhenshuai/.codex/plugins/cache/openai-primary-runtime/presentations/26.1026.10807/skills/presentations';
const {finalizePresentation}=await import(pathToFileURL(path.join(SKILL,'container_tools/artifact_tool_utils.mjs')).href);
await finalizePresentation({workspaceDir:ROOT,candidatePath:candidate,finalPath:path.join(OUT,'小学智能学习产品调研与小禾学伴定位建议.pptx'),pythonExecutable:'/Users/zhenshuai/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3',integrityValidatorPath:path.join(SKILL,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(SKILL,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...[2,3,4,5,6,9,10,11,12].flatMap(n=>['--require-native-table-slide',String(n)])],explicitTotalSlideCount:12,requiredNativeTableOwnerSlides:[2,3,4,5,6,9,10,11,12],requiredNativeChartOwnerSlides:[],fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(BUILD,'validation.json')});
await fs.writeFile(path.join(BUILD,'speaker-notes.txt'),pres.slides.items.map((sl,i)=>`第${i+1}页\n${sl.speakerNotes.text}`).join('\n\n'));
console.log(JSON.stringify({outputs:await fs.readdir(OUT),products:products.length,reviews:reviewRows.length,evidence:evid.length,slides:pres.slides.items.length}));
