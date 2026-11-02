import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Workbook, SpreadsheetFile, Presentation, PresentationFile } from '@oai/artifact-tool';
process.env.RUNTIME_NODE_MODULES='/Users/zhenshuai/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';

const ROOT='/Users/zhenshuai/Desktop/小禾学伴/市场调研-20261029';
const OUT=path.join(ROOT,'交付');
const BUILD=path.join(ROOT,'.build');
const PRE=path.join(BUILD,'previews-v2');
await fs.mkdir(PRE,{recursive:true});
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
 ['P0','补齐教材同步内容与变式题族','竞品已有教材章节练习、同步课和互动题库。小禾当前核心样题30道，持续学练范围有限。','zuoyebang-F1；xiaoyuan-F1；yangcong-F1/F4；APP-实施计划.md','选定一个四年级数学单元，补充教材映射、答案审核及不同结构的同知识点题目。','每个知识点有经审核的讲解、练习和未见复查题；不把重复做原题当作掌握。','教研审核、题目编写与内容授权，投入较高','待实施'],
 ['P1','连通讲解、再练与错题复查','小猿介绍错题分类与举一反三，洋葱介绍互动练习与核验。小禾已有提示记录、原答和间隔复查，可完善衔接。','xiaoyuan-F4/F5；yangcong-F4；APP-实施计划.md','讲解后接一道未见题；错题页显示独立作答与提示情况，并给出后续练习入口。','能从一道错题完成讲解、再练、订正及复查；帮助前后的表现可追溯。','交互调整、题库规则和真实用户测试，投入中等','待验证'],
 ['P1','让学习反馈给出下一步','硬件配套App展示学习轨迹、周报与管理功能。小禾已有本地档案和报告，但协作与可理解性仍待完善。','H03-F3；H04-F2/F4；APP-实施计划.md','报告显示知识点困难、提示依赖、下次练习和复查安排；先做可分享摘要。','使用者能说清当前困难和下一步；入口可执行，不仅展示分数或星级。','报告交互、用户测试，投入中等','待验证'],
 ['探索','受题库约束的AI追问','三个App均公开介绍AI讲解或提问。小禾目前采用固定提示。','zuoyebang-F2；xiaoyuan-F2；yangcong-F3；APP-实施计划.md','先在一个单元内支持连续追问，并核对解释、错误纠正和答案出现时机。','解释正确且可追溯；无法可靠回答时说明限制。','模型接入、内容审核与评测，投入较高','待研究'],
 ['探索','教师布置与跟进','小猿官网列班级批改与高频错题统计，洋葱有单独教师端。小禾正式协作尚未完成。','xiaoyuan-F7；yangcong-F6；APP-实施计划.md','有试点教师后，补最小布置练习、查看错题和回复流程。','教师可以完成任务并愿意持续使用。','身份、权限和云端协作，投入较高','待研究']
];
const matrix=[
 ['作业帮 App','多学段多学科\n按教材章节组织','单题／整页／作文拍题\n作业、口算、作文批改','AI解析与追问\n相似题视频','知识点与视频讲解','同步练习、单元测试','错题讲解\n归档与跨日规则待核','口算、词汇、诗词竞赛','面向家庭辅导\n独立报告入口待核','主App班级功能未查到','在线练习与试卷打印','免费下载＋VIP\n具体权益见产品概况'],
 ['小猿 AI','小学至高中\n教材同步课与练习','拍题搜题、多学科批改','分步解析\n视频与AI讲解','教材同步课程','同步练习、薄弱点匹配','错题分类、错因分析\n举一反三','学习计划（会员）\n口算PK、听写、词汇等','家长具体报告待实测','班级批改、高频错题统计\n资料打印分发','手机拍照、资料打印','免费下载＋VIP\n基础解析与AI权益待测'],
 ['洋葱学园','小学至中职资源\n具体教材套餐待核','拍题定位知识点\n批改能力未查到','动画讲解\n课中暂停向AI提问','教材同步动画课','互动题库\n思考训练与核验主张','自动归档和跨日规则待核','AI计划、每日任务\n游戏化激励','学生App家长报告待核','单独教师端\n不归入学生App','在线课程与题库','按科目、学段、时长购买\n小学现价与AI权益待核'],
 ['讯飞 T30 Ultra／产品线','教材同步、互动课\n小学英语阅读（产品线）','互动答疑与错因分析\n均为产品线介绍','启发式问答、板书讲解\n需核对机型与学段','同步内容与互动课\n按机型核权益','诊断、针对性练习\n学情组卷（产品线）','具体复习规则待核','部分规划功能限初中数学','家长具体权限待核','未查到具体协作范围','14.7英寸、翻转双摄\n手写笔（T30 Ultra）','硬件购买\n目录价与完整权益见概况'],
 ['学而思 T6 Max 页面','培优与英语资源\n具体版本套餐待核','拍题诊断、AI提问','AI讲题、互动培优','英语阅读、口语等内容\n页面级，套餐待核','学习规划、学情解读','错题处理与巩固','小思学伴与学习记忆\n不等于掌握判定','家长具体权限待核','未查到具体协作范围','纸屏互动\n配件按具体SKU核验','硬件＋会员入口\n精确SKU价与权益待核'],
 ['作业帮 P60／硬件生态','同步课程、启蒙\n为产品线能力','AI家教式作业辅导\n为P60商品页主张','AI老师／家教式辅导','同步课程（产品线）','练习细节待核','错题、生词同步*','配套App记录与周报*','远程功能管理*','未查到具体协作范围','学习平板\n具体规格待核','商品页2499—5499元区间\n不同SKU不直接等价'],
 ['小猿 S6／学练机生态','S6具体内容覆盖待核','答案解析查看控制*','本机讲解方式待核','本机课程范围待核','具体练习流程待核','每日学习轨迹*\n错题复查规则待核','使用时长约定*','多孩子、多设备管理*','未查到具体协作范围','官方目录归为学练机\n不套用手机App功能','硬件价格与完整权益待核'],
 ['作业帮 AI学习桌','小学语数英同步练\n10余种教材需逐项核对','指读、语音交互\n全题型批改待核','纸面伴学\n实际讲解范围待核','英语视频学习','同步练习、听写','未查到','计时与定时提醒','家长反馈能力待核','未查到','桌面设备、照明与桌椅\n连接纸面学习','具体在售SKU与价格待核']
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
const sm=sheet('功能对比','8 个竞品的 11 个功能与使用维度',['产品','内容与适配','输入与批改','讲解与追问','课程学习','练习与测试','错题与复习','计划与激励','家长反馈','教师协作','设备与纸面','收费方式'],matrix,[220,240,250,250,240,240,240,250,230,250,250,260]);
sm.getRange('A4').values=[['市场调研主表。按功能查看各产品当前公开能力。* 为硬件配套App能力，P60／S6适配待核。来源及范围见体验证据。']];
const overview=products.map((p,i)=>[p.id||`P${i+1}`,p.name,p.form,versionText(p),p.audience,p.positioning||(p.features||[]).map(f=>`（${f.scope||'App介绍'}）${f.claim}`).join('；'),pricingText(p)||(typeof p.pricing==='object'?JSON.stringify(p.pricing):'未核实'),(p.unknowns||[]).join('\n'),new Date('2026-10-29T00:00:00Z'),(p.sources||[]).slice(0,2).map(s=>s.url).join('\n')]);
const so=sheet('产品概况','产品身份、定位与收费口径',['编号','产品','形态','版本／型号','服务对象','产品主张／公开功能','收费与限制','待核实','检索日期','来源链接'],overview,[100,220,150,220,240,300,350,340,120,430]);
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
si.getRange('A4').values=[['小禾改进为附属部分。每项建议从竞品功能出发，对照小禾已有能力和差距，再决定借鉴动作。']];
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
let s=slide('小学教育 App 与伴学设备\n功能调研','约30秒。本次汇报以竞品目前提供的功能为主体。第1—10页约占12分钟，介绍8个样本、功能、使用方式、收费和公开反馈。第11—12页约占3分钟，讨论小禾可以借鉴的具体功能。依据2026-10-29的官网、商店和公开评论，尚无手机登录后操作实测。',appURLs.concat(hwURLs));
text(s,'作业帮、小猿 AI、洋葱学园',64,280,1152,50,32,C.green);
text(s,'讯飞、学而思、作业帮、小猿学习设备及作业帮 AI 学习桌',64,345,1152,72,29,C.ink);
text(s,'竞品功能调研约 80%    小禾改进参考约 20%',64,466,1152,56,31,C.green,true);
text(s,'公开资料版    2026 年 10 月 29 日',64,581,1152,40,24,C.muted);

s=slide('竞品已覆盖从作业处理到学习管理的多项功能','约60秒。这页概括功能范围。之后逐个产品展开，避免只罗列AI、错题本等功能名称。公开资料表明某能力被介绍过，不代表已经验证其准确性或四年级账号下的可用性。详情和证据范围记录在Excel。',appURLs.concat(hwURLs));
table(s,[['功能模块','目前公开介绍的具体功能','代表样本'],['作业处理','拍题检索、口算与多学科批改','作业帮、小猿 AI'],['讲解与提问','分步解析、视频或动画、AI追问','三个 App、讯飞、学而思'],['课程与练习','教材同步课程、互动练习、单元测试','三个 App及学习机产品线'],['错题与反馈','错因分析、同类再练、学情与学习记录','小猿 AI、学而思及配套 App'],['持续使用','学习计划、每日任务、PK与游戏激励','小猿 AI、洋葱学园'],['协作与设备','班级批改、远程管理、指读与语音伴学','教师端、配套 App、学习桌']],[210,585,357],{y:164,h:418,size:24});
text(s,'证据：官网、商店与60条公开评论。硬件按产品线、型号页和配套 App 区分。',64,615,1140,48,23,C.muted);

s=slide('作业帮：查题、批改与同步练习','约90秒。按六项功能介绍作业帮主App。当前中国区iOS版本14.55.0，10月27日更新。用户可以围绕拍题求助、检查作业和按章节练习使用。主App已介绍错题解析，但归档和跨日复习规则待实测。收费来自中国区商店展示，不代表安卓、首购或每个账号的实际价格。',apps[0].sources.map(x=>`${x.title}：${x.url}`));
await pic(s,'官方展示-zuoyebang-1.png',65,168,225,400);
table(s,[['功能','具体能力'],['拍题检索','支持单题、整页及作文检索'],['讲解追问','AI解析、追问与薄弱点讲解'],['作业批改','作业、口算、作文批改及错题讲解'],['同步练习','按教材章节推荐练习与单元测试'],['关联学习','推荐知识点与相似题视频'],['拓展练习','口算、词汇、诗词竞赛及打印练习']],[182,674],{x:360,y:174,w:856,h:378,size:23});
text(s,'iOS 14.55.0。VIP展示：连续包月19元，单月25元。具体免费范围待实测。',64,590,1150,42,23,C.ink);
text(s,'左图为开发者展示。来源：作业帮中国区 App Store。',64,635,1150,36,21,C.muted);

s=slide('小猿 AI：从批改延伸到错题再练与班级工具','约90秒。当前名称是小猿AI，原小猿口算，iOS3.145.0，10月26日更新。功能不仅是口算检查，还包括错因、举一反三、同步课程，以及官网列示的班级批改。班级工具配置、家长入口和使用规模要在实际账号核对。官网免费解析与会员AI讲解并存，应分功能核实。',apps[1].sources.map(x=>`${x.title}：${x.url}`));
await pic(s,'官方展示-xiaoyuan-1.png',65,168,225,400);
table(s,[['功能','具体能力'],['拍照批改','多学科错误标注与错因分析'],['拍题讲解','分步解析、视频与AI讲解'],['错题再练','错题分类与举一反三练习'],['同步学练','教材同步课程及同步练习'],['日常训练','口算PK、听写、诗词、词汇和口语'],['班级工具','批改、高频错题统计、资料打印分发']],[182,674],{x:360,y:174,w:856,h:378,size:23});
text(s,'iOS 3.145.0。VIP展示：连续包月25元、包年228元，含部分讲解与计划权益。',64,590,1150,42,23,C.ink);
text(s,'左图为开发者展示。来源：小猿 AI 官网与中国区 App Store。',64,635,1150,36,21,C.muted);

s=slide('洋葱学园：动画课程、互动练习与 AI 学伴','约90秒。iOS7.118.0，10月26日更新。洋葱从知识点课程切入，拍题后先定位知识点，以动画帮助理解，也支持课中暂停向AI私教提问。学习计划和每日任务用于持续学习。官网提到思考训练与核验，但不推断具体掌握算法。教师版为单独产品，不计入学生App能力。',apps[2].sources.map(x=>`${x.title}：${x.url}`));
await pic(s,'官方展示-yangcong-1.jpg',65,168,225,400);
table(s,[['功能','具体能力'],['同步课程','教材同步动画课、习题和试卷资源'],['拍题精学','定位知识点，动画讲解后引导解题'],['课中提问','暂停动画课并向AI私教提问'],['互动练习','互动题库与思考训练、核验主张'],['学习安排','AI定制计划与每日任务'],['持续激励','游戏化学习激励']],[182,674],{x:360,y:174,w:856,h:378,size:23});
text(s,'iOS 7.118.0。课程按学段、科目和时长收费，小学套餐与AI权益待核实。',64,590,1150,42,23,C.ink);
text(s,'左图为开发者展示。来源：洋葱官网与中国区 App Store。',64,635,1150,36,21,C.muted);

s=slide('讯飞与学而思：课程、诊断和讲练功能的组合','约90秒。讯飞样本型号为T30 Ultra，2024年发布，目前官网仍列示。多数功能来自产品线页，不默认全部可用于T30 Ultra和小学四年级。学而思为T6 Max页面，经典款标题与国际版素材并存，具体SKU需另核。这两种资料都说明学习机尝试把内容、辅导和后续练习组合起来。',hard.slice(0,2).flatMap(p=>p.sources.map(x=>`${x.title}：${x.url}`)));
table(s,[['模块','讯飞 T30 Ultra／产品线','学而思 T6 Max 页面'],['答疑讲解','启发式问答、错因分析\n板书讲解（产品线）','AI提问与讲题\n互动培优'],['诊断与练习','薄弱点诊断、针对性练习\n学情组卷（产品线）','拍题诊断、学情解读\n错题处理与学习规划'],['内容资源','教材同步、互动课程\n小学英语阅读（产品线）','英语阅读、口语家教\n课程按版本和套餐核实'],['设备与学伴','14.7英寸、翻转双摄\n手写笔（T30 Ultra型号）','纸屏互动、小思学伴\n学伴记忆范围待核'],['费用口径','官方目录11999元\n门店自提，非最终结算价','精确SKU价未核实\n完整会员权益待核']],[180,486,486],{y:160,h:438,size:23});
text(s,'讯飞部分规划功能限初中数学。T6 Max页面混合版本素材，不能推定全部权益。',64,617,1150,51,22,C.amber);

s=slide('作业帮与小猿设备：辅导功能及家长配套服务','约90秒。P60商品页、作业帮硬件产品线页和配套App属于三种不同证据范围。小猿S6的官方目录只能确认学练机分类，更多管理功能来自小猿智能设备App，不能拿小猿AI手机App功能补齐S6。星号项是配套App公开能力，具体机型绑定后支持范围待核。',hard.slice(2,4).flatMap(p=>p.sources.map(x=>`${x.title}：${x.url}`)));
table(s,[['模块','作业帮 P60／硬件生态','小猿 S6／学练机生态'],['设备定位','P60商品页：AI老师\n家教式全科作业辅导','官方目录列为学练机\n本机讲练功能需补资料'],['内容与讲解','同步课、启蒙、作业辅导\n为产品线介绍','配套App可控制\n答案解析查看*'],['学习记录','学习记录、每周报告\n错题与生词同步*','每日功能使用\n学习数据与轨迹*'],['家庭管理','配套App远程\n功能管理*','多孩子、多设备管理\n使用时长约定*'],['费用口径','商品页2499—5499元区间\n不同配置与套餐待区分','具体SKU价格未查到\n完整内容权益待核']],[180,486,486],{y:160,h:438,size:23});
text(s,'* 为硬件配套 App 能力，对 P60／S6 的具体适配待核。产品线能力不等于型号权益。',64,617,1150,51,22,C.amber);

s=slide('作业帮 AI 学习桌：把伴学功能放进纸面场景','约75秒。学习桌资料来自官网产品线图，具体在售型号、代际、库存和价格尚未确认。已展示小学同步练、指读、语音、听写、英语内容和计时提醒。指读不等于全题型讲解，提醒也不代表已验证专注效果。本页介绍功能和设备条件，不引用健康功效宣传。',hard[4].sources.map(x=>`${x.title}：${x.url}`));
await pic(s,'08-作业帮AI学习桌-官网.jpg',64,172,490,310);
text(s,'桌面设备连接纸面作业\n整合照明、桌椅与伴学入口',64,518,490,91,27,C.green);
table(s,[['功能','公开展示内容'],['教材同步','小学语数英同步练'],['指读','纸面指读交互'],['语音','语音交互与伴学'],['英语','英语视频学习内容'],['听写','配合纸面听写与练习'],['时间管理','计时与定时提醒']],[162,456],{x:598,y:172,w:618,h:402,size:24});
text(s,'官网展示10余种教材版本，需逐项核对。具体型号、价格与课程收费未核实。',64,627,1150,43,22,C.amber);

s=slide('8 个样本的核心功能对照','约90秒。将前面产品介绍放在相同维度下比较。表中是公开列示的能力，不做准确率打分。星号代表硬件产品线或页面资料，双星号是配套App能力。没有查到的项目保留为待核，不能解释为没有。更细的内容、激励、纸面、收费等11个维度见Excel主表。',appURLs.concat(hwURLs));
table(s,[['产品','答疑与讲解','课程与练习','错题与记录','家长／教师'],['作业帮 App','拍题、AI追问','同步练、单元测','错题讲解','主App报告待核'],['小猿 AI','批改、视频／AI','同步课、同步练','分类、错因、变式','班级批改统计'],['洋葱学园','拍题、动画／AI','动画课、互动练','归档规则待核','单独教师端'],['讯飞学习机','问答、板书*','诊断、针对性练*','错因、学情组卷*','权限细节待核'],['学而思学习机','诊断、AI讲题*','规划、培优内容*','错题处理*','权限细节待核'],['作业帮学习机','AI家教主张','同步课程*','错题与生词同步**','周报、远程管理**'],['小猿学练机','答案查看管控**','S6具体流程待核','每日学习轨迹**','家庭与时长管理**'],['作业帮学习桌','指读、语音伴学','同步练、英语内容','未查到','具体反馈待核']],[180,244,244,242,242],{y:158,h:432,size:21});
text(s,'*产品线／页面能力，须核对型号与学段。**配套App能力，P60／S6适配待核。',64,619,1150,50,22,C.amber);

s=slide('用户反馈提示需关注收费、答疑可靠性与操作体验','约90秒。公开RSS各取50条再按规则各选20条，共60条评论。用户身份和年级不明，不能当小学用户满意度调查，也不能用星级排名。将具体反馈当成下一步手机操作需要核对的问题。作业帮评论14717577839肯定解题思路，14718060879反馈解析与付费；小猿14716890086认可便利性，14717720398提到同题不同答案；洋葱14677388550肯定课程同时反馈旋转和AI符号问题，14690021504反馈iPad适配。',reviews.products.map(p=>`${p.product_name}：${p.reviews[0].source_url}`));
table(s,[['产品','评论中的肯定','评论中的问题线索','需核对的功能体验'],['作业帮','解题思路、批改\n辅导便利性','解析表现不一致\n付费边界争议','同题重复提问\n解析、视频与追问收费'],['小猿 AI','检查方便\n角色与互动吸引力','同题出现不同答案\nPK与激励规则争议','错题纠正、答案一致性\n激励是否影响任务'],['洋葱学园','课程讲解详细\n搜索与试用便利','横竖屏及iPad适配\nAI符号回答问题','课程后练习入口\n不同设备与AI交互']],[160,310,330,352],{y:171,h:342,size:24});
text(s,'总体可见：拍题与AI讲解、同步学练已被多家提供，差别在组织方式、内容和使用条件。',64,551,1150,80,27,C.green);
text(s,'各App20条目的抽样，反馈未经复现。不能据此计算缺陷率、满意度或产品排名。',64,636,1150,34,21,C.muted);

s=slide('竞品功能对小禾的直接参考','约90秒。前面约80%的内容介绍竞品功能，现在用两页说明小禾如何借鉴。根据2026-10-27项目文档，小禾已有练习、固定提示、原答、错题和间隔复查。这里先对照功能差距，不据此确定购买方或商业定位，也不把已有试用规则当作经过验证的学习效果。',appURLs.concat(hwURLs,['/Users/zhenshuai/Desktop/小禾学伴/APP-实施计划.md','/Users/zhenshuai/Desktop/小禾学伴/xiaohe-app/content-trial-notes.md']));
table(s,[['竞品已提供的功能','小禾已有与差距','可借鉴的具体动作'],['教材同步课、练习与题库\n作业帮／小猿／洋葱','已有组卷与30道核心样题\n教材映射和题族仍需补充','先完善一个数学单元\n补讲解、练习和复查题'],['错题分类、变式与核验\n小猿 AI／洋葱学园','已有提示、原答和间隔复查\n讲解后再练衔接需完善','讲解后接一道未见题\n明确独立完成与帮助状态'],['学习轨迹、周报与管理\n作业帮／小猿配套 App','已有本地档案与报告\n反馈理解和协作仍待完善','报告说明困难及下一步\n提供可执行的练习入口']],[375,390,387],{y:173,h:357,size:24});
text(s,'借鉴依据是已查到的功能和小禾现状。具体使用效果再通过实际操作验证。',64,588,1150,65,25,C.muted);

s=slide('小禾下一版：三项优先改进、两项后续探索','约90秒。优先级来自竞品功能差距与现有实现基础。先补内容，才能支持不同题目的持续学练；再连通错题再练与复查；最后让报告给出可执行的安排。题目需要人工教研审核，AI追问需要评测，教师协作需要身份和权限。这些是改进参考，不扩展成市场定位决策。',improvements.map(r=>`${r[1]}：${r[3]}`));
table(s,[['优先项','改进内容','完成检查','主要投入'],['P0 同步内容','一个单元的教材映射\n讲解与不同结构题族','内容通过教研审核\n复查有未见题可用','教研、题目编写\n内容授权'],['P1 错题再练','讲解后独立作答\n帮助状态与后续复查','一次错题任务可走通\n作答与提示记录可追溯','交互、题库规则\n真实使用测试'],['P1 学习反馈','知识点困难、提示依赖\n下一次练习与复查安排','使用者能说明困难\n并完成下一步操作','报告交互\n用户测试']],[190,365,365,232],{y:173,h:338,size:24});
text(s,'后续探索：受题库约束的 AI 追问；教师布置、查看与跟进。',64,557,1150,70,29,C.green,true);
text(s,'详细功能、出处和实测记录区见配套 Excel。先补实测，再调整具体改法。',64,631,1150,40,22,C.muted);

await fs.writeFile(path.join(BUILD,'deck-v2.json'),JSON.stringify(pres.toProto()));
for(const [i,sl]of pres.slides.items.entries()){const blob=await pres.export({slide:sl,format:'png',scale:1});await fs.writeFile(path.join(PRE,`slide-${String(i+1).padStart(2,'0')}.png`),new Uint8Array(await blob.arrayBuffer()));}
await fs.writeFile(path.join(BUILD,'deck-v2-inspect.txt'),(await pres.inspect({kind:'slide,textbox,table,notes',maxChars:20000})).ndjson);
const candidate=path.join(BUILD,'candidate-v2.pptx');await(await PresentationFile.exportPptx(pres)).save(candidate);
const SKILL='/Users/zhenshuai/.codex/plugins/cache/openai-primary-runtime/presentations/26.1026.10807/skills/presentations';
const {finalizePresentation}=await import(pathToFileURL(path.join(SKILL,'container_tools/artifact_tool_utils.mjs')).href);
const ownerSlides=[2,3,4,5,6,7,8,9,10,11,12];
await finalizePresentation({workspaceDir:ROOT,candidatePath:candidate,finalPath:path.join(OUT,'小学教育App与伴学设备功能调研.pptx'),pythonExecutable:'/Users/zhenshuai/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3',integrityValidatorPath:path.join(SKILL,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(SKILL,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...ownerSlides.flatMap(n=>['--require-native-table-slide',String(n)])],explicitTotalSlideCount:12,requiredNativeTableOwnerSlides:ownerSlides,requiredNativeChartOwnerSlides:[],fontPolicy:{basis:'design',families:[FONT]},verifyArtifactToolImport:true,receiptPath:path.join(BUILD,'validation-v2.json')});
await fs.writeFile(path.join(BUILD,'speaker-notes-v2.txt'),pres.slides.items.map((sl,i)=>`第${i+1}页\n${sl.speakerNotes.text}`).join('\n\n'));
console.log(JSON.stringify({slides:pres.slides.items.length,competitive_pages:10,xiaohe_pages:2,matrix_products:matrix.length,matrix_dimensions:matrix[0].length-1}));
