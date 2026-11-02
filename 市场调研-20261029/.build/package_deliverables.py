from pathlib import Path
from html.parser import HTMLParser
from urllib.parse import urlparse, unquote, quote
import json, zipfile, xml.etree.ElementTree as ET, hashlib, html, shutil

root=Path('/Users/zhenshuai/Desktop/小禾学伴/市场调研-20261029')
shots=[
 ('01-作业帮-官网.jpg','https://www.zybang.com/'),
 ('02-小猿AI-官网.jpg','https://www.xiaoyuankousuan.com/'),
 ('03-洋葱学园-官网.jpg','https://yangcongxueyuan.com/study-platform/'),
 ('05-学而思T6Max-官网.jpg','https://xpad.xueersi.com/T6MaxSeries?spuId=4694&spuUnionid=8260707083725051150264694&skuId=0'),
 ('08-作业帮AI学习桌-官网.jpg','https://smart.zuoyebang.com/eduDesk')]
index=root/'来源与截图/来源索引.html'
content=index.read_text()
if 'id="browser-shots"' not in content:
 rows=''.join(f'<tr><td><a href="{quote(n)}">{html.escape(n)}</a></td><td><a href="{html.escape(u)}" target="_blank" rel="noopener">{html.escape(u)}</a></td><td>2026-10-29 · 官网页面展示，非 App 实测</td></tr>' for n,u in shots)
 section='<section id="browser-shots"><h3>本次保存的官网页面截图</h3><div class="scroll"><table><thead><tr><th>截图文件</th><th>页面来源</th><th>日期与口径</th></tr></thead><tbody>'+rows+'</tbody></table></div></section>'
 content=content.replace('<h2 id="files">',section+'<h2 id="files">')
 index.write_text(content)
snap=root/'.build/source-snapshots';snap.mkdir(exist_ok=True)
for p in (root/'来源与截图').glob('*.txt'):shutil.move(str(p),snap/p.name)
for p in (root/'交付').glob('*.inspect.ndjson'):shutil.move(str(p),root/'.build'/p.name)

class LinkParser(HTMLParser):
 def __init__(self):super().__init__();self.links=[];self.ids=set()
 def handle_starttag(self,tag,attrs):
  d=dict(attrs)
  if 'id' in d:self.ids.add(d['id'])
  for k in ('href','src'):
   if k in d:self.links.append(d[k])
link_checks=[]
for p in [index,root/'实测与访谈手册.html']:
 parser=LinkParser();parser.feed(p.read_text())
 for raw in parser.links:
  u=urlparse(raw)
  if u.scheme or u.netloc:continue
  if not u.path:
   assert not u.fragment or u.fragment in parser.ids,(p,raw)
  else:assert (p.parent/unquote(u.path)).resolve().exists(),(p,raw)
 link_checks.append({'file':str(p.relative_to(root)),'links':len(parser.links),'status':'pass'})

reviews=json.loads((root/'research/reviews.json').read_text())['products']
assert len(reviews)==3 and all(len(p['reviews'])==20 for p in reviews)
ids=[r['review_id'] for p in reviews for r in p['reviews']]
assert len(set(ids))==60
pp=root/'交付/小学教育App与伴学设备功能调研.pptx'
xx=root/'交付/竞品对比表.xlsx'
with zipfile.ZipFile(pp) as z:
 assert z.testzip() is None
 slides=[n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')]
 notes=[n for n in z.namelist() if n.startswith('ppt/notesSlides/notesSlide') and n.endswith('.xml')]
 assert len(slides)==len(notes)==12
with zipfile.ZipFile(xx) as z:
 assert z.testzip() is None
 ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
 sheets=ET.fromstring(z.read('xl/workbook.xml')).find('s:sheets',ns)
 assert len(sheets)==4
 sheet_names=[s.attrib['name'] for s in sheets]
 assert set(sheet_names)=={'产品概况','功能对比','体验证据','改进机会'}

files=[pp,xx,root/'交付/汇报阅读说明.txt']+list((root/'来源与截图').glob('*'))+[root/'实测与访谈手册.html']
files += [root/'research'/n for n in ['apps.json','hardware.json','reviews.json','source_assets.json','local_and_protocol.json']]
files += list((root/'research/hardware-source-images').glob('*'))
files=[p for p in files if p.is_file()]
manifest={'date':'2026-10-29','scope':'案头调研版；12项手机实测、次日复查与真实用户验证待完成','slides':12,'worksheets':sheet_names,'product_count':8,'review_count':60,'link_checks':link_checks,'files':[{'path':str(p.relative_to(root)),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in files]}
(root/'.build/delivery-manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2))
archive=root/'小禾学伴竞品功能调研资料包-20261029.zip'
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as z:
 for p in files:z.write(p,root.name+'/'+str(p.relative_to(root)))
with zipfile.ZipFile(archive) as z:assert z.testzip() is None
print(json.dumps({'archive':str(archive),'file_count':len(files),'bytes':archive.stat().st_size,'local_links':link_checks,'package_checks':'pass'},ensure_ascii=False))
