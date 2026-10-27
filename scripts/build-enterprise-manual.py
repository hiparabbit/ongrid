"""Generate the formal Chinese operations manual; render and inspect every page before publication."""
import argparse
from pathlib import Path
from PIL import Image
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('--screenshots', type=Path, required=True)
parser.add_argument('--output', type=Path, required=True)
parser.add_argument('--url', default='http://192.168.1.11:18860')
parser.add_argument('--version', required=True)
args = parser.parse_args()
doc = Document()
s = doc.sections[0]
s.page_width, s.page_height = Cm(21), Cm(29.7)
s.top_margin, s.bottom_margin = Cm(3.7), Cm(3.5)
s.left_margin, s.right_margin = Cm(2.8), Cm(2.6)
s.footer_distance = Cm(2.5)
# GB/T 9704-style typography and hierarchy. Use available serif CJK fallback
# when the reader does not have the named Chinese typeface installed.
for name, font, size in [('Normal','仿宋',16),('Title','方正小标宋简体',22),('Heading 1','黑体',16),('Heading 2','楷体',16),('Caption','宋体',12)]:
    style=doc.styles[name]
    style.font.name=font
    style.element.get_or_add_rPr().get_or_add_rFonts().set(qn('w:eastAsia'),font)
    style.font.size=Pt(size)
    style.font.color.rgb=RGBColor(0,0,0)
    style.font.bold=False
    style.paragraph_format.line_spacing=Pt(28 if name!='Caption' else 20)
    style.paragraph_format.space_before=Pt(0)
    style.paragraph_format.space_after=Pt(0)
    style.paragraph_format.first_line_indent=Pt(32 if name=='Normal' else 0)
    style.paragraph_format.widow_control=True
for name in ['Heading 1','Heading 2']:
    doc.styles[name].paragraph_format.keep_with_next=True
for border in list(doc.styles.element.iter(qn('w:pBdr'))):
    border.getparent().remove(border)
footer=s.footer.paragraphs[0];footer.alignment=1
footer.style=doc.styles['Caption']
footer.add_run('— ')
field=OxmlElement('w:fldSimple');field.set(qn('w:instr'),'PAGE');footer._p.append(field)
footer.add_run(' —')
doc.core_properties.title='企业AI智能运维平台操作手册'
doc.core_properties.subject='平台操作与监控管理'
doc.core_properties.author='企业运维管理中心'
doc.core_properties.keywords='运维 监控 操作手册'
chapters=[
('一、平台登录','login','使用账号、密码表单进入企业运维工作台。',['访问平台地址，输入账号和密码。','账号至少３位，密码至少６位。','点击“登录”；退出使用账户菜单。'],'登录仅作本地表单校验，不提供服务端身份认证；不得输入其他系统的真实密码。'),
('二、工作台与业务导航','overview','顶部业务导航与左侧功能菜单共同组织操作入口。',['查看资源、风险和执行情况。','通过顶部“监控中心”切换区域。','窄屏使用左上角导航按钮。'],'所有业务数据为本地样本快照，不代表实时采集结果。'),
('三、统一监控总览','monitor-overview','集中查看基础设施与应用服务的监控分布。',['进入“监控中心—监控总览”。','核对监控对象和关注对象数量。','点击分类卡片查看监控明细。'],'总览按六类对象汇总；关注状态与对象快照保持一致。'),
('四、主机监控','monitor-hosts','检查主机处理器、内存、磁盘和系统负载。',['进入“主机监控”，搜索对象。','按健康状态筛选异常主机。','点击“查看详情”检查趋势。'],'详情支持６、１２、２４小时时间范围，并可跳转资源台账。'),
('五、应用服务监控','monitor-applications','查看请求速率、响应延迟、错误率与实例数量。',['进入“应用服务”，检索服务。','重点检查Ｐ９５延迟和错误率。','打开详情，比较各指标趋势。'],'延迟使用毫秒，流量使用每秒请求数；指标来自样本快照。'),
('六、数据库监控','monitor-databases','检查查询负载、连接使用率、复制延迟和缓存命中率。',['进入“数据库监控”。','按名称或业务域检索实例。','打开详情，核对连接及复制指标。'],'对连接池高使用率和复制延迟较大的对象优先排查。'),
('七、网络设备监控','monitor-network','检查入站和出站流量、丢包率以及链路延迟。',['进入“网络设备”。','筛选“需要关注”的设备。','查看详情并核对丢包和延迟。'],'流量单位为Ｍｂｐｓ，延迟单位为毫秒。'),
('八、容器集群监控','monitor-clusters','查看集群节点、运行容器组及资源分配情况。',['进入“容器集群”。','查看节点数及运行Ｐｏｄ数量。','打开详情，检查资源分配趋势。'],'资源分配率用于说明快照容量，不执行调度或扩缩容。'),
('九、接口可用性检查','monitor-availability','查看接口可用率、响应时间和证书剩余有效期。',['进入“接口可用性”。','按接口地址或名称检索。','核对响应时间与证书剩余天数。'],'接口地址为样本标识；页面不会向目标地址发起探测。'),
('十、日志检索','monitor-logs','按服务、关键字和日志级别检索结构化事件。',['进入“日志检索”，输入关键字。','选择日志级别以缩小结果范围。','查看详情，核对追踪编号。'],'共提供６０条样本记录；无结果时可清除筛选。'),
('十一、告警规则管理','monitor-rules','维护主机、接口、数据库等监控对象的阈值配置。',['进入“告警规则”。','点击“编辑阈值”，填写正数。','保存规则；取消时保持原配置。'],'规则仅保存在当前浏览器，不下发采集端，也不重新计算已有告警。'),
('十二、资源台账与指标','assets','按业务域查看资源信息、运行状态和容量。',['搜索资源名称或地址。','打开详情，查看关联告警。','跳转“指标监测”检查趋势。'],'台账包含２４项资源；表格可在自身容器内横向滚动。'),
('十三、指标趋势检查','metrics','按资源和时间范围检查处理器与内存趋势。',['选择目标资源。','切换６、１２、２４小时时间范围。','将鼠标移至曲线查看数值。'],'曲线为本地构造的模拟趋势，不表示现场采集值。'),
('十四、告警处置','incident','告警按待处置、处理中和已关闭三个状态管理。',['搜索并打开ＩＮＣ－１２０１。','点击“确认告警”进入处理中。','处置完成后确认关闭告警。'],'关闭需确认，取消不改变状态；已关闭告警不可再次确认。'),
('十五、智能分析','assistant','智能助手提供高负载、接口延迟和健康巡检分析入口。',['选择“高ＣＰＵ应急排查”。','查看告警证据和参考知识。','按需要进入关联流程。'],'回答为预设模拟分析，不调用真实模型，不作为生产诊断结论。'),
('十六、自动化流程','automation','流程按标准步骤组织健康巡检和异常排查作业。',['选择流程，点击“模拟执行”。','阅读确认信息后确认或取消。','完成后打开关联执行记录。'],'流程不会连接设备或执行远程命令，取消不会产生记录。'),
('十七、执行记录','executions','通过执行编号、时间和结果追溯流程操作。',['进入“执行记录”。','核对流程名称和执行时间。','刷新页面，检查记录保留情况。'],'初始６条记录，最多保留１００次模拟执行；完成不代表真实作业成功。'),
('十八、知识维护','knowledge','知识文档支持检索、阅读、新建与修改。',['按标题检索并打开知识。','点击“编辑文档”修改后保存。','新建时填写必填标题和正文。'],'编辑只修改当前语言；保存失败时保留文本副本。'),
('十九、个人设置与数据恢复','account','账户菜单提供语言、主题、数据恢复和退出入口。',['选择中英文语言或深浅主题。','恢复数据前备份需保留的文本。','确认恢复或取消，再核对结果。'],'恢复会替换本浏览器的业务记录；监控规则单独保存，不随业务数据恢复。'),
('二十、平台信息与使用说明','about','平台说明提供版本、数据范围和软件下载入口。',['核对当前运行版本。','需要时下载操作手册。','软件许可和对应源码见下载入口。'],'业务状态仅在当前浏览器保存；更换访问地址或清理站点数据后，原状态不可读取。'),
]
p=doc.add_paragraph('企业AI智能运维平台\n操作手册','Title');p.alignment=1
for text in ['企业运维管理中心',f'版本：{args.version}',f'访问地址：{args.url}','2026年10月']:
    p=doc.add_paragraph(text);p.alignment=1;p.paragraph_format.first_line_indent=0;p.paragraph_format.space_before=Pt(28)
doc.add_page_break();doc.add_heading('目　录',1)
for i,(title,*_) in enumerate(chapters):
    p=doc.add_paragraph(f'{title}　{i+3}');p.paragraph_format.first_line_indent=0
for title,shot,intro,items,result in chapters:
    doc.add_page_break();doc.add_heading(title,1)
    doc.add_paragraph(intro)
    for i,item in enumerate(items,1):doc.add_paragraph(f'{i}. {item}')
    path=args.screenshots/f'{shot}.png'
    with Image.open(path) as im:width=min(15.6,7.2*im.width/im.height)
    p=doc.add_paragraph();p.paragraph_format.first_line_indent=0;p.paragraph_format.line_spacing=1;p.alignment=1
    p.add_run().add_picture(str(path),width=Cm(width))
    p=doc.add_paragraph(f'图　{title.split("、",1)[1]}','Caption');p.alignment=1
    doc.add_paragraph(result)
args.output.parent.mkdir(parents=True,exist_ok=True);doc.save(args.output);print(args.output)
