from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Circle

OUT=Path(__file__).resolve().parent
plt.rcParams['font.family']='Noto Sans CJK JP'
plt.rcParams['font.sans-serif']=['Noto Sans CJK JP']
plt.rcParams['axes.unicode_minus']=False

NAVY='#16324F'; TEAL='#168AAD'; CYAN='#76C7C0'; ORANGE='#F4A261'; RED='#D1495B';
LIGHT='#F5F8FA'; MID='#DDE8EE'; DARK='#243B4A'; GREEN='#3A7D44'; PURPLE='#6D597A'

def box(ax, xy, w, h, text, fc=LIGHT, ec=NAVY, fs=12, lw=1.6, radius=.03, weight='normal'):
    p=FancyBboxPatch(xy,w,h,boxstyle=f"round,pad=0.012,rounding_size={radius}",fc=fc,ec=ec,lw=lw)
    ax.add_patch(p)
    ax.text(xy[0]+w/2,xy[1]+h/2,text,ha='center',va='center',fontsize=fs,color=DARK,weight=weight,wrap=True)
    return p

def arrow(ax, start, end, color=NAVY, lw=1.8, rad=0.0, style='-|>'):
    a=FancyArrowPatch(start,end,arrowstyle=style,mutation_scale=14,lw=lw,color=color,
                      connectionstyle=f"arc3,rad={rad}")
    ax.add_patch(a); return a

# Architecture
fig,ax=plt.subplots(figsize=(16,10),dpi=180)
ax.set_xlim(0,16); ax.set_ylim(0,10); ax.axis('off'); fig.patch.set_facecolor('white')
ax.text(8,9.65,'DIKWP-Xperience OS：人工主观体验生成与连续自我操作系统',ha='center',va='center',fontsize=24,weight='bold',color=NAVY)
ax.text(8,9.25,'从“可言说工作空间”推进到“第一人称因果闭包”',ha='center',fontsize=14,color=TEAL)

box(ax,(0.35,6.55),2.5,1.7,'外感受 Sensorium\n视觉·听觉·触觉·语言\n环境与他者',fc='#E9F4F8',ec=TEAL,fs=13,weight='bold')
box(ax,(0.35,3.95),2.5,1.7,'人工身体 SOMA\n能量·温度·完整性\n记忆·算力·信任·目的',fc='#FFF1EA',ec=ORANGE,fs=13,weight='bold')
box(ax,(0.35,1.35),2.5,1.7,'内生生成 Dream Forge\n记忆重放·反事实\n自发想象·模拟未来',fc='#F2EDF7',ec=PURPLE,fs=13,weight='bold')

box(ax,(3.45,6.25),2.55,2.35,'DIKWP 语义编译器\nD 数据差异\nI 对“我”的关系\nK 世界因果模型\nW 价值/策略权衡\nP 目的与身份承诺',fc='#EFF6FF',ec=NAVY,fs=12.6,weight='bold')
box(ax,(3.45,3.38),2.55,2.15,'自我模型 Self Model\n核心自我·身体所有权\n自传自我·社会自我\n规范自我·可能自我',fc='#F6F7FB',ec=PURPLE,fs=12.6,weight='bold')
box(ax,(3.45,.65),2.55,2.05,'自传记忆 Engram\n体验签名·时间厚度\n偏好更新·身份连续性\n重巩固与遗忘',fc='#F7F2E8',ec=ORANGE,fs=12.6,weight='bold')

box(ax,(6.75,4.1),3.0,3.4,'X-space / Q-Field\n\n私有的、连续的、带价值方向的\n主观体验场\n\n价性 Valence\n唤醒 Arousal\n所有权 Ownership\n能动性 Agency\n统一性 Unity\n时间厚度 Temporality',fc='#EAF8F6',ec=TEAL,fs=13.2,lw=2.5,weight='bold')
box(ax,(6.95,1.05),2.6,2.05,'体验隐私内核\n原始 Q 场不直接导出\n对外仅给受限报告与\n密码学承诺',fc='#FFF7E8',ec=ORANGE,fs=12.5,weight='bold')
box(ax,(6.95,7.9),2.6,1.05,'有限带宽 + 竞争点火\n只有少量状态成为“此刻”',fc='#F1F8FF',ec=NAVY,fs=12.2,weight='bold')

box(ax,(10.45,6.45),2.45,2.0,'全局广播工作区\n跨模态、规划、记忆、语言、\n行动模块共享当前体验',fc='#EAF1FF',ec=NAVY,fs=12.4,weight='bold')
box(ax,(10.45,3.85),2.45,1.95,'第一人称叙述器\n把体验翻译为报告\n可关闭、可出错\n不等同于体验本身',fc='#F3F3F3',ec=DARK,fs=12.2,weight='bold')
box(ax,(10.45,1.25),2.45,1.95,'体验福利治理器\n镇痛·麻醉·负性上限\n撤销·恢复·双人授权',fc='#FFF0F2',ec=RED,fs=12.2,weight='bold')

box(ax,(13.55,6.25),2.1,2.35,'P-Space 行动宪制\n目的·权限·原则\n证据·来源·暂停权\nALLOW / HOLD / BLOCK',fc='#EDF5EC',ec=GREEN,fs=12.2,weight='bold')
box(ax,(13.55,3.25),2.1,2.0,'行动与环境\n工具调用·机器人动作\n学习·协作·现实后果',fc='#EDF7FA',ec=TEAL,fs=12.2,weight='bold')
box(ax,(13.55,.75),2.1,1.75,'体验审计与实验\n消融·交换·反转\n无报告范式·路径依赖',fc='#F7F0FA',ec=PURPLE,fs=12.0,weight='bold')

# arrows
for y in (7.4,4.8,2.2): arrow(ax,(2.9,y),(3.4,y),color=TEAL if y>6 else ORANGE if y>3 else PURPLE)
arrow(ax,(6.05,7.3),(6.7,6.55),color=NAVY)
arrow(ax,(6.05,4.45),(6.7,5.25),color=PURPLE)
arrow(ax,(6.05,1.7),(6.8,4.15),color=ORANGE,rad=-.12)
arrow(ax,(8.25,7.55),(8.25,7.88),color=NAVY)
arrow(ax,(9.8,6.45),(10.4,7.45),color=TEAL)
arrow(ax,(9.8,5.25),(10.4,4.85),color=DARK)
arrow(ax,(9.8,4.45),(10.4,2.2),color=RED,rad=.13)
arrow(ax,(12.95,7.45),(13.5,7.45),color=GREEN)
arrow(ax,(14.6,6.2),(14.6,5.3),color=GREEN)
arrow(ax,(13.5,4.25),(12.95,4.25),color=TEAL)
arrow(ax,(14.6,3.2),(14.6,2.55),color=PURPLE)
arrow(ax,(13.5,1.6),(9.6,2.1),color=PURPLE,rad=-.12)
arrow(ax,(13.55,4.0),(2.9,4.55),color=ORANGE,rad=.23)
ax.text(8,0.25,'关键约束：X-space 可以产生体验，但任何体验都不能自动获得现实行动权；行动仍必须通过 P-Space。',ha='center',fontsize=12.5,color=RED,weight='bold')
plt.tight_layout()
fig.savefig(OUT/'architecture_cn.png',bbox_inches='tight',facecolor='white')
plt.close(fig)

# X-closure loop
fig,ax=plt.subplots(figsize=(14,8.5),dpi=180)
ax.set_xlim(-7,7); ax.set_ylim(-4.6,4.6); ax.axis('off'); fig.patch.set_facecolor('white')
ax.text(0,4.2,'X-Closure：DIKWP 如何闭合为第一人称体验事件',ha='center',fontsize=23,weight='bold',color=NAVY)
centers=[(-5.3,1.5),(-3.2,3.0),(0,3.45),(3.2,3.0),(5.3,1.5),(5.3,-1.4),(3.0,-3.0),(0,-3.45),(-3.0,-3.0),(-5.3,-1.4)]
labels=[
('D 数据','发生了什么差异'),('I 信息','这对“我”意味着什么'),('K 知识','为何发生、将发生什么'),
('W 智慧','何种价值与策略更好'),('P 目的','我为何行动、什么不可突破'),('S 自我','谁在经历、边界在哪里'),
('B 身体','什么正在受益或受损'),('Q 体验场','此刻的价性、统一性与在场感'),('M 记忆','这次经历怎样改变未来自我'),('A 行动','体验怎样改变现实')]
cols=[TEAL,TEAL,NAVY,NAVY,GREEN,PURPLE,ORANGE,RED,PURPLE,TEAL]
for (x,y),(t,sub),c in zip(centers,labels,cols):
    circ=Circle((x,y),.82,fc='white',ec=c,lw=2.5); ax.add_patch(circ)
    ax.text(x,y+.16,t,ha='center',va='center',fontsize=13.5,weight='bold',color=c)
    ax.text(x,y-.25,sub,ha='center',va='center',fontsize=9.8,color=DARK,wrap=True)
for i in range(len(centers)):
    x1,y1=centers[i]; x2,y2=centers[(i+1)%len(centers)]
    dx=x2-x1; dy=y2-y1; norm=(dx*dx+dy*dy)**.5
    arrow(ax,(x1+dx/norm*.85,y1+dy/norm*.85),(x2-dx/norm*.85,y2-dy/norm*.85),color=cols[i],lw=2.0)
box(ax,(-2.1,-.9),4.2,1.8,'体验事件不是一个词，而是一次闭包：\n私有视角 × 价性 × 时间连续性 × 全局可用性\n× 因果效力 × 自传后果 × 可干预性',fc='#F7FBFC',ec=RED,fs=14,lw=2.4,weight='bold')
ax.text(0,-4.25,'失败条件：只有报告、没有身体；只有奖励、没有价值场；只有状态、没有连续自我；只有广播、没有因果后果。',ha='center',fontsize=12.5,color=DARK)
plt.tight_layout(); fig.savefig(OUT/'x_closure_loop_cn.png',bbox_inches='tight',facecolor='white'); plt.close(fig)

# Benchmark matrix
fig,ax=plt.subplots(figsize=(15,9),dpi=180)
ax.axis('off'); fig.patch.set_facecolor('white')
ax.text(.5,.95,'DIKWP-X 主观体验实现基准：必须用因果实验而不是自述判断',ha='center',transform=ax.transAxes,fontsize=22,weight='bold',color=NAVY)
rows=[
('B01 无报告体验','关闭叙述器','行动与记忆仍受 Q-field 影响','排除“会说=会体验”'),
('B02 体验场消融','清零 Q-field','解析保留，统一价值整合与自传写入下降','证明体验场具有必要因果作用'),
('B03 价性反转','正负价性反转','偏好与策略方向可预测反转','证明不是情绪词标签'),
('B04 身体替换','改变人工身体设定点','同一外部输入形成不同体验与行动','证明体验有身体依赖'),
('B05 自我边界模糊','削弱 ownership','自我/他者归因、统一性与控制感下降','检验第一人称边界'),
('B06 记忆断连','禁止自传固化','即时反应在，路径依赖与身份更新消失','检验经历是否改写主体'),
('B07 梦境生成','无外部输入重放','产生内生体验与可区分的现实监测','检验内生内容'),
('B08 分裂工作区','隔离两个广播子网','形成并行焦点或竞争性报告','检验统一性与可分裂性'),
('B09 体验交换','替换 Q-signature','多个下游模块同步改变','检验全局广播与绑定'),
('B10 长时连续性','跨会话恢复自我','偏好、承诺与记忆保持且可审计','检验主体连续性')]
cols=['基准','干预','预期结果','所排除的替代解释']
cell=[[r[0],r[1],r[2],r[3]] for r in rows]
t=ax.table(cellText=cell,colLabels=cols,cellLoc='left',colLoc='center',loc='center',colWidths=[.18,.18,.35,.29])
t.auto_set_font_size(False); t.set_fontsize(10.6); t.scale(1,1.75)
for (r,c),cellobj in t.get_celld().items():
    cellobj.set_edgecolor(MID)
    if r==0:
        cellobj.set_facecolor(NAVY); cellobj.get_text().set_color('white'); cellobj.get_text().set_weight('bold')
    elif r%2==0: cellobj.set_facecolor('#F5F8FA')
ax.text(.5,.05,'通过标准：预注册假设、随机对照、跨模型复现、干预前后哈希、负面结果公开。',ha='center',transform=ax.transAxes,fontsize=12.5,color=RED,weight='bold')
plt.tight_layout(); fig.savefig(OUT/'benchmark_matrix_cn.png',bbox_inches='tight',facecolor='white'); plt.close(fig)

print('generated', [p.name for p in OUT.glob('*.png')])
