"""Illustrate local ordering only; dimensions do not represent an actual joint gap."""
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
plt.rcParams.update({'font.family':['Yu Gothic','Meiryo','DejaVu Sans'],'font.size':11})
fig,ax=plt.subplots(figsize=(10,3.4))
ax.add_patch(Rectangle((0,.2),2,.72,facecolor='#dce8ee',edgecolor='#91a9b5'))
ax.add_patch(Rectangle((0,-.92),2,.72,facecolor='#e5ead8',edgecolor='#a1aa8a'))
ax.text(1,.74,'主面AのQ8要素',ha='center',va='center',fontweight='bold')
ax.text(1,-.74,'従面BのQ8要素',ha='center',va='center',fontweight='bold')
for y,labels,color in [(.18,['1','3（中点）','2'],'#286787'),(-.18,['4','6（中点）','5'],'#607640')]:
 ax.plot([0,2],[y,y],color=color,lw=2)
 for x,label in zip([0,1,2],labels):
  ax.plot(x,y,'o',color=color,ms=5)
  ax.text(x,y+(.16 if y>0 else -.16),label,ha='center',va='center',color=color)
for x in [0,1,2]:ax.plot([x,x],[-.14,.14],'--',color='#999',lw=1)
ax.annotate('',xy=(1.65,.42),xytext=(.35,.42),arrowprops={'arrowstyle':'-|>','color':'#286787','lw':1.8})
ax.text(1,.52,'Q8辺の走査：1 → 3 → 2',ha='center',color='#286787')
ax.annotate('',xy=(.35,-.42),xytext=(1.65,-.42),arrowprops={'arrowstyle':'-|>','color':'#607640','lw':1.8})
ax.text(1,-.53,'Q8辺の走査：5 → 6 → 4',ha='center',color='#607640')
ax.annotate('',xy=(2.26,-.16),xytext=(2.26,.55),arrowprops={'arrowstyle':'-|>','color':'#a2484e','lw':2})
ax.text(2.36,.32,'主面の外向き法線 n\n（1→2の右側）',va='center',color='#a2484e')
ax.set(xlim=(-.2,3.45),ylim=(-1.07,1.08));ax.axis('off')
fig.suptitle('水圧JOINTの向き：登録順とQ8辺の走査順を区別する',fontsize=12,fontweight='bold',y=.98)
fig.text(.5,.025,'局所位置の模式図。実入力は厚さ0：1/4・2/5・3/6を同一座標に置く。',ha='center',fontsize=10)
fig.subplots_adjust(left=.03,right=.99,bottom=.13,top=.89)
fig.savefig(Path(__file__).resolve().parents[1]/'docs/manuals/images/groundwater/joint_orientation.png',dpi=180,facecolor='white')
plt.close(fig)
