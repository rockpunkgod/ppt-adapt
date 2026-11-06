"""Add an editable trigger lab using OfficeCLI before COM motion refinement."""
import argparse,shutil
from pathlib import Path
from build_visual_animation import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',default='officecli');p.add_argument('--artifacts',type=Path,required=True);a=p.parse_args()
root=a.artifacts;out=root/'refinement-base-v3.pptx'
if out.exists():p.error('Refusing to overwrite '+str(out))
shutil.copy2(root/'animation-lab.pptx',out)
d=Deck.__new__(Deck);d.exe=a.officecli;d.path=out;d.log=[];d.n=5
s=d.slide('171C23');ops=d.title(s,'点击标记，显示说明','指定形状触发 · 点击显示')
ops += [d.shape(s,'trigger-pin',7,8,2.6,2.6,fill='D4B46C',preset='ellipse'),d.shape(s,'trigger-label',12,7.4,17,5.5,'只有点击左侧标记\n才会显示这条说明',size=28,fill='355B67',preset='roundRect',margin='.7cm',valign='center'),d.shape(s,'instruction',4,13.5,26,2.5,'点击空白不应触发目标说明',size=22,color='C7C0D5',align='center'),d.note(s,'来源Top10 07:43–09:43。OfficeCLI建立基础对象；COM的InteractiveSequences添加指定pin触发的Fade入口；开关退出仍待验证，必须原生点击核验。')]
d.batch(ops);d.call('set',str(out),'/slide[5]/shape[@name=water]','--prop','y=9.5cm');d.finish();print(out)
