"""Compose actual external Blast/Overlay assets with editable native typography."""
import argparse
from pathlib import Path
from build_components import Deck

def main():
    p=argparse.ArgumentParser(description=__doc__);p.add_argument('--officecli',required=True);p.add_argument('--assets',type=Path,required=True);p.add_argument('--out',type=Path,required=True);a=p.parse_args()
    for n in ['channels.png','blast-right.png','overlay-two-photopea.png']:
        if not (a.assets/n).is_file():raise FileNotFoundError(a.assets/n)
    d=Deck(a.out,a.officecli)
    def picture(s,name,file,x,y,w,h):d.run('add',f'/slide[{s}]','--type','picture',props=dict(name=name,src=str((a.assets/file).resolve()),x=f'{x}cm',y=f'{y}cm',width=f'{w}cm',height=f'{h}cm'))
    for i,(title,caption) in enumerate([('分离通道，建立偏移','R 向右、GB 向左；原图仍保留。'),('只让局部产生拖尾','透明条带副本 · Blast · From the Right'),('让颜色形成方向','上左洋红 / 下右青色 · 两层分别 Overlay'),('保留信号，释放噪声','SIGNAL / A visual experiment')],1):
        s=d.slide('10141E')
        if i<4:
            d.shape(s,'heading',1.6,1,30.6,1.6,title,size=30,bold=True,color='F4F0E8');d.shape(s,'caption',1.65,2.7,30,1.1,caption,size=17,color='A9B8C5')
            if i<3:picture(s,'channel-base','channels.png',4.1,4.7,25.6,14.4)
            if i==2:picture(s,'actual-blast-strip-copy','blast-right.png',4.1,4.7,25.6,14.4)
            if i==3:picture(s,'external-overlay-composite','overlay-two-photopea.png',4.1,4.7,25.6,14.4)
        else:
            picture(s,'external-overlay-composite','overlay-two-photopea.png',9,1.8,24,13.5)
            d.shape(s,'eyebrow',1.7,1.2,20,1,'SIGNAL / 01',size=14,color='63DCE8')
            d.shape(s,'main-title',1.7,6.1,26,6.2,'KEEP THE\nSIGNAL.',font='Arial',size=64,bold=True,color='F4F0E8')
            d.shape(s,'closing',1.8,15.6,29,1.2,title,size=26,color='F4F0E8');d.shape(s,'footer',1.8,17.4,29,.8,'从清晰内容出发，让视觉变化服务于表达。',size=15,color='A9B8C5')
        d.notes(s,'Original channel fixture, reused band-copy mechanism. Actual Photopea Wind/Blast From the Right on transparent strips; independent magenta then cyan layers, each Overlay100%. external-render image effects; native text and separate page2 image objects remain editable. Not Adobe algorithm pixel equivalence, not animated glitch and not native PowerPoint blend filtering. New assets and combinations need actual playback.')
    d.run('close');d.run('validate');print(a.out)

if __name__=='__main__':main()
