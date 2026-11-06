import argparse,sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parent))
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args()
d=Deck(a.out,a.officecli,static=True)
samples=[('ONE IDEA AT A TIME','Three paragraphs, one editable text box.','Ask a clear question.\nChoose evidence that answers it.\nMake the conclusion visible.',30),
 ('LETTERS IN MOTION','One title, animated letter by letter.','DESIGN WITH PURPOSE',42),
 ('WORDS IN MOTION','One sentence, animated word by word.','MAKE EVERY WORD COUNT',38)]
for title,subtitle,text,size in samples:
 s=d.slide('F3F1EB');d.heading(s,title,subtitle)
 d.shape(s,'text-build-target',2,7.0,29,7.3,text,size=size,font='Arial',color='1C6561',valign='top')
 d.shape(s,'footer',1.5,17.2,30,1,'Original sample text',size=14,color='52636D')
d.finish()
