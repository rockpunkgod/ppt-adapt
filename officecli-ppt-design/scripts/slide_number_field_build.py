import sys,argparse
from pathlib import Path
from build_components import Deck
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--out',required=True);a=p.parse_args()
d=Deck(a.out,a.officecli,static=True)
for name in ['PLAN','DESIGN','DELIVER']:
 s=d.slide('F6F8FC');d.heading(s,name,'Master slide-number field / duplicate and reorder test')
 d.shape(s,'own-content',2,7,22,3,'Content belongs to this slide.',size=28)
d.finish()
