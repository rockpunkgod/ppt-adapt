from pathlib import Path
import argparse
from scene import Deck
from showcase import Showcase
parser=argparse.ArgumentParser()
parser.add_argument('--work',type=Path,required=True)
parser.add_argument('--officecli',required=True)
args=parser.parse_args()
d=Deck(args.work/'directions.pptx',args.officecli);a=Showcase(d)
for state in range(3):a.roam(state)
a.cover();a.brief();a.alternate(0);a.alternate(1)
a.flush()
for s in (2,3):d.run('set',f'/slide[{s}]',props={'transition':'morph-1800','advanceClick':True})
a.finish()
