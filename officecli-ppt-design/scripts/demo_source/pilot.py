from pathlib import Path
import argparse
from scene import Scene,Deck
parser=argparse.ArgumentParser()
parser.add_argument('--work',type=Path,required=True)
parser.add_argument('--officecli',required=True)
args=parser.parse_args()
d=Deck(args.work/'directions.pptx',args.officecli);a=Scene(d)
a.cover();a.brief();a.alternate(0);a.alternate(1)
for state in range(3):a.knowledge(state)
a.finish()
