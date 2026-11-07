from pathlib import Path
import argparse
from scene import Deck
from showcase import Showcase
parser=argparse.ArgumentParser()
parser.add_argument('--work',type=Path,required=True)
parser.add_argument('--officecli',required=True)
args=parser.parse_args()
d=Deck(args.work/'main.pptx',args.officecli);a=Showcase(d)
a.cover();a.brief();a.story()
for state in range(3):a.roam(state)
a.range(args.work);a.handoff();a.end()
a.flush()
for s in range(1,10):
 d.run('set',f'/slide[{s}]',props={'transition':('morph-1800' if s in [2,5,6] else 'fade-850'),'advanceClick':True})
a.finish()
print('Nine-page PPT Adapt city-scene presentation rebuilt from native objects.')
