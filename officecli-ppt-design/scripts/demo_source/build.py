from pathlib import Path
import argparse
from scene import Scene,Deck
parser=argparse.ArgumentParser()
parser.add_argument('--work',type=Path,required=True)
parser.add_argument('--officecli',required=True)
args=parser.parse_args()
d=Deck(args.work/'main.pptx',args.officecli);a=Scene(d)
a.cover();a.brief();a.story()
for state in range(3):a.knowledge(state)
a.panorama(args.work);a.handoff();a.end()
a.flush()
for s in range(1,10):
 d.run('set',f'/slide[{s}]',props={'transition':('morph-1600' if s in [2,3,5,6,8,9] else 'fade-750'),'advanceClick':True})
a.finish()
print('Nine-page original PPT Adapt presentation rebuilt from native objects.')
