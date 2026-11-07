from pathlib import Path
import sys
import argparse
parser=argparse.ArgumentParser()
parser.add_argument('--work',type=Path,required=True)
parser.add_argument('--officecli',required=True)
args=parser.parse_args()
SKILL=Path(__file__).resolve().parents[2]
O=args.work.resolve()
sys.path.insert(0,str(SKILL/'scripts'))
from build_components import Deck
from case_pages import Author
d=Deck(O/'directions.pptx',args.officecli);a=Author(d,SKILL/'assets/demo/aura-device.png')
a.case(0);a.case(1);a.alternate(0);a.alternate(1);a.flush();d.run('close');d.run('validate');d.run('close')
