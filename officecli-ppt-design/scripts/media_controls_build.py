import argparse,subprocess,sys
from pathlib import Path
p=argparse.ArgumentParser();p.add_argument('--officecli',required=True);p.add_argument('--ffmpeg',required=True);p.add_argument('--outdir',type=Path,required=True);a=p.parse_args()
if a.outdir.exists():raise FileExistsError(a.outdir)
a.outdir.mkdir(parents=True);here=Path(__file__).resolve().parent
subprocess.run([sys.executable,str(here/'media_controls_assets.py'),'--ffmpeg',a.ffmpeg,'--outdir',str(a.outdir/'assets')],check=True)
subprocess.run([sys.executable,str(here/'media_controls_layout.py'),'--officecli',a.officecli,'--assets',str(a.outdir/'assets'),'--outdir',str(a.outdir/'base')],check=True)
