"""Rebuild the current nine-slide example from local source, not a deck copy.

The Agent handles arbitrary presentation briefs; this command reproduces one
specific example. Python stdlib, existing OfficeCLI and Windows PowerPoint only.
"""
import argparse
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from zipfile import ZipFile
import xml.etree.ElementTree as ET

SKILL = Path(__file__).resolve().parents[1]
SOURCE = SKILL / 'scripts/demo_source'


def run(command):
    subprocess.run(list(map(str, command)), check=True)


def resolve_officecli(explicit):
    candidates = ([explicit] if explicit else [
        os.environ.get('PPT_ADAPT_OFFICECLI'),
        SKILL.parent / 'tools/officecli.exe',
        Path.cwd() / 'tools/officecli.exe',
        shutil.which('officecli'), shutil.which('officecli.exe'),
    ])
    tried = set()
    for candidate in filter(None, candidates):
        candidate = str(candidate)
        if candidate in tried:
            continue
        tried.add(candidate)
        if not Path(candidate).is_file():
            continue
        # Some installations report a version but fail to create a document.
        with tempfile.TemporaryDirectory(prefix='ppt-adapt-probe-') as folder:
            probe = Path(folder) / 'probe.pptx'
            result = subprocess.run([candidate, 'create', str(probe), '--json'],
                                    capture_output=True, timeout=45)
            if result.returncode == 0 and probe.is_file():
                subprocess.run([candidate, 'close', str(probe), '--json'],
                               capture_output=True, timeout=45, check=True)
                return str(Path(candidate).resolve())
    raise RuntimeError('No working OfficeCLI found. Agent: reuse the local installation '
                       'or prepare OfficeCLI, then pass --officecli internally.')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', required=True, type=Path)
    parser.add_argument('--officecli')
    parser.add_argument('--skip-video', action='store_true')
    parser.add_argument('--keep-work', action='store_true')
    args = parser.parse_args()
    if sys.platform != 'win32':
        parser.error('This native rebuild requires Windows desktop PowerPoint.')
    output = args.output.resolve()
    if output.exists():
        raise FileExistsError(f'Refusing to overwrite: {output}')
    exe = resolve_officecli(args.officecli)
    powershell = shutil.which('powershell.exe')
    if not powershell:
        raise RuntimeError('Windows PowerShell is required for native PowerPoint.')
    def ps(script, *params):
        run([powershell, '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', script, *params])
    # Check PowerPoint before generating anything. Do not quit the shared app.
    run([powershell, '-NoProfile', '-Command',
         "$ErrorActionPreference='Stop';$a=New-Object -ComObject PowerPoint.Application;"
         "Write-Output ('PowerPoint '+$a.Version)"])
    output.mkdir(parents=True)
    work = output / '.build-work'
    work.mkdir()
    print('1/4 Creating original design directions...', flush=True)
    run([sys.executable, '-X', 'utf8', SOURCE / 'pilot.py', '--work', work, '--officecli', exe])
    ps(SKILL / 'scripts/layout_story_native.ps1', '-InputDeck', work / 'directions.pptx',
       '-OutputDirectory', work / 'native-directions')
    print('2/4 Building the nine-slide presentation...', flush=True)
    run([sys.executable, '-X', 'utf8', SOURCE / 'build.py', '--work', work, '--officecli', exe])
    print('3/4 Finishing the native paper and document scenes...', flush=True)
    deck = output / 'presentation.pptx'
    ps(SOURCE / 'finish.ps1', '-Work', work, '-OutputDeck', deck)
    ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    with ZipFile(deck) as archive:
        root = ET.fromstring(archive.read('ppt/presentation.xml'))
        if len(root.findall('p:sldIdLst/p:sldId', ns)) != 9:
            raise RuntimeError('Expected nine rebuilt slides.')
    print('4/4 Exporting previews with native PowerPoint...', flush=True)
    ps(SKILL / 'scripts/render_preview.ps1', '-InputDeck', deck,
       '-OutputDirectory', output, *([] if args.skip_video else ['-Video']))
    if not args.keep_work:
        # Both directories were created by this invocation under the fresh output.
        shutil.rmtree(work)
        shutil.rmtree(output / '.preview-pages')
    print(f'Ready: {deck}\nPreview: {output / "preview.jpg"}', flush=True)
    if not args.skip_video:
        print(f'Native export: {output / "preview.mp4"}')
        print('Video is a native export preview, not a screen-recorded playback verification.')


if __name__ == '__main__':
    main()
