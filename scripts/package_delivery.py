"""Package tracked project files from local disk, without downloading anything."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
from zipfile import ZIP_DEFLATED, ZipFile


def main():
    root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--out', type=Path, default=root / 'ppt-skills.zip')
    args = parser.parse_args()
    output = args.out.resolve()
    names = subprocess.check_output(
        ['git', '-C', str(root), 'ls-files', '-z'],
    ).decode('utf-8').split('\0')
    if any((root / name).resolve() == output for name in names if name and name != 'ppt-skills.zip'):
        raise ValueError('Output would overwrite a project file.')
    files = []
    for name in sorted(set(filter(None, names))):
        path = root / name
        if name == 'ppt-skills.zip' or path.resolve() == output:
            continue
        if any(part in {'.git', '.work', 'tools', '__pycache__'} for part in path.relative_to(root).parts):
            raise ValueError(f'Private/runtime file unexpectedly tracked: {name}')
        if path.is_symlink() or not path.is_file():
            raise ValueError(f'Expected a local regular file: {name}')
        files.append((name, path))
    required = {
        'README.md', 'requirements.txt', 'scripts/package_delivery.py',
        'officecli-ppt-design/SKILL.md', 'ppt-director/SKILL.md',
        'ppt-composition-planner/SKILL.md',
        'examples/trace-product/trace-product.pptx',
    }
    if required.difference(name for name, _ in files):
        raise ValueError('Required delivery files are missing; stage new files before packaging.')
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output, 'w', ZIP_DEFLATED, compresslevel=8) as archive:
        for name, path in files:
            archive.write(path, name)
    with ZipFile(output) as archive:
        bad = archive.testzip()
        if bad:
            raise ValueError(f'Archive CRC failed: {bad}')
    print(json.dumps({
        'archive': str(output), 'files': len(files), 'bytes': output.stat().st_size,
        'sha256': hashlib.sha256(output.read_bytes()).hexdigest(),
    }))


if __name__ == '__main__':
    main()
