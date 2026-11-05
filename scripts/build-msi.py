#!/usr/bin/env python3
"""Package the legacy 32-bit installer without modifying its checked-in tables."""
import argparse
import os
from pathlib import Path
import shlex
import shutil
import struct
import subprocess
import tempfile


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--excel-build', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args()
    source = args.excel_build.resolve()
    output = args.output.resolve()
    root = Path(__file__).resolve().parent.parent
    dll = source / 'unioffice_excel.dll'
    with dll.open('rb') as handle:
        if handle.read(2) != b'MZ':
            parser.error('installer requires a Windows PE DLL')
        handle.seek(0x3c)
        offset = struct.unpack('<I', handle.read(4))[0]
        handle.seek(offset)
        if handle.read(6) != b'PE\0\0\x4c\x01':
            parser.error('legacy MSI supports only PE32 i386; a PE64 template is required')
    wine = shlex.split(os.environ.get('WINE', 'wine'))
    # IDT data is CP1251; Wine's importer decodes using the process ACP.
    wine_env = dict(os.environ, LC_ALL='ru_RU.UTF-8')
    winepath = wine + ['winepath.exe']

    def windows(path):
        return subprocess.check_output(winepath + ['-w', str(path)], text=True, env=wine_env).strip()

    scratch = Path.home() / 'tmp' / 'claude'
    scratch.mkdir(parents=True, exist_ok=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='unioffice-msi-', dir=scratch) as directory:
        work = Path(directory)
        shutil.copytree(root / 'MSI' / 'idt', work / 'idt')
        # Wine's IDT importer needs the database code page set separately.
        for archive in (work / 'idt').glob('*.idt'):
            lines = archive.read_bytes().splitlines()
            fields = lines[2].split(b'\t')
            if fields[0] == b'1251':
                lines[2] = b'\t'.join(fields[1:])
                archive.write_bytes(b'\r\n'.join(lines) + b'\r\n')
        codepage = work / '_ForceCodepage.idt'
        codepage.write_bytes(b'\r\n\r\n1251\t_ForceCodepage\r\n')
        payload = work / 'payload'
        payload.mkdir()
        for name in ('unioffice_excel.tlb', 'unioffice_excel.dll'):
            shutil.copy2(source / name, payload / name)
        table = work / 'idt' / 'File.idt'
        rows = table.read_bytes().splitlines()
        for index in range(3, len(rows)):
            columns = rows[index].split(b'\t')
            columns[3] = str((payload / columns[0].decode('ascii')).stat().st_size).encode('ascii')
            rows[index] = b'\t'.join(columns)
        table.write_bytes(b'\r\n'.join(rows) + b'\r\n')
        cabinet = work / 'ooffice.cab'
        subprocess.run(wine + [str(root / 'MSI/bin/CABARC.EXE'), 'n', windows(cabinet),
                               'unioffice_excel.tlb', 'unioffice_excel.dll'], cwd=payload, check=True, env=wine_env)
        package = work / 'unioffice.msi'
        shutil.copy2(root / 'MSI/bin/template.msi', package)
        msidb = wine + [str(root / 'MSI/bin/MsiDb.exe'), '-d', windows(package)]
        subprocess.run(msidb + ['-f', windows(work), '-i', '_ForceCodepage.idt'], cwd=work, check=True, env=wine_env)
        subprocess.run(msidb + ['-f', windows(work / 'idt'), '-i', '*'], cwd=work, check=True, env=wine_env)
        subprocess.run(msidb + ['-a', windows(cabinet)], cwd=work, check=True, env=wine_env)
        # Publish only a completely built package, on the output filesystem.
        staged = output.with_suffix('.msi.tmp')
        shutil.copy2(package, staged)
        staged.replace(output)
    print(output)


if __name__ == '__main__':
    main()
