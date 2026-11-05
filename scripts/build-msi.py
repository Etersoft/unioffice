#!/usr/bin/env python3
"""Package PE32 or PE64 Excel DLLs without modifying the checked-in MSI tables."""
import argparse
import os
from pathlib import Path
import shlex
import shutil
import struct
import subprocess
import tempfile
import uuid


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
        signature = handle.read(6)
        if signature == b'PE\0\0\x4c\x01':
            bits = 32
        elif signature == b'PE\0\0\x64\x86':
            bits = 64
        else:
            parser.error('installer supports only Windows PE i386 or AMD64 DLLs')
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
        if bits == 64:
            # Distinct identities allow independent install/uninstall of x86 and x64.
            component = work / 'idt' / 'Component.idt'
            rows = component.read_bytes().splitlines()
            for index in range(3, len(rows)):
                columns = rows[index].split(b'\t')
                columns[1] = b'{F2E409F1-18B1-5FAD-ADAD-184D12608F75}'
                columns[3] = str(int(columns[3]) | 256).encode('ascii')
                rows[index] = b'\t'.join(columns)
            component.write_bytes(b'\r\n'.join(rows) + b'\r\n')
            for name in ('Directory.idt', 'CustomAction.idt'):
                table = work / 'idt' / name
                table.write_bytes(table.read_bytes().replace(b'SystemFolder', b'System64Folder'))
            properties = work / 'idt' / 'Property.idt'
            data = properties.read_bytes()
            data = data.replace(b'{643315BA-DEF8-4CBE-AB26-C54B71535DFD}', b'{4C7393FD-6C4A-5B3F-8EC2-B7AD16262864}')
            data = data.replace(b'{CD87087B-70BE-4BE2-B7F3-987AB335659D}', b'{AB670464-4388-5009-B76D-7D4FA37EE2CF}')
            data = data.replace(b'ProductName\tUniOffice@Etersoft v0.5', b'ProductName\tUniOffice@Etersoft v0.5 (x64)')
            properties.write_bytes(data)
        summary = work / '_SummaryInformation.idt'
        package_code = '{' + str(uuid.uuid4()).upper() + '}'
        summary.write_text('PropertyId\tValue\ni2\tl255\n_SummaryInformation\tPropertyId\n'
                           + f'7\t{"x64" if bits == 64 else "Intel"};1033\n'
                           + f'9\t{package_code}\n14\t{200 if bits == 64 else 100}\n',
                           encoding='ascii', newline='\r\n')
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
        subprocess.run(msidb + ['-f', windows(work), '-i', '_SummaryInformation.idt'], cwd=work, check=True, env=wine_env)
        subprocess.run(msidb + ['-a', windows(cabinet)], cwd=work, check=True, env=wine_env)
        # Publish only a completely built package, on the output filesystem.
        staged = output.with_suffix('.msi.tmp')
        shutil.copy2(package, staged)
        staged.replace(output)
    print(f'PE{bits} MSI: {output} (package {package_code})')


if __name__ == '__main__':
    main()
