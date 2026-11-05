#!/usr/bin/env python3
"""Exercise MSI32/MSI64 lifecycle in a disposable prefix with both installed."""
import os
from pathlib import Path
import subprocess


def main():
    prefix = Path(os.environ['WINEPREFIX']).resolve()
    scratch = (Path.home() / 'tmp/claude').resolve()
    if not prefix.is_relative_to(scratch):
        raise SystemExit('Use a disposable WINEPREFIX under ~/tmp/claude')
    root = Path(__file__).resolve().parent.parent
    loader = os.environ.get('WINE', 'wine')
    log_dir = prefix.parent / 'coexistence-logs'
    log_dir.mkdir(exist_ok=True)
    products = {32: '{643315BA-DEF8-4CBE-AB26-C54B71535DFD}',
                64: '{4C7393FD-6C4A-5B3F-8EC2-B7AD16262864}'}
    sources = {32: root / 'unioffice_excel/build/mingw-32',
               64: root / 'unioffice_excel/build/winegcc-64'}
    installed = {32: prefix / 'drive_c/windows/syswow64',
                 64: prefix / 'drive_c/windows/system32'}
    packages = {32: root / 'MSI/build/mingw-32/unioffice.msi',
                64: root / 'MSI/build/winegcc-64/unioffice.msi'}

    def windows(path):
        return 'Z:' + str(path.resolve()).replace('/', '\\')

    def check_payload(bits, present=True):
        for name in ('unioffice_excel.dll', 'unioffice_excel.tlb'):
            path = installed[bits] / name
            if present:
                if path.read_bytes() != (sources[bits] / name).read_bytes():
                    raise RuntimeError(f'PE{bits} payload mismatch: {name}')
            elif path.exists():
                raise RuntimeError(f'PE{bits} payload remains: {name}')
        print(f'PE{bits} payload {"matches" if present else "removed"}', flush=True)

    def probe(bits, label):
        env = dict(os.environ, CLIENT_BITS=str(bits))
        with (log_dir / f'{label}.log').open('wb') as log:
            subprocess.run(['sh', str(root / 'tests/run-openoffice-backend.sh'),
                            'excel-backend.vbs', 'UniOffice integration: PASS'],
                           env=env, cwd=root, stdout=log, stderr=subprocess.STDOUT,
                           timeout=150, check=True)
        print(f'{label}: Automation PASS', flush=True)

    def msi(arguments, label):
        subprocess.run([loader, 'msiexec'] + arguments + ['/qn', '/l*v',
                       windows(log_dir / f'{label}.log')], timeout=60, check=True)
        print(f'{label}: MSI PASS', flush=True)

    for bits in (32, 64):
        check_payload(bits)
        probe(bits, f'initial-{bits}')
    msi(['/i', windows(packages[64]), 'REINSTALL=ALL', 'REINSTALLMODE=amusv'], 'repair-64')
    check_payload(64)
    probe(64, 'after-repair-64')
    msi(['/x', products[64]], 'remove-64-first')
    check_payload(64, False)
    check_payload(32)
    probe(32, 'surviving-32')
    msi(['/i', windows(packages[64])], 'reinstall-64')
    check_payload(64)
    msi(['/x', products[32]], 'remove-32-first')
    check_payload(32, False)
    check_payload(64)
    probe(64, 'surviving-64')
    msi(['/x', products[64]], 'remove-last-64')
    check_payload(64, False)
    result = subprocess.run([loader, 'reg', 'query', r'HKCR\Excel.Application\CLSID'],
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=15)
    if result.returncode != 1:
        raise RuntimeError(f'Expected missing shared ProgID; reg query returned {result.returncode}')
    print(f'MSI coexistence: PASS; logs: {log_dir}', flush=True)


if __name__ == '__main__':
    main()
