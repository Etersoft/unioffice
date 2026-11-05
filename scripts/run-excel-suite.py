#!/usr/bin/env python3
"""Run the seven legacy Excel VBS tests in a disposable Wine prefix."""
import os
from pathlib import Path
import re
import subprocess
import tempfile
import time


def main():
    if not os.environ.get('WINEPREFIX'):
        raise SystemExit('Set WINEPREFIX to a disposable OpenOffice + UniOffice prefix')
    root = Path(__file__).resolve().parent.parent
    scratch = Path.home() / 'tmp' / 'claude'
    scratch.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix='unioffice-excel-suite-', dir=scratch))
    print(f'Results: {work}', flush=True)
    loader = os.environ.get('WINE', 'wine')
    failures = 0
    for number in range(1, 8):
        source = root / f'tests/Excel/test_{number}.vbs'
        test_dir = work / f'test_{number}'
        test_dir.mkdir()
        # Preserve assertions/encoding, replace only completion dialogs. A separate
        # directory prevents the legacy code from launching a host of another bitness.
        data = re.sub(rb'(?m)^([ \t]*)MsgBox[^\r\n]*',
                      lambda match: match[1] + b'Set completedFile = fso.CreateTextFile("suite.done", True)\n'
                      + match[1] + b'completedFile.WriteLine gsuccess & " " & gfailed\n'
                      + match[1] + b'completedFile.Close', source.read_bytes())
        (test_dir / source.name).write_bytes(data)
        (test_dir / 'otchet.txt').write_bytes(b'')
        with (test_dir / 'office.log').open('wb') as office_log:
            office = subprocess.Popen([loader, r'C:\Program Files (x86)\OpenOffice 4\program\soffice.exe',
                                       '-headless', '-nofirststartwizard', '-norestore', '-nodefault'],
                                      stdout=office_log, stderr=subprocess.STDOUT)
            try:
                time.sleep(5)
                with (test_dir / 'script.log').open('wb') as script_log:
                    subprocess.run([loader, r'C:\windows\syswow64\cscript.exe', '//nologo',
                                    '/E:vbscript', source.name, str(number), '0', '0'],
                                   cwd=test_dir, stdout=script_log, stderr=subprocess.STDOUT,
                                   timeout=60, check=False)
            except subprocess.TimeoutExpired:
                print(f'test_{number}: TIMEOUT', flush=True)
                failures += 1
                # Stop here: a stranded COM call could contaminate following tests.
                break
            finally:
                office.terminate()
        report = (test_dir / 'otchet.txt').read_bytes()
        done = test_dir / 'suite.done'
        totals = done.read_text().split() if done.exists() else []
        print(f'test_{number}: totals={totals}; markers={report.count(b"[SUCCESS]")} success, '
              f'{report.count(b"[FAILED]")} failed', flush=True)
        for line in report.splitlines():
            if b'[FAILED]' in line:
                print(line.decode('cp1251', errors='replace'), flush=True)
        if len(totals) != 2 or totals[1] != '0':
            failures += 1
    return 1 if failures else 0


if __name__ == '__main__':
    raise SystemExit(main())
