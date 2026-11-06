#!/usr/bin/env python3
"""Run the seven legacy Excel VBS tests in a disposable Wine prefix."""
import argparse
import os
from pathlib import Path
import re
import subprocess
import tempfile
import time


# These are the assertions executed by the checked-in, noninteractive scripts.
EXPECTED_SUCCESS = {1: 32, 2: 13, 3: 22, 4: 33, 5: 32, 6: 11, 7: 13}
# The legacy scripts increment success once without logging an assertion here.
COUNTER_EXCESS = {4: 1, 6: 1}


def result_errors(number, totals, report, returncode):
    """Require complete assertion evidence, independently of cscript's exit code."""
    errors = []
    expected = EXPECTED_SUCCESS[number]
    if returncode:
        errors.append(f'script exited with status {returncode}')
    if len(totals) != 2 or any(not re.fullmatch(r'[0-9]+', value) for value in totals):
        errors.append('missing or malformed completion counters')
    elif (int(totals[0]), int(totals[1])) != (expected + COUNTER_EXCESS.get(number, 0), 0):
        errors.append(f'expected completion counters {expected + COUNTER_EXCESS.get(number, 0)} 0')
    success = report.count(b'[SUCCESS]')
    failed = report.count(b'[FAILED]')
    if success != expected:
        errors.append(f'expected {expected} successful assertions, got {success}')
    if failed:
        errors.append(f'report contains {failed} failed assertion(s)')
    return errors


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--tests', nargs='+', type=int, choices=range(1, 8), default=list(range(1, 8)))
    parser.add_argument('--bits', type=int, choices=(32, 64), default=32)
    args = parser.parse_args()
    script_host = r'C:\windows\system32\cscript.exe' if args.bits == 64 else r'C:\windows\syswow64\cscript.exe'
    if not os.environ.get('WINEPREFIX'):
        raise SystemExit('Set WINEPREFIX to a disposable OpenOffice + UniOffice prefix')
    root = Path(__file__).resolve().parent.parent
    scratch = Path.home() / 'tmp' / 'claude'
    scratch.mkdir(parents=True, exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix='unioffice-excel-suite-', dir=scratch))
    print(f'Results: {work}', flush=True)
    print(f'Client bits: {args.bits}; prefix: {os.environ["WINEPREFIX"]}', flush=True)
    loader = os.environ.get('WINE', 'wine')
    failures = 0
    for number in args.tests:
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
                time.sleep(3)
                with (test_dir / 'ready.log').open('wb') as ready_log:
                    ready = subprocess.run([loader, script_host, '//nologo',
                                            'Z:' + str(root / 'tests/openoffice-ready.vbs')],
                                           stdout=ready_log, stderr=subprocess.STDOUT, timeout=60, check=False)
                if ready.returncode or b'OpenOffice readiness: PASS' not in (test_dir / 'ready.log').read_bytes():
                    print(f'test_{number}: backend not ready; inspect ready.log', flush=True)
                    failures += 1
                    break
                with (test_dir / 'script.log').open('wb') as script_log:
                    result = subprocess.run([loader, script_host, '//nologo',
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
        try:
            totals = done.read_text(encoding='ascii').split() if done.exists() else []
        except (OSError, UnicodeError) as error:
            print(f'test_{number}: cannot read completion counters: {error}', flush=True)
            totals = []
        print(f'test_{number}: totals={totals}; markers={report.count(b"[SUCCESS]")} success, '
              f'{report.count(b"[FAILED]")} failed', flush=True)
        for line in report.splitlines():
            if b'[FAILED]' in line:
                print(line.decode('cp1251', errors='replace'), flush=True)
        errors = result_errors(number, totals, report, result.returncode)
        if errors:
            for error in errors:
                print(f'test_{number}: FAILED: {error}', flush=True)
            failures += 1
    return 1 if failures else 0


if __name__ == '__main__':
    raise SystemExit(main())
