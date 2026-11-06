# Run from 64-bit PowerShell in a disposable, interactive Windows test VM.
param(
    [ValidateSet(32, 64)][int[]]$Bits = @(32, 64),
    [ValidateSet('All', 'Suite', 'Probes')][string]$Mode = 'All',
    [string]$OutputRoot = "$env:USERPROFILE\tmp\claude"
)
$ErrorActionPreference = 'Stop'
if ([IntPtr]::Size -ne 8) { throw 'Use 64-bit PowerShell to select cscript architecture reliably' }
$root = Split-Path $PSScriptRoot -Parent
$work = Join-Path $OutputRoot ('unioffice-windows-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $work | Out-Null
$encoding = [Text.Encoding]::GetEncoding(1251)
$expected = @{1=32; 2=13; 3=22; 4=33; 5=32; 6=11; 7=13}
$extra = @{4=1; 6=1}

function Invoke-TestProcess {
    param([string]$Executable, [string[]]$Arguments, [string]$Directory, [string]$Log)
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $Executable
    $info.Arguments = ($Arguments | ForEach-Object { '"' + $_.Replace('"', '\"') + '"' }) -join ' '
    $info.WorkingDirectory = $Directory
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $info
    try {
        [void]$process.Start()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(60000)) {
            $process.Kill()
            throw "TIMEOUT: $Executable; stopping this run"
        }
        $text = $stdout.Result + $stderr.Result
        [IO.File]::WriteAllText($Log, $text)
        return [pscustomobject]@{ Code=$process.ExitCode; Text=$text }
    } finally {
        $process.Dispose()
    }
}

function Invoke-Passed {
    param([string]$Executable, [string[]]$Arguments, [string]$Marker, [string]$Name)
    $result = Invoke-TestProcess $Executable $Arguments $work (Join-Path $work "$bit-$Name.log")
    if ($result.Code -ne 0 -or -not $result.Text.Contains($Marker)) {
        throw "$Name failed (exit $($result.Code)): $($result.Text)"
    }
    Write-Host "PE${bit} ${Name}: PASS"
}

function Start-ReadyOffice {
    $office = 'C:\Program Files (x86)\OpenOffice 4\program\soffice.exe'
    Start-Process $office -ArgumentList '-headless -nofirststartwizard -norestore -nodefault' | Out-Null
    Start-Sleep -Seconds 3
    Invoke-Passed $cscript @('//nologo', (Join-Path $root 'tests\openoffice-ready.vbs')) 'OpenOffice readiness: PASS' 'ready'
}

Write-Host "Results: $work"
Write-Host "Windows: $([Environment]::OSVersion.VersionString); user: $([Security.Principal.WindowsIdentity]::GetCurrent().Name)"
try {
    foreach ($bit in $Bits) {
        if ($bit -eq 32) { $variant='mingw-32'; $system='SysWOW64' }
        else { $variant='winegcc-64'; $system='System32' }
        $cscript = Join-Path $env:windir "$system\cscript.exe"
        $build = Join-Path $root "unioffice_excel\build\$variant"
        Write-Host "Client: $cscript"
        if ($Mode -ne 'Suite') {
            $dll = Join-Path $env:windir "$system\unioffice_excel.dll"
            $tlb = Join-Path $env:windir "$system\unioffice_excel.tlb"
            # Inspect installation before any probe can register the DLL itself.
            if ($bit -eq 32) { $view=[Microsoft.Win32.RegistryView]::Registry32 }
            else { $view=[Microsoft.Win32.RegistryView]::Registry64 }
            $hive = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
            try {
                $key = $hive.OpenSubKey('SOFTWARE\Classes\CLSID\{00024500-0000-0000-C000-000000000046}\InprocServer32')
                if (-not $key) { throw "MSI registration missing for PE$bit" }
                try { $registered = [string]$key.GetValue('') } finally { $key.Dispose() }
                $registered = $registered.Trim('"')
                if ($registered -ne $dll -and $registered -ne 'unioffice_excel.dll') {
                    throw "MSI registered unexpected PE${bit} DLL: $registered"
                }
                Write-Host "PE${bit} installed COM registration: $registered"
            } finally { $hive.Dispose() }
            Invoke-Passed (Join-Path $build 'com-smoke.exe') @($dll, $tlb) 'COM smoke: 0 failure(s)' 'com-smoke'
            foreach ($probe in @(
                @('application-interface.exe', 'Application interface: 0 failure(s)'),
                @('typeinfo-lifetime.exe', 'TypeInfo lifetime: 0 failure(s)'),
                @('borders-enumerator.exe', 'Borders enumerator: 0 failure(s)'),
                @('range-enumerator.exe', 'Range enumerator: 0 failure(s)')
            )) {
                Start-ReadyOffice
                Invoke-Passed (Join-Path $build $probe[0]) @() $probe[1] $probe[0]
            }
            foreach ($probe in @(
                @('openoffice-backend.vbs', 'OpenOffice COM backend: PASS'),
                @('excel-backend.vbs', 'UniOffice integration: PASS'),
                @('font-script.vbs', 'Font script integration: PASS'),
                @('outline-levels.vbs', 'Outline integration: PASS'),
                @('border-diagonal.vbs', 'Border integration: PASS'),
                @('borders-enumeration.vbs', 'Borders enumeration integration: PASS'),
                @('border-weight.vbs', 'Border weight integration: PASS'),
                @('border-weight-mixed.vbs', 'Mixed border weight integration: PASS'),
                @('border-style-errors.vbs', 'Border style errors: PASS'),
                @('cells-indexing.vbs', 'Cells indexing: PASS'),
                @('range-item.vbs', 'Range Item: PASS'),
                @('cell-index-ownership.vbs', 'Cell index ownership: PASS'),
                @('cells-parent-lifetime.vbs', 'Cells parent lifetime: PASS'),
                @('worksheet-range-ownership.vbs', 'Worksheet Range ownership: PASS'),
                @('range-address-variable.vbs', 'Range address lifetime: PASS'),
                @('range-rows-columns.vbs', 'Range Rows and Columns: PASS'),
                @('sheet-rows-columns.vbs', 'Sheet Rows and Columns: PASS'),
                @('range-enumeration.vbs', 'Range enumeration: PASS')
            )) {
                Start-ReadyOffice
                Invoke-Passed $cscript @('//nologo', (Join-Path $root "tests\$($probe[0])")) $probe[1] $probe[0]
            }
            Invoke-Passed (Join-Path $build 'typelib-registration.exe') @($dll) '0 failure(s)' 'typelib'
        }
        if ($Mode -ne 'Probes') {
            foreach ($number in 1..7) {
                $directory = Join-Path $work "$bit-test-$number"
                New-Item -ItemType Directory $directory | Out-Null
                $data = $encoding.GetString([IO.File]::ReadAllBytes((Join-Path $root "tests\Excel\test_$number.vbs")))
                $replacement = [Text.RegularExpressions.MatchEvaluator]{ param($match)
                    $indent = $match.Groups[1].Value
                    return $indent + 'Set completedFile = fso.CreateTextFile("suite.done", True)' + "`r`n" +
                        $indent + 'completedFile.WriteLine gsuccess & " " & gfailed' + "`r`n" +
                        $indent + 'completedFile.Close'
                }
                $data = [regex]::Replace($data, '(?m)^([ \t]*)MsgBox[^\r\n]*', $replacement)
                $script = Join-Path $directory "test_$number.vbs"
                [IO.File]::WriteAllBytes($script, $encoding.GetBytes($data))
                [IO.File]::WriteAllBytes((Join-Path $directory 'otchet.txt'), [byte[]]@())
                Start-ReadyOffice
                $result = Invoke-TestProcess $cscript @('//nologo', '/E:vbscript', $script, "$number", '0', '0') $directory (Join-Path $directory 'script.log')
                $done = Join-Path $directory 'suite.done'
                if (-not (Test-Path $done)) { throw "PE${bit} test_${number}: no completion marker; $($result.Text)" }
                $totals = ([IO.File]::ReadAllText($done)).Trim() -split '\s+'
                $report = $encoding.GetString([IO.File]::ReadAllBytes((Join-Path $directory 'otchet.txt')))
                $success = [regex]::Matches($report, '\[SUCCESS\]').Count
                $failed = [regex]::Matches($report, '\[FAILED\]').Count
                $counter = $expected[$number] + $extra[$number]
                if ($result.Code -ne 0 -or $totals.Count -ne 2 -or $totals[0] -ne "$counter" -or
                    $totals[1] -ne '0' -or $success -ne $expected[$number] -or $failed -ne 0) {
                    throw "PE${bit} test_${number}: FAILED; counters=$totals markers=$success/$failed exit=$($result.Code); $report"
                }
                Write-Host "PE${bit} test_${number}: PASS; counters=$totals markers=$success/$failed"
            }
        }
    }
    [IO.File]::WriteAllText((Join-Path $work 'PASS'), 'Native Windows tests: PASS')
    Write-Host 'Native Windows tests: PASS'
} catch {
    [IO.File]::WriteAllText((Join-Path $work 'FAILED'), $_.Exception.ToString())
    Write-Host 'Native Windows tests: FAILED'
    Write-Host $_.Exception.ToString()
    exit 1
}
