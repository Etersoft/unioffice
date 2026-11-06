# Destructive lifecycle test: use only a disposable Windows VM with both MSIs installed.
param(
    [Parameter(Mandatory=$true)][string]$Msi32,
    [Parameter(Mandatory=$true)][string]$Msi64,
    [string]$OutputRoot = "$env:USERPROFILE\tmp\claude"
)
$ErrorActionPreference = 'Stop'
if ([IntPtr]::Size -ne 8) { throw 'Use 64-bit PowerShell' }
$root = Split-Path $PSScriptRoot -Parent
$work = Join-Path $OutputRoot ('unioffice-msi-windows-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $work | Out-Null
$products = @{32='{643315BA-DEF8-4CBE-AB26-C54B71535DFD}'; 64='{4C7393FD-6C4A-5B3F-8EC2-B7AD16262864}'}
$packages = @{32=(Resolve-Path $Msi32).Path; 64=(Resolve-Path $Msi64).Path}
$folders = @{32=(Join-Path $env:windir 'SysWOW64'); 64=(Join-Path $env:windir 'System32')}
$hashes = @{}
$installer = New-Object -ComObject WindowsInstaller.Installer

function Invoke-Child {
    param([string]$Executable, [string[]]$Arguments, [string]$Label)
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $Executable
    # msiexec parses switches itself; quoting every switch can open its usage UI.
    $info.Arguments = ($Arguments | ForEach-Object {
        if ($_ -match '\s') { '"' + $_ + '"' } else { $_ }
    }) -join ' '
    $info.UseShellExecute = $false
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $p = New-Object Diagnostics.Process
    $p.StartInfo = $info
    try {
        [void]$p.Start()
        $stdout = $p.StandardOutput.ReadToEndAsync()
        $stderr = $p.StandardError.ReadToEndAsync()
        if (-not $p.WaitForExit(60000)) { throw "$Label still running: PID $($p.Id)" }
        [IO.File]::WriteAllText((Join-Path $work "$Label.out"), $stdout.Result)
        [IO.File]::WriteAllText((Join-Path $work "$Label.err"), $stderr.Result)
        if ($p.ExitCode -ne 0) { throw "$Label exit $($p.ExitCode); inspect $work" }
    } finally { $p.Dispose() }
}

function Check-Payload {
    param([int]$Bit, [bool]$Present=$true)
    foreach ($file in 'unioffice_excel.dll','unioffice_excel.tlb') {
        $path = Join-Path $folders[$Bit] $file
        if ($Present) {
            if ((Get-FileHash $path -Algorithm SHA256).Hash -ne $hashes["$Bit-$file"]) { throw "PE${Bit} payload changed: $file" }
        } elseif (Test-Path $path) { throw "PE${Bit} payload remains: $file" }
    }
}

function Invoke-Msi {
    param([string[]]$Arguments, [string]$Label)
    Invoke-Child (Join-Path $env:windir 'System32\msiexec.exe') `
        ($Arguments + @('/qn','/norestart','/l*v',(Join-Path $work "$Label.log"))) $Label
    Write-Host "${Label}: MSI PASS"
}

function Probe-Client {
    param([int]$Bit, [string]$Label)
    Start-Process 'C:\Program Files (x86)\OpenOffice 4\program\soffice.exe' -ArgumentList '-headless -nofirststartwizard -norestore -nodefault' | Out-Null
    Start-Sleep -Seconds 3
    $script = Join-Path $root 'tests\excel-backend.vbs'
    Invoke-Child (Join-Path $folders[$Bit] 'cscript.exe') @('//nologo',$script) $Label
    if (-not ([IO.File]::ReadAllText((Join-Path $work "$Label.out"))).Contains('UniOffice integration: PASS')) {
        throw "$Label Automation failed; inspect $work"
    }
    Write-Host "${Label}: Automation PASS"
}

Write-Host "Results: $work"
try {
    foreach ($bit in 32,64) {
        if ($installer.ProductState($products[$bit]) -ne 5) { throw "PE${bit} is not installed for this user/machine" }
        foreach ($file in 'unioffice_excel.dll','unioffice_excel.tlb') {
            $hashes["$bit-$file"] = (Get-FileHash (Join-Path $folders[$bit] $file) -Algorithm SHA256).Hash
        }
        Probe-Client $bit "initial-$bit"
    }
    foreach ($bit in 32,64) {
        Invoke-Msi @('/famusv',$products[$bit]) "repair-$bit"
        foreach ($client in 32,64) {
            Check-Payload $client
            Probe-Client $client "after-repair-$bit-client-$client"
        }
    }
    Invoke-Msi @('/x',$products[64]) 'remove-64-first'
    Check-Payload 64 $false
    Check-Payload 32
    Probe-Client 32 'surviving-32'
    Invoke-Msi @('/i',$packages[64]) 'reinstall-64'
    Check-Payload 64
    Invoke-Msi @('/x',$products[32]) 'remove-32-first'
    Check-Payload 32 $false
    Check-Payload 64
    Probe-Client 64 'surviving-64'
    Invoke-Msi @('/x',$products[64]) 'remove-last-64'
    Check-Payload 64 $false
    if (Test-Path 'Registry::HKEY_CLASSES_ROOT\Excel.Application\CLSID') { throw 'Shared ProgID remains' }
    foreach ($bit in 32,64) {
        if ($installer.ProductState($products[$bit]) -ne -1) { throw "PE${bit} installer registration remains" }
    }
    [IO.File]::WriteAllText((Join-Path $work 'PASS'), 'Native MSI lifecycle: PASS')
    Write-Host 'Native MSI lifecycle: PASS'
} catch {
    [IO.File]::WriteAllText((Join-Path $work 'FAILED'), $_.Exception.ToString())
    Write-Host $_.Exception.ToString()
    exit 1
}
