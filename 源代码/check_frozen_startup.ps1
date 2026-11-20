param(
    [Parameter(Mandatory=$true)][string]$Executable,
    [Parameter(Mandatory=$true)][string]$ExpectedTitle,
    [string]$PythonExecutable = 'python'
)
$ErrorActionPreference = 'Stop'
$resolvedExe = (Resolve-Path -LiteralPath $Executable).Path
$started = Start-Process -FilePath $resolvedExe -WorkingDirectory (Split-Path -Parent $resolvedExe) -WindowStyle Hidden -PassThru
$ownedIds = @($started.Id)
$passed = $false
$deadline = (Get-Date).AddSeconds(45)
try {
    while ((Get-Date) -lt $deadline) {
        $children = & $PythonExecutable -c "import json,psutil,sys; n=int(sys.argv[1]); print(json.dumps([p.pid for p in psutil.Process(n).children(recursive=True)] if psutil.pid_exists(n) else []))" $started.Id
        if ($LASTEXITCODE -ne 0) { throw 'Could not discover the test process children (requires developer Python + psutil).' }
        # foreach unrolls the array; PS 5.1 ConvertFrom-Json returns '[]' as one empty-array item.
        foreach ($childId in ($children | ConvertFrom-Json)) { $ownedIds += [int]$childId }
        $ownedIds = @($ownedIds | Select-Object -Unique)
        foreach ($processId in $ownedIds) {
            $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
            if ($process -and $process.MainWindowTitle -eq $ExpectedTitle) {
                $passed = $true
                Write-Output "PASS: packaged application opened its main window (PID $processId)."
                break
            }
            if ($process -and $process.MainWindowTitle -match 'Unhandled exception|Fatal error') {
                throw "Packaged startup failed: $($process.MainWindowTitle)"
            }
        }
        if ($passed) { break }
        $started.Refresh()
        if ($started.HasExited) { throw "Packaged application exited before startup: $($started.ExitCode)" }
        Start-Sleep -Milliseconds 500
    }
    if (-not $passed) { throw 'Timed out waiting for the packaged application main window.' }
}
finally {
    # Close only the test instance and the child processes discovered under it.
    foreach ($processId in ($ownedIds | Sort-Object -Descending)) {
        Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
    }
}
