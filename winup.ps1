#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    # OOBE-safe hardening
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
    $ProgressPreference = 'Continue'
    $ConfirmPreference  = 'None'

    # NuGet and PSGallery
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }

    $psg = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psg -and $psg.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted | Out-Null
    }

    # PSWindowsUpdate
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck -AllowClobber -Scope AllUsers | Out-Null
    }
    Import-Module PSWindowsUpdate -Force

    # Ensure Microsoft Update is enabled (required for drivers)
    try {
        $mu = Get-WUServiceManager | Where-Object { $_.Name -match 'Microsoft Update' }
        if (-not $mu) {
            Add-WUServiceManager -MicrosoftUpdate -Confirm:$false | Out-Null
        }
    } catch { }

    # Scan (includes drivers)
    $updates = Get-WindowsUpdate `
        -MicrosoftUpdate `
        -IgnoreUserInput `
        -ErrorAction SilentlyContinue

    if (-not $updates -or $updates.Count -eq 0) {
        Write-Output "No updates (including drivers) available."
        exit 0
    }

    Write-Output "Installing $($updates.Count) update(s), including drivers..."

    # Install everything, drivers included
    Install-WindowsUpdate `
        -MicrosoftUpdate `
        -AcceptAll `
        -AutoReboot `
        -IgnoreUserInput `
        -Confirm:$false `
        -Verbose `
        -ErrorAction Stop

    Write-Output "Update pass completed. Rebooting if required."
    exit 0
}
catch {
    Write-Error $_
    exit 1
}
