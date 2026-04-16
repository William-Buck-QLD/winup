#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    # ----------------------------
    # OOBE-safe hardening
    # ----------------------------
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
    $ConfirmPreference  = 'None'
    $ProgressPreference = 'Continue'

    # ----------------------------
    # Bootstrap NuGet + PSGallery trust
    # ----------------------------
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }

    $psg = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psg -and $psg.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted | Out-Null
    }

    # ----------------------------
    # Install/import PSWindowsUpdate
    # ----------------------------
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue)) {
        Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck -AllowClobber -Scope AllUsers | Out-Null
    }
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop | Out-Null

    # ----------------------------
    # Ensure Microsoft Update is registered (Office/other MS products)
    # ----------------------------
    try {
        $mu = Get-WUServiceManager -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'Microsoft Update' }
        if (-not $mu) {
            Add-WUServiceManager -MicrosoftUpdate -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        }
    } catch { }

    function Invoke-InstallPass {
        param(
            [Parameter(Mandatory)]
            [ValidateSet('Software','Driver')]
            [string]$UpdateType,

            [Parameter(Mandatory)]
            [bool]$Optional
        )

        $class = if ($Optional) { 'Optional' } else { 'Required' }

        Write-Output ""
        Write-Output ("Scanning: {0} ({1})..." -f $UpdateType, $class)

        $scanParams = @{
            MicrosoftUpdate = $true
            UpdateType      = $UpdateType
            IgnoreUserInput = $true
            ErrorAction     = 'SilentlyContinue'
        }
        if ($Optional) { $scanParams['BrowseOnly'] = $true }

        $updates = Get-WindowsUpdate @scanParams
        $count = if ($updates) { $updates.Count } else { 0 }

        Write-Output ("Found: {0} {1} update(s)." -f $count, $class)

        if ($count -eq 0) { return }

        Write-Output ("Installing: {0} {1} update(s)..." -f $count, $class)

        $installParams = @{
            MicrosoftUpdate = $true
            UpdateType      = $UpdateType
            AcceptAll       = $true
            AutoReboot      = $true
            IgnoreUserInput = $true
            Confirm         = $false
            ErrorAction     = 'Stop'
        }
        if ($Optional) { $installParams['BrowseOnly'] = $true }

        # Use -Verbose as a switch, not inside the hashtable, to avoid parse/copy artefacts.
        Install-WindowsUpdate @installParams -Verbose | Out-Null

        Write-Output ("Pass complete: {0} ({1}). Rebooting if required." -f $UpdateType, $class)
    }

    # ----------------------------
    # Run all passes (everything)
    # ----------------------------
    Invoke-InstallPass -UpdateType Software -Optional:$false
    Invoke-InstallPass -UpdateType Software -Optional:$true
    Invoke-InstallPass -UpdateType Driver   -Optional:$false
    Invoke-InstallPass -UpdateType Driver   -Optional:$true

    Write-Output ""
    Write-Output "All passes completed. Rebooting if required."
    exit 0
}
catch {
    Write-Error $_
    exit 1
}
