#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
    $ConfirmPreference  = 'None'
    $ProgressPreference = 'Continue'

    # NuGet + PSGallery trust
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }
    $psg = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psg -and $psg.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted | Out-Null
    }

    # PSWindowsUpdate
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue)) {
        Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck -AllowClobber -Scope AllUsers | Out-Null
    }
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop | Out-Null

    # Ensure Microsoft Update is registered (Office/other MS products)
    try {
        $mu = Get-WUServiceManager -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'Microsoft Update' }
        if (-not $mu) {
            Add-WUServiceManager -MicrosoftUpdate -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        }
    } catch { }

    function Get-UpdateCount {
        param(
            [ValidateSet('Software','Driver')]
            [string]$UpdateType,
            [bool]$Optional
        )
        $p = @{
            MicrosoftUpdate = $true
            UpdateType      = $UpdateType     # Driver vs Software. [1](https://deepwiki.com/mgajda83/PSWindowsUpdate/7.4-filtering-and-search-criteria)[2](https://www.powershellgallery.com/packages/PSWindowsUpdate/2.2.1.5)
            IgnoreUserInput = $true
            ErrorAction     = 'SilentlyContinue'
        }
        if ($Optional) { $p['BrowseOnly'] = $true }  # Optional/BrowseOnly. [1](https://deepwiki.com/mgajda83/PSWindowsUpdate/7.4-filtering-and-search-criteria)[2](https://www.powershellgallery.com/packages/PSWindowsUpdate/2.2.1.5)
        $u = Get-WindowsUpdate @p
        if ($u) { return $u.Count }
        return 0
    }

    function Install-Pass {
        param(
            [ValidateSet('Software','Driver')]
            [string]$UpdateType,
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

        if ($count -eq 0) { return $false }

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

        Install-WindowsUpdate @installParams -Verbose | Out-Null
        Write-Output ("Pass complete: {0} ({1}). Rebooting if required." -f $UpdateType, $class)
        return $true
    }

    # 1) Required software first (the important stuff)
    Install-Pass -UpdateType Software -Optional:$false | Out-Null

    # 2) Optional/Driver content often appears only after scan refresh.
    # Retry discovery a few times before concluding "none".
    $maxAttempts = 5
    $sleepSeconds = 60

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        Write-Output ""
        Write-Output ("Refresh cycle {0}/{1}: checking for optional/software/driver content..." -f $attempt, $maxAttempts)

        $softOpt = Get-UpdateCount -UpdateType Software -Optional:$true
        $drvReq  = Get-UpdateCount -UpdateType Driver   -Optional:$false
        $drvOpt  = Get-UpdateCount -UpdateType Driver   -Optional:$true

        Write-Output ("Detected counts: Optional Software={0}, Required Drivers={1}, Optional Drivers={2}" -f $softOpt, $drvReq, $drvOpt)

        if (($softOpt + $drvReq + $drvOpt) -gt 0) { break }

        if ($attempt -lt $maxAttempts) {
            Write-Output ("Nothing yet. Waiting {0} seconds for scan results to populate..." -f $sleepSeconds)
            Start-Sleep -Seconds $sleepSeconds
        }
    }

    # 3) Now install the rest (everything)
    Install-Pass -UpdateType Software -Optional:$true  | Out-Null
    Install-Pass -UpdateType Driver   -Optional:$false | Out-Null
    Install-Pass -UpdateType Driver   -Optional:$true  | Out-Null

    Write-Output ""
    Write-Output "All passes completed. Rebooting if required."
    exit 0
}
catch {
    Write-Error $_
    exit 1
}
