#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    # ------------------------------------------------------------
    # OOBE-safe hardening / suppression
    # ------------------------------------------------------------
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
    $ConfirmPreference  = 'None'
    $ProgressPreference = 'Continue'

    # ------------------------------------------------------------
    # Bootstrap NuGet + PSGallery trust
    # ------------------------------------------------------------
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }

    $psg = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psg -and $psg.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted | Out-Null
    }

    # ------------------------------------------------------------
    # Install/import PSWindowsUpdate
    # ------------------------------------------------------------
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue)) {
        Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck -AllowClobber -Scope AllUsers | Out-Null
    }
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop | Out-Null

    # ------------------------------------------------------------
    # Ensure Microsoft Update is registered (Office/other MS products)
    # ------------------------------------------------------------
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
            [ValidateSet('Required','Optional')]
            [string]$Class
        )

        $isOptional = ($Class -eq 'Optional')

        Write-Output ""
        Write-Output ("Scanning: {0} ({1})..." -f $UpdateType, $Class)

        # PSWindowsUpdate pre-search filters:
        # - UpdateType: Software vs Driver
        # - BrowseOnly: Optional updates
        # These are native to the module’s two-stage filtering model. [3](https://deepwiki.com/mgajda83/PSWindowsUpdate/7.4-filtering-and-search-criteria)[4](https://www.powershellgallery.com/packages/PSWindowsUpdate/2.2.1.5)
        $scanParams = @{
            MicrosoftUpdate = $true
            UpdateType      = $UpdateType
            IgnoreUserInput = $true
            ErrorAction     = 'SilentlyContinue'
        }
        if ($isOptional) { $scanParams['BrowseOnly'] = $true }

        $updates = Get-WindowsUpdate @scanParams
        $count = if ($updates) { $updates.Count } else { 0 }

        Write-Output ("Found: {0} {1} update(s)." -f $count, $Class)

        if ($count -eq 0) { return }

        Write-Output ("Installing: {0} {1} update(s)..." -f $count, $Class)

        $installParams = @{
            MicrosoftUpdate = $true
            UpdateType      = $UpdateType
            AcceptAll       = $true
            AutoReboot      = $true
            IgnoreUserInput = $true
            Confirm         = $false
            Verbose         = $true
