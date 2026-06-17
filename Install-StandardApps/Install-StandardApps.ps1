# Install-StandardApps.ps1
# Run this script from an elevated PowerShell window (or as SYSTEM via RMM).
#
# Behavior:
# - Checks whether the Microsoft Office suite is already installed.
# - Reports whether Word, Excel, PowerPoint, and Outlook Classic are present.
# - Does NOT reinstall Office merely because Outlook Classic is missing.
# - Installs Microsoft 365 Apps only when no core Office applications are detected.
# - Installs Google Chrome and Adobe Acrobat Reader when they are missing.
# - Writes status messages to the console and a transcript log file.

#Requires -Version 5.1
#Requires -RunAsAdministrator

$ErrorActionPreference = "Stop"

# Use "O365ProPlusRetail" instead if the organization uses
# Microsoft 365 Apps for enterprise.
$OfficeProductId = "O365BusinessRetail"

$LogFile = Join-Path `
    $env:TEMP `
    "Standard-App-Install-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-Status {
    param (
        [Parameter(Mandatory)]
        [string]$Message
    )

    # Write-Host (not Write-Output) so status text is not emitted on the
    # success stream as a function "return value." Start-Transcript still
    # captures Write-Host, so these lines remain in the log file.
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
}

function Test-IsAdministrator {
    $CurrentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()

    $Principal = New-Object `
        Security.Principal.WindowsPrincipal($CurrentIdentity)

    return $Principal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
}

function Get-ProgramFilesFolders {
    $Folders = @(
        [Environment]::GetFolderPath("ProgramFiles"),
        [Environment]::GetFolderPath("ProgramFilesX86")
    )

    return $Folders |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($_)
        } |
        Select-Object -Unique
}

function Test-ExecutableRegistered {
    param (
        [Parameter(Mandatory)]
        [string]$ExecutableName,

        [string[]]$FallbackPaths = @()
    )

    $RegistryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$ExecutableName",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\$ExecutableName",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$ExecutableName"
    )

    foreach ($RegistryPath in $RegistryPaths) {
        if (Test-Path -LiteralPath $RegistryPath) {
            return $true
        }
    }

    foreach ($FallbackPath in $FallbackPaths) {
        if (
            -not [string]::IsNullOrWhiteSpace($FallbackPath) -and
            (Test-Path -LiteralPath $FallbackPath)
        ) {
            return $true
        }
    }

    return $false
}

function Test-InstalledProgramName {
    param (
        [Parameter(Mandatory)]
        [string]$Pattern
    )

    $RegistryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($RegistryPath in $RegistryPaths) {
        $Match = Get-ItemProperty `
            -Path $RegistryPath `
            -ErrorAction SilentlyContinue |
            Where-Object {
                $_.DisplayName -and $_.DisplayName -match $Pattern
            } |
            Select-Object -First 1

        if ($Match) {
            return $true
        }
    }

    return $false
}

function Get-OfficeApplicationStatus {
    $ProgramFilesFolders = Get-ProgramFilesFolders

    # Modern Click-to-Run Office (2016/2019/2021/365) lives under Office16.
    # Older MSI installs (Office 2013 = Office15) are not detected here.
    $OfficeFolders = foreach ($Folder in $ProgramFilesFolders) {
        Join-Path $Folder "Microsoft Office\root\Office16"
        Join-Path $Folder "Microsoft Office\Office16"
    }

    $OfficeApplications = [ordered]@{
        "Microsoft Word"              = "WINWORD.EXE"
        "Microsoft Excel"             = "EXCEL.EXE"
        "Microsoft PowerPoint"        = "POWERPNT.EXE"
        "Microsoft Outlook (classic)" = "OUTLOOK.EXE"
    }

    $Status = [ordered]@{}

    foreach ($Application in $OfficeApplications.GetEnumerator()) {
        $FallbackPaths = foreach ($Folder in $OfficeFolders) {
            Join-Path $Folder $Application.Value
        }

        $Status[$Application.Key] = Test-ExecutableRegistered `
            -ExecutableName $Application.Value `
            -FallbackPaths $FallbackPaths
    }

    return $Status
}

function Test-OfficeSuiteInstalled {
    param (
        [Parameter(Mandatory)]
        [System.Collections.IDictionary]$OfficeStatus
    )

    # Treat Microsoft Office as installed if any core desktop Office
    # application is present.
    #
    # Outlook Classic is intentionally excluded from this decision.
    # Its absence should not trigger a complete Office reinstallation.

    return (
        $OfficeStatus["Microsoft Word"] -or
        $OfficeStatus["Microsoft Excel"] -or
        $OfficeStatus["Microsoft PowerPoint"]
    )
}

function Test-NewOutlookInstalled {
    try {
        $NewOutlook = Get-AppxPackage `
            -AllUsers `
            -Name "Microsoft.OutlookForWindows" `
            -ErrorAction SilentlyContinue

        return [bool]$NewOutlook
    }
    catch {
        try {
            $NewOutlook = Get-AppxPackage `
                -Name "Microsoft.OutlookForWindows" `
                -ErrorAction SilentlyContinue

            return [bool]$NewOutlook
        }
        catch {
            return $false
        }
    }
}

function Test-GoogleChromeInstalled {
    $FallbackPaths = foreach ($Folder in (Get-ProgramFilesFolders)) {
        Join-Path $Folder "Google\Chrome\Application\chrome.exe"
    }

    $FallbackPaths += Join-Path `
        $env:LOCALAPPDATA `
        "Google\Chrome\Application\chrome.exe"

    return (
        (
            Test-ExecutableRegistered `
                -ExecutableName "chrome.exe" `
                -FallbackPaths $FallbackPaths
        ) -or
        (
            Test-InstalledProgramName `
                -Pattern "^Google Chrome$"
        )
    )
}

function Test-AdobeInstalled {
    $FallbackPaths = foreach ($Folder in (Get-ProgramFilesFolders)) {
        Join-Path $Folder "Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
        Join-Path $Folder "Adobe\Acrobat DC\Acrobat\Acrobat.exe"
    }

    return (
        (
            Test-ExecutableRegistered `
                -ExecutableName "AcroRd32.exe" `
                -FallbackPaths $FallbackPaths
        ) -or
        (
            Test-ExecutableRegistered `
                -ExecutableName "Acrobat.exe" `
                -FallbackPaths $FallbackPaths
        ) -or
        (
            Test-InstalledProgramName `
                -Pattern "^(Adobe Acrobat|Adobe Acrobat Reader|Adobe Reader)"
        )
    )
}

function Resolve-WinGetPath {
    # In an interactive elevated session winget.exe is normally on PATH.
    $Command = Get-Command "winget.exe" -ErrorAction SilentlyContinue
    if ($Command) {
        return $Command.Source
    }

    # When this script runs as SYSTEM (for example, pushed by an RMM agent),
    # winget.exe is NOT on PATH even though App Installer is provisioned.
    # Resolve the versioned executable directly from the WindowsApps folder
    # and use the most recently written copy.
    $SearchPattern = Join-Path `
        $env:ProgramFiles `
        "WindowsApps\Microsoft.DesktopAppInstaller_*__8wekyb3d8bbwe\winget.exe"

    $Resolved = Resolve-Path -Path $SearchPattern -ErrorAction SilentlyContinue

    if ($Resolved) {
        $Newest = $Resolved |
            ForEach-Object { Get-Item -LiteralPath $_.Path } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        return $Newest.FullName
    }

    return $null
}

function Install-WinGetPackage {
    param (
        [Parameter(Mandatory)]
        [string]$ApplicationName,

        [Parameter(Mandatory)]
        [string]$PackageId,

        [Parameter(Mandatory)]
        [string]$WinGetPath
    )

    Write-Status "$ApplicationName is not installed. Beginning installation."

    & $WinGetPath install `
        --id $PackageId `
        --exact `
        --source winget `
        --scope machine `
        --silent `
        --accept-package-agreements `
        --accept-source-agreements `
        --disable-interactivity

    # 0            = success
    # -1978335189  = no applicable upgrade / already installed (treated as OK)
    $AcceptableExitCodes = @(0, -1978335189)

    if ($AcceptableExitCodes -notcontains $LASTEXITCODE) {
        throw "$ApplicationName installation failed. WinGet exit code: $LASTEXITCODE"
    }

    Write-Status "$ApplicationName installation process completed."
}

function Install-Microsoft365Apps {
    param (
        [Parameter(Mandatory)]
        [string]$ProductId
    )

    $OfficeWorkingFolder = Join-Path `
        $env:TEMP `
        "Microsoft365-App-Deployment-$PID"

    $OfficeSetupPath = Join-Path `
        $OfficeWorkingFolder `
        "setup.exe"

    $OfficeConfigurationPath = Join-Path `
        $OfficeWorkingFolder `
        "configuration.xml"

    New-Item `
        -Path $OfficeWorkingFolder `
        -ItemType Directory `
        -Force |
        Out-Null

    $OfficeConfiguration = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="Current">
        <Product ID="$ProductId">
            <Language ID="en-us" />
        </Product>
    </Add>
    <Display Level="None" AcceptEULA="TRUE" />
    <Updates Enabled="TRUE" Channel="Current" />
</Configuration>
"@

    Set-Content `
        -Path $OfficeConfigurationPath `
        -Value $OfficeConfiguration `
        -Encoding UTF8

    Write-Status "Downloading the Microsoft Office Deployment Tool."

    [Net.ServicePointManager]::SecurityProtocol = `
        [Net.SecurityProtocolType]::Tls12

    Invoke-WebRequest `
        -Uri "https://officecdn.microsoft.com/pr/wsus/setup.exe" `
        -OutFile $OfficeSetupPath `
        -UseBasicParsing

    # Verify the downloaded bootstrapper is validly signed by Microsoft
    # before executing it.
    $Signature = Get-AuthenticodeSignature -FilePath $OfficeSetupPath

    if ($Signature.Status -ne "Valid") {
        throw "Office setup.exe signature is not valid (status: $($Signature.Status))."
    }

    if ($Signature.SignerCertificate.Subject -notmatch "Microsoft Corporation") {
        throw "Office setup.exe is not signed by Microsoft Corporation."
    }

    Write-Status "Office Deployment Tool signature verified."
    Write-Status "Beginning Microsoft 365 Apps installation."
    Write-Status "The Office deployment process may take several minutes."

    $OfficeInstallProcess = Start-Process `
        -FilePath $OfficeSetupPath `
        -ArgumentList "/configure `"$OfficeConfigurationPath`"" `
        -PassThru

    # Allow the Office deployment process to run for up to 45 minutes.
    $OfficeInstallCompleted = $OfficeInstallProcess.WaitForExit(2700000)

    if (-not $OfficeInstallCompleted) {
        try {
            Stop-Process `
                -Id $OfficeInstallProcess.Id `
                -Force `
                -ErrorAction SilentlyContinue
        }
        catch {
            # Continue to the error message if the process already exited.
        }

        throw "Microsoft 365 Apps installation exceeded the 45-minute timeout."
    }

    $OfficeInstallProcess.Refresh()

    if ($OfficeInstallProcess.ExitCode -ne 0) {
        throw "Microsoft 365 Apps installation failed. Exit code: $($OfficeInstallProcess.ExitCode)"
    }

    Write-Status "Microsoft 365 Apps installation process completed."

    Remove-Item `
        -Path $OfficeWorkingFolder `
        -Recurse `
        -Force `
        -ErrorAction SilentlyContinue
}

try {
    Start-Transcript `
        -Path $LogFile `
        -Append |
        Out-Null

    if (-not (Test-IsAdministrator)) {
        throw "This script must be run as Administrator."
    }

    Write-Status "Starting standard application check."

    # =========================================================
    # Microsoft Office
    # =========================================================

    Write-Status "Checking Microsoft Office applications."

    $OfficeStatus = Get-OfficeApplicationStatus

    foreach ($Application in $OfficeStatus.GetEnumerator()) {
        if ($Application.Value) {
            Write-Status "$($Application.Key) is already installed."
        }
        else {
            Write-Status "$($Application.Key) is not installed."
        }
    }

    if (Test-NewOutlookInstalled) {
        Write-Status "New Outlook for Windows is already installed."
    }
    else {
        Write-Status "New Outlook for Windows was not detected."
    }

    if (Test-OfficeSuiteInstalled -OfficeStatus $OfficeStatus) {
        Write-Status "Microsoft Office suite is already installed."
        Write-Status "Skipping Microsoft 365 Apps installation."

        if (-not $OfficeStatus["Microsoft Outlook (classic)"]) {
            Write-Status "Microsoft Outlook (classic) is not installed."
            Write-Status "A missing classic Outlook installation does not trigger a full Office suite reinstall."
        }
    }
    else {
        Write-Status "No core Microsoft Office desktop applications were detected."
        Write-Status "Beginning Microsoft 365 Apps installation."

        Install-Microsoft365Apps `
            -ProductId $OfficeProductId

        $OfficeStatus = Get-OfficeApplicationStatus

        if (Test-OfficeSuiteInstalled -OfficeStatus $OfficeStatus) {
            Write-Status "Microsoft Office suite installed successfully."
        }
        else {
            throw "Microsoft Office suite installation verification failed."
        }
    }

    # =========================================================
    # Google Chrome and Adobe Acrobat Reader
    # =========================================================

    $ChromeInstalled = Test-GoogleChromeInstalled
    $AdobeInstalled = Test-AdobeInstalled

    # Resolve WinGet once, only if we actually need to install something.
    $WinGetPath = $null

    if (-not $ChromeInstalled -or -not $AdobeInstalled) {
        $WinGetPath = Resolve-WinGetPath

        if (-not $WinGetPath) {
            throw @"
WinGet is not available on this computer.
Install or update Microsoft App Installer, then rerun this script.
"@
        }

        Write-Status "Using WinGet at: $WinGetPath"
    }

    # =========================================================
    # Google Chrome
    # =========================================================

    if ($ChromeInstalled) {
        Write-Status "Google Chrome is already installed."
    }
    else {
        Install-WinGetPackage `
            -ApplicationName "Google Chrome" `
            -PackageId "Google.Chrome" `
            -WinGetPath $WinGetPath

        if (Test-GoogleChromeInstalled) {
            Write-Status "Google Chrome installed successfully."
        }
        else {
            throw "Google Chrome installation verification failed."
        }
    }

    # =========================================================
    # Adobe Acrobat Reader
    # =========================================================

    if ($AdobeInstalled) {
        Write-Status "Adobe Acrobat or Adobe Acrobat Reader is already installed."
    }
    else {
        Install-WinGetPackage `
            -ApplicationName "Adobe Acrobat Reader" `
            -PackageId "Adobe.Acrobat.Reader.64-bit" `
            -WinGetPath $WinGetPath

        if (Test-AdobeInstalled) {
            Write-Status "Adobe Acrobat Reader installed successfully."
        }
        else {
            throw "Adobe Acrobat Reader installation verification failed."
        }
    }

    Write-Status "Standard application check completed successfully."
    Write-Status "Log file: $LogFile"

    # Explicit success code for RMM / automation callers.
    $script:ExitCode = 0
}
catch {
    Write-Status "ERROR: $($_.Exception.Message)"
    Write-Status "Review the log file for additional details: $LogFile"

    $script:ExitCode = 1
}
finally {
    try {
        Stop-Transcript | Out-Null
    }
    catch {
        # Do nothing if transcript logging was not started successfully.
    }
}

exit $script:ExitCode
