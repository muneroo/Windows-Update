# ============================================================
# WindowsUpdate.ps1
# Purpose:
#   Scan, download, and install Windows Updates.
#   Log results and detect whether a reboot is required.
# ============================================================

# ------------------------------------------------------------
# Log Configuration
# ------------------------------------------------------------

$LogDirectory = "C:\Logs\Update"

# Create log directory if it does not exist
if (!(Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

# Create a new log file for each run
$DateStamp = Get-Date -Format "yyyy-MM-dd_HHmm"
$LogFile = "$LogDirectory\WindowsUpdate_$DateStamp.log"

# ------------------------------------------------------------
# Logging Function
# ------------------------------------------------------------

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Entry = "$Timestamp [$Level] $Message"

    Add-Content -Path $LogFile -Value $Entry
    Write-Output $Entry
}

# ------------------------------------------------------------
# Start Logging
# ------------------------------------------------------------

Write-Log "============================================================"
Write-Log "Windows Update maintenance started."
Write-Log "Computer: $env:COMPUTERNAME"

try {

    # --------------------------------------------------------
    # Connect to the native Windows Update Agent
    # --------------------------------------------------------

    Write-Log "Connecting to Windows Update Agent..."

    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

    # --------------------------------------------------------
    # Scan for updates
    # --------------------------------------------------------

    Write-Log "Scanning for applicable Windows updates..."

    $SearchResult = $UpdateSearcher.Search(
        "IsInstalled=0 and IsHidden=0"
    )

    $UpdateCount = $SearchResult.Updates.Count

    Write-Log "Found $UpdateCount applicable update(s)."

    if ($UpdateCount -eq 0) {

        Write-Log "No updates are currently required."
        Write-Log "Windows Update maintenance completed successfully."
        Write-Log "============================================================"

        exit 0
    }

    # --------------------------------------------------------
    # Build update collection
    # --------------------------------------------------------

    $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl

    foreach ($Update in $SearchResult.Updates) {

        Write-Log "Found update: $($Update.Title)"

        if (!$Update.EulaAccepted) {
            $Update.AcceptEula()
        }

        [void]$UpdatesToInstall.Add($Update)
    }

    # --------------------------------------------------------
    # Download updates
    # --------------------------------------------------------

    Write-Log "Downloading $($UpdatesToInstall.Count) update(s)..."

    $Downloader = $UpdateSession.CreateUpdateDownloader()
    $Downloader.Updates = $UpdatesToInstall

    $DownloadResult = $Downloader.Download()

    Write-Log "Download operation completed. Result code: $($DownloadResult.ResultCode)"

    # --------------------------------------------------------
    # Only install successfully downloaded updates
    # --------------------------------------------------------

    $DownloadedUpdates = New-Object -ComObject Microsoft.Update.UpdateColl

    foreach ($Update in $UpdatesToInstall) {

        if ($Update.IsDownloaded) {

            Write-Log "Downloaded successfully: $($Update.Title)"
            [void]$DownloadedUpdates.Add($Update)

        }
        else {

            Write-Log "Download failed: $($Update.Title)" "ERROR"
        }
    }

    if ($DownloadedUpdates.Count -eq 0) {

        Write-Log "No updates were successfully downloaded." "ERROR"
        exit 1
    }

    # --------------------------------------------------------
    # Install updates
    # --------------------------------------------------------

    Write-Log "Installing $($DownloadedUpdates.Count) update(s)..."

    $Installer = $UpdateSession.CreateUpdateInstaller()
    $Installer.Updates = $DownloadedUpdates

    $InstallResult = $Installer.Install()

    Write-Log "Installation operation completed. Result code: $($InstallResult.ResultCode)"

    # --------------------------------------------------------
    # Log individual results
    # --------------------------------------------------------

    for ($i = 0; $i -lt $DownloadedUpdates.Count; $i++) {

        $Update = $DownloadedUpdates.Item($i)
        $Result = $InstallResult.GetUpdateResult($i)

        Write-Log "Update: $($Update.Title)"
        Write-Log "Result Code: $($Result.ResultCode)"
        Write-Log "HRESULT: $($Result.HResult)"
    }

    # --------------------------------------------------------
    # Reboot detection
    # --------------------------------------------------------

    if ($InstallResult.RebootRequired) {

    Write-Log "A system restart is required." "WARNING"
    Write-Log "Computer will restart automatically in 60 seconds." "WARNING"

    Start-Sleep -Seconds 60

    Restart-Computer -Force
}
else {

    Write-Log "No restart is currently required."
}

    Write-Log "Windows Update maintenance completed."
    Write-Log "============================================================"

    exit 0
}

catch {

    Write-Log "An unexpected error occurred." "ERROR"
    Write-Log $_.Exception.Message "ERROR"
    Write-Log "============================================================"

    exit 1
}
