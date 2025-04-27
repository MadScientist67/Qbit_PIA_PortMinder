Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ------------------------
# CONFIGURATION
# ------------------------

$logPath       = "$env:USERPROFILE\qbit_port_monitor.log"
$qbit_ini      = "$env:APPDATA\qBittorrent\qBittorrent.ini"
$qbit_exe      = "J:\Program Files\qBittorrent\qbittorrent.exe"
$piactl_path   = "C:\Program Files\Private Internet Access\piactl.exe"
$checkInterval = 900  # 15 minutes

# qBittorrent WebUI settings
$qbit_host     = "http://localhost:8080"  # Change if qBittorrent runs on a different IP/port
$qbit_username = "admin"                       # Your WebUI username
$qbit_password = "adminadmin"                  # Your WebUI password

# ------------------------
# FUNCTIONS
# ------------------------

function Log-Message($msg) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp :: $msg" | Out-File -Append -FilePath $logPath
}

function Get-PIAPort {
    $output = & "$piactl_path" get portforward 2>$null
    return $output -as [int]
}

function Get-QbitPort {
    $line = Select-String -Path $qbit_ini 'Session\\Port=\d*' | Select-Object -ExpandProperty Line
    return ($line -creplace '^[^0-9]*') -as [int]
}

function Update-QbitPort($newPort) {
    (Get-Content $qbit_ini) -replace "Session\\Port=\d*", "Session\Port=$newPort" | Set-Content -Path $qbit_ini
}

function GracefulShutdown-Qbit {
    try {
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $headers = @{
            "User-Agent" = "PowerShell"
        }

        # Login
        $loginResponse = Invoke-WebRequest -Uri "$qbit_host/api/v2/auth/login" `
                                           -Method Post `
                                           -Body "username=$qbit_username&password=$qbit_password" `
                                           -WebSession $session `
                                           -Headers $headers

        if ($loginResponse.StatusCode -ne 200) {
            Log-Message "Login failed. Status code: $($loginResponse.StatusCode)"
            return $false
        }

        # CSRF token and required headers
        $csrfToken = $session.Cookies.GetCookies($qbit_host)["SID"].Value
        $headers["Referer"] = $qbit_host
        $headers["Cookie"]  = "SID=$csrfToken"

        # Send shutdown command
        $shutdownResponse = Invoke-WebRequest -Uri "$qbit_host/api/v2/app/shutdown" `
                                              -Method Post `
                                              -WebSession $session `
                                              -Headers $headers

        if ($shutdownResponse.StatusCode -eq 200) {
            Log-Message "Graceful shutdown command sent."
            return $true
        } else {
            Log-Message "Shutdown failed. Status code: $($shutdownResponse.StatusCode)"
            return $false
        }
    } catch {
        Log-Message "Exception during shutdown: $_"
        return $false
    }
}

function Restart-QbitIfNeeded {
    $pia_port = Get-PIAPort
    $qbit_port = Get-QbitPort

    if ($pia_port -and ($qbit_port -ne $pia_port)) {
        Log-Message "Mismatch detected: qBit=$qbit_port, PIA=$pia_port. Restarting qBittorrent..."

        $stopped = GracefulShutdown-Qbit
        if (-not $stopped) {
            Stop-Process -Name "qbittorrent" -Force -ErrorAction SilentlyContinue
            Log-Message "Forced kill fallback used."
        }

        Start-Sleep -Seconds 30
        Update-QbitPort $pia_port
        Start-Sleep -Seconds 1

        Start-Process -FilePath $qbit_exe
        Log-Message "qBittorrent restarted with new port: $pia_port"
    } else {
        Log-Message "Ports match or PIA port unavailable. No action taken."
    }
}

# ------------------------
# SYSTRAY UI
# ------------------------

$icon = New-Object System.Windows.Forms.NotifyIcon
$icon.Icon = [System.Drawing.SystemIcons]::Information
$icon.Text = "qBittorrent Port Monitor"
$icon.Visible = $true

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$checkNowItem = $contextMenu.Items.Add("Check Now")
$viewLogItem  = $contextMenu.Items.Add("View Logs")
$exitItem     = $contextMenu.Items.Add("Exit")

$checkNowItem.Add_Click({
    Restart-QbitIfNeeded
})

$viewLogItem.Add_Click({
    if (Test-Path $logPath) {
        notepad.exe $logPath
    } else {
        [System.Windows.Forms.MessageBox]::Show("Log file not found at $logPath.", "Error")
    }
})

$exitItem.Add_Click({
    $icon.Visible = $false
    [System.Windows.Forms.Application]::Exit()
})

$icon.ContextMenuStrip = $contextMenu

# ------------------------
# BACKGROUND PORT CHECKING LOOP
# ------------------------

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = $checkInterval * 1000  # Convert seconds to milliseconds
$timer.Add_Tick({
    Restart-QbitIfNeeded
})
$timer.Start()


# ------------------------
# TRAY EVENT LOOP
# ------------------------

[System.Windows.Forms.Application]::Run()
