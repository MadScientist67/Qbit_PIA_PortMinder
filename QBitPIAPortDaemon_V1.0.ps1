# Configuration for qBittorrent WebUI API
$QB_URL = "http://localhost:8080"  # Change if qBittorrent runs on a different IP/port
$USERNAME = "admin"  # Change to your WebUI username
$PASSWORD = "adminadmin"  # Change to your WebUI password
$qbit_ini = "C:\Users\Jake\AppData\Roaming\qBittorrent\qBittorrent.ini"

function Get-Timestamp {
    return (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
}

Write-Output "$(Get-Timestamp) - qBittorrent Port Monitor started. Checking every 30 minutes..."
while ($true) {
    Write-Output "$(Get-Timestamp) - Checking qBittorrent port..."

    # Get current qBittorrent port
    $qbit_port = Select-String -Path $qbit_ini 'Session\\Port=\d*' | Select-Object -ExpandProperty Line
    $qbit_port = $qbit_port -creplace '^[^0-9]*'
    $qbit_process = Get-Process qbittorrent -ErrorAction SilentlyContinue

    # Get PIA assigned port
    $pia_port = & "C:\Program Files\Private Internet Access\piactl.exe" get portforward
    $pia_port = $pia_port -as [int]

    if ($qbit_port -ne $pia_port) {
        Write-Output "$(Get-Timestamp) - Port mismatch detected! qBittorrent: $qbit_port, PIA: $pia_port"
        
        if ($qbit_process) {
            Write-Output "$(Get-Timestamp) - Shutting down qBittorrent gracefully..."

            # Create a session to store cookies
            $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

            # Authenticate and capture cookies
            $loginResponse = Invoke-WebRequest -Uri "$QB_URL/api/v2/auth/login" `
                -Method Post `
                -Body @{username=$USERNAME; password=$PASSWORD} `
                -UseBasicParsing `
                -WebSession $session

            if ($loginResponse.Content -match "Ok") {
                Write-Output "$(Get-Timestamp) - Login successful. Sending shutdown request..."

                try {
                    # Send shutdown command using POST
                    $headers = @{
                        "Content-Type" = "application/x-www-form-urlencoded"
                        "Referer" = $QB_URL
                    }

                    $shutdownResponse = Invoke-WebRequest -Uri "$QB_URL/api/v2/app/shutdown" `
                        -Method Post `
                        -Headers $headers `
                        -WebSession $session `
                        -UseBasicParsing

                    Write-Output "$(Get-Timestamp) - Shutdown command sent successfully."
                } catch {
                    Write-Output "$(Get-Timestamp) - Failed to send shutdown command. Status Code: $($_.Exception.Response.StatusCode) - $($_.ErrorDetails.Message)"
                }
            } else {
                Write-Output "$(Get-Timestamp) - Login failed. Please check your credentials and WebUI settings."
            }

            # Wait for qBittorrent to fully close
            Start-Sleep -Seconds 30

            # Update qBittorrent config
            Write-Output "$(Get-Timestamp) - Updating qBittorrent config..."
            (Get-Content $qbit_ini) -replace "Session\\Port=\d*", "Session\Port=$pia_port" | Set-Content -Path $qbit_ini
            Start-Sleep -Seconds 1

            # Start qBittorrent again
            Write-Output "$(Get-Timestamp) - Starting qBittorrent..."
            Start-Process -FilePath "J:\Program Files\qBittorrent\qbittorrent.exe"
            Start-Sleep -Seconds 5

            # Verify qBittorrent restarted
            $qbit_process = Get-Process qbittorrent -ErrorAction SilentlyContinue
            if ($qbit_process) {
                $qbit_process.CloseMainWindow()
                Write-Output "$(Get-Timestamp) - qBittorrent started! Monitoring will continue..."
            } else {
                Write-Output "$(Get-Timestamp) - qBittorrent not started yet. Waiting..."
                Start-Sleep -Seconds 10
                $qbit_process = Get-Process qbittorrent -ErrorAction SilentlyContinue
                if ($qbit_process) {
                    $qbit_process.CloseMainWindow()
                    Write-Output "$(Get-Timestamp) - qBittorrent started! Monitoring will continue..."
                } else {
                    Write-Output "$(Get-Timestamp) - Failed to start qBittorrent!"
                }
            }
        }
    } else {
        Write-Output "$(Get-Timestamp) - Ports match, nothing to do..."
    }

    # Countdown Timer for 30 Minutes (Updates Every Minute)
    for ($i = 30; $i -gt 0; $i--) {
        Write-Host "$(Get-Timestamp) - Next check in $i minutes..." -NoNewline
        Start-Sleep -Seconds 60
        Write-Host "`r" -NoNewline
    }
}
