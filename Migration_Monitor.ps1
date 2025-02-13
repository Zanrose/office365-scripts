# Exchange Online Migration Tracker
# Author: ZoomSupport
# Description: This script continuously monitors mailbox migrations.
# You need to manually connect to Exchange Online powershell and run this. The Script will auto refresh the stats every 30 seconds. 

# Function to generate progress bar
function Get-ProgressBar ($percent) {
    $totalBars = 20  # Number of blocks in the progress bar
    $filledBars = [math]::Round(($percent / 100) * $totalBars)
    $emptyBars = $totalBars - $filledBars
    $filled = "".PadLeft($filledBars, [char]0x2588)  # Filled blocks
    $empty = "".PadLeft($emptyBars, [char]0x2591)  # Empty blocks
    return "$filled$empty $percent%"
}

# Start monitoring loop
while ($true) {
    Clear-Host

    # Get all move requests
    $moveRequests = Get-MoveRequest | Get-MoveRequestStatistics 

    # Filter Synced mailboxes (95% and above)
    $synced = $moveRequests | Where-Object { $_.PercentComplete -ge 95 }

    # Filter In Progress mailboxes (less than 95%)
    $inProgress = $moveRequests | Where-Object { $_.PercentComplete -lt 95 }

    # Calculate overall progress
    $totalMailboxes = $moveRequests.Count
    $completedMailboxes = $synced.Count
    if ($totalMailboxes -gt 0) {
        $completionPercentage = [math]::Round(($completedMailboxes / $totalMailboxes) * 100, 2)
    } else {
        $completionPercentage = 0
    }

    # Display overall migration progress
    Write-Host "==== Migration Progress: $completionPercentage% Complete ====" -ForegroundColor Cyan
    Write-Host ""

    # Display Synced Mailboxes
    if ($synced.Count -gt 0) {
        Write-Host "==== Synced Mailboxes (95% and above) ====" -ForegroundColor Green
        $synced | ForEach-Object {
            [PSCustomObject]@{
                DisplayName     = $_.DisplayName
                StatusDetail    = $_.StatusDetail
                Progress        = Get-ProgressBar $_.PercentComplete
            }
        } | Format-Table -Property DisplayName, StatusDetail, Progress -AutoSize
    } else {
        Write-Host "==== Synced Mailboxes (95% and above) ====" -ForegroundColor Green
        Write-Host "No synced mailboxes found." -ForegroundColor DarkGray
    }

    Write-Host ""

    # Display In Progress Mailboxes
    if ($inProgress.Count -gt 0) {
        Write-Host "==== In Progress Mailboxes (Below 95%) ====" -ForegroundColor Yellow
        $inProgress | ForEach-Object {
            [PSCustomObject]@{
                DisplayName     = $_.DisplayName
                StatusDetail    = $_.StatusDetail
                Progress        = Get-ProgressBar $_.PercentComplete
            }
        } | Format-Table -Property DisplayName, StatusDetail, Progress -AutoSize
    } else {
        Write-Host "==== In Progress Mailboxes (Below 95%) ====" -ForegroundColor Yellow
        Write-Host "No mailboxes in progress." -ForegroundColor DarkGray
    }

    Start-Sleep -Seconds 30
}
