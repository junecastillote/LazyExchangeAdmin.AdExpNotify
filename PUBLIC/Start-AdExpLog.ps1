#Function to Start Transaction Logging
Function Start-AdExpLog {
    param 
    (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$logFile
    )
    Stop-AdExpLog
    Start-Transcript $logFile -Append
}