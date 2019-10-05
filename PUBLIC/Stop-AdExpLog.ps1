#Function to Stop Transaction Logging
Function Stop-AdExpLog {
    $txnLog = ""
    Do {
        try {
            Stop-Transcript | Out-Null
        } 
        catch [System.InvalidOperationException] {
            $txnLog = "stopped"
        }
    } While ($txnLog -ne "stopped")
}
