
Function Get-AdExpUser {
    [CmdletBinding()]
    param (
        # Days before expiration threshold
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [int[]]$expireInXDays
    )

    # Start-Log -LogFile "$($env:windir)\temp\LazyExchangeAdmin.Get-AdExpUser.log"

    $today = Get-Date

    $oldest = ($expireInXDays | Sort-Object)[-1]
    $expiringAccounts = Search-ADAccount -UsersOnly -AccountExpiring -TimeSpan "$($oldest).$($today.Hour):$($today.Minute):$($today.Second)"

    $finalResult = @()
    foreach ($account in $expiringAccounts) {
        $user = Get-AdUser $account -Properties EmailAddress, Manager, Name, SamAccountName
        
        # Get Manager's email if $NotifyWho includes Manager.
        $managerEmail = ""
        if ($user.Manager) {
            try {
                $managerEmail = (Get-AdUser ($user.Manager) -Properties EmailAddress).EmailAddress
            }
            catch {
                $managerEmail = ""
            }
        }        
        
        # Calculate days left before account expiration (round-off)
        [int]$daysLeftToExpire = [int](New-TimeSpan -Start $today -End ($account.AccountExpirationDate)).TotalDays
        
        if ($expireInXDays -contains $daysLeftToExpire) {           
            
            # Build Expiring User Account Collection
            $tempObj = New-Object psobject -Property @{
                PSTypeName   = 'LazyExchangeAdmin.AdExpUser'
                Name         = $user.Name
                Login        = $user.SamAccountName
                Email        = $user.EmailAddress
                ManagerEmail = $managerEmail
                Expiration   = $account.AccountExpirationDate
                InDays       = $daysLeftToExpire
            }
            $finalResult += $tempObj
        }
    }
    # Stop-Log
    return $finalResult
}