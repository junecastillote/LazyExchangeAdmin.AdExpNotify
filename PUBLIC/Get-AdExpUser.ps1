
Function Get-AdExpUser {
    [CmdletBinding()]
    param (
        # Days before expiration threshold
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [int[]]$expireInXDays
    )

    $today = Get-Date # -Hour 0 -Minute 0 -Second 0
    $oldest = ($expireInXDays | Sort-Object)[-1]
    $expiringAccounts = Search-ADAccount -UsersOnly -AccountExpiring -TimeSpan "$($oldest).00:00:00"

    $finalResult = @()
    foreach ($account in $expiringAccounts) {
        $user = Get-AdUser $account -Properties EmailAddress, Manager, Name, SamAccountName
        
        # Get Manager's email if $NotifyWho includes Manager.
        $managerEmail = ""            
        try {
            $managerEmail = (Get-AdUser ($user.Manager) -Properties EmailAddress).EmailAddress
        }
        catch {
            $managerEmail = ""
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
    return $finalResult
}