## THIS IS AN EXAMPLE RUN FILE ##

# Unload module (if presently loaded)
if (Get-Module LazyExchangeAdmin.AdExpNotify) {Remove-Module LazyExchangeAdmin.AdExpNotify}
#& .\InstallMe.ps1 -ModulePath "C:\Program Files\WindowsPowerShell\Modules"

# Import Module
Import-Module LazyExchangeAdmin.AdExpNotify

# Start transaction logging
Start-AdExpLog -LogFile "$($env:windir)\temp\LazyExchangeAdmin.AdExpNotify.log"

# Notification parameters
$props = @{
    NotifyWho = 'Manager'
    AdminEmailAddress = 'admin@posh.lab'
    SenderAddress = 'AdExpNotify@posh.lab'
    SmtpServer = 'smtp.posh.lab'
    Port = 25
    AdditionalRecipient = 'ServiceDesk@posh.lab'
    CustomMessage = (Get-Content $PSScriptRoot\CustomMessage.HTML -Raw -Encoding UTF8 )
}

# Get expiring account in days (30,6,1,0), then send the notification
Get-AdExpUser -expireInXDays 30,6,1,0 -Verbose | Send-AdExpAlert @props -Verbose

# Stop transaction logging
Stop-AdExpLog