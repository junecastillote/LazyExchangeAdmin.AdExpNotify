Get-ChildItem .\PUBLIC_FUNCTIONS\ | Foreach-Object {. $_.FullName}
$x = Get-AdExpUser -expireInXDays (0..40)

$props = @{
    NotifyWho = 'User','Manager'
    AdminEmailAddress = 'poshlab.admin@posh.lab'
    SenderAddress = 'Notify@posh.lab'
    SmtpServer = 'smtp.posh.lab'
    Port = 25
}
$x | Send-AdExpAlert @props