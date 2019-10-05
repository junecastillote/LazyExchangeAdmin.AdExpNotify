Function Send-AdExpAlert {
    [CmdletBinding()]
    param (
        [parameter(Mandatory, ValueFromPipeline)]
        [PSTypeName('LazyExchangeAdmin.AdExpUser')]
        $InputObject,

        # Who to notify
        [parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('User', 'Manager')]
        [string[]]$NotifyWho,

        # Admin email address
        # This is always required, in case the Manager or User email address is missing.
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [mailaddress[]]$AdminEmailAddress,

        # --------------------------------
        # Start Mail Params
        # --------------------------------
        # Sender email address
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [mailaddress]$SenderAddress,

        # SMTP Server IP/FQDN/HOSTNAME
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SmtpServer,

        # SMTP Port. Default is 25 if not otherwise specified.
        [parameter()]
        [ValidateNotNullOrEmpty()]
        [int]$Port = 25,

        # SMTP Authentication Credential if required
        [parameter()]
        [ValidateNotNullOrEmpty()]
        [pscredential]$Credential,

        # Switch if SSL is required
        [parameter()]
        [switch]$UseSSL,

        # Additional Recipients (CC)
        [parameter()]
        [ValidateNotNullOrEmpty()]
        [mailaddress[]]$AdditionalRecipient,

        # Custom HTML message to append to the original email
        [parameter()]
        [string]$CustomMessage
        # --------------------------------
        # End Mail Params
        # --------------------------------
    )
    begin {
        
    }
    process {
        Write-Verbose "$($InputObject.Name), expires in $($InputObject.InDays) days"

        $body = @()
        $body += '<html><head><title></title></head><body style="font-family:Tahoma"><hr>'
        $body += '<h3>User Details</h3>'
        $body += '<p>'
        $body += ('<b>Account Expiration Date:</b> <font color="red">' + ("{0:dddd, MMMM dd, yyyy hh:MM tt}" -f ($InputObject.Expiration)) + '</font><br>')
        $body += ('<b>Name:</b> ' + ($InputObject.Name) + '<br>')
        $body += ('<b>Login:</b> ' + ($InputObject.Login) + '<br>')
        $body += ('<b>User Email:</b> ' + ($InputObject.Email) + '<br>')
        $body += ('<b>Manager Email:</b> ' + ($InputObject.ManagerEmail) + '<br>')
        $body += '</p><hr>'
        
        $body += '<h3>Notification</h3>'
        $body += '<p>'
    
        # Is user notified?
        if (($InputObject.Email) -and ($NotifyWho -contains 'User')) {
            $body += ('<b>User:</b> Yes<br>')
        }
        else {
            $body += ('<b>User:</b> No<br>')
        }        

        # Is manager notified?
        if (($InputObject.ManagerEmail) -and ($NotifyWho -contains 'Manager')) {
            $body += ('<b>Manager:</b> Yes<br>')
        }
        else {
            $body += ('<b>Manager:</b> No<br>')
        }

        $body += ('<b>Admin:</b> Yes<br>')
        $body += '</p><hr>'

        # Append CustomMessage if present
        if ($customMessage) {
            $body += $customMessage
        }

        $body += '</body></html>'

        # Build email parameters
        $mailParams = @{
            Subject    = "Account [$($InputObject.Name)] will expire in $($InputObject.InDays) days"
            Body       = $body -join "`n"
            From       = $SenderAddress
            SmtpServer = $SmtpServer
            Port       = $Port    
            BodyAsHtml = $true        
        }
        
        # --------------------------------     
        # Start 'To' Recipients
        # --------------------------------
        $To = @($AdminEmailAddress)

        # Add Manager Email as recipient
        if (($InputObject.ManagerEmail) -and ($NotifyWho -contains 'Manager')) {
            $To += ($InputObject.ManagerEmail)
        }

        # Add User Email as recipient
        if (($InputObject.Email) -and ($NotifyWho -contains 'User')) {
            $To += ($InputObject.Email)
        }

        # Add recipients to hash
        $mailParams += @{To = $To }
        # --------------------------------     
        # End 'To' Recipients
        # --------------------------------

        # --------------------------------     
        # Start 'CC' Recipients
        # --------------------------------
        if ($AdditionalRecipient) {
            $mailParams += @{cc = $AdditionalRecipient }
        }
        # --------------------------------
        # End 'To' Recipients
        # --------------------------------
        
        # If UseSSL is required
        if ($UseSSL) { $mailParams += @{UseSSL = $true } }

        # If Authenitcation is required
        if ($Credential) { $mailParams += @{Credential = $Credential } }
        
        try {
            Send-MailMessage @mailParams -Verbose
        }
        catch {
            Write-Host $_.Exception.Message
        }        
    }
    end {
        
    }
}