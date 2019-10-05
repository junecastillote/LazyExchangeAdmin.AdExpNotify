Function Send-AdExpAlert {
    [CmdletBinding()]
    param (
        [parameter(Mandatory, ValueFromPipeline)]
        [PSTypeName('LazyExchangeAdmin.AdExpUser')]
        $InputObject,

        # Who to notify
        [parameter(Mandatory)]
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

        [parameter()]
        [switch]$UseSSL,
        # --------------------------------
        # End Mail Params
        # --------------------------------

        [parameter()]
        [string]$CustomMessageFile
    )
    begin {
        #$PSBoundParameters
        # Validate custom message
        if ($CustomMessageFile) {
            if (!(Test-Path -Path $CustomMessageFile)) {
                Throw "The file $CustomMessageFile does not exist. Exiting script."
            }
            else {
                $customMessage = Get-Content -Path $CustomMessageFile -Raw -Encoding UTF8
            }
        }

        $html1 = @()
        $html1 += '<html><head><title></title></head><body style="font-family:Tahoma">'

    }
    process {
        Write-Host ($InputObject.Name)
        
        $html2 = '<p>'
        $html2 += ('Name: ' + ($InputObject.Name) + '<br>')
        $html2 += ('Login: ' + ($InputObject.Login) + '<br>')
        $html2 += ('User Email: ' + ($InputObject.Email) + '<br>')
        $html2 += ('Account Expiration Date: ' + ("{0:dddd, MMMM dd, yyyy hh:MM tt}" -f ($InputObject.Expiration)) + '<br>')
        $html2 += ('Manager Email: ' + ($InputObject.ManagerEmail) + '<br>')
        $html2 += '</p>'

        $body = $html1 + $html2

        if ($customMessage) {
            $body += $customMessage
        }
        $body += '</body></html>'
        $body = $body -join "`n"        

        $mailParams = @{
            Subject    = "Account [$($InputObject.Name)] will expire in $($InputObject.InDays) days"
            Body       = $body -join "`n"
            From       = $SenderAddress
            SmtpServer = $SmtpServer
            Port       = $Port    
            BodyAsHtml = $true        
        }
        
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
        
        # If UseSSL is required
        if ($UseSSL) { $mailParams += @{UseSSL = $true } }

        # If Authenitcation is required
        if ($Credential) { $mailParams += @{Credential = $Credential } }
        
        Send-MailMessage @mailParams -Verbose
    }
    end {

    }
}