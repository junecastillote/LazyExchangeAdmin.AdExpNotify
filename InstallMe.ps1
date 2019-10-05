[CmdletBinding()]
param (
    [parameter(Mandatory)]
    [string]$ModulePath
)
#$Moduleinfo = Test-ModuleManifest -Path ($PSScriptRoot+'\LazyExchangeAdmin.ExoAdminAuditLogReport.psd1')
$Moduleinfo = Test-ModuleManifest -Path ((Get-ChildItem $PSScriptRoot\*.psd1).FullName)
$ModulePath = $ModulePath + "\$($Moduleinfo.Name.ToString())\$($Moduleinfo.Version.ToString())"

if (!(Test-Path $ModulePath)) {
    New-Item -Path $ModulePath -ItemType Directory -Force | Out-Null
}

Get-ChildItem -Recurse | Unblock-File

Copy-Item -Path $PSScriptRoot\* -Include *.ps1,*.html,*.psd1,*.psm1 -Destination $ModulePath -Exclude 'installme.ps1' -Force -Confirm:$false
Copy-Item -Path $PSScriptRoot\Public\* -Destination (New-Item -ItemType Directory $ModulePath\Public -Force).FullName -Force -Confirm:$false
#Copy-Item -Path $PSScriptRoot\Private\* -Destination (New-Item -ItemType Directory $ModulePath\Private -Force).FullName -Force -Confirm:$false