<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-SharePointCredentials
{
    [CmdletBinding()]
    [Alias()]
    [OutputType()]
    Param
    (
        # Param1 help description
       # [Parameter(Mandatory=$Fa,
       #            ValueFromPipelineByPropertyName=$true,
       #            Position=0)]
       # $Param1

        # Param2 help description
#        [int]
 #       $Param2
    )

    Begin
    {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('domain')
    }
    Process
    {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('domain')
        if (test-path "c:\scripts\creds\${env:username}_spo_creds.xml")
        {
            
            $creds = Import-Clixml "c:\scripts\creds\${env:username}_spo_creds.xml"
           while (!($ds.ValidateCredentials($creds.UserName,$creds.GetNetworkCredential().password,[System.DirectoryServices.AccountManagement.ContextOptions]::Negotiate)))
            {
                $creds = Get-Credential -Message "Enter valid Sharepoint Online Credentials ex: fljpcnadmin@dow.com"
                $creds | Export-Clixml "c:\scripts\creds\${env:username}_spo_creds.xml"
            }

        }
        else
        {
            $creds = Get-Credential -Message "Enter Sharepoint Online Credentials ex : fljpcnadmin@dow.com"
            $creds | Export-Clixml "c:\scripts\creds\${env:username}_spo_creds.xml"
           while (!($ds.ValidateCredentials($creds.UserName,$creds.GetNetworkCredential().password,[System.DirectoryServices.AccountManagement.ContextOptions]::Negotiate)))
            {
                $creds = Get-Credential -Message "Enter valid Sharepoint Online Credentials ex: fljpcnadmin@dow.com"
                $creds | Export-Clixml "c:\scripts\creds\${env:username}_spo_creds.xml"
            }
        }
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)
    }
    End
    {
        $credentials
    }
}