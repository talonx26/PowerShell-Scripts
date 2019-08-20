#Requires -Version 3.0
function Get-SiemensLicenseCertificates
{

    <#
.SYNOPSIS
    Brief synopsis about the function.
 
.DESCRIPTION
    Detailed explanation of the purpose of this function.
 
.PARAMETER Param1
    The purpose of param1.

.PARAMETER Param2
    The purpose of param2.
 
.EXAMPLE
     Update-SiemensLicenseCertificate -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Update-SiemensLicenseCertificate

.EXAMPLE
     Update-SiemensLicenseCertificate -Param1 'Value1', 'Value2' -Param2 'Value'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author: Tony Turner
#>

    [CmdletBinding()]
    [OutputType([object[]])]
    Param
    (
        # Param1 help description
        <#
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string[]]$Images
#>      
    ) 

    Begin
    {
        $refs = @(".\Dll\Microsoft.SharePoint.Client.dll", ".\dll\Microsoft.SharePoint.Client.Runtime.dll")
        add-type -Path $refs
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('domain')
if (test-path "c:\scripts\creds\${env:username}_creds.xml")
{
    
    $creds = Import-Clixml "c:\scripts\creds\${env:username}_creds.xml"
    while (!($ds.ValidateCredentials($creds.UserName,$creds.GetNetworkCredential().password,[System.DirectoryServices.AccountManagement.ContextOptions]::Negotiate)))
    {
        $creds = Get-Credential -Message "Enter valid Sharepoint Online Credentials ex: fljpcnadmin@dow.com"
        $creds | Export-Clixml "c:\scripts\creds\${env:username}_creds.xml"
    }

}
else
{
    $creds = Get-Credential -Message "Enter Sharepoint Online Credentials ex : fljpcnadmin@dow.com"
    $creds | Export-Clixml "c:\scripts\creds\${env:username}_creds.xml"
    while (!($ds.ValidateCredentials($creds.UserName,$creds.GetNetworkCredential().password,[System.DirectoryServices.AccountManagement.ContextOptions]::Negotiate)))
    {
        $creds = Get-Credential -Message "Enter valid Sharepoint Online Credentials ex: fljpcnadmin@dow.com"
        $creds | Export-Clixml "c:\scripts\creds\${env:username}_creds.xml"
    }
}
        $ErrorActionPreference = "Continue"
        #Load SharePoint DLL's
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $webURL = "https://workspaces.bsnconnect.com/teams/LabAutomation"
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)
        $context.Credentials = $credentials
        $DiskLocation = "C:\Siemens"
    }
    Process
    {
        
            $web = $Context.Web
            $weblist = "Siemens License Tracker"
            $Context.Load($web) 
            $Context.ExecuteQuery() 

            $list = $web.Lists.GetByTitle($weblist)
           # $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            
            #$query.ViewXml = "<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Serial_x0020_Number'/><Value Type='Text'>$($SerialNumber)</Value></Eq></Where></Query></View>"
            
           # $query.viewXml = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            $items = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())  
            $context.Load($items)
            $context.ExecuteQuery()
            If ($items.count -gt 0)
            {
                foreach ($item in $items)
                {

                    $attachments = $item.AttachmentFiles
                    $context.Load($attachments)
                    $context.ExecuteQuery()
                    foreach ( $att in $item.AttachmentFiles)
                    {
                         $fileRef = $att.ServerRelativeURL
                         $fileinfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($context,$fileRef)
                         New-Item -ItemType Directory -Force -Path "$DiskLocation\$($item["Software_x003a_Part_x0020_Number"].LookupValue)"
                         $fileName = $DiskLocation + "\" + $item["Software_x003a_Part_x0020_Number"].LookupValue + "\"+  $att.fileName
                         $fileStream = [System.IO.File]::Create($fileName)
                         $fileinfo.Stream.CopyTo($fileStream)
                         $FileStream.close()
                    }
                }
            }
            
    } 
        

	   
   
    End
    {
	       
        # return $NetWorkIDs
    }
}
