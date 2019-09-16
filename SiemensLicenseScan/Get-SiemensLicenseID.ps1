#Requires -Version 3.0
function Get-SiemensLicenseID
{

    <# Synopsis
   #>
    [CmdletBinding()]
    [OutputType('PSCustomObject')]
    param (
        [Parameter(Mandatory, 
            ValueFromPipeline)]
        [object[]]$PartNumbers
    )

    BEGIN
    {
        $refs = @(".\Dll\Microsoft.SharePoint.Client.dll", ".\dll\Microsoft.SharePoint.Client.Runtime.dll")
        add-type -Path $refs
        Add-Type -AssemblyName PresentationCore, PresentationFramework
#set-location c:\scripts
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
        $webURL = "https://workspaces.bsnconnect.com/sites/LabAuto/Inventory"
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)
        $context.Credentials = $credentials
        $SoftwareIDs = @{}
    }
    Process
    {
        foreach ($partnumber in $partnumbers)
        {
            $web = $Context.Web
            $weblist = "MLKUPSiemensSoftware"
            $Context.Load($web) 
            $Context.ExecuteQuery() 
            #If PartNumber.Text is Null then object is string is Barcode object  
            if ($partnumber.Text -eq $null)
            { $pn = $partnumber}
            else
            {
                #PartNumber.Text is not Null so object is a Barcode Object
                $pn = $partnumber.text
            }
            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            
            $query.ViewXml = "<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Part_x0020_Number'/><Value Type='Text'>$pn</Value></Eq></Where></Query></View>"
            
            #$query.viewXml = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
            If ($items.count -eq 1)
            {
                If (!($SoftwareIDs.ContainsKey($pn)))
                {
                    $SoftwareIDs.Add($pn, $items[0]["ID"])
                }
            }
            
        }
    }

    END
    {
        #Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
        Write-Verbose $SoftwareIDs -Verbose
        return $SoftwareIDs
    }

}
