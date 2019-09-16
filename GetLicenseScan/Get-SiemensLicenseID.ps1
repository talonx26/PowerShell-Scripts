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
        foreach ($pn in $partnumbers)
        {
            $web = $Context.Web
            $weblist = "MLKUPSiemensSoftware"
            $Context.Load($web) 
            $Context.ExecuteQuery() 

            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            
            $query.ViewXml = "<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Part_x0020_Number'/><Value Type='Text'>$($pn.text)</Value></Eq></Where></Query></View>"
            
            #$query.viewXml = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
            If ($items.count -eq 1)
            {
                If (!($SoftwareIDs.ContainsKey($pn)))
                {
                    $SoftwareIDs.Add($pn.Text, $items[0]["ID"])
                }
            }
            
        }
    }

    END
    {
        #Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
        return $SoftwareIDs
    }

}
