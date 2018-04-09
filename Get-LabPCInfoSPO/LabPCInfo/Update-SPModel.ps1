#Requires -Version 3.0

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
function Update-SPModel
{
    [CmdletBinding()]
	
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [object[]]$Computers
    )

    Begin
    {
        $ErrorActionPreference = "Continue"
        #Edit to match full path and filename of where you want log file created
        #Load SharePoint DLL's
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $webURL = "https://workspaces.bsnconnect.com/teams/LabAutomation"
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)
        $context.Credentials = $credentials
        # $Context.Credentials = $creds
        $computers = $computers | ? { $_.model -notlike "Unable*"}
        $ModelID = @()
        $percentCounter = 0
    }
    Process
    {
        foreach ($computer in $Computers)
        {
            $percentCounter++
            write-progress -ParentId 2 -Activity "Processing Bios Information for $computer" -status "Updating SharePoint for $computer" -PercentComplete (($percentCounter / $Computers.count) * 100)
           
            $web = $Context.Web
            $weblist = "LKUPModel"
            $Context.Load($web) 
            $Context.ExecuteQuery() 

            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            $Model = $computer.Model #.Replace(" ","_x0020_")
            $query.ViewXml = "<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$Model</Value></Eq></Where></Query></View>"
            #[Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
        

            if ($items.count -eq 0)
            {
                # No Data info found.  Add new computer info
                $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
                #$itemCreateInfo
                $new = $list.AddItem($itemCreateInfo)
                $new["Title"] = $computer.Model
                $new["Hardware_x0020_Vendor"] = $Computer.Mfg
                $new.Update()
                $Context.ExecuteQuery()
                # Reload Query to get new Item
                $context.Load($items)
                $context.ExecuteQuery()

            }
            Else
            {
                $ID = "" | Select Computer, ID
                $ID.Computer = $computer.Computer
                If ($items[0]["ID"] -ne $null)
                {
                    $ID.ID = $items[0]["ID"]
                }
                else
                {
                    #No Data Found
                    $ID.ID = "Not Found"
                }
                $modelID += $ID

		
            }
		

        }
	   
		
    }
    End
    {
        write-progress -ParentId 2 -Activity "Processing Bios Information for $computer" -status "Updating SharePoint for $computer" -PercentComplete 100
            
        Return $ModelID
    }
}