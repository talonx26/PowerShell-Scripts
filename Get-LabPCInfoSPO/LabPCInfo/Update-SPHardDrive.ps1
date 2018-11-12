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
function Update-SPHardDrive
{
    [CmdletBinding()]
    [OutputType([object[]])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [object[]]$ComputerInfo,
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 1)]
        [string]$ComputerID
    )

    Begin
    {
        write-debug "Updating HD's for $($computerinfo.computer)"
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
      
        $ComputerInfo = $ComputerInfo | ? { $_.model -notlike "Unable*"}
        $HardDriveIDs = @()
        # Remove-Variable HDinfo -scope global #-ErrorAction SilentlyContinue
        #New-Variable HDInfo -Scope Global
        $percentCounter = 0
    }
    Process
    {
        foreach ($HD in $ComputerInfo.HD)
        {
            $percentCounter++
            write-progress -ParentId 2 -Activity "Processing Hard Drive Information for $computer" -status "Updating SharePoint for $computer" -PercentComplete (($percentCounter /($ComputerInfo.HD|measure-object).count) * 100)
           
            $web = $Context.Web
            $weblist = "LKUPHardDrives"
            $Context.Load($web) 
            $Context.ExecuteQuery() 

            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            $query.ViewXml = "<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($HD.VolPath)</Value></Eq></Where></Query></View>"
            #[Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
            If ($items.count -eq 0)
            {
                $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
                #$itemCreateInfo
                $new = $list.AddItem($itemCreateInfo)
                $new["Title"] = $HD.VolPath
			
                $new.Update()
                $Context.ExecuteQuery()
                $context.Load($items)
                $context.ExecuteQuery()
            }
            if ($items.count -gt 0)
            {
		   
                $items[0]["Total_Size"] = $hd.TotalSize
                $items[0]["Free_x0020_Space"] = $hd.Freespace
                $items[0]["Drive"] = $hd.drive
                $items[0]["Computer"] = $ComputerID
                $ID = "" | Select-Object Computer, ID
                $ID.Computer = $ComputerInfo.Computer
                $ID.ID = $items[0]["ID"]
                $HardDriveIDs += $ID
                $items[0].update()
                $context.ExecuteQuery()
            }
            Else
            {
                # No Data info found.  Add new computer info 
		   
                $ID = "" | Select Computer, ID
                $ID.Computer = $ComputerInfo.Computer.ToString()
                $ID.ID = "Not Found"
                $HardDriveIDs += $ID
            }

	   
        }
		

    }
	   
   
    End
    {
        write-progress -ParentId 2 -Activity "Processing Hard Drive Information for $computer" -status "Updating SharePoint for $$" -PercentComplete 100
           
      
        return $HardDriveIDs

	
    }
}


