#Requires -Version 3.0
function Update-SPSoftware
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [object]$Computers,
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 1)]
        [string]$ComputerID
    )

    Begin
    {
        write-debug "Updating Software's for $($computerinfo.computer)"
        $ErrorActionPreference = "Continue"
        #Edit to match full path and filename of where you want log file created
        #Load SharePoint DLL's

        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $webURL = "https://workspaces.bsnconnect.com/teams/LabAutomation"
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)
        $context.Credentials = $credentials
        $computers = $computers | ? { $_.model -notlike "Unable*"}
        $SoftwareIDs = @()
    }
    Process
    {
        foreach ($Software in $Computers.Software)
        {
            $VerbosePreference = $false
            $softwareID = -1
            $SoftwareID = Get-SPMasterSoftwareID -Software $Software #-Verbose
            Write-Verbose "Software ID : $softwareid"
            $web = $Context.Web
            $webList = "LKUPSoftware"
             
            $Context.Load($web) 
            $Context.ExecuteQuery() 
            $list = $web.Lists.GetByTitle($weblist)
            $context.load($list)
            $context.load($list.Fields)
            $context.ExecuteQuery()
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            
            $Query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><And><Eq><FieldRef Name='SoftwareID'/><Value Type='Text'>$SoftwareID</Value></Eq><Eq><FieldRef Name='ComputerID'/><Value Type='Text'>$ComputerID</Value></Eq></And></Where></Query></View>"
            #$Query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='ComputerID'/><Value Type='Text'>5</Value></Eq></Where></Query></View>"
            #$items = new-object Microsoft.SharePoint.Client.ListItemCollection($list.GetItems($Query))
            $items = $list.GetItems($Query)  
            $context.Load($items)

           # $context.Load($items,li => li.include(pi => pi.id))
            $context.ExecuteQuery()
            Write-Verbose $Software
            If ($items.count -eq 0)
            {
                #Record not found.  Create initial Record
                Write-Verbose "Record not Found  $($Software.ProdName) ::   $($software.VersionString)"
                $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
                #$itemCreateInfo
                $new = $list.AddItem($itemCreateInfo)
                
                $new["SoftwareID"] = $softwareid
                $new["ComputerID"] = $ComputerID
                Write-Verbose $new.FieldValues
               # $new.FieldValues
                $new.Update()
                $Context.ExecuteQuery()
                #Reload Items to get new Record ID
                $context.Load($items)
                $context.ExecuteQuery()
            }
            if ($items.count -eq 1)
            {
                $ids = "" | Select ComputerID, SoftwareID
                $ids.ComputerID = $ComputerID
                $ids.SoftwareID = $items[0]["ID"] 
                $SoftwareIDs += $ids 
            }
        }
  }

  End
    {
	       
        return $SoftwareIDs
    }
}
