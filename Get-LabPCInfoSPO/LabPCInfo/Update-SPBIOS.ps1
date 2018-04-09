#Requires -Version 3.0
function Update-SPBIOS
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
     Update-SPBIOS -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Update-SPBIOS

.EXAMPLE
     Update-SPBIOS -Param1 'Value1', 'Value2' -Param2 'Value'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author: Tony Turner
#>

    [CmdletBinding()]
    [OutputType('PSCustomObject')]
    param(
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [object[]]$Computers
        
    )
    BEGIN
    {
        #Used for prep. This code runs one time prior to processing items specified via pipeline input.
        
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
        $BIOSIDS = @()
        $percentCounter = 0
    }

    PROCESS
    {
        foreach ($computer in $Computers)
        {
            $web = $Context.Web
            $webList = "LKUPBIOS"
             
            $Context.Load($web) 
            $Context.ExecuteQuery() 
            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            $Query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><And><Eq><FieldRef Name='BIOS_x0020_Release_x0020_Date'/><Value Type='Date'>$($computer.BIOSReleaseDate)</Value></Eq><And><Eq><FieldRef Name='Title'/><Value Type='Text'>$($computer.BIOS)</Value></Eq><Eq><FieldRef Name='BIOS_x0020_Version'/><Value Type='Text'>$($computer.BiosVersion)</Value></Eq></And></And></Where></Query></View>"
           
    
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
            
            If ($items.count -eq 0)
            {
                #Record not found.  Create initial Record
                
                $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
                #$itemCreateInfo
                $new = $list.AddItem($itemCreateInfo)
                
                $new["Title"] = $computer.BIOS
                $new["BIOS_x0020_Version"] = $Computer.BiosVersion
                $new["BIOS_x0020_Release_x0020_Date"] = [datetime]$computer.BIOSReleaseDate
                $new.Update()
                $Context.ExecuteQuery()
                #Reload Items to get new Record ID
                $context.Load($items)
                $context.ExecuteQuery()
            }
            if ($items.count -eq 1)
            {
                $ids = "" | Select Computer, BIOSID
                $ids.Computer = $Computer.computer
                $ids.BIOSID = $items[0]["ID"] 
                $BIOSIDs += $ids 
            
            }
           	   
        }
		

    }
	   
   
    End
    {
	       
        return $BIOSIDS
    }

}
