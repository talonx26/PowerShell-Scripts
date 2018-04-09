#Requires -Version 3.0
function Update-SPLabPCInfo
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
     Update-SPLabInfo -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Update-SPLabInfo

.EXAMPLE
     Update-SPLabInfo -Param1 'Value1', 'Value2' -Param2 'Value'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author: Tony Turner
#>

    [CmdletBinding()]
    [OutputType('PSCustomObject')]
    param (
        [Parameter(Mandatory, 
            ValueFromPipeline)]
        [object[]]$Computers
    )

    BEGIN
    {
        #Used for prep. This code runs one time prior to processing items specified via pipeline input.
        $refs = @(".\Microsoft.SharePoint.Client.dll", ".\Microsoft.SharePoint.Client.Runtime.dll")
        add-type -Path $refs
        $webURL = "https://workspaces.bsnconnect.com/teams/LabAutomation"
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)
        $context.Credentials = $credentials
        $computers = $computers | ? { $_.model -notlike "Unable*"}
        $ModelID = Update-SPModel -Computers $Computers
        $BiosID = Update-SPBIOS -Computers $computers
    }

    PROCESS
    {
        #This code runs one time for each item specified via pipeline input.

        foreach ($computer in $Computers)
        {
            $VerbosePreference = $true
            Write-Verbose "$computer"
            $percentCounter++
            write-progress -ParentId 1 -Activity "Processing Computer $computer" -status "Updating SharePoint for $computer" -PercentComplete (($percentCounter / $Computers.count) * 100)
            Write-Verbose "Updating $computer"
            $web = $Context.Web
            $weblist = "Computer Inventory"
            $Context.Load($web) 
            $Context.ExecuteQuery() 
            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            $query.ViewXml = "<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($Computer.computer)</Value></Eq></Where></Query></View>"
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
            if ($items.count -eq 0)
            {
                $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
                #$itemCreateInfo
                $new = $list.AddItem($itemCreateInfo)
                $new["Title"] = $computer.Computer.toUpper()
                $new.Update()
                $context.ExecuteQuery()
                $context.Load($items)
                $context.ExecuteQuery()

            }
            if ($items.count -gt 0)
            {
                $items[0]["Title"] = $computer.Computer.toUpper()
                $items[0]["Memory"] = $Computer.TotalPhysicalMemory
                $items[0]["Serial_x0020_Number"] = $Computer.SerialNumber
                $items[0]["Model_x0020_Number"] = $ModelID[$ModelID.Computer.Indexof($Computer.Computer)].ID
                $items[0]["CSD"] = $computer.CSD
                $items[0]["Owner"] = $computer.Owner
                $items[0]["OwnerID"] = $computer.OwnerID
                $items[0]["Building"] = $computer.Building
                $items[0]["Description"] = $computer.Description
                $items[0]["BIOS"] = $BiosID[$BiosID.Computer.Indexof($Computer.Computer)].BIOSID
                $HDInfo = Update-SPHardDrive -ComputerInfo $computer -ComputerID $items[0]["ID"]
                $LookupCollection = @()
                foreach ($HD in $HDinfo)
                {
                    if ($hd.id -ne $null)
                    {
                        $lookupValue = New-Object Microsoft.SharePoint.Client.FieldLookupValue
                        $lookupvalue.LookupId = $HD.ID
                        $lookupCollection += $lookupValue
                    }
                }
                If ($LookupCollection.count -gt 0)
                {  
                    $items[0]["Hard_x0020_Drives"] = [Microsoft.SharePoint.Client.FieldLookupValue[]]$lookupCollection 
                }
                $NetInfo = Update-SPNetwork -ComputerInfo $computer -ComputerID $items[0]["ID"]
                $LookupCollection = @()
                foreach ($Net in $Netinfo)
                {
                    if ($Net.id -ne $null)
                    {
                        $lookupValue = New-Object Microsoft.SharePoint.Client.FieldLookupValue
                        $lookupvalue.LookupId = $Net.ID
		   
                        $lookupCollection += $lookupValue
                    }
                }
                If ($LookupCollection.count -gt 0)
                {  
                    $items[0]["Network"] = [Microsoft.SharePoint.Client.FieldLookupValue[]]$lookupCollection 
                }
                $items[0].update()
                $context.ExecuteQuery()
                Update-SPSoftware -Computers $computer -ComputerID $items[0]["ID"]

            }

            
        }
    }

    END
    {
        #Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
    }

}

#Update-SPLabPCInfo -Computers $compinfo