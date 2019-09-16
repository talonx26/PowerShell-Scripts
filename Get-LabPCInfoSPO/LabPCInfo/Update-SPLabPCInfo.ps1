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
        #Load Sharepoint Library Files
        $refs = @(".\Microsoft.SharePoint.Client.dll", ".\Microsoft.SharePoint.Client.Runtime.dll")
        add-type -Path $refs
        # Sharepoint Web Address and Login information
        $webURL = "https://workspaces.bsnconnect.com/sites/LabAuto/Inventory"
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)
        $context.Credentials = $credentials
        # Clear any computer objects with Model containing 'Unable'  - clears any objects with no data.
        $computers = $computers | ? { $_.model -notlike "Unable*"}

        #Calls function to Update Bios and Model Lists in Sharepoint
        #Function will got thru entire object and update lists and then return an master object with the Record Ids to be used
        #to link in the Computer Inventory List
        $ModelID = Update-SPModel -Computers $Computers
        $BiosID = Update-SPBIOS -Computers $computers
    }

    PROCESS
    {
        #This code runs one time for each item specified via pipeline input.

        foreach ($computer in $Computers)
        {
            #$VerbosePreference = $true
            Write-Verbose "$computer"
            $percentCounter++
            Write-Verbose "writing progress"
            write-progress -ParentId 1 -Activity "Processing Computer $computer" -status "Updating SharePoint for $computer" -PercentComplete (($percentCounter / ($Computers | Measure-Object).count) * 100)
            Write-Verbose "Updating $computer"
            $web = $Context.Web
            #Sharepoint list to get information from
            $weblist = "Computer Inventory"
            $Context.Load($web) 
            $Context.ExecuteQuery() 
            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            #Sharepoint Query
            $query.ViewXml = "<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($Computer.computer)</Value></Eq></Where></Query></View>"
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
            # If query came up with 0 items then record doesn't exist.  Create initial record with minimum required data
            if ($items.count -eq 0)
            {
                $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
                #$itemCreateInfo
                $new = $list.AddItem($itemCreateInfo)
                $new["Title"] = $computer.Computer.toUpper()
                $new.Update()
                $context.ExecuteQuery()
                #Run query again to get new record
                $context.Load($items)
                $context.ExecuteQuery()

            }
            # New record or existing record found.  Update all data
            if ($items.count -gt 0)
            {
                $items[0]["Title"] = $computer.Computer.toUpper()
                $items[0]["Memory"] = $Computer.TotalPhysicalMemory
                $items[0]["Serial_x0020_Number"] = $Computer.SerialNumber
                $items[0]["Model_x0020_Number"] = $ModelID[$ModelID.Computer.Indexof($Computer.Computer)].ID
                $items[0]["CSD"] = $computer.CSD
                $items[0]["City"] =$computer.city
                $items[0]["Owner"] = $computer.Owner
                $items[0]["OwnerID"] = $computer.OwnerID
                $items[0]["Building"] = $computer.Building
                $items[0]["Description"] = $computer.Description
                $items[0]["BIOS"] = $BiosID[$BiosID.Computer.Indexof($Computer.Computer)].BIOSID
               
                # Call function to update Hard drive Sharepoint list and return with Record Id's
                $HDInfo = Update-SPHardDrive -ComputerInfo $computer -ComputerID $items[0]["ID"]
                
                #Create Lookupcollection with Record Id's from Hard drive table to link to Computer List table
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
                #If collection has records then add record to Computer list object.
                If ($LookupCollection.count -gt 0)
                {  
                    $items[0]["Hard_x0020_Drives"] = [Microsoft.SharePoint.Client.FieldLookupValue[]]$lookupCollection 
                }
                #Call function to update Network Sharepoint list and return with Record Id's
                $NetInfo = Update-SPNetwork -ComputerInfo $computer -ComputerID $items[0]["ID"]
                #Create Lookupcollection with Record Id's from Network table to link to Computer List table
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
                #If collection has records then add record to Computer list object.
                If ($LookupCollection.count -gt 0)
                {  
                    $items[0]["Network"] = [Microsoft.SharePoint.Client.FieldLookupValue[]]$lookupCollection 
                }
                # Update all information from above to Sharepoint
                $items[0].update()
                $context.ExecuteQuery()
                #Call process to update Software information
                Update-SPSoftware -Computers $computer -ComputerID $items[0]["ID"] |Out-Null

            }

            
        }
    }

    END
    {
        #Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
    }

}

#Update-SPLabPCInfo -Computers $compinfo