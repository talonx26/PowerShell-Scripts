#Requires -Version 3.0
function Sync-SPSoftware 
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
     Sync-SPSoftware -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Sync-SPSoftware

.EXAMPLE
     Sync-SPSoftware -Param1 'Value1', 'Value2' -Param2 'Value'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author:  Mike F Robbins
    Website: http://mikefrobbins.com
    Twitter: @mikefrobbins
#>

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


    BEGIN {
        #Used for prep. This code runs one time prior to processing items specified via pipeline input.
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
       
     
    }

    PROCESS {
        #This code runs one time for each item specified via pipeline input.
            $web = $Context.Web
            $webList = "LKUPSoftware"
             
            $Context.Load($web) 
            $Context.ExecuteQuery() 
            $list = $web.Lists.GetByTitle($weblist)
            $context.load($list)
            $context.load($list.Fields)
            $context.ExecuteQuery()
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            # Query for all software records for this machine by ComputerID
            $Query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='ComputerID'/><Value Type='Text'>$ComputerID</Value></Eq></Where></Query></View>"
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
            $colFound = @()
        for ($i = 0; $i -lt $items.count; $i++) {
            #Use foreach scripting construct to make parameter input work the same as pipeline input (iterate through the specified items one at a time).
            # Create Object with info from Query to compare to $computers.Software
            $item = $items[$i]
            $found = "" | Select-Object ProdName, Version, Found
            $s = "" | Select-Object ProdGroup, ProdName, VersionString, Release, TechnVersion
            $s.ProdGroup = $item["SoftwareID_x003a_Product_x0020_G"].LookupValue
            $s.ProdName = $item["SoftwareID_x003a_Title"].LookupValue
            $s.VersionString = $item["SoftwareID_x003a_Software_x0020_"].LookupValue
            $s.Release = $item["SoftwareID_x003a_Release_x0020_V"].LookupValue
            $s.TechnVersion =$item["SoftwareID_x003a_Technical_x0020"].LookupValue

            $found.ProdName = $s.ProdName
            $found.Version  = $s.VersionString
            #$s
            if ($computers.Software -like $s) { 
                $found.Found = $true } 
                else { 
                    $items[$i].DeleteObject()
                    $context.ExecuteQuery()
                    $found.found = $false
                        }
            $colFound += $found
                    }
                  
             
             #$colFound | Format-Table
             }
                
    END {
        #Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
    }

}

<#
$dataValues = @()
$items.GetEnumerator() | % { 
    $dataValues += $_.FieldValues 
}


#>
<#
$colFound = @()
foreach ($item in $items)
{
#$item = $items[64]
$found = "" | Select ProdName, Version, Found
$s = "" | Select ProdGroup, ProdName, VersionString, Release, TechnVersion
$s.ProdGroup = $item["SoftwareID_x003a_Product_x0020_G"].LookupValue
$s.ProdName = $item["SoftwareID_x003a_Title"].LookupValue
$s.VersionString = $item["SoftwareID_x003a_Software_x0020_"].LookupValue
$s.Release = $item["SoftwareID_x003a_Release_x0020_V"].LookupValue
$s.TechnVersion =$item["SoftwareID_x003a_Technical_x0020"].LookupValue

$found.ProdName = $s.ProdName
$found.Version  = $s.VersionString
if ($compinfo[0].Software -like $s) { 
    $found.Found = $true } 
    else { $found.found = $false}
$colFound += $found
}
$colFound | ft
<#
$item = $items[0]
$s = "" | Select ProdGroup, ProdName, VersionString, Release, TechnVersion
$s.ProdGroup = $item["SoftwareID_x003a_Product_x0020_G"].LookupValue
$s.ProdName = $item["SoftwareID_x003a_Title"].LookupValue
$s.VersionString = $item["SoftwareID_x003a_Software_x0020_"].LookupValue
$s.Release = $item["SoftwareID_x003a_Release_x0020_V"].LookupValue
$s.TechnVersion =$item["SoftwareID_x003a_Technical_x0020"].LookupValue
$s
if ($compinfo[0].Software -like $s) { Write-Host "Found" } else { Write-host "not Found"}
#>
