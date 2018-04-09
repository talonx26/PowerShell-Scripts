#Requires -Version 3.0
function Get-SPMasterSoftwareID
{
    [CmdletBinding()]
   
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $false,
            Position = 0)]
        [object]$Software
       
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
     
        # $computers = $computers | ? { $_.model -notlike "Unable*"}
       
     
    }
    Process
    {
        <# foreach ($Software in $ComputerInfo.Software)
        { #>
        $web = $Context.Web
        $weblist = "MLKUPSoftware"
        $Context.Load($web) 
        $Context.ExecuteQuery() 
        $qry = @()
        if ($Software.ProdName -ne $null) { $qry += "<Eq><FieldRef Name='Title'/><Value Type='Text'>$($Software.ProdName)</Value></Eq>"}
        If ($Software.ProdGroup -ne $null) { $qry += "<Eq><FieldRef Name='Product_x0020_Group'/><Value Type='Text'>$($Software.ProdGroup)</Value></Eq>"}
        if ($Software.VersionString -ne $null) { $qry += "<Eq><FieldRef Name='Software_x0020_Version'/><Value Type='Text'>$($Software.VersionString)</Value></Eq>"}
        if ($Software.Release -ne $null) { $qry += "<Eq><FieldRef Name='Release_x0020_Version'/><Value Type='Text'>$($Software.Release)</Value></Eq>"}
        if ($Software.TechnVersion -ne $null) { $qry += "<Eq><FieldRef Name='Technical_x0020_Version'/><Value Type='Text'>$($Software.TechnVersion)</Value></Eq>"}
        $list = $web.Lists.GetByTitle($weblist)
        $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
        Switch ( $qry.count)
        {
            1 {$query.ViewXml = "<View Scope='RecursiveAll'><Query><Where>$($qry[0])</Where></Query></View>" }
            2 {$query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><And>$($qry[0])$($qry[1])</And></Where></Query></View>" }
            3 {$query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><And>$($qry[2])<And>$($qry[0])$($qry[1])</And></And></Where></Query></View>" }
            4 {$query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><And>$($qry[3])<And>$($qry[2])<And>$($qry[0])$($qry[1])</And></And></And></Where></Query></View>" }
            5 {$query.ViewXml = "<View Scope='RecursiveAll'><Query><Where><And>$($qry[4])<And>$($qry[3])<And>$($qry[2])<And>$($qry[0])$($qry[1])</And></And></And></And></Where></Query></View>" }
        }
    
        $items = $list.GetItems($Query)  
        $context.Load($items)
        $context.ExecuteQuery()
        Write-Verbose $Software
        If ($items.count -eq 0)
        {
            #Record not found.  Create initial Record
            Write-Verbose "Record not Found  $($Software.ProdName) ::   $($software.VersionString)"
            $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
            #$itemCreateInfo
            $new = $list.AddItem($itemCreateInfo)
            $new["Title"] = $software.ProdName
            $new["Release_x0020_Version"] = $Software.Release
            $new["Software_x0020_Version"] = $Software.VersionString
            $new["Technical_x0020_Version"] = $Software.TechnVersion
            $new["Product_x0020_Group"] = $Software.ProdGroup
            $new["Vendor"] = "Siemens"
            $new.Update()
            $Context.ExecuteQuery()
            #Reload Items to get new Record ID
            $context.Load($items)
            $context.ExecuteQuery()
        }
            
        if ($items.count -eq 1) {$id = $items[0]["ID"]}
        else
        {
            $id = -1
        }
           	   
        # }
		

    }
	   
   
    End
    {
        Write-Verbose "ID : $id"   
        return $id
    }
}