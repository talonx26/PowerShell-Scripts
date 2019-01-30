function Get-SPMACAddress
{
    [CmdletBinding()]
    [OutputType([object[]])]
    Param
    (
<#        
# Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [object[]]$ComputerInfo,
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 1)]
        [string]$ComputerID
  #>
  )

    Begin
    {
       
        $ErrorActionPreference = "Continue"
        #Edit to match full path and filename of where you want log file created
        #Load SharePoint DLL's
        $refs = @(".\Microsoft.SharePoint.Client.dll", ".\Microsoft.SharePoint.Client.Runtime.dll")
        add-type -Path $refs
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $webURL = "https://workspaces.bsnconnect.com/teams/LabAutomation"
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)
        $context.Credentials = $credentials
        $ComputerInfo = $ComputerInfo | ? { $_.model -notlike "Unable*"}
        $NetWorkIDs = @()
        $percentCounter = 0
    }  
    Process
    {
        
            $web = $Context.Web
            $weblist = "LKUPNetworkCard"
            $Context.Load($web) 
            $Context.ExecuteQuery() 

            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            
            $query.ViewXml = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            #"<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($NetWork.MAC)</Value></Eq></Where></Query></View>"
            #[Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $MAC = @()
            if ($items.count -gt 0)
            {
               $address = "" | Select Computer, MAC, Broadcast
               $address.Computer 
                #Update Record
               <#
                $s = "" ; $Network.IP | % { $s += "$_<br/>"}
               $items[0]["IP_x0020_Address"] = $s 
               $items[0]["ip_address"] = $network.ip | % {if ($_ -ne $null){if ($_.split('.').count -eq 4) {$_}} }
                $items[0]["Subnet"] = $network.subnet | % {if ($_ -ne $null){if ($_.split('.').count -eq 4) {$_}} }
                $items[0]["Gateway"] = $network.Gateway | % {if ($_ -ne $null){if ($_.split('.').count -eq 4) {$_}} }
                $items[0]["Broadcast_x0020_IP"] = $network.BroadcastIP 
                $items[0]["Computer"] = $ComputerID
                $items[0]["Network_x0020_Name"] = $NetWork.Name 
                $s = "" ; $Network.DNS | % { $s += "$_<br/>"}
                $items[0]["DNS"] = $s
                $s = "" ; $Network.DNSSearchSuffix | % { $s += "$_<br/>"}
                $items[0]["DNS_x0020_Search_x0020_Suffix"] = $s
                $ID = "" | Select Computer, ID
                $ID.Computer = $ComputerInfo.Computer
                $ID.ID = $items[0]["ID"]
                $NetworkIDs += $ID
                $items[0].update()
                $context.ExecuteQuery()
                #>
            }
            Else
            {
                # Record not found and unable to add new Record for some reason.  
                $ID = "" | Select Computer, ID
                $ID.Computer = $ComputerInfo.Computer.ToString()
                $ID.ID = "Not Found"
                $NetworkIDs += $ID
            }

	   
        
		

    }
	   
   
    End
    {
	       
        return $NetWorkIDs
    }
}

Get-SPMACAddress