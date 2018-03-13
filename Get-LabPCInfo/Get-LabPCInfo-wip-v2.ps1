$TargetComputer = "wpcs7lj013ss1"


#psexec \\10.100.65.222 netsh advfirewall firewall set rule group="Windows Management Instrumentation (WMI)" new enable=yes


#region Get-LabPCInfo
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
function Get-LabPCInfo {
    [CmdletBinding()]
	
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string[]]$Computers = $env:COMPUTERNAME
    )

    Begin {
        $ComputerInfo = @()
    }
    Process {
        Foreach ($computer in $computers) {
            try {
        
                Write-Host
                write-host $Computer
                write-host
                If (-not (Test-Connection -ComputerName $computer -count 1 -ErrorAction SilentlyContinue))
                {  throw "Can't ping $computer"}
                # Get WMI Data
                $info = "" | Select Computer, Mfg, Model, TotalPhysicalMemory, TotalMemoryGB, BIOS, BIOSVersion, BIOSReleaseDate, SerialNumber, OS, ServicePack, OSArchitecture, OSBuildNumber, HD, Network, Software
                $OS = Get-WmiObject win32_operatingsystem -ComputerName $Computer  -ErrorAction SilentlyContinue| select Caption, CSDVersion, OSArchitecture, Version
                if ($os -eq $null)
                {  throw "WMI not accessible on $computer"}
                $computersystem = Get-WmiObject win32_computersystem -ComputerName $Computer -ErrorAction SilentlyContinue| select Name, Manufacturer, Model, TotalPhysicalMemory, @{n = "TotalMemory(GB)"; e = {[math]::Round($_.TotalPhysicalMemory / 1GB, 3)}}
                $HDInfo = Get-WmiObject -query  "select * from win32_logicaldisk where DriveType = '3'" -ComputerName $computer -ErrorAction SilentlyContinue
                $bios = Get-WmiObject  win32_bios -ComputerName $computer -ErrorAction SilentlyContinue | select Name, Version, ReleaseDate, SerialNumber
                $adapters = Get-WmiObject Win32_NetworkAdapter -ComputerName $computer  -ErrorAction SilentlyContinue| ? {$_.NetEnabled -eq $true} | Select MACAddress, NetConnectionID, PNPDeviceID
                $configs = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $computer  -ErrorAction SilentlyContinue| ? { $_.IPEnabled -eq $true} | Select MACAddress, IPAddress, DNSDomainSuffixSearchOrder, DNSServerSearchOrder, DefaultIPGateway, IPSubnet
   
	   
                $info.Computer = $computersystem.Name
                $info.Mfg = $computersystem.Manufacturer
                $info.Model = $computersystem.Model
                $info.TotalPhysicalMemory = $computersystem.TotalPhysicalMemory
                $info.TotalMemoryGB = $computersystem.'TotalMemory(GB)'
                Remove-Variable computersystem
                $info.BIOS = $bios.Name
                $info.BIOSVersion = $bios.Version
                $info.BIOSReleaseDate = "$($Bios.ReleaseDate.Substring(0,4))-$($Bios.ReleaseDate.Substring(4,2))-$($Bios.ReleaseDate.Substring(6,2))"
                $info.SerialNumber = $bios.SerialNumber
                $info.OS = $os.Caption
                $info.ServicePack = $os.CSDVersion
                $info.OSArchitecture = $os.OSArchitecture
                $info.OSBuildNumber = $os.Version
		
                # Process HD Data
                $HDS = @()
                Foreach ($HD in $HDInfo) {
                    $h = "" | Select VolPath, TotalSize, FreeSpace, Drive
                    $h.Drive = $HD.DeviceID.replace(":", "")
                    $h.TotalSize = $HD.Size
                    $h.FreeSpace = $HD.Freespace
                    $h.VolPath = $HD.path.tostring().replace("""", "").replace("root\cimv2:Win32_LogicalDisk.DeviceID=", "").replace(":", "")
                    $HDS += $h

                }
                $info.HD = $HDS
                # Process Network Adapter Information
                $NetworkInfo = @()
                foreach ($Config in $configs) {
                    $NetInfo = "" | Select MAC, IP, SubNet, GateWay, Name, DNS, DNSSearchSuffix, BroadcastIP, WOL
                    $NetInfo.MAC = $config.MACAddress
                    $netinfo.Name = $adapters[$configs.MACAddress.IndexOf($config.MACAddress)].NetConnectionID
                    $NetInfo.IP = $config.IPAddress
                    $NetInfo.SubNet = $config.IPSubnet
                    $netinfo.GateWay = $config.DefaultIPGateway
                    $netinfo.DNS = $config.DNSServerSearchOrder
                    $netinfo.DNSSearchSuffix = $config.DNSDomainSuffixSearchOrder
                    $netinfo.BroadcastIP = Get-BroadcastAddress -IPAddress $config.ipAddress[0] -SubnetMask $config.IPSubnet[0]
                    $pnpDeviceID = $adapters[$configs.MACAddress.IndexOf($config.MACAddress)].PNPDeviceID
                    $nicPower = Get-WmiObject MSPower_DeviceWakeEnable -Namespace root\wmi -ComputerName $computer|
                        where {$_.instancename -match [regex]::escape($PNPDeviceID) }        
                    if ($nicPower -ne $null)
                    { $NetInfo.WOL = $true}
                    else
                    { $NetInfo.WOL = $false}

                    $NetworkInfo += $netinfo
                }
                $info.network = $NetworkInfo
                $info.Software = Get-SiemensSoftware -Computers $computer
                Remove-Variable H -ErrorAction SilentlyContinue
                Remove-Variable OS
                Remove-Variable bios
                $computerinfo += $info
            }
            catch {
                $info = "" | Select Computer, Mfg, Model, TotalPhysicalMemory, TotalMemoryGB, BIOS, BIOSVersion, BIOSReleaseDate, SerialNumber, OS, ServicePack, OSArchitecture, OSBuildNumber
                $info.computer = $Computer
                $info.Model = "Unable to connect to WMI"
          
         
                $computerinfo += $info
            }
        }
    }
    End {
        return $ComputerInfo
    }
}
#endregion Get-LabPCInfo


#region Update-SPLabPCInfo
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
function Update-SPLabPCInfo {
    [CmdletBinding()]
   
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [object[]]$Computers
    )

    Begin {
        $ErrorActionPreference = "stop"
        #Edit to match full path and filename of where you want log file created
        #Load SharePoint DLL's
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $weburl = "http://rndsharepoint.dow.com/sites/la/LASolutions/PCS7"
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl) 

        # $Context.Credentials = $creds
        $computers = $computers | ? { $_.model -notlike "Unable*"}
        $ModelID = UpdateSPModel -Computers $Computers
    }
    Process {
        foreach ($computer in $Computers) {
            Write-Verbose "Updating $computer"
            $web = $Context.Web
            $weblist = "Computer Inventory"
            $Context.Load($web) 
            $Context.ExecuteQuery() 

            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            $query.ViewXml = "<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($Computer.computer)</Value></Eq></Where></Query></View>"
            #$query.ViewXml = "<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($Computer.computer)</Value></Eq></Where></Query></View>"
            #[Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            # "<View><ViewFields><FieldRef Name='Product' /><FieldRef Name='Title'/></ViewFields></View>"
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
            if ($items.count -eq 0) {
                $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
                $itemCreateInfo
                $new = $list.AddItem($itemCreateInfo)
                $new["Title"] = $computer.Computer
                $new.Update()
                $context.ExecuteQuery()
                $context.Load($items)
                $context.ExecuteQuery()

            }
            if ($items.count -gt 0) {
                $items[0]["Title"] = $computer.Computer
                $items[0]["Memory"] = $Computer.TotalPhysicalMemory
                $items[0]["Serial_x0020_Number"] = $Computer.SerialNumber
                $items[0]["Model_x0020_Number"] = $ModelID[$ModelID.Computer.Indexof($Computer.Computer)].ID
                # Get HD Info	 
                $HDInfo = UpdateSPHardDrive -ComputerInfo $computer -ComputerID $items[0]["ID"]
                $LookupCollection = @()
                foreach ($HD in $HDinfo) {
                    if ($hd.id -ne $null) {
                        $lookupValue = New-Object Microsoft.SharePoint.Client.FieldLookupValue
                        $lookupvalue.LookupId = $HD.ID
                        $lookupCollection += $lookupValue
                    }
                }
                If ($LookupCollection.count -gt 0) {  
                    $items[0]["Hard_x0020_Drives"] = [Microsoft.SharePoint.Client.FieldLookupValue[]]$lookupCollection 
                }


                $NetInfo = UpdateSPNetwork -ComputerInfo $computer -ComputerID $items[0]["ID"]
                $LookupCollection = @()
                foreach ($Net in $Netinfo) {
                    if ($Net.id -ne $null) {
                        $lookupValue = New-Object Microsoft.SharePoint.Client.FieldLookupValue
                        $lookupvalue.LookupId = $Net.ID
		   
                        $lookupCollection += $lookupValue
                    }
                }
                If ($LookupCollection.count -gt 0) {  
                    $items[0]["Network"] = [Microsoft.SharePoint.Client.FieldLookupValue[]]$lookupCollection 
                }



                $items[0].update()

            }
		

        }
		
        $Context.ExecuteQuery()

    }
	
    End {
    }
}
#endregion Update-SPLabPCInfo


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
function UpdateSPModel {
    [CmdletBinding()]
	
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [object[]]$Computers
    )

    Begin {
        $ErrorActionPreference = "Continue"
        #Edit to match full path and filename of where you want log file created
        #Load SharePoint DLL's
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $weburl = "http://rndsharepoint.dow.com/sites/la/LASolutions/PCS7"
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl) 

        # $Context.Credentials = $creds
        $computers = $computers | ? { $_.model -notlike "Unable*"}
        $ModelID = @()
    }
    Process {
        foreach ($computer in $Computers) {
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
        

            if ($items.count -eq 0) {
                # No Data info found.  Add new computer info
                $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
                $itemCreateInfo
                $new = $list.AddItem($itemCreateInfo)
                $new["Title"] = $computer.Model
                $new["HW_x0020_Vendor"] = $Computer.Mfg
                $new.Update()
                $Context.ExecuteQuery()
                # Reload Query to get new Item
                $context.Load($items)
                $context.ExecuteQuery()

            }
            Else {
                $ID = "" | Select Computer, ID
                $ID.Computer = $computer.Computer
                If ($items[0]["ID"] -ne $null) {
                    $ID.ID = $items[0]["ID"]
                }
                else {
                    #No Data Found
                    $ID.ID = "Not Found"
                }
                $modelID += $ID

		
            }
		

        }
	   
		
    }
    End {
        Return $ModelID
    }
}



###############################################################################################

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
function UpdateSPHardDrive {
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

    Begin {
        write-debug "Updating HD's for $($computerinfo.computer)"
        $ErrorActionPreference = "Continue"
        #Edit to match full path and filename of where you want log file created
        #Load SharePoint DLL's

        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $weburl = "http://rndsharepoint.dow.com/sites/la/LASolutions/PCS7"
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl) 

        # $Context.Credentials = $creds
        $computers = $computers | ? { $_.model -notlike "Unable*"}
        $HardDriveIDs = @()
        # Remove-Variable HDinfo -scope global #-ErrorAction SilentlyContinue
        #New-Variable HDInfo -Scope Global
    }
    Process {
        foreach ($HD in $ComputerInfo.HD) {
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
            If ($items.count -eq 0) {
                $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
                $itemCreateInfo
                $new = $list.AddItem($itemCreateInfo)
                $new["Title"] = $HD.VolPath
			
                $new.Update()
                $Context.ExecuteQuery()
                $context.Load($items)
                $context.ExecuteQuery()
            }
            if ($items.count -gt 0) {
		   
                $items[0]["Total_x0020_Size"] = $hd.TotalSize
                $items[0]["Free_x0020_Space"] = $hd.Freespace
                $items[0]["Drive"] = $hd.drive
                $items[0]["Computer"] = $ComputerID
                $ID = "" | Select Computer, ID
                $ID.Computer = $ComputerInfo.Computer
                $ID.ID = $items[0]["ID"]
                $HardDriveIDs += $ID
                $items[0].update()
                $context.ExecuteQuery()
            }
            Else {
                # No Data info found.  Add new computer info 
		   
                $ID = "" | Select Computer, ID
                $ID.Computer = $ComputerInfo.Computer.ToString()
                $ID.ID = "Not Found"
                $HardDriveIDs += $ID
            }

	   
        }
		

    }
	   
   
    End {
        write-debug "End Updating HD"
      
        return $HardDriveIDs

	
    }
}




###############################################################################################

function UpdateSPNetWork {
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

    Begin {
        write-debug "Updating Networks's for $($computerinfo.computer)"
        $ErrorActionPreference = "Continue"
        #Edit to match full path and filename of where you want log file created
        #Load SharePoint DLL's

        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $weburl = "http://rndsharepoint.dow.com/sites/la/LASolutions/PCS7"
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl) 

        # $Context.Credentials = $creds
        $computers = $computers | ? { $_.model -notlike "Unable*"}
        $NetWorkIDs = @()
        Remove-Variable HDinfo -scope global #-ErrorAction SilentlyContinue
        New-Variable HDInfo -Scope Global
    }
    Process {
        foreach ($Network in $ComputerInfo.Network) {
            $web = $Context.Web
            $weblist = "LKUPNetworkCard"
            $Context.Load($web) 
            $Context.ExecuteQuery() 

            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            $Model = $computer.Model #.Replace(" ","_x0020_")
            $query.ViewXml = "<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($NetWork.MAC)</Value></Eq></Where></Query></View>"
            #[Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
            If ($items.count -eq 0) {
                #Record not found.  Create initial Record
                $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
                $itemCreateInfo
                $new = $list.AddItem($itemCreateInfo)
                $new["Title"] = $network.MAC
			
                $new.Update()
                $Context.ExecuteQuery()
                #Reload Items to get new Record ID
                $context.Load($items)
                $context.ExecuteQuery()
            }
            if ($items.count -gt 0) {
                #Update Record
                $items[0]["IP_x0020_Address"] = $Network.IP | % { $_}
                $items[0]["SUBNET"] = $network.subnet | % { $_}
                $items[0]["Gateway"] = $network.Gateway | % { $_}
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
            }
            Else {
                # Record not found and unable to add new Record for some reason.  
                $ID = "" | Select Computer, ID
                $ID.Computer = $ComputerInfo.Computer.ToString()
                $ID.ID = "Not Found"
                $NetworkIDs += $ID
            }

	   
        }
		

    }
	   
   
    End {
	       
        return $NetWorkIDs
    }
}

###############################################################################################
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
function Get-SiemensSoftware {
    [CmdletBinding()]
    [OutputType([object[]])]
    Param
    (
        # List of Computer(s) to run function against
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string[]]$Computers

    )

    Begin {
    }
    Process {
        $HKLM = 2147483650
        $Regkeys = @("SOFTWARE\Wow6432Node\Siemens\AUTSW", "SOFTWARE\Siemens\AUTSW")

        $objReg = [WMIClass]"\\$computer\root\default:StdRegProv"
        $values = @("ProdGroup", "ProdName", "VersionString", "Release", "TechnVersion")

        #$keys = $objReg.EnumKey($hklm, $regkey)
        $software = @()
        foreach ($regKey in $regkeys) {
            foreach ($key in $keys) {
                $subkeys = $objReg.EnumKey($HKLM, $Regkey)
                Foreach ($Sub in $subkeys.sNames) {

                    $v = "" | Select ProdGroup, ProdName, VersionString, Release, TechnVersion
                    $values | % { $v.$_ = ($objReg.GetStringValue($HKLM, "$regkey\$sub", $_)).svalue }
                    if ($v.ProdName -ne $null)
                    { $software += $v }

                }

            }
        }
    }
    End {
        return $software
    }
}



#http://community.idera.com/powershell/powertips/b/tips/posts/calculate-broadcast-address
function Get-BroadcastAddress {
    param
    (
        [Parameter(Mandatory = $true)]
        $IPAddress,
        $SubnetMask = '255.255.255.0'
    )

    filter Convert-IP2Decimal {
        ([IPAddress][String]([IPAddress]$_)).Address
    }


    filter Convert-Decimal2IP {
        ([System.Net.IPAddress]$_).IPAddressToString 
    }

    [UInt32]$ip = $IPAddress | Convert-IP2Decimal
    [UInt32]$subnet = $SubnetMask | Convert-IP2Decimal
    [UInt32]$broadcast = $ip -band $subnet 
    $broadcast -bor -bnot $subnet | Convert-Decimal2IP
}