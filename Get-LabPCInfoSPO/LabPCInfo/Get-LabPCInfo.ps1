#Requires -Version 3.0

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
function Get-LabPCInfo
{
    [CmdletBinding()]
	
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string[]]$Computers = $env:COMPUTERNAME
    )

    Begin
    {
        $ComputerInfo = @()
    }
    Process
    {
        Foreach ($computer in $computers)
        {
            try
            {
        
                Write-Host
                write-host $Computer
                write-host
                If (-not (Test-Connection -ComputerName $computer -count 1 -ErrorAction SilentlyContinue))
                {  throw "Unable to ping $computer"}
                # Get WMI Data
                $info = "" | Select Computer, Mfg, Model, CSD, Owner, OwnerID, City, Building, Description, TotalPhysicalMemory, TotalMemoryGB, BIOS, BIOSVersion, BIOSReleaseDate, SerialNumber, OS, ServicePack, OSArchitecture, OSBuildNumber, HD, Network, Software
                $OS = Get-WmiObject win32_operatingsystem -ComputerName $Computer  -ErrorAction SilentlyContinue| select Caption, CSDVersion, OSArchitecture, Version
                
                if ($os -eq $null)
                {  throw "Unable WMI not accessible on $computer"}
                $keys = @("CSD", "Owner", "OwnerID", "City", "Building", "Description")
                $regvalues = Get-RegValue -Computers $computer -Keys $keys -Path "Dow"
                $info.csd = $regvalues.value[$regvalues.key.IndexOf("CSD")]
                $info.Owner = $regvalues.value[$regvalues.key.IndexOf("Owner")]
                $info.OwnerID = $regvalues.value[$regvalues.key.IndexOf("OwnerID")]
                $info.City = $regvalues.value[$regvalues.key.indexof("City")]
                $info.Building = $regvalues.value[$regvalues.key.IndexOf("Building")]
                $info.Description = $regvalues.value[$regvalues.key.IndexOf("Description")]
                $computersystem = Get-WmiObject win32_computersystem -ComputerName $Computer -ErrorAction SilentlyContinue| select Name, Manufacturer, Model, TotalPhysicalMemory, @{n = "TotalMemory(GB)"; e = {[math]::Round($_.TotalPhysicalMemory / 1GB, 3)}}
                $HDInfo = Get-WmiObject -query  "select * from win32_logicaldisk where DriveType = '3'" -ComputerName $computer -ErrorAction SilentlyContinue
                $bios = Get-WmiObject  win32_bios -ComputerName $computer -ErrorAction SilentlyContinue | select Name, Version, ReleaseDate, SerialNumber
                $adapters = Get-WmiObject Win32_NetworkAdapter -ComputerName $computer  -ErrorAction SilentlyContinue| ? {$_.NetEnabled -eq $true} | Select MACAddress, NetConnectionID, PNPDeviceID
                $configs = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $computer  -ErrorAction SilentlyContinue| ? { $_.IPEnabled -eq $true } | Select MACAddress, IPAddress, DNSDomainSuffixSearchOrder, DNSServerSearchOrder, DefaultIPGateway, IPSubnet
   
	   
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
                Foreach ($HD in $HDInfo)
                {
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
                foreach ($Config in $configs)
                {
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
            catch
            {
                $info = "" | Select Computer, Mfg, Model, TotalPhysicalMemory, TotalMemoryGB, BIOS, BIOSVersion, BIOSReleaseDate, SerialNumber, OS, ServicePack, OSArchitecture, OSBuildNumber
                # $_ | select *
                $info.computer = $Computer
                $info.Model = $_.Exception.Message
          
         
                $computerinfo += $info
            }
        }
    }
    End
    {
        return $ComputerInfo
    }
}
#endregion Get-LabPCInfo


#$compinfo = Get-LabPCInfo -Computers (gc .\computers.txt)
#Update-SPLabPCInfo -Computers $compinfo
#https://workspaces.bsnconnect.com/teams/LabAutomation/Lists/Computer%20Inventory/AllItems.aspx