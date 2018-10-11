#Requires -Version 3.0
function New-RegistryValue
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
     New-RegistryKey -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | New-RegistryKey

.EXAMPLE
     New-RegistryKey -Param1 'Value1', 'Value2' -Param2 'Value'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    
#>

    [CmdletBinding()]
    [OutputType('PSCustomObject')]
    param (
        [Parameter(Mandatory, 
            ValueFromPipeline)]
        [string[]]$Computers,
        [string] $Path = "Software",
        [ValidateNotNullOrEmpty()]
        [string]$subkey,
    
        [ValidateSet([Microsoft.Win32.RegistryValueKind]::DWord, 
            [Microsoft.Win32.RegistryValueKind]::Binary,
            [Microsoft.Win32.RegistryValueKind]::MultiString,
            [Microsoft.Win32.RegistryValueKind]::QWord,
            [Microsoft.Win32.RegistryValueKind]::ExpandString,
            [Microsoft.Win32.RegistryValueKind]::String,
            [Microsoft.Win32.RegistryValueKind]::None,
            [Microsoft.Win32.RegistryValueKind]::Unknown)]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Win32.RegistryValueKind]$type,
        [ValidateNotNullOrEmpty()]
        $Key,
        [ValidateNotNullOrEmpty()]
        $Value
        <#,
        [ValidateSet([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryHive]::CurrentUser)]
        $hive
        #>
    )

    BEGIN
    {
        #Used for prep. This code runs one time prior to processing items specified via pipeline input.
        $Registrykeys = @()
    }

    PROCESS
    {
        #This code runs one time for each item specified via pipeline input.

        foreach ($Computer in $Computers)
        {
            Try
            {
                $Registrykey = "" | Select-Object Computer, SubKey, Key, Value, Created
                $Registrykey.Computer = $Computer
                $Registrykey.SubKey = "$Path\$subkey"
                $Registrykey.Key = $Key
                $Registrykey.value = $value
                if (!(test-connection $computer -count 1 -erroraction stop))
                {
                    throw [System.Management.Automation.MethodInvocationException]
                }
                $Service = Get-Service -ComputerName $computer -Name RemoteRegistry
                #Stores Startup Type to restore back to original status when finished.
                if ($service.StartType -eq "Disabled")
                {
                    $service | Set-Service -StartupType Manual
                    # Start-Sleep -Seconds 1
                    $blnRemoteRegistryDisabled = $true
                }
                else
                {
                    $blnRemoteRegistryDisabled = $false
                }
                $RemoteRegistryStatus = (Get-Service -ComputerName $computer -Name RemoteRegistry).Status
                  
                If ($RemoteRegistryStatus -eq 'Stopped') { Get-Service -ComputerName $computer -Name RemoteRegistry | Start-Service}
                
                $hive = [Microsoft.Win32.RegistryHive]::LocalMachine
                $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($hive, $computer)
           
                #Check if Key Exists
                
           
          
                if (!$baseKey.OpenSubKey("$Path\$subKey"))
                {
                    #Key doesn't exist Create it
                    #OpenSubkey with $true for write mode
                    $K = $baseKey.OpenSubKey("Software", $true)
               
                    [void]$k.CreateSubKey("$subKey") 
                  
                    
                }
                #OpenSubkey with $true for write mode
                $k = $baseKey.OpenSubKey("$Path\$subkey", $true)
                $k.SetValue($Key, $Value, $type)
                $Registrykey.Created = $true
            }
            catch [System.Management.Automation.MethodInvocationException]
            {
                $Registrykey.created = "Unable to reach computer. Verify Machine is online and not blocked by a firewall."
            }
            Catch [System.UnauthorizedAccessException]
            {
                $Registrykey.Created = "Registry key cannot be written to"
            }
            catch [System.Security.SecurityException]
            {
                $Registrykey.Computer = "User does not have permissions to write to registry"
            }
            catch [System.Net.NetworkInformation.PingException]
            {
                $Registrykey.created = "Unable to reach computer. Verify Machine is online and not blocked by a firewall."
            }


            $Registrykeys += $Registrykey
            #Stop RemoteRegistry service if it was stopped before.
            If ($RemoteRegistryStatus -eq 'Stopped') { Get-Service -ComputerName $computer -Name RemoteRegistry | stop-Service}
            if ($blnRemoteRegistryDisabled) { $Service | Set-Service -StartupType Disabled}
        }
    }

    END
    {
        #Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
        return $Registrykeys
    }

}

#New-RegistryValue -Computers "Wpcs7lj015es" -subkey "Dow" -type String -Key "CSD" -Value "2018-0002"
<#
Foreach ($c in $csv)
{
    Foreach ($prop in $c.psobject.properties)
    {
        If ($prop.Name -ne "Computer")
        {
                
            New-RegistryValue -Computers $c.computer -subkey "DOW" -type String -Key $Prop.Name -Value $prop.Value
        
        }
    }    


}

#>