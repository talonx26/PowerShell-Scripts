#Requires -Version 3.0
  
function Get-RegValue
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
     Get-RegValue -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Get-RegValue

.EXAMPLE
     Get-RegValue -Param1 'Value1', 'Value2' -Param2 'Value'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author:  Tony Turner
   
#>
    param (
        [CmdletBinding()]
        [OutputType('PSCustomObject')]
        [Parameter(Mandatory, 
            ValueFromPipeline)]
        [string[]]$Computers,
        [ValidateNotNullOrEmpty()]
        [string]$Path,
        [ValidateNotNullOrEmpty()]
        [string[]]$Keys
        
    )

    BEGIN
    {
        #Used for prep. This code runs one time prior to processing items specified via pipeline input.
        $registrykeys = @()
    }

    PROCESS
    {
        #This code runs one time for each item specified via pipeline input.

        foreach ($Computer in $Computers)
        {
            
            foreach ($key in $keys)
            {
                
                try
                {
                    $service = (Get-Service -ComputerName $computer -Name RemoteRegistry).Status
                    If ($service -eq 'Stopped') { Get-Service -ComputerName $computer -Name RemoteRegistry | Start-Service}
                    $Registrykey = "" | Select-Object Computer, SubKey, Key, Value
                    $Registrykey.Computer = $Computer
                    $Registrykey.SubKey = "Software\$Path"
                    $Registrykey.Key = $Key
                
                    $hive = [Microsoft.Win32.RegistryHive]::LocalMachine
                    $baseKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($hive, $computer)
                    $k = $baseKey.OpenSubKey("Software\$Path")
                    $Registrykey.Value = $k.getvalue($key)
                }
                catch [System.Management.Automation.MethodInvocationException]
                {
                    $Registrykey.Value = "Error: Unable to reach computer. Verify Machine is online and not blocked by a firewall."
                }
                catch [System.Security.SecurityException]
                {
                    $Registrykey.value = "Error: User doesn't have permission to access registry"
                }
                catch [System.Management.Automation.RuntimeException]
                {
                    #Key doesn't exists
                }
                #Use foreach scripting construct to make parameter input work the same as pipeline input (iterate through the specified items one at a time).
                $registrykeys += $Registrykey
                #Stop RemoteRegistry service if it was stopped before.
                If ($service -eq 'Stopped') { Get-Service -ComputerName $computer -Name RemoteRegistry | stop-Service}
            }
        }
    }

    END
    {
        return $registrykeys
        #Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
    }

}

#get-regvalue -Computers "wpcs7lj015es" -Path "DOW" -Key "CSD"