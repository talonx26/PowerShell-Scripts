function Get-SiemensSoftware
{
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

    Begin
    {
    }
    Process
    {
        foreach ($computer in $computers)
        {     
            <#  try
            { #>
            Write-Verbose "Enter Computers"
            $HKLM = 2147483650
            $Regkeys = @("SOFTWARE\Wow6432Node\Siemens\AUTSW", "SOFTWARE\Siemens\AUTSW")

            $objReg = [WMIClass]"\\$computer\root\default:StdRegProv"
            $values = @("ProdGroup", "ProdName", "VersionString", "Release", "TechnVersion")
            $keys = $objReg.EnumKey($hklm, $regkeys)
            $software = @()
            foreach ($regKey in $regkeys)
            {
                Write-Verbose "Enter Registry"
                foreach ($key in $keys)
                {
                    Write-Verbose "Enter Keys"
                    $subkeys = $objReg.EnumKey($HKLM, $Regkey)
                    Foreach ($Sub in $subkeys.sNames)
                    {
                        Write-Verbose "Enter SubKeys"
                        $v = "" | Select-Object ProdGroup, ProdName, VersionString, Release, TechnVersion
                        $values | % { $v.$_ = ($objReg.GetStringValue($HKLM, "$regkey\$sub", $_)).svalue }
                        if ($v.ProdName -ne $null)
                        { 
                            $index = -1
                            If ($software -ne $null)
                            {
                                $index = $software.ProdName.IndexOf($v.ProdName)
                            }
                            If ($index -eq -1 -or $software -eq $null )    
                            {$software += $v }
                            else
                            {
                                if ($software[$index].ProdGroup -eq $null) { $software[$index].ProdGroup = $v.ProdGroup}
                                if ($software[$index].VersionString -eq $null) { $software[$index].VersionString = $v.VersionString}
                                if ($software[$index].Release -eq $null) { $software[$index].Release = $v.Release}
                                if ($software[$index].TechnVersion -eq $null) { $software[$index].TechnVersion = $v.TechnVersion}
                            }
                        }
                        
                    }

                }

            }
            <#   }
          
            catch
            {
              
            } #>
            
        }
    }
    End
    {
        return $software
    }
}
$c = Get-SiemensSoftware -Computers "wpcs7lj014ss3"