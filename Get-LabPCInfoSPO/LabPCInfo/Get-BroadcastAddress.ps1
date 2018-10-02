#Requires -Version 3.0
function Get-BroadcastAddress
{
    [CmdletBinding()]
    [OutputType('PSCustomObject')]
    param (
        [Parameter(Mandatory, 
            ValueFromPipeline)]
        [string]$IPAddress,
        [string]$SubnetMask = '255.255.255.0'
        
    )
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
     Get-BroadcastAddress -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Get-BroadcastAddress

.EXAMPLE
     Get-BroadcastAddress -Param1 'Value1', 'Value2' -Param2 'Value'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author: Tony Turner
#>


   

    BEGIN
    {
        #Used for prep. This code runs one time prior to processing items specified via pipeline input.
    }

    PROCESS
    {

        filter Convert-IP2Decimal
        {
            ([IPAddress][String]([IPAddress]$_)).Address
        }


        filter Convert-Decimal2IP
        {
            ([System.Net.IPAddress]$_).IPAddressToString 
        }

        [UInt32]$ip = $IPAddress | Convert-IP2Decimal
        [UInt32]$subnet = $SubnetMask | Convert-IP2Decimal
        [UInt32]$broadcast = $ip -band $subnet 
        $broadcast -bor -bnot $subnet | Convert-Decimal2IP
       

    }

    END
    {
        #Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
    }

}
