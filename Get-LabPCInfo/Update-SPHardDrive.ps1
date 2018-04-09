#Requires -Version 3.0
function Update-SPHardDrive {

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
     Update-SPHardDrive -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Update-SPHardDrive

.EXAMPLE
     Update-SPHardDrive -Param1 'Value1', 'Value2' -Param2 'Value'
 
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
    [OutputType('PSCustomObject')]
    param (
        [Parameter(Mandatory, 
                   ValueFromPipeline)]
        [string[]]$Param1,

        [ValidateNotNullOrEmpty()]
        [string]$Param2
    )

    BEGIN {
        #Used for prep. This code runs one time prior to processing items specified via pipeline input.
    }

    PROCESS {
        #This code runs one time for each item specified via pipeline input.

        foreach ($Param in $Param1) {
            #Use foreach scripting construct to make parameter input work the same as pipeline input (iterate through the specified items one at a time).
        }
    }

    END {
        #Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
    }

}
