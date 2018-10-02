#Requires -Version 3.0
function Add-AcronisFolder
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
     Add-AcronisFolder -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Add-AcronisFolder

.EXAMPLE
     Add-AcronisFolder -Param1 'Value1', 'Value2' -Param2 'Value'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author: Tony Turner
#>

    [CmdletBinding()]
    
    param (
        [Parameter(Mandatory, 
            ValueFromPipeline)]
        [string]$Drive
    )

    BEGIN
    {
        #Used for prep. This code runs one time prior to processing items specified via pipeline input.
    }

    PROCESS
    {
        If (!(Test-path "$($drive):\Acronis Backup\Nightly"))
        {
            new-item -ItemType Directory -Path "$($drive):\Acronis Backup\Nightly" | out-null
        }
        If (!(Test-path "$($drive):\Acronis Backup\Weekly"))
        {
            new-item -ItemType Directory -Path "$($drive):\Acronis Backup\Weekly" | out-null
        }
        If (!(Test-path "$($drive):\Acronis Backup\Monthly"))
        {
            new-item -ItemType Directory -Path "$($drive):\Acronis Backup\Monthly" | out-null
        }
        #This code runs one time for each item specified via pipeline input.

        
    }

    END
    {
        #Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
    }

}
<#
$drive = "x"
If (!(Test-path "$($drive):\Acronis Backup\Nightly"))
        {
            new-item -ItemType Directory -Path "$($drive):\Acronis Backup\Nightly" | out-null
        }
        If (!(Test-path "$($drive):\Acronis Backup\Weekly"))
        {
            new-item -ItemType Directory -Path "$($drive):\Acronis Backup\Weekly" | out-null
        }
        If (!(Test-path "$($drive):\Acronis Backup\Monthly"))
        {
            new-item -ItemType Directory -Path "$($drive):\Acronis Backup\Monthly" | out-null
        }
#>