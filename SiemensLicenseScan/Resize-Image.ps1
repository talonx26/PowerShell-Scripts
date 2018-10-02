#Requires -Version 3.0
function Resize-Image
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
     Resize-Image -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Resize-Image

.EXAMPLE
     Resize-Image -Param1 'Value1', 'Value2' -Param2 'Value'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author: Tony Turner
#>
    [CmdletBinding()]
    
    param (
        # Param1 help description
        [Parameter(Mandatory = $true, 
            Position = 0)]
        [object]$Image,
        [Parameter(Mandatory = $false, 
            Position = 1)]
        [double]$scale = 0
        
    )
    
    begin
    {   
        #$VerbosePreference = "Continue"
    }
    
    process
    {
        
        write-verbose "Entering Resize-Image"
        $ImageProcess = new-object -ComObject WIA.ImageProcess
        Write-Verbose "Before"
        WRite-verbose "Image Width  : $($Image.Width)"
        write-verbose "Image Height : $($Image.Height)"
        write-verbose "Scale : $scale"
        if ($scale -ne 0)
        {
            Write-Verbose "Resizing"
            $ImageProcess.Filters.Add($ImageProcess.FilterInfos.Item("Scale").FilterID)
            $ImageProcess.filters
            $ImageProcess.Filters.Item(1).Properties.Item("MaximumWidth").Value = [string]($image.Width * $scale)
                
            $ImageProcess.Filters.Item(1).Properties.Item("MaximumHeight").Value = [string]($image.height * $scale)
            Write-Verbose "Width : $($ImageProcess.Filters.Item(1).Properties.Item("MaximumWidth").Value)"
            Write-Verbose "Height : $($ImageProcess.Filters.Item(1).Properties.Item("MaximumHeight").Value)"
        }
        else
        {
            Write-Verbose "Converting to JPG"
            $imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
            $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = $wiaFormatJPEG
            $imageProcess.Filters.Item(1).Properties.Item("Quality").Value = 50
        }
            
        $SavedImages = $imageProcess.Apply($image)
        Write-Verbose "After"
        WRite-verbose "Image Width  : $($SavedImages.Width)"
        write-verbose "Image Height : $($SavedImages.Height)"
        
    }
  
    end
    {
        write-verbose "Exiting Resize Image"
        return $SavedImages
        #$VerbosePreference = "SilentlyContinue"
    }
}
