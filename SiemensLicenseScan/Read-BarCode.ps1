#Requires -Version 3.0
function Read-BarCode
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
     Read-BarCode -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Read-BarCode

.EXAMPLE
     Read-BarCode -Param1 'Value1', 'Value2' -Param2 'Value'
 
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
        [Object]$Image
    )
    
    begin
    {
        $refs = @(
            ".\dll\ClearImageNet.70.dll"   

        ) 
        add-type -path $refs
        $barcodes = New-Object System.Collections.ArrayList; 
        if ($image.count -eq 2)
        {$image = $image[1]}
    }
            
    process
    {
        write-verbose "Processing Barcodes"
        $Imageconverter = New-Object System.Drawing.ImageConverter
        write-verbose "Converting Image"
        [System.Drawing.Bitmap]$Bitmap = $Imageconverter.ConvertFrom($image.FileData.BinaryData)

        #$image = $null
        $barcodes = New-Object System.Collections.ArrayList;
        Write-verbose "Scanning for Barcodes"
        $br = [Inlite.ClearImageNet.BarcodeReader]::new()
        # $barcodes = [Inlite.ClearImageNet.Barcode]::new()
        $br.Code39 = $true
        $br.qr = $true
        $barcodes = $br.Read($Bitmap)
        #[barcodeimaging]::ScanPage([ref]$barcodes, $bitmap, 1000, 2, 1)
        Foreach ($code in $barcodes)
        {
            $VerbosePreference = "Continue"
            Write-Verbose "$($code.Text)  "
            Write-Verbose "$($code.Length)"
        }
        If ($barcodes.count -lt 2)
        {
            Write-Verbose "Not enough Barcodes detected. Increasing picture size and rescanning" -Verbose
            $barcodes = New-Object System.Collections.ArrayList;
            If ($image.width -ge $image.height)
            {$scale = 20000 / $image.width}
            else
            { $scale = 20000 / $image.height }
                
            $im = Resize-Image -Image $image -scale $scale
            if ($im.count -eq 2) { $im = $im[1]}
            [System.Drawing.Bitmap]$Bitmap = $Imageconverter.ConvertFrom($im.FileData.BinaryData)
            #  [barcodeimaging]::ScanPage([ref]$barcodes, $bitmap, 1000, 2, 1)
            $im = $null
            Remove-Variable im
        }
       

  
        
    }
    End
    {
        write-verbose "Exit Barcodes."
        return $barcodes
    }
}