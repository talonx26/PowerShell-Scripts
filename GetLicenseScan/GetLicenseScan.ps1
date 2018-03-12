$PSScriptRoot


function get-LicenseScan
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $false, 
            Position = 0)]
        [string] $OutDirectory = "C:\scan"
    )

    Begin
    { 
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
       
        $deviceManager = New-Object -ComObject WIA.DeviceManager
        $device = $deviceManager.DeviceInfos.item(1).connect()
        # $wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
        $wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
        $barcodes = New-Object System.Collections.ArrayList
        # $wiaFormatTiff = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
        $scan = 0
        $problemScan = @()
    }
    Process
    { 
        foreach ($item in $device.Items)
        { 
            $MorePages = $true
            $item.Properties("Vertical Resolution").Value = 600
            $item.Properties("Horizontal Resolution").value = 600
         
            do
            {
                Try { $image = $item.Transfer($wiaFormatJPEG)  }
                Catch 
                {
                    $MorePages = $false
                }
                If ($MorePages)
                {
                    $scan++
                    #Image file for saving.  Smaller File for file size
                    $Saveimage = Resize-Image -Images $image
                    # Enlarge Image to send to barcode reader to increase accuracy. File stored in memory.
                    $image = Resize-Image  -Images $image -scale 2
        
                    $Imageconverter = New-Object System.Drawing.ImageConverter
                    [System.Drawing.Bitmap]$Bitmap = $Imageconverter.ConvertFrom($image.FileData.BinaryData)

                    #$image = $null
                    $barcodes = New-Object System.Collections.ArrayList; 
            
                    #[barcodeimaging]::ScanPage([ref]$barcodes, $bitmap, 1000, 2, 1)
                    Write-Verbose "Barcode Count Before: $($barcodes.count)"
                    $barcodes = Read-Barcode -Image $image
                    Write-Verbose "Barcodes :"
                    Write-Verbose $barcodes
                    Write-Verbose "Barcode Count after: $($barcodes.count)"
                    $filename = ""
                    $directory = ""
                    $barcodes = $barcodes | sort-object -Descending   
                    
                    foreach ($code in $barcodes)
                    {
                        if ($code.StartsWith("S"))
                        {
                            $fileName = $code
                            "Filename  $filename"
                        }
                        else
                        {
                            $directory = $code
                            "Dir : $directory"
                            <#      if (!(test-path $OutDirectory\$directory))
                            {
                                new-item -ItemType Directory -Path $OutDirectory\$directory 

                            } #>
                        }            
                    }
                    If ($barcodes.count -gt 2)
                    {
                        Write-Host "More then 2 Barcodes were detected."
                        $x = 1
                        $barcodes | % { Write-host "$x : $_"; $x++}
                        Write-host "1. Detected FileName : $filename"
                        Write-host "2. Detected Directory : $Directory"
                        
                        #  $file = Read-Host -Prompt "Please enter number for the correct filename" 
                        #  $dir = Read-Host -Prompt "Please enter number for the correct directory"
                    }
                    If ($barcodes.count -lt 2)
                    {
                        If ($filename -eq "")
                        {
                            $fileName = "No Serial -$scan.jpg"
                        
                            $directory = "$Problem\$directory"
                            Write-Verbose "No Serial Detected"
                            Write-Verbose "Directory : $directory"
                        }
                    }
                    If ($barcodes.count -ne 2)
                    {
                        $ScanProblem = "" | Select-Object ScanNumber, Directory, FileName
                        $ScanProblem.ScanNumber = $scan
                        $ScanProblem.FileName = $filename
                        $ScanProblem.Directory = $directory
                        $problemScan += $ScanProblem
                    }
                    if (!(Test-Path $OutDirectory\$Directory))
                    {
                        Write-Verbose "Creating $directory"
                       
                        new-item -ItemType Directory -Path $OutDirectory\$directory | out-null
                    }
                    $index = 1 
                    While (Test-Path ("$OutDirectory\$directory\$filename.jpg") )
                    {
                      
                        $fileName = "$filename-Scan-$index"
                        $index++
                        Write-Verbose "File exists changing file name to $filename"
                    }
                    Write-Verbose "Saving File : $OutDirectory\$directory\$filename.jpg"
                    $saveimage.SaveFile("$OutDirectory\$directory\$filename.jpg")  
                }

            } while ($MorePages)
           
        } 
        

        
    }
    End
    {
        If ($problemScan.count -ne 0)
        {
            Write-host "Detected problems. Please Review"
            $problemScan
        }
    }
}

<#
$s = New-Object -ComObject WIA.CommonDialog


$im= $s.ShowTransfer($item,[WIA.FormatID]::wiaFormatJPEG,$false)
get-history | { $_.commandline}



$device.Properties | ft
$item.Properties |ft
$device.Properties | ft
$device.WiaItem("TransferItemFlag")
$item.properties("Item Flags")
$item.properties("Item Flags").Type
$im
$im += $item.Transfer($wiaFormatJPEG)
$item.properties("Item Flags")
$im = @()
$im
$im += $item.Transfer($wiaFormatJPEG)
$item.properties("Item Flags")
$im += $item.Transfer($wiaFormatJPEG)
$item.properties("Item Flags")
$im += $item.Transfer($wiaFormatJPEG)
$item.properties("Item Flags")
$im += $item.Transfer($wiaFormatJPEG)
$im | % { $_.savefile("C:\scan\scantestx$index.jpg");$index++}
add-type -Path .\Interop.WIA.dll
[WIA.WiaItemFlags]
[WIA.WiaItemFlag]
[WIA.WiaItemFlag]::TransferItemFlag
#>


function Resize-Image
{
    [CmdletBinding()]
    [OutputType([object])]
    param (
        # Param1 help description
        [Parameter(Mandatory = $true, 
            Position = 0)]
        [object[]]$Images,
        [Parameter(Mandatory = $false, 
            Position = 1)]
        [double]$scale = 0
        
    )
    
    begin
    {   
        $SavedImages = @()
    }
    
    process
    {
        Foreach ($image in $images)
        {
            $ImageProcess = new-object -ComObject WIA.ImageProcess
            Write-Verbose "Before"
            WRite-verbose "Image Width  : $($Image.Width)"
            write-verbose "Image Height : $($Image.Height)"
            if ($scale -ne 0)
            {
                $ImageProcess.Filters.Add($ImageProcess.FilterInfos.Item("Scale").FilterID)
                $ImageProcess.Filters.Item(1).Properties.Item("MaximumWidth").Value = [string]($image.Width * $scale)
                
                $ImageProcess.Filters.Item(1).Properties.Item("MaximumHeight").Value = [string]($image.height * $scale)
            }
            else
            {
                $imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
                $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = $wiaFormatJPEG
                $imageProcess.Filters.Item(1).Properties.Item("Quality").Value = 50
            }
            
            $Savedimages += $imageProcess.Apply($image)
            Write-Verbose "After"
            WRite-verbose "Image Width  : $($Image.Width)"
            write-verbose "Image Height : $($Image.Height)"
        }
    }
  
    end
    {
        return $SavedImages
    }
}


function Read-Barcode
{
    [CmdletBinding()]
    param (
        # Param1 help description
        [Parameter(Mandatory = $true, 
            Position = 0)]
        [object]$Image
    )
    
    begin
    {
        $refs = @(
            ".\BarcodeImaging.dll"

        ) 
        add-type -path $refs
        $barcodes = New-Object System.Collections.ArrayList; 
    }
            
    process
    {
        $Imageconverter = New-Object System.Drawing.ImageConverter
        [System.Drawing.Bitmap]$Bitmap = $Imageconverter.ConvertFrom($image.FileData.BinaryData)

        #$image = $null
        $barcodes = New-Object System.Collections.ArrayList;
       
        [barcodeimaging]::ScanPage([ref]$barcodes, $bitmap, 1000, 2, 1)
        If ($barcodes.count -lt 2)
        {
            Write-Verbose "Not enough Barcodes detected. Increasing picture size and rescanning"
            $barcodes = New-Object System.Collections.ArrayList;
            $im = Resize-Image -Images $image -scale 2
            [System.Drawing.Bitmap]$Bitmap = $Imageconverter.ConvertFrom($im.FileData.BinaryData)
            [barcodeimaging]::ScanPage([ref]$barcodes, $bitmap, 1000, 2, 1)
            $im = $null
            Remove-Variable im
        }
        
        
    }
    
    end
    {
        return $barcodes
    }
}