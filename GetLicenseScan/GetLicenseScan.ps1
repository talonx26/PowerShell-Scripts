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
        add-type -Path .\Interop.WIA.dll
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
        Write-Verbose "Starting"
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
                    if ($_.exception -ilike ("*no documents left*") -and $scan -eq 0)
                    {
                        #There have been no scans so ADF was empty to begin with
                        Write-host "No documents in Scanner. Please check and restart script"
                    }
                    elseif (!($_.exception -ilike ("*no documents left*")) -and $scan -eq 0)
                    {
                        #This should catch error if Scanner doesn't start. 
                        Write-host "If Scanner did not run try restarting scanner Scan : $Scan"
                    }
                   
                    $MorePages = $false
                  
                   
                }
                If ($MorePages)
                {
                    $scan++
                    #Image file for saving.  Smaller File for file size
                    $Saveimage = Resize-Image -Image $image
                    # Enlarge Image to send to barcode reader to increase accuracy. File stored in memory.
                    #$BarCodeImage = New-Object 
                    $BarCodeImage = Resize-Image  -Image $image -scale 2
        
                    # $Imageconverter = New-Object System.Drawing.ImageConverter
                    # [System.Drawing.Bitmap]$Bitmap = $Imageconverter.ConvertFrom($image.FileData.BinaryData)

                    #$image = $null
                    $barcodes = New-Object System.Collections.ArrayList; 
            
                    #[barcodeimaging]::ScanPage([ref]$barcodes, $bitmap, 1000, 2, 1)
                    Write-Verbose "Barcode Count Before: $($barcodes.count)"
                    $barcodes = Read-Barcode -Image $BarCodeimage
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
                            $index = 1
                            $fileName = "No Serial-Scan-$index.jpg"
                        
                            $directory = "Problem\$directory"
                            
                            While (Test-Path ("$OutDirectory\$directory\$filename.jpg") )
                            {
                                $index++
                                $fileName = "No Serial-Scan-$index"
                                
                                Write-Verbose "File exists changing file name to $filename"
                            }
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




function Resize-Image
{
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
    }
}


function Read-Barcode
{
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
            ".\BarcodeImaging.dll"

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
        [barcodeimaging]::ScanPage([ref]$barcodes, $bitmap, 1000, 2, 1)
        If ($barcodes.count -lt 2)
        {
            Write-Verbose "Not enough Barcodes detected. Increasing picture size and rescanning"
            $barcodes = New-Object System.Collections.ArrayList;
            $im = Resize-Image -Image $image -scale 2
            [System.Drawing.Bitmap]$Bitmap = $Imageconverter.ConvertFrom($im.FileData.BinaryData)
            [barcodeimaging]::ScanPage([ref]$barcodes, $bitmap, 1000, 2, 1)
            $im = $null
            Remove-Variable im
        }
        
        
    }
    
    end
    {
        write-verbose "Exit Barcodes."
        return $barcodes
    }
}

get-licensescan -verbose