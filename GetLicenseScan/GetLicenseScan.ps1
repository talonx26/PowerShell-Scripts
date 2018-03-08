$PSScriptRoot


function get-LicenseScan
{
    [CmdletBinding()]
    Param
    (
        
    )

    Begin
    { 
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
        $refs = @(
            ".\BarcodeImaging.dll"

        ) 


        add-type -path $refs
        $deviceManager = New-Object -ComObject WIA.DeviceManager
        $device = $deviceManager.DeviceInfos.item(1).connect()
        # $wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
        $wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
        $barcodes = New-Object System.Collections.ArrayList
        $imageProcess = new-object -ComObject WIA.ImageProcess
        $resizeImageProcess = new-object -ComObject WIA.ImageProcess
    }
    Process
    { 
        foreach ($item in $device.Items)
        { 
            $item.Properties("Vertical Resolution").Value = 600
            $item.Properties("Horizontal Resolution").value = 600
            
            #  $item.Properties("Format") = $wiaFormatJPEG
            #  $item.Properties("Filename Extension") = "JPG"
            $image = $item.Transfer($wiaFormatJPEG) 
        } 
        $resizeImageProcess.Filters.Add($resizeImageProcess.FilterInfos.Item("Scale").FilterID)
        $imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
        $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = $wiaFormatJPEG
        $imageProcess.Filters.Item(1).Properties.Item("Quality").Value = 50
        $Saveimage = $imageProcess.Apply($image)
        $resizeImageProcess.Filters.Item(1).Properties.Item("MaximumWidth").Value = [string]($image.Width * 2)
        $resizeImageProcess.Filters.Item(1).Properties.Item("MaximumHeight").Value = [string]($image.height * 2)
        $image = $resizeImageProcess.Apply($image)
        
        #$file = "C:\Scan\temp01.$($image.fileExtension)"
        $Imageconverter = New-Object System.Drawing.ImageConverter
        [System.Drawing.Bitmap]$Bitmap = $Imageconverter.ConvertFrom($image.FileData.BinaryData)

        #$image = $null
        $barcodes = New-Object System.Collections.ArrayList; 
        $Bitmap.Width
        $Bitmap.Height
        [barcodeimaging]::ScanPage([ref]$barcodes, $bitmap, 1000, 2, 1)
        $filename = ""
        $directory = ""
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
                if (!(test-path c:\scan\$directory))
                {
                    new-item -ItemType Directory -Path c:\scan\$directory 

                }
            }            
        }

        $saveimage.SaveFile("c:\scan\$directory\$filename.jpg") 

        
    }
    End
    {
    }
}