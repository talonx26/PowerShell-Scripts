function get-LicenseScan
{
    [CmdletBinding()]
    Param
    (
        
    )

    Begin
    {
        $deviceManager = New-Object -ComObject WIA.DeviceManager
        $device = $deviceManager.DeviceInfos.item(1).connect()
        # $wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
        $wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
        $barcodes = New-Object System.Collections.ArrayList
        $imageProcess = new-object -ComObject WIA.ImageProcess

    }
    Process
    {
        foreach ($item in $device.Items)
        { 
            $item.Properties("Vertical Resolution") = 600
            $item.Properties("Horizontal Resolution") = 600
            $item.Properties("Format") = $wiaFormatJPEG
            $item.Properties("Filename Extension") = "JPG"
            $image = $item.Transfer($wiaFormatJPEG) 
        } 
        $imageProcess.Filters.Add($imageProcess.FilterInfos.Item("Convert").FilterID)
        $imageProcess.Filters.Item(1).Properties.Item("FormatID").Value = $wiaFormatJPEG
        #$imageProcess.Filters.Item(1).Properties.Item("Quality").Value = 5
        $image = $imageProcess.Apply($image)



        add-type -path ".\BarcodeImaging.dll"
        $file = "C:\Scan\temp01.$($image.fileExtension)"
        $image.SaveFile($file) 
        $image = $null
        $barcodes = New-Object System.Collections.ArrayList; [barcodeimaging]::ScanPage([ref]$barcodes, [System.Drawing.Bitmap]::FromFile($file), 600, 2, 1); 
        foreach ($code in $barcodes)
        {
            if ($code.StartsWith("S"))
            {
                $fileName = $code
            }
            else
            {
                $directory = $code
                if (!(test-path c:\scan\$directory))
                {
                    new-item -ItemType Directory -Path c:\scan\$directory 
                }
            }
            
        }
        move-item $file "c:\scan\$directory\$filename.jpg"
        
        <# 
        $refs = @(
            "c:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.Drawing.dll"

        ) 
        #>

    }
    End
    {
    }
}




Add-Type -Path $refs

<# 
$index = 1
do {
    $index++
} until (!(test-path "C:\Scan\test$index.$($image.fileExtension)"))
$index

    $image.SaveFile("C:\Scan\test$index.$($image.fileExtension)") 

 #>