#Requires -Version 3.0
function Get-SiemensLicenseScan {

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
     Get-SiemensLicenseScan -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Get-SiemensLicenseScan

.EXAMPLE
     Get-SiemensLicenseScan -Param1 'Value1', 'Value2' -Param2 'Value'

.INPUTS
    String

.OUTPUTS
    PSCustomObject

.NOTES
    Author: Tony Turner
#>

    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $false,
            Position = 0)]
        $OutDirectory = "c:\scan"
    )

    Begin {

        [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
        add-type -Path .\Dll\Interop.WIA.dll
        $deviceManager = New-Object -ComObject WIA.DeviceManager
        $device = $deviceManager.DeviceInfos.item(1).connect()
        # $wiaFormatPNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
        $wiaFormatJPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
        $barcodes = New-Object System.Collections.ArrayList
        # $wiaFormatTiff = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"
        $scan = 0
        $problemScan = @()
        $MorePages = $true
    }

    Process {
        Write-Verbose "Starting"
        foreach ($item in $device.Items) {
            $item.Properties("Vertical Resolution").Value = 600
            $item.Properties("Horizontal Resolution").value = 600


            do {
                Try { $image = $item.Transfer($wiaFormatJPEG) }
                Catch {
                    if ($_.exception -ilike ("*no documents left*") -and $scan -eq 0) {
                        #There have been no scans so ADF was empty to begin with
                        Write-host "No documents in Scanner. Please check and restart script"
                    }
                    elseif (!($_.exception -ilike ("*no documents left*")) -and $scan -eq 0) {
                        #This should catch error if Scanner doesn't start.
                        Write-host "If Scanner did not run try restarting scanner Scan : $Scan"
                    }

                    $MorePages = $false


                }

                If ($MorePages) {

                    $scan++
                    Write-Verbose "Scan Number $scan" -Verbose
                    #Image file for saving.  Smaller File for file size
                    $Saveimage = Resize-Image -Image $image
                    # Enlarge Image to send to barcode reader to increase accuracy. File stored in memory.
                    #$BarCodeImage = New-Object
                    #  $saveimage.SaveFile("$OutDirectory\temp.jpg")
                    # $BarCodeImage = Resize-Image  -Image $image -scale 2

                    # $Imageconverter = New-Object System.Drawing.ImageConverter
                    # [System.Drawing.Bitmap]$Bitmap = $Imageconverter.ConvertFrom($image.FileData.BinaryData)

                    #$image = $null
                    $barcodes = New-Object System.Collections.ArrayList;

                    #[barcodeimaging]::ScanPage([ref]$barcodes, $bitmap, 1000, 2, 1)
                    Write-Verbose "Barcode Count Before: $($barcodes.count)"
                    $barcodes = Read-Barcode -Image $image
                    Write-Verbose "Barcodes :"

                    Write-Verbose "Barcode Count after: $($barcodes.count)"
                    $filename = ""
                    $directory = ""
                    $barcodes = $barcodes | sort-object -Descending
                    #Check Barcodes to see if p/n is already in SharePoint Site
                    $SiemensIDs = Get-SiemensLicenseID -PartNumbers $barcodes
                    $blnSiemensSoftwareFound = $false
                    $FileTag = $Null
                    foreach ($code in $barcodes) {

                        If ($code.type -eq 'Code39') {
                            if ($code.text.StartsWith("S")) {
                                $fileName = $code.text
                                $filename = $filename.replace(" ", "")
                                "Filename  $filename"
                            }
                            else {
                                if ($SiemensIDs.ContainsKey($code.text)) {
                                    #
                                    $directory = $code.text
                                    "Dir : $directory"

                                    $blnSiemensSoftwareFound = $true
                                }
                                elseif (!$blnSiemensSoftwareFound)
                                { $directory = $code.text }
                            }
                        }
                        if ($code.type -eq 'QR') {
                            $fileTag = "#$($code.Text)#"
                        }
                    }

                    If (($barcodes.count -ne 2 -and !$blnSiemensSoftwareFound) -or $filename -eq '' -or !$blnSiemensSoftwareFound) {
                        $directory = "Problem\$directory"
                        $ScanProblem = "" | Select-Object ScanNumber, Directory, FileName, Description
                        $ScanProblem.ScanNumber = $scan
                        $ScanProblem.FileName = "$filename$filetag".replace("/", "-")
                        $ScanProblem.Directory = $directory
                        If ($barcodes.count -gt 2 ) {
                            $ScanProblem.Description = "Too many Barcodes detected.  Please verify "
                        }
                        If ($filename -eq '' ) {
                            $ScanProblem.Description = "Unable to detect Serial Number.  Please Verify"
                        }
                        If (!$blnSiemensSoftwareFound) {
                            $ScanProblem.Description = "Part Number not located in SharePoint.  Please verify and add to SharePoint."
                        }
                        $problemScan += $ScanProblem
                    }
                    # Replace invalid character in filename
                    $directory = $directory.replace("/", "-")
                    #Create Directory if it doesn't exist
                    if (!(Test-Path $OutDirectory\$Directory)) {
                        Write-Verbose "Creating $directory"

                        new-item -ItemType Directory -Path $OutDirectory\$directory | out-null
                    }
                    $index = 1
                    # If file exists increase index until file doesn't exist
                    $tempFileName = "$filename$FileTag"
                    $tempFileName = $tempFileName.replace("/", "-")
                    While (Test-Path ("$OutDirectory\$directory\$tempfilename.jpg") ) {

                        $TempFileName = "$filename-Scan-$index"
                        $index++
                        Write-Verbose "File exists changing file name to $filename"
                    }
                    $Filename = $tempFileName
                    Write-Verbose "Saving File : $OutDirectory\$directory\$filename.jpg"  -Verbose
                    $saveimage.SaveFile("$OutDirectory\$directory\$filename.jpg")
                    #  UpdateSiemensLicense -Images "$OutDirectory\$directory\$filename.jpg"
                }


            }while ($MorePages)
        }
    }
    End {
        If ($problemScan.count -ne 0) {
            Write-host "Detected problems. Please Review"
            $problemScan
        }
    }


}