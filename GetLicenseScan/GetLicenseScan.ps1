$PSScriptRoot


function get-LicenseScan
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $false, 
            Position = 0)]
        $OutDirectory = "c:\scan"
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
        $MorePages = $true
    }
    Process
    { 
        Write-Verbose "Starting"
        foreach ($item in $device.Items)
        { 
            $item.Properties("Vertical Resolution").Value = 1200
            $item.Properties("Horizontal Resolution").value = 1200

         
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
                    foreach ($code in $barcodes)
                    {
                        if ($code.StartsWith("S"))
                        {
                            $fileName = $code
                            "Filename  $filename"
                        }
                        else
                        {
                            if ($SiemensIDs.ContainsKey($code))
                            {
                                #
                                $directory = $code
                                "Dir : $directory"
                                
                                $blnSiemensSoftwareFound = $true
                            }
                        }            
                    }
                   
                    If ($barcodes.count -ne 2)
                    {
                        $directory = "Problem\$directory"
                        $ScanProblem = "" | Select-Object ScanNumber, Directory, FileName
                        $ScanProblem.ScanNumber = $scan
                        $ScanProblem.FileName = $filename
                        $ScanProblem.Directory = $directory
                        $problemScan += $ScanProblem
                    }
                    #Create Directory if it doesn't exist
                    if (!(Test-Path $OutDirectory\$Directory))
                    {
                        Write-Verbose "Creating $directory"
                       
                        new-item -ItemType Directory -Path $OutDirectory\$directory | out-null
                    }
                    $index = 1 
                    # If file exists increase index until file doesn't exist
                    $tempFileName = $filename
                    While (Test-Path ("$OutDirectory\$directory\$tempfilename.jpg") )
                    {
                      
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
            Write-Verbose "Not enough Barcodes detected. Increasing picture size and rescanning" -Verbose
            $barcodes = New-Object System.Collections.ArrayList;
            If ($image.width -ge $image.height)
            {$scale = 20000 / $image.width}
            else
            { $scale = 20000 / $image.height }
                
            $im = Resize-Image -Image $image -scale $scale
            if ($im.count -eq 2) { $im = $im[1]}
            [System.Drawing.Bitmap]$Bitmap = $Imageconverter.ConvertFrom($im.FileData.BinaryData)
            [barcodeimaging]::ScanPage([ref]$barcodes, $bitmap, 1000, 2, 1)
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


function UpdateSiemensLicense
{
    [CmdletBinding()]
    [OutputType([object[]])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string[]]$Images
      
    )

    Begin
    {
        $ErrorActionPreference = "Continue"
        #Load SharePoint DLL's
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $webURL = "https://workspaces.bsnconnect.com/sites/LabAuto/Inventory"
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)
        $context.Credentials = $credentials
      
    }
    Process
    {
        foreach ($image in $Images)
        {
            #Parse Path to locate Serial Number and Part Number
            $SerialNumber = $image.split("\")[-1].split(".")[0].toupper()
            $PartNumber = $image.split("\")[$image.split("\").count - 2]
            Write-Verbose "Image  :  $Image" -Verbose
            Write-Verbose "Part Number : $Partnumber" -verbose
            if ($partNumber -ne "")
            {
                #Check to see if Software Part Number is already in SharePoint
                $SiemensIDs = Get-SiemensLicenseID -PartNumbers $PartNumber
            }
            $FileName = Get-ChildItem $image
            $root = $filename.Directory.Parent.FullName
            $web = $Context.Web
            $weblist = "Siemens License Tracker"
            $Context.Load($web) 
            $Context.ExecuteQuery() 

            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            
            $query.ViewXml = "<View Scope='RecursiveAll'><Query>,<Where><Eq><FieldRef Name='Serial_x0020_Number'/><Value Type='Text'>$($SerialNumber)</Value></Eq></Where></Query></View>"
            
            #$query.viewXml = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
            If ($items.count -eq 0)
            {
                #Record not found.  Create initial Record
                If ($SiemensIDs.ContainsKey($PartNumber) -and ($SerialNumber.StartsWith("S")))
                {
                    write-host "Adding Record"
                    $itemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation 
                    # $itemCreateInfo
                    $new = $list.AddItem($itemCreateInfo)
                    #$lookupValue = New-Object Microsoft.SharePoint.Client.FieldLookupValue
                    #$lookupvalue.LookupId = $SiemensIDs[$PartNumber]
                    #$new["Software_x003a_Part_x0020_Number"] = $lookupValue
                    $new["Serial_x0020_Number"] = $SerialNumber			
                    $new.Update()
                    $Context.ExecuteQuery()
                    #Reload Items to get new Record ID
                    $context.Load($items)
                    $context.ExecuteQuery()
                } 
                else
                {
                    #Cannot find software 
                    $Root = "$root\Problem\$($FileName.directory.name)\$($Filename.Name)"              
                    
                }
               
            }
            if ($items.count -eq 1)
            {
                if ($Items[0]["Software"] -eq $null )
                {
                    $items[0]["Software"] = $SiemensIDs[$PartNumber]
                }
                #Update Record  - Add attachment if missing
                if (!($items[0].FieldValues.Attachments))
                {
                    $att = [Microsoft.SharePoint.Client.AttachmentCreationInformation]::new()
                    $file = [System.IO.FileStream]::new($image, [System.IO.FileMode]::Open)
                    $att.ContentStream = $file
                    $att.FileName = $file.name.split("\")[-1]
                    

                    [Microsoft.SharePoint.Client.Attachment]$attach = $items[0].AttachmentFiles.add($att)

                    $context.load($attach)
                    $context.ExecuteQuery()
                    $Root = "$root\Processed\$($FileName.directory.name)\$($Filename.Name)"
                }
                else
                {
                    #Attachment Exists move file to different folder for review
                    $Root = "$root\AttachmentExists\$($FileName.directory.name)\$($Filename.Name)"              
                    
                }
            }
            Else
            {
                #Cannot find software 
                $Root = "$root\Problem\$($FileName.directory.name)\$($Filename.Name)"              
                
            }
            $tempDir = $root.Replace($FileName.name, "")
            if (!(Test-Path $tempdir))
            {
                Write-Verbose "Creating $directory"
                       
                new-item -ItemType Directory -Path $tempDir | out-null
            }
            # Move file to correct folder
            Move-Item $FileName.FullName $root -Force
        }
        
        

    } 
        

	   
   
    End
    {
	       
        # return $NetWorkIDs
    }
}

#>
get-licensescan # -verbose
#$ListWebService = "http://rndsharepoint.dow.com/sites/la/LASolutions/PCS7/_vti_bin/lists.asmx"
#
$imagename = "c:\scan\1P6ES7653-2BA00-0XB5\SVPJ51557719.jpg"
#$imagename = "C:\scan\1P6ES7653-2BA00-0XB5\SVPH11544633.jpg"
#UpdateSiemensLicense -Images $imagename