Set-Strictmode -Version 3
#Requires -Modules SharePointPNPPowerShellOnline
<#
########  TODO
 

 


#######
#>



#region Helper Functions
<#
.SYNOPSIS
    Update Asset Manager SharePoint site.
 
.DESCRIPTION
    Reads in CSV file and updates the IR Asset Manager SharePoint Site.
    http://rndsharepoint.dow.com/sites/IR/AM

.NOTES
    Author:  Tony Turner
    Version: 3.0

 
#>


<#
   
#>

function Get-SiemensLicenseID {

    <# Synopsis
   #>
    [CmdletBinding()]
    [OutputType('PSCustomObject')]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline)]
        [object[]]$PartNumbers
    )

    BEGIN {
            
        $ErrorActionPreference = "Continue"
        #Load SharePoint DLL's
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $webURL = "https://workspaces.bsnconnect.com/sites/LabAuto/Inventory"
       # $context = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
       # $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)
        Connect-PnPOnline -Url $webURL -Credentials labautosp
       
        #$context.Credentials = $credentials
        $SoftwareIDs = @{ }
    }
    Process {
        foreach ($partnumber in $partnumbers) {
           # $web = $Context.Web
            $weblist = "MLKUPSiemensSoftware"
           # $Context.Load($web)
           # $Context.ExecuteQuery()
            #If PartNumber.Text is Null then object is string is Barcode object
            if ($null -eq $partnumber.Text)
            { $pn = $partnumber }
            else {
                #PartNumber.Text is not Null so object is a Barcode Object
                $pn = $partnumber.text
            }
           # $list = $web.Lists.GetByTitle($weblist)
           # $Query = New-Object Microsoft.SharePoint.Client.CamlQuery

            $query = "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='Part_x0020_Number'/><Value Type='Text'>$pn</Value></Eq></Where></Query></View>"

            #$query.viewXml = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            $items = get-pnplistitem -List $weblist -Query 
            $context.Load($items)
            $context.ExecuteQuery()
            If ($items.count -eq 1) {
                If (!($SoftwareIDs.ContainsKey($pn))) {
                    $SoftwareIDs.Add($pn, $items[0]["ID"])
                }
            }
         }
    }

    END {
        # Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
        # Write-Verbose $SoftwareIDs -Verbose
        return $SoftwareIDs
    }

}

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

function Update-SiemensLicenseCertificate
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
     Update-SiemensLicenseCertificate -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Update-SiemensLicenseCertificate

.EXAMPLE
     Update-SiemensLicenseCertificate -Param1 'Value1', 'Value2' -Param2 'Value'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author: Tony Turner
#>

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
        $refs = @(".\Dll\Microsoft.SharePoint.Client.dll", ".\dll\Microsoft.SharePoint.Client.Runtime.dll")
        add-type -Path $refs
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
$DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('domain')
if (test-path "c:\scripts\creds\${env:username}_creds.xml")
{
    
    $creds = Import-Clixml "c:\scripts\creds\${env:username}_creds.xml"
    while (!($ds.ValidateCredentials($creds.UserName,$creds.GetNetworkCredential().password,[System.DirectoryServices.AccountManagement.ContextOptions]::Negotiate)))
    {
        $creds = Get-Credential -Message "Enter valid Sharepoint Online Credentials ex: fljpcnadmin@dow.com"
        $creds | Export-Clixml "c:\scripts\creds\${env:username}_creds.xml"
    }

}
else
{
    $creds = Get-Credential -Message "Enter Sharepoint Online Credentials ex : fljpcnadmin@dow.com"
    $creds | Export-Clixml "c:\scripts\creds\${env:username}_creds.xml"
    while (!($ds.ValidateCredentials($creds.UserName,$creds.GetNetworkCredential().password,[System.DirectoryServices.AccountManagement.ContextOptions]::Negotiate)))
    {
        $creds = Get-Credential -Message "Enter valid Sharepoint Online Credentials ex: fljpcnadmin@dow.com"
        $creds | Export-Clixml "c:\scripts\creds\${env:username}_creds.xml"
    }
}
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
            if ($SerialNumber -match "#.*#")
            { 
                $notes = $matches[0]
            }
            $notes
            $PartNumber = $image.split("\")[$image.split("\").count - 2]
            $SerialNumber = $SerialNumber -replace "#.*#",""
            Write-Verbose "Image  :  $Image" -Verbose
            Write-Verbose "Serial Number : $SerialNumber" -Verbose
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
                if ($notes -ne $null)
                {
                        $items[0]["Notes"] = $notes -replace "#",""
                }
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
                    $att.FileName = $file.name.split("\")[-1] -replace "#.*#",""
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

function Update-Progressbar {
    [CmdletBinding()]
    param (
        #  [parameter(Mandatory=$true)]
        #  $SyncHash,
        [parameter(Mandatory = $true)]
        $Status,
        [parameter(Mandatory = $false)]
        $Percentage
    )
    " progress -  $status" | out-file C:\scripts\output.txt -Append
    if ($Percentage -eq -1) {
        $synchash.progressAsset.Dispatcher.Invoke([action] { $synchash.progressAsset.IsIndeterminate = $true }, "Normal")
    }
    else {
        $synchash.progressAsset.Dispatcher.Invoke([action] { $synchash.progressAsset.Value = $Percentage
                $synchash.progressAsset.IsIndeterminate = $false }, "Normal")
    }
    $synchash.lblProgressAsset.Dispatcher.Invoke([action] { $synchash.lblProgressAsset.Content = $status }, "Normal")

}

<#
function Get-RunspaceData {
    <#
.SYNOPSIS
    Get Runspace data and displays a progress bar
 
.DESCRIPTION
    Detailed explanation of the purpose of this function.
 
.PARAMETER Wait
    The purpose of param1.

.PARAMETER Activity
    Message to show up in Activity of the Write-Progress bar

.PARAMETER Jobs
    The object containing the runspacepools
 
.EXAMPLE
     Get-RunspaceData -Param1 'Value1', 'Value2'

.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author:  Tony Turner


    [CmdletBinding()]
    [OutputType('PSCustomObject')]
    param (
        [switch]$Wait,
        [System.Collections.ArrayList]$runspaces,
        [switch]$Indeterminate
        
    )
    $totalJobs = $runspaces.count
    Do {
        $more = $false
        # Update-Progressbar -Status "Runspaces " -Percentage -1
        $count = $runspaces.count
        # Update-Progressbar -Status "Runspaces $count " -Percentage -1
        if ($Indeterminate) {
            $p = -1
        }
        else {
            $p = ($count / $totalJobs * 100)
        }
        
        Update-Progressbar -Status "Remaining jobs $count/$($totalJobs)"  -Percentage $p
        Foreach ($runspace in $runspaces) {
            If ($runspace.Runspace.isCompleted) {
               
                $runspace.powershell.EndInvoke($runspace.Runspace)
                $runspace.powershell.dispose()
                $runspace.Runspace = $null
                $runspace.powershell = $null
               
            }
            ElseIf ($null -ne $runspace.Runspace ) {
          
                $more = $true
            }
        }

        [void][System.gc]::Collect()
        If ($more -AND $PSBoundParameters['Wait']) {
            Start-Sleep -Milliseconds 100
        }
        #Clean out unused runspace jobs
        $temphash = $runspaces.clone()
        $temphash | Where-Object {
            $null -eq $_.runspace
        } | ForEach-Object {
            # Write-verbose ("Removing {0}" -f $_.computer)
            $Runspaces.remove($_)
        }
        [void][System.gc]::Collect()
    } while ($more -AND $PSBoundParameters['Wait'])
    Update-Progressbar -Status "Closing Runspaces " -Percentage -1
    Get-Runspace | Where-Object { $_.RunspaceAvailability -eq "Available" } | ForEach-Object { $_.close() }
    Update-Progressbar -Status "Clearing Memory " -Percentage -1
    [void][System.gc]::Collect()
    #  Get-Runspace | ? { $_.RunspaceAvailability -eq "Available"} | % { $_.close();$_.Dispose()}
    Update-Progressbar -Status "Done " -Percentage 100
}
#>

#############################################################
function close-OrphanedRunSpaces() {
    Get-Runspace
    Write-output "closing"
    Get-Runspace | Where-Object { $_.RunspaceAvailability -eq "Available" } | % { $_.close(); $_.Dispose() }
    write-output "Closed"
    Get-Runspace
}

#############################################################
function Get-SyncHashValue {
    [CmdletBinding()]
    param (
        #  [parameter(Mandatory=$true)]
        #  $SyncHash,
        [parameter(Mandatory = $true)]
        $Object,
        [parameter(Mandatory = $false)]
        $Property
    )

    if ($TempVar) {
        Remove-Variable TempVar -Scope global
    }

    if ($Property) {
        $SyncHash.$Object.Dispatcher.Invoke([System.Action] { Set-Variable -Name TempVar -Value $($SyncHash.$Object.$Property) -Scope global }, "Normal")
    }
    else {
        $SyncHash.$Object.Dispatcher.Invoke([System.Action] { Set-Variable -Name TempVar -Value $($SyncHash.$Object) -Scope global }, "Normal")
    }

    Return $TempVar
}

#endregion



#Load Custom Functions into Session
$sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
# Load custom functions required for runspaces
$function = Get-Content function:\Get-SiemensLicenseID
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Get-SiemensLicenseID", $function
$SessionState.Commands.add($functionEntry)
$function = Get-Content Function:\Get-SiemensLicenseScan
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Get-SiemensLicenseScan", $function
$SessionState.Commands.add($functionEntry)
$function = Get-Content Function:\Read-BarCode
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Read-BarCode", $function
$SessionState.Commands.add($functionEntry)
$function = Get-Content Function:\Resize-Image
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Resize-Image", $function
$SessionState.Commands.add($functionEntry)
$function = Get-Content Function:\Update-SiemensLicenseCertificate
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Update-SiemensLicenseCertificate", $function
$SessionState.Commands.add($functionEntry)
$function = Get-Content Function:\Get-SyncHashValue
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Get-SyncHashValue", $function
$SessionState.Commands.add($functionEntry)
$function = Get-Content Function:\Update-ProgressBar
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Update-Progressbar", $function
$SessionState.Commands.add($functionEntry)

$assetResults = New-Object System.Collections.ObjectModel.ObservableCollection[object]

$Global:syncHash = [hashtable]::Synchronized(@{ })
$newRunspace = [runspacefactory]::CreateRunspace($sessionstate)
# Add Variables needed for backgroud threads to Synchash
$syncHash.path = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition) #$PWD.path
Set-Location $syncHash.path
$syncHash.assetResults = $assetResults
$syncHash.SessionState = $sessionstate

$newRunspace.Name = "GUI"
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
#Send Synchash variable to background threads
$newRunspace.SessionStateProxy.SetVariable("syncHash", $syncHash)



$psCmd = [PowerShell]::Create().AddScript( {

        [xml]$xaml = @"
  <Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="Symantec License Scanner" MinWidth="526" MinHeight="400" Height="287" Width="475"
  x:Name="mainWindow">
    <Grid x:Name="Grid">
        <DockPanel>
            <Border DockPanel.Dock="Top">
                <Grid Height="92">
                   <Label Content="Output Folder : " HorizontalAlignment="Left" Margin="10,18,0,0" VerticalAlignment="Top" Width="92"/>
                   <TextBox x:Name="txtFileName"  HorizontalAlignment="Stretch" Height="23" Margin="107,18,90,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" />
                   <Button x:Name="btnBrowse" Content="Browse" HorizontalAlignment="Right" Margin="0,18,10,0" VerticalAlignment="Top" Width="75" Height="23"/>
                   <Label x:Name="lblStatus" Content="" HorizontalContentAlignment="Center" Margin="10,38,0,0" VerticalAlignment="Top"/>
                   
                   <ProgressBar x:Name="progressAsset" Height="20" Margin="10,68,10,0" VerticalAlignment="Top" Width="{Binding Width, ElementName=dgResults}"/>
                   <Label x:Name="lblProgressAsset" Content="" Height="26" VerticalContentAlignment="Center" HorizontalAlignment="Left" Margin="10,66,10,0" VerticalAlignment="Top" Width="{Binding Width, ElementName=dgResults}"/>
                   
                </Grid>
            </Border>
            <Border DockPanel.Dock="Bottom">
                <Grid >
                    <DataGrid x:Name="dgResults"  Margin="10,0,10,40" />
                    
                    <TextBlock x:Name="txtBlock"  Visibility="Hidden" Margin="10,0,10,40" Background="DarkGray" TextWrapping="Wrap" ScrollViewer.HorizontalScrollBarVisibility="Auto" ScrollViewer.VerticalScrollBarVisibility="Auto">
                    <Run Text="test"/>
                    </TextBlock>
                    
                    <ToggleButton x:Name="btnHelp" Content="Help" HorizontalAlignment="Right" Visibility="Visible" VerticalAlignment="Bottom" Width="75" Margin="30,0,90,10" />
                    <Button x:Name="btnUpdate" Content="Update" HorizontalAlignment="Right" Margin="0,0,10,10" VerticalAlignment="Bottom" Width="75" IsEnabled="False"/>



                </Grid>
            </Border>
        </DockPanel>
      
       <Popup x:Name="Popup" Margin="10,10,0,13"  HorizontalAlignment="Left" VerticalAlignment="Top" Width="194" Height="200" IsOpen="{Binding IsChecked, ElementName=btnHelp}"
        PlacementTarget="{Binding mainWindow}">  
        <StackPanel>  
            <TextBlock Name="McTextBlock"   
             Background="LightBlue" >  
            This is popup text   
           </TextBlock>  
            <Button Content="This is button on a Pupup" />  
        </StackPanel>  
       </Popup>  

   </Grid>

   
   
</Window>
"@
        #  [xml]$xaml = Get-Content "C:\Users\nk23208\source\repos\PowerShellProject1\PowerShellProject1\WSUSReport.xaml"
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $syncHash.Window = [Windows.Markup.XamlReader]::Load( $reader )
        $form = [Windows.Markup.XamlReader]::Load( $reader )
        $syncHash.Host = $Host
        [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
        [System.Reflection.Assembly]::LoadWithPartialName("WindowsFormsIntegration")
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
        $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | % {
            #Find all of the form types and add them as members to the synchash
            $syncHash.Add($_.Name, $syncHash.Window.FindName($_.Name) )
        }
        $xaml.SelectNodes("//*[@Name]") | % { Set-Variable -Name "WPF$($_.Name)" -Value $form.FindName($_.Name) }
        $Script:JobCleanup = [hashtable]::Synchronized(@{ })
        $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
        #region Background runspace to clean up jobs
        $jobCleanup.Flag = $True
        $newRunspace = [runspacefactory]::CreateRunspace()
        $newRunspace.Name = "Cleanup"
        $newRunspace.ApartmentState = "STA"
        $newRunspace.ThreadOptions = "ReuseThread"
        $newRunspace.Open()
        $newRunspace.SessionStateProxy.SetVariable("jobCleanup", $jobCleanup)
        $newRunspace.SessionStateProxy.SetVariable("jobs", $jobs)
        $jobCleanup.PowerShell = [PowerShell]::Create().AddScript( {
                #Routine to handle completed runspaces
                Do {
                    Foreach ($runspace in $jobs) {
                        If ($runspace.Runspace.isCompleted) {
                            [void]$runspace.powershell.EndInvoke($runspace.Runspace)
                            $runspace.powershell.dispose()
                            $runspace.Runspace = $null
                            $runspace.powershell = $null
                        }
                    }
                    #Clean out unused runspace jobs
                    $temphash = $jobs.clone()
                    $temphash | Where-Object {
                        $_.runspace -eq $Null
                    } | ForEach-Object {
                        $jobs.remove($_)
                    }
                    Start-Sleep -Seconds 1
                } while ($jobCleanup.Flag)
            })
        $jobCleanup.PowerShell.Runspace = $newRunspace
        $jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()
        #endregion Background runspace to clean up jobs

  
        #region WPF Controls
        #region Asset File TextChanged
        $syncHash.txtFileName.add_TextChanged( {
                if ((Test-Path $syncHash.TxtFileName.Text) -and ($syncHash.txtFileName.Text).tolower().endswith("csv")) {
                    #Validate CSV header matches expected format
                    $header = "Asset,Subnumber,Company Code,Cost Center,Serial number,Inventory number,APC FY start,Dep. FY start,Bk.val.FY strt,Acquisition,Dep. for year,Capitalized on,Retirement,Dep.retir.,Curr.bk.val.,Asset description,Transfer,Dep.transfer,Currency,Post-capital.,Dep.post-cap.,Invest.support,Write-ups,Current APC,Accumul. dep."
                    $importHeader = Get-Content $syncHash.TxtFileName.Text | Select -First 1
                    if ($header -eq $importHeader) {
                        $assets = Import-Csv $syncHash.TxtFileName.Text
                        #Convert Array object to ArrayList so that we can remove objects from the list
                        $syncHash.assets = New-Object System.Collections.ArrayList(, $assets)
                        $syncHash.btnUpdate.Dispatcher.Invoke([action] { $syncHash.btnUpdate.IsEnabled = $true }, "Normal")
                        $syncHash.lblStatus.Dispatcher.Invoke([action] { $syncHash.lblStatus.Content = "" }, "Normal")
                    }
                    else {
                        $syncHash.lblStatus.Dispatcher.Invoke([action] { $syncHash.lblStatus.Content = "CSV file doesn't match expected format" }, "Normal")
                        $syncHash.btnUpdate.Dispatcher.Invoke([action] { $syncHash.btnUpdate.IsEnabled = $false }, "Normal")
                    }
                }
                else {
                    $syncHash.btnUpdate.Dispatcher.Invoke([action] { $syncHash.btnUpdate.IsEnabled = $false }, "Normal")
                }
            })
        #endregion Asset File TextChanged

        #region Browse to locate Asset CSV File
        $syncHash.btnBrowse.add_Click( {
                # Loads list of KBs into Listbox.  This will be the KBs that you want to pull patch status for.
                $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog 
                $openfiledialog.InitialDirectory = $syncHash.path
                $OpenFileDialog.filter = "csv Files(*.csv)|*.csv"
                $OpenFileDialog.ShowDialog() | out-null
                if (Test-Path $OpenFileDialog.FileName) {
                    $syncHash.assets = Import-Csv $OpenFileDialog.FileName
                    $syncHash.btnUpdate.Dispatcher.Invoke([action] { $syncHash.btnUpdate.IsEnabled = $true }, "Normal")
                    $synchash.txtFileName.Dispatcher.Invoke([action] { $syncHash.txtFileName.text = $OpenFileDialog.FileName }, "Normal")
                }
                else {
                    $syncHash.btnUpdate.Dispatcher.Invoke([action] { $syncHash.btnUpdate.IsEnabled = $false }, "Normal")
                }
            })
        #endregion Browse to locate Asset CSV File


        #region  Help menu
        $syncHash.btnHelp.add_Click( {
                # Loads list of KBs into Listbox.  This will be the KBs that you want to pull patch status for.
                if ($syncHash.btnHelp.IsChecked) {
                    $syncHash.txtBlock.Visibility = "Visible"
                }
                else {
                    $syncHash.txtBlock.Visibility = "Hidden"
                }
           })
        #endregion  Help menu


        #region Update - Begin processessing Asset List
        $syncHash.btnUpdate.add_Click( {
                $syncHash.btnUpdate.Dispatcher.Invoke([action] { $syncHash.btnUpdate.IsEnabled = $false }, "Normal")
                $newRunspace = [runspacefactory]::CreateRunspace($syncHash.sessionState)
                #$syncHash.reportComputers  = $reportComputers
                $newRunspace.ApartmentState = "STA"
                $newRunspace.Name = "Assets"
                $newRunspace.ThreadOptions = "ReuseThread"
                $newRunspace.Open()
                $newRunspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
                $PowerShell = [PowerShell]::Create().AddScript( {
                        Update-IRAssets 
                   })
                $PowerShell.Runspace = $newRunspace
                [void]$Jobs.Add((
                        [pscustomobject]@{
                            PowerShell = $PowerShell
                            Runspace   = $PowerShell.BeginInvoke()
                        }
                    ))
         })

        #endregion Update - Begin processessing Asset List
        #endregion WPF Controls
   
   
        #region Window Close
        $syncHash.Window.Add_Closed( {
                Write-Verbose 'Halt runspace cleanup job processing'
                $jobCleanup.Flag = $False
                #Stop all runspaces
                $jobCleanup.PowerShell.Dispose()
         })

        #endregion Window Close
        #endregion Boe's Additions

        #$x.Host.Runspace.Events.GenerateEvent( "TestClicked", $x.test, $null, "test event")

        #$syncHash.Window.Activate()
        $syncHash.Window.ShowDialog() | Out-Null
        $syncHash.Error = $Error
    })
$psCmd.Runspace = $newRunspace

1..5 | ForEach-Object { write-host "!!!!!!!!!!!!!!!!!!" }
Write-host "to run display your UI, run:  " -NoNewline
write-host -foregroundcolor Green '$data = $psCmd.BeginInvoke()'
$data = $psCmd.BeginInvoke()
Start-sleep -Milliseconds 500
[void](close-OrphanedRunSpaces)






