#Requires -Version 3.0
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
        $ErrorActionPreference = "Continue"
        #Load SharePoint DLL's
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
        $webURL = "https://workspaces.bsnconnect.com/teams/LabAutomation"
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
