#Requires -Version 3.0
function Sync-SharePointProjectTree
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
     Update-SPLabInfo -Param1 'Value1', 'Value2'

.EXAMPLE
     'Value1', 'Value2' | Update-SPLabInfo

.EXAMPLE
     Update-SPLabInfo -Param1 'Value1', 'Value2' -Param2 'Value'
 
.INPUTS
    String
 
.OUTPUTS
    PSCustomObject
 
.NOTES
    Author: Tony Turner
#>

    [CmdletBinding()]
    [OutputType('PSCustomObject')]
    param (
        [Parameter(Mandatory=$false, 
            ValueFromPipeline)]
        [object[]]$Computers
    )
 
    BEGIN
    {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement
        $DS = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('domain')
        if (test-path "c:\scripts\creds\${env:username}_spo_creds.xml")
        {
            
            $creds = Import-Clixml "c:\scripts\creds\${env:username}_spo_creds.xml"
           while (!($ds.ValidateCredentials($creds.UserName,$creds.GetNetworkCredential().password,[System.DirectoryServices.AccountManagement.ContextOptions]::Negotiate)))
            {
                $creds = Get-Credential -Message "Enter valid Sharepoint Online Credentials ex: fljpcnadmin@dow.com"
                $creds | Export-Clixml "c:\scripts\creds\${env:username}_spo_creds.xml"
            }
        
        }
        else
        {
            $creds = Get-Credential -Message "Enter Sharepoint Online Credentials ex : fljpcnadmin@dow.com"
            $creds | Export-Clixml "c:\scripts\creds\${env:username}_spo_creds.xml"
           while (!($ds.ValidateCredentials($creds.UserName,$creds.GetNetworkCredential().password,[System.DirectoryServices.AccountManagement.ContextOptions]::Negotiate)))
            {
                $creds = Get-Credential -Message "Enter valid Sharepoint Online Credentials ex: fljpcnadmin@dow.com"
                $creds | Export-Clixml "c:\scripts\creds\${env:username}_spo_creds.xml"
            }
        }
        #Used for prep. This code runs one time prior to processing items specified via pipeline input.
        #Load Sharepoint Library Files
        $refs = @(".\Microsoft.SharePoint.Client.dll", ".\Microsoft.SharePoint.Client.Runtime.dll")
        add-type -Path $refs
        # Sharepoint Web Address and Login information
        $webURL = "https://workspaces.bsnconnect.com/sites/LabAutomation/LAInternal/Projects"
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($webURL)
        $credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)
        $context.Credentials = $credentials
        # Clear any computer objects with Model containing 'Unable'  - clears any objects with no data.
      #  $computers = $computers | ? { $_.model -notlike "Unable*"}

        #Calls function to Update Bios and Model Lists in Sharepoint
        #Function will got thru entire object and update lists and then return an master object with the Record Ids to be used
        #to link in the Computer Inventory List
       # $ModelID = Update-SPModel -Computers $Computers
       # $BiosID = Update-SPBIOS -Computers $computers
    }

    PROCESS
    {
        #This code runs one time for each item specified via pipeline input.

        
            #$VerbosePreference = $true
           <#
            Write-Verbose "$computer"
            $percentCounter++
            Write-Verbose "writing progress"
            write-progress -ParentId 1 -Activity "Processing Computer $computer" -status "Updating SharePoint for $computer" -PercentComplete (($percentCounter / ($Computers | Measure-Object).count) * 100)
            Write-Verbose "Updating $computer"
            #>
            $web = $Context.Web
            #Sharepoint list to get information from
            $weblist = "Tasks"
            $Context.Load($web) 
            $Context.ExecuteQuery() 
            $list = $web.Lists.GetByTitle($weblist)
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            #Sharepoint Query
            $query.ViewXml = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery().Viewxml
            
            $items = $list.GetItems($Query)  
            $context.Load($items)
            $context.ExecuteQuery()
            # If query came up with 0 items then record doesn't exist.  Create initial record with minimum required data
            Foreach ($item in $items)

            {                
                $x = 0    
               $site = $item["Project_x0020_Site"].url
               if ($site -ne $null)
               {
               $subcontext = New-Object Microsoft.SharePoint.Client.ClientContext($site)
               $status = Get-SPProjectStatus -site $site -Task $item["Title"].trim()
               If ($status -ne $null)
               {
                    $item["PercentComplete"] = $status.PercentComplete
                    $item["Business"] = $status.business
                    $item["Project_x0020_Site"].Description = "Link"
                    $item["Project_x0020_Status"] = $status.status
                    $item["DueDate"] = $status.DueDate
                    $item["AssignedTo"] = $status.AssignedTo
                    $item["Location"] = $status.location
                    $item["FTE"] = $status.FTE
                    $item["Global_x0020_ID"] = $status.GlobalID
                    $item.Update()
                    $Context.ExecuteQuery()
                    #Reload Items to get new Record ID
                   

               }
              }
            }
            

            
        
    }

    END
    {
        #Used for cleanup. This code runs one time after all of the items specified via pipeline input are processed.
    }


}
#Update-SPLabPCInfo -Computers $compinfo