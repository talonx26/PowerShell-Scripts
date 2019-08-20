Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Publishing.dll"


#Read more: https://www.sharepointdiary.com/2016/10/sharepoint-online-check-if-site-collection-subsite-exists-powershell-csom.html#ixzz5vvCQWItK
<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function HelperSPOSubSite {
    [CmdletBinding()]
    [Alias()]
    
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$Parent = "https://workspaces.bsnconnect.com/sites/LabAutomation/LAInternal/Projects",

        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Blank", "Project")]
        [string]$SiteType
    )
    

    Begin {  
        # $refs = @("C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll", 
        #      "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll")
        # add-type -Path $refs
       
       $creds = Get-SharePointCredentials
        $ProjectTemplate = "{FABDC41D-933A-4A22-AC37-8F3054D9F084}#SharePoint Project Management Template"
        $BlankTemplate = "STS#1"
      
    }
    Process {
        
        If (!(Check-SiteExists -SiteURL "$parent/$($name.Replace(' ',''))" -Credentials $creds))
        {
            # create Context
            $context = New-Object Microsoft.SharePoint.Client.ClientContext($Parent)
            $context.Credentials = $Creds

            #create Subsite object
            $Subsite = New-Object Microsoft.SharePoint.Client.WebCreationInformation
        
            # Sets Subsite Title to name and set URL to the Name but removes all spaces
            $Subsite.Title = $Name
            $Subsite.url = $Name.Replace(' ', '')
      
            # Sets Language to English 
            $Subsite.Language = "1033"
            $Subsite.UseSamePermissionsAsParentSite = $true
            If ($SiteType -eq "Blank") 
            {
                $Subsite.WebTemplate = $BlankTemplate
                $subweb = $context.web.webs.Add($subsite)
                $context.Load($subweb)
                $context.ExecuteQuery()
            }
            elseIf ( $SiteType -eq "Project") 
            {
                $Subsite.WebTemplate = $ProjectTemplate
                $subweb = $context.web.webs.Add($subsite)
                $context.Load($subweb)
                $context.ExecuteQuery()
            
            }
        }
        
        #Clean up the Page
        $subContext = New-Object Microsoft.SharePoint.Client.ClientContext("$parent/$($name.Replace(' ',''))")
        $subContext.Credentials = $creds
        $subWeb = $subContext.web
        $subContext.Load($subweb)
        $subContext.ExecuteQuery()
        If ($SiteType -eq "Blank")
        {
            
            $quicklaunch = $subweb.Navigation.QuickLaunch
            $subContext.load($quicklaunch)
            $subContext.ExecuteQuery()
            foreach ($link in $quicklaunch) {
                    $subContext.Load($link)
                    $subContext.load($link.Children)
                    $subContext.ExecuteQuery()
            
                    $link.Title
                    if ($link.Title -eq "Home") {
                        $p = $subweb.ParentWeb
                        $subContext.load($p)
                        $subContext.ExecuteQuery()
                        $link.Title = $p.Title
                        $link.Url = "Https://workspaces.bsnconnect.com/$($p.ServerRelativeUrl)"
                        $link.Update()
                        $subContext.ExecuteQuery()                
                    }
                    if ($link.Title -eq "Site Contents") { 
                        $link.DeleteObject()
                        $link.Update()
                       #throws error but everything works.  Try catch to hide error output
                       try{
                        $subContext.ExecuteQuery() | Out-Null
                        }
                        catch{}
                    }
          
                    foreach ($child in $link.Children) {
                        $subContext.Load($child)
                        $subContext.ExecuteQuery()
                        $child.title

                    }
               



        }
    }
        if ($SiteType -eq "Project")
        {
         # Clean up Project Tasks from dummy information
            $List = $subWeb.Lists.GetByTitle("Project Tasks")
            $subContext.Load($list)
            $subContext.ExecuteQuery()
            $Query = New-Object Microsoft.SharePoint.Client.CamlQuery
            $Query.ViewXml = '<Query><Where><IsNull><FieldRef Name="ParentID"/></IsNull></Where></Query>'
            $listItems = $List.GetItems($query)
            $subContext.load($listItems)
            $subContext.ExecuteQuery()  
            foreach ($item in $listItems)
        {
           #Delete all parent Tasks to clear dummy information
           if ( $item["ParentID"] -eq $null)
           {
              
              $List.GetItemById($item.ID).deleteobject()
              $list.Update()
            }
        }
        try
        {
            $subcontext.ExecuteQuery()
        }
        catch
        {}

        #Delete Getting Started WebPart
        $file = $subContext.web.GetFileByServerRelativeUrl("$($subContext.url.Replace("https://workspaces.bsnconnect.com",''))/SitePages/Home.aspx")
        $subContext.load($file)
        $subContext.ExecuteQuery()
        $wpManager = $file.GetLimitedWebPartManager([Microsoft.SharePoint.Client.WebParts.PersonalizationScope]::Shared)
        $webparts = $wpManager.Webparts  
        $subContext.Load($webparts)  
        $subContext.ExecuteQuery() 
        if($webparts.Count -gt 0){  
        foreach($webpart in $webparts){  
        $subContext.Load($webpart.WebPart.Properties)  
        $subContext.ExecuteQuery()  
        $propValues = $webpart.WebPart.Properties.FieldValues  
        if ($webpart.WebPart.Properties.FieldValues["Title"] -eq "Get started with your project")
            {
                    $webpart.DeleteWebPart()
                    $subContext.ExecuteQuery()
            }
             
    }  
}  
        }

        # Edit Navigation in Parent site
        If ((Check-SiteExists -SiteURL "$parent/$($name.Replace(' ',''))" -Credentials $creds))
        {
           Add-SPQuickLaunch -Parent $Parent -Node $Name -Header Blank
        }
    }
    End {
       # $subweb
    }
}





Function Add-SPQuickLaunch{
[CmdletBinding()]
    [Alias()]
    
    Param
    (
        # Param1 help description
        
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$Parent = "https://workspaces.bsnconnect.com/sites/LabAutomation/LAInternal/Projects/ProjectTest",
        [Parameter(Mandatory = $true)]
        [string]$Node,
        [Parameter(Mandatory = $true)]
        [ValidateSet("Blank", "Hood")]
        [string]$Header
    )
    Begin
    {
         $creds = Get-SharePointCredentials
    }

    Process 
    {
            #Get Context and Web, Quicklaunch object
            $context = New-Object Microsoft.SharePoint.Client.ClientContext($Parent)
            $context.Credentials = $Creds
            $quicklaunch = $context.web.Navigation.QuickLaunch
            $web = $context.web
            $context.load($web)
            $context.load($quicklaunch)
            $context.ExecuteQuery()
            $link = $null
            #Locate Parent Node
            $link = $quicklaunch | ? { $_.Title -eq $web.Title}
           
            #If Parent node is not found create new Parent Node
            if ($Link -eq $null)
            {
                $navigationNode = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
                $navigationNode.Title = $web.Title
                $navigationnode.Url = $null
                $navigationNode.AsLastNode = $true
                $context.load($quicklaunch.Add($navigationNode))
                $context.ExecuteQuery()
                
                #Search Quicklaunch for Parent Node after creation
                $link = $quicklaunch | ? { $_.Title -eq $web.Title}

            }
            # Get Subsite con
            $subContext = New-Object Microsoft.SharePoint.Client.ClientContext("$parent/$($node.Replace(' ',''))")
            $subcontext.Credentials = $Creds
            $subweb = $subContext.Web
            $subContext.Load($subweb)
            $subcontext.ExecuteQuery()
            $context.load($link)
            $context.ExecuteQuery()
            # Attempt to load the children.  Inside of try block because it will fail if there are no child links
            Try{$context.load($link.children)
                $context.ExecuteQuery()
                $newNode = $link.Children | ? { $_.title -eq $subweb.Title } 
                }
            catch {}
            if ($newnode  -eq $null)
            {
                
                 $navigationNode = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation

            $navigationNode.Title = $subweb.Title
            $navigationNode.url = "$parent/$($node.Replace(' ',''))"
            $navigationNode.AsLastNode = $true
            $context.Load( $link.children.add($navigationNode))
            $context.ExecuteQuery()


            }
           
            #$context.load($link)
            #$contest.load($link.children)
            
           
           
    }

    End 
    {

    }


}




Function Add-SubsitesToQuickLaunch{
[CmdletBinding()]
    [Alias()]
    
    Param
    (
        # Param1 help description
        
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$Parent = "https://workspaces.bsnconnect.com/sites/LabAutomation/LAInternal/Projects/ProjectTest"
       
    )
    Begin
    {
         $creds = Get-SharePointCredentials
    }

    Process 
    {
            #Get Context and Web, Quicklaunch object
            $context = New-Object Microsoft.SharePoint.Client.ClientContext($Parent)
            $context.Credentials = $Creds
            $web = $context.web
            $subwebs = $web.webs
            $context.load($web)
            $context.load($subwebs)
            $context.ExecuteQuery()

            # if This site contains subsites then add them to the quick launch
            If ($subwebs.count -gt 0)
            {
                $quicklaunch = $context.web.Navigation.QuickLaunch
                $context.load($quicklaunch)
                $context.ExecuteQuery()
                $link = $null
                #Locate Parent Node
                $link = $quicklaunch | ? { $_.Title -eq $web.Title}
                  #If Parent node is not found create new Parent Node
                if ($Link -eq $null)
                {
                    $navigationNode = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation
                    $navigationNode.Title = $web.Title
                    $navigationnode.Url = ""
                    $navigationNode.AsLastNode = $false
                    $context.load($quicklaunch.Add($navigationNode))
                    $context.ExecuteQuery()
                
                    #Search Quicklaunch for Parent Node after creation
                    $link = $quicklaunch | ? { $_.Title -eq $web.Title}

                }
                foreach ($subweb in $subwebs)
                {
                   # $subContext = New-Object Microsoft.SharePoint.Client.ClientContext("$parent/$($node.Replace(' ',''))")
                   # $subcontext.Credentials = $Creds
                   # $subweb = $subContext.Web
                   # $subContext.Load($subweb)
                   # $subcontext.ExecuteQuery()

                    $context.load($link)
                    $context.ExecuteQuery()
                    # Attempt to load the children.  Inside of try block because it will fail if there are no child links
                    Try{$context.load($link.children)
                        $context.ExecuteQuery()
                        $newNode = $link.Children | ? { $_.title -eq $subweb.Title } 
                    }
                    catch {}
                    if ($newnode  -eq $null)
                    {
                
                        $navigationNode = New-Object Microsoft.SharePoint.Client.NavigationNodeCreationInformation

                        $navigationNode.Title = $subweb.Title
                        $navigationNode.url = $subweb.url
                        $navigationNode.AsLastNode = $false
                        $context.Load( $link.children.add($navigationNode))
                        $context.ExecuteQuery()


                    }
                    Add-SubsitesToQuickLaunch -Parent $subweb.url
           
                }


            }


         
           
          
            # Get Subsite con
           
            
            #$context.load($link)
            #$contest.load($link.children)
            
           
           
    }

    End 
    {

    }


}



Function Check-SiteExists ( $SiteURL, $Credentials) {
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $ctx.Credentials = $Credentials
    $web = $ctx.Web
    $ctx.Load($web)
    try {
        $ctx.ExecuteQuery()
        return $true
    }
    catch {
        return $false
    }


}




Function Add-SPSubSite
{
 [CmdletBinding()]
    [Alias()]
    
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $false, Position = 0)]
        [string]$Parent = "https://workspaces.bsnconnect.com/sites/LabAutomation/LAInternal/Projects",

        [Parameter(Mandatory = $true)]
        [string]$Sites
    )

    Begin
    {
        $creds = Get-SharePointCredentials
    }
    Process
    {
        $subsites = $sites.Split('/')
        $p = $Parent
        foreach($site in $subSites)
        {
            
         #   $p
            #write-host "$p   $(Check-SiteExists -SiteUrl $p -Credentials $creds )"
            
            If (!(Check-SiteExists -SiteUrl "$p/$($site.Replace(' ',''))" -Credentials $creds ))
            {
               $s = "$p/$($site.Replace(' ',''))"
                if ($subsites.indexof($site) -eq ($subsites.length-1))
                {
                   write-host "creating site : $s"
                    HelperSPOSubSite -Parent $p -Name $site -SiteType Project 
                }
                else
                {
                    write-host "creating site : $s"
                    HelperSPOSubSite -Parent $p -Name $site -SiteType Blank

                }
            }
            $p = "$p/$($site.Replace(' ',''))"  
            $p
        }
    }
    End
    {

    }


    


}

$site = "project test/Texas/Lake Jackson/ECB/Mod 36/Hood 3"
 