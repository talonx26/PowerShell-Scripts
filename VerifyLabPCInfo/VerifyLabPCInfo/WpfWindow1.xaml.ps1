    #ERASE ALL THIS AND PUT XAML BELOW between the @" "@ 
$inputXML = @"
<Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

  Title="MainWindow" Height="486.885" Width="705.328">

    <Grid>
        <TextBlock x:Name="tbBIT9_1" HorizontalAlignment="Left" Margin="112,24,0,0" TextWrapping="Wrap" Text="test" VerticalAlignment="Top" Height="122" Width="273" FontSize="8"/>
        <Label Content="BIT9" HorizontalAlignment="Left" Margin="10,19,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.11,0.039"/>
        <DataGrid x:Name="grdCSD" HorizontalAlignment="Left" Height="34" Margin="292,189,0,0" VerticalAlignment="Top" Width="336"/>
        <DataGrid x:Name="grdDOWCA" HorizontalAlignment="Left" Height="100" Margin="292,240,0,0" VerticalAlignment="Top" Width="336"/>
        <DataGrid x:Name="grdSepINFO" HorizontalAlignment="Left" Height="43" Margin="292,345,0,0" VerticalAlignment="Top" Width="336"/>
        <TextBlock x:Name="tbBIT9_2" HorizontalAlignment="Left" Margin="385,24,0,0" TextWrapping="Wrap" Text="test" VerticalAlignment="Top" Height="122" Width="285" FontSize="8" RenderTransformOrigin="0.419,0.512"/>
    </Grid>

</Window>
 
"@ 
$inputXML = Get-Content -Path .\WpfWindow1.xaml
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch [System.Management.Automation.MethodInvocationException] {
    Write-Warning "We ran into a problem with the XAML code.  Check the syntax for this control..."
    write-host $error[0].Exception.Message -ForegroundColor Red
    if ($error[0].Exception.Message -like "*button*"){
        write-warning "Ensure your &lt;button in the `$inputXML does NOT have a Click=ButtonClick property.  PS can't handle this`n`n`n`n"}
}
catch{#if it broke some other way😀
    Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
        }
 
#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
 
Get-FormVariables
 
#===========================================================================
    # Use this space to add code to the various form elements in your GUI
    #===========================================================================
                                                                    
     
    #Reference 
 
    #Adding items to a dropdown/combo box
      #$vmpicklistView.items.Add([pscustomobject]@{'VMName'=($_).Name;Status=$_.Status;Other="Yes"})
     
    #Setting the text of a text box to the current PC name    
      #$WPFtextBox.Text = $env:COMPUTERNAME
     
    #Adding code to a button, so that when clicked, it pings a system
    # $WPFbutton.Add_Click({ Test-connection -count 1 -ComputerName $WPFtextBox.Text
    # })
    #===========================================================================
    # Shows the form
    #===========================================================================
write-host "To show the form, run the following" -ForegroundColor Cyan
'$Form.ShowDialog() | out-null'
 
 

 
function Load-Xaml {
$PSScriptRoot
	[xml]$xaml = Get-Content -Path .\WpfWindow1.xaml
	$manager = New-Object System.Xml.XmlNamespaceManager -ArgumentList $xaml.NameTable
	$manager.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml");
	$xamlReader = New-Object System.Xml.XmlNodeReader $xaml
	[Windows.Markup.XamlReader]::Load($xamlReader)
}

$window = Load-Xaml
#$window.ShowDialog()



#This script will check to if all necessary stuff is installed



#Checking CSD Agent
<#
.Synopsis
   Check to see if CSAD Agent is installed
.DESCRIPTION
   Check to see if CSAD Agent is installed
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Verify-CSADTask
{
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Computer=$env:COMPUTERNAME
    )

    Begin
    {
    }
    Process
    {
      try 
      {
        
        $csv = $csv = Convertfrom-csv (schtasks /query  /fo csv /v ) | ? {$_.TaskName -like "*CSAD-Task*"} | select TaskName, Status,"Schedule Type", Schedule
      }
      catch
      {
        $csv =  "" | select TaskName, Status,"Schedule Type", Schedule
        $csv.TaskName = "Not Installed" 
          
      }
      If ($csv -eq $null)
      {
        $csv =  "" | select TaskName, Status,"Schedule Type", Schedule
        $csv.TaskName = "Not Installed" 
      }
    }
    End
    {
    return $csv
    }
    
}




<#
.Synopsis
   Will check to see if the 9 DOW CA certs are installed
.DESCRIPTION
   Will check to see if the 9 DOW CA certs are installed
.EXAMPLE
   Find-DOWRootCA
#>
function Verify-DOWRootCA

{
    [CmdletBinding()]
    Param()
    Begin
    {
     $Certs = @()
    }
    Process
    {
     push-Location cert:\CurrentUser\Root
     $c = "" | select Cert, Name, Expiration    
     
     $cert = Get-ChildItem | ? { $_.subject -like "*DOW Chemical Production Root CA*"} | select Subject, @{Name ="Expiration Date";Expression ={ $_.NotAfter}} -First 1
     $c = "" | select Cert, Name, Expiration
     $c.Cert = "DOW Prod Root CA"
     $c.Name = $cert.Subject
     $c.Expiration = $cert.'Expiration Date'
     $certs += $c
     
     $cert = Get-ChildItem | ? { $_.subject -like "*BSNC Root Authority*"} | select Subject, @{Name ="Expiration Date";Expression ={ $_.NotAfter}} -First 1
     $c = "" | select Cert, Name, Expiration
     $c.Cert = "BSN Root CA"
     $c.Name = $cert.Subject
     $c.Expiration = $cert.'Expiration Date'
     $certs += $c
     
     $cert = Get-ChildItem | ? { $_.subject -like "*DOW Chemical Root CA*"} | select Subject, @{Name ="Expiration Date";Expression ={ $_.NotAfter}} -First 1
     $c = "" | select Cert, Name, Expiration
     $c.Cert = "DOW Root CA"
     $c.Name = $cert.Subject
     $c.Expiration = $cert.'Expiration Date'
     $certs += $c

     $cert = Get-ChildItem | ? { $_.subject -like "*DOW Corning PKI Production Root CA - 02*"} | select Subject, @{Name ="Expiration Date";Expression ={ $_.NotAfter}} -First 1
     $c = "" | select Cert, Name, Expiration
     $c.Cert = "DOW Corning PKI Root CA"
     $c.Name = $cert.Subject
     $c.Expiration = $cert.'Expiration Date'
     $certs += $c

     push-Location cert:\CurrentUser\CA
     $cert = Get-ChildItem | ? { $_.subject -like "*DOW Chemical Production SSL CA*"} | select Subject, @{Name ="Expiration Date";Expression ={ $_.NotAfter}} -First 1
     $c = "" | select Cert, Name, Expiration
     $c.Cert = "DOW SSL CA"
     $c.Name = $cert.Subject
     $c.Expiration = $cert.'Expiration Date'
     $certs += $c
     
     $cert = Get-ChildItem | ? { $_.subject -like "*BSNC Mach CA*"} | select Subject, @{Name ="Expiration Date";Expression ={ $_.NotAfter}} -First 1
     $c = "" | select Cert, Name, Expiration
     $c.Cert = "BSNC Mach CA"
     $c.Name = $cert.Subject
     $c.Expiration = $cert.'Expiration Date'
     $certs += $c
     
     $cert = Get-ChildItem | ? { $_.subject -like "*DOW Corning Issuing CA 31*"} | select Subject, @{Name ="Expiration Date";Expression ={ $_.NotAfter}} -First 1
     $c = "" | select Cert, Name, Expiration
     $c.Cert = "DOW Corning CA 31"
     $c.Name = $cert.Subject
     $c.Expiration = $cert.'Expiration Date'
     $certs += $c
     
     $cert = Get-ChildItem | ? { $_.subject -like "*DOW Corning Issuing CA 32*"} | select Subject, @{Name ="Expiration Date";Expression ={ $_.NotAfter}} -First 1
     $c = "" | select Cert, Name, Expiration
     $c.Cert = "DOW Corning CA 32"
     $c.Name = $cert.Subject
     $c.Expiration = $cert.'Expiration Date'
     $certs += $c
     
     $cert = Get-ChildItem | ? { $_.subject -like "*DOW Chemical Issuing CA*"} | select Subject, @{Name ="Expiration Date";Expression ={ $_.NotAfter}} -First 1 
     $c = "" | select Cert, Name, Expiration
     $c.Cert = "DOW Issuing CA"
     $c.Name = $cert.Subject
     $c.Expiration = $cert.'Expiration Date'
     $certs += $c


    }
    End
    {
    Pop-Location
    Pop-Location
    return $certs
    }
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Verify-Bit9
{
    [CmdletBinding()]
    Param
    (
       
    )

    Begin
    {
    }
    Process
    {
    Try
{
    $bit9 = & 'C:\Program Files (x86)\Bit9\Parity Agent\DasCLI.exe' status 
    $version = ($bit9 | Out-String) -match "(?ms)^Version.*?Cache"
    if ($matches.count -gt 0 ) 
{
    #write-host "BIT9 Installed" 
    $matches[0]
}
}
Catch
{
  #Write-host "BIT9 not installed"
  return "BIT9 not Installed"
}

    }
    End
    {
    }
}

<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-SEPInfo
{
    [CmdletBinding()]
    
    Param()

    Begin
    {
    }
    Process
    {
        $RegPath ="hklm:\SOFTWARE\Wow6432Node"
        $SepInfo = "" | Select Server, Group, Version, DefinitionDate
        $sylink = Get-ItemProperty "$regpath\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink"  -ErrorAction SilentlyContinue
       
        # If sylink is null then no data found.  Check for 32BIT version
        If ($sylink -eq $null) 
        {
            $sylink = Get-ItemProperty "hklm:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink"  -ErrorAction SilentlyContinue
            $RegPath = "hklm:\SOFTWARE"

        }
        if ($sylink -ne $null) 
        {
            #Get AV Definition Date
            $av = Get-ItemProperty "$RegPath\Symantec\Symantec Endpoint Protection\AV" -ErrorAction SilentlyContinue
            $SepInfo.DefinitionDate = [datetime]"$($av.PatternFileDate[1]+1) - $($av.PatternFileDate[2]) - $($av.PatternFileDate[0]+1970)" | get-date -format "MM-dd-yyyy"
            $SEP = Get-ItemProperty "$RegPath\Symantec\Symantec Endpoint Protection\CurrentVersion"
            $SEPInfo.Version = $SEP.PRODUCTVERSION
            $sepinfo.Group = $sylink.CurrentGroup
            $Server = $sylink.CommunicationStatus.Split(':')[1]
            if ($Server -match  "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}")
            { 
               try
               { 
                $Server = ([System.Net.Dns]::GetHostbyAddress($matches[0]) ).hostname 
                }
                catch
                {
                 
                }
            }
            $sepinfo.Server =$Server
        }
        else
        {
          $SepInfo.Server = "Not Installed"
        }
    }
    End
    {
     return $SepInfo
    }
}




<#
.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
#>
function Get-WSUSInfo
{
    [CmdletBinding()]
    Param()

    Begin
    {
    }
    Process
    { 
        $WSUS = "" | Select Server
        $wsus.server = (Get-ItemProperty "hklm:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate").WUServer
    }
    End
    {
       return $wsus
    }
}
##Write-host "Checking SEP"
#Get-SEPInfo
#Write-host "Checking WSUS"
#Get-WSUSInfo
#Write-host "Checking BIT9"




 $bit9 = Verify-Bit9
 $bit9
 $WPFtbBIT9_1.text = $bit9.Split("`n")[0..4]
 $WPFtbBIT9_2.text = $bit9.Split("`n")[6..9]

 #$form.ShowDialog()