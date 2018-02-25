# This script works with Powershell 2.0
    ############################################################
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
    $CSAD = @()
    }
    Process
    {
      try 
      {
        
        $csv = $csv = Convertfrom-csv (schtasks /query  /fo csv /v ) | ? {$_.TaskName -like "*CSAD-Task*"} | select TaskName, Status,Frequency, Schedule
      }
      catch
      {
        $csv =  "" | select TaskName, Status,Frequency, Schedule
        $csv.TaskName = "Not Installed" 
          
      }
      If ($csv -eq $null)
      {
        $csv =  "" | select TaskName, Status,Frequency, Schedule
        $csv.TaskName = "Not Installed" 
      }
      $csad += $csv
    }
    End
    {
    return $csad
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
    
function Get-SyncHashValue
{
    [CmdletBinding()]
    param (
      #  [parameter(Mandatory=$true)]
      #  $SyncHash,
        [parameter(Mandatory=$true)]
        $Object,
        [parameter(Mandatory=$false)]
        $Property
    )

    if ($TempVar)
    {
        Remove-Variable TempVar -Scope global
    }

    if ($Property)
    {
        $SyncHash.$Object.Dispatcher.Invoke([System.Action]{Set-Variable -Name TempVar -Value $($SyncHash.$Object.$Property) -Scope global},"Normal")
    }
    else
    {
        $SyncHash.$Object.Dispatcher.Invoke([System.Action]{Set-Variable -Name TempVar -Value $($SyncHash.$Object) -Scope global},"Normal")
    }

    Return $TempVar
}


$Global:syncHash = [hashtable]::Synchronized(@{})
$initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$function = Get-Content function:\Verify-BIT9
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Verify-Bit9", $function
$initialSessionState.Commands.add($functionEntry)
$function = Get-Content function:\verify-CSADTask
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Verify-CSADTask", $function
$initialSessionState.Commands.add($functionEntry)
$function = Get-Content function:\verify-DOWRootCA
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Verify-DOWRootCA", $function
$initialSessionState.Commands.add($functionEntry)
$function = Get-Content function:\Get-SEPINFO
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Get-SEPInfo", $function
$initialSessionState.Commands.add($functionEntry)
$function = Get-Content function:\GET-WSUSInfo
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Get-WSUSInfo", $function
$initialSessionState.Commands.add($functionEntry)

$newRunspace =[runspacefactory]::CreateRunspace($initialSessionState)
$Global:syncHash.Path = (Get-Location).Path
##$newRunspace.Name = "GUI"
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$syncHash.InitialSessionState = $initialSessionState
$newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)

$psCmd = [PowerShell]::Create().AddScript({
  
  $inputXML = @"
 <Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

  Title="Lab PC Check" Height="486.885" Width="782.377" ShowInTaskbar="False">

    <Grid>
        <Button x:Name="btnBIT9" Content="BIT9" HorizontalAlignment="Left" Margin="10,24,0,0" VerticalAlignment="Top" Width="75"/>
        <TextBlock x:Name="tbBIT9_1" HorizontalAlignment="Left" Margin="112,24,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="69" Width="273" FontSize="8"  Text=""/>
        <ListView x:Name="grdCSD" HorizontalAlignment="Left" Margin="112,100,0,0" VerticalAlignment="Top" FontSize="8" MinWidth="300">
           <ListView.View>
                <GridView>
                    <GridViewColumn DisplayMemberBinding="{Binding TaskName}" Header="Task"/>
                    <GridViewColumn DisplayMemberBinding="{Binding Status}" Header="Status"/>
                    <GridViewColumn DisplayMemberBinding="{Binding Frequency}" Header="Frequency"/>
                    <GridViewColumn DisplayMemberBinding="{Binding Schedule}" Header="Schedule"/>
                </GridView>
            </ListView.View>
        </ListView>
        <ListView x:Name="grdDOWCA" HorizontalAlignment="Left" Margin="112,138,0,0" VerticalAlignment="Top" FontSize="8" MaxHeight="145" MinHeight="50" MinWidth="300">
         <ListView.View>
                <GridView>
                    <GridViewColumn DisplayMemberBinding="{Binding Cert}" Header="Cert"/>
                    <GridViewColumn DisplayMemberBinding="{Binding Name}" Header="Name"/>
                    <GridViewColumn DisplayMemberBinding="{Binding Expiration}" Header="Expiration"/>
                  </GridView>
            </ListView.View>
        </ListView>
        <ListView x:Name="grdSepINFO" HorizontalAlignment="Left" Margin="112,345,0,0" VerticalAlignment="Top" FontSize="8" MinWidth="300">
          <ListView.View>
          <GridView>
                    <GridViewColumn DisplayMemberBinding="{Binding Server}" Header="Server"/>
                    <GridViewColumn DisplayMemberBinding="{Binding Group}" Header="Group"/>
                    <GridViewColumn DisplayMemberBinding="{Binding Version}" Header="Expiration"/>
                    <GridViewColumn DisplayMemberBinding="{Binding DefinitionDate}" Header="Definition Date"/>
                </GridView>
            </ListView.View>
          </ListView>
        <TextBlock x:Name="tbBIT9_2" HorizontalAlignment="Left" Margin="385,24,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Height="69" Width="285" FontSize="8" RenderTransformOrigin="0.419,0.512" />
        <Button x:Name="btnCSD" Content="CSD Info Agent" HorizontalAlignment="Left" Margin="10,100,0,0" VerticalAlignment="Top" Width="75" FontSize="10" />
        <Button x:Name="btnDOWCA" Content="DOW Root CA" HorizontalAlignment="Left" Margin="10,138,0,0" VerticalAlignment="Top" Width="75" FontSize="10"/>
        <Button x:Name="btnSEP" Content="SEP" HorizontalAlignment="Left" Margin="10,345,0,0" VerticalAlignment="Top" Width="75"/>
        <Button x:Name="btnWSUS" Content="WSUS" HorizontalAlignment="Left" Margin="10,401,0,0" VerticalAlignment="Top" Width="75"/>
        <TextBlock x:Name="tbWSUS" HorizontalAlignment="Left" Margin="134,401,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Width="372"/>
    </Grid>

</Window>
"@

    Push-Location $syncHash.path
 	$syncHash.Host = $Host

    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
	[System.Reflection.Assembly]::LoadWithPartialName("WindowsFormsIntegration")
	[void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Controls.Data")
	$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
    [xml]$XAML = $inputXML
    #[xml]$xaml 
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
	$form=[Windows.Markup.XamlReader]::Load( $reader )
        $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
        #Find all of the form types and add them as members to the synchash
        $syncHash.Add($_.Name,$syncHash.Window.FindName($_.Name) )
    }
	#$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $form.FindName($_.Name)}
    $Script:JobCleanup = [hashtable]::Synchronized(@{})
    $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

    #region Background runspace to clean up jobs
    $jobCleanup.Flag = $True
    $newRunspace =[runspacefactory]::CreateRunspace()
    #$newRunspace.Name = "Cleanup"
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"
    $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("jobCleanup",$jobCleanup)
    $newRunspace.SessionStateProxy.SetVariable("jobs",$jobs)
    $jobCleanup.PowerShell = [PowerShell]::Create().AddScript({
        #Routine to handle completed runspaces
        Do {
            Foreach($runspace in $jobs) {
                If ($runspace.Runspace.isCompleted) {
                    [void]$runspace.powershell.EndInvoke($runspace.Runspace)
                    $runspace.powershell.dispose()
                    $runspace.Runspace = $null
                    $runspace.powershell = $null
                }
            }
            #Clean out unused runspace jobs
            $temphash = $jobs.clone()
            $temphash | Where {
                $_.runspace -eq $Null
            } | ForEach {
                $jobs.remove($_)
            }
            Start-Sleep -Seconds 1
        } while ($jobCleanup.Flag)
    })
    $jobCleanup.PowerShell.Runspace = $newRunspace
    $jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()
    #endregion Background runspace to clean up jobs


    $syncHash.btnBIT9.Add_Click({
      Try
{
    $bit9 = & 'C:\Program Files (x86)\Bit9\Parity Agent\DasCLI.exe' status 
    $version = ($bit9 | Out-String) -match "(?ms)^Version.*?Cache"
    if ($matches.count -gt 0 ) 
{
    #write-host "BIT9 Installed" 
    $Bit9 = $matches[0]
}
}
Catch
{
  #Write-host "BIT9 not installed"
  $bit9 =  "BIT9 not Installed"
}
$syncHash.tbBIT9_1.text = $bit9.split("`n")[0..4]
 $syncHash.tbBIT9_2.text = $bit9.split("`n")[6..9]
 
})


 $syncHash.btnCSD.Add_Click({
		$CSD = Verify-CSADTask
        $CSAD = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        $C =  "" | select TaskName, Status,Frequency, Schedule
        $c.TaskName = $csd.TaskName
        $c.Frequency = $csd.Frequency
        $c.Status = $csd.Status
        $c.Schedule = $csd.Schedule
        $csad.add($csd)
        $syncHash.grdCSD.ItemsSource = $csad
       

    
    })
#>
    $syncHash.btnDOWCA.Add_Click({
		$cert = Verify-DOWRootCA
        $certs = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        $cert | % {
        $C =  "" | select Cert, Name, Expiration
        $c.Cert = $_.cert
        $c.Name = $_.Name
        $c.Expiration = $_.Expiration
        $certs.add($c)
        }
        $syncHash.grdDOWCA.ItemsSource = $certs
       })
	

    $syncHash.btnSEP.Add_Click({
	    $sep = Get-SEPInfo
        #$sepinfo = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        $sepinfo = @()
        $sepinfo += $sep
        $syncHash.grdSepINFO.ItemsSource = $sepinfo
       })
	
    $syncHash.btnWSUS.Add_Click({
	    $wsus = Get-WSUSInfo
       $syncHash.tbwsus.text = $wsus.server
    })
	
#>
    #region Window Close
    $synchash.tbBit9_1.Text = get-location
    $syncHash.Window.Add_Closed({
        Write-Verbose 'Halt runspace cleanup job processing'
        $jobCleanup.Flag = $False

        #Stop all runspaces
        $jobCleanup.PowerShell.Dispose()
    })
    #endregion Window Close
    #endregion Boe's Additions

 
  
    $syncHash.btnBIT9.RaiseEvent((new-object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
    $syncHash.btnCSD.RaiseEvent((new-object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
    $syncHash.btnDOWCA.RaiseEvent((new-object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
    $syncHash.btnSEP.RaiseEvent((new-object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
    $syncHash.btnWSUS.RaiseEvent((new-object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
    #>
    #$syncHash.Window.Activate()
    $syncHash.Window.ShowDialog() | Out-Null
    $syncHash.Error = $Error

})
$psCmd.Runspace = $newRunspace

1..5 | %{ write-host "!!!!!!!!!!!!!!!!!!"}
Write-host "to run display your UI, run:  " -NoNewline
write-host -foregroundcolor Green '$data = $psCmd.BeginInvoke()'
$data = $psCmd.BeginInvoke()
Write-host "Please wait while data is gathered.   GUI will load when finished"
sleep -Milliseconds 500
    

function close-OrphanedRunSpaces()
{
   Get-Runspace
   Write-Host "closing"
    Get-Runspace | ? { $_.RunspaceAvailability -eq "Available"} | % { $_.close();$_.Dispose()}
   write-host "Closed"
   Get-Runspace
}




