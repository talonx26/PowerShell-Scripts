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
function Get-CSADTask {
    [CmdletBinding()]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        $Computer = $env:COMPUTERNAME
    )

    Begin {
        $CSAD = @()
    }
    Process {
        try {
        
            $csv = $csv = Convertfrom-csv (schtasks /query  /fo csv /v ) | Where-Object {$_.TaskName -like "*CSAD-Task*"} | Select-Object TaskName, Status, "Schedule Type", Schedule
        }
        catch {
            $csv = "" | Select-Object TaskName, Status, "Schedule Type", Schedule
            $csv.TaskName = "Not Installed" 
          
        }
        If ($csv -eq $null) {
            $csv = "" | Select-Object TaskName, Status, "Schedule Type", Schedule
            $csv.TaskName = "Not Installed" 
        }
        $csad += $csv
    }
    End {
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
function Get-DOWRootCA {
    [CmdletBinding()]
    Param()
    Begin {
        $Certs = @()
    }
    Process {
        push-Location cert:\CurrentUser\Root
        $c = "" | Select-Object Cert, Name, Expiration    
     
        $cert = Get-ChildItem | Where-Object { $_.subject -like "*DOW Chemical Production Root CA*"} | Select-Object Subject, @{Name = "Expiration Date"; Expression = { $_.NotAfter}} -First 1
        $c = "" | Select-Object Cert, Name, Expiration
        $c.Cert = "DOW Prod Root CA"
        $c.Name = $cert.Subject
        $c.Expiration = $cert.'Expiration Date'
        $certs += $c
     
        $cert = Get-ChildItem | Where-Object { $_.subject -like "*BSNC Root Authority*"} | Select-Object Subject, @{Name = "Expiration Date"; Expression = { $_.NotAfter}} -First 1
        $c = "" | Select-Object Cert, Name, Expiration
        $c.Cert = "BSN Root CA"
        $c.Name = $cert.Subject
        $c.Expiration = $cert.'Expiration Date'
        $certs += $c
     
        $cert = Get-ChildItem | Where-Object { $_.subject -like "*DOW Chemical Root CA*"} | Select-Object Subject, @{Name = "Expiration Date"; Expression = { $_.NotAfter}} -First 1
        $c = "" | Select-Object Cert, Name, Expiration
        $c.Cert = "DOW Root CA"
        $c.Name = $cert.Subject
        $c.Expiration = $cert.'Expiration Date'
        $certs += $c

        $cert = Get-ChildItem | Where-Object { $_.subject -like "*DOW Corning PKI Production Root CA - 02*"} | Select-Object Subject, @{Name = "Expiration Date"; Expression = { $_.NotAfter}} -First 1
        $c = "" | Select-Object Cert, Name, Expiration
        $c.Cert = "DOW Corning PKI Root CA"
        $c.Name = $cert.Subject
        $c.Expiration = $cert.'Expiration Date'
        $certs += $c

        push-Location cert:\CurrentUser\CA
        $cert = Get-ChildItem | Where-Object { $_.subject -like "*DOW Chemical Production SSL CA*"} | Select-Object Subject, @{Name = "Expiration Date"; Expression = { $_.NotAfter}} -First 1
        $c = "" | Select-Object Cert, Name, Expiration
        $c.Cert = "DOW SSL CA"
        $c.Name = $cert.Subject
        $c.Expiration = $cert.'Expiration Date'
        $certs += $c
     
        $cert = Get-ChildItem | Where-Object { $_.subject -like "*BSNC Mach CA*"} | Select-Object Subject, @{Name = "Expiration Date"; Expression = { $_.NotAfter}} -First 1
        $c = "" | Select-Object Cert, Name, Expiration
        $c.Cert = "BSNC Mach CA"
        $c.Name = $cert.Subject
        $c.Expiration = $cert.'Expiration Date'
        $certs += $c
     
        $cert = Get-ChildItem | Where-Object { $_.subject -like "*DOW Corning Issuing CA 31*"} | Select-Object Subject, @{Name = "Expiration Date"; Expression = { $_.NotAfter}} -First 1
        $c = "" | Select-Object Cert, Name, Expiration
        $c.Cert = "DOW Corning CA 31"
        $c.Name = $cert.Subject
        $c.Expiration = $cert.'Expiration Date'
        $certs += $c
     
        $cert = Get-ChildItem | Where-Object { $_.subject -like "*DOW Corning Issuing CA 32*"} | Select-Object Subject, @{Name = "Expiration Date"; Expression = { $_.NotAfter}} -First 1
        $c = "" | Select-Object Cert, Name, Expiration
        $c.Cert = "DOW Corning CA 32"
        $c.Name = $cert.Subject
        $c.Expiration = $cert.'Expiration Date'
        $certs += $c
     
        $cert = Get-ChildItem | Where-Object { $_.subject -like "*DOW Chemical Issuing CA*"} | Select-Object Subject, @{Name = "Expiration Date"; Expression = { $_.NotAfter}} -First 1 
        $c = "" | Select-Object Cert, Name, Expiration
        $c.Cert = "DOW Issuing CA"
        $c.Name = $cert.Subject
        $c.Expiration = $cert.'Expiration Date'
        $certs += $c


    }
    End {
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
function Get-SEPInfo {
    [CmdletBinding()]
    
    Param()

    Begin {
    }
    Process {
        $RegPath = "hklm:\SOFTWARE\Wow6432Node"
        $SepInfo = "" | Select-Object Server, Group, Version, DefinitionDate
        $sylink = Get-ItemProperty "$regpath\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink"  -ErrorAction SilentlyContinue
       
        # If sylink is null then no data found.  Check for 32BIT version
        If ($sylink -eq $null) {
            $sylink = Get-ItemProperty "hklm:\SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SYLINK\SyLink"  -ErrorAction SilentlyContinue
            $RegPath = "hklm:\SOFTWARE"

        }
        if ($sylink -ne $null) {
            #Get AV Definition Date
            $av = Get-ItemProperty "$RegPath\Symantec\Symantec Endpoint Protection\AV" -ErrorAction SilentlyContinue
            $SepInfo.DefinitionDate = [datetime]"$($av.PatternFileDate[1]+1) - $($av.PatternFileDate[2]) - $($av.PatternFileDate[0]+1970)" | get-date -format "MM-dd-yyyy"
            $SEP = Get-ItemProperty "$RegPath\Symantec\Symantec Endpoint Protection\CurrentVersion"
            $SEPInfo.Version = $SEP.PRODUCTVERSION
            $sepinfo.Group = $sylink.CurrentGroup
            $Server = $sylink.CommunicationStatus.Split(':')[1]
            if ($Server -match "\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}") { 
                try { 
                    $Server = ([System.Net.Dns]::GetHostbyAddress($matches[0]) ).hostname 
                }
                catch {
                 
                }
            }
            $sepinfo.Server = $Server
        }
        else {
            $SepInfo.Server = "Not Installed"
        }
    }
    End {
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
function Get-WSUSInfo {
    [CmdletBinding()]
    Param()

    Begin {
    }
    Process { 
        $WSUS = "" | Select-Object Server
        $wsus.server = (Get-ItemProperty "hklm:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate").WUServer
    }
    End {
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
function Get-Bit9 {
    [CmdletBinding()]
    Param
    (
       
    )

    Begin {
    }
    Process {
        Try {
            $bit9 = & 'C:\Program Files (x86)\Bit9\Parity Agent\DasCLI.exe' status 
            $version = ($bit9 | Out-String) -match "(?ms)^Version.*?Cache"
            if ($matches.count -gt 0 ) {
                #write-host "BIT9 Installed" 
                $matches[0]
            }
        }
        Catch {
            #Write-host "BIT9 not installed"
            return "BIT9 not Installed"
        }

    }
    End {
    }
} 
    
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
        $SyncHash.$Object.Dispatcher.Invoke([System.Action] {Set-Variable -Name TempVar -Value $($SyncHash.$Object.$Property) -Scope global}, "Normal")
    }
    else {
        $SyncHash.$Object.Dispatcher.Invoke([System.Action] {Set-Variable -Name TempVar -Value $($SyncHash.$Object) -Scope global}, "Normal")
    }

    Return $TempVar
}


$Global:syncHash = [hashtable]::Synchronized(@{})
$initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$function = Get-Content function:\Get-BIT9
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Get-Bit9", $function
$initialSessionState.Commands.add($functionEntry)
$function = Get-Content function:\Get-CSADTask
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Get-CSADTask", $function
$initialSessionState.Commands.add($functionEntry)
$function = Get-Content function:\Get-DOWRootCA
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Get-DOWRootCA", $function
$initialSessionState.Commands.add($functionEntry)
$function = Get-Content function:\Get-SEPINFO
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Get-SEPInfo", $function
$initialSessionState.Commands.add($functionEntry)
$function = Get-Content function:\GET-WSUSInfo
$functionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "Get-WSUSInfo", $function
$initialSessionState.Commands.add($functionEntry)

$newRunspace = [runspacefactory]::CreateRunspace($initialSessionState)
$Global:syncHash.Path = (Get-Location).Path
$newRunspace.Name = "GUI"
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$syncHash.InitialSessionState = $initialSessionState
$newRunspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

$psCmd = [PowerShell]::Create().AddScript( {
  
        [xml]$xaml = @"
 <Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

  Title="Lab PC Check" Height="486.885" Width="782.377" ShowInTaskbar="False">

    <Grid>
        <Button x:Name="btnBIT9" Content="BIT9" HorizontalAlignment="Left" Margin="10,24,0,0" VerticalAlignment="Top" Width="75"/>
        <TextBlock x:Name="tbBIT9_1" HorizontalAlignment="Left" Margin="112,24,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="69" Width="273" FontSize="8"  Text=""/>
        <DataGrid x:Name="grdCSD" HorizontalAlignment="Left" Margin="112,100,0,0" VerticalAlignment="Top" FontSize="8" MinWidth="300">
           
        </DataGrid>
        <DataGrid x:Name="grdDOWCA" HorizontalAlignment="Left" Margin="112,138,0,0" VerticalAlignment="Top" FontSize="8" MaxHeight="145" MinHeight="50" MinWidth="300"/>
        <DataGrid x:Name="grdSepINFO" HorizontalAlignment="Left" Margin="112,345,0,0" VerticalAlignment="Top" FontSize="8" MinWidth="300"/>
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
        $inputXML = Get-Content -Path .\WpfWindow1.xaml
    
        # $reader=(New-Object System.Xml.XmlNodeReader $xaml)
        # $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
        #	$form=[Windows.Markup.XamlReader]::Load( $reader )

        $syncHash.Host = $Host

        [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
        [System.Reflection.Assembly]::LoadWithPartialName("WindowsFormsIntegration")
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")

        #$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
        #[xml]$XAML = $inputXML
        #[xml]$xaml 
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $syncHash.Window = [Windows.Markup.XamlReader]::Load( $reader )
        $form = [Windows.Markup.XamlReader]::Load( $reader )
        $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
            #Find all of the form types and add them as members to the synchash
            $syncHash.Add($_.Name, $syncHash.Window.FindName($_.Name) )
        }
        $xaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name "WPF$($_.Name)" -Value $form.FindName($_.Name)}
        $Script:JobCleanup = [hashtable]::Synchronized(@{})
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


        $syncHash.btnBIT9.Add_Click( {
		
                #region Boe's Additions
                $newRunspace = [runspacefactory]::CreateRunspace($synchash.initialSessionState)
		
                $newRunspace.ApartmentState = "STA"
                $newRunspace.Name = "BIT9"
                $newRunspace.ThreadOptions = "ReuseThread"
                $newRunspace.Open()
                $newRunspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
                $PowerShell = [PowerShell]::Create().AddScript( {
        
                        Try {
                            $bit9 = & 'C:\Program Files (x86)\Bit9\Parity Agent\DasCLI.exe' status 
                            $version = ($bit9 | Out-String) -match "(?ms)^Version.*?Cache"
                            if ($matches.count -gt 0 ) {
                                #write-host "BIT9 Installed" 
                                $Bit9 = $matches[0]
                            }
                        }
                        Catch {
                            #Write-host "BIT9 not installed"
                            $bit9 = "BIT9 not Installed"
                        }
     

	   
                        $syncHash.tbBIT9_1.Dispatcher.Invoke([action] {$syncHash.tbBIT9_1.text = $bit9.split("`n")[0..4]}, "Normal")
                        $syncHash.tbBIT9_1.Dispatcher.Invoke([action] { $syncHash.tbBIT9_2.text = $bit9.split("`n")[6..9]}, "Normal")

                    })
                $PowerShell.Runspace = $newRunspace
                [void]$Jobs.Add((
                        [pscustomobject]@{
                            PowerShell = $PowerShell

                            Runspace   = $PowerShell.BeginInvoke()
                        }
                    ))
            })

        $syncHash.btnCSD.Add_Click( {
                #region Boe's Additions
                $newRunspace = [runspacefactory]::CreateRunspace($synchash.initialSessionState)
                $newRunspace.ApartmentState = "STA"
                $newRunspace.Name = "CSD"
                $newRunspace.ThreadOptions = "ReuseThread"
                $newRunspace.Open()
                $newRunspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
                $PowerShell = [PowerShell]::Create().AddScript( {
                        $CSD = Get-CSADTask
                        $c = @()
                        $c += $csd
                        $syncHash.grdCSD.Dispatcher.Invoke([action] {$syncHash.grdCSD.ItemsSource = $c}, "Normal")
       

                    })
	
                $PowerShell.Runspace = $newRunspace
                [void]$Jobs.Add((
                        [pscustomobject]@{
                            PowerShell = $PowerShell

                            Runspace   = $PowerShell.BeginInvoke()
                        }
                    ))

    
            })

        $syncHash.btnDOWCA.Add_Click( {
                #region Boe's Additions
                $newRunspace = [runspacefactory]::CreateRunspace($synchash.initialSessionState)
                $newRunspace.ApartmentState = "STA"
                $newRunspace.Name = "CSD"
                $newRunspace.ThreadOptions = "ReuseThread"
                $newRunspace.Open()
                $newRunspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
                $PowerShell = [PowerShell]::Create().AddScript( {
                        $cert = Get-DOWRootCA
       
                        $syncHash.grdDOWCA.Dispatcher.Invoke([action] {$syncHash.grdDOWCA.ItemsSource = $cert}, "Normal")
       

                    })
	
                $PowerShell.Runspace = $newRunspace
                [void]$Jobs.Add((
                        [pscustomobject]@{
                            PowerShell = $PowerShell

                            Runspace   = $PowerShell.BeginInvoke()
                        }
                    ))

    
            })

        $syncHash.btnSEP.Add_Click( {
                #region Boe's Additions
                $newRunspace = [runspacefactory]::CreateRunspace($synchash.initialSessionState)
                $newRunspace.ApartmentState = "STA"
                $newRunspace.Name = "CSD"
                $newRunspace.ThreadOptions = "ReuseThread"
                $newRunspace.Open()
                $newRunspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
                $PowerShell = [PowerShell]::Create().AddScript( {
                        $sep = Get-SEPInfo
                        $sepinfo = @()
                        $sepinfo += $sep
       
                        $syncHash.grdSepINFO.Dispatcher.Invoke([action] {$syncHash.grdSepINFO.ItemsSource = $sepinfo}, "Normal")
       

                    })
	
                $PowerShell.Runspace = $newRunspace
                [void]$Jobs.Add((
                        [pscustomobject]@{
                            PowerShell = $PowerShell

                            Runspace   = $PowerShell.BeginInvoke()
                        }
                    ))

    
            })

        $syncHash.btnWSUS.Add_Click( {
                #region Boe's Additions
                $newRunspace = [runspacefactory]::CreateRunspace($synchash.initialSessionState)
                $newRunspace.ApartmentState = "STA"
                $newRunspace.Name = "WSUS"
                $newRunspace.ThreadOptions = "ReuseThread"
                $newRunspace.Open()
                $newRunspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
                $PowerShell = [PowerShell]::Create().AddScript( {
                        $wsus = Get-WSUSInfo
                        $syncHash.tbwsus.Dispatcher.Invoke([action] {$syncHash.tbwsus.text = $wsus.server}, "Normal")
       

                    })
	
                $PowerShell.Runspace = $newRunspace
                [void]$Jobs.Add((
                        [pscustomobject]@{
                            PowerShell = $PowerShell

                            Runspace   = $PowerShell.BeginInvoke()
                        }
                    ))

    
            })
        #>
        #region Window Close
        $synchash.tbBit9_1.Text = get-location
        $syncHash.Window.Add_Closed( {
                Write-Verbose 'Halt runspace cleanup job processing'
                $jobCleanup.Flag = $False

                #Stop all runspaces
                $jobCleanup.PowerShell.Dispose()
            })
        #endregion Window Close
        #endregion Boe's Additions

        #$x.Host.Runspace.Events.GenerateEvent( "TestClicked", $x.test, $null, "test event")
        $syncHash.btnBIT9.RaiseEvent((new-object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
        $syncHash.btnCSD.RaiseEvent((new-object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
        $syncHash.btnDOWCA.RaiseEvent((new-object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
        $syncHash.btnSEP.RaiseEvent((new-object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
        $syncHash.btnWSUS.RaiseEvent((new-object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
        #$syncHash.Window.Activate()
        $syncHash.Window.ShowDialog() | Out-Null
        $syncHash.Error = $Error












    })
$psCmd.Runspace = $newRunspace

1..5 | ForEach-Object { write-host "!!!!!!!!!!!!!!!!!!"}
Write-host "to run display your UI, run:  " -NoNewline
write-host -foregroundcolor Green '$data = $psCmd.BeginInvoke()'
$data = $psCmd.BeginInvoke()
Start-Sleep -Milliseconds 500
#$syncHash.txtCurrDir.Dispatcher.Invoke([action]{ $syncHash.txtCurrDir.text = $syncHash.path},"Normal")
#$syncHash.IPResults.Dispatcher.Invoke([action]{$syncHash.IPResults.ItemsSource = $ips},"Normal")
#$syncHash.Window.Dispatcher.Invoke([action]{$syncHash.IPS.Clear()},"Normal")
# How to add to Background UI
#$syncHash.Window.Dispatcher.Invoke([action]{$ips.Add($ip)},"Normal")
<#
$syncHash.txtInput.Add_LostFocus({
	$syncHash.txtinput.dispatcher.invoke([action]{$global:test = $syncHash.txtinput.text})
})
#>


function close-OrphanedRunSpaces() {
    Get-Runspace
    Write-Host "closing"
    Get-Runspace | Where-Object { $_.RunspaceAvailability -eq "Available"} | ForEach-Object { $_.close(); $_.Dispose()}
    write-host "Closed"
    Get-Runspace
}




<#
###  Examples
	$syncHash.txtInput.Add_LostFocus({
	if ($syncHash.txtInput.Text -like "*\*")
		{
			$file = $syncHash.txtInput.Text
		}
		else
		{
			$file = "$($syncHash.txtcurrdir.text)\$($synchash.txtinput.text)"
		}
	if (Test-Path  $file)
     { $syncHash.btnStart.Dispatcher.Invoke([action]{$syncHash.btnStart.IsEnabled = $True},"Normal")
     }
     else
     { $syncHash.btnStart.Dispatcher.Invoke([action]{$syncHash.btnStart.IsEnabled = $false},"Normal")}
})

$syncHash.btnExport.add_Click({
    if ($synchash.txtOutput.Text.Trim().Length -eq 0) 
    {
        "blank" | Out-File "C:\Users\nk23208\Source\Repos\PowerShell-Scripts\NSLookup-WPF\NSLookup-WPF\test.txt"
        $file = "$($synchash.txtCurrDir.text)\$($synchash.txtInput.Text)"
        $file = (gci $file).BaseName
        $synchash.txtOutput.Text = "$file-$(get-date -f "MM-dd-yyyy").csv"
    } 
	if ($syncHash.txtOutput.Text -notlike "*.csv")
	{ 
        $syncHash.txtOutput.Text += ".csv"
    }
	$syncHash.IPS | Export-Csv "$($syncHash.txtcurrdir.text)\$($synchash.txtoutput.text)" -NoTypeInformation
})

    $syncHash.btnStart.Add_Click({
		$syncHash.btnStart.IsEnabled = $false
		if ($syncHash.txtInput.Text -like "*\*")
		{
			$file = $syncHash.txtInput.Text
		}
		else
		{
			$file = "$($syncHash.txtcurrdir.text)\$($synchash.txtinput.text)"
		}

		$syncHash.IPS.clear()
		 #region Boe's Additions
        $newRunspace =[runspacefactory]::CreateRunspace()
		$syncHash.File = Get-content $file | select -Unique
        $newRunspace.ApartmentState = "STA"
		$newRunspace.Name = "DNSQuery"
        $newRunspace.ThreadOptions = "ReuseThread"
        $newRunspace.Open()
        $newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash)
        $PowerShell = [PowerShell]::Create().AddScript({
	    $ServerCount = $synchash.file.count
        $i= 0
$synchash.file | % {
	$i = $synchash.file.IndexOf($_) +1
	$p = [math]::round($i / $ServerCount * 100)
	$synchash.lblProgress.Dispatcher.Invoke([action]{$synchash.lblProgress.Content = "$i/$servercount  $p%"},"Normal")
	$synchash.Progress.Dispatcher.Invoke([action]{$syncHash.Progress.Value = $p},"Normal")
	
    #$syncHash.Progress.Value = 
	remove-variable R -ErrorAction SilentlyContinue | out-null

   if ($_ -as [ipaddress] )
   {
    Try
    {
    #Write-host "By IP Address $_"
    $name = $_
    $R =  [System.Net.Dns]::GetHostbyAddress($_)
    $IP = New-Object psobject
    $IP | Add-Member -Type NoteProperty -Name Target -Value $_
    $IP | Add-Member -Type NoteProperty -Name HostName -Value $R.HostName.ToUpper()
    $IP | Add-Member -Type NoteProperty -Name IPAddress -Value $R.AddressList.IPAddressToString
    $syncHash.Window.Dispatcher.Invoke([action]{$synchash.ips.Add($ip)},"Normal")
    }
    Catch
    {
    #Write-host "IP Address Exception $Name"
    $IP = New-Object psobject
    $IP | Add-Member -Type NoteProperty -Name Target -value $name 
    $IP | Add-Member -Type NoteProperty -Name HostName -Value "Not Found"
    $IP | Add-Member -Type NoteProperty -Name IPAddress -Value $name
   $syncHash.Window.Dispatcher.Invoke([action]{$synchash.ips.Add($ip)},"Normal")
    }
   }
   else
   {
      Try
       {
       #Write-host "By HostName $_"
       $name = $_.toUpper()
       $R = [System.Net.Dns]::GetHostAddresses($_)
       foreach ($i  in $R )
       {
       $IP = New-Object psobject
       $IP | Add-Member -Type NoteProperty -Name Target  -value $_
       $IP | Add-Member -type NoteProperty -name HostName -value $_.toUpper()
       $IP | Add-Member -Type NoteProperty -Name IPAddress -value $i.IPAddressToString
       $syncHash.Window.Dispatcher.Invoke([action]{$synchash.ips.Add($ip)},"Normal")
		  # $synchash.ips.add($ip)
       }
       }
       catch
       {
          #Write-host "HostName Exeption $name"
          $IP = New-Object psobject
          $IP | Add-Member -Type NoteProperty -Name Target -Value $name
          $IP | Add-Member -Type NoteProperty -Name HostName -Value $name
          $IP | Add-Member -Type NoteProperty -Name IPAddress -Value "Not Found"
         $syncHash.Window.Dispatcher.Invoke([action]{$synchash.ips.Add($ip)},"Normal")
       }
   }

    remove-variable R -ErrorAction SilentlyContinue | out-null
    }
		$synchash.btnExport.Dispatcher.Invoke([action]{$synchash.btnExport.IsEnabled = $true},"Normal")
})
		$SyncHash.Host.UI.Write( "button")
        #Start-Job -Name Sleeping -ScriptBlock {start-sleep 5}
        #while ((Get-Job Sleeping).State -eq 'Running'){
            $x+= "."
        $PowerShell.Runspace = $newRunspace
        [void]$Jobs.Add((
            [pscustomobject]@{
                PowerShell = $PowerShell

                Runspace = $PowerShell.BeginInvoke()
            }
        ))

    
    })
#>