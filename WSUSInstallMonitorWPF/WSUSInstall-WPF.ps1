Get-EventSubscriber | ForEach-Object { Unregister-Event $_.SubscriptionId}

$computers = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$stat = New-Object psobject
$stat | Add-Member -Type NoteProperty -Name Computer -Value "test1"
$stat | Add-Member  -type NoteProperty -Name Action -Value "Install"
$stat | Add-Member -type NoteProperty -Name Time -Value "12:00:00"
$stat | Add-Member -Type NoteProperty -Name Description -Value "test install"
$stat | Add-Member -type NoteProperty -Name Progress -Value $(Get-Random -Maximum 100)
$computers.add($stat)

$Global:syncHash = [hashtable]::Synchronized(@{})
$syncHash.add("Host", $Host)
$newRunspace = [runspacefactory]::CreateRunspace()
$syncHash.path = $PWD.path
$syncHash.computers = $computers
$syncHash.stat = $stat
$newRunspace.Name = "GUI"
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

$psCmd = [PowerShell]::Create().AddScript( {
        [xml]$xaml = @"
 <Window
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="WSUS Install Monitor" Height="456.329" Width="894.903" MinWidth="1000" MinHeight="432">

     <Grid Margin="0,0,0,0">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="4*" />
            <ColumnDefinition />

        </Grid.ColumnDefinitions>
        <Button x:Name="btnStart" Content="Start" HorizontalAlignment="Right" Margin="0,10,120,0" VerticalAlignment="Top" Width="75" Grid.Column="1" IsEnabled="False"/>
        <Button x:Name="btnStop" Content="Stop" HorizontalAlignment="Right" Margin="0,10,40,0" VerticalAlignment="Top" Width="75" Grid.Column="1" IsEnabled="True"/>
        
        <Label Content="Data Folder" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="txtDataFolder" HorizontalAlignment="Left" Height="23" Margin="100,10,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="572" Grid.ColumnSpan="2"/> 
         <DataGrid x:Name="WSUSResults" Grid.Column="0"   Margin="10,50,10,10" AutoGenerateColumns="False" Grid.ColumnSpan="2" MinHeight="300" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" UseLayoutRounding="True">
            <DataGrid.RowStyle>
                <Style TargetType ="DataGridRow">
                    <Style.Triggers>
                        <DataTrigger Binding="{Binding Action}" Value="Reboot">
                            <Setter Property="Background" Value="Violet" />
                            <Setter Property="Foreground" Value="Yellow"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Action}" Value="Install">
                            <Setter Property="Background" Value="LawnGreen" />
                            <Setter Property="Foreground" Value="Black"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Action}" Value="Download">
                            <Setter Property="Background" Value="OrangeRed" />
                            <Setter Property="Foreground" Value="Yellow"/>
                        </DataTrigger>
                       <DataTrigger Binding="{Binding Action}" Value="Catalog">
                            <Setter Property="Background" Value="OrangeRed" />
                            <Setter Property="Foreground" Value="Yellow"/>
                        </DataTrigger>
                        <DataTrigger Binding="{Binding Action}" Value="Search">
                            <Setter Property="Background" Value="Yellow" />
                            <Setter Property="Foreground" Value="Black"/>
                        </DataTrigger>
                    </Style.Triggers>
                </Style>
               
            </DataGrid.RowStyle>
            <DataGrid.Resources>
                <DataTemplate x:Key="MyDataTemplate">
                    <Grid >
                        <ProgressBar x:Name="ProgCell" Minimum="0" Maximum="100" FlowDirection="LeftToRight"  Value="{Binding Progress}" Width="{Binding Path=Width, ElementName=ProgressCell}" Height="{Binding Path=Height, ElementName=ProgressCell}" Margin="0"/>
                        <TextBlock Text="{Binding Progress, StringFormat={}{0}%}" Width="{Binding Path=Width, ElementName=ProgressCell}" Height="{Binding Path=Height, ElementName=ProgressCell}"  HorizontalAlignment="Center"/>
                    </Grid>
                </DataTemplate>
            </DataGrid.Resources>
            <DataGrid.GroupStyle>
                <GroupStyle>
                <GroupStyle.HeaderTemplate>
                        <DataTemplate>
                            <TextBlock FontWeight="Bold" FontSize="14" Text="{Binding Path=Action}"/>
                        </DataTemplate>
                    </GroupStyle.HeaderTemplate>
                </GroupStyle>
            </DataGrid.GroupStyle>
            <DataGrid.Columns>
                <DataGridTextColumn Binding="{Binding Computer}" Width="120" ClipboardContentBinding="{x:Null}" Header="Computer" />
                <DataGridTextColumn Binding="{Binding Action}" Width="100" ClipboardContentBinding="{x:Null}" Header="Action"/>
                <DataGridTextColumn Binding="{Binding Time}"  Width="130" ClipboardContentBinding="{x:Null}" Header="Time"/>
                <DataGridTemplateColumn x:Name="ProgressCell" ClipboardContentBinding="{x:Null}" Header="Progress" Width="200" CellTemplate="{StaticResource MyDataTemplate}"/>
                <DataGridTextColumn Binding="{Binding Description}"  Width="*" MinWidth="400" ClipboardContentBinding="{x:Null}" Header="Status"/>
            </DataGrid.Columns>
        </DataGrid>

    </Grid>
</Window>
"@
        [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
        [System.Reflection.Assembly]::LoadWithPartialName("WindowsFormsIntegration")
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $syncHash.Window = [Windows.Markup.XamlReader]::Load( $reader )
        #$form=[Windows.Markup.XamlReader]::Load( $reader )

        

    

        #$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
        [xml]$XAML = $xaml
        $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | % {
            #Find all of the form types and add them as members to the synchash
            $syncHash.Add($_.Name, $syncHash.Window.FindName($_.Name) )
        }
        $Script:JobCleanup = [hashtable]::Synchronized(@{})
        $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

        #region Background runspace to clean up jobs
        $jobCleanup.Flag = $True
        $newRunspace = [runspacefactory]::CreateRunspace()
        $newRunspace.Name = "Cleanup"
        $newRunspace.ApartmentState = "STA"
        $newRunspace.ThreadOptions = "ReuseThread"
        $newRunspace.Open()
        $syncHash.jobs = $Script:Jobs
        $newRunspace.SessionStateProxy.SetVariable("jobCleanup", $jobCleanup)
        $newRunspace.SessionStateProxy.SetVariable("jobs", $script:jobs)
        $newRunspace.SessionStateProxy.SetVariable("synchash", $synchash)
        $jobCleanup.PowerShell = [PowerShell]::Create().AddScript( {
                #Routine to handle completed runspaces
                Do
                {
                    Foreach ($runspace in $jobs)
                    {
                        If ($runspace.Runspace.isCompleted)
                        {
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


        $syncHash.txtDataFolder.Add_LostFocus( {
               
                if (Test-Path  $syncHash.txtDataFolder.text)
                {
                    $syncHash.btnStart.Dispatcher.Invoke([action] {$syncHash.btnStart.IsEnabled = $True}, "Normal")
                }
                else
                { $syncHash.btnStart.Dispatcher.Invoke([action] {$syncHash.btnStart.IsEnabled = $false}, "Normal")}
            })

        $syncHash.btnStart.Add_Click( {
                $Global:timer = new-object System.Windows.Threading.DispatcherTimer
                #Fire off every 5 seconds
                
                Write-Verbose “Adding 1 second interval to timer object”
                $timer.Interval = [TimeSpan]"0:0:1.00"
                #Add event per tick
                Write-Verbose "Adding Tick Event to timer object"
                $global:timer.Add_Tick( {
                        $syncHash.wsusresults.Dispatcher.Invoke([action] { $syncHash.wsusresults.items.refresh()}, "Normal")
                    })
                #Start timer
                Write-Verbose “Starting Timer”
                $timer.Start()
                #$syncHash.IPS.clear()
                #region Boe's Additions
                $newRunspace = [runspacefactory]::CreateRunspace()
                $SyncHash.path = $SyncHash.txtDataFolder.text
                $newRunspace.ApartmentState = "STA"
                $newRunspace.Name = "WSUSResults"
                $newRunspace.ThreadOptions = "ReuseThread"
                $newRunspace.Open()
                $newRunspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
                $PowerShell = [PowerShell]::Create().AddScript( {
                        Param($computers)    
                        Get-EventSubscriber | ForEach-Object { Unregister-Event $_.SubscriptionId}
                        #** Debug $synchash.host.ui.Writeline("Click Runspace")
                        #** Debug $synchash.host.ui.Writeline("Runspace PID $PID")
                        #** Debug $synchash.host.ui.WriteLine("Synced: `n $($synchash.computers.IsSynchronized)")
                        $syncHash.Window.Dispatcher.Invoke([action] {$syncHash.computers.Add($syncHash.stat)}, "Normal")
                        function Test-dipatch
                        {
                            $syncHash.Host.ui.WriteLine("Test")
                            $syncHash.stat.Computer = "TestInstall"
                            $syncHash.Window.Dispatcher.Invoke([action] {$syncHash.computers.add($syncHash.stat)}, "Normal")
                        }
                        Test-dipatch
                        function Get-LastLine
                        {
                            [cmdletBinding()]
                            Param($path)
                            Wait-Debugger
                            $synchash.host.ui.WriteLine("Get-LastLine")
                            #$syncHash.Window.Dispatcher.Invoke([action] {$syncHash.computers.Add($syncHash.stat)}, "Normal")

                            $synchash.host.ui.WriteLine("Enter Invoke")    
                            $lines = Get-Content $path
                            $lines = $lines.split("`n")
                                    
                            if ( $Lines[$lines.count - 1].Trim().Length -gt 0 ) 
                            { $line = $lines[$line.count - 1] }
                            else 
                            { $line = $lines[$lines.count - 2] }
                            $line = $line.Split(';')
                            $stat = New-Object psobject
                            $stat | Add-Member -Type NoteProperty -Name Computer -Value $line[0] 
                            $stat | Add-Member  -type NoteProperty -Name Action -Value $line[1]
                            $stat | Add-Member -type NoteProperty -Name Time -Value $line[2]
                            $stat | Add-Member -Type NoteProperty -Name Description -Value $line[3]
                            #** Debug $synchash.host.ui.writeline("Lines : $line[3]")
                            $progress = 0
                            if ($line[3] -ilike "*Total Progress*")
                            {
                                $progress = $line[3].Substring($line[3].IndexOf("Total Progress"))
                               
                                #** Debug $synchash.host.ui.writeline("Progress : $progress")
                                if ($line[3].Substring($line[3].IndexOf("Total Progress")) -match "\d{1,3}")
                                {
                                    $progress = $Matches[0]
                                }
                                else
                                {
                                    $progress = 0
                                }
                            }
                          
                            $stat | Add-Member -type NoteProperty -Name Progress -Value $progress
                            #** Debug $synchash.host.ui.WriteLine("stat : $stat")
                            #write-host "Count $($comp.count)"
                            #** Debug $synchash.host.ui.WriteLine(" $PID Computer Count : $($synchash.computers.count)")
                            #$synchash.Window.dispatcher.invoke([action] {$synchash.computers.add($stat)}, "Normal")
                            
                            if ($synchash.computers.count -eq 0 )
                            {
                                #Write-host "Zero"
                                #** Debug $synchash.host.ui.WriteLine("$PID Computer Count :initializing")
                                
                                try
                                {
                                    # $synchash.computers.add($stat)
                                    $global:synchash.computers.add($stat)
                                }
                                catch
                                {
                                    #** Debug $synchash.host.ui.WriteLine("$PID Init Error: $_")
                                }
                              
                                #** Debug $synchash.host.ui.WriteLine("$PID after Init Computer Count : $($synchash.computers.count)")
                            }
                            
                            else
                            {
                                if ($synchash.computers.computer.Contains($stat.computer))
                                {
                                    #** Debug $synchash.host.ui.WriteLine("$PID Found Computer")
                                    $index = $synchash.computers.computer.IndexOf($stat.computer)
                                    try
                                    {
                                        #** Debug $synchash.host.ui.WriteLine("Updating Computer")
                                        #** Debug $synchash.host.ui.WriteLine("Index :$index")
                                        
                                        $synchash.computers[$index] = $stat
                                            
                                    }
                                    catch
                                    {
                                        #** Debug $synchash.host.ui.WriteLine("Update Error: $_")
                                    }
                                
                                
                                }
                                else
                                {  
                                    #** Debug $synchash.host.ui.WriteLine("Adding new computer")
                                    try
                                    {
                                        $synchash.computers.add($stat)
                                    }
                                    catch
                                    {
                                        #** Debug $synchash.host.ui.WriteLine("add Error : $_")
                                    }
                                
          
                                } 
    
                            }
              
                            # $Synchash.Host.Runspace.Events.GenerateEvent("ListViewChanged", $syncHash.listView, $null, "ListView Changed")
                            #   Register-EngineEvent -SourceIdentifier "ListViewChanged" -Action {$synchash.host.ui.Writeline("Event Happened inside")} -Forward
                            #   $test = "na"
                            #  if ($syncHash.listView.items.count -gt 0) {write-host "test 321" ; $test = "test 321"; $syncHash.listView.Items.Refresh()} else {$test = "nope" ; $syncHash.listView.ItemsSource = $computers}
                        }
                       
                        #** debug $syncHash.host.ui.Writeline("Register Events")
                        #** debug $syncHash.host.ui.writeline("Path : $($Synchash.path)")
                        # Register-EngineEvent -SourceIdentifier "ListViewChanged" -Action {$synchash.host.ui.Writeline("Event Happened outside")}
                        $fsw = New-Object System.IO.FileSystemWatcher $syncHash.path, "*.csv" 
                        $event = Register-ObjectEvent -InputObject $fsw -EventName "Changed" -action {
                            $syncHash.host.ui.writeline("PID $PID"); 
                            Get-LastLine($event.sourceEventArgs.fullpath)}       
                           
                    }).addargument($synchash.computers)
              
                #Start-Job -Name Sleeping -ScriptBlock {start-sleep 5}
                #while ((Get-Job Sleeping).State -eq 'Running'){
                
                $PowerShell.Runspace = $newRunspace
                [void]$Jobs.Add((
                        [pscustomobject]@{
                            PowerShell = $PowerShell

                            Runspace   = $PowerShell.BeginInvoke()
                        }
                    ))
                $SyncHash.Host.ui.WriteLine("Jobs : $($jobs.count)")

    
            })
       

        #region Window Close
        $syncHash.btnStop.add_Click( {
                Get-EventSubscriber | ForEach-Object { Unregister-Event $_.SubscriptionId}
                if ($timer.IsEnabled)
                {
                    $timer.stop()
                }
            })
        $syncHash.Window.Add_Closed( {
                Write-Verbose 'Halt runspace cleanup job processing'
                $SyncHash.Host.UI.Writeline( "Windows CLose")
                $SyncHash.Host.UI.Writeline( "CLeanupflag $($jobCleanup.Flag)")
                $jobCleanup.Flag = $False

                #Stop all runspaces
                $jobCleanup.PowerShell.Dispose()
            })
        #endregion Window Close
        #endregion Boe's Additions

        #$x.Host.Runspace.Events.GenerateEvent( "TestClicked", $x.test, $null, "test event")

        #$syncHash.Window.Activate()
        # $Binding = New-Object System.Windows.Data.Binding 
        #$Binding.Mode = [System.Windows.Data.BindingMode]::OneWay
        #$syncHash.WSUSResults.DataContext = $computers
        #[void][System.Windows.Data.BindingOperations]::SetBinding($syncHash.WSUSResults, [System.Windows.Controls.DataGrid]::ItemsSourceProperty, $Binding)

        $syncHash.Window.ShowDialog() | Out-Null
        $syncHash.Error = $Error
    })
$psCmd.Runspace = $newRunspace

#Add-Type -AssemblyName System.windows.data.Binding
#[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Data.Binding')
#$Binding = New-Object System.Windows.Data.Binding 
#$Binding.Mode = [System.Windows.Data.BindingMode]::OneWay
#$syncHash.WSUSResults.DataContext = $computers
#[void][System.Windows.Data.BindingOperations]::SetBinding($syncHash.WSUSResults, [System.Windows.Controls.DataGrid]::ItemsSourceProperty, $Binding)
#[System.Windows.Data.BindingOperations]::
1..5 | % { write-host "!!!!!!!!!!!!!!!!!!"}
Write-host "to run display your UI, run:  " -NoNewline
write-host -foregroundcolor Green '$data = $psCmd.BeginInvoke()'
$data = $psCmd.BeginInvoke()
Start-Sleep -Milliseconds 1000
$synchash.Window.Dispatcher.Invoke([action] {$synchash.txtDataFolder.text = $PSScriptRoot}, "Normal")
#$syncHash.txtCurrDir.Dispatcher.Invoke([action] { $syncHash.txtCurrDir.text = $syncHash.path}, "Normal")
$syncHash.Window.Dispatcher.Invoke([action] {$syncHash.WSUSResults.ItemsSource = $computers}, "Normal")
Start-Sleep -Milliseconds 500
$syncHash.Window.Dispatcher.Invoke([action] {$syncHash.Computers.Clear()}, "Normal")
# How to add to Background UI
#$syncHash.Window.Dispatcher.Invoke([action]{$ips.Add($ip)},"Normal")
<#
$syncHash.txtInput.Add_LostFocus({
	$syncHash.txtinput.dispatcher.invoke([action]{$global:test = $syncHash.txtinput.text})
})
#>



function Get-SyncHashValue
{
    [CmdletBinding()]
    param (
        #  [parameter(Mandatory=$true)]
        #  $SyncHash,
        [parameter(Mandatory = $true)]
        $Object,
        [parameter(Mandatory = $false)]
        $Property
    )

    if ($TempVar)
    {
        Remove-Variable TempVar -Scope global
    }

    if ($Property)
    {

        $SyncHash.$Object.Dispatcher.Invoke([System.Action] {Set-Variable -Name TempVar -Value $($SyncHash.$Object.$Property) -Scope global}, "Normal")
    }
    else
    {
        $SyncHash.Host.UI.writeline("Object")
        $SyncHash.$Object.Dispatcher.Invoke([System.Action] {Set-Variable -Name TempVar -Value $($SyncHash.$Object) -Scope global}, "Normal")
    }
    $SyncHash.Host.UI.writeline("com : $tempvar")
    Return $TempVar
}

function Get-LastLine
{
    [cmdletBinding()]
    Param($path)
   
    #$oldConsole = [console]::TreatControlCAsInput
    #[console]::TreatControlCAsInput = $true
    #write-host "enter"
    #write-host "computers : $($global:computers -isnot [System.Array])"
    $synchash.Host.ui.WriteLine("Get-LastLine")
    # $synchash.Host.ui.d
    write-debug "entrance"
    write-verbose "Entering LastLine"
   
    # $synchash.window.Dispatcher.invoke([action] {
    $synchash.Host.ui.WriteLine("Enter Invoke")    
    #$stat = "" | Select-Object Computer, Action, Time, Progress, Description
    $lines = Get-Content $path
    $lines = $lines.split("`n")
                                    
    if ( $Lines[$lines.count - 1].Trim().Length -gt 0 ) 
    { $line = $lines[$line.count - 1] }
    else 
    { $line = $lines[$lines.count - 2] }
    $line = $line.Split(';')
    $stat = New-Object psobject
    $stat | Add-Member -Type NoteProperty -Name Computer -Value $line[0] 
    $stat | Add-Member  -type NoteProperty -Name Action -Value $line[1]
    $stat | Add-Member -type NoteProperty -Name Time -Value $line[2]
    $stat | Add-Member -Type NoteProperty -Name Description -Value $line[3]
    $stat | Add-Member -type NoteProperty -Name Progress -Value $(Get-Random -Maximum 100)
    
    #write-host "Count $($comp.count)"
    $synchash.Host.ui.WriteLine(" $PID Computer Count : $($synchash.computers.count)")
    #$synchash.Window.dispatcher.invoke([action] {$synchash.computers.add($stat)}, "Normal")
                            
    if ($synchash.computers.count -eq 0 )
    {
        #Write-host "Zero"
        $synchash.Host.ui.WriteLine("$PID Computer Count :initializing")
                                
        try
        {
            # $synchash.computers.add($stat)
            $global:synchash.computers.add($stat)
        }
        catch
        {
            $synchash.Host.ui.WriteLine("$PID Init Error: $_")
        }
                              
        $synchash.Host.ui.WriteLine("$PID after Init Computer Count : $($synchash.computers.count)")
    }
                            
    else
    {
        if ($synchash.computers.computer.Contains($stat.computer))
        {
            $synchash.Host.ui.WriteLine("$PID Found Computer")
            $index = $synchash.computers.computer.IndexOf($stat.computer)
            try
            {
                $synchash.Host.ui.WriteLine("Updating Computer")
                $synchash.Host.ui.WriteLine("Index :$index")
                                        
                $synchash.computers[$index] = $stat
                                            
            }
            catch
            {
                $synchash.Host.ui.WriteLine("Update Error: $_")
            }
                                
                                
        }
        else
        {  
            $synchash.Host.ui.WriteLine("Adding new computer")
            try
            {
                $synchash.computers.add($stat)
            }
            catch
            {
                $synchash.Host.ui.WriteLine("add Error : $_")
            }
                                
          
        } 
    
    }
    #$synchash.WSUSResults.Dispatcher.Invoke([action] {$synchash.WSUSResults.items.Refresh()}, "Normal")       
}

#$fsw = New-Object System.IO.FileSystemWatcher "C:\Users\Talonx\source\repos\VSCode\PowerShell-Scripts\WSUSInstallMonitorWPF\Test Data", "*.csv" 
#$event = Register-ObjectEvent -InputObject $fsw -EventName "Changed" -action { Get-LastLine($event.sourceEventArgs.fullpath); $syncHash.host.ui.Writeline("Change")}       
                          
Register-EngineEvent -SourceIdentifier "ListViewChanged" -Action {$synchash.host.ui.Writeline("Event Happened outside"); 
    $synchash.host.UI.writeline(" WSUSResults1: ")  
    
    try
    {
        #  $global:synchash.Window.dispatcher.invoke([action] {$global:synchash.WSUSResults.items.Refresh() }, "Normal")
    }
    catch
    {
        $synchash.host.UI.writeline("Event Error: $_") 
    }  
    
}

function close-OrphanedRunSpaces()
{
    Get-Runspace
    Write-Host "closing"
    Get-Runspace | ? { $_.RunspaceAvailability -eq "Available"} | % { $_.close(); $_.Dispose()}
    write-host "Closed"
    Get-Runspace
}

<#
$status = @("Search", "Catalog", "Download", "Install", "Reboot")
1..100 | % {
    $s = "" | Select Computer, Action, Time,  Description
    $s.Computer = "Computer-$(get-random -Maximum 25)"
    $s.Action = $status[$(get-random -Maximum 4)]
    $s.Time = get-date -f "hh:mm:ss"
    $Progress = Get-Random -Maximum 100
    $s.Description = "Test - $_ Total Progress($progress%)"
    ($s |  ConvertTo-Csv -NoTypeInformation -Delimiter ";"   | % { $_.replace("""", '').replace(",", ";").replace("Computer;Action;Time;Progress;Description `n", "")})[1] | Out-File ".\test data\test.csv" -Append  
    Start-Sleep -Milliseconds (Get-Random -Maximum 1000)

} #>