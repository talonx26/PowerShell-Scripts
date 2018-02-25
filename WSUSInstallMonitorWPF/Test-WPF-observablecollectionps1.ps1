Get-EventSubscriber | % { Unregister-Event $_.SubscriptionId}
$DebugPreference = "Continue"
$Global:syncHash = [hashtable]::Synchronized(@{})
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open() 
$newRunspace.Name = "SyncHash"
$newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
$newRunspace.SessionStateProxy.SetVariable("Computers",$Computers)
$syncHash.add("Host",$Host)

$psCmd = [PowerShell]::Create().AddScript({
    [xml]$xaml = @"
  <Window
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="WSUS Install Monitor" Height="456.329" Width="1131.962">

    <Grid Margin="123,-55,3.4,54.2">
        <Grid.ColumnDefinitions>
        <ColumnDefinition Width="4*" />
        <ColumnDefinition />

     </Grid.ColumnDefinitions>
        <Button x:Name="btnStart" Content="Start" HorizontalAlignment="Left" Margin="861,402,0,0" VerticalAlignment="Top" Width="75"/>
        <ListView Grid.Column="0" x:Name="listView" HorizontalAlignment="Stretch" Height="350" Margin="-101,81,0,0" VerticalAlignment="Stretch" Width="Auto" >
            <ListView.Resources>
                <Style TargetType="{x:Type ListViewItem}">
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
                        <DataTrigger Binding="{Binding Action}" Value="Search">
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
                <DataTemplate x:Key="MyDataTemplate">
                    <Grid Margin="-6">
                        <ProgressBar x:Name="ProgCell" Minimum="0" Maximum="100" FlowDirection="LeftToRight"  Value="{Binding [Progress]}" Width="{Binding Path=Width, ElementName=ProgressCell}" Height="20" Margin="0"/>
                        <TextBlock Text="{Binding Progress, StringFormat={}{0}%}" HorizontalAlignment="Center"/>
                    </Grid>
                </DataTemplate>
            </ListView.Resources>
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Computer" Width="120"  DisplayMemberBinding="{Binding Computer}"/>
                    <GridViewColumn Header="Action" Width="100" DisplayMemberBinding="{Binding Action}"/>
                    <GridViewColumn Header="Time" Width="130" DisplayMemberBinding="{Binding Time}"/>
                  <GridViewColumn x:Name="ProgressCell" Header="Progress" Width="200" CellTemplate="{StaticResource MyDataTemplate}"/>

                    <GridViewColumn Header="Status" Width="400" DisplayMemberBinding="{Binding Description}"/>
                </GridView>
            </ListView.View>
           <ListView.GroupStyle>
                <GroupStyle>
                    <GroupStyle.HeaderTemplate>
                        <DataTemplate>
                            <TextBlock FontWeight="Bold" FontSize="14" Text="{Binding Path=Action}"/>
                        </DataTemplate>
                    </GroupStyle.HeaderTemplate>
                </GroupStyle>
            </ListView.GroupStyle>

        </ListView>
       
    </Grid>
</Window>

"@
#<TextBlock Text="{Binding Progress, StringFormat={}{0}%}" HorizontalAlignment="Center"/>
 #<GridViewColumn x:Name="ProgressCell" Header="Progress" Width="200" CellTemplate="{StaticResource MyDataTemplate}"/>

    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
    
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $xaml
        $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
        #Find all of the form types and add them as members to the synchash
        $syncHash.Add($_.Name,$syncHash.Window.FindName($_.Name) )
        $syncHash.add("Computers",$Computers)

    }

    $Script:JobCleanup = [hashtable]::Synchronized(@{})
    $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

    #region Background runspace to clean up jobs
    $jobCleanup.Flag = $True
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"          
    $newRunspace.Open()  
    $newRunspace.name = "Cleanup"      
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

    $syncHash.button.Add_Click({
        #Start-Job -Name Sleeping -ScriptBlock {start-sleep 5}
        #while ((Get-Job Sleeping).State -eq 'Running'){
            $x+= "."
        #region Boe's Additions
        $newRunspace =[runspacefactory]::CreateRunspace()
        $newRunspace.ApartmentState = "STA"
        $newRunspace.ThreadOptions = "ReuseThread"          
        $newRunspace.Open()
        $newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash) 
        $PowerShell = [PowerShell]::Create().AddScript({
Function Update-Window {
        Param (
            $Control,
            $Property,
            $Value,
            [switch]$AppendContent
        )

        # This is kind of a hack, there may be a better way to do this
        If ($Property -eq "Close") {
            $syncHash.Window.Dispatcher.invoke([action]{$syncHash.Window.Close()},"Normal")
            Return
        }

        # This updates the control based on the parameters passed to the function
        $syncHash.$Control.Dispatcher.Invoke([action]{
            # This bit is only really meaningful for the TextBox control, which might be useful for logging progress steps
            If ($PSBoundParameters['AppendContent']) {
                $syncHash.$Control.AppendText($Value)
            } Else {
                $syncHash.$Control.$Property = $Value
            }
        }, "Normal")
    }                        
<#
Update-Window -Control StarttextBlock -Property ForeGround -Value White                                                       
start-sleep -Milliseconds 850
$x += 1..15000000
update-window -Control ProgressBar -Property Value -Value 25

update-window -Control TextBox -property text -value $x -AppendContent
Update-Window -Control ProcesstextBlock -Property ForeGround -Value White                                                       
start-sleep -Milliseconds 850
update-window -Control ProgressBar -Property Value -Value 50

Update-Window -Control FiltertextBlock -Property ForeGround -Value White                                                       
start-sleep -Milliseconds 500
update-window -Control ProgressBar -Property Value -Value 75

Update-Window -Control DonetextBlock -Property ForeGround -Value White                                                       
start-sleep -Milliseconds 200
update-window -Control ProgressBar -Property Value -Value 100
#>




Function Submit-RunspaceChange 
{ [cmdletBinding()]
Param($path)
 write-debug "Enter Submit";
get-LastLine($event.sourceEventArgs.fullpath)
$syncHash.listView.Dispatcher.Invoke([action]{ if ($syncHash.listView.items.count -gt 0) 
{write-verbose "Refresh" ;$syncHash.listView.Items.Refresh()} 
else {write-verbose "add"; $syncHash.listView.ItemsSource = $computers},"Normal"}) 

}


function Get-LastLine
{ [cmdletBinding()]
Param($path)
   
    #$oldConsole = [console]::TreatControlCAsInput
    #[console]::TreatControlCAsInput = $true
    #write-host "enter"
    #write-host "computers : $($global:computers -isnot [System.Array])"
    $synchash.Host.ui.write("Get-LastLine")
   # $synchash.Host.ui.d
    write-debug "entrance"
    write-verbose "Entering LastLine"
   if (!$global:computers) 
       { write-host "No Global"
         $global:computers = New-Object System.Collections.ObjectModel.ObservableCollection[object]

#@()
        }
    if ($global:computers -isnot [System.Collections.ObjectModel.ObservableCollection[object]])
       {
         Write-Host "no array"
         $global:computers = New-Object System.Collections.ObjectModel.ObservableCollection[object]



       }
  
        $comp = $global:computers  
        #write-host "Comp $($comp -is [System.Array])"
   
    $stat = "" | select Computer,Action,Time, Progress, Description
    $lines = Get-Content $path
    $lines = $lines.split("`n")
    
    if ( $Lines[$lines.count-1].Trim().Length -gt 0 ) 
        { $line = $lines[$line.count-1] }
    else 
        { $line = $lines[$lines.count-2] }
    $line = $line.Split(';')
    $stat.computer = $line[0]
    $stat.action = $line[1]
    $stat.Time = $line[2]
    $stat.Description = $line[3]
    $stat.Progress = get-random -Maximum 100
    #write-host "Count $($comp.count)"
    if ($Comp.count -eq 0 )
      {  #Write-host "Zero"
         $comp.Add($stat)}


    if ($comp.computer.Contains($stat.computer))
        {
            $index = $comp.computer.IndexOf($stat.computer)
          # write-host "INdex : $index"
           #write-host "Computer : $($comp[$index])"
            $comp[$index] = $stat
        }
        else
        {  
            #write-host "adding"
            $comp.add($stat)
          
       }
    
    #write-host "exit"
     $global:computers = $comp
    
     $global:status = $stat
    write-host "refresh"
    Write-Verbose "refresh 2"
    $synchash.window.Dispatcher.invoke([action]{$synchash.listview.items.refresh(),"Normal"})
     $Synchash.Host.Runspace.Events.GenerateEvent("ListViewChanged",$syncHash.listView,$null,"ListView Changed")
    Register-EngineEvent -SourceIdentifier "ListViewChanged" -Action {$synchash.host.ui.write("Event Happened inside")} -Forward
    $test = "na"
    if ($syncHash.listView.items.count -gt 0) {write-host "test 321" ;$test = "test 321";$syncHash.listView.Items.Refresh()} else {$test = "nope" ;$syncHash.listView.ItemsSource = $computers}
}


        })
        $PowerShell.Runspace = $newRunspace
        [void]$Jobs.Add((
            [pscustomobject]@{
                PowerShell = $PowerShell
                Runspace = $PowerShell.BeginInvoke()
            }
        ))
    })


  


    #region Window Close 
    $syncHash.Window.Add_Closed({
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
$data = $psCmd.BeginInvoke()
Start-Sleep -Seconds 2

If (!(Test-Path variable:computers) )
{
    #COmputers Variable doesn't exist, create empty structure
   #$computers=@()    
   $computers =   New-Object System.Collections.ObjectModel.ObservableCollection[object]

}


Register-EngineEvent -SourceIdentifier "ListViewChanged" -Action {$synchash.host.ui.write("Event Happened outside")}
$fsw = New-Object System.IO.FileSystemWatcher "C:\wsus\test2\nk23208", "*.csv" 
$event = Register-ObjectEvent -InputObject $fsw -EventName "Changed" -action { write-Debug "changed"; 
Write-debug "$($event.sourceEventArgs.fullpath)"; Submit-RunspaceChange($event.sourceEventArgs.fullpath)}

#;cls ;$computers | ft * -AutoSize | out-host}


$synchash.listView.Dispatcher.Invoke([action]{$syncHash.listView.ItemsSource = $computers},"Normal")
    $VerbosePreference = "continue"
    $Global:timer = new-object System.Windows.Threading.DispatcherTimer

    #Fire off every 5 seconds

    Write-Verbose “Adding 5 second interval to timer object”

    $timer.Interval = [TimeSpan]”0:0:5.00"

    #Add event per tick

    Write-Verbose "Adding Tick Event to timer object"

    $global:timer.Add_Tick({

        
        write-host "updating window"
        Write-Verbose “Updating Window”
        $syncHash.listView.Dispatcher.Invoke([action]{ if ($syncHash.listView.items.count -gt 0) {$syncHash.listView.Items.Refresh()} else {$syncHash.listView.ItemsSource = $computers},"Normal"})

        })

    #Start timer

    Write-Verbose “Starting Timer”

    $timer.Start()



 


#$syncHash.listView.Dispatcher.Invoke([action]{$syncHash.listView.ItemsSource=$computers},"Normal")
#$syncHash.listView.Dispatcher.Invoke([action]{$syncHash.listView.Items.Refresh()},"Normal")

#$syncHash.listView.add_sourceupdated({$syncHash.listView.Dispatcher.Invoke([action]{ if ($syncHash.listView.items.count -gt 0) {$syncHash.listView.Items.Refresh()} else {$syncHash.listView.ItemsSource = $computers},"Normal"})})


# cleanup
# Get-Runspace | ? {$_.RunspaceAvailability -eq "Available"} | % { $_.close(); $_.Dispose()}  

 