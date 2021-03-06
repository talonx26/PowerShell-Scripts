﻿cls
$Global:syncHash = [hashtable]::Synchronized(@{})
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)


$psCmd = [PowerShell]::Create().AddScript({
    [xml]$xaml = @"
 <Window
           xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
           xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
           xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
       Title="WSUS Install Monitor" Height="700" Width="1131.962">
  <Grid Margin="0,0,-185.4,0.2">
    
   
    <DataGrid x:Name="dataGrid" HorizontalAlignment="Left" Margin="44,30.2,0,0" VerticalAlignment="Top" Height="Auto" Width="Auto">
      <DataGrid.Resources>
         <Style TargetType="{x:Type DataGridRow}">
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
            <ProgressBar Maximum="100" Value="{Binding Progress}" Width="100" Height="10" Margin="0"/>
            <TextBlock Text="{Binding Progress, StringFormat={}{0}%}" HorizontalAlignment="Center"/>
          </Grid>
        </DataTemplate>
      </DataGrid.Resources>
      <DataGrid.Columns>
        <DataGridTextColumn Binding="{Binding Computer}"  Header="Computer"/>
        <DataGridTextColumn Binding="{Binding Action}" Header="Action"/>
        <DataGridTextColumn Binding="{Binding Time}" Header="Time"/>
        <DataGridTemplateColumn Header="Progress" Width="100">
           <DataGridTemplateColumn.CellTemplate>
              <DataTemplate>
                 <ProgressBar Value="{Binding Path=Progress, Mode=OneWay}" Minimum="0" Maximum="100" />
                  
              </DataTemplate>
           </DataGridTemplateColumn.CellTemplate>
        </DataGridTemplateColumn>
        <DataGridTextColumn Binding="{Binding Description}" Header="Status" />
        
      </DataGrid.Columns>
    </DataGrid>
  </Grid>
</Window>

"@


    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
    
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $xaml
        $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
        #Find all of the form types and add them as members to the synchash
        $syncHash.Add($_.Name,$syncHash.Window.FindName($_.Name) )

    }

    $Script:JobCleanup = [hashtable]::Synchronized(@{})
    $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

    #region Background runspace to clean up jobs
    $jobCleanup.Flag = $True
    $newRunspace =[runspacefactory]::CreateRunspace()
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


 #
 
 
 function Update-DataGrid 
 { [CmdletBinding()]
	param(
		[Parameter(Mandatory=$True,
		ValueFromPipeline=$True)]
		[array[]]$c
	)
	BEGIN {}
	PROCESS {
    $syncHash.dataGrid.Dispatcher.invoke([action]{
    $w = $synchash.dataGrid.items.computer.IndexOf($c.Computer)
    if($w -gt -1 ) 
      {$synchash.dataGrid.items[$w]=$c}
    else 
      {$synchash.dataGrid.AddChild($c)}
    $synchash.dataGrid.items.Refresh()
    },"Normal")  }
	END {}


     
 }

  #$syncHash.dataGrid.Dispatcher.invoke([action]{ $w = $synchash.dataGrid.items.computer.IndexOf($c.Computer);if($w -gt -1 ) {$synchash.dataGrid.items[$w]=$c} else {$synchash.dataGrid.AddChild($c)};$synchash.dataGrid.items.Refresh()},"Normal")  
 #$syncHash.dataGrid.Dispatcher.Invoke([action]{$syncHash.dataGrid.ItemsSource = $computers})