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

$ips = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$IP = "" | Select Target, HostName, IPAddress
$ip.Target = "test1"
$ip.HostName = "Test2"
$ip.IPAddress =  "127.0.0.1"
$ips.add($ip)

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

    Title="NSLookup Utility" Height="365.061" Width="488.725" SizeToContent="WidthAndHeight">
    <Grid>
        <TextBox x:Name="txtInput" HorizontalAlignment="Left" Height="23" Margin="121,17,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="120"/>
        <Label x:Name="label" Content="Input File" HorizontalAlignment="Left" Margin="10,17,0,0" VerticalAlignment="Top" Height="27"/>
        <Label x:Name="label1" Content="Output file" HorizontalAlignment="Left" Margin="10,49,0,0" VerticalAlignment="Top" Width="88"/>
        <TextBox x:Name="txtOutput" HorizontalAlignment="Left" Height="23" Margin="121,49,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="120"/>
        <Button x:Name="btnStart" Content="Start" HorizontalAlignment="Left" Margin="336,13,0,0" VerticalAlignment="Top" Width="75"/>
        <ListView x:Name="IPResults" HorizontalAlignment="Left" Height="239" Margin="10,80,0,0" VerticalAlignment="Top" Width="450" >
            <ListView.View>
                <GridView>
                   <GridViewColumn Header="Target"      Width="150" DisplayMemberBinding="{Binding Target}"/>
                    <GridViewColumn Header="HostName" Width="150" DisplayMemberBinding="{Binding HostName}"/>
                    <GridViewColumn Header="IPAddress" Width="150" DisplayMemberBinding="{Binding IPAddress}"/>
                </GridView>
            </ListView.View>
        </ListView>

    </Grid>

</Window>
"@


    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
    $syncHash.Add("IPAdd",$ips)
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
			Param ($IPAddress)
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
        }).AddArgument($IPS)
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


1..5 | %{ write-host "!!!!!!!!!!!!!!!!!!"}
Write-host "to run display your UI, run:  " -NoNewline
write-host -foregroundcolor Green '$data = $psCmd.BeginInvoke()'
$data = $psCmd.BeginInvoke()
sleep -Milliseconds 500

$syncHash.IPResults.Dispatcher.Invoke([action]{$syncHash.IPResults.ItemsSource = $ips},"Normal")