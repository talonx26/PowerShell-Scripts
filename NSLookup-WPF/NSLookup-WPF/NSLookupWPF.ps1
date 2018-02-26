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
$syncHash.path = $PWD.path
$syncHash.IPS = $ips
$syncHash.IP = $ip
$newRunspace.Name = "GUI"
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)

$psCmd = [PowerShell]::Create().AddScript({
    [xml]$xaml = @"
  <Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

    Title="NSLookup Utility" SizeToContent="WidthAndHeight" Width="485.6" Height="395.909" MinWidth="520" MinHeight="378" MaxWidth="840" MaxHeight="700">
    <Grid Margin="0,0,2,0">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="472*"/>
            <ColumnDefinition Width="5*"/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="txtInput" HorizontalAlignment="Left" Height="23" Margin="145,49,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="160"/>
        <Label x:Name="label" Content="Input File" HorizontalAlignment="Left" Margin="15,49,0,0" VerticalAlignment="Top" Height="27" Width="72"/>
        <Label x:Name="label1" Content="Output file" HorizontalAlignment="Left" Margin="15,81,0,0" VerticalAlignment="Top" Width="88" Height="30"/>
        <TextBox x:Name="txtOutput" HorizontalAlignment="Left" Height="23" Margin="145,81,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="160"/>
        <Button x:Name="btnStart" Content="Start" HorizontalAlignment="Right" Margin="0,49,-0.2,0" Width="75" VerticalAlignment="Top" Grid.ColumnSpan="2" Height="26" IsEnabled="False"/>

        <DataGrid x:Name="IPResults" AutoGenerateColumns="False" Margin="15,129,10,9.2" MinHeight="200" MinWidth="450"  MaxHeight="600" MaxWidth="800" HorizontalContentAlignment="Stretch"  >
            <DataGrid.Columns>
                <DataGridTextColumn Binding="{Binding Target}" ClipboardContentBinding="{x:Null}" Header="Target" MinWidth="150" Width="*"/>
                <DataGridTextColumn Binding="{Binding HostName}" ClipboardContentBinding="{x:Null}" Header="HostName" MinWidth="150" Width="*"/>
                <DataGridTextColumn Binding="{Binding IPAddress}" ClipboardContentBinding="{x:Null}" Header="IPAddress" MinWidth="150" Width="*"/>
            </DataGrid.Columns>
        </DataGrid>
        <TextBox x:Name="txtCurrDir" Margin="145,21,-0.2,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Grid.ColumnSpan="2" Height="26"/>
        <Label Content="Current Directory" HorizontalAlignment="Left" Margin="15,21,0,0" VerticalAlignment="Top" Height="30" Width="125"/>
		<ProgressBar x:Name="Progress" HorizontalAlignment="Left" Height="20" Margin="15,109,0,0" VerticalAlignment="Top" Width="337"/>
        <Label x:Name="lblProgress" Content="Label" HorizontalAlignment="Left" Margin="15,104,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.455,1.02" Width="337" HorizontalContentAlignment="Center"/>
	    <Button x:Name="btnExport" Grid.ColumnSpan="2" Content="Export" HorizontalAlignment="Right" Margin="402,78,-0.2,0" VerticalAlignment="Top" Width="75" IsEnabled="False"/>

    </Grid>

</Window>
"@
        [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
        [System.Reflection.Assembly]::LoadWithPartialName("WindowsFormsIntegration")
        [void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
	$form=[Windows.Markup.XamlReader]::Load( $reader )

	$syncHash.Host = $Host

    

	#$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
    [xml]$XAML = $xaml
        $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
        #Find all of the form types and add them as members to the synchash
        $syncHash.Add($_.Name,$syncHash.Window.FindName($_.Name) )
    }
	$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $form.FindName($_.Name)}
    $Script:JobCleanup = [hashtable]::Synchronized(@{})
    $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

    #region Background runspace to clean up jobs
    $jobCleanup.Flag = $True
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.Name = "Cleanup"
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
$syncHash.txtCurrDir.Dispatcher.Invoke([action]{ $syncHash.txtCurrDir.text = $syncHash.path},"Normal")
$syncHash.IPResults.Dispatcher.Invoke([action]{$syncHash.IPResults.ItemsSource = $ips},"Normal")
$syncHash.Window.Dispatcher.Invoke([action]{$syncHash.IPS.Clear()},"Normal")
# How to add to Background UI
#$syncHash.Window.Dispatcher.Invoke([action]{$ips.Add($ip)},"Normal")
<#
$syncHash.txtInput.Add_LostFocus({
	$syncHash.txtinput.dispatcher.invoke([action]{$global:test = $syncHash.txtinput.text})
})
#>


function close-OrphanedRunSpaces()
{
   Get-Runspace
   Write-Host "closing"
    Get-Runspace | ? { $_.RunspaceAvailability -eq "Available"} | % { $_.close();$_.Dispose()}
   write-host "Closed"
   Get-Runspace
}