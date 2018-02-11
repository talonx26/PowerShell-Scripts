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

    Title="NSLookup Utility" SizeToContent="WidthAndHeight" Width="536" Height="378" MinWidth="520" MinHeight="378">
    <Grid Margin="0,0,2,0">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="404*"/>
            <ColumnDefinition Width="45*"/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="txtInput" HorizontalAlignment="Left" Height="23" Margin="126,49,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="120"/>
        <Label x:Name="label" Content="Input File" HorizontalAlignment="Left" Margin="15,49,0,0" VerticalAlignment="Top" Height="27"/>
        <Label x:Name="label1" Content="Output file" HorizontalAlignment="Left" Margin="15,81,0,0" VerticalAlignment="Top" Width="88"/>
        <TextBox x:Name="txtOutput" HorizontalAlignment="Left" Height="23" Margin="126,81,0,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Width="120"/>
        <Button x:Name="btnStart" Content="Start" HorizontalAlignment="Right" Margin="0,49,21,0" Width="75" VerticalAlignment="Top" Grid.ColumnSpan="2"/>
        <ListView x:Name="IPResults" Margin="10,114,21,21" RenderTransformOrigin="0.508,0.533" MinHeight="200" MinWidth="450" Grid.ColumnSpan="2" >
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Target"      Width="150" DisplayMemberBinding="{Binding Target}"/>
                    <GridViewColumn Header="HostName" Width="150" DisplayMemberBinding="{Binding HostName}"/>
                    <GridViewColumn Header="IPAddress" Width="150" DisplayMemberBinding="{Binding IPAddress}"/>
                </GridView>
            </ListView.View>

        </ListView>
        <TextBox x:Name="txtCurrDir" Margin="126,21,21,0" TextWrapping="Wrap" Text="TextBox" VerticalAlignment="Top" Grid.ColumnSpan="2"/>
        <Label Content="Current Directory" HorizontalAlignment="Left" Margin="15,21,0,0" VerticalAlignment="Top"/>

    </Grid>

</Window>
"@


    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $syncHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
	$form=[Windows.Markup.XamlReader]::Load( $reader )
	
	$syncHash.Host = $Host

    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
	[System.Reflection.Assembly]::LoadWithPartialName("WindowsFormsIntegration")
	[void][System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
	

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
	<#$SyncHash.btnStart.add_Click({
	
	 #$syncHash.txtinput.text.dispatcher.invoke([action]{$syncHash.txtinput.Text = "test"},"Normal")
		$syncHash.txtinput.text = "test"
	 })
	#>
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
    $syncHash.btnStart.Add_Click({
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
		$syncHash.File = Get-content $file
        $newRunspace.ApartmentState = "STA"
		$newRunspace.Name = "DNSQuery"
        $newRunspace.ThreadOptions = "ReuseThread"          
        $newRunspace.Open()
        $newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash) 
        $PowerShell = [PowerShell]::Create().AddScript({

$synchash.file | % {
   remove-variable R -ErrorAction SilentlyContinue | out-null
   
   if ($_.split(".").count -eq 4)
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
    $synchash.ips.add($ip)

    }
    Catch
    {
    #Write-host "IP Address Exception $Name"
    $IP = New-Object psobject
    $IP | Add-Member -Type NoteProperty -Name Target -value $name
    $IP | Add-Member -Type NoteProperty -Name HostName -Value "Not Found"
    $IP | Add-Member -Type NoteProperty -Name IPAddress -Value $name
   $synchash.ips.add($ip)
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
         $synchash.ips.add($ip)
       }
  
   }

    remove-variable R -ErrorAction SilentlyContinue | out-null
    } 
		if ($syncHash.txtOutput.Text -notlike "*.csv")
		{ $syncHash.txtOutput.Text += ".csv"}
		$syncHash.IPS | Export-Csv "$($syncHash.txtcurrdir.text)\$($synchash.txtoutput.text)" -NoTypeInformation

})
		$SyncHash.Host.UI.Write( "button")
        #Start-Job -Name Sleeping -ScriptBlock {start-sleep 5}
        #while ((Get-Job Sleeping).State -eq 'Running'){
            $x+= "."
        #region Boe's Additions
        <#
        $newRunspace =[runspacefactory]::CreateRunspace()
        $newrunspace.Name ="btnStart"
        $newRunspace.ApartmentState = "STA"
        $newRunspace.ThreadOptions = "ReuseThread"          
        $newRunspace.Open()
        $newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash) 
        $PowerShell = [PowerShell]::Create().AddScript({
			
			Write-Host "click"
			#$syncHash.txtinput.text.dispatcher.invoke([action]{$syncHash.txtinput.Text = "test"},"Normal")

})   #>
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

# How to add to Background UI
#$syncHash.Window.Dispatcher.Invoke([action]{$ips.Add($ip)},"Normal")
<#
$syncHash.txtInput.Add_LostFocus({
	$syncHash.txtinput.dispatcher.invoke([action]{$global:test = $syncHash.txtinput.text})
	
})
#>

	function RunspacePing {
param($syncHash)
if ($Count -eq $null)
    {NullCount; break}
 
$syncHash.Host = $host
$Runspace = [runspacefactory]::CreateRunspace()
$Runspace.ApartmentState = "STA"
$Runspace.ThreadOptions = "ReuseThread"
$Runspace.Open()
$Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash) 
#$Runspace.SessionStateProxy.SetVariable("count",$count)
#$Runspace.SessionStateProxy.SetVariable("ComputerName",$ComputerName)
#$Runspace.SessionStateProxy.SetVariable("TargetBox",$TargetBox)
 
$code = {
    $syncHash.Window.Dispatcher.invoke([action]{$Global:t = $syncHash.txtInput.Text},"Normal")
	Write-Host $t
    
}
$PSinstance = [powershell]::Create().AddScript($Code)
$PSinstance.Runspace = $Runspace
$job = $PSinstance.BeginInvoke()
}




	#Write-Host $return
	<#
Get-content $syncHash.txtInput.Text | % {
	
   remove-variable R -ErrorAction SilentlyContinue | out-null
   
   if ($_.split(".").count -eq 4)
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
    $ips.add($IP)

    }
    Catch
    {
    #Write-host "IP Address Exception $Name"
    $IP = New-Object psobject
    $IP | Add-Member -Type NoteProperty -Name Target -value $name
    $IP | Add-Member -Type NoteProperty -Name HostName -Value "Not Found"
    $IP | Add-Member -Type NoteProperty -Name IPAddress -Value $name
    $ips.add($IP)
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
       $ips.add($IP)
       }
       }
       catch
       {
          #Write-host "HostName Exeption $name"
          $IP = New-Object psobject
          $IP | Add-Member -Type NoteProperty -Name Target -Value $name
          $IP | Add-Member -Type NoteProperty -Name HostName -Value $name
          $IP | Add-Member -Type NoteProperty -Name IPAddress -Value "Not Found"
          $ips.add($IP)
       }
  
   }

    remove-variable R -ErrorAction SilentlyContinue | out-null
    }
	
	})#>



function close-OrphanedRunSpaces()
{
   Get-Runspace
   Write-Host "closing"
    Get-Runspace | ? { $_.RunspaceAvailability -eq "Available"} | % { $_.close();$_.Dispose()}
   write-host "Closed"
   Get-Runspace
}



