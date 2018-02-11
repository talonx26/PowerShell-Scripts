#
# Script3.ps1
#
# WPF GUI Script for Folder Creation

#region Initialization
# Get the script location
$ScriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

# Load Windows Presentation Framework
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

# Create a wShell object for simple Popup boxes
$wshell = New-Object -ComObject Wscript.Shell

#region Import ActiveDirectory PowerShell Module
try {
    $null = Import-Module ActiveDirectory -ea stop
} catch {
    $catch = $wshell.Popup("Error: ActiveDirectory Module Not Available",0,"ERROR",0x2)
    switch ($catch) {
        "3" { # Abort
                exit 
            }
        "4" { # Retry
                try {
                    $null = Import-Module ActiveDirectory -ea stop
                } catch {
                    $catch = $wshell.Popup("Error: ActiveDirectory Module Not Available",0,"ERROR",0x0)
                    exit
                }
            }
    } #end switch
} #end catch
#endregion

#region Import NTFSSecurity PowerShell Module
try {
    $null = Import-Module $ScriptDir\NTFSSecurity -ea stop
} catch {
    $catch = $wshell.Popup("Error: NTFSSecurity Module Not Available",0,"ERROR",0x2)
    switch ($catch) {
        "3" { # Abort
                exit 
            }
        "4" { # Retry
                try {
                    $null = Import-Module $ScriptDir\NTFSSecurity -ea stop
                } catch {
                    $catch = $wshell.Popup("Error: NTFSSecurity Module Not Available",0,"ERROR",0x0)
                    exit
                }
            }
    } #end switch
} #end catch
#endregion

#endregion

#region XAML GUI Code
$inputXML = @"
<Window x:Class="WpfApplication1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApplication1"
        mc:Ignorable="d"
        Title="Foldernator 9000" Height="505" Width="790" WindowStartupLocation="CenterScreen" Topmost="False" WindowStyle="ThreeDBorderWindow" Cursor="Hand">
    <Grid>
        <GroupBox x:Name="groupBoxCommon" Header="Common Options" HorizontalAlignment="Left" Margin="520,0,0,0" VerticalAlignment="Top" Width="247" Height="125"/>
        <GroupBox x:Name="groupBoxImport" Header="Import CSV Folder List" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" Height="125" Width="250" Foreground="Red" BorderBrush="Red"/>
        <ComboBox x:Name="comboBoxCsv" HorizontalAlignment="Left" Margin="24,30,0,0" VerticalAlignment="Top" Width="225" Height="18" FontSize="10" SelectedIndex="0"/>
        <Button x:Name="buttonRefresh" Content="Refresh the CSV File List" HorizontalAlignment="Left" Margin="24,76,0,0" VerticalAlignment="Top" Width="225" Height="18" FontSize="10" Background="#FFF592FF" ToolTip="Updates the above list. Useful if you dropped a new CSV in the folder"/>
        <Button x:Name="buttonImport" Content="Import Selected CSV File" HorizontalAlignment="Left" Margin="24,99,0,0" VerticalAlignment="Top" Width="225" Height="18" FontSize="10" Background="#FFFF6D6D"/>
        <GroupBox x:Name="groupBoxManual" Header="Manual Folder Input" HorizontalAlignment="Left" Margin="265,0,0,0" VerticalAlignment="Top" Height="125" Width="250" Foreground="#FF009FFF" BorderBrush="#FF009FFF"/>
        <TextBox x:Name="textBoxFolder" HorizontalAlignment="Left" Height="18" Margin="278,30,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="225" FontSize="10" Padding="0,1,0,0">
            <TextBox.ToolTip>
                <TextBlock>
                    Enter a folder name (eg: D:\Shares\GroupData\Admin)
                </TextBlock>
            </TextBox.ToolTip>
        </TextBox>
        <CheckBox x:Name="checkBoxBreakInheritance" Content="Break Inheritance" HorizontalAlignment="Left" Margin="278,53,0,0" VerticalAlignment="Top" Height="15" Width="100" FontSize="10">
            <CheckBox.ToolTip>
                <TextBlock>
                    Inheritance is Disabled for this folder. Existing permissions are copied.
                    <LineBreak />
                    Permission groups matching `$NETBIOS\FL- are stripped from the folder.
                    <LineBreak />
                    New security groups will be created and permissions applied accordingly.
                </TextBlock>
            </CheckBox.ToolTip>
        </CheckBox>
        <CheckBox x:Name="checkBoxShare" Content="Share" HorizontalAlignment="Left" Margin="403,73,0,0" VerticalAlignment="Top" Width="100" FontSize="10">
            <CheckBox.ToolTip>
                <TextBlock>
                    A 'hidden' share will be created for the folder.
                    <LineBreak />
                    Access Based Enumeration will be enabled for the share.
                </TextBlock>
            </CheckBox.ToolTip>
        </CheckBox>
        <CheckBox x:Name="checkBoxThisFolderOnly" Content="This Folder Only" HorizontalAlignment="Left" Margin="278,73,0,0" VerticalAlignment="Top" FontSize="10" Width="100">
            <CheckBox.ToolTip>
                <TextBlock>
                    Permissions will be added to This Folder Only.
                    <LineBreak />
                    They will not propagate to subfolders or files.
                </TextBlock>
            </CheckBox.ToolTip>
        </CheckBox>
        <CheckBox x:Name="checkBoxUserFolder" Content="User Folder" HorizontalAlignment="Left" Margin="403,53,0,0" VerticalAlignment="Top" Width="100" FontSize="10">
            <CheckBox.ToolTip>
                <TextBlock>
                    Permissions will be added to this folder so Domain Users can create subfolders
                    <LineBreak />
                    that they have full control over (think: Folder Redirection).
                </TextBlock>
            </CheckBox.ToolTip>
        </CheckBox>
        <Button x:Name="buttonFolderAdd" Content="Add Folder to List" HorizontalAlignment="Left" Margin="278,99,0,0" VerticalAlignment="Top" Width="225" Height="18" FontSize="10" Background="#FFA6AEFF"/>
        <Button x:Name="buttonProcess" Content="Process List and Create Folders" HorizontalAlignment="Left" Margin="531,99,0,0" VerticalAlignment="Top" Width="225" FontSize="10" Height="18" Background="Lime"/>
        <Label x:Name="labelShorten" Content="Shorten Group Names n Levels:" HorizontalAlignment="Left" Margin="560,53,0,0" VerticalAlignment="Top" FontSize="10" Width="150" Height="18" Padding="5,4,5,2">
            <Label.ToolTip>
                <TextBlock>
                    You can use this to strip the drive letter and some of the folder
                    <LineBreak />
                    path out of the automatically generated Group Names.
                    <LineBreak />
                    <LineBreak />
                    Example: If your path is D:\Shares\GroupData\Admin
                    <LineBreak />
                    0 = FL-D!Shares!GroupData!Admin
                    <LineBreak />
                    1 = FL-Shares!GroupData!Admin
                    <LineBreak />
                    2 = FL-GroupData!Admin
                    <LineBreak />
                    etc etc
                </TextBlock>
            </Label.ToolTip>
        </Label>
        <ComboBox x:Name="comboBoxShorten" HorizontalAlignment="Left" Margin="721,53,0,0" VerticalAlignment="Top" Width="35" Height="18" FontSize="10" SelectedIndex="2"/>
        <Button x:Name="buttonClear" Content="Clear Selected Items" HorizontalAlignment="Left" Margin="531,76,0,0" VerticalAlignment="Top" Width="225" FontSize="10" Height="18" Background="#FFF9FF83" ToolTip="Select rows from the list then click this to remove them."/>
        <ListView x:Name="listView" HorizontalAlignment="Left" Height="312" Margin="10,130,0,0" VerticalAlignment="Top" Width="755" BorderBrush="#FF000000" AllowDrop="True" SelectionMode="Extended" FontSize="10">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Folder" DisplayMemberBinding ="{Binding Folder}" Width="255"/>
                    <GridViewColumn Header="Group Name (Base)" DisplayMemberBinding ="{Binding GroupName}" Width="230"/>
                    <GridViewColumn Header="BreakInheritance" DisplayMemberBinding ="{Binding BreakInheritance}" Width="50"/>
                    <GridViewColumn Header="ThisFolderOnly" DisplayMemberBinding ="{Binding ThisFolderOnly}" Width="50"/>
                    <GridViewColumn Header="Share" DisplayMemberBinding ="{Binding Share}" Width="40"/>
                    <GridViewColumn Header="UserFolder" DisplayMemberBinding ="{Binding UserFolder}" Width="50"/>
                    <GridViewColumn Header="Status" DisplayMemberBinding ="{Binding Status}" Width="50"/>
                </GridView>
            </ListView.View>
        </ListView>
        <ProgressBar x:Name="ProgressBar" HorizontalAlignment="Left" Height="14" Margin="10,447,0,0" VerticalAlignment="Top" Width="755" Background="Black" BorderBrush="Black" Foreground="Lime" Minimum="0" Maximum="100"/>
        <ComboBox x:Name="comboBoxOU" HorizontalAlignment="Left" Margin="531,30,0,0" VerticalAlignment="Top" Width="225" SelectedIndex="0" FontSize="10" Height="18"/>
    </Grid>
</Window>
"@
#endregion

#region XAML Import
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'

[xml]$XAML = $inputXML

# Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{
    Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
    return
}

# Store form objects in variables
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}

#endregion

#region Functions

Function Do-GroupName {
    param(
        [parameter(Mandatory=$true,Position=0)][string]$Folder,
        [parameter(Mandatory=$true,Position=1)][string]$Shorten
    )

    $objInput = $Folder.Replace(":","").Replace("\","!").Split("!")
    $objOutput = "FL-"

    [int]$i=0
    foreach ($objSplit in $objInput) {
        if ($i -ge $Shorten) {
          $objOutput += $objSplit
            if ($i+1 -lt $objInput.Count) {
                $objOutput += "!"
            }
        }
        $i++
    }

    return $objOutput
}

Function Do-ProcessList {

    $SyncHash = [hashtable]::Synchronized(@{Form = $Form; WPFlistView = $WPFlistView; WPFProgressBar = $WPFProgressBar})
    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash)
    $Worker = [PowerShell]::Create().AddScript({
        foreach ($i in $SyncHash.WPFlistView.Items) {
            $i.Status="Working"
            start-sleep -Milliseconds 100
            $SyncHash.WPFProgressBar.Value++
        }
        foreach ($i in(1..100)) {
            start-sleep -Milliseconds 10
            $SyncHash.WPFProgressBar.Value++
        }
    })
    $Worker.Runspace = $Runspace
    $Worker.BeginInvoke()

}

#endregion

#region Set Actions/Values/Properties For GUI Objects

#region Combo Boxes
# Combo Box: Csv
$CSVList = @('Select a CSV file')
$CSVList += (gci $ScriptDir -Filter *.csv).Name
$WPFcomboBoxCsv.ItemsSource = $CSVList

# Combo Box: OUList
$OUList = @('Select an OU to create groups in')
$OUList += ((Get-ADOrganizationalUnit -filter *).DistinguishedName | Sort-Object)
$WPFcomboBoxOU.ItemsSource = $OUList

# Combo Box: Shorten
$ShortenItems = @('0','1','2','3','4')
$WPFcomboBoxShorten.ItemsSource = $ShortenItems
#endregion

#region Checkboxes
# CheckBox: BreakInheritance
$WPFcheckBoxBreakInheritance.Add_Checked({
    $WPFcheckBoxUserFolder.IsChecked=$false
})
$WPFcheckBoxBreakInheritance.Add_UnChecked({
    $WPFcheckBoxThisFolderOnly.IsChecked=$false
})

# CheckBox: ThisFolderOnly
$WPFcheckBoxThisFolderOnly.Add_Checked({ 
    $WPFcheckBoxBreakInheritance.IsChecked=$true
    $WPFcheckBoxUserFolder.IsChecked=$false
})

# CheckBox: UserFolder
$WPFcheckBoxUserFolder.Add_Checked({
    $WPFcheckBoxBreakInheritance.IsChecked=$false
})
#endregion

#region Buttons
# Button: Refresh
$WPFbuttonRefresh.Add_Click({
    $CSVList = @('Select a CSV file')
    $CSVList += (gci $ScriptDir -Filter *.csv).Name
    $WPFcomboBoxCsv.ItemsSource = $CSVList
})

# Button: Import
$WPFbuttonImport.Add_Click({
    $CSVfile = $WPFcomboBoxCsv.SelectedValue
    if (Test-Path $ScriptDir\$CSVfile) {
        $CSV = Import-Csv $ScriptDir\$CSVfile
        if ($CSV.Folder) {
            $Shorten = $WPFcomboBoxShorten.Text
            $CSV | ForEach-Object {
                $Status = "Queued"
                if ($_.BreakInheritance -eq 'True') {
                    if ($_.UserFolder -eq 'True') {
                        $Status="Warning"
                    }
                } else {
                    if ($_.ThisFolderOnly -eq 'True') {
                        $Status="Warning"
                    }
                }
                if ( !($_.Status) ) {
                    Add-Member -InputObject $_ -NotePropertyName 'Status' -NotePropertyValue $Status
                }
                if ( !($_.GroupName) ) {
                    if ($_.BreakInheritance -eq "False") {
                        $GroupName = "None"
                    } else {
                        $GroupName = Do-GroupName $_.Folder $Shorten
                    }
                    Add-Member -InputObject $_ -NotePropertyName 'GroupName' -NotePropertyValue $GroupName
                }
                $WPFlistView.Items.Add($_)
            } #end Foreach-Object
        } #end if $CSV.Folder
    } #end if Test-Path
}) #end Add_Click

# Button: FolderAdd
$WPFbuttonFolderAdd.Add_Click({
    if ($WPFcheckBoxBreakInheritance.IsChecked -eq $true) {
        $GroupName = Do-GroupName $WPFtextBoxFolder.Text $WPFcomboBoxShorten.Text
    } else {
        $GroupName = "None"
    }
    $Status = "Queued"
    $WPFlistView.Items.Add([pscustomobject]@{
        'Folder'=$WPFtextBoxFolder.Text;
        'GroupName'=$GroupName;
        'BreakInheritance'=$WPFcheckBoxBreakInheritance.IsChecked;
        'ThisFolderOnly'=$WPFcheckBoxThisFolderOnly.IsChecked;
        'Share'=$WPFcheckBoxShare.IsChecked;
        'UserFolder'=$WPFcheckBoxUserFolder.IsChecked;
        'Status'=$Status})
})

# Button: Clear
$WPFbuttonClear.Add_Click({
    while ($WPFlistView.SelectedItems.Count -gt 0) {
        $WPFlistView.Items.RemoveAt($WPFlistView.SelectedIndex)
    }
})

# Button: Process
$WPFbuttonProcess.Add_Click({
    
    $OU = $WPFcomboBoxOU.SelectedValue
    $FolderList = @()
    $FolderList += $WPFlistView.Items

    if ( !($OU.StartsWith("OU=")) ) {
        if ( !($FolderList) ) {
            $wshell.Popup("Add Folders to the list and select an OU",0,"OK",0x1)
            #return
        } else {
            $wshell.Popup("Select an OU",0,"OK",0x1)
            #return
        }
    }
    if ( !($FolderList) ) {
        $wshell.Popup("Add Folders to the list",0,"OK",0x1)
        #return
    }

    Do-ProcessList

})
#endregion

#region ListView
# List View: ListView
$WPFlistView.Add_MouseLeftButtonUp({

    $WPFtextBoxFolder.Text = $WPFlistView.SelectedItem.Folder
    if ($WPFlistView.SelectedItem.BreakInheritance -eq 'True') {
        $WPFcheckBoxBreakInheritance.IsChecked=$true
    } else {
        $WPFcheckBoxBreakInheritance.IsChecked=$false
    }
    if ($WPFlistView.SelectedItem.Share -eq 'True') {
        $WPFcheckBoxShare.IsChecked=$true
    } else {
        $WPFcheckBoxShare.IsChecked=$false
    }
    if ($WPFlistView.SelectedItem.ThisFolderOnly -eq 'True') {
        $WPFcheckBoxThisFolderOnly.IsChecked=$true
    } else {
        $WPFcheckBoxThisFolderOnly.IsChecked=$false
    }
    if ($WPFlistView.SelectedItem.UserFolder -eq 'True') {
        $WPFcheckBoxUserFolder.IsChecked=$true
    } else {
        $WPFcheckBoxUserFolder.IsChecked=$false
    }  

})

$WPFlistView.Add_MouseDoubleClick({

    $WPFlistView.SelectAll()

})

$WPFListView.Add_SourceUpdated({
    $WPFlistView.SourceUpdated
    $WPFlistView.Items.Refresh()
})

#endregion

#endregion


#===========================================================================
# Show the form
#===========================================================================

[void]$Form.ShowDialog()