function Load-Xaml {
	[xml]$xaml = @" 
	<Window

  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"

  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"

  Title="MainWindow" Height="350" Width="525">

    <Grid>

        <Label Content="Label" HorizontalAlignment="Left" Margin="68,38,0,0" VerticalAlignment="Top" Width="197"/>

        <Button Content="Button" HorizontalAlignment="Left" Margin="307,41,0,0" VerticalAlignment="Top" Width="75"/>

    </Grid>

</Window>
"@
	#Get-Content -Path $PSScriptRoot\WpfWindow2.xaml
	$manager = New-Object System.Xml.XmlNamespaceManager -ArgumentList $xaml.NameTable
	$manager.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml");
	$xamlReader = New-Object System.Xml.XmlNodeReader $xaml
	[Windows.Markup.XamlReader]::Load($xamlReader)
}

$window = Load-Xaml
$window.ShowDialog()