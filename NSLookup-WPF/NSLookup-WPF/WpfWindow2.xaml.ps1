function Load-Xaml {
	[xml]$xaml = Get-Content -Path $PSScriptRoot\WpfWindow2.xaml
	$manager = New-Object System.Xml.XmlNamespaceManager -ArgumentList $xaml.NameTable
	$manager.AddNamespace("x", "http://schemas.microsoft.com/winfx/2006/xaml");
	$xamlReader = New-Object System.Xml.XmlNodeReader $xaml
	[Windows.Markup.XamlReader]::Load($xamlReader)
}

$window = Load-Xaml
$window.ShowDialog()