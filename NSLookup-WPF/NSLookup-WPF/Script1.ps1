#
# Script1.ps1
#
$syncHash.btnStart.add_Click({
1..10 | % { 
if ($ips -eq $null) { $ips = @()}
$IP = "" | Select Target, HostName, IPAddress
$ip.Target = "test1"
$ip.HostName = "Test2"
$ip.IPAddress =  "127.0.0.1"
$ips += $ip



$synchash.dataGrid.Dispatcher.Invoke([action]{$synchash.dataGrid.ItemsSource = $ips},"Normal")
sleep -Milliseconds 1000
}
})


$synchash.txtInput.add_LostFocus({get-syc)
