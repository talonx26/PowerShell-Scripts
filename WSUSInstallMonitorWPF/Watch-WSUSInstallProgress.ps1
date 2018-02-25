#Update path to match location for log files
$SCOMLogPath = "C:\wsus\test2\nk23208"

Get-EventSubscriber | % { Unregister-Event $_.SubscriptionId}
#$comp = @()
Remove-Variable computers -ErrorAction SilentlyContinue| out-null
function Get-LastLine($path)
{
   
    #$oldConsole = [console]::TreatControlCAsInput
    #[console]::TreatControlCAsInput = $true
    #write-host "enter"
    #write-host "computers : $($global:computers -isnot [System.Array])"
    if (!$global:computers) 
       { #write-host "No Global"
         $global:computers = @()
        }
    if ($global:computers -isnot [System.Array])
       {
         #Write-Host "no array"
         $global:computers = @()

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
    $Stat.progress = Get-Random -Maximum 100
    #write-host "Count $($comp.count)"
    if ($Comp.count -eq 0 )
      {  #Write-host "Zero"
         $comp += $stat}


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
            $comp  += $stat
          
       }
    
    #write-host "exit"
     $global:computers = $comp
    
     $global:status = $stat
     cls
    "{0,-20}{1,-20}{2,-25}{3,-50}" -f "Computer", "Action","Time","Description"  |write-host -BackgroundColor White   -ForegroundColor Black
    "----------------------------------------------------------------------------------------------------------" | write-host
    #write-host "test"
    $comp | %{
         $msg = $_
      switch($_.action)
        {
          # "{0,-20}{1,-20}{2,-20}{3,-10}{4,-50}" -f $msg.Computer, $msg.Action, $msg.time, "$($msg.percent)%", $msg.description 
          "Search" 
                   {  "{0,-20}{1,-20}{2,-25}{3,-50}" -f $msg.Computer, $msg.Action, $msg.time, $msg.description  |write-host -BackgroundColor Green  -ForegroundColor Black }
          "Install" 
                   {  "{0,-20}{1,-20}{2,-25}{3,-50}" -f $msg.Computer, $msg.Action, $msg.time,$msg.description  |write-host -BackgroundColor Green  -ForegroundColor Black }
          "Reboot"
                   {  "{0,-20}{1,-20}{2,-25}{3,-50}" -f $msg.Computer, $msg.Action, $msg.time,$msg.description  |write-host -BackgroundColor Magenta  -ForegroundColor Black}
          "Download"
                   {  "{0,-20}{1,-20}{2,-25}{3,-50}" -f $msg.Computer, $msg.Action, $msg.time,$msg.description  |write-host -BackgroundColor Yellow  -ForegroundColor Black}
          "RebootRequired"
                   {  "{0,-20}{1,-20}{2,-25}{3,-50}" -f $msg.Computer, $msg.Action, $msg.time,$msg.description  |write-host -BackgroundColor Red  -ForegroundColor Black}
           Default {  "{0,-20}{1,-20}{2,-25}{3,-50}" -f $msg.Computer, $msg.Action, $msg.time, $msg.description |write-host }
        }

       

    }
    
}

$fsw = New-Object System.IO.FileSystemWatcher $SCOMLogPath , "*.csv" 
$event = Register-ObjectEvent -InputObject $fsw -EventName "Changed" -action { Get-LastLine($event.sourceEventArgs.fullpath)}#;cls ;$computers | ft * -AutoSize | out-host}
