##sample script showing usage of runspaces
#gets some basic info about all servers in domain

Function process-parallel($arr, $SB){
#Create the Runspace Pool
$RunspacePool = [RunspaceFactory]::CreateRunspacePool()

#Open the pool
$RunspacePool.Open()

#Create a runspace collection to hold the results
$RunspaceCollection = New-Object system.collections.arraylist    

foreach ($obj in $arr){        

    #Create a powershell object
    $Powershell = [PowerShell]::Create()

    #Specify the runspace to use
    $Powershell.RunspacePool = $RunspacePool

    #Add the script block to the processing  make sure this is piped to out-null or it will pollute your output
    $Powershell.AddScript($SB) | Out-Null

    #Add the current object as the parameter to the script block make sure this is piped to out-null or it will pollute your output
    $Powershell.AddArgument($obj) | Out-Null

    #create a temporary runspace object
    $RS = New-Object -TypeName PSObject -Property @{
        Runspace   = $PowerShell.BeginInvoke() 
        PowerShell = $PowerShell
    }

    #add the runspace to the collection, make sure this is piped to out-null or it will pollute your output with diag messages
    $RunspaceCollection.Add($RS) | Out-Null
}

#create a return array, this will be filled with the results of your query
$return = @()


#While loop waits for 
While($RunspaceCollection){
    #iterate through all runspaces in the collection, checking for results
    ForEach($Runspace in $RunspaceCollection.ToArray()){
        #check if the runspace is commpleted
        if($Runspace.RunSpace.IsCompleted -eq $true){
            #if the runspace is completed, end invoking and assign the results to the $return variable
            $return += $Runspace.Powershell.EndInvoke($Runspace.RunSpace)

            #dispose of the powershell object
            $Runspace.Powershell.dispose()

            $RunspaceCollection.Remove($Runspace)
        }
    }
}

return $return
}


$arr = get-adcomputer -filter *


$SB = {
param($Computer)

$return = "" | Select-Object OS,Driver,Disk,Share


$return.OS = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer.name
$return.Driver = Get-WmiObject Win32_SystemDriver -ComputerName $Computer.name
$return.Disk = Get-WmiObject Win32_logicalDisk -ComputerName $Computer.name
$return.share = Get-WmiObject win32_share -ComputerName $Computer.name

return $Return

}

process-parallel $arr $SB