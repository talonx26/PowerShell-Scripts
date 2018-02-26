#----------------------------------------------
#region Application Functions
#----------------------------------------------
 
function OnApplicationLoad {
    #Note: This function is not called in Projects
    #Note: This function runs before the form is created
    #Note: To get the script directory in the Packager use: Split-Path $hostinvocation.MyCommand.path
    #Note: To get the console output in the Packager (Windows Mode) use: $ConsoleOutput (Type: System.Collections.ArrayList)
    #Important: Form controls cannot be accessed in this function
    #TODO: Add modules and custom code to validate the application load
     
    return $true #return true for success or false for failure
}
 
function OnApplicationExit {
    #Note: This function is not called in Projects
    #Note: This function runs after the form is closed
    #TODO: Add custom code to clean up and unload modules when the application exits
     
    $script:ExitCode = 0 #Set the exit code for the Packager
}
 
#endregion Application Functions
 
#----------------------------------------------
# Generated Form Function
#----------------------------------------------
function Call-MD-RunningProcesses_psf {
 
    #----------------------------------------------
    #region Import the Assemblies
    #----------------------------------------------
    [void][reflection.assembly]::Load('mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
    [void][reflection.assembly]::Load('System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
    [void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
    [void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
    [void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
    [void][reflection.assembly]::Load('System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
    [void][reflection.assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
    [void][reflection.assembly]::Load('System.Core, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
    [void][reflection.assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
    #endregion Import Assemblies
 
    #----------------------------------------------
    #region Generated Form Objects
    #----------------------------------------------
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $frmMain = New-Object 'System.Windows.Forms.Form'
    $timerCheckRunSpaceFinished = New-Object 'System.Windows.Forms.Timer'
    $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
    #endregion Generated Form Objects
 
    #----------------------------------------------
    # User Generated Script
    #----------------------------------------------
     
     
     
     
     
     
     
    #Dynamically create a datagridview within a synchronized hash table so that it is available in another workflow/pool
     
    $hashDGProcs = [hashtable]::Synchronized(@{ })
    $hashDGProcs.DGRunningProcs = New-Object System.Windows.Forms.DataGridView
    $hashDGProcs.dgRunningProcs.AutoSizeColumnsMode = "None"
    $hashDGProcs.dgRunningProcs.AutoSizeRowsMode = "None"
    $hashDGProcs.dgRunningProcs.Dock = "None"
    $hashDGProcs.dgRunningProcs.Location = New-Object System.Drawing.Point(18, 17)
    $hashDGProcs.dgRunningProcs.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
    $hashDGProcs.dgRunningProcs.ScrollBars = 'Both'
    $hashDGProcs.dgRunningProcs.ScrollBars = 'Both'
    $hashDGProcs.dgRunningProcs.Size = New-Object Drawing.Size(622, 309)
    $hashDGProcs.dgRunningProcs.ColumnCount = 4
    $hashDGProcs.dgRunningProcs.Columns[0].HeaderText = "Process ID"
    $hashDGProcs.dgRunningProcs.Columns[1].HeaderText = "Name"
    $hashDGProcs.dgRunningProcs.Columns[2].HeaderText = "Username"
    $hashDGProcs.dgRunningProcs.Columns[3].HeaderText = "CreationDate"
    $hashDGProcs.dgRunningProcs.SelectionMode = 'FullRowSelect'
     
    #Get a hostname
    If (!($strComputer)) { $strComputer = Read-Host "Enter a computer name" }
     
     
    $frmMain_Load= {
        #add the gridview to our form
        $frmMain.Controls.Add($hashDGProcs.dgRunningProcs)
         
        #Call our function to populate the datagrid
        fnProcScriptBlocks
    }
     
     
    function fnProcScriptBlocks
    {
     
        #Define the code we want to run in another thread
            $sbProcScript = {
                Param ($hashDGProc, $strComputer)
                #Get an array or collection of the Current Running Processes...
                #Wait a minute...that doesn't look like a get-wmi command?
                #I am utilizing psremoting here so my query will run locally on the remote computer...therefore it runs faster
                $arrProc = Invoke-Command -ComputerName $strComputer -ScriptBlock{
                 
                    #Define function within the invoke-command to convert the start time of the process to date time format
                    Function WMIDateStringToDate($crdate)
                    {
                        If ($crdate -match ".\d*-\d*")
                        {
                            $crdate = $crdate -replace $matches[0], " "
                            $idate = [System.Int64]$crdate
                            $date = [DateTime]::ParseExact($idate, 'yyyyMMddHHmmss', $null)
                            return $date
                        }
                    }
                    #Get the processes from wmi and add a noteproperty for owner and date...then select the properties we want and return it up through the invoke-command
                    gwmi -class win32_process | ForEach             {
                        #Add the owner as a note property
                        Try { $objProcess = Add-Member -InputObject $_ -MemberType NoteProperty -Name UserName -Value ($_.GetOwner().User) -PassThru }
                        Catch [Exception] { $objProcess = Add-Member -InputObject $_ -MemberType NoteProperty -Name UserName -Value "" }
                        Finally { }
                         
                        #Add the reformated date as a note property
                        Try { Add-Member -InputObject $objProcess -MemberType NoteProperty -Name refCreationDate -Value (WMIDateStringToDate($_.CreationDate)) -PassThru }
                        Catch [Exception] { Add-Member -InputObject $objProcess -MemberType NoteProperty -Name refCreationDate -Value "" }
                        Finally { }
                         
                    } | Select-Object ProcessID, Name, Username, refCreationDate
                     
                } -ErrorAction Stop
             
            #At this point, $arrProc is an array of processes
                        #If there is no date in our datgrid view...
                If ($hashDGProc.dgRunningProcs.Rows.Count -le 1)
            {
                 
                    #We must call invoke when we modify the datagridview, because this code will be executed in another runspace
                    $hashDGProc.DGRunningProcs.Invoke([action]{
                        ForEach ($objProcess in $arrProc)
                        {
                            #Add the process to the gridview
                            $dgIndex = $hashDGProc.dgRunningProcs.Rows.Add($objProcess.ProcessID, $objProcess.Name, $objProcess.Username, $objProcess.refCreationDate)
                        }
                })
                 
                }
                else #more than 1 row in the DG view
                {
                    #find old procs
                    $arrAgedProcs = @()
                 
                    #Check each existing row in the datagridview
                    ForEach ($dgrow in $hashDGProc.dgRunningProcs.Rows)
                    {
                        $objProcStillRunning = $false
                        #Check if the process is still running
                        $objProcStillRunning = $arrProc | Where-Object -Property ProcessID -eq $dgrow.Cells[0].Value
                        If (!($objProcStillRunning))
                        {
                            #Remove the row if the process is no longer running
                            $hashDGProc.DGRunningProcs.Invoke([action]{
                                $hashDGProc.DGRunningProcs.Rows.Remove($dgrow)
                            })
                        }
                        else #Process is still running
                        {
                            #Update any property value changes here
                            $arrAgedProcs += $dgrow.Cells[0].Value
                        }
                    }
                     
                    #Add Missing Procs
                     
                    #Once again, we need to invoke as we are running this from another runspace
                    $hashDGProc.DGRunningProcs.Invoke([action]{
                        ForEach ($objNewProc in ($arrProc | Where-Object -Property ProcessID -NotIn $arrAgedProcs))
                        {
                            $dgIndex = $hashDGProc.dgRunningProcs.Rows.Add($objNewProc.ProcessID, $objNewProc.Name, $objNewProc.Username, $objNewProc.refCreationDate)
                        }
                    })
                     
                }
            }
         
            #Here I have defined a script block that I will execute when the runspace 
        $sbProcComplete = {
            #Rerun the function
                    Start-Sleep -Milliseconds 100 #Give time to delete the old runspace so we don't trip over ourselves
            fnProcScriptBlocks 
            }
         
        #Here we get into the runspace/multi-threading
         
            #First make sure we have a runspace pool....not really necessary, but I like the extra control...See Boe's blog posts for more on this
            If (!($runspacepool))
            {
                $Script:runspaces = New-Object System.Collections.ArrayList
                $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
                $runspacepool = [runspacefactory]::CreateRunspacePool(1, 10, $sessionstate, $host)
                $runspacepool.Open()
            }
         
            #Create the new runspace
            $powershellRunSpace = [powershell]::Create()
         
            #Add our script and arguments...note that we are passing our hashtable which contains our datagridview as well as our script block
            $powershellRunSpace.AddScript($sbProcScript).AddArgument($hashDGProcs).AddArgument($strComputer)
            $powershellRunSpace.RunspacePool = $runspacepool
         
            #Create a custom object to hold the properties of our runspace instance     
            $InstRunSpace = "" | Select-Object name, powershell, runspace, Computer, CompletedScript
            $InstRunSpace.Name = (Get-Random)
             
            $InstRunSpace.Computer = $strComputer
            $instRunSpace.Powershell = $powershellRunSpace
            $InstRunSpace.CompletedScript = $sbProcComplete
            $InstRunSpace.RunSpace = $powershellRunSpace.BeginInvoke() #This line kicks off the runspace..which runs our scriptblock
            $runspaces.Add($InstRunSpace) | Out-Null #add the runspace instance to the array list of runspaces
             
            #Not required unless you want to run another script after the script block has completed
            If (!($timerCheckRunSpaceFinished.Enabled))
            {
            $timerCheckRunSpaceFinished.Enabled = $true
            $timerCheckRunSpaceFinished.Start()
            }
             
             
        }
     
    #This function will run from our timer ... it simply checks if a runspace has completed and then executes any defined completedscripts
    Function fnGet-RunspaceData
    {
        Foreach ($runspace in $runspaces)
        {
            If ($runspace.Runspace.isCompleted) #If the runspace is done
            {
                Try
                {
                    $runspace.powershell.EndInvoke($runspace.Runspace) #Close the runspace
                    $runspace.powershell.dispose() #put the runspace in the garbage
                }
                Catch [Exception]{
                    #Do nothing....most of the time, you don't need the endinvoke, but if you try to run the endinvoke that isn't running it throws an exception
                }
                $runspace.Runspace = $null #nullify it
                $runspace.powershell = $null
                If ($runspace.CompletedScript) #If we defined a completedscript property....run it
                {
                    &amp; $runspace.completedScript
                }
            }
        }
        #Clean out unused runspace jobs
        $temphash = $runspaces.clone()
        $temphash | Where-Object -Property Runspace -eq $Null | ForEach-Object{ $Runspaces.remove($_) }
    }
     
    $timerCheckRunSpaceFinished_Tick={
        #Here I am checking if the runspace has completed...each time the timer ticks
        fnGet-RunspaceData
    }
     
    # --End User Generated Script--
    #----------------------------------------------
    #region Generated Events
    #----------------------------------------------
     
    $Form_StateCorrection_Load=
    {
        #Correct the initial state of the form to prevent the .Net maximized form issue
        $frmMain.WindowState = $InitialFormWindowState
    }
     
    $Form_Cleanup_FormClosed=
    {
        #Remove all event handlers from the controls
        try
        {
            $frmMain.remove_Load($frmMain_Load)
            $timerCheckRunSpaceFinished.remove_Tick($timerCheckRunSpaceFinished_Tick)
            $frmMain.remove_Load($Form_StateCorrection_Load)
            $frmMain.remove_FormClosed($Form_Cleanup_FormClosed)
        }
        catch [Exception]
        { }
    }
    #endregion Generated Events
 
    #----------------------------------------------
    #region Generated Form Code
    #----------------------------------------------
    #
    # frmMain
    #
    $frmMain.ClientSize = '667, 349'
    $frmMain.Name = "frmMain"
    $frmMain.Text = "Form"
    $frmMain.add_Load($frmMain_Load)
    #
    # timerCheckRunSpaceFinished
    #
    $timerCheckRunSpaceFinished.Interval = 1000
    $timerCheckRunSpaceFinished.add_Tick($timerCheckRunSpaceFinished_Tick)
    #endregion Generated Form Code
 
    #----------------------------------------------
 
    #Save the initial state of the form
    $InitialFormWindowState = $frmMain.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $frmMain.add_Load($Form_StateCorrection_Load)
    #Clean up the control events
    $frmMain.add_FormClosed($Form_Cleanup_FormClosed)
    #Show the Form
    return $frmMain.ShowDialog()
 
} #End Function
 
#Call OnApplicationLoad to initialize
if((OnApplicationLoad) -eq $true)
{
    #Call the form
    Call-MD-RunningProcesses_psf | Out-Null
    #Perform cleanup
    OnApplicationExit
}