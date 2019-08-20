Function GetUpdateState {
    param([string[]]$kbnumber,
        [string]$wsusserver,
        [string]$port
    )
    $report = @()
    $updateServer = "143.219.39.47"
    [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
    #$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer($wsusserver,$False,80)
    $AdminProxy = New-Object -TypeName Microsoft.UpdateServices.Administration.AdminProxy
    $wsus = $AdminProxy.GetRemoteUpdateServerInstance($updateServer , $fALSE, 80)
    $CompSc = new-object Microsoft.UpdateServices.Administration.ComputerTargetScope
    $CompSc.IncludeDownstreamComputerTargets = $true
    $group = $wsus.GetComputerTargetGroups() | ? { $_.name -eq "PCS7 Production" } 
    $CompSc.ComputerTargetGroups.Add($group)
    $updateScope = new-object Microsoft.UpdateServices.Administration.UpdateScope; 
    $updateScope.UpdateApprovalActions = [Microsoft.UpdateServices.Administration.UpdateApprovalActions]::Install
    $Report = @()
    $kbCount = 0
    foreach ($kb in $kbnumber) {
        $kbCount++
        Write-Progress -Id 1 -Activity "Processing KBs" -Status "Gather data for $KB ($KBCount/$($kbnumber.count)" -PercentComplete (($kbCount / $kbnumber.count) * 100)
        #Loop against each KB number passed to the GetUpdateState function 
        $updates = $wsus.GetUpdates($updateScope) | ? { $_.Title -match $kb } #Getting every update where the title matches the $kbnumber
        $updateCount = 0
        foreach ($update in $updates) {
            $updateCount++
            Write-Progress -Id 2 -Activity "Processing Updates" -Status "Gather data for $($Update.KnowledgebaseArticles[0]) ($updateCount/$($updates.count)" -PercentComplete (($updateCount / $updates.count) * 100) -ParentId 1
            #Loop against the list of updates I stored in $updates in the previous step
            $Systems = $update.GetUpdateInstallationInfoPerComputerTarget($CompSc) | ? { $_.UpdateApprovalAction -eq "Install" -and $_.UpdateInstallationState -ne "NotApplicable" } 
            $systemCount = 0
            foreach ($system in $Systems) { 
                
                #for the current update
                #Getting the list of computer object IDs where this update is supposed to be installed ($_.UpdateApprovalAction -eq "Install")
                $Comp = $wsus.GetComputerTarget($_.ComputerTargetId)# using Computer object ID to retrieve the computer object properties (Name, IP address)
                $systemCount++
                Write-Progress -Id 3 -Activity "Processing Computers" -Status "Gather data for $($comp.Fulldomainname) ($systemCount/$($Systems.count)" -PercentComplete (($systemCount / $Systems.count) * 100) -ParentId 2
                If ($Report.count -eq 0) {
                    # $Report is blank
                    # Create and add object to collection
    
                    $computer = New-Object PSObject
                    $computer | Add-Member NoteProperty Name($Comp.FullDomainName)
                    $computer | add-member NoteProperty IPAddress($comp.IPAddress)
                    $computer | Add-Member NoteProperty OS($Comp.OSDescription)
                    $computer | add-member NoteProperty LastSyncTime($Comp.LastSyncTime)
                    $computer | add-member NoteProperty LastReportedStatusTime($Comp.LastReportedStatusTime)
                    if ($Comp.SyncsFromDownstreamServer) {   
                        $server = $Comp.GetParentServer().FullDomainName 
                   
                    }
                    else {
                        If (($wsus.name.split('.')).count -eq 4 ) {
                            #WSUS name is IP address.  Convert to name.
                            $Server = (Resolve-DnsName ($wsus.Name)).namehost
                
                        }
                        else {
                            $Server = $wsus.Name
                        }
                    }
                    $computer | Add-Member NoteProperty WSUSServer($Server)
                    $computer | Add-Member NoteProperty PatchInstalled($false)
                    Foreach ( $k in $kbnumber) {
                        $computer | Add-Member NoteProperty $k($null)
                        $computer | Add-Member NoteProperty "$($K)Approval"($null)
    
    
                    }
                    $Report += $computer
    
                }
                If ($Report.Name.Indexof($comp.FullDomainName) -eq -1) {
                    # Computer is not in the collection
                    # Create and add object to collection
    
                    $computer = New-Object PSObject
                    $computer | Add-Member NoteProperty Name($Comp.FullDomainName)
                    $computer | add-member NoteProperty IPAddress($comp.IPAddress)
                    $computer | Add-Member NoteProperty OS($Comp.OSDescription)
                    $computer | add-member NoteProperty LastSyncTime($Comp.LastSyncTime)
                    $computer | add-member NoteProperty LastReportedStatusTime($Comp.LastReportedStatusTime)
                    if ($Comp.SyncsFromDownstreamServer) {   
                        $server = $Comp.GetParentServer().FullDomainName 
                   
                    }
                    else {
                        If (($wsus.name.split('.')).count -eq 4 ) {
                            #WSUS name is IP address.  Convert to name.
                            $Server = (Resolve-DnsName ($wsus.Name)).namehost
                
                        }
                        else {
                            $Server = $wsus.Name
                        }
                    }
                    $computer | Add-Member NoteProperty WSUSServer($Server)
                    $computer | Add-Member NoteProperty PatchInstalled($false)
                    Foreach ( $k in $kbnumber) {
                        $computer | Add-Member NoteProperty $k($null)
                        $computer | Add-Member NoteProperty "$($K)Approval"($null)
    
    
                    }
                    $Report += $computer
    
                }
    
                $index = $Report.Name.Indexof($comp.FullDomainName)
             
                #  If Update is installed change PatchInstalled Flag to True.  To show that at least one of the requested patches has been installed.
                if ($system.UpdateInstallationState -eq "Installed") { 
               
                    $Report[$index].PatchInstalled = $true
                  
                }
             
               
                $Report[$index]."KB$($Update.KnowledgebaseArticles[0])" = $system.UpdateInstallationState
                $Report[$index]."KB$($Update.KnowledgebaseArticles[0])Approval" = $system.UpdateApprovalAction
               
                
            }
            Write-Progress -id 3 -Completed
        }
    }
    
    #Filtering the report
    #$report | ?{$_.UpdateInstallationStatus -ne 'NotApplicable' -and $_.UpdateInstallationStatus -ne 'Unknown'} | Export-Csv C:\Temp\unpatched.csv -Force -en ASCII -NoTypeInformation
    
    $Report | Export-Csv C:\Temp\unpatched1.csv -Force -en ASCII -NoTypeInformation
    
} 