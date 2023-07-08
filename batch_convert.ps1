// CreateMPTTemplate:
//          Function to generate a unique template file for the conversion of an App-V package into an MSIX one.
function CreateMPTTemplate($conversionParam, $refId,  $virtualMachine, $workingDirectory)
{
    # create template file for this conversion
    $templateFilePath = [System.IO.Path]::Combine($workingDirectory, "MPT_Templates", "MsixPackagingToolTemplate_Ref$($refId).xml")
    $conversionMachine = "<mptv2:RemoteMachine ComputerName=""$($virtualMachine.Name)"" Username=""$($virtualMachine.Credential.UserName)"" />"

    $saveFolder = [System.IO.Path]::Combine($workingDirectory, "MSIX")
    $xmlContent = @"
<MsixPackagingToolTemplate
    xmlns="http://schemas.microsoft.com/appx/msixpackagingtool/template/2018"
    xmlns:mptv2="http://schemas.microsoft.com/msix/msixpackagingtool/template/1904">
<Installer Path="$($conversionParam.InstallerPath)" Arguments="$($conversionParam.InstallerArguments)" />
$conversionMachine
<SaveLocation PackagePath="$saveFolder" />
<PackageInformation
    PackageName="$($conversionParam.PackageName)"
    PackageDisplayName="$($conversionParam.PackageDisplayName)"
    PublisherName="$($conversionParam.PublisherName)"
    PublisherDisplayName="$($conversionParam.PublisherDisplayName)"
    Version="$($conversionParam.PackageVersion)">
</PackageInformation>
</MsixPackagingToolTemplate>
"@
    Set-Content -Value $xmlContent -Path $templateFilePath
    $templateFilePath
}

function RunRemotePreInstaller($virtualMachineName, $Installer)
{
    # INPUTS:
    #    $virtualMachineName:  Name of the Windows machine to remote to
    #    $Installer:  Full path to a powershell ps1 file.
    # PURPOSE: The installer file, along with all files in the same folder, will be copied to the named machine
    #          into a temporary folder and the ps1 file will be run on that machine from that temp location.
    Write-Output "---------------- Starting PreInstaller $($Installer) Remotely..." 
    try
    {
    $leaffile = Split-Path -Path $Installer -Leaf
    $fullfolder = Split-Path -Path $Installer -Parent
    $leaffolder = Split-Path -Path $fullfolder -Leaf
    $localpath = "C:\TempInstaller"
    $sharename = "TempInstaller"
    $shareCommand = {
       if (!(Test-Path $args[0]))
       {
           Write-Output "     create directory $($args[0])" 
           New-Item -ItemType Directory -Path $args[0] 
           New-Item -ItemType Directory -Path "$($args[0])\xxx"
       }
       if (!(Get-SmbShare -name $args[1] -ErrorAction SilentlyContinue))
       {
           Write-Output "     create share $($args[0])"
           New-SmbShare -Name $args[1] -Path $args[0] -Description "Temp Share for PreInstaller" -FullAccess Everyone
       }
       Write-Output "     check share"
       Get-SmbShare -name $args[1]
    }
    $ArgumentArr =  $localpath, $sharename 
    
    Write-Output "     create dir and share remotely" 
    Invoke-Command -ComputerName $virtualMachineName -ScriptBlock $shareCommand -ArgumentList $ArgumentArr 
    
    Write-Output "     copy item $($fullfolder) to \\$($virtualMachineName)\$($sharename)"  
    Copy-Item -Recurse -Path "$($fullfolder)" -Destination "\\$($virtualMachineName)\$($sharename)"  
    
    write-output "     confirm contents at \\$($virtualMachineName)\$($sharename)"  
    dir -Recurse "\\$($virtualMachineName)\$($sharename)"  
    
    Write-Output "     Invoke preinstaller $($localpath)\$($leaffolder)\$($leaffile) remotely"  
    $cmd2run = @{
       ComputerName = $virtualMachineName
       ScriptBlock = { Invoke-Expression -Command "Set-ExecutionPolicy -ExecutionPolicy Unrestricted"
                       Invoke-Expression -Command "$($args)"  }
       ArgumentList = "$($localpath)\$($leaffolder)\$($leaffile)"
    }
    Invoke-Command @cmd2run  
    }
    catch
    {
        Write-Output "***EXCEPTION in RunRemotePreInstaller $($_)"
    }
    Write-output "Is installed check:"
    invoke-Command -ComputerName $virtualMachineName {get-ChildItem "C:\Program Files" }

    Write-Output "---------------- PreInstaller done." 
}

$FunctionRunRemotePreInstaller = @"

Function RunRemotePreInstaller
{
    $(Get-Command RunRemotePreInstaller | Select -expand Definition)
}

"@

function SetupOutputFolders($workingDirectory, $CleanupOutputFolderAtStart)
{
    Write-Host "Cleanup = $($CleanupOutputFolderAtStart)"
    #Cleanup previous run
    if ($CleanupOutputFolderAtStart -eq $true)
    {
        Write-Host 'Cleanups:'
        get-job | Stop-Job |Remove-Job
        if (Test-Path  ($workingDirectory))
        {
            if (Test-Path  ([System.IO.Path]::Combine($workingDirectory, "MPT_Templates")))
            {
                Write-Host '   Cleanup: MPT_Templates'
                Remove-item ([System.IO.Path]::Combine($workingDirectory, "MPT_Templates")) -recurse
            }
            if (Test-Path  ([System.IO.Path]::Combine($workingDirectory, "MSIX")))
            {
                Write-Host '   Cleanup: MSIX'
                Remove-item ([System.IO.Path]::Combine($workingDirectory, "MSIX")) -recurse
            }
            if (Test-Path  ([System.IO.Path]::Combine($workingDirectory, "LOGS")))
            {
                Write-Host '   Cleanup: LOGS'
                Remove-item ([System.IO.Path]::Combine($workingDirectory, "Logs")) -recurse
            }
        }
        else
        {
            Write-Host "Create: $($workingDirectory)"
            New-Item -Force -Type Directory ($workingDirectory)
        }

        #Set up this run
        Write-Host 'Create Folder Strucure'
        New-Item -Force -Type Directory ([System.IO.Path]::Combine($workingDirectory, "MPT_Templates"))
        New-Item -Force -Type Directory ([System.IO.Path]::Combine($workingDirectory, "MSIX"))
        New-Item -Force -Type Directory ([System.IO.Path]::Combine($workingDirectory, "LOGS"))
    }
    else
    {
        Write-Host 'Cleanups skipped.'
    }
}

function RunConversionJobs($AppConversionParameters, $virtualMachines, $workingDirectory, $retryBad, $CleanupOutputFolderAtStart, $skipFirst)
{
    Write-Host 'RunConversionJobs'

    SetupOutputFolders $workingDirectory $CleanupOutputFolderAtStart


    $logfolder  = ([System.IO.Path]::Combine($workingDirectory, "LOGS"))

    # Ensure that the dynamic memory parameters are reset
    foreach ($conv in $AppConversionParameters )
    {
        JobSetStartedValue   $conv $false
        JobSetCompletedValue $conv $false
    }
  

    # create list of the indices of $AppConversionParameters that haven't started running yet
    $remainingConversionIndexes = @()
    $AppConversionParameters | Foreach-Object { $i = $skipFirst } { $remainingConversionIndexes += ($i++) }

    $failedConversionIndexes = New-Object -TypeName "System.Collections.ArrayList"
    foreach ($fci in $failedConversionIndexes)
    {
        $failedConversionIndexes.Remove($fci)
    }
    
    # Next schedule jobs on virtual machines which can be checkpointed/re-used
    # keep a mapping of VMs and the current job they're running, initialized ot null
    $virtMachinesArray = New-Object -TypeName "System.Collections.ArrayList"
    foreach ($vm in $virtualmachines)
    {
        if ($vm.enabled -eq $true)
        { 
            $virtMachine = New-Object -TypeName PSObject
            $virtMachine | Add-Member -NotePropertyName npVmCfgObj -NotePropertyValue $vm
            $virtMachine | Add-Member -NotePropertyName npVmGetObj -NotePropertyValue (get-vm -ComputerName $vm.Host -Name $vm.Name)
            $virtMachine | Add-Member -NotePropertyName npRefId -NotePropertyValue -1
            $virtMachine | Add-Member -NotePropertyName npInUse -NotePropertyValue $false
            $virtMachine | Add-Member -NotePropertyName npJobObj -NotePropertyValue $nul
            $virtMachine | Add-Member -NotePropertyName npAppName -NotePropertyValue ""    
            $virtMachine | Add-Member -NotePropertyName npPreInstall -NotePropertyValue $nul 
            $virtMachine | Add-Member -NotePropertyName npErrorCount -NotePropertyValue 0  
            $virtMachine | Add-Member -NotePropertyName npDisabled -NotePropertyValue $false  
            $virtMachine | Add-Member -NotePropertyName npAppConfiguration -NotePropertyValue $nul  
            $virtMachinesArray.Add($virtMachine)    > $xxx ## $xxx is just to avoid unwanted console output
        }
    }

    # Use a semaphore to signal when a machine is available. Note we need a global semaphore as the jobs are each started in a different powershell process
    # Make sure prior runs are cleared out first
    $semaphore = New-Object -TypeName System.Threading.Semaphore -ArgumentList @($virtMachinesArray.Count, $virtMachinesArray.Count, "Global\MPTBatchConversion")
    $semaphore.Close()
    $semaphore.Dispose()

    $semaphore = New-Object -TypeName System.Threading.Semaphore -ArgumentList @($virtMachinesArray.Count, $virtMachinesArray.Count, "Global\MPTBatchConversion")
    Write-Host ""

    while ($semaphore.WaitOne(-1))
    {
        if ($remainingConversionIndexes.Count -gt 0)
        {
            # select a job to run 
            Write-Host "Determining next job to run..."
            $conversionParam = $AppConversionParameters[$remainingConversionIndexes[0]]
            if ( $conversionParam.Enabled)
            {
                # select a VM to run it on. Retry a few times due to race between semaphore signaling and process completion status
                $vm = $nul
                while (-not $vm) { $vm = $virtMachinesArray | where { $_.npInUse -eq $false -and $_.npDisabled -eq $false } | Select-Object -First 1 }
               

                # Capture the ref index and update list of remaining conversions to run
                $refId = $remainingConversionIndexes[0]
                $remainingConversionIndexes = $remainingConversionIndexes | where { $_ -ne $remainingConversionIndexes[0] }
                Write-Host "Dequeue for conversion Ref $($refId) for app $($conversionParam.PackageName) on VM $($vm.npVmGetObj.Name)." -Foreground Cyan
                
                $vm.npRefId = $refId
                $vm.npAppConfiguration = $conversionParam
                $vm.npAppName = $conversionParam.PackageName
                if ($conversionParam.PreInstallerArguments -ne $nul)
                {
                    $vm.npPreInstall = $conversionParam.PreInstallerArguments
                }
                else
                {
                    $vm.npPreInstall = $nul
                }
                $vm.npInUse = $true
                $conversionParam.Started = $true
                $conversionParam.Completed = $false

                $templateFilePath = CreateMPTTemplate $conversionParam $refId $vm.npVmCfgObj $workingDirectory 
                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($vm.npVmCfgObj.Credential.Password)
                $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

                $useLogFile = "$($logfolder)\$($conversionParam.PackageName)_$(Get-Date -format FileDateTime).txt"
                Write-Host "Start Job  $($vm.npVmGetObj.ComputerName)  $($conversionParam.PackageName) " 
                Write-Output "Start Job  $($vm.npVmGetObj.ComputerName)  $($conversionParam.PackageName) " > $useLogFile
                $jobObject = start-job -ScriptBlock {  
                    param($refId, $vMachine,  $machinePassword, $templateFilePath, $initialSnapshotName,$logFile, $funcs)
                    
                    Write-Output "Starting with params:  $($refId), $($vMachine.npVmGetObj.ComputerName)/$($vMachine.npVmGetObj.Name), $($vmsCount), $($machinePassword), $($templateFilePath), $($initialSnapshotName), $($logFile)" >> $logFile

                    try
                    {
                        Invoke-Expression $funcs
                        Write-Output "debug: be4 get snapshot" >> $logFile
                        $snap = Get-VMSnapshot -Name $initialSnapshotName -VMName $vMachine.npVmCfgObj.Name -ComputerName $vMachine.npVmCfgObj.Host -ErrorAction Continue
                        Write-Output "debug: after get snapshot" >> $logFile
                        if ( $snap)
                        {
                            Write-Output "Reverting VM snapshot for  $($vMachine.npVmCfgObj.Host) / $($vMachine.npVmCfgObj.Name): $($initialSnapshotName)" >> $logFile
                            Restore-VMSnapshot -ComputerName $vMachine.npVmCfgObj.Host -VMName $vMachine.npVmCfgObj.Name -Name $initialSnapshotName -Confirm:$false
                            Write-Output "debug: after revert" >> $logFile
                            ####we probably don't need to replace the vm object, but once had an issue so let's be sure...
                            Start-Sleep 2
                            $vMachine.npVmGetObj = (get-vm -ComputerName $vMachine.npVmCfgObj.Host -Name $vMachine.npVmCfgObj.Name)
                            Set-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name -Notes "Preparing $($vMachine.npAppName)"
                            
                            if ( $vMachine.npVmGetObj.state -eq 'Off' -or $vMachine.npVmGetObj.state -eq 'Saved' )
                            {
                                Write-Output "Starting VM" >> $logFile
                                Start-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name 
                                #Write-Output "after Starting VM" >> $logFile
                                $limit = 60
                                while ($vMachine.npVmGetObj.state -ne 'Running')
                                {
                                    Start-Sleep 5
                                    $limit = $limit - 1
                                    if ($limit -eq 0)
                                    {
                                        Write-Output "TIMEOUT while starting restored checkpoint' state=$($vMachine.npVmGetObj.state)." >> $logFile
                                        $vMachine.npErrorCount = $vMachine.npErrorCount + 1
                                        if ($vMachine.npVmGetObj.state -ne 'Off')
                                        {
                                            Write-Host "Debug: Stop VM $( $vMachine.npVmGetObj.Name)" >> $logFile
                                            Stop-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name -TurnOff -Force -ErrorAction SilentlyContinue
                                        }
                                        break;
                                    }
                               }
                            }
                            else
                            {
                                Write-Output "Debug: state is $($vMachine.npVmGetObj.state)" >> $logFile
                            }
                        }
                        else
                        {
                            Write-Output "Get-VMSnapshot error" >> $logFile
                        }

                        $waiting = $true
                        $waitcount = 0
                        ## Let VM Settle a little.  At times the VMs get a little busy thanks to MS and more time seems to work better.
                        while ($waiting)
                        {
                            if ( $vMachine.npVmGetObj.state -eq 'Running' -and $vMachine.npVmGetObj.upTime.TotalSeconds -gt 120  )
                            {
                                $waiting = $false

                                if ($vMachine.npPreInstall -ne $nul)
                                {
                                    Write-output "Before RunRemotePreInstaller"
                                    RunRemotePreInstaller $vMachine.npVmGetObj.Name $vMachine.npPreInstall >> $logFile
                                    Write-output "After RunRemotePreInstaller"
                                }


                                Set-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name -Notes "Packaging $($vMachine.npAppName)"
                                Write-Output "" >> $logFile
                                Write-Output "==========================Starting package..." >> $logFile
                                MsixPackagingTool.exe create-package --template $templateFilePath --machinePassword $machinePassword -v >> $logFile
                                Write-Output "==========================Packaging tool done." >> $logFile
                                Write-Output "" >> $logFile
                            }
                            $waitcount += 1
                            if ($waitcount -gt 360)
                            {
                                $waiting = $false
                                Write-Output "Timeout waiting for OS to start" >> $logFile
                                $vMachine.npErrorCount = $vMachine.npErrorCount + 1
                                if ($vMachine.npVmGetObj.state -ne 'Off')
                                {
                                    Write-Host "Debug: Stop VM $( $vMachine.npVmGetObj.Name)" >> $logFile
                                    Stop-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name -TurnOff -Force -ErrorAction SilentlyContinue
                                }
                            }
                            start-sleep 1
                        }
                        Write-Output "Debug: job ready for finalizing." >> $logFile
                    }
                    catch
                    {
                        Write-Output "***EXECPTION***"
                        Write-Output "***EXECPTION***" >> $logFile
                        Write-Output $_
                        Write-Output $_ >> $logfile
                    }
                    finally
                    {
                        Write-Output "Finalizing." >> $logFile
                        Stop-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name -TurnOff -Force -ErrorAction SilentlyContinue

                        #Read-Host -Prompt 'Press any key to exit this window '
                        Write-Output "Complete." >> $logFile
                    }

                }  -ArgumentList $refId,  $vm,  $password, $templateFilePath, $vm.npVmCfgObj.initialSnapshotName, $useLogFile, "$($FunctionRunRemotePreInstaller)"
                $vm.npJobObj = $jobObject
                write-host "Ref$($refId): job is named $($jobObject.Name)"
                start-sleep 10
            }
            else {
                $refId = $remainingConversionIndexes[0]
                $remainingConversionIndexes = $remainingConversionIndexes | where { $_ -ne $remainingConversionIndexes[0] }
                Write-Host "Ref $($refId): $($conversionParam.PackageName) skipped by request." -ForegroundColor gray
                $semaphore.Release()
            }
        }
        else
        {
            $semaphore.Release()
            break;
        }


        $tempFiledConversionIndexes = WaitForFreeVM $virtMachinesArray $workingDirectory
        foreach ($tempFail in $tempFiledConversionIndexes)
        {
            if ($tempFail -gt 0 -and $AppConversionParameters[$tempFail].Enabled -eq $true -and $AppConversionParameters[$tempFail].Completed -eq $false)
            {
                Write-host "Fail Ref $($tempFail)" -Foreground yellow
                $failedConversionIndexes.Add($tempFail) > $xxx
            }
        }
        Write-host "One or more VMs are available for scheduling..."        
    }

    Write-Host "Finished scheduling all jobs, wait for final jobs to complete."
    #$virtualMachines | foreach-object { if ($vmsCurrentJobNameMap[$_.Name]) { $vmsCurrentJobNameMap[$_.Name].WaitForExit() } }
    $countInUse = $virtMachinesArray.Count
    $firstposttime = $true
    $countInUse = CountEnabledInuseVMs $virtMachinesArray
    Write-Host "There are $($countInUse) VMs still running" 
    while ($countInUse -gt 0)
    {
        $tempFiledConversionIndexes = WaitForFreeVM $virtMachinesArray $workingDirectory
        foreach ($tempFail in $tempFiledConversionIndexes)
        {
            if ($tempFail -gt 0 -and $AppConversionParameters[$tempFail].Enabled -eq $true -and $AppConversionParameters[$tempFail].Completed -eq $false)
            {
                Write-host "Fail Ref $($tempFail)" -Foreground yellow
                $failedConversionIndexes.Add($tempFail) > $xxx
            }
        }
        $countInUse = CountEnabledInuseVMs $virtMachinesArray
        if ($firstposttime -eq $true)
        {
            Write-Host "There are now $($countInUse) VMs still running" 
            $firstposttime = $false            
        }
        Sleep(5)
    }

    $semaphore.Dispose()
    Write-Host "Finished running initial attempt on all packaging jobs."

    if ($retryBad)
    {
        #Get the best VM today
        $redoVirtMachinesArray = New-Object -TypeName "System.Collections.ArrayList"
        $bestvmachine = $nul
        foreach ($vmach in $virtMachinesArray)
        {
            Write-Host "$($vmach.npVmGetObj.Name) aka $($vmach.npVmCfgObj.Name) dis=$($vmach.npDisabled) err=$($vmach.npErrorCount)"
            if ($vmach.npDisabled -eq $false)
            {
                if ($bestvmachine -eq $nul -or $vmach.npErrorCount -lt $bestvmachine.npErrorCount)
                {
                    Write-Host "Redo possibly uses $($vmach.npVmCfgObj.Name)"
                    $bestvmachine = $vmach
                }
            }
        }
        $redoVirtMachinesArray.Add($bestmachine) > $xxx

        Write-host "There are $($failedConversionIndexes.Count) packages for redo" -ForegroundColor Cyan
        foreach ($failedConversionIndex in $failedConversionIndexes)
        {
            $failedConversionParameter = $AppConversionParameters[$failedConversionIndex]
            if ($failedConversionParameter.Completed -eq $false)
            {
                Write-Host "Redo[$($failedConversionIndex)] for $($failedConversionParameter.PackageName) on $($bestvmachine.npVmCfgObj.Name)  "  -ForegroundColor Cyan
                $bestvmachine.npRefId = $failedConversionIndex
                $bestvmachine.npAppConfiguration = $failedConversionParameter
                $bestvmachine.npAppName = $failedConversionParameter.PackageName
                if ($failedConversionParameter.PreInstallerArguments -ne $nul)
                {
                    Write-host "   Has preinstall $($failedConversionParameter.PreInstallerArguments)"
                    $bestvmachine.npPreInstall = $failedConversionParameter.PreInstallerArguments
                }
                else
                {
                    $bestvmachine.npPreInstall = $nul
                }
                $bestvmachine.npInUse = $true                
                $failedConversionParameter.Started = $true
                $failedConversionParameter.Completed = $false

                $redoid = $AppConversionParameters.Count + $failedConversionIndex + 1
                $templateFilePath = CreateMPTTemplate $failedConversionParameter $redoid $bestvmachine.npVmCfgObj $workingDirectory 
                $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($bestvmachine.npVmCfgObj.Credential.Password)
                $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    
    
                PackageThis $bestvmachine $password $templateFilePath $bestvmachine.npVmCfgObj.initialSnapshotName "$($logfolder)\$($failedConversionParameter.PackageName)_Redo_$(Get-Date -format FileDateTime).txt"
        
                #####WaitForFreeVM $redoVirtMachinesArray $workingDirectory
            }
            else
            {
                Write-host "Redo[$($failedConversionIndex)] skipped"
            }
        }    
    }

}

function WaitForFreeVM($virtMachinesArray, $workingDirectory)
{
    $thisPassFailedConversionIndexes =  New-Object -TypeName "System.Collections.ArrayList"
    # ensure array is clear
    foreach ($pfci in $thisPassFailedConversionIndexes)
    {
        $thisPassFailedConversionIndexes.Remove($pfci)
    }

    $CountEnabled = 0
    foreach ($vm in $virtMachinesArray)
    {
        if (-not $vm.npDisabled)
        {
            $CountEnabled = $CountEnabled + 1
        }
    }
    $numAvailable = 0
        
    while ($numAvailable -eq 0)
    {
        #Sleep(1)
        foreach ($vm in $virtMachinesArray)
        {
            if ($vm.npDisabled -eq $false)
            { 
                if ($vm.npInUse -eq $true)
                { 
                    if ($vm.npJobObj.State -eq 'Running') 
                    { 
                        if ($vm.npVmGetObj.upTime.TotalHours -gt 2.0)
                        {
                            Write-Host "Timeout on $($vm.npVmGetObj.Name) processing $($vm.npAppName)." -ForegroundColor Red
                            Stop-Job -Job $vm.npJobObj
                            Remove-Job -Job $vm.npJobObj -Force
                            $vm.npJobObj = $nul
                            Checkpoint-VM -ComputerName $vm.npVmGetObj.ComputerName -Name $vm.npVmGetObj.Name -SnapshotName "$($vm.npAppName)_$(get-date)"
                            Stop-VM -ComputerName $vm.npVmGetObj.ComputerName -Name $vm.npVmGetObj.Name -TurnOff -ErrorAction SilentlyContinue
                            Set-VM -ComputerName $vm.npVmGetObj.ComputerName -Name $vm.npVmGetObj.Name -Notes 'none'
                            $thisPassFailedConversionIndexes.Add(($vm.npRefId)) > $xxx                            
                            $vm.npRefId = -1
                            $vm.npAppConfiguration = $nul
                            $vm.npAppName = ''
                            $vm.npInUse = $false
                            $vm.npErrorCount = $vm.npErrorCount + 1
                            if ($vm.npErrorCount -gt 5 -and $CountEnabled -gt 1)
                            {
                                $vm.npDisabled = $true
                                $CountEnabled -= 1
                                Write-Host "Disabling $($vm.npVmGetObj.Name) due to exess errors" -BackgroundColor DarkRed -ForegroundColor White
                            }
                            else    
                            {
                                $semaphore.Release()
                                $numAvailable += 1
                            }
                        }
                        else
                        {
                            #$countInUse += 0
                        } 
                    }
                    else
                    {
                        write-host "debug: job $($vm.npJobObj.Name) state $($vm.npJobObj.State) "
                        if (Test-Path -Path "$($workingDirectory)\MSIX\$($vm.npAppName)_*.msix")
                        {
                            Write-Host "Completion of  $($vm.npAppName) on $($vm.npVmGetObj.Name)." -ForegroundColor Green
                            $vm.npAppConfiguration.Completed = $true
                            $thisPassFailedConversionIndexes.Add(-1) > $xxx
                        }
                        else
                        {
                            Write-Host "Completion without package of  $($vm.npAppName) on $($vm.npVmGetObj.Name)." -ForegroundColor Red
                            Checkpoint-VM -ComputerName $vm.npVmGetObj.ComputerName -Name $vm.npVmGetObj.Name -SnapshotName "$($vm.npAppName)_$(get-date)"
                            $thisPassFailedConversionIndexes.Add(($vm.npRefId)) > $xxx
                            $vm.npErrorCount = $vm.npErrorCount + 1
                        }
                        Stop-Job -Job $vm.npJobObj
                        Remove-Job -Job $vm.npJobObj -Force
                        $vm.npJobObj = $nul
                        if ($vm.npVmGetObj.State -eq 'Running')
                        {
                            Stop-VM -ComputerName $vm.npVmGetObj.ComputerName -Name $vm.npVmGetObj.Name -TurnOff -ErrorAction SilentlyContinue
                        }
                        Set-VM -ComputerName $vm.npVmGetObj.ComputerName -Name $vm.npVmGetObj.Name -Notes 'none'
                        $vm.npRefId = -1
                        $vm.npAppConfiguration = $nul
                        $vm.npAppName = ''
                        $vm.npInUse = $false
                        if ($vm.npErrorCount -gt 5 -and $CountEnabled -gt 1)
                        {
                            $vm.npDisabled = $true
                            $CountEnabled -= 1
                            Write-Host "Disabling $($vm.npVmGetObj.Name) due to exess errors" -BackgroundColor DarkRed -ForegroundColor White
                        }
                        else
                        {
                            $semaphore.Release()
                            $numAvailable += 1
                        }
                    }
                }
                else
                {
                    ## VM already not in use
                    $numAvailable += 1
                }
            }
            else
            {
                # VM already Disabled
            }
        }
    }
    return $thisPassFailedConversionIndexes 
}

function CountEnabledInuseVMs($virtMachinesArray)
{
    $Count = 0
    foreach ($vm in $virtMachinesArray)
    {
        if ($vm.npDisabled -eq $false -and $vm.npInuse -eq $true)
        {
            $Count = $Count + 1
        }
    }
    return $Count
}

function PackageThis( $vMachine, $machinePassword, $templateFilePath, $initialSnapshotName, $logFile)
{
    Write-Output "Starting with params:   $($vMachine.npVmGetObj.ComputerName)/$($vMachine.npVmGetObj.Name),  $($machinePassword), $($templateFilePath), $($initialSnapshotName), $($logFile)" > $logFile

    try
    {
        Write-Output "debug: be4 get snapshot" >> $logFile
        $snap = Get-VMSnapshot -Name $initialSnapshotName -VMName $vMachine.npVmCfgObj.Name -ComputerName $vMachine.npVmCfgObj.Host -ErrorAction Continue
        Write-Output "debug: after get snapshot" >> $logFile
        if ( $snap)
        {
            Write-Output "Reverting VM snapshot for  $($vMachine.npVmCfgObj.Host) / $($vMachine.npVmCfgObj.Name): $($initialSnapshotName)" >> $logFile
            Restore-VMSnapshot -ComputerName $vMachine.npVmCfgObj.Host -VMName $vMachine.npVmCfgObj.Name -Name $initialSnapshotName -Confirm:$false
            Write-Output "debug: after revert" >> $logFile
            ####probably don't need to replace this, but once had an issue...
            Start-Sleep 5
            $vMachine.npVmGetObj = (get-vm -ComputerName $vMachine.npVmCfgObj.Host -Name $vMachine.npVmCfgObj.Name)
            Set-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name -Notes "Preparing $($vMachine.npAppName)"
            
            if ( $vMachine.npVmGetObj.state -eq 'Off'-or $vMachine.npVmGetObj.state -eq 'Saved')
            {
                Write-Output "Starting VM" >> $logFile
                Start-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name 
                #Write-Output "after Starting VM" >> $logFile
                $limit = 60
                while ($vMachine.npVmGetObj.state -ne 'Running')
                {
                    Start-Sleep 2
                    $limit = $limit - 1
                    if ($limit -eq 0)
                    {
                        Write-Output "TIMEOUT while starting restored checkpoint' state=$($vMachine.npVmGetObj.state)." >> $logFile
                        $vMachine.npErrorCount = $vMachine.npErrorCount + 1
                        if ($vMachine.npVmGetObj.state -ne 'Off')
                        {
                            Write-Host "Debug: Stop VM $( $vMachine.npVmGetObj.Name)" >> $logFile
                            Stop-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name -TurnOff -Force -ErrorAction SilentlyContinue
                        }
                        break;
                    }
                }
            }
            else
            {
                Write-Output "debug state is $($vMachine.npVmGetObj.state)" >> $logFile
            }

        }
        else
        {
            Write-Output "Get-VMSnapshot error" >> $logFile
        }

        $waiting = $true
        $waitcount = 0
        ## Let VM Settle a little.  At times the VMs get a little busy thanks to MS and more time seems to work better.
        while ($waiting)
        {
            if ( $vMachine.npVmGetObj.state -eq 'Running' -and $vMachine.npVmGetObj.upTime.TotalSeconds -gt 30  )
            {
                $waiting = $false

                
                if ($vMachine.npPreInstall -ne $nul)
                {
                    Write-output "Before RunRemotePreInstaller"
                    RunRemotePreInstaller $vMachine.npVmGetObj.Name $vMachine.npPreInstall >> $logFile
                    Write-output "After RunRemotePreInstaller"
                }

                Set-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name -Notes "Packaging $($vMachine.npAppName)"
                Write-Output "" >> $logFile
                Write-Output "==========================Starting package..." >> $logFile
                MsixPackagingTool.exe create-package --template $templateFilePath --machinePassword $machinePassword -v >> $logFile
                Write-Output "==========================Packaging tool done." >> $logFile
                Write-Output "" >> $logFile
            }
            $waitcount += 1
            if ($waitcount -gt 360)
            {
                $waiting = $false
                Write-Output "Timeout waiting for OS to start" >> $logFile
                $vMachine.npErrorCount = $vMachine.npErrorCount + 1
                if ($vMachine.npVmGetObj.state -ne 'Off')
                {
                    Write-Host "Debug: Stop VM $( $vMachine.npVmGetObj.Name)" >> $logFile
                    Stop-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name -TurnOff -Force -ErrorAction SilentlyContinue
                }
            }
            start-sleep 1
        }
        Write-Output "Debug: job ready for finalizing." >> $logFile
    }
    catch
    {
        Write-Output "***EXECPTION***"
        Write-Output "***EXECPTION***" >> $logFile
        Write-Output $_
        Write-Output $_ >> $logfile
    }
    finally
    {
        Write-Output "Finalizing." >> $logFile
        Stop-VM -ComputerName $vMachine.npVmGetObj.ComputerName -Name $vMachine.npVmGetObj.Name -TurnOff -Force -ErrorAction SilentlyContinue

        #Read-Host -Prompt 'Press any key to exit this window '
        Write-Output "Complete." >> $logFile
        Write-Output "Complete."
    }
}

function JobSetEnabledValue($conversionParam, $value)
{
    $conversionParam.Enabled = $value
}
function JobSetStartedValue($conversionParam, $value)
{
    $conversionParam.Started = $value
}
function JobSetCompletedValue($conversionParam,$value)
{
    $conversionParam.Completed = $value
}


function FindUndoneJobs($AppConversionParameters,$workingDirectory)
{
    # NO longer used
    $undoneConversionArray = @() 
    foreach ($conversionParam in $AppConversionParameters)
    {
        if ( $conversionParam.Enabled)
        {
            if (-not (Test-Path -Path "$($workingDirectory)\MSIX\$($conversionParam.PackageName)_*.msix"))
            {
                $undoneConversionArray += $conversionParam
            }
        }
    }
    return $undoneConversionArray
}

function SignPackages($msixFolder, $signtoolPath, $certfile, $certpassword, $timestamper, $doOnlyThisPackageName)
{

    Get-ChildItem $msixFolder | foreach-object {
        $msixPath = $_.FullName
        if ($doOnlyThisPackageName -eq '')
        {
            Write-Host "Running: $signtoolPath sign /f $certfile /p "*redacted*" /fd SHA256 /t $timestamper $msixPath"
            & $signtoolPath sign /f $certfile  /p $certpassword /fd SHA256 /t $timestamper $msixPath
        }
        elseif ($msixPath -like "*$($doOnlyThisPackageName)*")
        {
            Write-Host "Running: $signtoolPath sign /f $certfile /p "*redacted*" /fd SHA256 /t $timestamper $msixPath"
            & $signtoolPath sign /f $certfile  /p $certpassword /fd SHA256 /t $timestamper $msixPath
        }
    }
}

function AutoFixPackages($AppConversionParameters, $inputfolder, $outputFolder)
{
    $Toolpath = 'TMEditX.exe'
    $Processed = 0
    $Attempted = 0

    if (-not (Test-Path -Path $outputFolder))
    {
        Write-Host 'Creating msixpsf directory.'
        New-Item -Force -Type Directory ($outputFolder)
    }

    foreach ($conversionParam in $AppConversionParameters)
    {
        $msixPath = [System.IO.Path]::Combine($inputfolder, $conversionParam.PackageName)
        $msixOutPath = [System.IO.Path]::Combine($outputFolder, $conversionParam.PackageName)
        if ( $conversionParam.Enabled )
        {
            if (  $conversionParam.Fixups -ne "")
            {
            $Attempted += 1
            $found = $false
            Get-ChildItem $inputfolder | foreach-object {
                if ($_.FullName.StartsWith($msixPath))
                {
                    $defangedConversionParamFixups = $conversionParam.Fixups
                    if ( $conversionParam.Fixups.StartsWith('"'))
                    {
                        $defangedConversionParamFixups = $conversionParam.Fixups.Trim('"').TrimEnd('"')
                    }
                    ##Write-Host "Running: $($Toolpath) $($defangedConversionParamFixups) $($_.FullName) $($msixOutPath)" -ForegroundColor Cyan
                    ##& $Toolpath $defangedConversionParamFixups /AutoSaveAsMsix $msixOutPath
                    Write-Host "Running: $($Toolpath) $($_.FullName) $($defangedConversionParamFixups)" -ForegroundColor Cyan
                    ##& $Toolpath $defangedConversionParamFixups $_.FullName
                    $arglist = "$($_.FullName) $($defangedConversionParamFixups)" 
                    Start-Process -Wait -FilePath $Toolpath -ArgumentList $arglist
                    $found = $true
                    $Processed += 1
                    Start-Sleep 5
                }
            }
            if (!$found)
            {
                Write-host "$($msixPath)_* not found" -ForegroundColor Yellow
            }
        }
        }
        else
        {
            Write-host "$($msixPath)_* skipped by request." -ForegroundColor Gray
        }
    }
     $countPackagesFix = (get-item "$($outputFolder)\*.msix").Count
    Write-Host "$($Processed) of $($Attempted) packages fixed; total $($countPackagesFix) fixed packages in output folder." -ForegroundColor Green

}

Function Find-BestPowerShellExe()
{
    $path = "powershell.exe"

    if (Test-Path "$($env:ProgramFiles)\PowerShell\8")
    {
        return "$($env:ProgramFiles)\PowerShell\8\powershell.exe"
    }
    if (Test-Path "$($env:ProgramFiles)\PowerShell\7")
    {
        return "$($env:ProgramFiles)\PowerShell\7\powershell.exe"
    }
    return "$($env:WINDIR)\system32\WindowsPowerShell\v1.0\powershell.exe"
}