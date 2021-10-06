<#
    .NOTES
        Author: Mark McGill, VMware
        Last Edit: 10-05-2021
        Version 1.1
    .SYNOPSIS
        Copy-VMDK will copy the source harddrive from a running source VM to the destination VM and attach it
    .DESCRIPTION
        1. Takes a snapshot of the source VM
        2. Clones the source VM from the snapshot
        3. Copies the source HD to the destination VM
        4. Renames the HD based on the source and destination names
        5. Attaches the new HD to the destination VM
        6. Deletes the VM Clone
        7. Deletes the source VM snapshot
    .EXAMPLE
        Copy-VMDK -sourceVmName "virtualMachine1" -sourceHdNames "Hard disk 2" -destinationVmName "virtualMachine2"
    .EXAMPLE
        #Copy multiple vmdks to the destination VM
        $harddisks = "Hard disk 1","Hard disk 2"
        Copy-VMDK -sourceVmName "virtualMachine1" -sourceHdNames $harddisks -destinationVmName "virtualMachine2"
    .EXAMPLE
        #unless you use the -overwrite flag, VMDKs on the destination VM will NOT be overwritten
        Copy-VMDK -sourceVmName "virtualMachine1" -sourceHdNames "Hard disk 2" -destinationVmName "virtualMachine2" -overwrite
    .OUTPUTS
        Returns object containing details of source, clone, and destination VMs and hard disks
        Returns error specifics on failure
#>
function Copy-VMDK
{
    #Requires -Version 5.0
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory=$true)][string]$sourceVmName,
        [Parameter(Mandatory=$true)][array]$sourceHdNames,
        [Parameter(Mandatory=$true)][string]$destinationVmName,
        [Parameter(Mandatory=$false)][switch]$overwrite = $false
    )

    Function CleanUp($snapshot,$cloneVm)
    {
        If ($snapshot -ne $null)
        {
            #Delete Snapshot
            Write-Verbose "Deleting snapshot $($snapshot.Name)"
            $snapshot | Remove-Snapshot -Confirm:$false
        }
        If ($cloneVm -ne $null)
        {
            #Delete cloned VM
            Write-Verbose "Deleting clone VM $($cloneVm.Name)"
            $cloneVm | Remove-VM -DeletePermanently -Confirm:$false
        }
        $ErrorActionPreference = "Continue"
    }

    $ErrorActionPreference = "Stop"

    #create VM Snapshot on source VM
    Try
    {
        $sourceVm = Get-VM -Name $sourceVmName
        $destinationVm = Get-VM $destinationVmName
        $destinationHd0 = ($destinationVm | Get-HardDisk)[0]
        $destinationDs = Get-Datastore -Id "$($destinationHd0.ExtensionData.Backing.Datastore.Type)-$($destinationHd0.ExtensionData.Backing.Datastore.Value)"
        $destinationVmPath = ($destinationHd0.Filename -replace $($destinationHd0.FileName.Split("/")[-1]),"").Trim("/")
        $cloneVmName = "Clone_of_$sourceVmName-$(get-date -Format MM-dd-yyyy-HH-mm-ss)"
        $cloneFolder = $sourceVm.ExtensionData.Parent
        $snapshotName = "$sourceVMName-CopyVMDK-Script-$(get-date -Format MM-dd-yyyy-HH-mm-ss)"
        Write-Verbose "Creating snapshot $snapshotName for VM $sourceVmName"
        $snapshot = New-Snapshot -VM $sourceVmName -Name $snapshotName -Quiesce
    }
    Catch
    {
        CleanUp $snapshot $cloneVm
        Return "Unable to take VM snapshot: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
    }

    Try
    {
        #create objects to pass for new VM creation
        $vmRelocateSpec = New-Object Vmware.Vim.VirtualMachineRelocateSpec -Verbose:$verbose
        $vmRelocateSpec.datastore = $datastore.MoRef
        $vmRelocateSpec.DeviceChange = $vDeviceConfigSpec

        $vmCloneSpec = New-Object VMware.Vim.VirtualMachineCloneSpec -Verbose:$verbose
        $vmCloneSpec.location = $vmRelocateSpec
        $vmCloneSpec.powerOn = $false
        $vmCloneSpec.template = $false
        $vmCloneSpec.Snapshot = $snapshot.Id

        #Clone source VM
        Write-Verbose "Cloning $($sourceVm.Name) to $cloneVmName"
        $cloneTask = $SourceVM.ExtensionData.CloneVM($cloneFolder, $cloneVmName, $vmCloneSpec)
        $cloneVm = Get-VM $cloneVmName
        
    }
    Catch
    {
        CleanUp $snapshot $cloneVm
        Return "Unable to clone VM: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
    }
    
    $returnArr = @()
    foreach($sourceHdName in $sourceHdNames)
    {
        $returnObj = "" | Select SourceVM,SourceHD,SourceVMDK,Snapshot,CloneVM,DestinationVM,DestinationHD,DestinationVMDK

        #copy specified VMDK to destination VM
        Try
        {
            $sourceHd = $sourceVm | Get-HardDisk -Name $sourceHdName
            $cloneHd = $cloneVm | Get-HardDisk -Name $sourceHdName
            $cloneDs = Get-Datastore -Id "$($cloneHd.ExtensionData.Backing.Datastore.Type)-$($cloneHd.ExtensionData.Backing.Datastore.Value)"
            $destinationVmdkName = $cloneHd.Filename.Split("/")[-1] -replace $cloneVmName,$destinationVmName
            $destinationHds = $destinationVm | Get-HardDisk
            If(($destinationHds.Filename | Where {$_ -match $destinationVmdkName}).Count -gt 0)
            {
                If ($overwrite -eq $false)
                {
                    Write-Error "$destinationVmdkName already exists on $(destinationVm.Name). Re-run with -overwrite option or delete the HD"
                }
                else 
                {
                    $destinationHd = $destinationHds | Where{$_.Filename -match $destinationVmdkName}
                    Write-Verbose "Deleting existing vmdk $destinationHdName"
                    $removeHD = $destinationHd | Remove-HardDisk -DeletePermanently -Confirm:$false
                }
            }
            $cloneVmdk = $cloneHd.FileName.Split("/")[-1]
            Write-Verbose "Copying VMDK $($cloneHd.Filename) to $($destinationVmPath)"
            $hdCopy = $cloneHd | Copy-HardDisk $destinationVmPath -Confirm:$false
        }
        Catch
        {
            #Delete cloned VM
            Write-Host "Unable to copy VMDK file $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
            break
        }

        Try
        {
            Write-Verbose "Renaming VMDK $cloneVmdk to $destinationVmdkName"
            Rename-Item vmstore:\$(($destinationVm | Get-Datacenter).Name)\$($destinationDs.Name)\$destinationVmName\$cloneVmdk -NewName $destinationVmdkName
            #attach HD
            Write-Verbose "Attaching VMDK $destinationVmdkName to VM $destinationVmName"
            $attachHd = $destinationVm | New-HardDisk -DiskPath "$destinationVMPath/$destinationVmdkName" -Confirm:$false
        }
        Catch
        {
            Write-Host "Unable to rename or attach VMDK: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
            break    
        }
        $returnObj.SourceVM = $sourceVm.Name
        $returnObj.SourceHd = $sourceHd.Name
        $returnObj.SourceVMDK = $sourceHd.Filename
        $returnObj.Snapshot = $snapshot.Name
        $returnObj.CloneVM = $cloneVm.Name
        $returnObj.DestinationVM = $destinationVm.Name
        $returnObj.DestinationHD = $destinationHdName
        $returnObj.DestinationVMDK = "$destinationVMPath/$destinationHdName"
        $returnArr += $returnObj
    } #end foreach 

    CleanUp $snapshot $cloneVm

    Return $returnArr
}