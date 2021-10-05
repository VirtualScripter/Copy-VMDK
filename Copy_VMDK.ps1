<#
    .NOTES
        Author: Mark McGill, VMware
        Last Edit: 10-05-2021
        Version 1.0
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
        Copy-VMDK -sourceVmName "virtualMachine1" -sourceHdName "Hard disk 2" -destinationVmName "virtualMachine2"
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
        [Parameter(Mandatory=$true)][string]$sourceHdName,
        [Parameter(Mandatory=$true)][string]$destinationVmName
    )

    $returnObj = "" | Select SourceVM,SourceHD,SourceVMDK,Snapshot,CloneVM,DestinationVM,DestinationHD,DestinationVMDK
    $ErrorActionPreference = "Stop"
    $cloneVmName = "Clone_of_$sourceVmName-$(get-date -Format MM-dd-yyyy-HH-mm-ss)"

    #create VM Snapshot on source VM
    Try
    {
        $sourceVm = Get-VM -Name $sourceVmName
        $sourceHd = $sourceVm | Get-HardDisk -Name $sourceHdName
        $cloneFolder = $sourceVm.ExtensionData.Parent
        $snapshotName = "$sourceVMName-CopyVMDK-Script-$(get-date -Format MM-dd-yyyy-HH-mm-ss)"
        Write-Verbose "Creating snapshot $snapshotName for VM $sourceVmName"
        $snapshot = New-Snapshot -VM $sourceVmName -Name $snapshotName
    }
    Catch
    {
        $ErrorActionPreference = "Continue"
        Return "Unable to take VM snapshot: $($_.Exception.Message)"
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
    }
    Catch
    {
        Write-Verbose "Deleting snapshot $($snapshot.Name)"
        $snapshot | Remove-Snapshot -Confirm:$false
        $ErrorActionPreference = "Continue"
        Return "Unable to clone VM: $($_.Exception.Message)"
    }

    #copy specified VMDK to destination VM
    Try
    {
        $cloneVm = Get-VM $cloneVmName
        $cloneHd = $cloneVm | Get-HardDisk -Name $sourceHdName
        $cloneDs = Get-Datastore -Id "$($cloneHd.ExtensionData.Backing.Datastore.Type)-$($cloneHd.ExtensionData.Backing.Datastore.Value)"

        $destinationVm = Get-VM $destinationVmName
        $destinationHd0 = ($destinationVm | Get-HardDisk)[0]
        $destinationDs = Get-Datastore -Id "$($destinationHd.ExtensionData.Backing.Datastore.Type)-$($destinationHd.ExtensionData.Backing.Datastore.Value)"

        $cloneVmdk = $cloneHd.FileName.Split("/")[-1]

        $destinationVmPath = ($destinationHd0.Filename -replace $($destinationHd.FileName.Split("/")[-1]),"").Trim("/")
        Write-Verbose "Copying VMDK $($cloneHd.Filename) to $($destinationVmPath)"
        $hdCopy = $cloneHd | Copy-HardDisk $destinationVmPath -Confirm:$false
    }
    Catch
    {
        #Delete cloned VM
        Write-Verbose "Deleting clone VM $($cloneVm.Name)"
        $cloneVM | Remove-VM -DeletePermanently -Confirm:$false
        #Delete Snapshot
        Write-Verbose "Deleting snapshot $($snapshot.Name)"
        $snapshot | Remove-Snapshot -Confirm:$false
        $ErrorActionPreference = "Continue"
        Write-Error "Unable to copy VMDK file"
        Return "$($_.Exception.Message)"
    }

    Try
    {
        $destinationHdName = $cloneHd.Filename.Split("/")[-1] -replace $cloneVmName,$destinationVmName
        Write-Verbose "Renaming $cloneVmdk to $destinationHdName"
        Rename-Item vmstore:\$(($destinationVm | Get-Datacenter).Name)\$($destinationDs.Name)\$destinationVmName\$cloneVmdk -NewName $destinationHdName
        #attach HD
        Write-Verbose "Attaching VMDK $destinationHdName to VM $destinationVmName"
        $renamedHd = $destinationVm | New-HardDisk -DiskPath "$destinationVMPath/$destinationHdName" -Confirm:$false
    }
    Catch
    {
        #Delete cloned VM
        Write-Verbose "Deleting clone VM $($cloneVm.Name)"
        $cloneVM | Remove-VM -DeletePermanently -Confirm:$false
        #Delete Snapshot
        Write-Verbose "Deleting snapshot $($snapshot.Name)"
        $snapshot | Remove-Snapshot -Confirm:$false
        $ErrorActionPreference = "Continue"
        Return "Unable to rename or attach VMDK: $($_.Exception.Message)"        
    }

    #Delete cloned VM
    Write-Verbose "Deleting clone VM $($cloneVm.Name)"
    $cloneVM | Remove-VM -DeletePermanently -Confirm:$false
    #Delete Snapshot
    Write-Verbose "Deleting snapshot $($snapshot.Name)"
    $snapshot | Remove-Snapshot -Confirm:$false

    $ErrorActionPreference = "Continue"
    $returnObj.SourceVM = $sourceVm.Name
    $returnObj.SourceHd = $sourceHd.Name
    $returnObj.SourceVMDK = $sourceHd.Filename
    $returnObj.Snapshot = $snapshot.Name
    $returnObj.CloneVM = $cloneVm.Name
    $returnObj.DestinationVM = $destinationVm.Name
    $returnObj.DestinationHD = $destinationHdName
    $returnObj.DestinationVMDK = "$destinationVMPath/$destinationHdName"

    Return $returnObj
}