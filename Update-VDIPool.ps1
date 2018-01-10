<# 
    .NOTES
    ===============================================================================================
    Author                      : James Wood
    Author email                : woodj@vmware.com
    Version                     : 1.0
    ===============================================================================================
    Tested Against Environment:
    Horizon View Server Version : 7.0.2
    PowerCLI Version            : PowerCLI 6.5
    PowerShell Version          : 5.0, 5.1
    ===============================================================================================
    Changelog:
    01.2017 ver 1.0 Initial Version

    ===============================================================================================

    Copyright © 2016 VMware, Inc. All Rights Reserved.
    
    Permission is hereby granted, free of charge, to any person obtaining a copy of
    this software and associated documentation files (the "Software"), to deal in
    the Software without restriction, including without limitation the rights to
    use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies
    of the Software, and to permit persons to whom the Software is furnished to do
    so, subject to the following conditions:
    
    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.
    
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
    ===============================================================================================

    .Synopsis
        Updates desktops in manual pool.
    .DESCRIPTION
        Updates desktops in manual pool by removing the existing desktops from the pool, deleting them
        from vCenter, clones parent image to new desktops then adds the new desktops to the manual pool.
    .OUTPUTS
        None
#>

Function New-RandomVMName {
    $randomName = "Win7-" + -join ((48..57) + (65..90) + (97..122) | Get-Random -count 10 | ForEach-Object {[char]$_})
    return $randomName
}

Function Cleanup-Session {
    param (
    [Parameter(Mandatory = $false)]
    [string[]]
    $err,

    [Parameter(Mandatory = $false)]
    [string]
    $vc,

    [Parameter(Mandatory = $false)]
    [string]
    $hv
    )

    If ($vc)
    {
        Disconnect-VIServer $vc -Confirm:$false
    }
    
    If ($hv)
    {
        Disconnect-HVServer $hv -Confirm:$false
    }

    If ($err)
    {
        Write-Error "Errors Encountered:"
        ForEach ($e in $err)
        {
            Write-Error $e.Exception.Message
        }

        $scriptPath = Split-Path -Parent $PSCommandPath
        $err | ConvertTo-Csv | Out-File -FilePath $scriptPath\Errors.csv

        Write-Host "The error details have been written to $scriptPath\Errors.csv"
    }
    Exit
}

#Get credentials for connecting to Horizon Admin and vCenter
$cred = Get-Credential -Message "Enter domain\user and password."

#Get vCenter address
$txtvCenter = Read-Host -Prompt "Enter vCenter FQDN"

#Get Horizon Admin address
$txtHVAdmin = Read-Host -Prompt "Enter Horizon Admin FQDN"

#Get Horizon Pool ID
$txtPoolID = Read-Host -Prompt "Enter Horizon Manual Pool ID"

#Get Parent VM Name
$txtParentName = Read-Host -Prompt "Enter Parent VM Name"

#Get Destination vCenter Folder Name
$txtFolderName = Read-Host -Prompt "Enter Folder Name for Desktop VM's"

#Get Resource Pool for New VMs
$txtCluster = Read-Host -Prompt "Enter the vCenter Cluster Name for the new Desktop VMs"

#Get Datastore for New VMs
$txtDatastore = Read-Host -Prompt "Enter the vCenter Datastore name for the new Desktop VMs"

#Get number of desktops to create
$NumDesktops = Read-Host -Prompt "Enter the Number of Desktops to Create"

#Get Guest Customizaton Specification Name
$txtOSCustSpec = Read-Host -Prompt "Enter the Name of the Desktop Customization Specification from vCenter"

#Connect to vCenter
Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow "Connecting to vCneter..."
$objvCenter = $null
Do{
    Try
    {
        $objvCenter = Connect-VIServer -Server $txtvCenter -Credential $cred -ErrorAction Stop
    }
    Catch [VMware.VimAutomation.Sdk.Types.V1.ErrorHandling.VimException.ViServerConnectionException]
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error connecting to $txtvCenter."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Re-enter the vCenter FQDN/IP"
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        $txtvCenter = Read-Host -Prompt "Please re-enter the FQDN of your vCenter server"
    }
    Catch [VMware.VimAutomation.ViCore.Types.V1.ErrorHandling.InvalidLogin]
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Invalid login credentials."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Re-enter your user name/pswd"
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        $cred = Get-Credential -Message "Please re-enter your user name and password."
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  An unhandled error has occured."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  The script will now exit!"
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Cleanup-Session -err $_ -vc $objvCenter
    }
}
Until ($objvCenter -ne $null)
Write-host -ForegroundColor DarkYellow "Connected to $txtvCenter!"

Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow "Connecting to Horizon Administrator..."
#Connect to Horizon Admin
$objHVAdmin = $null
Do{
    Try
    {
        $objHVAdmin = Connect-HVServer -Server $txtHVAdmin -Credential $cred -ErrorAction Stop
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error connecting to $txtHVAdmin."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Re-enter connection information."
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        $TryHVConnection++
        If($TryHVConnection -ge 3)
        {
            "Connection failed 3 times. Exiting Script!"
            Cleanup-Session -err $_ -vc $objvCenter -hv $objHVAdmin
        }
        Start-Sleep -Seconds 3
        $cred = Get-Credential -Message "Please re-enter your user name and password."
        $txtHVAdmin = Read-Host -Prompt "Please re-enter the FQDN of your Horizon Admin server"    
    }
}
Until ($objHVAdmin -ne $null)

Write-host -ForegroundColor DarkYellow "Connected to $txtHVAdmin!"

Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow -NoNewline "Disabling Desktop Pool $txtPoolID..."
#Disable Desktop Pool
Try
{
    Set-HVPool -PoolName $txtPoolID -Disable
}
Catch
{
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error disabling Desktop Pool $txtPoolID."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  The script will now exit!"
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Cleanup-Session -err $_ -vc $objvCenter -hv $objHVAdmin
}

Write-Host -ForegroundColor DarkYellow " Disabled!"

Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow "Powering off existing Desktop VMs..."
#Poweroff VDI guests
$vms = $null
Do{
    Try
    {
        $vms = Get-VM -Location $txtFolderName -ErrorAction Stop
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error finding folder $txtFolderName."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Re-enter the folder name"
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        $txtFolderName = Read-Host -Prompt "Please re-enter the folder name"
    }
}
Until ($vms -ne $null)

ForEach ($vm in $vms)
{
    Try
    {
        Stop-VM -VM $vm -Confirm:$false -ErrorAction Stop
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error powering off $vm."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Check that VM is powered off in vCenter."
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
    }
}
Write-Host -ForegroundColor DarkYellow "Power Operation Complete!"

#Remove Existing Desktops from Manual Pool
#ToDo - I haven't figured out how to programmatically remove a desktop from a pool.  Suggestions here are welcome!
Read-Host -Prompt "Remove the existing desktops from the $txtPoolID pool.  Press enter when complete to continue the script"

Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow -NoNewline "Deleting existing Desktop VMs..."
#Delete VMs
ForEach ($vm in $vms)
{
    Try
    {
        Remove-VM -VM $vm -DeletePermanently -Confirm:$false -ErrorAction Stop
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error deleting $vm."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Check that VM has been deleted in vCenter."
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
    }
}
Write-Host -ForegroundColor DarkYellow " Done!"

Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow "Cloning new Desktop VMs..."
#Clone new VDI Guests
$objCustSpec = $null
Do
{
    Try
    {
        $objCustSpec = Get-OSCustomizationSpec -Name $txtOSCustSpec -ErrorAction Stop
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error loading customization spec $txtOSCustSpec."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Please re-enter the customizaton spec name."
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        $txtOSCustSpec = Read-Host -Prompt "Please re-enter the customization spec name."
    }
}
Until ($objCustSpec -ne $null)

ForEach ($i in 1..$NumDesktops)
{
    Try
    {
        $newName = New-RandomVMName
        New-VM -Name $newName -VM $txtParentName -Location $txtFolderName -ResourcePool $txtCluster -Datastore $txtDatastore -OSCustomizationSpec $objCustSpec -Confirm:$false -ErrorAction Stop
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error creating VM $newName."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  VM will not be created."
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Error $_.Exception.Message
    }
}
Write-Host -ForegroundColor DarkYellow "New Desktops Created!"

Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow "Powering on new Desktop VMs..."
#Start new VDI Guests
$vms = $null
Do{
    Try
    {
        $vms = Get-VM -Location $txtFolderName -ErrorAction Stop
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error finding folder $txtFolderName."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Re-enter the folder name"
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        $txtFolderName = Read-Host -Prompt "Please re-enter the folder name"
    }
}
Until ($vms -ne $null)

ForEach ($vm in $vms)
{
    Try
    {
        Start-VM -VM $vm -Confirm:$false -ErrorAction Stop
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error powering on $vm."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Check that VM is powered on in vCenter."
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
    }
}
Write-Host -ForegroundColor DarkYellow "Power Operation Completed!"
Read-Host "Press Enter when Sys-prep process has completed"
Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow -NoNewline "Adding new Desktop VMs to $txtPoolID..."
#Add new desktops to pool
Try
{
    Add-HVDesktop -PoolName $txtPoolID -Machines $vms -Vcenter $objvCenter -HvServer $objHVAdmin -ErrorAction Stop
}
Catch
{
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error adding VMs to Desktop Pool $txtPoolID."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  You may need to add them manualy."
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
}
Write-Host -ForegroundColor DarkYellow " Done!"
Write-Host -ForegroundColor DarkYellow "Waiting for desktops to check in with Horizon Admin Portal..."
Start-Sleep -Seconds 10
#Check for desktops to register with Horizon Admin Portal
Do
{
    $wait = Get-HVMachine -PoolName $txtPoolID | Where-Object {$_.Base.BasicState -ne "Available"}
    Write-Host -NoNewline "."
    Start-sleep -Seconds 10
}
While ($wait.count -gt 0)
Write-Host ""
Write-Host -ForegroundColor DarkYellow "All desktops have checked in!"
Start-Sleep -Seconds 10
#Poweroff VDI guests
Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow "Shutting down new Desktop VMs..."
$vms = $null
Do{
    Try
    {
        $vms = Get-VM -Location $txtFolderName -ErrorAction Stop
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error finding folder $txtFolderName."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Re-enter the folder name"
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        $txtFolderName = Read-Host -Prompt "Please re-enter the folder name"
    }
}
Until ($vms -ne $null)

ForEach ($vm in $vms)
{
    Try
    {
        Stop-VMGuest -VM $vm -Confirm:$false -ErrorAction Stop
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error powering off $vm."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Check that VM is powered off in vCenter."
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
    }
}
Write-Host -ForegroundColor DarkYellow "Waiting for all guest OS to shutdown..."
Do
{
    $vmOn = Get-VM -Location $txtFolderName | Where-Object {$_.PowerState -ne "PoweredOff"}
    Write-Host -ForegroundColor DarkYellow -NoNewline "."
    Start-Sleep -Seconds 5
}
While ($vmOn.Count -gt 0)
Write-Host ""
Write-Host -ForegroundColor DarkYellow "Power Operation Complete!"

#Set disk persistence
Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow "Setting up guest disks..."
ForEach ($vm in $vms)
{
    Try
    {
        Get-HardDisk -VM $vm | Set-HardDisk -Persistence IndependentNonPersistent -Confirm:$false -ErrorAction Stop
    }
    Catch
    {
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error setting disk mode on $vm."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  You will need to set the disk mode manualy."
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
    }
}
Write-Host -ForegroundColor DarkYellow "Disk Settings Complete!"
#Enable Desktop Pool
Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow -NoNewline "Enable Desktop Pool..."
Try
{
    Set-HVPool -PoolName $txtPoolID -Enable
}
Catch
{
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  Error enabling Desktop Pool $txtPoolID."
        Write-Host -ForegroundColor Red -BackgroundColor Black "|  The script will now exit!"
        Write-Host -ForegroundColor Red -BackgroundColor Black "----------"
        Cleanup-Session -err $_ -vc $objvCenter -hv $objHVAdmin
}
Write-Host -ForegroundColor DarkYellow " Done!"


Write-Host -ForegroundColor DarkYellow "****"
Write-Host -ForegroundColor DarkYellow "Cleaning House!"
Cleanup-Session -err $null -vc $vCenter -hv $txtHVAdmin