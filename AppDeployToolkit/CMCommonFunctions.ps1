#region supporting functions
If (Test-Path "$env:WinDir\CCM\ccmexec.exe")
{
	$ccmexecPath = "$env:WinDir\CCM\ccmexec.exe"
}

Function Get-CMVersion
{
	If ($ccmexecPath)
	{
		$CMVersion = (Get-Item $ccmexecPath).VersionInfo.ProductVersion
		Write-Log	"Info: Current SCCM client version is $CMVersion" -Source Get-CMVersion
	}
}

Function WaitThirtySeconds
{
    Write-Log -Message "Waiting 30 seconds"
    Start-Sleep -Seconds 30
}

Function WaitFiveMinutes
{
	Write-Log -Message "Waiting 5 minutes"
	Start-Sleep -Seconds 300
}

function Parse-Date ([string]$DateString, [string]$Format, [switch]$Trim) #region Parse-Date
{
	if ($Trim) { $DateString = $DateString.SubString(0, $Format.Length) }
	[DateTime]::ParseExact($DateString, $Format, [Globalization.CultureInfo]::InvariantCulture)
} #endregion Parse-Date
#endregion supporting functions

#region SCCM client service functions
Function Test-CMClientServiceExists
{
	If (Test-ServiceExists -Name 'CcmExec')
	{
		Write-Log -Message "Info: SCCM Client service exists" -Source Test-CMClientServiceExists
		$true
	}
	Else
	{
		Write-Log -Message "Warning: SCCM Client service does not exist" -Source Test-CMClientServiceExists -Severity 2
		$false
	}
}
Function Test-CMClientServiceRunning
{
	If ((Get-Service 'CcmExec').Status -eq "Running")
	{
		Write-Log -Message "Info: SCCM Client service is running" -Source Test-CMClientServiceRunning
		$true
	}
	Else
	{
		Write-Log -Message "Warning: SCCM Client service does not exist" -Source Test-CMClientServiceExists -Severity 2
		$false
	}
}
Function Start-CMClientService {
    If (Test-CMClientServiceExists) {
        If (!(Test-CMClientServiceRunning)) {
            Try {
                Write-Log -Message "Info: attempting to start SCCM Client service" -Source Start-CMClientService
                Start-ServiceAndDependencies -Name 'CcmExec'
                WaitThirtySeconds
                Test-CMClientServiceRunning
            }
            Catch {
                Write-Log -Message "Error: failed to start SCCM Client service" -Source Start-CMClientService
            }
        }
        Else {Write-Log -Message "Warning: SCCM Client service already running" -Source Start-CMClientService -Severity 2}
    }
    Else {
        Write-Log "Fatal error: ccmexec service does not exist. SCCM client could not be installed." -Source Start-CMClientService -Severity 3
        Exit-Script -ExitCode 69013
    }
}

Function Stop-CMClientService {
    Stop-ServiceAndDependencies -Name 'CcmExec'
    (Get-Service 'CcmExec').Status|Write-Log -Source Stop-CMClientService
}

Function Wait-CMClientService
{
	#Wait for CCMExec to appear, if it doesn't happen within 5 minutes from finishing CCMSetup, fail the installation
	$counter = 5
	Do
	{
		Write-Log -Message "Info: Waiting for another $counter minute(s) for service to start" -Source Wait-CMClientService
		if ($counter -le 0)
		{
			Try { Repair-CMClient }
			Catch
			{
				Write-Log "Fatal error: ccmexec service could not be started" -Source Wait-CMClientService -Severity 3
				Exit-Script -ExitCode 69012
			}
		}
		Start-Sleep -s 60
		$counter = $counter - 1
	}
	While (!(Test-CMClientServiceRunning))
}
#endregion SCCM client service functions

#region ccm WMI namespace functions
Function Test-WMIRepository {
    Try {   
        Write-Log -Message "Testing WMI Repository" -Source Test-WMIRepository
        Execute-Process -Path "Winmgmt"  -Parameters "/verifyrepository"
        }
        Catch {Write-Log -Message "Warning: WMI Repository inconsistent" -Source Test-WMIRepository -Severity 2}
}

Function Salvage-WMIRepository {
    Write-Log -Message "Warning: Attempting to salvage WMI Repository" -Source Salvage-WMIRepository -Severity 2
    Execute-Process -Path "Winmgmt"  -Parameters "/salvagerepository"
}

Function Reset-WMIRepository {
    Write-Log -Message "Warning: Attempting to reset WMI Repository" -Source Reset-WMIRepository -Severity 2
    Execute-Process -Path "Winmgmt"  -Parameters "/resetrepository"
}

Function Test-CCMWMINamespace {
    If (Get-WMIobject -namespace root -class __NAMESPACE -filter "NAME='CCM'") {
        $true
        Write-Log -Message  "Info: CCM WMI Namespace exists" -Source Test-CCMWMINamespace
    }
    Else {
        Write-Log -Message  "Warning: CCM WMI Namespace doesn't exist" -Source Test-CCMWMINamespace -Severity 2
    }
}

Function Delete-CCMWMINamespace {
    If (Test-CCMWMINamespace)
        {
            Try {
                Write-Log -Message  "Info: Deleting CCM WMI Namespace" -Source Delete-CCMWMINamespace
                gwmi -query "SELECT * FROM __Namespace WHERE Name='CCM'" -Namespace "root" | Remove-WmiObject
            }
            Catch {
                Write-Log -Message  "Info: Failed to delete CCM WMI Namespace" -Source Delete-CCMWMINamespace -Severity 2
            }

        }
}

Function Repair-WMIRepository {
    Write-Log -Message "Warning: WMI Repository Broken" -Source Repair-WMIRepository -Severity 2
    Try {#Try to salvage WMI repository, fall back to resetting it
        Try {#Try to salvage WMI repository
            Salvage-WMIRepository
            WaitThirtySeconds
            If (!(Test-WMIRepository)){
                If (Test-CMClientServiceRunning) {
                    Stop-CMClientService
                }
            }
        }
        Catch {Write-Log -Message  "Warning: Attempt to salvage WMI Repository failed" -Source Repair-WMIRepository -Severity 2}
    }
    Catch {
        Try {#Try to reset WMI repository
            Write-Log -Message "Warning: Resetting WMI Repository" -Source Repair-WMIRepository -Severity 2
            Reset-WMIRepository
        }
        Catch {
            Write-Log -Message "Error: Failed to reset WMI Repository" -Source Repair-WMIRepository -Severity 3
            Exit-Script -ExitCode 69020
        }
    }
}
#endregion ccm WMI namespace functions

#region SCCM client installation/repair functions
Function Remove-CMClientFolders {
    Write-Log -Message "Removing folders related to SCCM client" -Source Remove-CMClientFolders
    If (Test-Path "$env:WinDir\ccmsetup") {Remove-Folder -Path "$env:WinDir\ccmsetup"}
    If (Test-Path "$env:WinDir\ccmcache") {Remove-Folder -Path "$env:WinDir\ccmcache"}
    If (Test-Path "$env:WinDir\CCM") {Remove-Folder -Path "$env:WinDir\CCM"}
}

Function Remove-CMClient {
    Write-Log -Message "Executing ccmclean"
	Try
	{
        Execute-Process -Path "$dirSupportFiles\ccmclean.exe" -Parameters "/q /all /logdir:`"$configToolkitLogDir`""
        Write-Log -Message "Waiting for CCMSetup to actually finish" -Source Remove-CMClient
        WaitThirtySeconds
        Get-Process -Name CCMSetup | Wait-Process
        WaitThirtySeconds
        If ($ccmexecPath) {
            Try {Execute-Process -Path "$dirFiles\ccmsetup.exe" -Parameters "/uninstall"}
            Catch{Write-Log -Message "Warning: Failed to uninstall SCCM client using ccmsetup /uninstall" -Source Remove-CMClient -Severity 2}
        }
		Remove-CMClientFolders
		Delete-CCMWMINamespace
        Exit-Script -ExitCode 3010
    }
    Catch {
        Write-Log -Message "Failed to uninstall sccm client using ccmclean" -Source Remove-CMClient -Severity 3
        Exit-Script -ExitCode 69012
    }
}
                               
Function Repair-CMClient {
    If ($ccmexecPath) {
        Try {
            Write-Log -Message "Info: Repairing SCCM Client" -Source Repair-CMClient
            $oSCCM = [wmiclass] “\root\ccm:sms_client”
            $oSCCM.RepairClient()

            WaitThirtySeconds
                                
            Write-Log -Message "Info: Waiting for CCMRepair to actually finish" -Source Repair-CMClient
            Get-Process -Name CCMRepair | Wait-Process
                                
            Write-Log -Message "Info: Waiting for CCMSetup to actually finish" -Source Repair-CMClient
            Get-Process -Name CCMSetup | Wait-Process
        }

        Catch {Write-Log -Message "Warning: SCCM client could not be repaired." -Source Repair-CMClient -Severity 2}
    }
    ElseIf (Get-Service 'CcmExec') {Write-Log -Message "Warning: SCCM client is not installed. Could not start the repair" -Source Repair-CMClient -Severity 2}
}

Function Install-CMClient {
    Write-Log -Message "Info: Installing SCCM Client" -Source Install-CMClient
    Try {
        If($args){
            Execute-Process -Path "$dirFiles\ccmsetup.exe" -Parameters "$args /noservice" -IgnoreExitCodes 7}
        Else {
            Execute-Process -Path "$dirFiles\ccmsetup.exe" -Parameters "/noservice" -IgnoreExitCodes 7}
        #Execute-Process -Path "$dirFiles\ccmsetup.exe" -Parameters "/source:`"$dirFiles\`" /noservice"}# /forceinstall $iSMSCACHESIZE $iSMSMP $iSMSMP $iFSP $iSMSSLP $iDNSSUFFIX"
    }
    Catch {
        Write-Log "Fatal Error: ccmsetup failed. SCCM client could not be installed." -Source Install-CMClient
        Exit-Script -ExitCode 69011
    }
    Write-Log -Message "Info: Waiting for CCMSetup to actually finish" -Source Install-CMClient
    Start-Sleep -Seconds 60
    Get-Process -Name CCMSetup | Wait-Process
    Start-Sleep -Seconds 120
    If (Test-CMClientServiceExists){
        If (!(Test-CMClientServiceRunning)){
            Start-CMClientService
        }
    }
    Else {
        Write-Log "Fatal error: ccmexec service does not exist. SCCM client could not be installed." -Source Install-CMClient -Severity 3
        Exit-Script -ExitCode 69013
    }
    Wait-CMClientService
}
#endregion SCCM client installation/repair functions

#region SCCM client site/location functions

Function Get-CMSite
{
	$([WmiClass]"\ROOT\ccm:SMS_Client").getassignedsite() | Select sSiteCode | Write-Log -Source Get-CMSite
	$CMSite = $([WmiClass]"\ROOT\ccm:SMS_Client").getassignedsite() | Select sSiteCode -ExpandProperty sSiteCode
}

Function Set-CMSite {
    Try {
        Write-Log -Message "Info: Automatically assigning site" -Source Set-CMSite

        get-wmiobject -query "SELECT * FROM Win32_Service" -namespace "ROOT\cimv2"
        $a=([wmi]"ROOT\ccm:SMS_Client=@");$a.EnableAutoAssignment=$True;$a.Put()
        get-wmiobject -query "SELECT * FROM Win32_Service WHERE Name ='CcmExec'" -namespace "ROOT\cimv2"

        Write-Log -Message "Info: Restarting CCMExec" -Source Set-CMSite
        Stop-CMClientService
        WaitThirtySeconds
        Start-CMClientService
        (Get-Service 'CcmExec').Status|Write-Log -Source Set-CMSite
        WaitThirtySeconds
        Get-CMSite|Write-Log -Source Set-CMSite
    }
    Catch {Write-Log -Message "Warning: failed to autodiscover SCCM site code" -Source Set-CMSite -Severity 2}
}
Function Get-CMPrimarySiteServer
{
	<#
		#	Created on:   	07.07.2017 11:47
		#	Created by:   	Mieszko Ślusarczyk
		#	Organization: 	International Paper
    .SYNOPSIS
    Get Configuration Manager Primary Site Server name.
    
    .DESCRIPTION
	The script will check the currently assigned SCCM primary site and display it's FQDN

    
    .EXAMPLE
    Get-CMPrimarySiteServer

    .DEPENDENT FUNCTIONS
    Write-Log

    #>
	Try
	{
		$SMSSiteCode = ([wmiclass]"ROOT\ccm:SMS_Client").GetAssignedSite().sSiteCode
		If ($SMSSiteCode)
		{
			$script:PrimarySiteServer = ([wmi]"ROOT\ccm:SMS_Authority.Name=`"SMS:$SMSSiteCode`"").CurrentManagementPoint
			If ($PrimarySiteServer)
			{
				Write-Log "Info: Primary Site Server FQDN is $PrimarySiteServer" -Source "Get-CMPrimarySiteServer"
			}
			
		}
	}
	Catch
	{
		Write-Log "Error: Failed to find Primary Site Server FQDN" -Severity 3 -Source "Get-CMPrimarySiteServer"
	}
	
}
#endregion SCCM client site/location functions

#region SCCM client schedules
Function Reset-CMMachinePolicy {
    Try {
        Write-Log -Message "Info: Resetting Machine Policy" -Source Reset-CMMachinePolicy
        ([wmiclass]'ROOT\ccm:SMS_Client').ResetPolicy(1)}
    Catch {Write-Log -Message "Error: Failed to reset Machine Policy" -Source Reset-CMMachinePolicy -Severity 2}
}

Function Execute-CMHardwareInventory {
    Try {
        Write-Log -Message "Info: Executing Hardware Inventory" -Source Execute-CMHardwareInventory
        [void]([wmiclass]'ROOT\ccm:SMS_Client').TriggerSchedule('{00000000-0000-0000-0000-000000000001}')
        }
    Catch {Write-Log -Message "Error: Failed to execute Hardware Inventory" -Source Execute-CMHardwareInventory -Severity 2}
}

Function Request-CMMachinePolicy {
    Try {
        Write-Log -Message "Info: Requesting Machine Policy" -Source Request-CMMachinePolicy
        [void]([wmiclass]'ROOT\ccm:SMS_Client').TriggerSchedule('{00000000-0000-0000-0000-000000000021}')
        }
    Catch {Write-Log -Message "Error: Failed to request Machine Policy" -Source Request-CMMachinePolicy -Severity 2}
}

Function Evaluate-CMMachinePolicy {
    Try {
        Write-Log -Message "Info: Evaluating machine policy" -Source Evaluate-CMMachinePolicy
        ([wmiclass]'ROOT\ccm:SMS_Client').TriggerSchedule('{00000000-0000-0000-0000-000000000022}')
        }
    Catch{Write-Log -Message "Error: Failed to evaluate Machine Policy" -Source Evaluate-CMMachinePolicy -Severity 2}
}

Function Scan-CMupdates {
    Try {
        Write-Log -Message "Info: Scanning for Software Updates" -Source Scan-CMupdates
        [void]([wmiclass]'ROOT\ccm:SMS_Client').TriggerSchedule('{00000000-0000-0000-0000-000000000113}')
        }
    Catch {Write-Log -Message "Error: Failed to scan for Software Updates" -Source Scan-CMupdates -Severity 2}
}

Function Evaluate-CMupdates {
    Try {
        Write-Log -Message "Info: Evaluating installation of Software Updates" -Source Evaluate-CMupdates
        [void]([wmiclass]'ROOT\ccm:SMS_Client').TriggerSchedule('{00000000-0000-0000-0000-000000000108}')
        }
    Catch {Write-Log -Message "Error: Failed to evaluate Software Updates" -Source Evaluate-CMupdates -Severity 2}
}

Function Evaluate-CMapplications {
    Try {
        Write-Log -Message "Info: Evaluating installation of applications" -Source Evaluate-CMapplications
        [void]([wmiclass]'ROOT\ccm:SMS_Client').TriggerSchedule('{00000000-0000-0000-0000-000000000121}')
        }
    Catch {Write-Log -Message "Error: Failed to evaluate applications" -Source Evaluate-CMapplications -Severity 2}
}


Function Refresh-CMPolicies {
    Try{
        Set-CMSite
        WaitThirtySeconds
        Reset-CMMachinePolicy
        WaitFiveMinutes            
        Request-CMMachinePolicy
        WaitThirtySeconds
        Evaluate-CMMachinePolicy
        WaitFiveMinutes
        Scan-CMupdates
        WaitThirtySeconds
        Execute-CMHardwareInventory
        Evaluate-CMupdates
        Evaluate-CMapplications
        WaitThirtySeconds
        Install-SCCMSoftwareUpdates
        Request-CMMachinePolicy
        WaitThirtySeconds
        Evaluate-CMapplications
        Install-SCCMSoftwareUpdates
        Evaluate-CMMachinePolicy
        
    }
	Catch
	{
		Write-Log -Message "Warning: failed to refresh SCCM policies" -Source Refresh-CMPolicies -Severity 2
	}
}
#endregion SCCM client schedules

#region SCCM PFE Remediation related functions
Function Get-PFEServer
{
	<#
		#	Created on:   	07.07.2017 11:47
		#	Created by:   	Mieszko Ślusarczyk
		#	Organization: 	International Paper
    .SYNOPSIS
    Get SCCM PFE Remediation Agent Server name.
    
    .DESCRIPTION
	The script will check the currently assigned SCCM primary site and display it's FQDN

    
    .EXAMPLE
    Get-PFEServer

    .DEPENDENT FUNCTIONS
    Write-Log

    #>
	If (Test-Path "HKLM:\SOFTWARE\Microsoft\Microsoft PFE Remediation for Configuration Manager")
	{
		Try
		{
			$script:PFEServer = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft PFE Remediation for Configuration Manager").PrimarySiteName
			If ($PFEServer)
			{
				Write-Log "Info: PFE server name is $PFEServer" -Source "Get-PFEServer"
			}
			Else
			{
				Write-Log "Error: Could not get PFE server name" -Severity 3 -Source "Get-PFEServer"
			}
			
		}
		Catch
		{
			Write-Log "Error: Could not get PFE server name" -Severity 3 -Source "Get-PFEServer"
		}
	}
	Else
	{
		Write-Log "Error: `"HKLM:\SOFTWARE\Microsoft\Microsoft PFE Remediation for Configuration Manager`" does not exist" -Severity 3 -Source "Get-PFEServer"
	}
	
}

Function Set-PFEServer
{
	<#
		#	Created on:   	07.07.2017 11:47
		#	Created by:   	Mieszko Ślusarczyk
		#	Organization: 	International Paper
    .SYNOPSIS
    Set SCCM PFE Remediation Agent Server name.
    
    .DESCRIPTION
	The script will assign PFE Remediation Agent with SCCM primary site and display it's FQDN

    
    .EXAMPLE
    Set-PFEServer

    .DEPENDENT FUNCTIONS
    Write-Log

    #>
	If ($PrimarySiteServer)
	{
		If (Test-Path "HKLM:\SOFTWARE\Microsoft\Microsoft PFE Remediation for Configuration Manager")
		{
			Try
			{
				
				Write-Log "Info: Setting PFE server name to $PrimarySiteServer" -Source "Get-PFEServer"
				Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Microsoft PFE Remediation for Configuration Manager" -Name PrimarySiteName -Value "$PrimarySiteServer"
				Try
				{
					Write-Log "Info: PFE server name changed, restarting PFERemediation service" -Source "Get-PFEServer"
					Restart-Service PFERemediation
				}
				Catch
				{
					Write-Log "Error: Failed restart PFERemediation service" -Severity 3 -Source "Get-PFEServer"
				}
			}
			Catch
			{
				Write-Log "Error: Failed to set PFE server name to $PrimarySiteServer" -Severity 3 -Source "Get-PFEServer"
			}
		}
		Else
		{
			Write-Log "Error: `"HKLM:\SOFTWARE\Microsoft\Microsoft PFE Remediation for Configuration Manager`" does not exist." -Severity 3 -Source "Get-PFEServer"
		}
	}
	Else
	{
		Write-Log "Error: No Primary Site Server FQDN detected" -Severity 3 -Source "Get-PFEServer"
	}
}

Function Update-PFEServerAssignment
{
	<#
		#	Created on:   	07.07.2017 11:47
		#	Created by:   	Mieszko Ślusarczyk
		#	Organization: 	International Paper
    .SYNOPSIS
    Check SCCM PFE Remediation Agent Server assignment.
    
    .DESCRIPTION
	The script will Check PFE Remediation Agent server assignment, and update the assignment (based on SCCM client assignment) if necessary

    
    .EXAMPLE
    Update-PFEServerAssignment

    .DEPENDENT FUNCTIONS
	Get-CMPrimarySiteServer
	Get-PFEServer
	Write-Log

    #>
	Get-CMPrimarySiteServer
	Get-PFEServer
	If ("$PFEServer" -ne "$PrimarySiteServer")
	{
		Write-Log "Info: Trying to update PFE server assignment" -Source "Update-PFEServerAssignment"
		Try
		{
			Set-PFEServer
		}
		Catch
		{
			Write-Log "Error: Failed to update PFE server assignment" -Severity 3 -Source "Update-PFEServerAssignment"
		}
	}
	Else
	{
		Write-Log "Info: PFE server assignment is up to date" -Source "Update-PFEServerAssignment"
	}
}

#endregion SCCM PFE Remediation related functions

#region ccmcache functions
Function Clean-CMCacheOldItems #region Clean-CMCacheOldItems
{
	$ccmCacheItems = `
	Try
	{
		get-wmiobject -query "SELECT * FROM CacheInfoEx" -namespace "ROOT\ccm\SoftMgmtAgent"
	}
	Catch
	{
		Write-Log "Error: Failed to get ccmcache items"
	}
	
	foreach ($ccmCacheItem in $ccmCacheItems)
	{
		[datetime]$ccmCacheItemLastReferencedDate = Parse-Date $ccmCacheItem.LastReferenced "yyyyMMddHHmmss" -Trim
		[guid]$ccmCacheItemCacheId = $ccmCacheItem.CacheId
		[string]$ccmCacheItemLocation = $ccmCacheItem.Location
		[int]$ccmCacheItemContentSize = $ccmCacheItem.ContentSize / 1024
		
		If (($ccmCacheItemLastReferencedDate -le (Get-Date).AddMonths(-1)) -or ($runningTaskSequence) ) # Older than a month or running in a task sequence
		{
			If (($ccmCacheItem.PersistInCache -eq 1) -and (!($runningTaskSequence))) # Don't delete persisted items if not running in a task sequence
			{
				Write-Log "$ccmCacheItemCacheId is persisted, skipping" -Source Clean-CMCacheOldItems
			}
			Else
			{
				Write-Log "Info: Deleting $ccmCacheItemCacheId from $ccmCacheItemLocation `($ccmCacheItemContentSize`MB`)" -Source Clean-CMCacheOldItems
				Try
				{
					[wmi]"ROOT\ccm\SoftMgmtAgent:CacheInfoEx.CacheId=`"$ccmCacheItemCacheId`"" | Remove-WmiObject
				}
				Catch
				{
					Write-Log "Error: Failed to delete $ccmCacheItemCacheId from $ccmCacheItemLocation" -Source Clean-CMCacheOldItems -Severity 2
				}
			}
		}
	}
} #endregion Clean-CMCacheOldItems

function Clean-CMCacheOrphanedItems #region Clean-CMCacheOrphanedItems
{
	Write-Log "Cleaning up orphaned folders in ccmcache" -Source Clean-CMCacheOrphanedItems
	$UsedFolders = $CacheElements | % { Select-Object -inputobject $_.Location }
	[string]$CCMCache = ([wmi]"ROOT\ccm\SoftMgmtAgent:CacheConfig.ConfigKey='Cache'").Location
	if ($CCMCache.EndsWith('ccmcache'))
	{
		Get-ChildItem($CCMCache) | ?{ $_.PSIsContainer } | WHERE { $UsedFolders -notcontains $_.FullName } | % { Remove-Item $_.FullName -recurse; $Cleaned++ }
	}
	If ($Cleaned -ge 1)
	{
		Write-Log "Info: Cleaned $Cleaned orphaned items from ccmcache" -Source Clean-CMCacheOrphanedItems
	}
	Else
	{
		Write-Log "Info: No orphaned items in ccmcache to clean" -Source Clean-CMCacheOrphanedItems
	}
	
} #endregion Clean-CMCacheOrphanedItems
#endregion ccmcache functions

# SIG # Begin signature block
# MIITkQYJKoZIhvcNAQcCoIITgjCCE34CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUxwOH0Sgo+YidYGXwqdRjoGlO
# wV+ggg7fMIIEFDCCAvygAwIBAgILBAAAAAABL07hUtcwDQYJKoZIhvcNAQEFBQAw
# VzELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExEDAOBgNV
# BAsTB1Jvb3QgQ0ExGzAZBgNVBAMTEkdsb2JhbFNpZ24gUm9vdCBDQTAeFw0xMTA0
# MTMxMDAwMDBaFw0yODAxMjgxMjAwMDBaMFIxCzAJBgNVBAYTAkJFMRkwFwYDVQQK
# ExBHbG9iYWxTaWduIG52LXNhMSgwJgYDVQQDEx9HbG9iYWxTaWduIFRpbWVzdGFt
# cGluZyBDQSAtIEcyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAlO9l
# +LVXn6BTDTQG6wkft0cYasvwW+T/J6U00feJGr+esc0SQW5m1IGghYtkWkYvmaCN
# d7HivFzdItdqZ9C76Mp03otPDbBS5ZBb60cO8eefnAuQZT4XljBFcm05oRc2yrmg
# jBtPCBn2gTGtYRakYua0QJ7D/PuV9vu1LpWBmODvxevYAll4d/eq41JrUJEpxfz3
# zZNl0mBhIvIG+zLdFlH6Dv2KMPAXCae78wSuq5DnbN96qfTvxGInX2+ZbTh0qhGL
# 2t/HFEzphbLswn1KJo/nVrqm4M+SU4B09APsaLJgvIQgAIMboe60dAXBKY5i0Eex
# +vBTzBj5Ljv5cH60JQIDAQABo4HlMIHiMA4GA1UdDwEB/wQEAwIBBjASBgNVHRMB
# Af8ECDAGAQH/AgEAMB0GA1UdDgQWBBRG2D7/3OO+/4Pm9IWbsN1q1hSpwTBHBgNV
# HSAEQDA+MDwGBFUdIAAwNDAyBggrBgEFBQcCARYmaHR0cHM6Ly93d3cuZ2xvYmFs
# c2lnbi5jb20vcmVwb3NpdG9yeS8wMwYDVR0fBCwwKjAooCagJIYiaHR0cDovL2Ny
# bC5nbG9iYWxzaWduLm5ldC9yb290LmNybDAfBgNVHSMEGDAWgBRge2YaRQ2XyolQ
# L30EzTSo//z9SzANBgkqhkiG9w0BAQUFAAOCAQEATl5WkB5GtNlJMfO7FzkoG8IW
# 3f1B3AkFBJtvsqKa1pkuQJkAVbXqP6UgdtOGNNQXzFU6x4Lu76i6vNgGnxVQ380W
# e1I6AtcZGv2v8Hhc4EvFGN86JB7arLipWAQCBzDbsBJe/jG+8ARI9PBw+DpeVoPP
# PfsNvPTF7ZedudTbpSeE4zibi6c1hkQgpDttpGoLoYP9KOva7yj2zIhd+wo7AKvg
# IeviLzVsD440RZfroveZMzV+y5qKu0VN5z+fwtmK+mWybsd+Zf/okuEsMaL3sCc2
# SI8mbzvuTXYfecPlf5Y1vC0OzAGwjn//UYCAp5LUs0RGZIyHTxZjBzFLY7Df8zCC
# BJ8wggOHoAMCAQICEhEh1pmnZJc+8fhCfukZzFNBFDANBgkqhkiG9w0BAQUFADBS
# MQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2lnbiBudi1zYTEoMCYGA1UE
# AxMfR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0EgLSBHMjAeFw0xNjA1MjQwMDAw
# MDBaFw0yNzA2MjQwMDAwMDBaMGAxCzAJBgNVBAYTAlNHMR8wHQYDVQQKExZHTU8g
# R2xvYmFsU2lnbiBQdGUgTHRkMTAwLgYDVQQDEydHbG9iYWxTaWduIFRTQSBmb3Ig
# TVMgQXV0aGVudGljb2RlIC0gRzIwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEK
# AoIBAQCwF66i07YEMFYeWA+x7VWk1lTL2PZzOuxdXqsl/Tal+oTDYUDFRrVZUjtC
# oi5fE2IQqVvmc9aSJbF9I+MGs4c6DkPw1wCJU6IRMVIobl1AcjzyCXenSZKX1GyQ
# oHan/bjcs53yB2AsT1iYAGvTFVTg+t3/gCxfGKaY/9Sr7KFFWbIub2Jd4NkZrItX
# nKgmK9kXpRDSRwgacCwzi39ogCq1oV1r3Y0CAikDqnw3u7spTj1Tk7Om+o/SWJMV
# TLktq4CjoyX7r/cIZLB6RA9cENdfYTeqTmvT0lMlnYJz+iz5crCpGTkqUPqp0Dw6
# yuhb7/VfUfT5CtmXNd5qheYjBEKvAgMBAAGjggFfMIIBWzAOBgNVHQ8BAf8EBAMC
# B4AwTAYDVR0gBEUwQzBBBgkrBgEEAaAyAR4wNDAyBggrBgEFBQcCARYmaHR0cHM6
# Ly93d3cuZ2xvYmFsc2lnbi5jb20vcmVwb3NpdG9yeS8wCQYDVR0TBAIwADAWBgNV
# HSUBAf8EDDAKBggrBgEFBQcDCDBCBgNVHR8EOzA5MDegNaAzhjFodHRwOi8vY3Js
# Lmdsb2JhbHNpZ24uY29tL2dzL2dzdGltZXN0YW1waW5nZzIuY3JsMFQGCCsGAQUF
# BwEBBEgwRjBEBggrBgEFBQcwAoY4aHR0cDovL3NlY3VyZS5nbG9iYWxzaWduLmNv
# bS9jYWNlcnQvZ3N0aW1lc3RhbXBpbmdnMi5jcnQwHQYDVR0OBBYEFNSihEo4Whh/
# uk8wUL2d1XqH1gn3MB8GA1UdIwQYMBaAFEbYPv/c477/g+b0hZuw3WrWFKnBMA0G
# CSqGSIb3DQEBBQUAA4IBAQCPqRqRbQSmNyAOg5beI9Nrbh9u3WQ9aCEitfhHNmmO
# 4aVFxySiIrcpCcxUWq7GvM1jjrM9UEjltMyuzZKNniiLE0oRqr2j79OyNvy0oXK/
# bZdjeYxEvHAvfvO83YJTqxr26/ocl7y2N5ykHDC8q7wtRzbfkiAD6HHGWPZ1BZo0
# 8AtZWoJENKqA5C+E9kddlsm2ysqdt6a65FDT1De4uiAO0NOSKlvEWbuhbds8zkSd
# wTgqreONvc0JdxoQvmcKAjZkiLmzGybu555gxEaovGEzbM9OuZy5avCfN/61PU+a
# 003/3iCOTpem/Z8JvE3KGHbJsE2FUPKA0h0G9VgEB7EYMIIGIDCCBAigAwIBAgIT
# UgAAABxN1D0InJGIbQAAAAAAHDANBgkqhkiG9w0BAQsFADA/MRMwEQYKCZImiZPy
# LGQBGRYDY29tMRYwFAYKCZImiZPyLGQBGRYGaXBhcGVyMRAwDgYDVQQDEwdJUFN1
# YkNBMB4XDTE3MDQxNzE1MTcyNVoXDTIwMDQxNjE1MTcyNVowHTEbMBkGA1UEAxMS
# SW50ZXJuYXRpb25hbFBhcGVyMIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQCN
# DHhqB9jdBRakFFWXJnNdj+ZuOYmN9g+U5v/xsgDam173evS5GV0Zj5w6yFmag1Fj
# kjhyQmiZcilVrb2z+CU4mgmbkduCZai5/1NN6wiHqTsQyTlwyNCoGRTmHAzWdRgs
# OJ4SHcW5YWKDSicThtUlEiCipueqLf6J55W0vSnvEwIDAQABo4ICuTCCArUwPgYJ
# KwYBBAGCNxUHBDEwLwYnKwYBBAGCNxUIgfvTWYSZhyiDsYMahKiqM4SB+SaBGYW0
# wReFytkCAgFkAgEMMBMGA1UdJQQMMAoGCCsGAQUFBwMDMAsGA1UdDwQEAwIHgDAM
# BgNVHRMBAf8EAjAAMBsGCSsGAQQBgjcVCgQOMAwwCgYIKwYBBQUHAwMwHQYDVR0O
# BBYEFKkBcsWCUlbQsQuvlGn8TdH+HNybMB8GA1UdIwQYMBaAFO/UWWdgalRzmf9r
# cPCTbIgUSaTrMIHxBgNVHR8EgekwgeYwgeOggeCggd2GK2h0dHA6Ly9jZXJ0Lmlw
# YXBlci5jb20vQ2VydERhdGEvSVBTdWJDQS5jcmyGga1sZGFwOi8vL0NOPUlQU3Vi
# Q0EsQ049UzAyQVNDQSxDTj1DRFAsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMs
# Q049U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1pcGFwZXIsREM9Y29tP2Nl
# cnRpZmljYXRlUmV2b2NhdGlvbkxpc3Q/YmFzZT9vYmplY3RDbGFzcz1jUkxEaXN0
# cmlidXRpb25Qb2ludDCB8QYIKwYBBQUHAQEEgeQwgeEwNwYIKwYBBQUHMAKGK2h0
# dHA6Ly9jZXJ0LmlwYXBlci5jb20vQ2VydERhdGEvSVBTdWJDQS5jcnQwgaUGCCsG
# AQUFBzAChoGYbGRhcDovLy9DTj1JUFN1YkNBLENOPUFJQSxDTj1QdWJsaWMlMjBL
# ZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1cmF0aW9uLERDPWlw
# YXBlcixEQz1jb20/Y0FDZXJ0aWZpY2F0ZT9iYXNlP29iamVjdENsYXNzPWNlcnRp
# ZmljYXRpb25BdXRob3JpdHkwDQYJKoZIhvcNAQELBQADggIBAGryNXWYcEAErDeM
# nIORRnlVZchJd2qx6SdEmzX3YFYcoAqnLS6WMGyWLI6Ak2Yirf3xGrzanK42Jlkp
# D+tP0fBzAUXlcA9cb3w/HIvWFdM+OZqI1YCnLv7tDeJu6cmN9IFhbaowcYC3VgBM
# w8E3nVWfk99OxeGu/2Q0qZl2ry0LQfw+oI8smPMQekby1rR75EF2phMFef8nQtAz
# 5NeLWxAU7AgM1vWig4N8LtoaUK8BmRLwV2D39Iddhds0XkWxQfyjYtAhhGH4oMq3
# L7eFZ0Hi6psuAE3iWFhgXqgV+rlvX6f4Itt+UiktfkTxhuMAzQeNNhaKXwloD7/F
# 122bD7k0mORD9xx1+jHPbpC0FgfXFccNTfyBTagIylaKDtrEdg0+gUGNCBhlgZfX
# cM8QlT4aJ1IRB98mv58KbzNUPIVG8C+5bUCPhrFG4yo4AboP9706VplR/g8lT/nY
# vOD8uqs+orXNlJ4AXB9Wxn6iik4H5ZyioBF5Zhma8yG/Lb1pzXvmZtwwxA61JOOD
# HO8HQMJQDow4mkQFxs1HRx2k+UWzx0NfvnbzyR4ip39uS6UqXKMrLzIJ2/Gczm/i
# cmNqo9xRc+kXjJXjFFgnJguG85OY9ALrpBfCUvqf1N27DnCHsfHCLLXny6SU+MaH
# D69vPYBagfulKZqFHdaEdy/ZMPzoMYIEHDCCBBgCAQEwVjA/MRMwEQYKCZImiZPy
# LGQBGRYDY29tMRYwFAYKCZImiZPyLGQBGRYGaXBhcGVyMRAwDgYDVQQDEwdJUFN1
# YkNBAhNSAAAAHE3UPQickYhtAAAAAAAcMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQ5J0GUSkzU
# oDkTwqSUWI+e231LzTANBgkqhkiG9w0BAQEFAASBgE4TL+mUROsnx1Ft13PITyeG
# Bch9KF4076WuQp+e9ngaqCSWgl9QayFuoJHvNq4zZCTaz24Pyr0EVYR7eRApVO7a
# X/XVu3BObtLQ4sibxYvv3iwvWQBRsPYhOAlC0XY4xeCozbq0iXNmM6cxlzH2eday
# TVVScS4xUsSoz1+5Y90+oYICojCCAp4GCSqGSIb3DQEJBjGCAo8wggKLAgEBMGgw
# UjELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExKDAmBgNV
# BAMTH0dsb2JhbFNpZ24gVGltZXN0YW1waW5nIENBIC0gRzICEhEh1pmnZJc+8fhC
# fukZzFNBFDAJBgUrDgMCGgUAoIH9MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEw
# HAYJKoZIhvcNAQkFMQ8XDTE3MDczMTE0MzIyMVowIwYJKoZIhvcNAQkEMRYEFKpM
# 1mPMgQWSpxZvfu7kMNNoAG7pMIGdBgsqhkiG9w0BCRACDDGBjTCBijCBhzCBhAQU
# Y7gvq2H1g5CWlQULACScUCkz7HkwbDBWpFQwUjELMAkGA1UEBhMCQkUxGTAXBgNV
# BAoTEEdsb2JhbFNpZ24gbnYtc2ExKDAmBgNVBAMTH0dsb2JhbFNpZ24gVGltZXN0
# YW1waW5nIENBIC0gRzICEhEh1pmnZJc+8fhCfukZzFNBFDANBgkqhkiG9w0BAQEF
# AASCAQB1BNODvA64AJC4Ya7Md3NsYS3Jv8nnG+xZ9mT9jZs5lJ6nivbvcyWR0lEX
# C7cKLTYcZ5EpqNfPL7VHv34eWXxVcxi6yniohJBBO1FGNrRvFXfLtWMIMQZbMxcE
# 8J+rOqW/2LRLgRX6dzgQW0Ls5AyU/ZSAw0/nuL+Mv85CP+F6buxqu6fJQKhcAWeV
# JbAhSEUrc8OqCtRx3n+Kxm0sUlxJhFfONNBGonRUUbaiQdC4aWT4bZPu+2xCFKyE
# QvL8p2ZbzzIaFrznGYCqAl2jq5+/RvFK5tmTIcj1rvL38v1IOnY3pEYNRqY1k4Pl
# xwpPFhPy7cuv/F1e5ANUEGx0l4FJ
# SIG # End signature block
