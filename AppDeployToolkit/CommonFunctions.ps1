Function WaitOneMinute #region WaitOneMinute
{
	#Wait one minute
	Write-Log -Message "Waiting 60 Seconds"
	Start-Sleep -Seconds 60
}#endregion WaitOneMinute

Function Set-ScheduledTask #region Set-ScheduledTask
{
<#
.SYNOPSIS
Creates or removes a scheduled task on local computer or remote computer
.DESCRIPTION
Creates a scheduled task on local computer or remote computer using COM object and XML text.
Returns $True or $False

Simply create a task in "Task Scheduler" and export it to XML.
Then use that XML from inside a Here-String [@" "@] or from the xml file

CAVEAT: Only handles task folders ONE level deep
.PARAMETER TaskXmlContent
Contents of a Task XML file as [String], [XmlDocument] or [xml].
.PARAMETER TaskNamePath
Name of the task to be created/removed and/or TaskFolder to be created/removed
Can also contain a TaskFolder. TaskFolder is auto-created if needed.
[TODO SomeDay: Optional *if* TaskXmlContent Description's Last line is @@@MyTaskFolder\MyTaskName]
.PARAMETER TaskUser
Default: $null
TaskUser can be $null if task is running as SYSTEM or a group (e.g. BUILT-IN\USERS)
.PARAMETER TaskPwd
Default: $null
TaskPwd can be $null if task is running as SYSTEM or a group (e.g. BUILT-IN\USERS)
.PARAMETER ComputerName
Name of Computer where to create the task. Default: localhost
.PARAMETER Remove
Remove the task specified by TaskNamePath [and ComputerName]
.PARAMETER RemoveTaskFolderIfEmpty
Works with -Remove parameter. Delete folder holding the targeted task.
.PARAMETER ContinueOnError
Continue if an error is encountered. Default: $false.
.EXAMPLE
[XML]$NewTaskXmlFileContent = Get-Content "C:\stuff\ExportedTask.xml"
Set-ScheduledTask -TaskNamePath $NewTaskName -TaskXmlContent $NewTaskXmlFileContent -ComputerName $Computer
Creates task in Root Task folder (\) from "C:\stuff\ExportedTask.xml"
.EXAMPLE
Set-ScheduledTask -TaskNamePath $NewTaskName -Remove -ComputerName $Computer
Removes task from Root Task folder (\)
.EXAMPLE
[xml]$NewTaskXmlContent = @" …<xml stuff>… "@
Set-ScheduledTask -TaskNamePath "MyTaskFolder\MyTaskName" -TaskXmlContent $NewTaskXmlContent
Creates task in MyTaskFolder Task folder (\MyTaskFolder) from [xml]$NewTaskXmlContent
CAVEAT: Cannot create more than ONE level deep.
.EXAMPLE
Set-ScheduledTask -TaskNamePath "MyTaskFolder\MyTaskName" -Remove
Set-ScheduledTask -TaskNamePath "\MyTaskFolder\MyTaskName" -Remove
Removes task from MyTaskFolder Task folder (\MyTaskFolder)
CAVEAT: Cannot delete from more than ONE level deep.
.EXAMPLE
Set-ScheduledTask -TaskNamePath "MyTaskFolder\MyTaskName" -Remove -RemoveTaskFolderIfEmpty
Removes task from MyTaskFolder Task folder (\MyTaskFolder)
Removes MyTaskFolder Task folder if it then becomes empty
CAVEAT: Cannot delete from more than ONE level deep.
.EXAMPLE
Set-ScheduledTask -TaskNamePath "\MyTaskFolder\*" -Remove -RemoveTaskFolderIfEmpty
Removes MyTaskFolder Task folder and its Tasks. (BE CAREFUL!!)
CAVEAT: Cannot delete more than ONE level deep.
.NOTES
Version 1.0 (22-APR-2015)
Denis St-Pierre (Ottawa, Canada)
LIMITATION: cannot handle more than one task folder deep
Based on http://psappdeploytoolkit.codeplex.com
For syntax use: Get-Help Set-ScheduledTask
#>
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $True)]
		[ValidateNotNullorEmpty()]
		[string]$TaskNamePath = "",
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		$TaskXmlContent,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[Switch]$Remove,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[string]$ComputerName = "localhost",
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[string]$TaskUser = $null,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullorEmpty()]
		[string]$TaskPwd = $null,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[switch]$RemoveTaskFolderIfEmpty,
		[Parameter(Mandatory = $false)]
		[ValidateNotNullOrEmpty()]
		[bool]$ContinueOnError = $true
	)
	
	Begin
	{
		## Get the name of this function and write header
		[string]${CmdletName} = $PSCmdlet.MyInvocation.MyCommand.Name
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -CmdletBoundParameters $PSBoundParameters -Header
		Remove-Variable TaskFolderPath -ErrorAction SilentlyContinue #Needed for debug to make sense
		
		Try
		{
			[System.__ComObject]$ScheduleObj = New-Object -ComObject ("Schedule.Service") #Need Admin do to this!
			$ScheduleObj.connect($ComputerName)
			[System.__ComObject]$TaskRootFolderObj = $ScheduleObj.GetFolder("\")
		}
		Catch
		{
			[string]$exceptionMessage = "$($_.Exception.Message) ($($_.ScriptStackTrace))"
			[String]$message = "ERROR: Cannot connect to [Schedule.Service]. Are we running elevated? $exceptionMessage"
			If ($ContinueOnError)
			{
				Write-Log $message -Source ${CmdletName}
				return $false #exit function
			}
			Else
			{
				Throw $message
			}
		}
		#TODO SomeDay: Retrieve TaskNamePath from TaskXmlContent Description's Last line is @@@MyTaskFolder\MyTaskName
		Write-Log "TaskNamePath is [$TaskNamePath]" -Source ${CmdletName}
		
		If ($TaskNamePath -match '\\')
		{
			#Get $TaskName and $TaskFolderPath from TaskNamePath
			[string]$TaskName = [System.IO.Path]::GetFileName($TaskNamePath)
			[string]$TaskFolderPath = [System.IO.Path]::GetDirectoryName($TaskNamePath)
		}
		Else
		{
			[string]$TaskName = $TaskNamePath
			#	[string]$TaskFolderPath = "\"	#RootTaskFolder
		}
	}
	Process
	{
		#Remove ALL Tasks in ONE Task Folder (Except ROOT Task Folder)
		If (($Remove) -and ($TaskName -eq "*"))
		{
			If ($TaskRootFolderObj.Path -eq $TaskFolderObj.Path)
			{
				[String]$message = "Will NOT delete all tasks in [ROOT Task Folder]. Allowing this would break too many things!"
				If ($ContinueOnError)
				{
					Write-log $Message -Source ${CmdletName}
					return $false #exit function
				}
				Else
				{
					Throw "ERROR: $Message"
				}
			}
			
			Try
			{
				[System.__ComObject]$TaskFolderObj = $ScheduleObj.GetFolder($TaskFolderPath)
			}
			Catch
			{
				[string]$exceptionMessage = "$($_.Exception.Message)] ($($_.ScriptStackTrace))"
				[String]$message = "Task folder [$TaskFolderPath] does not exist. `nNothing to Delete. $exceptionMessage"
				If ($ContinueOnError)
				{
					Write-log $message -Source ${CmdletName}
					return $True #exit function
				}
				Else
				{
					Throw "ERROR: $message"
				}
			}
			
			Write-Log "Deleting Task folder [$($TaskFolderObj.Path)] regardless of the number of tasks in the folder" -Source ${CmdletName}
			Try
			{
				[System.__ComObject]$AllTasks = $TaskFolderObj.GetTasks(1) #1= include hidden tasks too
			}
			Catch
			{
				[string]$exceptionMessage = "$($_.Exception.Message)] ($($_.ScriptStackTrace))"
				[String]$message = "Unable to get Tasks in Task folder [$TaskFolderPath] $exceptionMessage"
				If ($ContinueOnError)
				{
					Write-log $message -Source ${CmdletName}
					return $false #exit function
				}
				Else
				{
					Throw "ERROR: $message"
				}
			}
			[Int32]$TotalNumTasks = $AllTasks.count
			Write-log "[$TaskFolderPath] has $TotalNumTasks task(s)" -Source ${CmdletName}
			If ($TotalNumTasks -gt 0)
			{
				ForEach ($Task in $AllTasks)
				{
					Try
					{
						$Task.Stop(0) #Just in case, should test if running first
						Start-Sleep -Seconds 1
						$TaskFolderObj.DeleteTask($Task.Name, 0)
						Write-Log ("Task [$($Task.Name)] was deleted") -Source ${CmdletName}
					}
					Catch
					{
						[String]$message = "Cannot delete task [$($Task.Name)]. Might not exist or stopped"
						If ($ContinueOnError)
						{
							Write-log $message -Source ${CmdletName}
						}
						Else
						{
							Throw "ERROR: $message"
						}
					}
				}
				
				#Are they all gone?
				[System.__ComObject]$AllTasks = $TaskFolderObj.GetTasks(1) #1= include hidden tasks too
				If ($($AllTasks.Count) -ne 0)
				{
					[string]$ErrMess = "ERROR: Not all tasks have been deleted from the task folder."
					ForEach ($Task in $AllTasks)
					{
						$ErrMess = $ErrMess + "`nTask [$($Task.Name)]"
					}
					If ($ContinueOnError)
					{
						Write-log $ErrMess -Source ${CmdletName}
						return $false #exit function
					}
					Else
					{
						Throw "ERROR: $ErrMess"
					}
				}
				Else
				{
					#Write-log "INFO:No Tasks to delete in task folder[$SubFolderPath]." -Source ${CmdletName}
				}
			}
			Else
			{
				Write-Log "INFO: No tasks to delete in task folder [$SubFolderPath]." -Source ${CmdletName}
			}
			
			#Delete the Task Folder
			If ($RemoveTaskFolderIfEmpty)
			{
				#CAVEAT : you must use .DeleteFolder method with the PARENT of the task folder you want to delete
				#CAVEAT2: I didn't bother to add code for subfolders b/c I didn't need it.
				$SubFolderName = Split-Path -Path $TaskFolderObj.Path -Leaf
				$ParentFolderPath = Split-Path -Path $TaskFolderObj.Path -Parent
				Try
				{
					[System.__ComObject]$ParentFolderObj = $ScheduleObj.GetFolder($ParentFolderPath)
				}
				Catch
				{
					[string]$exceptionMessage = "$($_.Exception.Message) ($($_.ScriptStackTrace))"
					Throw "Task folder [${ParentFolderPath}] does not exist. `nNothing to Delete. $exceptionMessage"
					Return $true #exit Function
				}
				#No need to check for left-over tasks. We did this already above
				
				Try
				{
					$ParentFolderObj.DeleteFolder($SubFolderName, $null)
					Write-Log "Task folder [$SubFolderName] was deleted" -Source ${CmdletName}
					return $true #exit Function
				}
				Catch
				{
					$exceptionMessage = "$($_.Exception.Message) ($($_.ScriptStackTrace))"
					[String]$message = "Unable to delete Task folder [$SubFolderName] $exceptionMessage"
					If ($ContinueOnError)
					{
						Write-log $message -Source ${CmdletName}
						return $false #exit Function
					}
					Else
					{
						Throw "ERROR: $message"
					}
				}
			}
			Else
			{
				#Write-log "Not attempting to delete Task Folder [$($TaskFolderObj.Path)]" -Source ${CmdletName}
				return $true #exit Function
			}
			
		}
		
		#Remove ONE Task
		If (($Remove) -and ($TaskName -ne ""))
		{
			Try
			{
				[System.__ComObject]$TaskFolderObj = $ScheduleObj.GetFolder($TaskFolderPath)
			}
			Catch
			{
				[String]$message = "Task folder [${TaskFolderPath}] does not exist. `nNothing to Delete."
				If ($ContinueOnError)
				{
					Write-log $message -Source ${CmdletName}
					return $true #exit Function
				}
				else { Throw "ERROR: $message" }
			}
			
			Write-Log ("Task [$TaskName] will be removed") -Source ${CmdletName}
			Try
			{
				[System.__ComObject]$task = $TaskFolderObj.gettask($TaskName)
				$Task.Stop(0) #Stop the task, Just in case
				Start-Sleep -Seconds 1
				$TaskFolderObj.DeleteTask($TaskName, 0)
				Write-Log ("Task [$TaskName] was deleted") -Source ${CmdletName}
			}
			Catch
			{
				[String]$message = "INFO:Cannot delete task [$TaskName]. It might not exist."
				Write-log $message -Source ${CmdletName}
				return $True #exit Function
			}
			
			#TODO: Check if the Task is still in $TaskFolderObj or not ( use gettasks() ?)
			
			If ($RemoveTaskFolderIfEmpty)
			{
				If ($TaskRootFolderObj.Path -eq $TaskFolderObj.Path)
				{
					[String]$message = "INFO: Cannot delete [ROOT task] folder."
					If ($ContinueOnError)
					{
						Write-log $message -Source ${CmdletName}
					}
					Else
					{
						Throw "ERROR: $message"
					}
				}
				Else
				{
					Try
					{
						#CAVEAT : you must use .DeleteFolder method with the PARENT of the task folder you want to delete
						#CAVEAT2: I didn't bother to add code for subfolders b/c I didn't need it.
						$SubFolderName = Split-Path -Path $TaskFolderObj.Path -Leaf
						$ParentFolderPath = Split-Path -Path $TaskFolderObj.Path -Parent
						
						Try
						{
							[System.__ComObject]$ParentFolderObj = $ScheduleObj.GetFolder($ParentFolderPath)
						}
						Catch
						{
							[string]$exceptionMessage = "$($_.Exception.Message) ($($_.ScriptStackTrace))"
							Throw "Cannot Get Task folder [${ParentFolderPath}] in order to delete $SubFolderName. $exceptionMessage"
							Return $false #exit Function
						}
						#Is $TaskFolder empty?
						[System.__ComObject]$AllTasks = $TaskFolderObj.GetTasks(1) #1= include hidden tasks too
						If ($($AllTasks.Count) -eq 0)
						{
							$ParentFolderObj.DeleteFolder($SubFolderName, $null)
							Write-Log "Task folder [$SubFolderName] was deleted" -Source ${CmdletName}
						}
						Else
						{
							Write-Log "INFO: Cannot delete Task folder [$SubFolderName]. It still contains $($AllTasks.Count) task(s)." -Source ${CmdletName}
						}
						return $True #exit Function
					}
					Catch
					{
						$exceptionMessage = "$($_.Exception.Message) ($($_.ScriptStackTrace))"
						[String]$message = "Unable to delete Task folder [$SubFolderName] $exceptionMessage"
						If ($ContinueOnError)
						{
							Write-log $message -Source ${CmdletName}
							return $false #exit Function
						}
						Else
						{
							Throw "ERROR: $message"
						}
					}
				} #Else
			}
			Else
			{
				#Write-log "Not attempting to delete Task Folder [$($TaskFolderObj.Path)]" -Source ${CmdletName}
				return $true #exit Function
			}
		}
		
		#Create Task
		If ($TaskName -eq "*")
		{
			[String]$message = "Cannot create task named $TaskName"
			If ($ContinueOnError)
			{
				Write-log $message -Source ${CmdletName}
				return $false #exit Function
			}
			Else
			{
				Throw "ERROR: $message"
			}
		}
		
		If (-not ($TaskXmlContent))
		{
			[String]$message = "Cannot create task without -TaskXmlContent parameter"
			If ($ContinueOnError)
			{
				Write-log $message -Source ${CmdletName}
				Return $false #exit Function
			}
			Else
			{
				Throw "ERROR: $message"
			}
		}
		
		If ($TaskFolderPath)
		{
			#Create TaskFolder if needed
			Write-Log "Creating Task folder [${TaskFolderPath}] If needed." -Source ${CmdletName}
			Try
			{
				$TaskRootFolderObj.CreateFolder($TaskFolderPath)
			}
			Catch { } #ignore already exists error
			Try
			{
				[System.__ComObject]$TaskFolderObj = $ScheduleObj.GetFolder($TaskFolderPath)
			}
			Catch
			{
				[String]$message = "Task folder [${TaskFolderPath}] does not exist. Cannot create task in non-existing task folder."
				If ($ContinueOnError)
				{
					Write-Log $message -Source ${CmdletName}
					Return $false
				}
				Else
				{
					Throw "ERROR: $message"
				}
			}
		}
		Else
		{
			#Task will be created in the "Root" Task Folder
			[System.__ComObject]$TaskFolderObj = $TaskRootFolderObj
		}
		
		#Creating task (In sub folder if needed)
		[System.__ComObject]$NewTask = $ScheduleObj.NewTask($null) #Create blank task
		If ($($TaskXmlContent.gettype().name) -eq "XmlDocument")
		{
			$NewTask.XmlText = $TaskXmlContent.OuterXml -as [string] #load XmlText property
		}
		ElseIf ($($TaskXmlContent.gettype().name) -eq "String")
		{
			$NewTask.XmlText = $TaskXmlContent #load XmlText property
		}
		Else
		{
			[String]$message = " -TaskXmlContent as [$($TaskXmlContent.gettype().name)] is not supported. Please cast as [XmlDocument] or [string]."
			If ($ContinueOnError)
			{
				Write-log $message -Source ${CmdletName}
				return $false #exit Function
			}
			Else
			{
				Throw "ERROR: $message"
			}
		}
		
		#It can overwrite an existing task just fine provided that the task folder exists already
		# but just in case…
		Try
		{
			Write-Log "Creating Task [$TaskName]…" -Source ${CmdletName}
			$RegistrationResult = $TaskFolderObj.RegisterTaskDefinition($TaskName, $NewTask, 6, $TaskUser, $TaskPwd, 1, $null)
			#Write-log "DEV: $RegistrationResult" -Source ${CmdletName}
		}
		Catch
		{
			[string]$exceptionMessage = "$($_.Exception.Message) ($($_.ScriptStackTrace))"
			[String]$message = "Unable to import task [$TaskName] $exceptionMessage"
			If ($ContinueOnError)
			{
				Write-log $message -Source ${CmdletName}
				return $false
			}
			Else
			{
				throw "ERROR: $message"
			}
		}
		#$RegistrationResult is a Massive block of text but not as [string]
		#[String]$RegistrationResultString = Out-String -InputObject $RegistrationResult
		#return $RegistrationResultString
		Return $true
	}
	End
	{
		Write-FunctionHeaderOrFooter -CmdletName ${CmdletName} -Footer
	}
} #endregion Set-ScheduledTask

Function Check-PendingReboot #region Check-PendingReboot
{
	If ((Get-PendingReboot).IsSystemRebootPending)
	{
		Write-Log -Message "The machine needs to be restarted before attempting installation" -Source Check-PendingReboot
		If ($runningTaskSequence)
		{
			Write-Log -Message "The machine is running a task sequence" -Source Check-PendingReboot
			$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
			[bool]$tsenv.Value("_SMSTSRebootRequired") = $true
			Show-DialogBox -Title 'Restart needed' -Text 'The machine needs to be restarted before attempting installation' -Icon 'Exclamation' -Timeout 3600 -Buttons "OK"
		}
		Else
		{
			Show-InstallationRestartPrompt -Countdownseconds 3600 -CountdownNoHideSeconds 600
		}
        Exit-Script -ExitCode 69010
	}
}#endregion Check-PendingReboot

Function Close-Processes #region Close-Processes
{
	$processList = $null; $processes | %{ $processList = $processList + $_ + ',' }; $processList = $processList.TrimEnd(","); $processList
	
	ForEach ($process in $processes)
		{
			If (Get-Process -Name $process -ea 0) { $runningProcesses = $process + "," + $runningProcesses }
		}
	If ($runningProcesses)
	{
		$runningProcesses = $runningProcesses.TrimEnd(",")
		Show-InstallationWelcome -CloseApps $processList -CloseAppsCountdown 3600 -BlockExecution -AllowDefer -DeferTimes 2 -CheckDiskSpace -PersistPrompt
	}
	Else
	{
		Show-InstallationWelcome -CheckDiskSpace -CloseApps $processList -BlockExecution
	}
	
}#endregion Close-Processes

function Get-ItemType #region Get-ItemType
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$Path
	)
	
	if ((Get-ChildItem ("$Path")).PSIsContainer)
	{
		$Script:ItemType = "Folder"
	}
	else
	{
		$Script:ItemType = "File"
	}
	Write-Log "Debug: $Path is a $ItemType" -Source Get-ItemType -DebugMessage
}#region Get-ItemType

function Get-ItemsPaths #region Get-ItemsPaths
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$Path
	)
	Get-ItemType -Path $Path
	
	If ($ItemType -eq "Folder")
	{
		Get-ChildItem "$Path" | Where-Object { $_.PSIsContainer } | Select-Object FullName -ExpandProperty FullName
	}
	ElseIf ($ItemType -eq "File")
	{
		Get-ChildItem "$Path" | Where-Object { ! $_.PSIsContainer } | Select-Object FullName -ExpandProperty FullName
	}
}#endregion Get-ItemsPaths

function Get-ItemSize #region Get-ItemSize
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$Path
	)
	
	Try
	{
		"{0:N2} MB" -f ((Get-ChildItem $Path -Recurse | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB) | Write-Log -Source Get-ItemSize -DebugMessage -PassThru
	}
	Catch { Write-Log "Warning: Failed to calculate size" -Severity 2 -Source Get-ItemSize} 
}#endregion Get-ItemSize

function Remove-TempFiles
{
	param
	(
		[parameter(Mandatory = $true)]
		[string]$Path
	)
	
	$WindowsTempItems = Get-ItemsPaths -Path $Path
	foreach ($WindowsTempItem in $WindowsTempItems)
	{
		$WindowsTempItemSize = Get-ItemSize -Path $WindowsTempItem
		If ($WindowsTempItemSize -gt "0")
		{
			Write-Log "Info: Trying to delete $WindowsTempItem ($WindowsTempItemSize)"
		}
		Else
		{
			Write-Log "Info: Trying to delete $WindowsTempItem"
		}
		try
		{
			If ($ItemType -eq 'File')
			{
				Write-Log "Debug: $WindowsTempItem is a file" -Source Remove-TempFiles -DebugMessage
				Remove-File -Path $WindowsTempItem -Recurse
			}
			ElseIf ($ItemType -eq 'Folder')
			{
				Write-Log "Debug: $WindowsTempItem is a folder" -Source Remove-TempFiles -DebugMessage
				Remove-Folder -Path $WindowsTempItem
			}
		}
		catch
		{
			Write-Log "Error: Failed to delete $WindowsTempItem"
		}
	}
	
}

# SIG # Begin signature block
# MIITkQYJKoZIhvcNAQcCoIITgjCCE34CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU/gdF+Z0l1kMIds+Fm/8su8eC
# Isyggg7fMIIEFDCCAvygAwIBAgILBAAAAAABL07hUtcwDQYJKoZIhvcNAQEFBQAw
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQtLZlalsH7
# UlYrpVsWohzqV5ivXzANBgkqhkiG9w0BAQEFAASBgGUGEgjs9Ul/8HTZze1qIxY2
# J6U1zp5df1S+6l2U9Xg+Xj5l4wuUVGoRNaCxoRCpwmCjiiIN8qFVcVpZ0ZkMLDsW
# k0S2t91+azJyoa5RDZTUwqJdovLgqJiir5skBNHLB4UeSAvpEA1cagxWVRLQg5t7
# zD07iPeZP23R/4ogFJofoYICojCCAp4GCSqGSIb3DQEJBjGCAo8wggKLAgEBMGgw
# UjELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExKDAmBgNV
# BAMTH0dsb2JhbFNpZ24gVGltZXN0YW1waW5nIENBIC0gRzICEhEh1pmnZJc+8fhC
# fukZzFNBFDAJBgUrDgMCGgUAoIH9MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEw
# HAYJKoZIhvcNAQkFMQ8XDTE3MDgwMTA4NDIzNVowIwYJKoZIhvcNAQkEMRYEFGJf
# 96rouApOn6gLqNqe9g0U3tcyMIGdBgsqhkiG9w0BCRACDDGBjTCBijCBhzCBhAQU
# Y7gvq2H1g5CWlQULACScUCkz7HkwbDBWpFQwUjELMAkGA1UEBhMCQkUxGTAXBgNV
# BAoTEEdsb2JhbFNpZ24gbnYtc2ExKDAmBgNVBAMTH0dsb2JhbFNpZ24gVGltZXN0
# YW1waW5nIENBIC0gRzICEhEh1pmnZJc+8fhCfukZzFNBFDANBgkqhkiG9w0BAQEF
# AASCAQCTos/JbK4EwhD/LDzTWXi7ip25qGdwXinXZ5krWteZZYv/fP0uVDmnfxe+
# JqnJHYSNo4lUy9TcaRwW54sjIOkkaarmwBOJ3VppBYOLFeDprdvcodc6rleECZo/
# sJUzZ8rm9iPMyBEoJLfWvbnVWQ4YBNzMrTgV0ZEDWvbMe8tKz4SacwElnBZ+Qu34
# sIaBvAIfoNxyhqFQgIf1dmbH4zK1xryCaVEhpX5j3q2wz72k0/AcaW2ZB1udjTzU
# hOKPiDK+fJ1GQcDRXwcZjImdAmjqD98FPHeUcZG4l/QXsX2eUo8hAnJ91iIwdeHQ
# 1i4O+0gG/NIKkFPlGumGA8KoQB0u
# SIG # End signature block
