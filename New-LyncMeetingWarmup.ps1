<#
.SYNOPSIS
	Depending upon the parameters passed to it, this script will either execute a PowerShell command to
"warmup" your Lync-SfB meetings (activating a napping IIS application pool), or create the required
Scheduled Tasks to automate the process.

.DESCRIPTION
	Depending upon the parameters passed to it, this script will either execute a PowerShell command to
"warmup" your Lync-SfB meetings (activating a napping IIS application pool), or create the required
Scheduled Tasks to automate the process.

It requires Server 2012 or later for full functionality. Refer the script's blog page for how you can use
it with Server 2008 or 2008 R2.

It is based upon (and 100% inspired by) this EXCELLENT post and suggestion by Drago Totev:
http://www.lynclog.com/2013/12/user-might-experince-delay-when-join.html


.NOTES
    Version				: 1.5
	Date				: 12th November 2019
	Author    			: Greig Sheridan
	There are lots of credits at the bottom of the script

	Revision History 	:
				v1.5: 12th November 2019
					Added test for Server 2019.
					Added my auto-update code.

				v1.4: 3rd March 2018
					Thank you @TrevorAMiller for pointing out MS changed the versioning in Server 2016 between Preview and GA. Updated test.

				v1.3: 10 February 2015
					Whoops: fixed a tiny typo in the updated version test that broke it for Win 8.1 & Server 2012 R2

				v1.2: 7 February 2015
					My colleague Tristan highlighted that the script fails to generate the tasks if your o/s is Server 2008.
					I've amended it to work as best it can with 2008. It can't create the tasks - that step you'll
						still need to do manually, however the tasks you create can still just call this script and it will
						"warm up" your pools for you. I've added more how-to guidance on the blog post.
					Neatened the CmdletBinding. Makes for a more accurate "get-help" output & blocks unsupported "-WhatIf" and "-Confirm".
					Tweaked the .EXAMPLES.

				v1.1: 17 January 2015
					Realised v1.0 wouldn't work correctly for EE pools, and that the "meetFqdn" isn't actually required.
					Changed the ScheduledTask Arguments to an exection policy of "AllSigned". (Was unrestricted, leftover from design)
					Changed write-host to write-output & added "-NoProfile" to Task Arguments (thanks Pat)

				v1.0: 16 January 2015
					Initial release.


.LINK
    https://greiginsydney.com/New-LyncMeetingWarmup

.EXAMPLE
	.\New-LyncMeetingWarmup.ps1

	Description
	-----------
    With no input parameters passed to it, the script will dump this help text to screen.


.EXAMPLE
	.\New-LyncMeetingWarmup.ps1 -CreateTasks

	Description
	-----------
	If the "CreateTasks" flag is set, two new Scheduled tasks will be created.
	(This and "-GetScheduledTaskInfo" are the only two ways you as a user will normally run this script).

.EXAMPLE
	.\New-LyncMeetingWarmup.ps1 -internal

	Description
	-----------
	If the "Internal" flag is set, a web call will be made to the host's FQDN on port 443, to
	"warm up" a sleeping Application in IIS, resulting in a faster join experience for users.
	(This mode will normally be run by the Scheduled Task created after using the "-CreateTasks" step).

.EXAMPLE
	.\New-LyncMeetingWarmup.ps1 -external

	Description
	-----------
	If the "External" flag is set, a web call will be made to the host's FQDN on port 4443, to
	"warm up" a sleeping Application in IIS, resulting in a faster join experience for users.
	(This mode will normally be run by the Scheduled Task created after using the "-CreateTasks" step).

.EXAMPLE
	.\New-LyncMeetingWarmup.ps1 -GetScheduledTaskInfo

	Description
	-----------
	This runs a "GetScheduledTaskInfo" query to display the last time the scheduled tasks ran.
	(This and "-CreateTasks" are the only two ways you as a user will normally run this script).

.PARAMETER CreateTasks
		Boolean. If $True (or simply present) and the MeetFQDN is also provided, the required Scheduled Tasks will be created.
		Their presence will be tested first and additional identical tasks won't be created if this script is called accidentally/repeatedly.

.PARAMETER Internal
		Boolean. If $True (or simply present) and the MeetFQDN is also provided, the Internal meeting URL will be called.

.PARAMETER External
		Boolean. If $True (or simply present) and the MeetFQDN is also provided, the External meeting URL will be called.

.PARAMETER GetScheduledTaskInfo
		Boolean. If $True (or simply present), the script will generate a query of the scheduled tasks.

.PARAMETER SkipUpdateCheck
	Boolean. Skips the automatic check for an Update. Courtesy of Pat: http://www.ucunleashed.com/3168

#>

[CmdletBinding(SupportsShouldProcess = $False, DefaultParameterSetName='None')]
Param(
	[Parameter(ParameterSetName='Create', Mandatory = $true)]
	[alias("create")][switch]$CreateTasks,
	[Parameter(ParameterSetName='Int', Mandatory = $true)]
    [alias("int")][switch]$Internal,
	[Parameter(ParameterSetName='Ext', Mandatory = $true)]
	[alias("ext")][switch]$External,
	[Parameter(ParameterSetName='Query', Mandatory = $true)]
	[alias("info")][switch]$GetScheduledTaskInfo
)


#--------------------------------
# START FUNCTIONS ---------------
#--------------------------------

function Get-UpdateInfo
{
  <#
      .SYNOPSIS
      Queries an online XML source for version information to determine if a new version of the script is available.
	  *** This version customised by Greig Sheridan. @greiginsydney https://greiginsydney.com ***

      .DESCRIPTION
      Queries an online XML source for version information to determine if a new version of the script is available.

      .NOTES
      Version               : 1.2 - See changelog at https://ucunleashed.com/3168 for fixes & changes introduced with each version
      Wish list             : Better error trapping
      Rights Required       : N/A
      Sched Task Required   : No
      Lync/Skype4B Version  : N/A
      Author/Copyright      : © Pat Richard, Office Servers and Services (Skype for Business) MVP - All Rights Reserved
      Email/Blog/Twitter    : pat@innervation.com  https://ucunleashed.com  @patrichard
      Donations             : https://www.paypal.me/PatRichard
      Dedicated Post        : https://ucunleashed.com/3168
      Disclaimer            : You running this script/function means you will not blame the author(s) if this breaks your stuff. This script/function
                            is provided AS IS without warranty of any kind. Author(s) disclaim all implied warranties including, without limitation,
                            any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use
                            or performance of the sample scripts and documentation remains with you. In no event shall author(s) be held liable for
                            any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss
                            of business information, or other pecuniary loss) arising out of the use of or inability to use the script or
                            documentation. Neither this script/function, nor any part of it other than those parts that are explicitly copied from
                            others, may be republished without author(s) express written permission. Author(s) retain the right to alter this
                            disclaimer at any time. For the most up to date version of the disclaimer, see https://ucunleashed.com/code-disclaimer.
      Acknowledgements      : Reading XML files
                            http://stackoverflow.com/questions/18509358/how-to-read-xml-in-powershell
                            http://stackoverflow.com/questions/20433932/determine-xml-node-exists
      Assumptions           : ExecutionPolicy of AllSigned (recommended), RemoteSigned, or Unrestricted (not recommended)
      Limitations           :
      Known issues          :

      .EXAMPLE
      Get-UpdateInfo -Title "Compare-PkiCertificates.ps1"

      Description
      -----------
      Runs function to check for updates to script called <Varies>.

      .INPUTS
      None. You cannot pipe objects to this script.
  #>
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
	[string] $title
	)
	try
	{
		[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
		if ($HasInternetAccess)
		{
			write-verbose "Performing update check"
			# ------------------ TLS 1.2 fixup from https://github.com/chocolatey/choco/wiki/Installation#installing-with-restricted-tls
			$securityProtocolSettingsOriginal = [System.Net.ServicePointManager]::SecurityProtocol
			try {
			  # Set TLS 1.2 (3072). Use integers because the enumeration values for TLS 1.2 won't exist in .NET 4.0, even though they are
			  # addressable if .NET 4.5+ is installed (.NET 4.5 is an in-place upgrade).
			  [System.Net.ServicePointManager]::SecurityProtocol = 3072
			} catch {
			  Write-verbose 'Unable to set PowerShell to use TLS 1.2 due to old .NET Framework installed.'
			}
			# ------------------ end TLS 1.2 fixup
			[xml] $xml = (New-Object -TypeName System.Net.WebClient).DownloadString('https://greiginsydney.com/wp-content/version.xml')
			[System.Net.ServicePointManager]::SecurityProtocol = $securityProtocolSettingsOriginal #Reinstate original SecurityProtocol settings
			$article  = select-XML -xml $xml -xpath "//article[@title='$($title)']"
			[string] $Ga = $article.node.version.trim()
			if ($article.node.changeLog)
			{
				[string] $changelog = "This version includes: " + $article.node.changeLog.trim() + "`n`n"
			}
			if ($Ga -gt $ScriptVersion)
			{
				$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
				$updatePrompt = $wshell.Popup("Version $($ga) is available.`n`n$($changelog)Would you like to download it?",0,"New version available",68)
				if ($updatePrompt -eq 6)
				{
					Start-Process -FilePath $article.node.downloadUrl
					Write-Warning "Script is exiting. Please run the new version of the script after you've downloaded it."
					exit
				}
				else
				{
					write-verbose "Upgrade to version $($ga) was declined"
				}
			}
			elseif ($Ga -eq $ScriptVersion)
			{
				write-verbose "Script version $($Scriptversion) is the latest released version"
			}
			else
			{
				write-verbose "Script version $($Scriptversion) is newer than the latest released version $($ga)"
			}
		}
		else
		{
		}

	} # end function Get-UpdateInfo
	catch
	{
		write-verbose "Caught error in Get-UpdateInfo"
		if ($Global:Debug)
		{
			$Global:error | fl * -f #This dumps to screen as white for the time being. I haven't been able to get it to dump in red
		}
	}
}


function CreateNewScheduledTask()
{
    param ([string]$IntExt)

	$global:scriptpath

	if ($IntExt -eq "Int")
	{
		$TaskName = "Warmup Lync-SfB Internal App Pool"
		$TaskArg = "-Executionpolicy AllSigned -NoProfile -file ""$($scriptpath)"" -internal"
	}
	else
	{
		$TaskName = "Warmup Lync-SfB External App Pool"
		$TaskArg = "-Executionpolicy AllSigned -NoProfile -file ""$($scriptpath)"" -external"
	}

	$TaskDescr = "Executes a dummy Lync-SfB meeting join attempt to start the App pool running"
	$TaskCommand = "c:\windows\system32\WindowsPowerShell\v1.0\powershell.exe"
	$TaskAction = New-ScheduledTaskAction -Execute "$TaskCommand" -Argument "$TaskArg"
	$TaskTrigger = New-ScheduledTaskTrigger -Once -At 1am 	# This is a dummy trigger, we'll replace it in a sec

	#Now check the task doesn't already exist, and if it's not there, create it:
	if (Get-ScheduledTask -taskname "$TaskName" -ea silentlycontinue)
	{
		write-warning "A task by the name ""$($TaskName)"" already exists"
	}
	else
	{
		#Create the task:
		Register-ScheduledTask -Action $TaskAction -Trigger $Tasktrigger -TaskName "$TaskName" -TaskPath "Microsoft\Windows" -User "System" -RunLevel Highest | out-Null

		#Export the task:
		[xml]$EncryptSyncST = Export-ScheduledTask "$TaskName" -TaskPath "Microsoft\Windows"
		#Edit the XML for the trigger:
		if ($IntExt -eq "Int")
		{
			$UpdatedXML = [xml]'<EventTrigger xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task"><Enabled>true</Enabled><Subscription>&lt;QueryList&gt; &lt;Query Id="0" Path="System"&gt; &lt;Select Path="System"&gt; *[System[Provider[@Name=''Microsoft-Windows-WAS''] and (EventID=5074)]] and *[EventData[Data[@Name=''AppPoolID''] and (Data=''LyncIntFeature'')]] &lt;/Select&gt; &lt;/Query&gt; &lt;/QueryList&gt;</Subscription></EventTrigger>'
		}
		else
		{
			$UpdatedXML = [xml]'<EventTrigger xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task"><Enabled>true</Enabled><Subscription>&lt;QueryList&gt; &lt;Query Id="0" Path="System"&gt; &lt;Select Path="System"&gt; *[System[Provider[@Name=''Microsoft-Windows-WAS''] and (EventID=5074)]] and *[EventData[Data[@Name=''AppPoolID''] and (Data=''LyncExtFeature'')]] &lt;/Select&gt; &lt;/Query&gt; &lt;/QueryList&gt;</Subscription></EventTrigger>'
		}
		$EncryptSyncST.Task.Triggers.InnerXml = $UpdatedXML.InnerXML

		#Discard the original and replace it with our updated copy:
		Get-ScheduledTask -taskname "$TaskName" | Unregister-ScheduledTask -Confirm:$false
		Register-ScheduledTask "$TaskName" -TaskPath "Microsoft\Windows" -Xml $EncryptSyncST.OuterXml | out-Null
		write-verbose "Created task ""$($TaskName)"" OK"
	}
}


function QueryScheduledTask()
{
    param ([string]$Taskname)
	if (Get-Command Get-ScheduledTask -ea silentlycontinue)
	{
		if (Get-ScheduledTask -taskname "$TaskName" -ea silentlycontinue)
		{
			Get-ScheduledTask -taskname "$TaskName" | Get-ScheduledTaskInfo
		}
		else
		{
			write-warning "No task by the name ""$($TaskName)"" exists"
		}
	}
	else
	{
		write-warning "The ""Get-ScheduledTask"" command is not valid for this operating system"
	}
}

#--------------------------------
# END FUNCTIONS -----------------
#--------------------------------

$ScriptVersion = "1.5"
$Error.Clear()
$scriptpath = $MyInvocation.MyCommand.Path
$HostFqdn = "$env:computername.$env:userdnsdomain"

if ($skipupdatecheck)
{
	write-verbose "Skipping update check"
}
else
{
	write-progress -id 1 -Activity "Performing update check" -Status "Running Get-UpdateInfo" -PercentComplete (50)
	Get-UpdateInfo -title "New-LyncMeetingWarmup.ps1"
	write-progress -id 1 -Activity "Back from performing update check" -Status "Running Get-UpdateInfo" -Completed
}

#Thank you Pat Richard: www.ehloworld.com/1697
$OSVersion = Get-WMIObject -Class Win32_OperatingSystem

if ($CreateTasks)
{
	If (!(($OSVersion.Version -match "6.2.920") -or ` # Win 8.0 or Server 2012
		  ($OSVersion.Version -match "6.3.960") -or ` # Win 8.1 or Server 2012 R2
		  ($OSVersion.Version -match "6.4.")	-or ` # Server 'vNext' (Server 2016's working title)
		  ($OSVersion.Version -match "10.")     -or   # Server 2016
		  ($OSVersion.Caption -like "*2019*")))		  # Server 2019 reports a Version of 10.0.17763 - but the Caption IDs it more conclusively
	{
		write-output ""
		write-output "Sorry, this script requires Server 2012 or later to fully automate the Task creation process."
		write-output "Refer http://www.lynclog.com/2013/12/user-might-experince-delay-when-join.html for the manual process"
		write-output "or https://greiginsydney.com/New-LyncMeetingWarmup/#Server2008 for a partial solution."
		write-output ""
		exit
	}

	#Pat again:
	if (! (New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
	{
		write-output ""
		write-output "Sorry, the script needs to be run Elevated to create new tasks"
		write-output ""
	}
	else
	{
		CreateNewScheduledTask "Int"
		CreateNewScheduledTask "Ext"
	}
	exit
}
if ($Internal)
{
	Invoke-WebRequest ("https://" + $HostFqdn + "/Meet?key=A1B2C3D4") -verbose:$false | Out-Null
	write-verbose "Actioned Internal web request to ""https://$($HostFqdn)/Meet?key=A1B2C3D4"""
	exit
}
if ($External)
{
	Invoke-WebRequest ("https://" + $HostFqdn + ":4443/Meet?key=A1B2C3D4") -verbose:$false | Out-Null
	write-verbose "Actioned External web request to ""https://$($HostFqdn):4443/Meet?key=A1B2C3D4"""
	exit
}
if ($GetScheduledTaskInfo)
{
	$TaskName = "Warmup Lync-SfB Internal App Pool"
	QueryScheduledTask $TaskName
	$TaskName = "Warmup Lync-SfB External App Pool"
	QueryScheduledTask $TaskName
	exit
}

#Nothing to do!
get-help .\New-LyncMeetingWarmup.ps1


# Credits:
#--------------------
# It is based upon (and 100% inspired by) this EXCELLENT post and suggestion by Drago Totev:
# http://www.lynclog.com/2013/12/user-might-experince-delay-when-join.html
#
# Creating a task in PowerShell:
# http://www.verboon.info/2013/12/powershell-creating-scheduled-tasks-with-powershell-version-3/
#
# Tricky Task creation:
# http://stackoverflow.com/questions/20108886/scheduled-task-with-daily-trigger-and-repetition-interval
# https://p0w3rsh3ll.wordpress.com/2013/07/05/deprecated-features-of-the-task-scheduler/

# Code signing certificate with thanks to DigiCert:
# SIG # Begin signature block
# MIIceAYJKoZIhvcNAQcCoIIcaTCCHGUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUqOcVAjkiPzbtRFbBRi/h4nuq
# fGGgghenMIIFMDCCBBigAwIBAgIQA1GDBusaADXxu0naTkLwYTANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTIwMDQxNzAwMDAwMFoXDTIxMDcw
# MTEyMDAwMFowbTELMAkGA1UEBhMCQVUxGDAWBgNVBAgTD05ldyBTb3V0aCBXYWxl
# czESMBAGA1UEBxMJUGV0ZXJzaGFtMRcwFQYDVQQKEw5HcmVpZyBTaGVyaWRhbjEX
# MBUGA1UEAxMOR3JlaWcgU2hlcmlkYW4wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQC0PMhHbI+fkQcYFNzZHgVAuyE3BErOYAVBsCjZgWFMhqvhEq08El/W
# PNdtlcOaTPMdyEibyJY8ZZTOepPVjtHGFPI08z5F6BkAmyJ7eFpR9EyCd6JRJZ9R
# ibq3e2mfqnv2wB0rOmRjnIX6XW6dMdfs/iFaSK4pJAqejme5Lcboea4ZJDCoWOK7
# bUWkoqlY+CazC/Cb48ZguPzacF5qHoDjmpeVS4/mRB4frPj56OvKns4Nf7gOZpQS
# 956BgagHr92iy3GkExAdr9ys5cDsTA49GwSabwpwDcgobJ+cYeBc1tGElWHVOx0F
# 24wBBfcDG8KL78bpqOzXhlsyDkOXKM21AgMBAAGjggHFMIIBwTAfBgNVHSMEGDAW
# gBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQUzBwyYxT+LFH+GuVtHo2S
# mSHS/N0wDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1Ud
# HwRwMG4wNaAzoDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3Vy
# ZWQtY3MtZzEuY3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hh
# Mi1hc3N1cmVkLWNzLWcxLmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgG
# CCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEE
# ATCBhAYIKwYBBQUHAQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wTgYIKwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFNIQTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMB
# Af8EAjAAMA0GCSqGSIb3DQEBCwUAA4IBAQCtV/Nu/2vgu+rHGFI6gssYWfYLEwXO
# eJqOYcYYjb7dk5sRTninaUpKt4WPuFo9OroNOrw6bhvPKdzYArXLCGbnvi40LaJI
# AOr9+V/+rmVrHXcYxQiWLwKI5NKnzxB2sJzM0vpSzlj1+fa5kCnpKY6qeuv7QUCZ
# 1+tHunxKW2oF+mBD1MV2S4+Qgl4pT9q2ygh9DO5TPxC91lbuT5p1/flI/3dHBJd+
# KZ9vYGdsJO5vS4MscsCYTrRXvgvj0wl+Nwumowu4O0ROqLRdxCZ+1X6a5zNdrk4w
# Dbdznv3E3s3My8Axuaea4WHulgAvPosFrB44e/VHDraIcNCx/GBKNYs8MIIFMDCC
# BBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0Ew
# HhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEV
# MBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29t
# MTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5n
# IENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfT
# CzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdgl
# rA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRn
# iolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7
# MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPr
# CGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z
# 3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8E
# BAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0g
# BEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nED
# wGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqG
# SIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9
# D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQG
# ivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEeh
# emhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJ
# RZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5
# gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIGajCCBVKgAwIBAgIQAwGa
# Ajr/WLFr1tXq5hfwZjANBgkqhkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEw
# HwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAw
# WhcNMjQxMDIyMDAwMDAwWjBHMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNl
# cnQxJTAjBgNVBAMTHERpZ2lDZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBT
# qZ8fZFnmfGt/a4ydVfiS457VWmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWR
# n8YUOawk6qhLLJGJzF4o9GS2ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRV
# fRiGBYxVh3lIRvfKDo2n3k5f4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3v
# J+P3mvBMMWSN4+v6GYeofs/sjAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA
# 8bLOcEaD6dpAoVk62RUJV5lWMJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGj
# ggM1MIIDMTAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8E
# DDAKBggrBgEFBQcDCDCCAb8GA1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIB
# kjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQG
# CCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMA
# IABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMA
# IABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMA
# ZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkA
# bgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgA
# IABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUA
# IABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAA
# cgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQAS
# KxOYspkH7R7for5XDStnAs0wHQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9
# MH0GA1UdHwR2MHQwOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRENBLTEuY3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkw
# JAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcw
# AoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElE
# Q0EtMS5jcnQwDQYJKoZIhvcNAQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI
# //+x1GosMe06FxlxF82pG7xaFjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7ea
# sGAm6mlXIV00Lx9xsIOUGQVrNZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8Oxw
# YtNiS7Dgc6aSwNOOMdgv420XEwbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQN
# JsQOfxu19aDxxncGKBXp2JPlVRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNt
# omHpigtt7BIYvfdVVEADkitrwlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbN
# MIIFtaADAgECAhAG/fkDlgOt6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJ
# BgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5k
# aWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBD
# QTAeFw0wNjExMTAwMDAwMDBaFw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5j
# b20xITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/J
# M/xNRZFcgZ/tLJz4FlnfnrUkFcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPs
# i3o2CAOrDDT+GEmC/sfHMUiAfB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ
# 8DIhFonGcIj5BZd9o8dD3QLoOz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNu
# gnM/JksUkK5ZZgrEjb7SzgaurYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJr
# GGWxwXOt1/HYzx4KdFxCuGh+t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3ow
# ggN2MA4GA1UdDwEB/wQEAwIBhjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUH
# AwIGCCsGAQUFBwMDBggrBgEFBQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIB
# xTCCAbQGCmCGSAGG/WwAAQQwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIw
# ggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQA
# aQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUA
# cAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMA
# UAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEA
# cgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkA
# dAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8A
# cgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIA
# ZQBuAGMAZQAuMAsGCWCGSAGG/WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsG
# AQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqg
# OKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURS
# b290Q0EuY3JsMB0GA1UdDgQWBBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSME
# GDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+
# ybcoJKc4HbZbKa9Sz1LpMUerVlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6
# hnKtOHisdV0XFzRyR4WUVtHruzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5P
# sQXSDj0aqRRbpoYxYqioM+SbOafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke
# /MV5vEwSV/5f4R68Al2o/vsHOE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qqu
# AHzunEIOz5HXJ7cW7g/DvXwKoO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQ
# nHcUwZ1PL1qVCCkQJjGCBDswggQ3AgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EANRgwbrGgA18btJ2k5C8GEwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFGwkffSUyK2rtP/pu3mc
# vZr1q4JuMA0GCSqGSIb3DQEBAQUABIIBAFon6VnhNtP0yzyHhMEBQuhodldrPZkX
# Oaev2KBIfD5QnzX6Xy1wrzOd9hqkUGLlEMKv2LH3Mwrfl97hAxXAj1rbMyKy2Vk3
# mBsBl5ZviVUj7bHCcZlBAEAfi283rfmrexBfuPiHNVGAC7Km5LcIfd+qcYr+ou9q
# 98e2zRG1Z0TrwkFrELdeyGACBej7wLULDVBNc1292ulcBaQsrFx7KcIllUqnTsiB
# Mcoc3H6sSXCqMspTTMD42KHZftEq2bGPeGlD4fePQPTSTu5/TcnKlnnn9olyi8rL
# d30tAw11Hwc96WU/03FtbzyaoZWr4Mzsd1fGwQpReDRsYj+w57lBhx2hggIPMIIC
# CwYJKoZIhvcNAQkGMYIB/DCCAfgCAQEwdjBiMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYD
# VQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTECEAMBmgI6/1ixa9bV6uYX8GYw
# CQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcN
# AQkFMQ8XDTIwMDUwNTExMzEyNlowIwYJKoZIhvcNAQkEMRYEFPkRyCM/gZ/qUCIB
# K8qCP8UvPOcGMA0GCSqGSIb3DQEBAQUABIIBAEitMOFhmXBGRtFTfCL/LTiVQ+Y/
# YvjCHeWiCkRqe1wKmMPiMaao/CG/lPlZmW9mHmlDTcVF893hS1J5hLl8Ap4tyr7C
# EKNuFV2K+yjex8PoR/0PRnqS66xOFjGT5jS11ZRU/oQQ+XmrrTpGy+DSYqOgK1Yu
# ufuScwPHYfLc/J4W/UXPDd0JGXL+5aFDzrdLeii3STLBjOk4WNGS+YQl1Qp5Jg5d
# 3dNLF+biHo8E/iprI81awA8whYylfaF8xnoNHGN/VSIeqzRlW8VZYEWCM0qR6vsF
# SEty29KeaPWcGCaxhTkqOwzGPZ0oqrVnVTdlxayYjYS07OhgFIVqqED4X6Y=
# SIG # End signature block
