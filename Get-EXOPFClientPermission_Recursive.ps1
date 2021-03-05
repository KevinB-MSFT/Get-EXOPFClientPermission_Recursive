#########################################################################################
# LEGAL DISCLAIMER
# This Sample Code is provided for the purpose of illustration only and is not
# intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
# RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
# EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
# MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
# nonexclusive, royalty-free right to use and modify the Sample Code and to
# reproduce and distribute the object code form of the Sample Code, provided
# that You agree: (i) to not use Our name, logo, or trademarks to market Your
# software product in which the Sample Code is embedded; (ii) to include a valid
# copyright notice on Your software product in which the Sample Code is embedded;
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
# against any claims or lawsuits, including attorneys’ fees, that arise or result
# from the use or distribution of the Sample Code.
# 
# This posting is provided "AS IS" with no warranties, and confers no rights. Use
# of included script samples are subject to the terms specified at 
# https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx.
#
# Exchange Online Device partnership inventory
# Get-EXOPFClientPermission_Recursive
#  
# Created by: Kevin Bloom and Garrin Thompson 03/05/2021 Kevin.Bloom@Microsoft.com  *** "Borrowed" a few quality-of-life functions from Start-RobustCloudCommand.ps1 and added EXOv2 connection
#
#########################################################################################
#
#########################################################################################

##Define variables and constants
#Array to collect and gather all of the results
$Global:Records = @()
$Global:ParentGroup = ""
$Global:NestedGroup = ""
# Writes output to a log file with a time date stamp
Function Write-Log {
	Param ([string]$string)
	$NonInteractive = 1
	# Get the current date
	[string]$date = Get-Date -Format G
	# Write everything to our log file
	( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
	# If NonInteractive true then supress host output
	if (!($NonInteractive)){
		( "[" + $date + "] - " + $string) | Write-Host
	}
}

# Sleeps X seconds and displays a progress bar
Function Start-SleepWithProgress {
	Param([int]$sleeptime)
	# Loop Number of seconds you want to sleep
	For ($i=0;$i -le $sleeptime;$i++){
		$timeleft = ($sleeptime - $i);
		# Progress bar showing progress of the sleep
		Write-Progress -Activity "Sleeping" -CurrentOperation "$Timeleft More Seconds" -PercentComplete (($i/$sleeptime)*100);
		# Sleep 1 second
		start-sleep 1
	}
	Write-Progress -Completed -Activity "Sleeping"
}

# Setup a new O365 Powershell Session using RobustCloudCommand concepts
Function New-CleanO365Session {
	#Prompt for UPN used to login to EXO 
   Write-log ("Removing all PS Sessions")

   # Destroy any outstanding PS Session
   Get-PSSession | Remove-PSSession -Confirm:$false
   
   # Force Garbage collection just to try and keep things more agressively cleaned up due to some issue with large memory footprints
   [System.GC]::Collect()
   
   # Sleep 10s to allow the sessions to tear down fully
   Write-Log ("Sleeping 10 seconds to clear existing PS sessions")
   Start-Sleep -Seconds 10

   # Clear out all errors
   $Error.Clear()
   
   # Create the session
   Write-Log ("Creating new PS Session")
	#OLD BasicAuth method create session
	#$Exchangesession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Credential -Authentication Basic -AllowRedirection
   # Check for an error while creating the session
	If ($Error.Count -gt 0){
		Write-log ("[ERROR] - Error while setting up session")
		Write-log ($Error)
		# Increment our error count so we abort after so many attempts to set up the session
		$ErrorCount++
		# If we have failed to setup the session > 3 times then we need to abort because we are in a failure state
		If ($ErrorCount -gt 3){
			Write-log ("[ERROR] - Failed to setup session after multiple tries")
			Write-log ("[ERROR] - Aborting Script")
			exit		
		}	
		# If we are not aborting then sleep 60s in the hope that the issue is transient
		Write-log ("Sleeping 60s then trying again...standby")
		Start-SleepWithProgress -sleeptime 60
		
		# Attempt to set up the sesion again
		New-CleanO365Session
	}
   
   # If the session setup worked then we need to set $errorcount to 0
   else {
	   $ErrorCount = 0
   }
   # Import the PS session/connect to EXO
	$null = Connect-ExchangeOnline -UserPrincipalName $EXOLogonUPN -DelegatedOrganization $EXOtenant -ShowProgress:$false -ShowBanner:$false
   # Set the Start time for the current session
	Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
}

# Verifies that the connection is healthy; Goes ahead and resets it every "$ResetSeconds" number of seconds (14.5 mins) either way 
Function Test-O365Session {
	# Get the time that we are working on this object to use later in testing
	$ObjectTime = Get-Date
	# Reset and regather our session information
	$SessionInfo = $null
	$SessionInfo = Get-PSSession
	# Make sure we found a session
	if ($SessionInfo -eq $null) { 
		Write-log ("[ERROR] - No Session Found")
		Write-log ("Recreating Session")
		New-CleanO365Session
	}	
	# Make sure it is in an opened state if not log and recreate
	elseif ($SessionInfo.State -ne "Opened"){
		Write-log ("[ERROR] - Session not in Open State")
		Write-log ($SessionInfo | fl | Out-String )
		Write-log ("Recreating Session")
		New-CleanO365Session
	}
	# If we have looped thru objects for an amount of time gt our reset seconds then tear the session down and recreate it
	elseif (($ObjectTime - $SessionStartTime).totalseconds -gt $ResetSeconds){
		Write-Log ("Session Has been active for greater than " + $ResetSeconds + " seconds" )
		Write-log ("Rebuilding Connection")
		
		# Estimate the throttle delay needed since the last session rebuild
		# Amount of time the session was allowed to run * our activethrottle value
		# Divide by 2 to account for network time, script delays, and a fudge factor
		# Subtract 15s from the results for the amount of time that we spend setting up the session anyway
		[int]$DelayinSeconds = ((($ResetSeconds * $ActiveThrottle) / 2) - 15)
		
		# If the delay is >15s then sleep that amount for throttle to recover
		if ($DelayinSeconds -gt 0){
			Write-Log ("Sleeping " + $DelayinSeconds + " addtional seconds to allow throttle recovery")
			Start-SleepWithProgress -SleepTime $DelayinSeconds
		}
		# If the delay is <15s then the sleep already built into New-CleanO365Session should take care of it
		else {
			Write-Log ("Active Delay calculated to be " + ($DelayinSeconds + 15) + " seconds no addtional delay needed")
		}
		# new O365 session and reset our object processed count
		New-CleanO365Session
	}
	else {
		# If session is active and it hasn't been open too long then do nothing and keep going
	}
	# If we have a manual throttle value then sleep for that many milliseconds
	if ($ManualThrottle -gt 0){
		Write-log ("Sleeping " + $ManualThrottle + " milliseconds")
		Start-Sleep -Milliseconds $ManualThrottle
	}
}

Function Enumerate-PFAccess
{
    Param ($Item)
    #Get's the PF's access control list
    Write-log ("$Item.Identity -Enumerate PFAccess function" )
    Test-O365Session
    $PfPerms = Get-PublicFolderClientPermission -Identity $($item.Identity)

    Foreach ($PfPerm in $PfPerms)
    {
        #If the ACL entry is default, simply add it to the global array
        If ($PfPerm.User.DisplayName -eq "Default")
        {
            $Record = "" | select Identity,User,AccessRights,ParentGroup,NestedGroup
            $Record.Identity = $PfPerm.Identity
            $Record.User = $PfPerm.User
            $Record.AccessRights = $PfPerm.AccessRights
            $Record.ParentGroup = "*Directly Assigned"
            $Record.NestedGroup = "*Directly Assigned"
            $global:Records += $Record
        }
        #If the ACL entry is Anonymous, simply add it to the global array
        ElseIf ($PfPerm.User.DisplayName -eq "Anonymous")
        {
            $Record = "" | select Identity,User,AccessRights,ParentGroup,NestedGroup
            $Record.Identity = $PfPerm.Identity
            $Record.User = $PfPerm.User
            $Record.AccessRights = $PfPerm.AccessRights
            $Record.ParentGroup = "*Directly Assigned"
            $Record.NestedGroup = "*Directly Assigned"
            $global:Records += $Record
        }
        #If the ACL entry is not default or anonymous, check the recipient type for user or group
        elseif ($PfPerm.User.DisplayName -ne "Anonymous" -or $PfPerm.User.DisplayName -ne "Default")
        {
            #Get the recipient type, user or group
            Test-O365Session
            $IDRecipientType = (Get-Recipient -Identity $($PfPerm.user.RecipientPrincipal.PrimarySmtpAddress)).RecipientType
            #If the ACL entry is not a group, simply add it to the global array
            If ($IDRecipientType -notlike "*group*")
            {
                $Record = "" | select Identity,User,AccessRights,ParentGroup,NestedGroup
                $Record.Identity = $PfPerm.Identity
                $Record.User = $PfPerm.User
                $Record.AccessRights = $PfPerm.AccessRights
                $Record.ParentGroup = "*Directly Assigned"
                $Record.NestedGroup = "*Directly Assigned"
                $global:Records += $Record
            }
            #If the ACL entry is  a group, call the enumerate group function
            Elseif ($IDRecipientType -like "*group*")
            {
                $Global:ParentGroup = $PfPerm.User.DisplayName
                Enumerate-Group ($PfPerm)
            }
        }
    }
}

#Function to enumerate group memberships including nested groups
Function Enumerate-Group
{
    Param ($Group)
    Write-log ("$Group.Identity -Enumerate Group function" )
    #Gets the members of the group
    if (!$Group.User.RecipientPrincipal.PrimarySmtpAddress)
    {
        Test-O365Session
        $GroupMembers = Get-DistributionGroupMember -ResultSize Unlimited -Identity $($group.User)
    }
    else 
    {
        Test-O365Session
        $GroupMembers = Get-DistributionGroupMember -ResultSize Unlimited -Identity $($group.User.RecipientPrincipal.PrimarySmtpAddress)
    }
       
    #Loops through the members
    Foreach ($GroupMember in $GroupMembers)
    {
        #If entry is not a group, add the values to a hash table and add the record to the $Records array
        If ($GroupMember.RecipientTypeDetails -notlike "*group*")
        {
            $Record = "" | select Identity,User,AccessRights,ParentGroup,NestedGroup
            $Record.Identity = $Group.Identity
            $Record.User = $GroupMember.DisplayName
            $Record.AccessRights = $Group.AccessRights
            $Record.ParentGroup = $Global:ParentGroup
            $Record.NestedGroup = $Global:NestedGroup
            $global:Records += $Record
        }
        #If entry is  a group, send the group to the Enumerate-Group function
        Elseif ($GroupMember.RecipientTypeDetails -like "*group*")
        {
            $SubGroup = "" | select Identity,User,AccessRights,ParentGroup,NestedGroup
            $SubGroup.Identity = $Group.Identity
            $SubGroup.User = $GroupMember.DisplayName
            $SubGroup.AccessRights = $Group.AccessRights
            $SubGroup.ParentGroup = $Global:ParentGroup
            $SubGroup.NestedGroup = ""
            $Global:NestedGroup = $GroupMember.DisplayName
            #Calls itself so the nested group can be enumerated
            Enumerate-Group ($SubGroup)
        }
        
    }
    $Global:NestedGroup = ""
}

#------------------v
#ScriptSetupSection
#------------------v

#Set Variables
$logfilename = '\Add-EXOPFClientPermissions_logfile_'
$outputfilename = '\Add-EXOPFClientPermissions_Output_'
$execpol = get-executionpolicy
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force  #this is just for the session running this script
Write-Host;$EXOLogonUPN=Read-host "Type in UPN for account that will execute this script"
$EXOtenant=Read-host "Type in your tenant domain name (eg <domain>.onmicrosoft.com)";write-host "...pleasewait...connecting to EXO..."
# Set $OutputFolder to Current PowerShell Directory
[IO.Directory]::SetCurrentDirectory((Convert-Path (Get-Location -PSProvider FileSystem)))
$outputFolder = [IO.Directory]::GetCurrentDirectory()
$DateTicks = (Get-Date).Ticks
$logFile = $outputFolder + $logfilename + $DateTicks + ".txt"
$OutputFile= $outputfolder + $outputfilename + $DateTicks + ".csv"
[int]$ManualThrottle=0
[double]$ActiveThrottle=.25
[int]$ResetSeconds=870

# Setup our first session to O365
$ErrorCount = 0
New-CleanO365Session
Write-Log ("Connected to Exchange Online")
write-host;write-host -ForegroundColor Green "...Connected to Exchange Online as $EXOLogonUPN";write-host

#Gets all Public Folders in EXO
Write-Log ("Getting all PFs")
Test-O365Session
$Session = Get-PSSession
$PFs = Invoke-Command -Session $Session -ScriptBlock{Get-PublicFolder \ -Recurse }

#Loops through the PFs and calls the functions to enumerate the nested members
Foreach ($pf in $pfs)
{
    Enumerate-PFAccess ($pf)
}
#Exports the global array to a .csv
$Global:Records | select Identity,user,@{name='AccessRights';Expression={[string]::join(";",($_.accessrights))}},ParentGroup,NestedGroup |Export-Csv $OutputFile -NoTypeInformation