################################################################################################
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
# SCRIPT:  Exchange Online- Get OWA Text Notifications In Use
#  -Dependent on EXOv2 module: (https://www.powershellgallery.com/packages/ExchangeOnlineManagement)
#
#  EXO_OWATextNotificationsInUse_1.1.ps1
#  - Download latest version from https://github.com/garrint 
#  
#  Created by: Garrin Thompson 9/15/2020 garrint@microsoft.com *** 
#
################################################################################################
# This script checks all mailbox recipients provided in an import file with a PrimarySmtpAddress 
#   attribute available/populated for the value of $true in the Get-TextMessagingAccount
#   for the attribute: NotificationPhoneNumberVerified.  If True, we gather relevant information.
#
# The output can be used to target the users with OWA Text Notification enabled for communication 
#   about the deprecation of this feature in coming weeks.
################################################################################################

#----------------------v
#ScriptFunctionsSection
#----------------------v

# Writes output to a log file with a time date stamp
    Function Write-Log {
        Param ([string]$string)
        $NonInteractive = 1
        # Get the current date
        [string]$date = Get-Date -Format G
        # Write everything to our log file
        ( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append
        # If NonInteractive true then supress host output
        If (!($NonInteractive)){
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

# Setup a new O365 Powershell Session using RobustCloudCommand concepts to help maintain the session
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
        $null = Connect-ExchangeOnline -UserPrincipalName $EXOLogonUPN -ShowProgress:$false -ShowBanner:$false
        # Set the Start time for the current session
        Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
    }

# Verifies that the connection is healthy; Resets it every "$ResetSeconds" number of seconds (14.5 mins) either way 
    Function Test-O365Session {
        # Get the time that we are working on this object to use later in testing
        $ObjectTime = Get-Date
        # Reset and regather our session information
        $SessionInfo = $null
        $SessionInfo = Get-PSSession
        # Make sure we found a session
        If ($SessionInfo -eq $null) { 
            Write-log ("[ERROR] - No Session Found")
            Write-log ("Recreating Session")
            New-CleanO365Session
        }	
        # Make sure it is in an opened state If not log and recreate
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
            If ($DelayinSeconds -gt 0){
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
        If ($ManualThrottle -gt 0){
            Write-log ("Sleeping " + $ManualThrottle + " milliseconds")
            Start-Sleep -Milliseconds $ManualThrottle
        }
    }

#------------------v
#ScriptSetupSection
#------------------v

#Set Variables
	$logfilename = 'EXO_ScriptExecution_logfile_'
	$outputfilename = 'EXO_Script_Output_'
	$execpol = get-executionpolicy
	Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force  #this is just for the session running this script
	Write-Host; $EXOLogonUPN=Read-host "Type in UPN for account that will execute this script"; write-host "...pleasewait...connecting to EXO..."
	$SmtpCreds = (get-credential -Message "Provide EXO account Pasword" -UserName "$EXOLogonUPN")
	# Set $OutputFolder to Current PowerShell Directory
	[IO.Directory]::SetCurrentDirectory((Convert-Path (Get-Location -PSProvider FileSystem)))
    $outputFolder = [IO.Directory]::GetCurrentDirectory()
    $StarttimeTicks = (Get-Date).Ticks
	$logFile = $outputFolder + '\' + $logfilename + $StarttimeTicks + ".txt"
	$OutputFile= $outputfolder + '\' + $outputfilename + $StarttimeTicks + ".csv"
	[int]$ManualThrottle=0
	[double]$ActiveThrottle=.25
	[int]$ResetSeconds=870
# Setup our first session to O365
	$ErrorCount = 0
	New-CleanO365Session
	Write-Log ("Connected to Exchange Online")
	write-host;write-host -ForegroundColor Green "...Connected to Exchange Online as $EXOLogonUPN";write-host
# Get when we started the script for estimating time to completion
	$ScriptStartTime = Get-Date
	$startDate = Get-Date
	write-progress -id 1 -activity "Beginning..." -PercentComplete (5) -Status "initializing variables"
# Clear the error log so that sending errors to file relate only to this run of the script
	$error.clear()

#-------------------------v
#Start CUSTOM CODE Section
#-------------------------v

# Import MailboxList CSV
	Write-Progress -Id 1 -Activity "Importing all EXO UserMailboxes" -PercentComplete (15) -Status "Import-Csv"
	$exombxlist = Import-Csv c:\temp\allexombxs.csv

# Set a counter and some variables to use for periodic write/flush and reporting for loop to create Hashtable
	$currentProgress = 1
	[TimeSpan]$caseCheckTotalTime=0
	# report counter
		$c = 0
	# running counter
		$i = 0
	# Set the number of objects to cycle before writing to disk and sending stats, i'd consider 5000 max
		$statLimit = 1000
	# Get the total number of objects, which we use in some stat calculations
		$t = $exombxlist.count
	# Set some timedate variables for the stats report
		$loopStartTime = Get-Date
		$loopCurrentTime = Get-Date

# Create Array
	# Prepare a new array for the Objects you want to get from EXO and set the attributes list
	[System.Collections.ArrayList]$OWAtexters = New-Object System.Collections.ArrayList($null)
	$OWAtexters | Select PrimarySmtpAddress,Identity,NotificationPhoneNumber,NotificationPhoneNumberVerified
	$OWAtexters.Clear()

#  Create a new array for use in a Hashtable containing calculated properties indexed by a property from the both lists
    [System.Collections.ArrayList]$ResultsList = New-Object System.Collections.ArrayList($null)
    $ResultsList | Select PrimarySmtpAddress,Identity,NotificationPhoneNumber,NotificationPhoneNumberVerified
	$ResultsList.Clear()

# Get Data from EXO and populate arrays 
	# Update Log
        $progressActions = $exombxlist.count
        Write-Log ("Starting data collection");Write-Log ("Exchange Online mailboxes being checked: " + ($progressActions));Write-Log ("Elapsed time to collect data from Get-TextMessagingAccount: " + ($($invokeElapsedTime)))
	# Update Screen
        Write-host;Write-host -foregroundcolor Cyan "Starting data collection...(counters will display for every 1000 objects retrieved)";;sleep 2;write-host "-------------------------------------------------"
        Write-Host -NoNewline "Total EXO UserMailboxes being checked: ";Write-Host -ForegroundColor Green $progressActions
    Foreach ($mbx in $exombxlist) {
		#Check that we still have a valid EXO PS session
            Test-O365Session
		# Total up the running count 
			$i++
		# Dump the $ResultsList to CSV at every $statLimit number of objects (defined above); also send status e-mail with some metrics at each dump.
            If (++$c -eq $statLimit) {
                # Moved this from the bottom of the script, and added -Append parameter
                    $ResultsList | select PrimarySmtpAddress,Identity,NotificationPhoneNumber,NotificationPhoneNumberVerified | export-csv -path $OutputFile -notypeinformation -Append
                    $loopLastTime = $loopCurrentTime
                    $loopCurrentTime = Get-Date
                    $currentRate = $statLimit/($loopCurrentTime-$loopLastTime).TotalHours
                    $avgRate = $i/($loopCurrentTime-$loopStartTime).TotalHours
                # Send a status email each time we write $statimit number of objects to the file (requires $SmtpCreds to be defined)
                    $old_ErrorActionPreference = $ErrorActionPreference
                    $ErrorActionPreference = 'SilentlyContinue'
                    Send-MailMessage -From "$EXOLogonUPN" -To "$EXOLogonUPN" -Subject "$OutputFilename Progress" -Body "$OutputFilename PROGRESS report`n`nCurrentTime: $loopCurrentTime`nStartTime: $loopStartTime`n`nCounter: $i out of $t devices, at a current rate of $currentRate per hour.`n`nBased on the overall average rate, we will be done in $($(1/($avgRate*24)*($t-$i)) - $((Get-Date).TotalDays)) days on $((Get-Date).AddDays($(1/($avgRate*24)*($t-$i))))." -SmtpServer 'smtp.office365.com' -Port:25 -UseSsl:$true -BodyAsHtml:$false -Credential:$SmtpCreds
                    $ErrorActionPreference = $old_ErrorActionPreference
                # Update Log
                    Write-Log ("Counter: $i out of $t objects at $currentRate per hour. Estimated Completion on $((Get-Date).AddDays($(1/($avgRate*24)*($t-$i)))))")
                # Update Screen
                    Write-host "Counter: $i out of $t objects at $currentRate per hour. Estimated Completion on $((Get-Date).AddDays($(1/($avgRate*24)*($t-$i)))))" 
                # Clear StatLimit and $ResultsLost for next run
                    $c = 0
                    $ResultsList.Clear()
                    write-progress -id 1 -Activity "Checking for NotificationPhoneNumberVerified" -PercentComplete ((1/($avgRate*24)*($t-$i))) -Status "Get-TextMessagingAccount is running"
            }
        # Test for NotificationPhoneNumberVerified = $true, Get-TextMessagingAccount
            $TexterList = $null
            If ((Get-TextMessagingAccount $mbx.PrimarySmtpAddress).NotificationPhoneNumberVerified) {
                $TexterList = Get-TextMessagingAccount $mbx.PrimarySmtpAddress | Select-Object @{n='PrimarySmtpAddress';e={$mbx.PrimarySmtpAddress}},Identity,NotificationPhoneNumber,NotificationPhoneNumberVerified
                # Update the pivotal index Hashtable (using shorthand notation for add-member)
                    $line = @{
                        PrimarySmtpAddress=$TexterList.PrimarySmtpAddress
                        Identity=$TexterList.Identity
                        NotificationPhoneNumber=$TexterList.NotificationPhoneNumber
                        NotificationPhoneNumberVerified=$TexterList.NotificationPhoneNumberVerified
                    }
                # Since arraylist.add returns highest index, we need a way to ignore that value with Out-Null
                    $ResultsList.Add((New-Object PSobject -property $line)) | Out-Null
                }
        # Update Progress
            $currentProgress++
	}
# Update Elapsed Time
    $invokeEndDate = Get-Date
    $invokeElapsedTime = $invokeEndDate - $startDate
    Write-Host -NoNewline "Elapsed time to collect data for OWA Text Notifications in use:  ";write-host -ForegroundColor Yellow "$($invokeElapsedTime)"

# Disconnect from EXO and cleanup the PS session
	Get-PSSession | Remove-PSSession -Confirm:$false -ErrorAction silentlycontinue

# Create the Output File (report) using the attributes created in the Hashtable by exporting to CSV
	# Update Progress
		write-progress -id 1 -activity "Creating Output Report" -PercentComplete (95) -Status "$outputFolder"
	# Create Report
		$ResultsList | select PrimarySmtpAddress,Identity,NotificationPhoneNumber,NotificationPhoneNumberVerified | export-csv -path $OutputFile -notypeinformation -Append

# Separately capture any PowerShell errors and output to an errorfile
	$errfilename = $outputfolder + '\' + $logfilename + "_ERRORs_" + (Get-Date).Ticks + ".txt" 
	write-progress -id 1 -activity "Error logging" -PercentComplete (99) -Status "$errfilename"
	ForEach ($err in $error) {  
		$logdata = $null 
		$logdata = $err 
		If ($logdata) 
			{ 
				out-file -filepath $errfilename -Inputobject $logData -Append 
			} 
	}
#Clean Up and Show Completion in session and logs
	# Update Progress	
		write-progress -id 1 -activity "Complete" -PercentComplete (100) -Status "Success!"
	# Update Log
		$endDate = Get-Date
		$elapsedTime = $endDate - $startDate
		Write-Log ("Report started at: " + $($startDate));Write-Log ("Report ended at: " + $($endDate));Write-Log ("Total Elapsed Time: " + $($elapsedTime)); Write-Log ("Data Collection Completed!")
	# Update Screen
		write-host;Write-Host -NoNewLine "Report started at    ";write-host -ForegroundColor Yellow "$($startDate)"
		Write-Host -NoNewLine "Report ended at      ";write-host -ForegroundColor Yellow "$($endDate)"
		Write-Host -NoNewLine "Total Elapsed Time:   ";write-host -ForegroundColor Yellow "$($elapsedTime)"
		Write-host "-------------------------------------------------";write-host -foregroundcolor Cyan "Data collection Complete!";write-host;write-host -foregroundcolor Green "...The $OutputFile CSV and execution logs were created in $outputFolder";write-host;write-host;sleep 1

#------------------------^
#End CUSTOM CODE Section
#------------------------^
#End of Script