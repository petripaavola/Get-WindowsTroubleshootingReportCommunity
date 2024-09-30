<#
.SYNOPSIS
    Get-WindowsTroubleshootingReportCommunity.ps1 is a comprehensive Windows and Intune troubleshooting tool. It reads Windows Event logs and structured log files, generating a detailed HTML report that supports real-time filtering and scenario-based views.

.DESCRIPTION
    This script processes Windows Event logs (.evtx) and log files, combining them into a unified timeline for troubleshooting.
    It can analyze logs from both live systems and diagnostics packages (.zip) downloaded from Intune. The output is an interactive HTML report that includes real-time search, filtering, and event categorization using custom event rules defined in the 'EventRules.json' file.
    
    The script can filter and display all events or only known events, as defined in the event rules, and supports both short time snapshots and longer-term log analysis (up to 365 days). Community-contributed event rules enhance the tool’s ability to focus on meaningful data, discarding irrelevant events. 

.PARAMETER LogFile
    The path to a specific Windows Event (.evtx) or structured log file for analysis.

.PARAMETER LogFilesFolder
    The path to a folder containing multiple event or log files to be processed.

.PARAMETER LastMinutes
    Specifies the number of minutes back from the current time to retrieve events. Default is 5 minutes.

.PARAMETER LastHours
    Specifies the number of hours back from the current time to retrieve events. Default is 1 hour.

.PARAMETER LastDays
    Specifies the number of days back from the current time to retrieve events. Default is 1 day.

.PARAMETER MinutesBeforeLastBoot
    Specifies the number of minutes before the last boot to retrieve events.

.PARAMETER MinutesAfterLastBoot
    (Mandatory) Specifies the number of minutes after the last boot to retrieve events.

.PARAMETER StartTime
    Specifies the start time (in the format 'yyyy-MM-dd HH:mm:ss') from which to retrieve events.

.PARAMETER EndTime
    Specifies the end time (in the format 'yyyy-MM-dd HH:mm:ss') until which to retrieve events.

.PARAMETER LogViewerUI
    Launches a UI for viewing and filtering logs directly.

.PARAMETER RealtimeLogViewerUI
    Launches a real-time log viewer UI.

.PARAMETER AllEvents
    Displays all events in the report, including those not categorized by known event rules.

.PARAMETER IncludeSelectedKnownRulesCategoriesOnly
    Filters the report to show only events from the specified known rule categories.

.PARAMETER ExcludeSelectedKnownRulesCategories
    Excludes events from the specified known rule categories from the report.

.PARAMETER SortDescending
    Sorts the report with the most recent events first.

.EXAMPLE
    Get-WindowsTroubleshootingReportCommunity.ps1 -LogFile "C:\Logs\System.evtx" -LastHours 12
    This example processes the specified event log file and retrieves all events from the last 12 hours.

.EXAMPLE
    Get-WindowsTroubleshootingReportCommunity.ps1 -LogFilesFolder "C:\Logs\" -StartTime "2024-09-01 08:00:00" -EndTime "2024-09-01 18:00:00"
    This example processes all event logs in the folder and retrieves events from the specified time range.

.NOTES
    Author: Petri Paavola (Microsoft MVP - Windows and Intune)
    Get-WindowsTroubleshootingReportCommunity.ps1 enables efficient troubleshooting through powerful event log and log file analysis, designed for the IT Pro community with a focus on sharing event rules.

.LINK
	https://github.com/petripaavola/Get-WindowsTroubleshootingReportCommunity
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
Param(
	[Parameter(Mandatory=$false,
				HelpMessage = 'Enter Windows Events (.evtx) or .log file path',
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
				[Alias("FullName")]
				[String]$LogFile = $null,
	[Parameter(Mandatory=$false,
				HelpMessage = 'Enter Windows Event (.evtx) and/or .log files folder path',
                ValueFromPipeline=$false,
                ValueFromPipelineByPropertyName=$false)]
				[String]$LogFilesFolder = $null,
	[Parameter(Mandatory=$false,
				ParameterSetName = 'LastMinutes',
				HelpMessage = 'How many hours back to get Windows Events (default=5)')]
				[int]$LastMinutes = 5,
	[Parameter(Mandatory=$false,
				ParameterSetName = 'LastHours',
				HelpMessage = 'How many hours back to get Windows Events (default=1)')]
				[int]$LastHours = 1,
	[Parameter(Mandatory=$false,
				ParameterSetName = 'LastDays',
				HelpMessage = 'How many days back to get Windows Events (default=1)')]
				[int]$LastDays = 1,
	[Parameter(Mandatory=$false,
				ParameterSetName = 'MinutesSinceLastBoot',
				HelpMessage = 'How many minutes before last boot')]
				[int]$MinutesBeforeLastBoot=$null,
	[Parameter(Mandatory=$true,
				ParameterSetName = 'MinutesSinceLastBoot',
				HelpMessage = 'How many minutes after last boot')]
				[int]$MinutesAfterLastBoot=$null,
	[Parameter(Mandatory=$false,
				ParameterSetName = 'StartEndTimeSpecified',
				HelpMessage = 'Specify StartTime for example in format: yyyy-MM-dd HH:mm:ss (2024-12-31 22:05:00)')]
				[String]$StartTime,
	[Parameter(Mandatory=$false,
				ParameterSetName = 'StartEndTimeSpecified',
				HelpMessage = 'Specify EndTime for example in format: yyyy-MM-dd HH:mm:ss (2024-12-31 23:05:00)')]
				[String]$EndTime,
	[Parameter(Mandatory=$false)]
				[Switch]$LogViewerUI,
	[Parameter(Mandatory=$false,
				ParameterSetName = 'RealtimeLogViewerUI')]
				[Switch]$RealtimeLogViewerUI,
	[Parameter(Mandatory=$false)]
				[Switch]$AllEvents,
	[Parameter(Mandatory=$false)]
				[String[]]$IncludeSelectedKnownRulesCategoriesOnly,
	[Parameter(Mandatory=$false)]
				[String[]]$ExcludeSelectedKnownRulesCategories,
	[Parameter(Mandatory=$false)]
				[Switch]$SortDescending
)


$ScriptVersion = "0.71"
#$TimeOutBetweenGraphAPIRequests = 300

# Save timeline objects to this List
$observedTimeline = [System.Collections.Generic.List[PSObject]]@()

# TimeLine entry index
# This might be used in HTML table for sorting entries
$Script:observedTimeLineIndexToHTMLTable=0



################ Functions Start ################

# region Functions

    Function RecordStatusToTimeline {
        param
        (
			[Parameter(Mandatory=$false)] $dateTimeObject=$null,
			[Parameter(Mandatory=$false)] [String] $date,
			[Parameter(Mandatory=$true)] [String] $logName,
			[Parameter(Mandatory=$false)] [String] $providerName=$null,
			[Parameter(Mandatory=$false)] [String] $id,
			[Parameter(Mandatory=$false)] [String] $levelDisplayName,
            [Parameter(Mandatory=$false)] [String] $message,
			[Parameter(Mandatory=$false)] [String] $messageToolTip,
			[Parameter(Mandatory=$false)] [String] $KnownCategoryName=$null,
			[Parameter(Mandatory=$false)] [String] $color=$null,
			[Parameter(Mandatory=$false)] [int] $eventNumber=$null
        )

		# Some .evtx files have lines without message
		# So that is reason $message is not mandatory parameter but it is checked here
		if($message) {

			# Add index counter to timeline to possibly help sorting events later in the report
			$Script:observedTimeLineIndexToHTMLTable++


			# Add Color based on LevelDisplayName text if Color was NOT specified manually
			# If color was specified manually then we use manually specified value as is
			if(-not $color) {
				if($levelDisplayName -like '*Success') {
					$color = 'White'
				}

				if($levelDisplayName -like '*Warning') {
					$color = 'Yellow'
				}


				if($levelDisplayName -like '*Error') {
					$color = 'Red'
				}
			}


			# Set date string formatting
			#$dateToTimeLine = "$(Get-Date $date -Format 'yyyy-MM-dd HH:mm:ss')"
			try {
				# If trying to change dateTime string format fails
				# Then we show date as it was presented in Log
				#
				# This is because EN-US and for example nordic Windows may have different date syntax which may fail otherwise
				
				# This will replace : time separator with . which is bad
				# Because with log files time separator is :
				#$dateToTimeLine = Get-Date $date -Format 'yyyy-MM-dd HH:mm:ss.fff' -ErrorAction SilentlyContinue

				# Not sure if this Try-Catch format check/change should be taken away
				# and just rely what we got as parameter
				$dateToTimeLine = $date

			} catch {
				$dateToTimeLine = $date
			}

			$CustomTimelineObject = ([PSCustomObject]@{
				'DateTimeObject' = $dateTimeObject
				'Index' = $Script:observedTimeLineIndexToHTMLTable
				'EventNumber' = $eventNumber
				'Date' = $dateToTimeLine
				'LogName' = $logName
				'ProviderName' = $providerName
				'Id' = $id
				'LevelDisplayName' = $levelDisplayName
				'Message' = $message
				'MessageToolTip' = $messageToolTip
				'KnownCategoryName' = $KnownCategoryName
				'Color' = $color
				})

			# DEBUG $CustomTimelineObject
			#$CustomTimelineObject
			#Pause

			$observedTimeline.add($CustomTimelineObject)

			#Start-Sleep -Milliseconds 10
			
			# If $RealtimeLogViewerUI is specified then show events in Out-GridView as they are processed realtime
			if($RealtimeLogViewerUI) {
				$Script:OutGridViewRealtime.Process(($CustomTimelineObject | Select-Object -Property * -ExcludeProperty Color))
			}
		} else {
			# Message was empty
			#Write-Verbose "DEBUG: `$Message was empty"
		}
	}

	### HTML Report helper functions ###
	function Fix-HTMLSyntax {
		Param(
			$html
		)

		$html = $html.Replace('&lt;', '<')
		$html = $html.Replace('&gt;', '>')
		$html = $html.Replace('&quot;', '"')

		return $html
	}

	function Fix-HTMLColumns {
		Param(
			$html
		)

		# Rename column headers
		$html = $html -replace '<th>@odata.type</th>','<th>App type</th>'
		$html = $html -replace '<th>displayname</th>','<th>App name</th>'
		$html = $html -replace '<th>assignmentIntent</th>','<th>Assignment Intent</th>'
		$html = $html -replace '<th>assignmentTargetGroupDisplayName</th>','<th>Target Group</th>'
		$html = $html -replace '<th>assignmentFilterDisplayName</th>','<th>Filter name</th>'
		$html = $html -replace '<th>FilterIncludeExclude</th>','<th>Filter Intent</th>'
		$html = $html -replace '<th>publisher</th>','<th>Publisher</th>'
		$html = $html -replace '<th>productVersion</th>','<th>Version</th>'
		$html = $html -replace '<th>filename</th>','<th>Filename</th>'
		$html = $html -replace '<th>createdDateTime</th>','<th>Created</th>'
		$html = $html -replace '<th>lastModifiedDateTime</th>','<th>Modified</th>'

		return $html
	}

 # startregion function_read_file
function Read-LogFile {
	Param(
			[Parameter(Mandatory=$true,
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
				[Alias("FullName")]
				$Path = $null,
				[Parameter(Mandatory=$true)]
				$StartTimeObject,
				[Parameter(Mandatory=$true)]
				$EndTimeObject
	)
	
	# Check that log file exists
	if(-not (Test-Path $Path)) {
		# Log file does not exist
		# Return with $false
		Write-Host "Log file does NOT exist. Skipping file: $Path" -ForegroundColor 'Red'
		return
	} else {
		# File exists
		
		# Get file object so we get more properties from file
		$LogFileObject = Get-ChildItem -File -Path $Path
		
		if(-not $LogFileObject) {
			Write-Host "ERROR: Skipping file... Could not read log file information: $Path" -ForegroundColor 'Red'
			return
		}
	}

	# Get FileName without full Path
	
	# This would work if we didn't have file object
	#$LogFileName = Split-Path -Path $Path -Leaf

	# Get Name from file object
	$LogFileName = $LogFileObject.Name

	# File name without extension .log
	$LogFileBaseName = $LogFileObject.BaseName
	
	# File extension .log
	$LogFileExtension = $LogFileObject.Extension


	# Check if we have Known events for this file
	$KnownLogEventRuleObjectsForThisLogFileType = @()
	
	Foreach($EventRuleObject in ($EventRulesArray.KnownEventRules | Where-Object LogType -eq "$LogFileExtension")) {
		$EventRuleLogFileName = $EventRuleObject.LogFileName
		$EventRuleLogFileNameWithoutExtension = $EventRuleLogFileName.Replace("$LogFileExtension",'')
		

		if($LogFileBaseName -like "$($EventRuleLogFileNameWithoutExtension)*") {

			# Escape special characters in the Message field using regex escape
			# Later we use -match operator to test if message is a match to rule
			$EventRuleObject.Message = [regex]::Escape($EventRuleObject.Message)
			
			# Add rule object to array
			# Yes this is kind of wrong way to add to array variables (copies whole array and add new object) but good enough for now :)
			$KnownLogEventRuleObjectsForThisLogFileType += $EventRuleObject
		}
	}
	
	if($KnownLogEventRuleObjectsForThisLogFileType) {
		# Found Event Rules for this log file type
		
		# DEBUG
		#$KnownLogEventRuleObjectsForThisLogFileType
		#Pause

	} else {
		if(-not $AllEvents) {
			# Did not find Event Rules for this log file type
			# And -AllEvents is NOT specified
			Write-Host "No Event Rules for this file. Skipping to next file..."
			return
		}
	}


	# Assign variables
	
	$LogFileTypeRecognized = $null
	$OldestLogEntryDateTimeObject = $null
	$NewestLogEntryDateTimeObject = $null

	
	# CMTrace type log file
	# This matches single line full log entry
	#$CMTraceSingleLineRegex = '^\<\!\[LOG\[(.*)]LOG\].*\<time="([0-9]{1,2}):([0-9]{1,2}):([0-9]{1,2}).([0-9]{1,})".*date="([0-9]{1,2})-([0-9]{1,2})-([0-9]{4})" component="(.*?)" context="(.*?)" type="(.*?)" thread="(.*?)" file="(.*?)">$'
	
	$CMTraceSingleLineRegex ='^\<\!\[LOG\[(.*)]LOG\].*\<time="([0-9]{1,2}):([0-9]{1,2}):([0-9]{1,2}).([0-9]{1,}).*".*date="([0-9]{1,2})-([0-9]{1,2})-([0-9]{4})" component="(.*?)" context="(.*?)" type="(.*?)" thread="(.*?)" file="(.*?)">$'

	# Start of multiline log entry
	$CMTraceFirstLineOfMultiLineLogRegex = '^\<\!\[LOG\[(.*)$'

	# End of multiline log entry
	#$CMTraceLastLineOfMultiLineLogRegex = '^(.*)\]LOG\]\!>\<time="([0-9]{1,2}):([0-9]{1,2}):([0-9]{1,2}).([0-9]{1,})".*date="([0-9]{1,2})-([0-9]{1,2})-([0-9]{4})" component="(.*?)" context="(.*?)" type="(.*?)" thread="(.*?)" file="(.*?)">$'
	
	$CMTraceLastLineOfMultiLineLogRegex ='^(.*)\]LOG\]\!>\<time="([0-9]{1,2}):([0-9]{1,2}):([0-9]{1,2}).([0-9]{1,}).*".*date="([0-9]{1,2})-([0-9]{1,2})-([0-9]{4})" component="(.*?)" context="(.*?)" type="(.*?)" thread="(.*?)" file="(.*?)">$'


	# Office C2RClient logs in C:\Windows\Logs -directory
	# First line in log file is
	# 'Timestamp	Process	TID	Area	Category	EventID	Level	Message	Correlation'
	$OfficeC2RClientLogHeaders = 'Timestamp	Process	TID	Area	Category	EventID	Level	Message	Correlation'

	# Get total number of log entries in variable before adding log file events
	$LogEntriesBeforeLogfileProcessing = $observedTimeline.Count

	#####################
	# Read log file
	$LogFileLines = Get-Content -Path $Path

	# Figure out what log file syntax is used
	Foreach($Line in $LogFileLines) {

		# Test for CMTrace-type log file
		if(($Line -Match $CMTraceSingleLineRegex) -or ($Line -Match $CMTraceFirstLineOfMultiLineLogRegex) -or ($Line -Match $CMTraceLastLineOfMultiLineLogRegex)) {
			# Recognized CMTrace type log file
			$LogFileTypeRecognized = 'CMTrace'
			
			Write-Host "Log file recognized as CMTrace-type"
			
			# Get CMTrace log file oldest log entry
			if($Line -Match $CMTraceSingleLineRegex) {
				$Hour = $Matches[2]
				$Minute = $Matches[3]
				$Second = $Matches[4]
				
				$MilliSecondFull = $Matches[5]
				# Cut milliseconds to 0-999
				# Time unit is so small that we don't even bother to round the value
				$MilliSecond = $MilliSecondFull.Substring(0,3)

				$Month = $Matches[6]
				$Day = $Matches[7]
				$Year = $Matches[8]
				
				$Component = $Matches[9]
				$Context = $Matches[10]
				$Type = $Matches[11]
				$Thread = $Matches[12]
				$File = $Matches[13]

				$Param = @{
					Hour=$Hour
					Minute=$Minute
					Second=$Second
					MilliSecond=$MilliSecond
					Year=$Year
					Month=$Month
					Day=$Day
				}

				$OldestLogEntryDateTimeObject = Get-Date @Param
				
				# Break out from Foreach-loop
				Break
			}

			if($Line -Match $CMTraceLastLineOfMultiLineLogRegex) {
				$Hour = $Matches[2]
				$Minute = $Matches[3]
				$Second = $Matches[4]
				
				$MilliSecondFull = $Matches[5]
				# Cut milliseconds to 0-999
				# Time unit is so small that we don't even bother to round the value
				$MilliSecond = $MilliSecondFull.Substring(0,3)
				
				$Month = $Matches[6]
				$Day = $Matches[7]
				$Year = $Matches[8]
				
				$Component = $Matches[9]
				$Context = $Matches[10]
				$Type = $Matches[11]
				$Thread = $Matches[12]
				$File = $Matches[13]

				$Param = @{
					Hour=$Hour
					Minute=$Minute
					Second=$Second
					MilliSecond=$MilliSecond
					Year=$Year
					Month=$Month
					Day=$Day
				}

				$OldestLogEntryDateTimeObject = Get-Date @Param
				
				# Break out from Foreach-loop
				Break
			}
		}

		# Test for OfficeClick2Run client logs
		if($Line -eq $OfficeC2RClientLogHeaders) {
			# Recognized Office Click2Run log file
			$LogFileTypeRecognized = 'OfficeC2R'
			
			Write-Host "Log file recognized as Office Click2Run-type"
			
			# Break out from Foreach-loop
			Break
		}
		
	}


	# Get CMTrace log file newest log entry
	if($LogFileTypeRecognized -eq 'CMTrace') {
		# Process file from end to start
		$i = $LogFileLines.Count - 1

		While($i -ge 0) {
			$Line = $LogFileLines[$i]

			if($Line -Match $CMTraceSingleLineRegex) {
				$Hour = $Matches[2]
				$Minute = $Matches[3]
				$Second = $Matches[4]
				
				$MilliSecondFull = $Matches[5]
				# Cut milliseconds to 0-999
				# Time unit is so small that we don't even bother to round the value
				$MilliSecond = $MilliSecondFull.Substring(0,3)

				$Month = $Matches[6]
				$Day = $Matches[7]
				$Year = $Matches[8]
				
				$Component = $Matches[9]
				$Context = $Matches[10]
				$Type = $Matches[11]
				$Thread = $Matches[12]
				$File = $Matches[13]

				$Param = @{
					Hour=$Hour
					Minute=$Minute
					Second=$Second
					MilliSecond=$MilliSecond
					Year=$Year
					Month=$Month
					Day=$Day
				}

				$NewestLogEntryDateTimeObject = Get-Date @Param
				
				# Break out from While-loop
				Break
			}

			if($Line -Match $CMTraceLastLineOfMultiLineLogRegex) {
				$Hour = $Matches[2]
				$Minute = $Matches[3]
				$Second = $Matches[4]
				
				$MilliSecondFull = $Matches[5]
				# Cut milliseconds to 0-999
				# Time unit is so small that we don't even bother to round the value
				$MilliSecond = $MilliSecondFull.Substring(0,3)
				
				$Month = $Matches[6]
				$Day = $Matches[7]
				$Year = $Matches[8]
				
				$Component = $Matches[9]
				$Context = $Matches[10]
				$Type = $Matches[11]
				$Thread = $Matches[12]
				$File = $Matches[13]

				$Param = @{
					Hour=$Hour
					Minute=$Minute
					Second=$Second
					MilliSecond=$MilliSecond
					Year=$Year
					Month=$Month
					Day=$Day
				}

				$NewestLogEntryDateTimeObject = Get-Date @Param
				
				# Break out from While-loop
				Break
			}

			$i--
		}
	}
	
	# Office Click2Run log files are csv which can be imported as is
	if($LogFileTypeRecognized -eq 'OfficeC2R') {
		# Import csv file

		# Not supported for now
		return

		Write-Host "Import Office Click2Run log file (csv-type): $Path"
		$LogFileLines = Import-Csv -Path $Path -Delimiter "`t"
		$Success = $?
		if($Success) {
			Write-Host "Success" -ForegroundColor 'Green'
		} else {
			return
		}
	}

	# Check if we detected log file type
	if(-not $LogFileTypeRecognized) {
		Write-Host "Unknown log file type. Skipping this file..." -ForegroundColor Yellow
		return
	}



	# Check if log file entries are within specified StartTime and EndTime
	# and skip file if isn't
	if($OldestLogEntryDateTimeObject -and $NewestLogEntryDateTimeObject) {
		# We have oldest and newest log entry DateTime
		
		# Oldest log entry is newer than specified EndTime
		if($OldestLogEntryDateTimeObject -gt $EndTimeObject) {
			Write-Host "Log's oldest entry is newer than specified EndTime. Skipping file..." -ForegroundColor Yellow
			return
		}

		# Newest log entry is older than specified StartTime
		if($NewestLogEntryDateTimeObject -lt $StartTimeObject) {
			Write-Host "Log's latest entry is older than specified StartTime. Skipping file..." -ForegroundColor Yellow
			return
		}
	}


	# CMTrace helper variables for multiline processing
	$MultilineLogEntryStartsArrayIndex=0
	$MultilineLogEntryStartFound=$False

	#$StartTime=(Get-Date).AddDays(-1)
	
	# Process log file line by line
	$Line = $null
	
	# We'll save linenumbers also so we can sort log entries with exactly same dateTime value
	# Some log entries have exactly same dateTime value with 7 digits in milliseconds
	# Which break sorting unless we sort also by lineNumber property
	$LineNumber = 0
	foreach($Line in $LogFileLines) {

		# Increase LineNumber value first here
		# Do not move to end because we will Continue Foreach which will skip rest of code
		$LineNumber++

		# Testing if this line removes UTF8-BOM character from start of the line
		# This has been seen in few lines in smsts.log and it causes problems
		$Line = $Line -replace "^\uFEFF", ""


		if($LogFileTypeRecognized -eq 'CMTrace') {
			
			# Get data from CurrentLogEntry
			if($Line -Match $CMTraceSingleLineRegex) {
				# This matches single line log entry

				$MultilineLogEntryStartFound=$False
				
				# Regex found match
				$LogMessage = $Matches[1].Trim()

				$Hour = $Matches[2]
				$Minute = $Matches[3]
				$Second = $Matches[4]
				
				$MilliSecondFull = $Matches[5]
				# Cut milliseconds to 0-999
				# Time unit is so small that we don't even bother to round the value
				$MilliSecond = $MilliSecondFull.Substring(0,3)

				$Month = $Matches[6]
				$Day = $Matches[7]
				$Year = $Matches[8]
				
				$Component = $Matches[9]
				$Context = $Matches[10]
				$Type = $Matches[11]
				$Thread = $Matches[12]
				$File = $Matches[13]

				$Param = @{
					Hour=$Hour
					Minute=$Minute
					Second=$Second
					MilliSecond=$MilliSecond
					Year=$Year
					Month=$Month
					Day=$Day
				}

				$LogEntryDateTimeObject = Get-Date @Param
				#Write-Host "DEBUG `$LogEntryDateTime: $LogEntryDateTime" -ForegroundColor Yellow
				
				# This works for humans but does not sort
				#$DateTimeToLogFile = "$($Hour):$($Minute):$($Second).$MilliSecondFull $Day/$Month/$Year"

				# Add leading 0 so sorting works right
				if($Month -like "?") { $Month = "0$Month" }

				# Add leading 0 so sorting works right
				if($Day -like "?") { $Day = "0$Day" }

				# This does sorting right way
				#$DateTimeToLogFile = "$Year-$Month-$Day $($Hour):$($Minute):$($Second).$MilliSecondFull"
				$DateTimeToLogFile = "$Year-$Month-$Day $($Hour):$($Minute):$($Second).$MilliSecond"


				# Save log line to report (object array variable)
				#RecordStatusToTimeline -date $LogEntryDateTime -logName $LogFileObject.Name -id '' -levelDisplayName '' -message $LogMessage
				#RecordStatusToTimeline -dateTimeObject $LogEntryDateTimeObject -date $DateTimeToLogFile -logName $LogFileObject.Name -id '' -levelDisplayName '' -message $LogMessage -eventNumber $LineNumber

				# Save log line to report
				if(($LogEntryDateTimeObject -ge $StartTimeObject) -and ($LogEntryDateTimeObject -le $EndTimeObject)) {

					$MessageAlreadyRecorded = $false

					# Check if LogMessage is specified in known rules
					Foreach($KnownLogEventRuleObject in $KnownLogEventRuleObjectsForThisLogFileType) {
						# DEBUG
						#$LogMessage
						#$KnownLogEventRuleObject
						#Pause
						
						# -match handles special characters in EventRules.Json
						# and -match is faster than -like !!!
						if ($LogMessage -match $KnownLogEventRuleObject.Message) {
							if($KnownLogEventRuleObject.Color) {
								$Color = $KnownLogEventRuleObject.Color
							} else {
								$Color = $null
							}

							# Set MessageToolTip if exists
							if($KnownLogEventRuleObject.ToolTip) {
								$MessageToolTip = $KnownLogEventRuleObject.ToolTipText
							} else {
								$MessageToolTip = $null
							}
							
							RecordStatusToTimeline -dateTimeObject $LogEntryDateTimeObject -date $DateTimeToLogFile -logName $LogFileObject.Name -ProviderName $Component -id '' -levelDisplayName '' -message $LogMessage -eventNumber $LineNumber -Color $Color -MessageToolTip $MessageToolTip -KnownCategoryName $KnownLogEventRuleObject.CategoryName
							
							$MessageAlreadyRecorded = $true
					
							# DEBUG
							#Write-Host "Found known LogMessage: $LogMessage"
							#$KnownLogEventRuleObject
							#Pause
							
							# Break out from Foreach-loop
							Break
						}
					}

					if($MessageAlreadyRecorded) {
						# Continue to next line in log file (go to next foreach loop)
						Continue						
					}

					#if($AllEvents -and (-not $MessageAlreadyRecorded)) {
					if($AllEvents) {
						# We save all events
						
						RecordStatusToTimeline -dateTimeObject $LogEntryDateTimeObject -date $DateTimeToLogFile -logName $LogFileObject.Name -ProviderName $Component -id '' -levelDisplayName '' -message $LogMessage -eventNumber $LineNumber
					}
				}

				# Continue to next line in log file (go to next foreach loop)
				Continue

			} elseif ($Line -Match $CMTraceFirstLineOfMultiLineLogRegex) {
				# This is start of multiline log entry
				# Single line regex did not get results so we are dealing multiline case separately here

				#Write-Host "DEBUG Start of multiline regex: $CurrentLogEntry" -ForegroundColor Yellow

				$MultilineLogEntryStartFound=$True

				# Regex found match
				$LogMessage = $Matches[1].Trim()
				
				$LogEntryDateTimeObject = $null
				$LogEntryDateTime = $null
				$DateTimeToLogFile = $null
				#$Component = ''
				#$Context = ''
				#$Type = ''
				#$Thread = ''
				#$File = ''

				
				# Save log line to report (object array variable)
				# We save this always because we don't yet know if this log entry is within
				# specified start and endtime
				#
				# We will remove this later if we find dateTime to be out of scope
				RecordStatusToTimeline -dateTimeObject $LogEntryDateTimeObject -date $LogEntryDateTime -logName $LogFileObject.Name -id '' -levelDisplayName '' -message $LogMessage -eventNumber $LineNumber


				# Save index for multiline first log entry
				$MultilineLogEntryStartsArrayIndex = $observedTimeline.Count - 1
				
				# Continue to next line in log file (go to next foreach loop)
				Continue
				
			} elseif ($Line -Match $CMTraceLastLineOfMultiLineLogRegex) {
				# This is end of multiline log entry
				# Single line regex did not get results so we are dealing multiline case separately here

				# Regex found match
				$LogMessage = $Matches[1].Trim()

				$Hour = $Matches[2]
				$Minute = $Matches[3]
				$Second = $Matches[4]
				
				$MilliSecondFull = $Matches[5]
				# Cut milliseconds to 0-999
				# Time unit is so small that we don't even bother to round the value
				$MilliSecond = $MilliSecondFull.Substring(0,3)
				
				$Month = $Matches[6]
				$Day = $Matches[7]
				$Year = $Matches[8]
				
				$Component = $Matches[9]
				$Context = $Matches[10]
				$Type = $Matches[11]
				$Thread = $Matches[12]
				$File = $Matches[13]

				$Param = @{
					Hour=$Hour
					Minute=$Minute
					Second=$Second
					MilliSecond=$MilliSecond
					Year=$Year
					Month=$Month
					Day=$Day
				}

				$LogEntryDateTimeObject = Get-Date @Param
				#Write-Host "DEBUG `$LogEntryDateTime: $LogEntryDateTime" -ForegroundColor Yellow

				# This works for humans but does not sort
				#$DateTimeToLogFile = "$($Hour):$($Minute):$($Second).$MilliSecondFull $Day/$Month/$Year"

				# Add leading 0 so sorting works right
				if($Month -like "?") { $Month = "0$Month" }

				# Add leading 0 so sorting works right
				if($Day -like "?") { $Day = "0$Day" }

				# This does sorting right way
				#$DateTimeToLogFile = "$Year-$Month-$Day $($Hour):$($Minute):$($Second).$MilliSecondFull"
				$DateTimeToLogFile = "$Year-$Month-$Day $($Hour):$($Minute):$($Second).$MilliSecond"


				if($MultilineLogEntryStartFound) {
					$Multiline = '---->'
				} else {
					$Multiline = ''
				}


				# Add message and DateTime value to last object in $observedTimeline
				if($MultilineLogEntryStartFound) {

					# Reset multiline variable
					$MultilineLogEntryStartFound=$False

					if(-not ($observedTimeline[-1])) {
						Write-Host "DEBUG: Current `$Line: $Line"
						Write-Host "Previous lines"
						$LogFileLines[$LineNumber-1]
						$LogFileLines[$LineNumber-2]
						$LogFileLines[$LineNumber-3]
						
						Write-Host "DEBUG: `$observedTimeline[-1] ja -2"
						Write-Host ""
						$observedTimeline[-1]
						Write-Host ""
						$observedTimeline[-2]
						Write-Host ""
						$observedTimeline[-3]
						Write-Host ""
					}



					# Get last object array index
					#$LastObjectIndex = $observedTimeline.Count - 1

					# Add ProviderName (CMTrace component)
					($observedTimeline[-1]).ProviderName = $Component

					
					# Add message text to existing multiline log entry
					# last array object is index -1
					($observedTimeline[-1]).Message = "$($observedTimeline[-1].Message)`n$Message"

					# Add log dateTime to existing (last) multiline log entry
					
					# 6 numbers in milliseconds which will make sorting fail?
					#$observedTimeline[-1].Date = $DateTimeToLogFile

					# 3 numbers in milliseconds which is same than with other log events
					# so sorting works (at least better)
					($observedTimeline[-1]).Date = $DateTimeToLogFile
					
					($observedTimeline[-1]).DateTimeObject = $LogEntryDateTimeObject

					# Delete last log entry if it is out of timing scope
					# Before StartTime OR after EndTime
					if(($LogEntryDateTimeObject -le $StartTimeObject) -or ($LogEntryDateTimeObject -ge $EndTimeObject)) {
						$observedTimeline.RemoveAt($observedTimeline.Count - 1)

						# Continue to next line in log file (go to next foreach loop)
						Continue
					}


					# Check if LogMessage is specified in known rules
					$MessageMatchesKnownRule = $false
					Foreach($KnownLogEventRuleObject in $KnownLogEventRuleObjectsForThisLogFileType) {
						# DEBUG
						#$LogMessage
						#$KnownLogEventRuleObject
						#Pause

						# -match handles special characters in EventRules.Json
						# and -match is faster than -like !!!
						if($observedTimeline[-1].Message -match $KnownLogEventRuleObject.Message) {
							# DEBUG
							#Write-Host "Found known LogMessage: $LogMessage"
							#$KnownLogEventRuleObject
							#Pause

							$MessageMatchesKnownRule = $true
							
							# Mark log message as known
							$observedTimeline[-1].Color = $KnownLogEventRuleObject.Color

							# Set MessageToolTip value
							$MessageToolTip = $KnownLogEventRuleObject.ToolTipText
							if($MessageToolTip) {
								$observedTimeline[-1].MessageToolTip = $MessageToolTip
							}

							# Set KnownCategoryName
							$observedTimeline[-1].KnownCategoryName = $KnownLogEventRuleObject.KnownCategoryName

							# Break out from Foreach-loop
							Break
						}
					}
					
					if($MessageMatchesKnownRule) {
						# We keep the log message
					} else {
						# No Event Rules found for this log message

						if($AllEvents) {
							# -AllEvents specified so we are keeping this LogMessage (and all LogMessages)
						} else {
							# Remove this log entry
							$observedTimeline.RemoveAt($observedTimeline.Count - 1)
						}
					}
					
				} else {
					# Found Multiline log end message but start message hasn't been found earlier
					# We should never get here
					Write-Host "Found multiline log entry ending message without start log entry. This should not be possible..." -ForegroundColor Yellow
					
					# DEBUG
					Write-Host "DEBUG: $LogMessage"
					#Pause
					
					# Reset multiline variable
					$MultilineLogEntryStartFound=$False
				}
				
				# Continue to next line in log file (go to next foreach loop)
				Continue
					
			} else {
				# We didn't catch log entry with our regex
				# This should be multiline log entry but not first or last line in that log entry
				# This can also be some line that should be matched with (other) regex
				
				#Write-Host "DEBUG: $Line"  -ForegroundColor Yellow

				# Add log message to last object in $observedTimeline
				if($MultilineLogEntryStartFound) {
					# Get last object array index
					#$LastObjectIndex = $observedTimeline.Count - 1
					
					# Add message text to existing multiline log entry
					# last array object is index -1
					$observedTimeline[-1].Message = "$($observedTimeline[-1].Message)`n$Line"
				} else {
					# We should not get here unless string is not on CMTrace syntax
					# This sometimes happens when log text gets truncated and is not save entirely
					Write-Verbose "CMTrace log entry wrong format: $Line"
				}
				
				# Continue to next line in log file (go to next foreach loop)
				Continue
				
			}
		} # End CMTrace If-clause

		if($LogFileTypeRecognized -eq 'OfficeC2R') {
			
		} # End OfficeC2R If-clause

	} # Foreach end single log file line by line foreach

	$LogEntriesAfterLogfileProcessing = $observedTimeline.Count
	$LogEntriesAdded = $LogEntriesAfterLogfileProcessing -$LogEntriesBeforeLogfileProcessing
	
	Write-Host "Log entries added: $LogEntriesAdded"
	Write-Host ""
} # endregion function_read_file


# startregion function_read_eventlog
function Read-EventLog {
	Param(
			[Parameter(Mandatory=$false,
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$false)]
				$EventLog = $null,
				[Parameter(Mandatory=$false,
                ValueFromPipeline=$false,
                ValueFromPipelineByPropertyName=$false)]
				$EventLogPath = $null,
				[Parameter(Mandatory=$true)]
				$StartTimeObject,
				[Parameter(Mandatory=$true)]
				$EndTimeObject
	)	

	if($EventLog) {
		# Getting EventLog from live OS

		$EventLogName = $EventLog.LogName
		Write-Host "Processing $EventLogName"
		
		if($AllEvents) {
			# Get all Events between StartTime and EndTime

			# Getting events from live OS
			$GatheredWinEvents = $null
			$GatheredWinEvents = Get-WinEvent -FilterHashtable @{ 
				LogName=$EventLogName
				StartTime=$StartTimeObject
				EndTime=$EndTimeObject
			} -ErrorAction SilentlyContinue
			
			Write-Host "Found events count $($GatheredWinEvents.Count)"
		} else {
			# Get only known Events
			
			# Get Id's we know from EventRules.json
			
			# Before EventsRules.json syntax change
			#$KnownEventRulesIds = $EventRulesArray | Where-Object LogName -eq $EventLogName | Select-Object -ExpandProperty Id

			# Before EventRules.json syntax change
			$KnownEventRulesIds = $EventRulesArray | Foreach-Object { $_.KnownEventRules | Where-Object Channel -eq $EventLogName | Select-Object -ExpandProperty Id }

			# Note this will get extra events which are not known
			# because Event is identified by Id and ProviderName
			# But this approach still gets us significantly less Events than just getting all events
			
			# Get all Events from current eventlog (in foreach loop)
			$GatheredWinEvents = $null

			if($EventLog -and $EventLogName) {
				# Getting only known events from live OS
				if($KnownEventRulesIds) {
					$GatheredWinEvents = Get-WinEvent -FilterHashtable @{
						LogName=$EventLogName
						Id=$KnownEventRulesIds
						StartTime=$StartTimeObject
						EndTime=$EndTimeObject
					} -ErrorAction SilentlyContinue
					
					Write-Host "Found events count $($GatheredWinEvents.Count)"
					
				} else {
					Write-Host "No known events for this Event log. Skipping file..."
				}
			}			
			
		}
	} elseif ($EventLogPath) {
		# Reading saved EventLog file
			
		# Check that log file exists
		if(-not (Test-Path $EventLogPath)) {
			# Log file does not exist
			# Return with $false
			Write-Host "Log file does NOT exist. Skipping file: $EventLogPath" -ForegroundColor 'Red'
			return
		} else {
			# File exists
			
			# Get file object so we get more properties from file
			$EventLogFileObject = Get-ChildItem -File -Path $EventLogPath
			
			if(-not $EventLogFileObject) {
				Write-Host "ERROR: Skipping file... Could not read log file information: $EventLogPath" -ForegroundColor 'Red'
				return
			}
		}

		# Get FileName without full Path
		
		# This would work if we didn't have file object
		#$LogFileName = Split-Path -Path $Path -Leaf

		# Get Name from file object
		$LogFileName = $EventLogFileObject.Name

		# File name without extension .log
		$LogFileBaseName = $EventLogFileObject.BaseName

		Write-Host "Getting Events from saved file: $LogFileName"

		##########################
		# Get Events from file
		
		$GatheredWinEvents = $null
		$GatheredWinEvents = Get-WinEvent -Path $EventLogPath -Oldest

		if(-not $GatheredWinEvents) {
			Write-Host "Did not find any Windows events in the file. Skipping to next file..." -ForegroundColor Yellow
			return
		}
		
		# Filter out null messages
		$GatheredWinEvents = $GatheredWinEvents | Where-Object { $_.Message -ne $null }

		# Get LogName property from first log event
		$EventLogName = $GatheredWinEvents[0].LogName

		Write-Host "Found LogName $EventLogName"
		Write-Host "Total events found $($GatheredWinEvents.Count)"

		# Filter Found events using StartTime and EndTime
		# Filter out messages outside StartTimeObject and EndTimeObject
		$GatheredWinEvents = $GatheredWinEvents | Where-Object { ($_.TimeCreated -ge $StartTimeObject) -and ($_.TimeCreated -le $EndTimeObject) }
		Write-Host "Total events after filtering $($GatheredWinEvents.Count)"
		
	} else {
		# We should never get here
		Write-Host "Eventlog name or path was not specified. Skipping to next Event log..."
		return
	}


	# Dirty? workaround!
	# Make a copy of $GatheredWinEvents Events Array because original Events have read-only Properties which we want to change
	# We can change Property values for copied objects
	$GatheredWinEvents = $GatheredWinEvents | Select-Object -Property *

	# Add KnownEvents to HashTable which we can query way faster than using | Where-Object syntax in each round in Foreach-loop
	$KnownEventsHashTable = @{}

	# Add KnownEvents to HashTable for quicker search
	
	# .etl files may not have LogName property
	if($EventLogName) {
		foreach($EventRule in ($EventRulesArray | Foreach-Object { $_.KnownEventRules | Where-Object Channel -eq $EventLogName })) {
			# Create variables
			
			#$PropertyName = $EventRule.Id
			
			# PowerShell (core) will import json numbers as int64
			# Event IDs are Int (32-bit).
			# If you create HashTable key with json number Int64 then that is not same than Event ID Int when comparing later
			# Because HashTable requires exactly same variable type
			# This problem is only with Powershell (core). With Windows Powershell json import create Int numbers
			#
			# Original with only Id as key
			#$PropertyName = [int]$EventRule.Id
			
			# Better. Key is Id+Source
			$PropertyName = "$($EventRule.Id)-$($EventRule.ProviderName)"
			$PropertyValue = $EventRule
			
			# Add hashtable entry
			$KnownEventsHashTable.add($PropertyName, $PropertyValue)
		}
	}	
	#Start-Sleep -Seconds 1
	
	# DEBUG. Yes me and you should use Debugger properly and not like this! :)
	#Write-Host "DEBUG `$KnownEventsHashTable"
	#$KnownEventsHashTable
	#Write-Host ""
	#Pause
	
	# Process and add information to known Events
	#$KnownEventRulesForCurrentEventLog = $EventRulesArray | Where-Object LogName -eq $EventLogName

	Write-Host "Find and process known events..."

	Foreach($EventObject in $GatheredWinEvents) {
		# DEBUG
		#Write-Host "DEBUG: Process Event $($EventObject.id)"
		
		# HashTable is much faster than below syntax!
		# Check if we have Event Rule for current event Id
		# We will go trough all events.
		# Maybe optimize this later to only process Event Id which are known
		#if($KnownEventsHashTable.ContainsKey([int]$EventObject.Id)) {
			
		$HashTableKeyValueToSearch = "$($EventObject.Id)-$($EventObject.ProviderName)"
			
		if($KnownEventsHashTable.ContainsKey($HashTableKeyValueToSearch)) {
			# Current event is known and specified in EventRules.json

			# DEBUG
			#Write-Host "Found Known Event in $($EventObject.LogName) Id $($EventObject.Id)"

			# Get Known Event rule for this Id from HashTable
			$HashKeyToFind = "$($EventObject.Id)-$($EventObject.ProviderName)"
			$KnownEventRuleObject = $KnownEventsHashTable[$HashKeyToFind]
	
			# DEBUG
			#$KnownEventRuleObject
			#Pause
	
			# Add 'Known' before LevelDisplayName
			$EventObject.LevelDisplayName = "Known $($EventObject.LevelDisplayName)"

			# Add MessageToolTip information
			if($KnownEventRuleObject.ToolTipText) {
				# Add ToolTipText property to Event object
				$EventObject | Add-Member -MemberType noteProperty -Name MessageToolTip -Value "$($KnownEventRuleObject.ToolTipText)"
			}

			# Add KnownCategoryName property to event object
			$EventObject | Add-Member -MemberType noteProperty -Name KnownCategoryName -Value "$($KnownEventRuleObject.CategoryName)"

			$EventObjectXMLData = $null
			
			# Add XML data to Message if configured in EventRules.json
			if($KnownEventRuleObject.IncludeEventXMLDataInMessage -eq $True) {
				
				# DEBUG
				#Write-Host "DEBUG: `$KnownEventRuleObject.IncludeEventXMLDataInMessage is True " -ForegroundColor Yellow
				
				# Get XML data to Multiline string
				$EventObjectXMLDataString = $null
				$EventObject.Properties.Value | Foreach-Object { $EventObjectXMLDataString += "$_"}

				# Add XML data to Event Message
				$EventObject.Message = "$($EventObject.Message)`n`n$EventObjectXMLDataString"
			}

			# Add XML data to ToolTipText if configured in EventRules.json
			if($KnownEventRuleObject.IncludeEventXMLDataInToolTip -eq $True) {

				#Write-Host "DEBUG: IncludeEventXMLDataInToolTip -eq True" -ForegroundColor Yellow
				
				# Add ToolTipText property to Event object
				$EventObject | Add-Member -MemberType noteProperty -Name ToolTipText -Value ''
				
				# We may have $EventObjectXMLData from earlier step above
				if(-not $EventObjectXMLData) {
					# Get XML data to Multiline string
					$EventObject.Properties.Value | Foreach-Object { $EventObjectXMLData = "$XMLData`n$_"}
				}

				$EventObject.ToolTipText = "$($EventObject.ToolTipText)`n`n$EventObjectXMLData"
			}

			# Add Color-information if specified in Known Events
			if($KnownEventRuleObject.Color) {
				
				# Add Color property to Event object
				$EventObject | Add-Member -MemberType noteProperty -Name Color -Value "$($KnownEventRuleObject.Color)"
			}



			# DEBUG Known $EventObject
			#$EventObject
			#Pause
			
		} else {
			# No Known Events for this Event Id
			#Write-Host "No Known Events for $($EventObject.Id)"

<#
			# Add Color-information if LevelDisplayName contains Warning
			if($EventObject.LevelDisplayName -Like '*Warning') {
				
				# Add Color property to Event object
				$EventObject | Add-Member -MemberType noteProperty -Name Color -Value 'Yellow'
			}


			# Add Color-information if LevelDisplayName contains Error
			if($EventObject.LevelDisplayName -Like '*Error') {
				
				# Add Color property to Event object
				$EventObject | Add-Member -MemberType noteProperty -Name Color -Value 'Red'
			}
#>
			if(-not $AllEvents) {
				# Selected only include Known Events
				
				#Write-Host "DEBUG: Unknown event $($EventObject.Id) $($EventObject.ProviderName)" -ForegroundColor Yellow
			
				# Add .RemoveMe = $True property
				# so we will remove this event from array in next step
				$EventObject | Add-Member -MemberType noteProperty -Name RemoveMe -Value $true

			}
		}
	}


	#############################################################
	# Explicit exception handling for Event Logs

	# Intune MDM logs have also falsepositive Error messages which we change from Error to Information
	if($EventLogName -eq 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin') {
		$PossibleFalsePositiveMDMErrors = $GatheredWinEvents | Where-Object { ($_.ProviderName -eq 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider') -and ($_.Id -eq '404') }
		
		Foreach($EventObject in $PossibleFalsePositiveMDMErrors) {

			if($EventObject.Message -like "MDM ConfigurationManager: Command failure status*(./Device/Vendor/MSFT/Policy/ConfigOperations/ADMXInstall/Receiver/Properties/Policy/FakePolicy/Version)*") {
				$EventObject.LevelDisplayName = 'Known Information'
				$EventObject.Color = $null
			}

			# Remove this fake error policy from report is only Known events are shown
			# This is really annoying falseposive event
			if(-not $AllEvents) {
				# Add .RemoveMe = $True property
				# so we will remove this event from array in next step
				$EventObject | Add-Member -MemberType noteProperty -Name RemoveMe -Value $true
			}

		}
	}
	
	# Change MonitorSettingChange Powerevent to human understandable text (Monitor Turned Off and Monitor Turned On)
	if($EventLogName -eq 'Microsoft-Windows-PushNotification-Platform/Operational') {
		$MonitorOnOffMessageEventObjects = $GatheredWinEvents | Where-Object { ($_.ProviderName -eq 'Microsoft-Windows-PushNotifications-Platform') -and ($_.Id -eq '1025') }
		
		Foreach($EventObject in $MonitorOnOffMessageEventObjects) {
			if($EventObject.Message -eq "A Power event was fired: MonitorSettingChange [PowerEventType] true [Enabled].") {
				$EventObject.Message = 'Monitor Turned On'
			}
			
			if($EventObject.Message -eq "A Power event was fired: MonitorSettingChange [PowerEventType] false [Enabled].") {
				$EventObject.Message = 'Monitor Turned Off'
			}
		}
	}

<#
	# Not working - FIX this. Yes you! :)
	if((($EventObject.Id -eq 19) -and ($EventLogName -eq 'WindowsUpdateClient')) -or (($EventObject.Id -eq 43) -and ($EventLogName -eq 'WindowsUpdateClient'))) {

		# Remove newlines from Updates to make report more readable
		$EventObject.Message = ($EventObject.Message) -replace "`r`n|`n|`r", " "		
		#$EventObject.Message = ($EventObject.Message).Replace("`r`n", " ").Replace("`n", " ").Replace("`r", " ")
	}
#>

	
	#############################################################


	# Remove property .RemoveMe $True objects
	# These were unknown events
	$FilteredWinEvents =  $GatheredWinEvents | Where-Object RemoveMe -ne $True
	
	# Sort events
	$FilteredWinEvents =  $FilteredWinEvents | Sort-Object -Property TimeCreated | Select-Object -Property * -ExcludeProperty Bookmark
	
	# DEBUG Save sorted Error events to json file for debugging
	#$DebugSaveFileName = $EventLogName.Replace('/','_')
	#$FilteredWinEvents | ConvertTo-Json -depth 3 | Out-File -Filepath "$PSScriptRoot\_$($DebugSaveFileName).json" -Encoding UTF8 -Force

	Write-Host "Events added to report after filtering $($FilteredWinEvents.Count)"

	# Go through events, process information and save event to timeline
	Foreach ($event in $FilteredWinEvents) {
		
		$LogEntryDateTimeObject = $event.TimeCreated
		$DateTimeToLogFile = "{0:0000}-{1:00}-{2:00} {3:00}:{4:00}:{5:00}.{6:000}" -f $LogEntryDateTimeObject.Year, $LogEntryDateTimeObject.Month, $LogEntryDateTimeObject.Day, $LogEntryDateTimeObject.Hour, $LogEntryDateTimeObject.Minute, $LogEntryDateTimeObject.Second, $LogEntryDateTimeObject.Millisecond
		
		# Remove Microsoft-Windows- from ProviderName so it looks same as Source column in Event Viewer
		$ProviderName = ($event.ProviderName).Replace('Microsoft-Windows-','')

		# Windows Events logName
		$LogName = $event.LogName
		
		if($event.color) {
			$Color = $event.color
		} else {
			$Color = $null
		}
		
		RecordStatusToTimeline -dateTimeObject $LogEntryDateTimeObject -date $DateTimeToLogFile -logName $LogName -providerName $ProviderName -id $event.Id -levelDisplayName $event.levelDisplayName -message $event.message -messageToolTip $event.ToolTipText -KnownCategoryName $event.KnownCategoryName -Color $Color
		
		# DEBUG
		#Write-Host "$($event.TimeCreated) $($event.message)"
		
	}
	Write-Host ""

} # endregion function_read_eventlog




# endregion Functions
################ Functions End ################


###############################################
# Main starts here

Write-Host "Get-WindowsTroubleshootingReportCommunity.ps1 $ScriptVersion" -ForegroundColor Cyan
Write-Host "Author: Petri.Paavola@yodamiitti.fi / Microsoft MVP - Windows and Intune"
Write-Host "Co-author: IT Community"
Write-Host ""

Write-Host "Gathering Windows and log events may take some time, hold on..."
#$GatheredWinEvents = Get-WinEvent -FilterHashtable @{ LogName='*'; StartTime=(Get-Date).AddDays(-$Days) }

# Set events gathering start time based on the specific parameter set
switch ($PsCmdlet.ParameterSetName) {
    'LastMinutes' {
		Write-Host "Getting events from $LastMinutes minutes ago"
		$StartTimeObject = (Get-Date).AddMinutes(-$LastMinutes)

		# Set EndTime to current dateTime + 24hours
		$EndTime = Get-Date
		#$EndTimeObject = (Get-Date).AddHours(24)
		$EndTimeObject = Get-Date

    }
    'LastHours' {
		Write-Host "Getting events from $LastHours hours ago"
		$StartTimeObject = (Get-Date).AddHours(-$LastHours)
		
		# Set EndTime to current dateTime
		$EndTime = Get-Date
		#$EndTimeObject = (Get-Date).AddHours(24)
		$EndTimeObject = Get-Date
    }
    'LastDays' {
		Write-Host "Getting events from $LastDays days ago"
		$StartTimeObject = (Get-Date).AddDays(-$LastDays)
	
		# Set EndTime to current dateTime
		$EndTime = Get-Date
		#$EndTimeObject = (Get-Date).AddHours(24)
		$EndTimeObject = Get-Date
    }
    'MinutesSinceLastBoot' {
        #Write-Host "Retrieving events from the last $LastXDays days."
		
		# Get LastBootTime
		$LastBootTimeObject = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime

		if($MinutesBeforeLastBoot) {
			# Start X minutes before boot
			$StartTimeObject = (Get-Date $LastBootTimeObject).AddMinutes(-$MinutesBeforeLastBoot)
		} else {
			# Start from boot
			$StartTimeObject = Get-Date $LastBootTimeObject
		}

		$EndTimeObject = $LastBootTimeObject.AddMinutes($MinutesAfterLastBoot)

		Write-Host "Gettings events $MinutesSinceLastBoot minutes onwards from last reboot ('$StartTime' - '$EndTime')"
		
    }
	'StartEndTimeSpecified' {

		# Try to create $StartTimeObject
		Try {
			$StartTimeObject = Get-Date $StartTime
		} catch {
			Write-Host "ERROR: Could not parse -StartTime variable to DateTime" -ForegroundColor 'Red'
			Write-Host "Script will exit!"
			Exit 0
		}

		# Try to create $EndTimeObject
		Try {
			$EndTimeObject = Get-Date $EndTime
		} catch {
			Write-Host "ERROR: Could not parse -EndTime variable to DateTime" -ForegroundColor 'Red'
			Write-Host "Script will exit!"
			Exit 0
		}
	
	}	
    'Default' {
		if($AllEvents) {
			# Set StartTime default value to 10 minutes ago
			Write-Host "StartTime not specified so getting events from 10 minutes ago (default)"
			$StartTimeObject = (Get-Date).AddMinutes(-10)
		} else {
			# Set StartTime default value to 2 hours ago
			Write-Host "StartTime not specified so getting only known events from 2 hours ago (default)"
			$StartTimeObject = (Get-Date).AddHours(-2)
		}

		
		# Set EndTime to current dateTime
		$EndTime = Get-Date
		#$EndTimeObject = (Get-Date).AddHours(24)
		$EndTimeObject = Get-Date
    }
}

# Create StartTime and EndTime strings
$StartTimeString = Get-Date $StartTimeObject -Format 'yyyy-MM-dd HH:mm:ss.fff'
$EndTimeString = Get-Date $EndTimeObject -Format 'yyyy-MM-dd HH:mm:ss.fff'

Write-Host "StartTime = $StartTimeString"
Write-Host "EndTime   = $EndTimeString"

# DEBUG - This was to check possible timezone affect for Time
# Which is not affecting, possibly, maybe. Because it actually is affecting but not here...
#Write-Host "DEBUG `$StartTimeObject"
#$StartTimeObject
#Write-Host "DEBUG `$EndTimeObject"
#$EndTimeObject


$EventRulesArray = @()
if(Test-Path "$PSScriptRoot\EventRules") {
	
	# Get EventRules files
	$EventRulesFiles = Get-ChildItem "$PSScriptRoot\EventRules\EventRules*.json" -File

	if($EventRulesFiles) {
		Foreach($File in $EventRulesFiles) {
			Write-Host "Import $($File.Name) file"
			$EventRulesArrayTEMP = Get-Content $File.FullName -Raw | ConvertFrom-Json
			if($EventRulesArrayTEMP) {
				Write-Host "SUCCESS" -ForegroundColor 'Green'
				Write-Host "Imported $($EventRulesArrayTEMP.Count) rulesets ($(($EventRulesArrayTEMP.KnownEventRules).Count) rules)"
				$EventRulesArray += $EventRulesArrayTEMP
			} else {
				Write-Host "FAILED to import $($File.Name)! Skipping to next file..." -ForegroundColor 'Red'
				Pause
			}
		}
	} else {
		Write-Host "EventRules files does not exist.`nMake sure you have downloaded at least EventRules.json from GitHub project page" -ForegroundColor Yellow
		Write-Host "Tool will not recognize known events without this/these EventRules files"
		Exit 1	
	}
} else {
	Write-Host "EventRules folder does not exist.`nMake sure you have downloaded at least EventRules.json from GitHub project page" -ForegroundColor Yellow
	Write-Host "Tool will not recognize known events without this/these EventRules files"
	Exit 1
}

if(-not $EventRulesArray) {
	Write-Host "Could not find or import EventRules .json files" -ForegroundColor Red
	Write-Host "Scripts brains are in these files. Make sure you download and copy files to folder:`n$PSScriptRoot\EventRules" -ForegroundColor Yellow
	Exit 1
}


# Filter out others than specified KnownRules Categories
if($IncludeSelectedKnownRulesCategoriesOnly) {
	
	$TempEventRulesArray = @()
	Foreach($SpecifiedCategoryNameToInclude in $IncludeSelectedKnownRulesCategoriesOnly) {
		# We sholud be ok to use += with array in this case
		# because we should not have THAT many objects to handle
		# "Real" solution would be using ArrayList and .Add operation
		
		$TempEventRulesArray += $EventRulesArray | Where-Object { $_.CategoryName -eq $SpecifiedCategoryNameToInclude }
		
	}
	$EventRulesArray = $TempEventRulesArray
}


# Filter out specified KnownRules Categories to exclude
if ($ExcludeSelectedKnownRulesCategories) {
    
    # Initialize an empty array for filtered results
    $TempEventRulesArray = @()

    # Loop through all categories and exclude the ones in the $ExcludeSelectedKnownRulesCategories
    foreach ($EventRule in $EventRulesArray) {
        # Check if the category is NOT in the exclusion list
        if ($ExcludeSelectedKnownRulesCategories -notcontains $EventRule.CategoryName) {
            # Add to temp array if it's not in the exclusion list
            $TempEventRulesArray += $EventRule
        }
    }

    # Assign the filtered results back to the main array
    $EventRulesArray = $TempEventRulesArray
}

# DEBUG: Print filtered results to verify
#$EventRulesArray | ForEach-Object { Write-Host "Category after exclusion: $($_.CategoryName)" }
# Pause


Write-Host ""

if($RealtimeLogViewerUI) {
	# Realtime log Viewer
	# Starting events from 15 seconds ago and keep updating events every 15 seconds

	Write-Host "Parameter -RealtimeLogViewerUI specified."
	Write-Host "Gettings events from last 15 seconds and keep updating events every 15 seconds"

	try {

		# Start Out-Gridview
		$Script:OutGridViewRealtime = {Out-Gridview}.GetSteppablePipeline()
		$Script:OutGridViewRealtime.Begin($true)

		$StartTimeObject=(Get-Date).AddSeconds(-15)

		While($Script:OutGridViewRealtime) {
			
			if($LogFile -or $LogFilesFolder) {
				# Realtime view from selected 1 log file

				# Not implemented yet
				Write-Host "Realtime view is for now supported only for Windows Events." -ForegroundColor Yellow
				Write-Host "Do not specify parameter -LogFile" -ForegroundColor Yellow
				Exit 0
				
			} else {		
				# "Realtime" from Event Logs

				# Clear variable just in case
				$GatheredWinEvents = $null

				# Get Windows Events
				$GatheredWinEvents = Get-WinEvent -FilterHashtable @{ LogName='*'; StartTime=$StartTimeObject }
				#$GatheredWinEvents = Get-WinEvent -FilterHashtable @{ LogName='*'; StartTime=$StartTime; EndTime=$EndTime }	
				Write-Host "Found $($GatheredWinEvents.Count) events"

				Write-Host "Processing found events. This may take another while..."

				# Filter events to Errors only and save sorted Error events to variable
				#$FilteredWinEvents =  $GatheredWinEvents | Where-Object -Property LevelDisplayName -eq 'Error' | Sort-Object -Property TimeCreated | Select-Object -Property * -ExcludeProperty Bookmark

				# Clear variable just in case
				$FilteredWinEvents = $null

				# Sort events
				$FilteredWinEvents =  $GatheredWinEvents | Sort-Object -Property TimeCreated | Select-Object -Property * -ExcludeProperty Bookmark

				# DEBUG Save sorted Error events to json file for debugging
				#$FilteredWinEvents | ConvertTo-Json -depth 3 | Out-File -Filepath "$PSScriptRoot\GatheredWinEvents.json" -Encoding UTF8

				# Go through events, process information and save event to timeline
				Foreach ($event in $FilteredWinEvents) {
					RecordStatusToTimeline -date $event.TimeCreated -logName $event.LogName -id $event.Id -levelDisplayName $event.levelDisplayName -message $event.message
				}

				# Get last event
				$LastEvent = $FilteredWinEvents | Select-Object -Last 1

				# Clear variable just in case
				$FilteredWinEvents = $null

				# Set StartTime to 1 microseconds onwards from last event log time
				$StartTime=(Get-Date $LastEvent.TimeCreated).AddMilliseconds(1)
				
				Write-Host "Sleeping 15 seconds"
				Start-Sleep -Seconds 15
				Write-Host "Press Ctrl-C to break realtime event logging..."
			}
		}
	} catch {
		# This is not working because Ctrl-C breaks out from whole script :/

		# User pressed Ctrl-C to break out from While Loop
		Write-Host "Realtime view canceled with Ctrl-C..."
		
	} finally {
		# Let's do some cleanup
		$Script:OutGridViewRealtime.End()		
		
	}

} else {
	# Not realtime log Viewer
	# This is normal case

	if($LogFilesFolder -or $LogFile) {
		# Getting events and logs from specified file or folder

		# We'll change this when we find name information
		$ComputerNameForReport = 'unknown'

		if($LogFile) {
			# -LogFile
			
			Write-Host "Parameter -LogFile specified"
			Write-Host "Loading events or log entries from file:`n$LogFile`n"
			
			if(-not (Test-Path $LogFile)) {
				Write-Host "File does NOT exist: $LogFile" -ForegroundColor Red
				Write-Host "Script will exit"
				Exit 0
			}
			
			$LogFileObjects = Get-ChildItem -Path $LogFile
			
		} else {
			# -LogFilesFolder
			
			Write-Host "Parameter -LogFilesFolder specified"
			Write-Host "Loading events and log files from folder:`n$LogFilesFolder`n"

			if(-not (Test-Path $LogFilesFolder)) {
				Write-Host "Folder does NOT exist: $LogFilesFolder" -ForegroundColor Red
				Write-Host "Script will exit"
				Exit 0
			}
					
			# PRODUCTION line
			$LogFileObjects = Get-ChildItem -Path "$LogFilesFolder\*" -Include *.evtx,*.log -Depth 2
			
			# Not used if .etl files are needed. They lack information when using Get-WinEvent
			# .etl files need to be processed other way
			#$LogFileObjects = Get-ChildItem -Path "$LogFilesFolder\*" -Include *.evtx,*.etl,*.log -Depth 2
			
			# Sort by file Extension
			$LogFileObjects = $LogFileObjects | Sort-Object -Property Extension
			
		}
		

		Foreach($FileObject in $LogFileObjects) {

			Write-Host "Processing file: $($FileObject.Name)"

			if(($FileObject.extension -eq '.evtx') -or ($FileObject.extension -eq '.etl')) {
				# Process saved EventLog file
				Read-EventLog -EventLogPath $FileObject.FullName -StartTimeObject $StartTimeObject -EndTimeObject $EndTimeObject
				
			} elseif ($FileObject.extension -eq '.log') {
				# Read and parse .log files
				Write-Host "Processing log file: $($FileObject.Name)"
				Read-LogFile -Path $FileObject.FullName -StartTimeObject $StartTimeObject -EndTimeObject $EndTimeObject
				Write-Host ""
				
			} else {
				# Unknown file type
				# We should never get here if we filter our file search earlier
				Write-Host "ERROR: File type not supported ($($FileObject.Extension)). Skipping to next file..." -ForegroundColor 'Red'
				
				# Continue to next file in Foreach loop
				Continue
			}
		}
		
		
	} else {
		# Getting events and logs from live Windows system

		$ComputerNameForReport = $env:ComputerName

		########################################################
		# Process Event logs from live Windows

		Write-Host "Get Windows Events Logs"

		# Get EventLogs which have Events
		#$EventLogsList = Get-WinEvent -ListLog * | Where-Object RecordCount -gt 0
		$EventLogsList = Get-WinEvent -ListLog * | Where-Object RecordCount -gt 0 -ErrorAction SilentlyContinue

		# Filter in EventLogs where LastWriteTime property is later than $StartTimeObject
		$EventLogsList = $EventLogsList | Where-object LastWriteTime -ge $StartTimeObject
		
		Write-Host "Processing $($EventLogsList.Count) Event logs"
		
		# Production 0.5 command. Restore me back
		#$GatheredWinEvents = Get-WinEvent -FilterHashtable @{ LogName='*'; StartTime=$StartTimeObject; EndTime=$EndTimeObject }

		if($AllEvents) {
			Write-Host "Gettings all events (this creates bigger file size)"
			Write-Host ""
		} else {
			Write-Host "Gettings only known events in Eventrules.json (this creates smaller file size)"	
			Write-Host ""
		}
		
		Foreach($EventLogObject in $EventLogsList) {
			Read-EventLog -EventLog $EventLogObject -StartTimeObject $StartTimeObject -EndTimeObject $EndTimeObject
		}


		########################################################
		# Process log files from live Windows

		# Log folders to process from Windows
		$LogFoldersFromWindows = @(
			'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs'
			'C:\Windows\CCM\Logs'
			'C:\Windows\CCMSetup\Logs'
			'C:\Windows\Logs\CBS\CBS.log'
			"C:\Windows\Logs\DISM\dism.log"
		)

		Write-Host "Get .log files"

		# PRODUCTION line
		#
		# -ErrorAction SilentlyContinue
		# This is specified because non-existing path will produce error.
		$LogFileObjects = Get-ChildItem -Path $LogFoldersFromWindows -Include *.log -Depth 2 -ErrorAction SilentlyContinue

		Write-Host "Found $($LogFileObjects.Count) .log files"
		
		# Sort by file Extension
		#$LogFileObjects = $LogFileObjects | Sort-Object -Property Extension

		Foreach($FileObject in $LogFileObjects) {

			Write-Host "Processing file: $($FileObject.Name)"

			if(($FileObject.extension -eq '.evtx') -or ($FileObject.extension -eq '.etl')) {
				# Process saved EventLog file
				Read-EventLog -EventLogPath $FileObject.FullName -StartTimeObject $StartTimeObject -EndTimeObject $EndTimeObject
				
			} elseif ($FileObject.extension -eq '.log') {
				# Read and parse .log files
				Write-Host "Processing log file: $($FileObject.Name)"
				Read-LogFile -Path $FileObject.FullName -StartTimeObject $StartTimeObject -EndTimeObject $EndTimeObject
				Write-Host ""
				
			} else {
				# Unknown file type
				# We should never get here if we filter our file search earlier
				Write-Host "ERROR: File type not supported ($($FileObject.Extension)). Skipping to next file..." -ForegroundColor 'Red'
				
				# Continue to next file in Foreach loop
				Continue
			}
		}


		########################################################
		# Get ScheduledTask information

		# Check if Scheduled Tasks CategoryName exists $EventRulesArray
		if($EventRulesArray | Where categoryname -eq 'Scheduled Tasks') {

			$ScheduledTasks = Get-ScheduledTask | Get-ScheduledTaskInfo
			
			# Include Tasks that did run in specified timeframe
			$ScheduledTasksInSpecifiedTimeScope = $ScheduledTasks | Where-Object { ($_.LastRunTime -ge $StartTimeObject) -and ($_.LastRunTime -le $EndTimeObject) }

			# Process Task that run in specified timeframe
			Foreach($ScheduledTaskObject in $ScheduledTasksInSpecifiedTimeScope) {
				$LogEntryDateTimeObject = $ScheduledTaskObject.LastRunTime
				
				$DateTimeToLogFile = "{0:0000}-{1:00}-{2:00} {3:00}:{4:00}:{5:00}.{6:000}" -f $LogEntryDateTimeObject.Year, $LogEntryDateTimeObject.Month, $LogEntryDateTimeObject.Day, $LogEntryDateTimeObject.Hour, $LogEntryDateTimeObject.Minute, $LogEntryDateTimeObject.Second, $LogEntryDateTimeObject.Millisecond
				
				$LastTaskResult = $ScheduledTaskObject.LastTaskResult
				$TaskName = $ScheduledTaskObject.TaskName
				$TaskPath = $ScheduledTaskObject.TaskPath

				#$Message = "ScheduledTask Run: $($TaskName) Result:$($LastTaskResult) Path:$($TaskPath)"
				$Message = "ScheduledTask Run: $($TaskName) ($($LastTaskResult))"
				$ToolTip = "Path:$($TaskPath)"

				$LogName = 'Scheduled Task'
				$ProviderName = 'Scheduled Task'
				$EventId = $null
				


				if($LastTaskResult -eq 0) {
					# Last TaskResult succeeded
					$Color = 'Green'
					$LevelDisplayName = 'Scheduled Task Success'
				} else {
					# Last TaskResult failed
					$Color = 'Red'
					$LevelDisplayName = 'Scheduled Task Failed'
				}
				
				RecordStatusToTimeline -dateTimeObject $LogEntryDateTimeObject -date $DateTimeToLogFile -logName $LogName -providerName $ProviderName -id $EventId -levelDisplayName $LevelDisplayName -message $Message -Color $Color -MessageToolTip $ToolTip -KnownCategoryName 'Scheduled Tasks'
			}

			# Include Tasks in the future
			$ScheduledTasksInTheFuture = $ScheduledTasks | Where-Object { $_.NextRunTime -ge $EndTimeObject }

			# Process Task that run in specified timeframe
			Foreach($ScheduledTaskObject in $ScheduledTasksInTheFuture) {
				$LogEntryDateTimeObject = $ScheduledTaskObject.NextRunTime
				
				$DateTimeToLogFile = "{0:0000}-{1:00}-{2:00} {3:00}:{4:00}:{5:00}.{6:000}" -f $LogEntryDateTimeObject.Year, $LogEntryDateTimeObject.Month, $LogEntryDateTimeObject.Day, $LogEntryDateTimeObject.Hour, $LogEntryDateTimeObject.Minute, $LogEntryDateTimeObject.Second, $LogEntryDateTimeObject.Millisecond
				
				$LastTaskResult = $null
				$TaskName = $ScheduledTaskObject.TaskName
				$TaskPath = $ScheduledTaskObject.TaskPath

				#$Message = "Future ScheduledTask run:$($TaskName) Path:$($TaskPath)"
				$Message = "Future ScheduledTask run:$($TaskName)"
				$ToolTip = "Path:$($TaskPath)"

				$LogName = 'Scheduled Task Future'
				$ProviderName = 'Scheduled Task Future'
				$EventId = $null

				$LevelDisplayName = 'Future Task'
				#$Color = 'White'
				
				RecordStatusToTimeline -dateTimeObject $LogEntryDateTimeObject -date $DateTimeToLogFile -logName $LogName -providerName $ProviderName -id $EventId -levelDisplayName $LevelDisplayName -message $Message -Color $null -MessageToolTip $ToolTip -KnownCategoryName 'Scheduled Tasks'
			}
		}

	} # End Getting events and logs from live Windows system
}

	#if($ExportHTML) {
	if($True) {


		########################################################################################################################
		# Create HTML report

	$head = @'
	<style>
		body {
			background-color: #FFFFFF;
			font-family: Arial, sans-serif;
		}

		  header {
			background-color: #444;
			color: white;
			padding: 10px;
			display: flex;
			align-items: center;
			
			position: sticky; /* Makes the header stick to the top */
			top: 0; /* Sticks the header at the top of the page */
			z-index: 1000; /* Ensures the header stays above other content */
			padding: 10px;
			box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); /* Adds a shadow for visual effect */
		  }

		  header h1 {
			margin: 0;
			font-size: 24px;
			margin-right: 20px;
		  }

		  header .additional-info {
			display: flex;
			flex-direction: column;
			align-items: flex-start;
			justify-content: center;
		  }

		  header .additional-info p {
			margin: 0;
			line-height: 1.2;
		  }

		  header .start-end-time-group {
			display: flex;
			flex-direction: column;
			align-items: flex-start;
			justify-content: center;
			margin-left: 20px; /* Adds a bit of space between the two groups */
		  }

		  header .start-end-time-group p {
			margin: 0;
			line-height: 1.2;
		  }

		  header .author-info {
			display: flex;
			flex-direction: column;
			align-items: flex-end;
			justify-content: center;
			margin-left: auto;
		  }

		  header .author-info p {
			margin: 0;
			line-height: 1.2;
		  }
		  
		  header .author-info a {
			color: white;
			text-decoration: none;
		  }

		  header .author-info a:hover {
			text-decoration: underline;
		  }

		  .filtering {
			background-color: white;
			color: black;
			padding: 10px;
			display: flex;
			/* align-items: center; */
			
			position: sticky; /* Makes the header stick to the top */
			top: 0; /* Sticks the header at the top of the page */
			/* background-color: white; *//* Ensures the background of the header is opaque */
			z-index: 1000; /* Ensures the header stays above other content */
			padding: 10px;
			box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); /* Adds a shadow for visual effect */
		  }

		.search-container {
			background-color: white;
			padding: 5px;
			display: flex;
			align-items: center;
			position: sticky; /* Sticky behavior */
			top: 180px; /* Adjust based on your layout, ensures it's below the dropdowns */
			z-index: 1000;
			box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
		}

		#searchInput {
			padding: 3px;
			width: 200px;
		}

		#clearSearch, #resetFilter {
			margin-left: 5px;
			padding: 3px;
		}

		.filter-row {
		  display: flex;
		  gap: 20px;
		  margin-top: 10px;
		}

		.control-group {
		  margin: 0 10px;
		}

		table {
			border-collapse: collapse;
			width: 100%;
			text-align: left;
		}

		table, table#TopTable {
			border: 2px solid #1C6EA4;
			/* background-color: #f7f7f4; */
			background-color: #012456; /* dark background */
			color: #f0f0f0;  /* light color for the text */
		}

		table td, table th {
			/* border: 2px solid #AAAAAA; */
			border: 1px solid #1a2f6a;
			padding: 2px;
		}

		table td {
			font-size: 15px;
			white-space: pre; /* or 'nowrap' of 'pre' if you don't want any wrapping. 'pre-wrap' will wrap */
		}

		table th {
			font-size: 18px;
			font-weight: bold;
			color: #FFFFFF;
			background: #1C6EA4;
			background: -moz-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
			background: -webkit-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
			background: linear-gradient(to bottom, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
			white-space: pre; /* or 'nowrap' of 'pre' if you don't want any wrapping. 'pre-wrap' will wrap */white-space: nowrap;
		}


		/* Highlight table line when hovering over */
		table tr:hover {
			background-color: #1C6EA4;
            cursor: pointer;
		}

		table#TopTable td, table#TopTable th {
			vertical-align: top;
			text-align: center;
		}

		table thead th:first-child {
			border-left: none;
		}

		table thead th span {
			font-size: 14px;
			margin-left: 4px;
			opacity: 0.7;
		}

		table tfoot {
			font-size: 16px;
			font-weight: bold;
			color: #FFFFFF;
			background: #D0E4F5;
			background: -moz-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
			background: -webkit-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
			background: linear-gradient(to bottom, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
			border-top: 2px solid #444444;
		}

		table tfoot .links {
			text-align: right;
		}

		table tfoot .links a {
			display: inline-block;
			background: #1C6EA4;
			color: #FFFFFF;
			padding: 2px 8px;
			border-radius: 5px;
		}

		/* Every other table row to different background color
		table tbody tr:nth-child(even) {
		  background-color: #D0E4F5;
		}
		*/

		/* Set table row background color */
		table tbody tr {
		  /* background-color: #f7f7f4; */
		  background-color: #012456;
		}

		select {
		  font-family: "Courier New", monospace;
		}

	  /* Set DownloadTable settings */
	  #ApplicationDownloadStatistics {
		width: auto;
		max-width: 1000px;
		margin: 0 auto; /* Centers the table if it doesn't span the entire width of its container */
		
		/* Float -> align table to left
		   This setting needs to be cleared after table with html syntax
		   <div style="clear:both;"></div>
		   because otherwise page bottom element will float on right of table
		 */
		float: left;
	  }

		
	  footer {
		background-color: #444;
		color: white;
		padding: 10px;
		display: flex;
		align-items: center;
		justify-content: center;
	  }


	  footer .creator-info {
		display: flex;
		flex-direction: row;
		align-items: center;
		margin-right: 20px;
	  }

	  footer .creator-info p {
		line-height: 1.2;
		margin: 0;
	  }

	  footer .creator-info p.author-text {
		margin-right: 20px; /* Add margin-right rule here */
	  }

	  .profile-container {
		position: relative;
		width: 50px;
		height: 50px;
		border-radius: 50%;
		overflow: hidden;
		margin-right: 10px;
	  }

	  .profile-container img {
		width: 100%;
		height: 100%;
		object-fit: cover;
		transition: opacity 0.3s;
	  }

	  .profile-container img.black-profile {
		position: absolute;
		top: 0;
		left: 0;
		z-index: 1;
	  }

	  .profile-container:hover img.black-profile {
		opacity: 0;
	  }

	  footer .company-logo {
		width: 100px;
		height: auto;
		margin: 0 20px;
	  }

	  footer a {
		color: white;
		text-decoration: none;
	  }

	  footer a:hover {
		text-decoration: underline;
	  }

	  
		.filter-row {
		  display: flex;
		  align-items: center;
		}

		.control-group {
		  display: flex;
		  flex-direction: column;
		  align-items: flex-start;
		  margin-right: 16px;
		}

		.control-group label {
		  font-weight: bold;
		}

		/* Tooltip container */
		.tooltip {
		  position: relative;
		  display: inline-block;
		  cursor: pointer;
		  /* text-decoration: none; */ /* Remove underline from hyperlink */
		  color: inherit; /* Make the hyperlink have the same color as the text */
		}

		/* Tooltip text */
		.tooltip .tooltiptext {
		  visibility: hidden;
		  
		  /* make it up to the viewport width minus some padding 
		     bigger value is used so reaaaaaally wide tooltips would not overflow from left
			 which it still will do depending on window size
		  */
		  max-width: calc(100vw - 160px); 
		  
		  /* width: 120px; */
		  background-color: #555;
		  color: #fff;
		  text-align: left;
		  border-radius: 6px;
		  position: absolute;
		  z-index: 1;
		  bottom: 125%;

		  /* ensures the tooltip is centered with respect to the hovered element */
		  /* These are important settings! */
		  transform: translateX(-50%);
		  left: 50%;

		  /* margin-left: -60px; */
		  opacity: 0;
		  transition: opacity 1s;
		  /* white-space: pre; */
		  padding: 10px; /* Change this value to suit your needs */
		  
		  /* monospaced font so output is more readable */
		  font-family: 'Courier New', monospace;
		  /* word-wrap: break-word; */ /* break long words to fit within the tooltip */
		  
		  overflow-x: auto; /* This adds horizontal scrolling if content exceeds max-width */
		  
		  /* white-space: nowrap; */ /* This ensures content stays in a single line */
		  
		  /* preserve whitespaces */
		  white-space: pre;
		  
		  
		}

		/* Show tooltip text when hovering */
		.tooltip:hover .tooltiptext {
		  visibility: visible;
		  opacity: 1;
		}
	</style>
'@

		############################################################
		# Create HTML report

		Write-Host "Create HTML report information"

		######################
		#Write-Host "Create Observed	Timeline HTML fragment."

		try {

			$PreContent = @"
				<div class="filtering">
					<div class="filter-row">
						<div class="control-group">
						  <label><input type="checkbox" class="filterCheckbox" value="Green" onclick="toggleCheckboxes(this)"> Green</label>
						  <label><input type="checkbox" class="filterCheckbox" value="Yellow" onclick="toggleCheckboxes(this)"> Yellow</label>
						  <label><input type="checkbox" class="filterCheckbox" value="Red" onclick="toggleCheckboxes(this)"> Red</label>
						</div>
						<!-- Dropdown 1 -->
						<div class="control-group">
							<label for="dropdown1">LogName</label>
							<select id="filterDropdown1" multiple size="8">
							  <option value="all" selected>All</option>
							</select>
						</div>
						<!-- Dropdown 2 -->
						<div class="control-group">
							<label for="dropdown2">ProviderName</label>
							<select id="filterDropdown2" multiple size="8">
							  <option value="all" selected>All</option>
							</select>
						</div>
						<!-- Dropdown 3 -->
						<div class="control-group">
						<label for="dropdown3">LevelDisplayName</label>
						<select id="filterDropdown3" multiple size="8">
							  <option value="all" selected>All</option>
							</select>
						</div>
						<!-- Dropdown 4 -->
						<div class="control-group">
						<label for="dropdown4">KnownCategoryName</label>
						<select id="filterDropdown4" multiple size="8">
							  <option value="all" selected>All</option>
							</select>
						</div>
					</div>
				</div>
				<div class="search-container">
					<input type="text" id="searchInput" placeholder="Search...">
					<button id="clearSearch" onclick="clearSearch()">X</button>
					<button id="resetFilter" onclick="resetFilters()">Reset filters</button>
				</div>
			</div>
"@
			
			# Sort $observedTimeline
			
			#Write-Verbose "DEBUG: Objects in `$observedTimeline BEFORE sorting: $($observedTimeline.Count)"
			
			try {
				
				# Sort descending to get newest entries to topmost
				# This is same behavior than in Event Viewer
				# Though log files have newest in bottom
				if($SortDescending) {
					# Sort Descending
					# This is like Event logs where newest is topmost
					$observedTimeline = $observedTimeline | Sort-Object -Property @{
						Expression = { $_.Date }
						Ascending = $false # Sort Date in descending order
					}, @{
						Expression = { $_.Index }
						Ascending = $false  # Sort Index in descending order
					}
				} else {

					# Sort oldest first and newest last
					# Sort Ascending
					# This is like log files where oldest is topmost
					$observedTimeline = $observedTimeline | Sort-Object -Property @{
						Expression = { $_.Date }
						Ascending = $true # Sort Date in descending order
					}, @{
						Expression = { $_.Index }
						Ascending = $true  # Sort Index in descending order
					}
				}
				
			} catch {

				Write-Host "Sorting by property dateTimeObject failed" -ForegroundColor 'Yellow'
			}

			#Write-Verbose "DEBUG: Objects in `$observedTimeline AFTER  sorting: $($observedTimeline.Count)"

			# Create HTML from observedTimeline
			$observedTimelineHTML = $observedTimeline | Select-Object -Property * -ExcludeProperty dateTimeObject, EventNumber | ConvertTo-Html -As Table -Fragment -PreContent $PreContent

			# Fix &lt; &quot; etc...
			#$observedTimelineHTML = Fix-HTMLSyntax $observedTimelineHTML

			# Fix column names
			#$AllAppsByDisplayNameHTML = Fix-HTMLColumns $AllAppsByDisplayNameHTML

			# Add TableId
			$TableId = 'ObservedTimeline'
			$observedTimelineHTML = $observedTimelineHTML.Replace('<table>',"<table id=`"$TableId`">")

			# Convert HTML Array to String which is requirement for HTTP PostContent
			$observedTimelineHTML = $observedTimelineHTML | Out-String


			# DEBUG save $html1 to file
			#$observedTimelineHTML | Out-File "$PSScriptRoot\observedTimelineHTML.html"

		}
		catch {
			Write-Error "$($_.Exception.GetType().FullName)"
			Write-Error "$($_.Exception.Message)"
			Write-Error "Error creating HTML fragment information"
			Write-Host "Script will exit..."
			Pause
			Exit 1        
		}
		#############################
		# Create html
		Write-Host "Creating HTML report..."

		try {

			$ReportRunDateTime = (Get-Date).ToString("yyyyMMddHHmm")
			$ReportRunDateTimeHumanReadable = (Get-Date).ToString("yyyy-MM-dd HH:mm")
			$ReportRunDateFileName = (Get-Date).ToString("yyyyMMddHHmm")

			if($ExportHTMLReportPath) {
				if(Test-Path $ExportHTMLReportPath) {
					$ReportSavePath = $ExportHTMLReportPath
				} else {
					Write-Host "Warning: Parameter -ExportHTMLReportPath specified but destination directory does not exists ($ExportHTMLReportPath) !" -ForegroundColor Yellow
					Write-Host "Warning: Defaulting to running directory ($PSScriptRoot)!" -ForegroundColor Yellow
					$ReportSavePath = $PSScriptRoot
				}
			} else {
				$ReportSavePath = $PSScriptRoot
			}
			
			
			if($ComputerNameForReport -eq 'N/A') {
				$HTMLFileName = "$($ReportRunDateFileName)_WindowsTroubleShooting_Report.html"	
			} else {
				$HTMLFileName = "$($ReportRunDateFileName)_$($ComputerNameForReport)_WindowsTroubleShooting_Report.html"	
			}
			
			
			$PreContent = @"
		<header>
		  <h1>Get-WindowsTroubleShootingReportCommunity ver $ScriptVersion</h1>
		  <div class="additional-info">
			<p><strong>Report run:</strong> $ReportRunDateTimeHumanReadable</p>
			<p><strong>Computer Name:</strong> $ComputerNameForReport</p>
		  </div>
		  <div class="start-end-time-group">
			<p><strong>Start time:</strong> $StartTimeString</p>
			<p><strong>End time  :</strong> $EndTimeString</p>
			<!-- <p><strong>Tenant name:</strong> $TenantDisplayName</p> -->
			<!-- <p><strong>Tenant id:</strong> $($ConnectMSGraph.TenantId)</p> -->
		  </div>
		  <div class="author-info">
			<p><a href="https://github.com/petripaavola/Get-WindowsTroubleshootingReportCommunity" target="_blank"><strong>Download Report tool from GitHub</strong></a><br>Author: Petri Paavola - Microsoft MVP</p>
		  </div>
		</header>
		<br>
"@

			$JavascriptPostContent = @'
		<div style="clear:both;"></div>
		<p><br></p>
		<footer>
			<div class="creator-info">
			<p class="author-text">Author:</p>
			  <div class="profile-container">
				<img src="data:image/png;base64,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" alt="Profile Picture Petri Paavola">
				<img class="black-profile" src="data:image/png;base64,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" alt="Black Profile Picture Petri Paavola">
			  </div>
			  <p><strong>Petri Paavola</strong><br>
				<a href="mailto:Petri.Paavola@yodamiitti.fi">Petri.Paavola@yodamiitti.fi</a><br>
				Senior Modern Management Principal
			  </p>
			</div>
			<img class="company-logo" src="data:image/png;base64,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" alt="Microsoft MVP">
			<p style="margin: 0;">
			  <a href="https://github.com/petripaavola/Get-WindowsTroubleshootingReportCommunity" target="_blank"><strong>Download Get-WindowsTroubleshootingReportCommunity from GitHub</strong<</a>
			</p>
		</footer>
		<script>
		
			const HTML_TABLE_ID = "ObservedTimeline";
			const COLUMN_INDEX_FOR_DROPDOWN1 = 2;
			const COLUMN_INDEX_FOR_DROPDOWN2 = 3;
			const COLUMN_INDEX_FOR_DROPDOWN3 = 5;
			const COLUMN_INDEX_FOR_DROPDOWN4 = 8;
			const COLUMN_INDEX_FOR_CHECKBOXES = 9;
			const SORT_BY_COLUMN_INDEX = 0;
			const SORT_AS_INTEGER_COLUMNS = [0];
			const INTENT_COLUMN_INDEX = 5;

			// Specify which columns will be sorted as integers
			const integerColumns = [0];

			// Add column names you want to hide
			const COLUMN_NAMES_TO_HIDE = ['Index', 'Color', 'MessageToolTip', 'KnownCategoryName'];
			
			// Use for testing we have values in hidden columns
			//const COLUMN_NAMES_TO_HIDE = ['Index'];

			// Add column names which are set to bold text
			//const COLUMN_NAMES_TO_BOLD = ['Date', 'Message'];
			const COLUMN_NAMES_TO_BOLD = ['Date'];

			// Change table row background color based on Color column value
			// Green to success, red to fail and Yellow for Info
			const COLOR_COLUMN_NAME = 'Color';
			const NAMED_COLUMNS_TO_COLOR = ['Date', 'Id', 'LevelDisplayName', 'Message'];

	
			// Constant object with predefined hex color values
			
			const FONT_COLOR_CONSTANTS = {
				'Green': '#00ff00',
				'Red': '#ff0000',
				'Yellow': '#ffffcc'
				// Add other colors as needed
			};

	
			function updateRowBackgroundOnColumnValueChange(tableId, columnIndex) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");
			  let previousValue = null;
			  let currentColor = "rgba(208, 228, 245, 1)";

			  table.setAttribute("data-last-color", currentColor);

			  for (let i = 1; i < rows.length; i++) {
				let row = rows[i];

				// Skip hidden rows
				if (row.style.display === "none") {
				  continue;
				}

				let currentValue = row.getElementsByTagName("td")[columnIndex].textContent;

				if (previousValue !== null && currentValue !== previousValue) {
				  currentColor = table.getAttribute("data-last-color") === "rgba(242, 242, 242, 1)" ? "rgba(208, 228, 245, 1)" : "rgba(242, 242, 242, 1)";
				  table.setAttribute("data-last-color", currentColor);
				}

				row.style.backgroundColor = currentColor;
				previousValue = currentValue;
			  }
			}


			function setColumnBold(tableId, columnIndex) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");

			  // Unbold previously selected column
			  if (table.hasAttribute("data-bold-column")) {
				let previousBoldColumn = parseInt(table.getAttribute("data-bold-column"));

				rows[0].getElementsByTagName("th")[previousBoldColumn].style.fontWeight = "normal";
				for (let i = 1; i < rows.length; i++) {
				  rows[i].getElementsByTagName("td")[previousBoldColumn].style.fontWeight = "normal";
				}
			  }

			  // Set header text to bold
			  rows[0].getElementsByTagName("th")[columnIndex].style.fontWeight = "bold";

			  // Set column values to bold
			  for (let i = 1; i < rows.length; i++) {
				rows[i].getElementsByTagName("td")[columnIndex].style.fontWeight = "bold";
			  }

			  // Save current bold column index
			  table.setAttribute("data-bold-column", columnIndex);
			}


			function mergeSort(arr, comparator) {
			  if (arr.length <= 1) {
				return arr;
			  }

			  const mid = Math.floor(arr.length / 2);
			  const left = mergeSort(arr.slice(0, mid), comparator);
			  const right = mergeSort(arr.slice(mid), comparator);

			  return merge(left, right, comparator);
			}

			function merge(left, right, comparator) {
			  let result = [];
			  let i = 0;
			  let j = 0;

			  while (i < left.length && j < right.length) {
				if (comparator(left[i], right[j]) <= 0) {
				  result.push(left[i]);
				  i++;
				} else {
				  result.push(right[j]);
				  j++;
				}
			  }

			  return result.concat(left.slice(i)).concat(right.slice(j));
			}

			// Declare a sortingDirections object outside the sortTable function to store the sorting directions for each column.
			const sortingDirections = {};


			function sortTable(n, tableId, dateColumns = []) {
			  let table, rows;
			  table = document.getElementById(tableId);

			  // Initialize the sorting direction for the column if it hasn't been set yet
			  if (!(n in sortingDirections)) {
				sortingDirections[n] = "asc";
			  }

			  // Remove existing arrow icons
			  let headerRow = table.getElementsByTagName("th");
			  for (let i = 0; i < headerRow.length; i++) {
				headerRow[i].innerHTML = headerRow[i].innerHTML.replace(
				  /<span>.*<\/span>/,
				  "<span>&#8597;</span>"
				);
			  }

			  const isDateColumn = dateColumns.includes(n);
			  rows = Array.from(table.rows).slice(1);

			  const comparator = (a, b) => {
				const x = a.cells[n].innerHTML.toLowerCase();
				const y = b.cells[n].innerHTML.toLowerCase();
				const isIntegerColumn = integerColumns.includes(n);

				if (isDateColumn) {
				  const xDate = getDateFromString(x);
				  const yDate = getDateFromString(y);

				  if (sortingDirections[n] === "asc") {
					return xDate - yDate;
				  } else {
					return yDate - xDate;
				  }
				} else if (isIntegerColumn) {
					const xInt = parseInt(x, 10) || 0;  // Use 0 if parsing fails
					const yInt = parseInt(y, 10) || 0;  // Use 0 if parsing fails
					if (sortingDirections[n] === "asc") {
					  return xInt - yInt;
					} else {
					  return yInt - xInt;
					}
				} else {
				  if (sortingDirections[n] === "asc") {
					return x.localeCompare(y);
				  } else {
					return y.localeCompare(x);
				  }
				}
			  };

			  const sortedRows = mergeSort(rows, comparator);

			  // Reinsert sorted rows into the table
			  for (let i = 0; i < sortedRows.length; i++) {
				table.tBodies[0].appendChild(sortedRows[i]);
			  }

			  // Update arrow icon for the last sorted column
			  if (sortingDirections[n] === "asc") {
				headerRow[n].innerHTML = headerRow[n].innerHTML.replace(
				  /<span>.*<\/span>/,
				  "<span>&#x25B2;</span>"
				);
			  } else {
				headerRow[n].innerHTML = headerRow[n].innerHTML.replace(
				  /<span>.*<\/span>/,
				  "<span>&#x25BC;</span>"
				);
			  }

			  // Create row coloring based on selected column
			  updateRowBackgroundOnColumnValueChange(HTML_TABLE_ID, n);

			  // Bold selected column header and text
			  setColumnBold(HTML_TABLE_ID, n);

			  // Toggle sorting direction for the next click
			  sortingDirections[n] = sortingDirections[n] === "asc" ? "desc" : "asc";
			}


			function getDateFromString(dateStr) {
				let [day, month, year, hours, minutes, seconds] = dateStr.split(/[. :]/);
				return new Date(year, month - 1, day, hours, minutes, seconds);
			}

			// Populates the given select element with unique values from the specified table and column
			function populateSelectWithUniqueColumnValues(tableId, column, selectId) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");
			  let uniqueValues = {};

			  for (let i = 1; i < rows.length; i++) {
				let cellValue = rows[i].getElementsByTagName("td")[column].innerText;

				if (uniqueValues[cellValue]) {
				  uniqueValues[cellValue]++;
				} else {
				  uniqueValues[cellValue] = 1;
				}
			  }

			  let select = document.getElementById(selectId);

			  // Convert the uniqueValues object to an array of key-value pairs
			  let uniqueValuesArray = Object.entries(uniqueValues);

			  // Sort the array by the keys (unique column values)
			  uniqueValuesArray.sort((a, b) => a[0].localeCompare(b[0]));

			  // Find the longest text
			  let longestTextLength = Math.max(...uniqueValuesArray.map(([value, count]) => (value + " (" + count + ")").length));

			  // Loop through the sorted array to create the options with padded number values
			  for (let [value, count] of uniqueValuesArray) {
				let optionText = value + " (" + count + ")";
				let paddingLength = longestTextLength - optionText.length;
				let padding = "\u00A0".repeat(paddingLength);
				let option = document.createElement("option");
				option.value = value;
				option.text = value + padding + " (" + count + ")";
				select.add(option);
			  }
			}

			// This is used to extract AssignmentGroupDisplayName from complex a tag with span
			function getDirectChildTextNotWorking(parentNode) {
				let childNodes = parentNode.childNodes;
				let textContent = '';

				for(let i = 0; i < childNodes.length; i++) {
					if(childNodes[i].nodeType === Node.TEXT_NODE) {
						textContent += childNodes[i].nodeValue;
					}
				}

				return textContent.trim();  // Remove leading/trailing whitespaces
			}


			// Returns the textContent of the node if the node doesn't have any child nodes
			// In some Columns we have a href tag so we in those cases we need to get display value from <a> tag
			function getDirectChildText(node) {
			  if (!node || !node.hasChildNodes()) return node.textContent.trim();
			  return Array.from(node.childNodes)
				.filter(child => child.nodeType === Node.TEXT_NODE)
				.map(textNode => textNode.textContent)
				.join("");
			}


			// Filters the table based on the selected dropdown values, checkboxes, and search input
			// Should find AssignmentGroup and Filter displayNames for filtering from a href tag
			function combinedFilter(tableId, columnIndexForDropdown1, columnIndexForDropdown2, columnIndexForDropdown3, columnIndexForDropdown4, columnIndexForCheckboxes) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");
			  let checkboxes = document.getElementsByClassName("filterCheckbox");
			  let dropdown1 = document.getElementById("filterDropdown1");
			  let dropdown2 = document.getElementById("filterDropdown2");
			  let dropdown3 = document.getElementById("filterDropdown3");
			  let dropdown4 = document.getElementById("filterDropdown4");
			  let searchInput = document.getElementById("searchInput");
			  let searchText = searchInput.value.toLowerCase();

			  // Add dropdown here if you need to support multiple dropdown values filtering
			  let selectedDropdownValues1 = Array.from(dropdown1.selectedOptions).map(option => option.value);
			  let selectedDropdownValues2 = Array.from(dropdown2.selectedOptions).map(option => option.value);
			  let selectedDropdownValues3 = Array.from(dropdown3.selectedOptions).map(option => option.value);
			  let selectedDropdownValues4 = Array.from(dropdown4.selectedOptions).map(option => option.value);

			  for (let i = 1; i < rows.length; i++) {
				let row = rows[i];
				let cell1 = row.getElementsByTagName("td")[columnIndexForDropdown1];
				let cell2 = row.getElementsByTagName("td")[columnIndexForDropdown2];
				let cell3 = row.getElementsByTagName("td")[columnIndexForDropdown3];
				let cell4 = row.getElementsByTagName("td")[columnIndexForDropdown4];

				let cellValueDropdown1 = getDirectChildText(cell1.querySelector('a') || cell1);
				let cellValueDropdown2 = getDirectChildText(cell2.querySelector('a') || cell2);
				let cellValueDropdown3 = getDirectChildText(cell3.querySelector('a') || cell3);
				let cellValueDropdown4 = getDirectChildText(cell4.querySelector('a') || cell4);

				// Supports multiple selections from dropdownbox
				let showRowByDropdown1 = selectedDropdownValues1.includes("all") || selectedDropdownValues1.includes(cellValueDropdown1);
				let showRowByDropdown2 = selectedDropdownValues2.includes("all") || selectedDropdownValues2.includes(cellValueDropdown2);
				let showRowByDropdown3 = selectedDropdownValues3.includes("all") || selectedDropdownValues3.includes(cellValueDropdown3);
				let showRowByDropdown4 = selectedDropdownValues4.includes("all") || selectedDropdownValues4.includes(cellValueDropdown4);

				// Supports only 1 selection from dropdownbox
				//let showRowByDropdown1 = dropdown1.value === "all" || cellValueDropdown1 === dropdown1.value;
				//let showRowByDropdown2 = dropdown2.value === "all" || cellValueDropdown2 === dropdown2.value;
				//let showRowByDropdown3 = dropdown3.value === "all" || cellValueDropdown3 === dropdown3.value;
				//let showRowByDropdown4 = dropdown4.value === "all" || cellValueDropdown4 === dropdown4.value;



				/*
				// Supports only 1 CheckBox selection

				let showRowByCheckboxes = true;
				for (let checkbox of checkboxes) {
				  if (checkbox.checked) {
					let cellValue = row.getElementsByTagName("td")[columnIndexForCheckboxes].textContent;
					let checkboxValues = checkbox.value.split(",");
					if (!checkboxValues.includes(cellValue)) {
					  showRowByCheckboxes = false;
					  break;
					}
				  }
				}

				let showRowBySearch = true;
				if (searchText) {
				  showRowBySearch = false;
				  let cells = row.getElementsByTagName("td");
				  for (let cell of cells) {
					if (getDirectChildText(cell.querySelector('a') || cell).toLowerCase().includes(searchText)) {
					  showRowBySearch = true;
					  break;
					}
				  }
				}
				*/


				// Checkboxes filtering logic (supporting multiple selections)
				let showRowByCheckboxes = true;
				let anyCheckboxChecked = Array.from(checkboxes).some(checkbox => checkbox.checked);

				if (anyCheckboxChecked) {
				  showRowByCheckboxes = false; // Start with false, then look for matching checkboxes
				  for (let checkbox of checkboxes) {
					if (checkbox.checked) {
					  let cellValue = row.getElementsByTagName("td")[columnIndexForCheckboxes].textContent;
					  let checkboxValues = checkbox.value.split(",");
					  if (checkboxValues.includes(cellValue)) {
						showRowByCheckboxes = true;
						break; // As soon as one matches, no need to check further
					  }
					}
				  }
				}

/*
				// Search input filter - searching one word = whole string as is - ORIGINAL one word text search
				let showRowBySearch = true;
				if (searchText) {
				  showRowBySearch = false;
				  let cells = row.getElementsByTagName("td");
				  for (let cell of cells) {
					if (getDirectChildText(cell.querySelector('a') || cell).toLowerCase().includes(searchText)) {
					  showRowBySearch = true;
					  break;
					}
				  }
				}
*/


				// Search input filter
				let showRowBySearch = true;
				if (searchText) {
				  if (searchText.includes(",")) {
					// If searchText contains a comma, split and check for any word match
					let searchWords = searchText.split(",").map(word => word.trim());
					showRowBySearch = false;
					let cells = row.getElementsByTagName("td");

					for (let cell of cells) {
					  let cellText = getDirectChildText(cell.querySelector('a') || cell).toLowerCase();

					  // Check if any of the search words match the content
					  if (searchWords.some(word => cellText.includes(word))) {
						showRowBySearch = true;
						break;
					  }
					}
				  } else {
					// If searchText does not contain a comma, search for the whole string
					showRowBySearch = false;
					let cells = row.getElementsByTagName("td");

					for (let cell of cells) {
					  if (getDirectChildText(cell.querySelector('a') || cell).toLowerCase().includes(searchText)) {
						showRowBySearch = true;
						break;
					  }
					}
				  }
				}


				// Determine if the row should be visible
				row.style.display = (showRowByDropdown1 && showRowByDropdown2 && showRowByDropdown3 && showRowByDropdown4 && showRowByCheckboxes && showRowBySearch) ? "" : "none";
			  }


			  // Handle no visible rows
			  let visibleRowCount = 0;
			  for (let i = 1; i < rows.length; i++) {
				if (rows[i].style.display !== 'none') {
				  visibleRowCount++;
				}
			  }

			  const noResultsMessage = document.getElementById('noResultsMessage');
			  if (visibleRowCount === 0) {
				noResultsMessage.style.display = 'block';
			  } else {
				noResultsMessage.style.display = 'none';
			  }
			}
			// function combinedFilter ends


			// toggleCheckBoxes event function
			function toggleCheckboxes(checkbox) {
			  
			  /*
			  // Unchecks the other checkboxes in the group and updates the table filters
			  let checkboxes = document.getElementsByClassName("filterCheckbox");
			  for (let cb of checkboxes) {
				if (cb !== checkbox) {
					cb.checked = false;
			    }
			  }
              */
			  
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_DROPDOWN4, COLUMN_INDEX_FOR_CHECKBOXES);
			}


			// Clears the search input and updates the table filters
			function clearSearch() {
			  let searchInput = document.getElementById("searchInput");
			  searchInput.value = "";
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_DROPDOWN4, COLUMN_INDEX_FOR_CHECKBOXES);
			}

			// Resets all filters and updates the table
			function resetFilters() {
			  let searchInput = document.getElementById("searchInput");
			  searchInput.value = "";
			  
			  let checkboxes = document.getElementsByClassName("filterCheckbox");
			  for (let checkbox of checkboxes) {
				checkbox.checked = false;
			  }
			  
			  let filterDropdown1 = document.getElementById("filterDropdown1");
			  filterDropdown1.value = "all";
			  
			  let filterDropdown2 = document.getElementById("filterDropdown2");
			  filterDropdown2.value = "all";
			  
			  let filterDropdown3 = document.getElementById("filterDropdown3");
			  filterDropdown3.value = "all";
			  
			  let filterDropdown4 = document.getElementById("filterDropdown4");
			  filterDropdown4.value = "all";
			  
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_DROPDOWN4, COLUMN_INDEX_FOR_CHECKBOXES);
			}


			
			// Change Intent column cell background color based on value (Included or Excluded)
			function colorCells(tableId, columnIndex) {
			  let table = document.getElementById(tableId);
			  for (let i = 0; i < table.rows.length; i++) {
				let cell = table.rows[i].cells[columnIndex];
				if (cell) {
				  switch (cell.innerText.trim()) {
					case 'Included':
					  cell.style.backgroundColor = 'lightgreen';
					  break;
					case 'Excluded':
					  cell.style.backgroundColor = 'lightSalmon';
					  break;
					default:
					  break;
				  }
				}
			  }
			}


			// Event listeners for the dropdowns and checkboxes
			document.getElementById("filterDropdown1").addEventListener("change", function() {
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_DROPDOWN4, COLUMN_INDEX_FOR_CHECKBOXES);
			});

			document.getElementById("filterDropdown2").addEventListener("change", function() {
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_DROPDOWN4, COLUMN_INDEX_FOR_CHECKBOXES);
			});
			
			document.getElementById("filterDropdown3").addEventListener("change", function() {
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_DROPDOWN4, COLUMN_INDEX_FOR_CHECKBOXES);
			});

			document.getElementById("filterDropdown4").addEventListener("change", function() {
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_DROPDOWN4, COLUMN_INDEX_FOR_CHECKBOXES);
			});

			let checkboxes = document.getElementsByClassName("filterCheckbox");
			for (let checkbox of checkboxes) {
			  checkbox.addEventListener("change", function() {
				combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_DROPDOWN4, COLUMN_INDEX_FOR_CHECKBOXES);
			  });
			}

			// Add an event listener for the search input
			document.getElementById("searchInput").addEventListener("input", function() {
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_DROPDOWN4, COLUMN_INDEX_FOR_CHECKBOXES);
			});

		// Another approach to get function to run when loading page
		// This is not needed but left here on purpose just in case needed in the future
		//window.addEventListener('load', function() {
		//  sortTable(2, HTML_TABLE_ID, SORT_AS_INTEGER_COLUMNS);
		//});		



		// Hide named columns
		function hideColumnsByNames(tableId, columnNames) {
			let table = document.getElementById(tableId);
			let headerCells = table.getElementsByTagName('th');
			let columnIndices = [];

			// Find the indices of the columns with the given names
			for (let i = 0; i < headerCells.length; i++) {
				if (columnNames.includes(headerCells[i].innerText.trim())) {
					columnIndices.push(i);
				}
			}

			// If columns with the given names are found, hide them
			for (let index of columnIndices) {
				for (let row of table.rows) {
					row.cells[index].style.display = 'none';
				}
			}
		}

		// This static approach is used
		// if we don't use column sorting which changes bolding also to sorted column
		function boldColumnValues(tableId, columnNames) {
			let table = document.getElementById(tableId);
			let headerCells = table.getElementsByTagName('th');
			let columnIndices = [];

			// Find the indices of the columns with the given names
			for (let i = 0; i < headerCells.length; i++) {
				if (columnNames.includes(headerCells[i].innerText.trim())) {
					columnIndices.push(i);
				}
			}

			// If columns with the given names are found, set their text to bold
			for (let index of columnIndices) {
				for (let row of table.rows) {
					let cell = row.cells[index];
					if (cell) {
						cell.style.fontWeight = 'bold';
						// No need to change the background color, as it will be preserved.
					}
				}
			}
		}


		// Change html table row background colors based on Color column value
		function adjustColumnColorsByAnotherColumn(tableId, colorColumnName, targetColumnNames) {
			let table = document.getElementById(tableId);
			let headerCells = table.getElementsByTagName('th');
			let colorColumnIndex = null;
			let targetColumnIndices = [];

			// Find the index of the "Color" column (or the column specified in colorColumnName)
			for (let i = 0; i < headerCells.length; i++) {
				if (headerCells[i].innerText.trim() === colorColumnName) {
					colorColumnIndex = i;
				}
				if (targetColumnNames.includes(headerCells[i].innerText.trim())) {
					targetColumnIndices.push(i);
				}
			}

			if (colorColumnIndex === null) {
				console.error(`Couldn't find the '${colorColumnName}' column.`);
				return;
			}

			// Iterate over each row in the table
			for (let row of table.rows) {
				let colorValue = row.cells[colorColumnIndex].innerText.trim();
				if (colorValue) {
					for (let targetIndex of targetColumnIndices) {
						row.cells[targetIndex].style.backgroundColor = colorValue;
					}
				}
			}
		}


		// Change html table cell text colors based on Color column value
		function adjustColumnColorsByAnotherColumn(tableId, colorColumnName, targetColumnNames) {
			let table = document.getElementById(tableId);
			let headerCells = table.getElementsByTagName('th');
			let colorColumnIndex = null;
			let targetColumnIndices = [];

			// Find the index of the "Color" column (or the column specified in colorColumnName)
			for (let i = 0; i < headerCells.length; i++) {
				if (headerCells[i].innerText.trim() === colorColumnName) {
					colorColumnIndex = i;
				}
				if (targetColumnNames.includes(headerCells[i].innerText.trim())) {
					targetColumnIndices.push(i);
				}
			}

			if (colorColumnIndex === null) {
				console.error(`Couldn't find the '${colorColumnName}' column.`);
				return;
			}

			// Iterate over each row in the table
			for (let row of table.rows) {
				let colorValue = row.cells[colorColumnIndex].innerText.trim();
				if (colorValue) {
					let fontColor = FONT_COLOR_CONSTANTS[colorValue] || colorValue;
					for (let targetIndex of targetColumnIndices) {
						row.cells[targetIndex].style.color = fontColor;
					}
				}
			}
		}
		
		
		// Add HoverOn ToolTip from named ToolTip Column to named Target Column
		function addTooltipFromSourceColumnToTargetColumn(tableId, tooltipSourceColumnName, targetColumnNames) {
			let table = document.getElementById(tableId);
			let headerCells = table.getElementsByTagName('th');
			let sourceColumnIndex = null;
			let targetColumnIndices = [];

			// Find the index of the tooltip source column
			for (let i = 0; i < headerCells.length; i++) {
				if (headerCells[i].innerText.trim() === tooltipSourceColumnName) {
					sourceColumnIndex = i;
				}
				if (targetColumnNames.includes(headerCells[i].innerText.trim())) {
					targetColumnIndices.push(i);
				}
			}

			if (sourceColumnIndex === null) {
				console.error(`Couldn't find the '${tooltipSourceColumnName}' column.`);
				return;
			}

			// Iterate over each row in the table
			for (let row of table.rows) {
				if (row.cells[sourceColumnIndex] && row.cells[sourceColumnIndex].innerText.trim() !== "") {
					let tooltipText = row.cells[sourceColumnIndex].innerText.trim();
					
					for (let targetIndex of targetColumnIndices) {
						if (row.cells[targetIndex]) {
							// Create the tooltip span element
							let tooltipSpan = document.createElement('span');
							tooltipSpan.className = 'tooltiptext';
							tooltipSpan.innerText = tooltipText;

							// Append the tooltip to the target cell and add the tooltip class to the cell
							row.cells[targetIndex].classList.add('tooltip');
							row.cells[targetIndex].appendChild(tooltipSpan);
						}
					}
				}
			}
		}
		
		
		window.onload = function() {
			
			// Call this function to populate the first dropdown with unique values from the specified table and column
			populateSelectWithUniqueColumnValues(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, "filterDropdown1");

			// Call this function to populate the second dropdown with unique values from the specified table and column
			populateSelectWithUniqueColumnValues(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN2, "filterDropdown2");
			
			// Call this function to populate the third dropdown with unique values from the specified table and column
			populateSelectWithUniqueColumnValues(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN3, "filterDropdown3");

			// Call this function to populate the third dropdown with unique values from the specified table and column
			populateSelectWithUniqueColumnValues(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN4, "filterDropdown4");

			// Not needed anymore because sorting will also do this automatically
			//updateRowBackgroundOnColumnValueChange(HTML_TABLE_ID, 2);

			// Sort table by name so user knowns which column was sorted
			//sortTable(SORT_BY_COLUMN_INDEX, HTML_TABLE_ID, SORT_AS_INTEGER_COLUMNS);
			
			// Change Intent column background color
			colorCells(HTML_TABLE_ID, INTENT_COLUMN_INDEX);
			
			// Hide columns
			hideColumnsByNames(HTML_TABLE_ID, COLUMN_NAMES_TO_HIDE);
			
			// Bold named column(s)
			boldColumnValues(HTML_TABLE_ID, COLUMN_NAMES_TO_BOLD);

			// Change row background color value based on Color column value
			//adjustColumnColorsByAnotherColumn(HTML_TABLE_ID, COLOR_COLUMN_NAME, NAMED_COLUMNS_TO_COLOR);
			
			// Change cell text color value based on Color column value
			adjustColumnColorsByAnotherColumn(HTML_TABLE_ID, COLOR_COLUMN_NAME, NAMED_COLUMNS_TO_COLOR);
			
			// Add ToolTips to Detail column
			addTooltipFromSourceColumnToTargetColumn(HTML_TABLE_ID, 'MessageToolTip', 'Message')
			
		};
		</script>
'@


		$Title = "Get-WindowsTroubleShootingCommunity Observed Events Report"
		ConvertTo-HTML -head $head -PostContent $observedTimelineHTML, $JavascriptPostContent -PreContent $PreContent -Title $Title | Out-File "$ReportSavePath\$HTMLFileName"
		$Success = $?

		if (-not ($Success)) {
			Write-Error "Error creating HTML file."
			Write-Host "Script will exit..."
			Pause
			Exit 1
		}
		else {
			Write-Host "Get-WindowsTroubleShootingCommunity report HTML file created:`n`t$ReportSavePath\$HTMLFileName`n" -ForegroundColor Green
		}
			
		############################################################
		# Open HTML file

		# Check file exists and is bigger than 0
		# File should exist already but years ago slow computer/disk caused some problems
		# so this is hopefully not needed workaround
		# Wait max. of 20 seconds

		if(-not $DoNotOpenReportAutomatically) {
			$i = 0
			$filesize = 0
			do {
				Write-Host "Double check HTML file creation is really done (round $i)"
				$filesize = 0
				Start-Sleep -Seconds 2
				try {
					$HTMLFile = Get-ChildItem "$ReportSavePath\$HTMLFileName"
					$filesize = $HTMLFile.Length
				}
				catch {
					# Something went wrong, waiting for next round.
					Write-Host "Trouble getting file size, waiting 2 seconds and trying again..."
				}
				if ($filesize -eq 0) { Write-Host "Filesize is 0kB so waiting for a while for file creation to finish" }

				$i += 1
			} while (($i -lt 10) -and ($filesize -eq 0))

			Write-Host "Opening created file:`n$ReportSavePath\$HTMLFileName`n"
			try {
				Invoke-Item "$ReportSavePath\$HTMLFileName"
			}
			catch {
				Write-Host "Error opening file automatically to browser. Open file manually:`n$ReportSavePath\$HTMLFileName`n" -ForegroundColor Red
			}
		} else {
			Write-Host "`nNote! Parameter -DoNotOpenReportAutomatically specified. Report was not opened automatically to web browser`n" -ForegroundColor Yellow
		}
		
		} # Try creation HTML report end
		  catch {
			Write-Error "$($_.Exception.GetType().FullName)"
			Write-Error "$($_.Exception.Message)"
			Write-Error "Error creating HTML report: $ReportSavePath\$HTMLFileName"
		}
			
			
	} # If $ExportHTML end


	if($ShowLogViewerUI -or $LogViewerUI) {
		Write-Host "Show logs in Out-GridView"

		$SelectedLines = $observedTimeline | Select-Object -Property Index, Date, LogName, Multiline, ProcessRunTime, Id, LevelDisplayName, Message, MessageToolTip, Component, Context, Type, Thread, File, FileName, Line | Out-GridView -Title "Get-WindowsTroubleshootingReportCommunity $version logs viewer" -OutputMode Multiple

	} else {

		#Write-Host "`nTip: Use Parameter -ShowLogViewerUI to get LogViewerUI for graphical log viewing/debugging`n" -ForegroundColor Cyan
		
	}


























