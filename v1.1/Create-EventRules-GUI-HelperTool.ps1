# This tool will show all Windows Event provider in Out-GridView
# You can then select event provider where you want to create Known EventRules to
# Get-WindowsTroubleshootingReportCommunity tool
#
# Petri.Paavola@yodamiitti.fi
# Microsoft MVP - Windows and Intune


# Change your information here
# This will be shown on EventRules.json file
# And shared in project GitHub
$Author = "Firstname.Lastname@company.com / Super IT-Admin"


# Create Eventrules folder if not exists
if(-not (Test-Path "$PSScriptRoot\EventRules")) {
	New-Item -ItemType Directory -Path "$PSScriptRoot\EventRules" -Force
}


Write-Host "Select Event Provider to get Event IDs from"

# List Event log Providers in Out-GridView
$SelectedEventProviders = Get-WinEvent -ListProvider '*' -ErrorAction SilentlyContinue | Sort-Object -Property ProviderName | Out-GridView -Title 'Select Event Viewer log to show available Event IDs' -OutputMode Single

Write-Host "Selected $($SelectedEventProviders.ProviderName)"

Foreach ($EventProviderObject in $SelectedEventProviders) {
	# Show selected EventProvider
	$EventProviderObject | Format-List -Property *

	# Make a copy of events so we can add our own custom properties later
	$Events = (Get-WinEvent -ListProvider $EventProviderObject.ProviderName).Events

	# Add Level as clear text (Informational, Error, Warning)
	Foreach ($Event in $Events) {
		# DEBUG
		#Write-Host "DEBUG event:"
		#$Event | ConvertTo-Json

		# Add DisplayName property for Event
		$Event | Add-Member -MemberType NoteProperty -Name LevelDisplayName -Value $Event.Level.DisplayName

		# Add LogName property for Event
		$Event | Add-Member -MemberType NoteProperty -Name LogName -Value $Event.LogLink.LogName

	}

	Write-Host "Select Event IDs you want to add to your custom EventRules json file"

	# Show available Event Provider Event Ids
	$SelectedEventIds = $Events | Select-Object -Property Id, Description, LevelDisplayName, LogName, LogLink, Level | Out-GridView -Title 'Select objects for KnownRules.json' -OutputMode Multiple

	if(-not $SelectedEventIds) {
		Write-Host "No Event IDs selected. Skipping EventProvider $($EventProviderObject.ProviderName)"
		Continue
	}

	$CategoryName = $null
	While(-not $CategoryName) {
		$CategoryName = Read-Host "Enter CategoryName for KnownRules.json: "
	}
	
	$EventRulesFileFullPath = "$PSScriptRoot\EventRules\EventRules-$($CategoryName).json"
	Write-Host "Create EventRules file: $EventRulesFileFullPath"

	$EventRulesArray = [System.Collections.Generic.List[PSObject]]@()

	# Translate selected Event provider Ids to KnownRules.json syntax
	Foreach($SelectedEventId in $SelectedEventIds | Sort-Object -Property Id) {
		# Create custom Powershell object

		# DEBUG
		#$SelectedEventId | fl *
		#$SelectedEventId | ConvertTo-Json -Depth 4 | Set-Clipboard

		# Event type: Informational, Error
		Switch ($SelectedEventId.Level.Value) {
			'2' { 
					# Error
					$Color = 'Red'
				}
			'3' { 
					# Warning
					$Color = ''
				}
			'4' { 
					# Informational
					$Color = ''
				}
			Default {
				$Color = ''
			}
		}

		$EventRuleCustomObject = ([PSCustomObject]@{
			"CategoryName" = $CategoryName
			"LogType" = '.evtx'
			"Channel" = $SelectedEventId.LogLink.LogName
			"Id" = $SelectedEventId.Id
			"LevelDisplayName" = $SelectedEventId.LevelDisplayName
			"ProviderName" = "$($EventProviderObject.ProviderName)"
			"IncludeEventXMLDataInMessage" = $false
			"IncludeEventXMLDataInToolTip" = $false
			"ToolTipText" = ''
			"Color" = $Color
			"DeveloperNotes" = "$($SelectedEventId.Description)" 
			"Author" = $Author
			"LinkToBlogArticle" = ''
		})

		# DEBUG $CustomTimelineObject
		#$CustomTimelineObject
		#Pause

		$EventRulesArray.add($EventRuleCustomObject)
	}

	# Create main CategoryObject
	# And add previously selected EventIds to Property KnownEventRules
	$EventRulesObjectForJSONFile = ([PSCustomObject]@{
		"CategoryName" = $CategoryName
		"KnownEventRules" = $EventRulesArray
	})



	# DEBUG
	#$EventRulesArray
	#$EventRulesArray | ConvertTo-Json | Set-Clipboard
	
	# DEBUG
	#$EventRulesObjectForJSONFile

	# Save EventRules to json file
	$EventRulesObjectForJSONFile | ConvertTo-Json -Depth 3 | Out-File -FilePath "$EventRulesFileFullPath" -Force
	$Success = $?
	
	if($Success) {
		Write-Host "File saved successfully" -ForegroundColor Green
	} else {
		Write-Host "Failed to save file!" -ForegroundColor Red
	}

}
