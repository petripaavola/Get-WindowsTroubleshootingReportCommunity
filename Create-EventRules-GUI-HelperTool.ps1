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


Write-Host "Select Event Provider to get Event IDs from"

# List Event log Providers in Out-GridView
$SelectedEventProviders = Get-WinEvent -ListProvider '*' -ErrorAction SilentlyContinue | Sort-Object -Property ProviderName | Out-GridView -Title 'Select Event Viewer log to show available Event IDs' -OutputMode Single

Write-Host "Selected $($SelectedEventProviders.ProviderName)"

Foreach ($EventProviderObject in $SelectedEventProviders) {
	# Show selected EventProvider
	$EventProviderObject | Format-List -Property *

	Write-Host "Select Event IDs you want to add to your custom EventRules json file"

	# Show available Event Provider Event Ids
	$SelectedEventIds = (Get-WinEvent -ListProvider $EventProviderObject.ProviderName).Events | Select-Object -Property Id, Description, Level, LogLink | Out-GridView -Title 'Select objects for KnownRules.json' -OutputMode Multiple

	$CategoryName = $null
	While(-not $CategoryName) {
		$CategoryName = Read-Host "Enter CategoryName for KnownRules.json: "
	}
	
	$EventRulesFileFullPath = "$PSScriptRoot\EventRules-$($CategoryName).json"
	Write-Host "Create EventRules file: $EventRulesFileFullPath"

	$EventRulesArray = [System.Collections.Generic.List[PSObject]]@()

	# Translate selected Event provider Ids to KnownRules.json syntax
	Foreach($SelectedEventId in $SelectedEventIds) {
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
