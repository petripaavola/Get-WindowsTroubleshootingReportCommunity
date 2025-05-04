

[CmdletBinding()]
Param()


$ScriptVersion = "0.91"
$TimeOutBetweenGraphAPIRequests = 300


Write-Host "Get-Intune-Apps-and-Scripts-GUIDs-and-Names.ps1 $ScriptVersion" -ForegroundColor Cyan
Write-Host "Author: Petri.Paavola@yodamiitti.fi / Microsoft MVP - Windows and Intune"
Write-Host ""



################ Functions ################

	Function Get-AppIntent {
		Param(
			$AppId
			)

		$intent = 'Unknown Intent'

		if($AppId) {
			if($IdHashtable.ContainsKey($AppId)) {
				$AppPolicy=$IdHashtable[$AppId]

				if($AppPolicy.Intent) {
					Switch ($AppPolicy.Intent)
					{
						0	{ $intent = 'Not Targeted' }
						1	{ $intent = 'Available Install' }
						3	{ $intent = 'Required Install' }
						4	{ $intent = 'Required Uninstall' }
						default { $intent = 'Unknown Intent' }
					}
				}				
			}
		}
		
		return $intent
	}

	Function Get-AppIntentNameForNumber {
		Param(
			$IntentNumber
			)

		Switch ($IntentNumber)
		{
			0	{ $intent = 'Not Targeted' }
			1	{ $intent = 'Available Install' }
			3	{ $intent = 'Required Install' }
			4	{ $intent = 'Required Uninstall' }
			default { $intent = 'Unknown Intent' }
		}
		
		return $intent
	}

	Function Get-AppDetectionResultForNumber {
		Param(
			$DetectionNumber
			)

		Switch ($DetectionNumber)
		{
			0	{ $DetectionState = 'Unknown' }
			1	{ $DetectionState = 'Detected' }
			2	{ $DetectionState = 'Not Detected' }
			3	{ $DetectionState = 'Unknown' }
			4	{ $DetectionState = 'Unknown' }
			5	{ $DetectionState = 'Unknown' }
			default { $DetectionState = 'Unknown' }
		}
		
		return $DetectionState
	}


	Function Get-AppName {
		Param(
			$AppId
			)

		$AppName = $null

		if($AppId) {
			if($IdHashtable.ContainsKey($AppId)) {
				$AppPolicy=$IdHashtable[$AppId]

				if($AppPolicy.Name) {
					$AppName = $AppPolicy.Name
				}				
			}
		}
		
		return $AppName
	}


	Function Get-AppType {
		Param(
			$AppId
			)

		$AppType = 'App'

		if($AppId) {
			if($IdHashtable.ContainsKey($AppId)) {
				$AppPolicy=$IdHashtable[$AppId]

				if($AppPolicy.InstallerData) {
					# This should be New Store App
					$AppType = 'WinGetApp'
				} else {
					# This should be Win32App
					$AppType = 'Win32App'
				}
			}
		}
		
		return $AppType
	}

	Function Convert-AppDetectionValuesToHumanReadable {
		Param(
			$DetectionRulesObject
		)

		# Object has DetectionType and DetectionText objects
		# Object is array of objects
		<#		
		[
		  {
			"DetectionType":  2,
			"DetectionText":  {
								  "Path":  "C:\\Program Files (x86)\\Foo",
								  "FileOrFolderName":  "bar.exe",
								  "Check32BitOn64System":  true,
								  "DetectionType":  1,
								  "Operator":  0,
								  "DetectionValue":  null
							  }
		  }
		]
		#>

		foreach($DetectionRule in $DetectionRulesObject) {


			# Change DetectionText properties values to text
			
			# DetectionType: Registry
			if($DetectionRule.DetectionType -eq 0) {
			
				# Registry Detection Type values
				# https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobappregistrydetectiontype?view=graph-rest-beta
				Switch ($DetectionRule.DetectionText.DetectionType)  {
					0 { $DetectionRule.DetectionText.DetectionType = 'Not configure' }
					1 { $DetectionRule.DetectionText.DetectionType = 'Value exists' }
					2 { $DetectionRule.DetectionText.DetectionType = 'Value does not exist' }
					3 { $DetectionRule.DetectionText.DetectionType = 'String comparison' }
					4 { $DetectionRule.DetectionText.DetectionType = 'Integer comparison' }
					5 { $DetectionRule.DetectionText.DetectionType = 'Version comparison' }
				}
				
				# Registry Detection Operation values
				# https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobappruleoperator?view=graph-rest-beta
				Switch ($DetectionRule.DetectionText.Operator)  {
					0 { $DetectionRule.DetectionText.Operator = 'Not configured' }
					1 { $DetectionRule.DetectionText.Operator = 'Equals' }
					2 { $DetectionRule.DetectionText.Operator = 'Not equal to' }
					4 { $DetectionRule.DetectionText.Operator = 'Greater than' }
					5 { $DetectionRule.DetectionText.Operator = 'Greater than or equal to' }
					8 { $DetectionRule.DetectionText.Operator = 'Less than' }
					9 { $DetectionRule.DetectionText.Operator = 'Less than or equal to' }
				}
			}


			# DetectionType: File
			if($DetectionRule.DetectionType -eq 2) {

				# File Detection Type values
				# https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobappfilesystemdetectiontype?view=graph-rest-beta
				Switch ($DetectionRule.DetectionText.DetectionType)  {
					0 { $DetectionRule.DetectionText.DetectionType = 'Not configure' }
					1 { $DetectionRule.DetectionText.DetectionType = 'File or folder exists' }
					2 { $DetectionRule.DetectionText.DetectionType = 'Date modified' }
					3 { $DetectionRule.DetectionText.DetectionType = 'Date created' }
					4 { $DetectionRule.DetectionText.DetectionType = 'String (version)' }
					5 { $DetectionRule.DetectionText.DetectionType = 'Size in MB' }
					6 { $DetectionRule.DetectionText.DetectionType = 'File or folder does not exist' }
				}
				
				# File Detection Operator values
				# https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobappdetectionoperator?view=graph-rest-beta
				Switch ($DetectionRule.DetectionText.Operator)  {
					0 { $DetectionRule.DetectionText.Operator = 'Not configured' }
					1 { $DetectionRule.DetectionText.Operator = 'Equals' }
					2 { $DetectionRule.DetectionText.Operator = 'Not equal to' }
					4 { $DetectionRule.DetectionText.Operator = 'Greater than' }
					5 { $DetectionRule.DetectionText.Operator = 'Greater than or equal to' }
					8 { $DetectionRule.DetectionText.Operator = 'Less than' }
					9 { $DetectionRule.DetectionText.Operator = 'Less than or equal to' }
				}
			}
			

			# DetectionType: Custom script
			if($DetectionRule.DetectionType -eq 3) {

				# Convert base64 script to clear text
				#$DetectionRule.DetectionText.ScriptBody

				# Decode Base64 content
				$b = [System.Convert]::FromBase64String("$($DetectionRule.DetectionText.ScriptBody)")
				$DetectionRule.DetectionText.ScriptBody = [System.Text.Encoding]::UTF8.GetString($b)

			}
			
			
			<#
			# Change DetectionType value to text
			Switch ($DetectionRule.DetectionType) {
				0 { $DetectionRule.DetectionType = 'Registry' }
				1 { $DetectionRule.DetectionType = 'MSI' }
				2 { $DetectionRule.DetectionType = 'File' }
				3 { $DetectionRule.DetectionType = 'Custom script' }
				default { $DetectionRule.DetectionType = $DetectionRule.DetectionType }
			}
			#>
			
			# Add new property with DetectionType value as text
			Switch ($DetectionRule.DetectionType) {
				0 { $DetectionRule | Add-Member -MemberType noteProperty -Name DetectionTypeAsText -Value 'Registry' }
				1 { $DetectionRule | Add-Member -MemberType noteProperty -Name DetectionTypeAsText -Value 'MSI' }
				2 { $DetectionRule | Add-Member -MemberType noteProperty -Name DetectionTypeAsText -Value 'File' }
				3 { $DetectionRule | Add-Member -MemberType noteProperty -Name DetectionTypeAsText -Value 'Custom script' }
				default { $DetectionRule | Add-Member -MemberType noteProperty -Name DetectionTypeAsText -Value $DetectionRule.DetectionType }
			}			
			
		}			
		
		return $DetectionRulesObject
	}


	function Invoke-MgGraphRequestGetAllPages {
		param (
			[Parameter(Mandatory = $true)]
			[String]$uri
		)

		$MgGraphRequest = $null
		$AllMSGraphRequest = $null

		Start-Sleep -Milliseconds $TimeOutBetweenGraphAPIRequests

		try {

			# Save results to this variable
			$allGraphAPIData = @()

			do {

				$MgGraphRequest = $null
				$MgGraphRequest = Invoke-MgGraphRequest -Uri $uri -Method 'Get' -OutputType PSObject -ContentType "application/json"

				if($MgGraphRequest) {

					# Test if object has attribute named Value (whether value is null or not)
					#if((Get-Member -inputobject $MgGraphRequest -name 'Value' -Membertype Properties) -and (Get-Member -inputobject $MgGraphRequest -name '@odata.context' -Membertype Properties)) {
					if(Get-Member -inputobject $MgGraphRequest -name 'Value' -Membertype Properties) {
						# Value property exists
						$allGraphAPIData += $MgGraphRequest.Value

						# Check if we have value starting https:// in attribute @odate.nextLink
						# and check that $Top= parameter was NOT used. With $Top= parameter we can limit search results
						# but that almost always results .nextLink being present if there is more data than specified with top
						# If we specified $Top= ourselves then we don't want to fetch nextLink values
						#
						# So get GraphAllPages if there is valid nextlink and $Top= was NOT used in url originally
						if (($MgGraphRequest.'@odata.nextLink' -like 'https://*') -and (-not ($uri.Contains('$top=')))) {
							# Save nextLink url to variable and rerun do-loop
							$uri = $MgGraphRequest.'@odata.nextLink'
							Start-Sleep -Milliseconds $TimeOutBetweenGraphAPIRequests

							# Continue to next round in Do-loop
							Continue

						} else {
							# We dont have nextLink value OR
							# $top= exists so we return what we got from first round
							#return $allGraphAPIData
							$uri = $null
						}
						
					} else {
						# Sometimes we get results without Value-attribute (eg. getting user details)
						# We will return all we got as is
						# because there should not be nextLink page in this case ???
						return $MgGraphRequest
					}
				} else {
					# Invoke-MGGraphRequest failed so we return false
					return $null
				}
				
			} while ($uri) # Always run once and continue if there is nextLink value


			# We should not end here but just in case
			return $allGraphAPIData

		} catch {
			Write-Error "There was error with MGGraphRequest with url $url!"
			return $null
		}
	}


	function Get-IntunePowershellScriptContentInCleartext {
		Param(
			[Parameter(Mandatory=$true)]
			[String]$PowershellScriptPolicyId
		)

		# Check if we already have value
		if($IdHashtable[$PowershellScriptPolicyId].scriptContentClearText) {
			# Powershell script in clear text already exists in HashTable object
			# So we can return it
			
			return $IdHashtable[$PowershellScriptPolicyId].scriptContentClearText

		} else {
			# Property scriptContentClearText does NOT exist in HashTable object

			# Download Powershell script in cleartext if -Online parameter has been specified
			# PowerShell script is in clear text in IME log files so this is always shown
			# And later version might get PowerShell script from IME log instead from Graph
			if($Online) {

				#Write-Verbose "Downloading Powershell script: $PowershellScriptPolicyId"
				
				$uri = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts/$PowershellScriptPolicyId"
				$IntunePowershellScriptInformation = Invoke-MgGraphRequestGetAllPages -Uri $uri

				if($IntunePowershellScriptInformation) {
					#Write-Verbose "Done" -ForegroundColor Green
					
					if($IntunePowershellScriptInformation.scriptContent) {

						Try {
							# Convert Intune Powershell script base64 content to clear text
							$b = [System.Convert]::FromBase64String("$($IntunePowershellScriptInformation.scriptContent)")
							$IntuneScriptContentInClearText = [System.Text.Encoding]::UTF8.GetString($b)
						} catch {
							# Some fatal error converting base64 to cleartext
							Write-Error "Error converting Intune Powershell script base64 to cleartext" -ForegroundColor Red
							return 'N/A'
						}

						# Add new property to HashTable object if it doesn't already exist
						if(-not $IdHashtable[$PowershellScriptPolicyId].scriptContentClearText) {
							
							$IdHashtable[$PowershellScriptPolicyId] | Add-Member -MemberType noteProperty -Name scriptContentClearText -Value $IntuneScriptContentInClearText
						} else {
							$IdHashtable[$PowershellScriptPolicyId].scriptContentClearText = $IntuneScriptContentInClearText
						}

						return $IntuneScriptContentInClearText
					} else {
						# Did not get scriptContent information
						# We should never get here if we got anything successfully from Graph API
						Write-Verbose "Failed to download Powershell scriptContent property" -ForegroundColor Yellow
						return 'N/A'
					}
				} else {
					# Could not get Intune Powershell information from Intune
					# Doing nothing
					Write-Host "Failed to download Powershell script from Intune" -ForegroundColor Yellow
					return 'N/A'
				}
			} else {
				# -Online not selected and we didn't have value so return $null
				return 'N/A'
			}
		}
		# We should not get here
		return 'N/A'
	}


	function Get-IntuneRemediationDetectionScriptContentInCleartext {
		Param(
			[Parameter(Mandatory=$true)]
				[String]$ScriptPolicyId
		)


		# Check if we already have value
		if($IdHashtable[$ScriptPolicyId].detectionScriptContentClearText) {
			# Powershell script in clear text already exists in HashTable object
			# So we can return it
			
			return $IdHashtable[$ScriptPolicyId].detectionScriptContentClearText

		} else {
			# Property detectionScriptContent does NOT exist in HashTable object

			# Download Remediation Detection script in cleartext if -Online parameter has been specified
			# and -DoNotDownloadClearTextRemediationScriptsToReport is NOT specified
			if($Online -and (-not $DoNotDownloadClearTextRemediationScriptsToReport)) {

				#Write-Verbose "Downloading Remediation Detection script: $ScriptPolicyId"
				
				$uri = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/$ScriptPolicyId"
				#Write-Verbose "URI: $uri"
				$IntuneRemediationScriptInformation = Invoke-MgGraphRequestGetAllPages -Uri $uri

				if($IntuneRemediationScriptInformation) {
					#Write-Verbose "Done" -ForegroundColor Green
					
					if($IntuneRemediationScriptInformation.detectionScriptContent) {

						Try {
							# Convert Intune Remediation Detection script base64 content to clear text
							$b = [System.Convert]::FromBase64String("$($IntuneRemediationScriptInformation.detectionScriptContent)")
							$IntuneScriptContentInClearText = [System.Text.Encoding]::UTF8.GetString($b)
						} catch {
							# Some fatal error converting base64 to cleartext
							Write-Error "Error converting Intune Remediation Detection script base64 to cleartext" -ForegroundColor Red
							return 'N/A'
						}

						# Add new property to HashTable object if it doesn't already exist
						if(-not $IdHashtable[$ScriptPolicyId].detectionScriptContentClearText) {
							
							$IdHashtable[$ScriptPolicyId] | Add-Member -MemberType noteProperty -Name detectionScriptContentClearText -Value $IntuneScriptContentInClearText
						} else {
							$IdHashtable[$PowershellScriptPolicyId].detectionScriptContentClearText = $IntuneScriptContentInClearText
						}

						return $IntuneScriptContentInClearText
					} else {
						# Did not get detectionScriptContent information
						# We should never get here if we got anything successfully from Graph API
						Write-Verbose "Failed to download Remediation Detect detectionScriptContent property" -ForegroundColor Yellow
						return 'N/A'
					}
				} else {
					# Could not get Intune Remediation Detect information from Intune
					# Doing nothing
					Write-Host "Failed to download Remediation Detect script from Intune" -ForegroundColor Yellow
					return 'N/A'
				}
			} else {
				# -Online not selected and we didn't have value so return $null
				return 'N/A'
			}
		}
		# We should not get here
		return 'N/A'
	}


	function Get-IntuneRemediationRemediateScriptContentInCleartext {
		Param(
			[Parameter(Mandatory=$true)]
				[String]$ScriptPolicyId
		)


		# Check if we already have value
		if($IdHashtable[$ScriptPolicyId].RemediateScriptContentClearText) {
			# Powershell script in clear text already exists in HashTable object
			# So we can return it
			
			return $IdHashtable[$ScriptPolicyId].RemediateScriptContentClearText

		} else {
			# Property remediationScriptContent does NOT exist in HashTable object

			# Download Remediation Remediate script in cleartext if -Online parameter has been specified
			# and -DoNotDownloadClearTextRemediationScriptsToReport is NOT specified
			if($Online -and (-not $DoNotDownloadClearTextRemediationScriptsToReport)) {

				#Write-Verbose "Downloading Remediation Remediate script: $ScriptPolicyId"
				
				$uri = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/$ScriptPolicyId"
				$IntuneRemediationScriptInformation = Invoke-MgGraphRequestGetAllPages -Uri $uri

				if($IntuneRemediationScriptInformation) {
					#Write-Verbose "Done" -ForegroundColor Green
					
					if($IntuneRemediationScriptInformation.remediationScriptContent) {

						Try {
							# Convert Intune Remediation Detection script base64 content to clear text
							$b = [System.Convert]::FromBase64String("$($IntuneRemediationScriptInformation.remediationScriptContent)")
							$IntuneScriptContentInClearText = [System.Text.Encoding]::UTF8.GetString($b)
						} catch {
							# Some fatal error converting base64 to cleartext
							Write-Error "Error converting Intune Remediation Detection script base64 to cleartext" -ForegroundColor Red
							return 'N/A'
						}

						# Add new property to HashTable object if it doesn't already exist
						if(-not $IdHashtable[$ScriptPolicyId].RemediateScriptContentClearText) {
							
							$IdHashtable[$ScriptPolicyId] | Add-Member -MemberType noteProperty -Name RemediateScriptContentClearText -Value $IntuneScriptContentInClearText
						} else {
							$IdHashtable[$PowershellScriptPolicyId].RemediateScriptContentClearText = $IntuneScriptContentInClearText
						}

						return $IntuneScriptContentInClearText
					} else {
						# Did not get remediationScriptContent information
						# We should never get here if we got anything successfully from Graph API
						Write-Verbose "Failed to download Powershell remediationScriptContent property" -ForegroundColor Yellow
						return 'N/A'
					}
				} else {
					# Could not get Intune Remediation script information from Intune
					# Doing nothing
					Write-Host "Failed to download Remediation Remediate script from Intune" -ForegroundColor Yellow
					return 'N/A'
				}
			} else {
				# -Online not selected and we didn't have value so return $null
				return 'N/A'
			}
		}
		# We should not get here
		return 'N/A'
	}


	function Get-IntuneCustomComplianceScriptContentInCleartext {
		Param(
			[Parameter(Mandatory=$true)]
				[String]$ScriptPolicyId
		)


		# Check if we already have value
		if($IdHashtable[$ScriptPolicyId].detectionScriptContentClearText) {
			# Powershell script in clear text already exists in HashTable object
			# So we can return it
			
			return $IdHashtable[$ScriptPolicyId].detectionScriptContentClearText

		} else {
			# Property detectionScriptContent does NOT exist in HashTable object

			# Download Custom Compliance script in cleartext if -Online parameter has been specified
			# and -DoNotDownloadClearTextRemediationScriptsToReport is NOT specified
			if($Online -and (-not $DoNotDownloadClearTextRemediationScriptsToReport)) {

				#Write-Verbose "Downloading Custom Compliance script: $ScriptPolicyId"
				
				$uri = "https://graph.microsoft.com/beta/deviceManagement/deviceComplianceScripts/$ScriptPolicyId"
				$IntuneRemediationScriptInformation = Invoke-MgGraphRequestGetAllPages -Uri $uri

				if($IntuneRemediationScriptInformation) {
					#Write-Verbose "Done" -ForegroundColor Green
					
					if($IntuneRemediationScriptInformation.detectionScriptContent) {

						Try {
							# Convert Intune Custom Compliance script base64 content to clear text
							$b = [System.Convert]::FromBase64String("$($IntuneRemediationScriptInformation.detectionScriptContent)")
							$IntuneScriptContentInClearText = [System.Text.Encoding]::UTF8.GetString($b)
						} catch {
							# Some fatal error converting base64 to cleartext
							Write-Error "Error converting Intune Custom Compliance script base64 to cleartext" -ForegroundColor Red
							return 'N/A'
						}

						# Add new property to HashTable object if it doesn't already exist
						if(-not $IdHashtable[$ScriptPolicyId].detectionScriptContentClearText) {
							
							$IdHashtable[$ScriptPolicyId] | Add-Member -MemberType noteProperty -Name detectionScriptContentClearText -Value $IntuneScriptContentInClearText
						} else {
							$IdHashtable[$PowershellScriptPolicyId].detectionScriptContentClearText = $IntuneScriptContentInClearText
						}

						return $IntuneScriptContentInClearText
					} else {
						# Did not get detectionScriptContent information
						# We should never get here if we got anything successfully from Graph API
						Write-Verbose "Failed to download Custom Compliance detectionScriptContent property" -ForegroundColor Yellow
						return 'N/A'
					}
				} else {
					# Could not get Intune Custom Compliance information from Intune
					# Doing nothing
					Write-Host "Failed to download Custom Compliance script from Intune" -ForegroundColor Yellow
					return 'N/A'
				}
			} else {
				# -Online not selected and we didn't have value so return $null
				return 'N/A'
			}
		}
		# We should not get here
		return 'N/A'
	}


################ Functions ################


# Download Powershell scripts, Proactive Remediation scripts and custom Compliance Policy scripts
if ($True) {

	$GraphAuthenticationModule = $null
	$MgContext = $null

	Write-Host "Connecting to Intune using module Microsoft.Graph.Authentication"

	Write-Host "Import module Microsoft.Graph.Authentication"
	Import-Module Microsoft.Graph.Authentication
	$Success = $?

	if($Success) {
		# Module imported successfully
		Write-Host "Success`n" -ForegroundColor Green
	} else {
		Write-Host "Failed"  -ForegroundColor Red
		Write-Host "Make sure you have installed module Microsoft.Graph.Authentication"
		Write-Host "You can install module without admin rights to your user account with command:`n`nInstall-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser" -ForegroundColor Yellow
		Write-Host "`nor you can install machine-wide module with with admin rights using command:`nInstall-Module -Name Microsoft.Graph.Authentication"
		Write-Host ""
		Exit 1
	}

	Write-Host "Connect to Microsoft Graph API"

	$scopes = "DeviceManagementConfiguration.Read.All"
	$MgGraph = Connect-MgGraph -scopes $scopes
	$Success = $?

	if ($Success -and $MgGraph) {
		Write-Host "Success`n" -ForegroundColor Green

		# Get MgGraph session details
		$MgContext = Get-MgContext
		
		if($MgContext) {
		
			$TenantId = $MgContext.TenantId
			$AdminUserUPN = $MgContext.Account

			Write-Host "Connected to Intune tenant:`n$TenantId`n$AdminUserUPN`n"

		} else {
			Write-Host "Error getting MgContext information!`nScript will exit!" -ForegroundColor Red
			Exit 1
		}
		
	} else {
		Write-Host "Could not connect to Graph API!" -ForegroundColor Red
		Exit 1
	}

	# Download Intune scripts information if we have connection to Microsoft Graph API
	if($MgContext) {

		Write-Host "Download Intune Powershell scripts"
		# Get PowerShell Scripts
		$uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts'
		$AllIntunePowershellScripts = Invoke-MgGraphRequestGetAllPages -Uri $uri

		if($AllIntunePowershellScripts) {
			#Write-Host "Done" -ForegroundColor Green
			
			# Add Name Property to object
			$AllIntunePowershellScripts | Foreach-Object { $_ | Add-Member -MemberType noteProperty -Name name -Value $_.displayName }
			
			# Add all PowershellScripts to Hashtable
			#$AllIntunePowershellScripts | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }
			
			# Save to json file
			$AllIntunePowershellScripts | ConvertTo-Json -Depth 4 | Out-File -FilePath "$PSScriptRoot\KnownGUIDS-IntunePowershellScripts.json" -Force

			Write-Host "Done. Downloaded $($AllIntunePowershellScripts.Count) PowerShell scripts" -ForegroundColor Green

		} else {
			Write-Host "Did not find Intune Powershell scripts`n" -ForegroundColor Yellow
		}

		Start-Sleep -MilliSeconds 500
		
		Write-Host "Download Intune Remediations Scripts"
		# Get Proactive Remediations Scripts
		$uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts'
		$AllIntuneProactiveRemediationsScripts = Invoke-MgGraphRequestGetAllPages -Uri $uri

		if($AllIntuneProactiveRemediationsScripts) {
			#Write-Host "Done" -ForegroundColor Green
			
			# Add Name Property to object
			$AllIntuneProactiveRemediationsScripts | Foreach-Object { $_ | Add-Member -MemberType noteProperty -Name name -Value $_.displayName }
			
			# Add policyType 6
			# Which is Remediation script type
			$AllIntuneProactiveRemediationsScripts | Foreach-Object { $_ | Add-Member -MemberType noteProperty -Name policyType -Value 6 }
			
			# Add all PowershellScripts to Hashtable
			#$AllIntuneProactiveRemediationsScripts | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }
				
			# Save to json file
			$AllIntuneProactiveRemediationsScripts | ConvertTo-Json -Depth 4 | Out-File -FilePath "$PSScriptRoot\KnownGUIDS-IntuneRemediationScripts.json" -Force
			
			Write-Host "Done. Downloaded $($AllIntuneProactiveRemediationsScripts.Count) Remediation scripts`n" -ForegroundColor Green

		} else {
			Write-Host "Did not find Intune Remediation scripts" -ForegroundColor Yellow
		}

		Start-Sleep -MilliSeconds 500

		Write-Host "Download Intune Windows Device Compliance custom Scripts"
		# Get Windows Device Compliance custom Scripts
		$uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceComplianceScripts'
		$AllIntuneCustomComplianceScripts = Invoke-MgGraphRequestGetAllPages -Uri $uri

		if($AllIntuneCustomComplianceScripts) {
			#Write-Host "Done" -ForegroundColor Green
			
			# Add Name Property to object
			$AllIntuneCustomComplianceScripts | Foreach-Object { $_ | Add-Member -MemberType noteProperty -Name name -Value $_.displayName }

			# Add policyType 8
			# Which is Custom Compliance script type
			$AllIntuneCustomComplianceScripts | Foreach-Object { $_ | Add-Member -MemberType noteProperty -Name policyType -Value 8 }
			
			# Add all PowershellScripts to Hashtable
			#$AllIntuneCustomComplianceScripts | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }
			
			# Save to json file
			$AllIntuneCustomComplianceScripts | ConvertTo-Json -Depth 4 | Out-File -FilePath "$PSScriptRoot\KnownGUIDS-IntuneCustomComplianceScripts.json" -Force

			Write-Host "Done. Downloaded $($AllIntuneCustomComplianceScripts.Count) Custom Compliance scripts`n" -ForegroundColor Green

		} else {
			Write-Host "Did not find Intune Windows custom Compliance scripts" -ForegroundColor Yellow
		}
		
		Start-Sleep -MilliSeconds 500
		
		Write-Host "Download Intune Filters"
		$uri = 'https://graph.microsoft.com/beta/deviceManagement/assignmentFilters?$select=*'
		$AllIntuneFilters = Invoke-MgGraphRequestGetAllPages -Uri $uri

		if($AllIntuneFilters) {
			#Write-Host "Done" -ForegroundColor Green
			
			# Add all Filters to Hashtable
			#$AllIntuneFilters | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }

			# Save to json file
			$AllIntuneFilters | ConvertTo-Json -Depth 4 | Out-File -FilePath "$PSScriptRoot\KnownGUIDS-IntuneFilters.json" -Force

			Write-Host "Done. Downloaded $($AllIntuneFilters.Count) Filters`n" -ForegroundColor Green

		} else {
			Write-Host "Did not find Intune filters" -ForegroundColor Yellow
		}

		
		Write-Host "Download Intune Apps"
		$Uri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=(isof(%27microsoft.graph.win32CatalogApp%27)%20or%20isof(%27microsoft.graph.windowsStoreApp%27)%20or%20isof(%27microsoft.graph.microsoftStoreForBusinessApp%27)%20or%20isof(%27microsoft.graph.officeSuiteApp%27)%20or%20(isof(%27microsoft.graph.win32LobApp%27)%20and%20not(isof(%27microsoft.graph.win32CatalogApp%27)))%20or%20isof(%27microsoft.graph.windowsMicrosoftEdgeApp%27)%20or%20isof(%27microsoft.graph.windowsPhone81AppX%27)%20or%20isof(%27microsoft.graph.windowsPhone81StoreApp%27)%20or%20isof(%27microsoft.graph.windowsPhoneXAP%27)%20or%20isof(%27microsoft.graph.windowsAppX%27)%20or%20isof(%27microsoft.graph.windowsMobileMSI%27)%20or%20isof(%27microsoft.graph.windowsUniversalAppX%27)%20or%20isof(%27microsoft.graph.webApp%27)%20or%20isof(%27microsoft.graph.windowsWebApp%27)%20or%20isof(%27microsoft.graph.winGetApp%27))%20and%20(microsoft.graph.managedApp/appAvailability%20eq%20null%20or%20microsoft.graph.managedApp/appAvailability%20eq%20%27lineOfBusiness%27%20or%20isAssigned%20eq%20true)&`$orderby=displayName&"
		$AllIntuneApps = Invoke-MgGraphRequestGetAllPages -Uri $uri

		if($AllIntuneApps) {
			#Write-Host "Done" -ForegroundColor Green
			
			# Add all Filters to Hashtable
			#$AllIntuneApps | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }

			# Save to json file
			$AllIntuneApps | ConvertTo-Json -Depth 4 | Out-File -FilePath "$PSScriptRoot\KnownGUIDS-IntuneApps.json" -Force

			Write-Host "Done. Downloaded $($AllIntuneApps.Count) Apps`n" -ForegroundColor Green

		} else {
			Write-Host "Did not find Intune Apps" -ForegroundColor Yellow
		}

	} else {
		Write-Host "Not connected to Microsoft Intune, skip downloading script names...." -ForegroundColor Yellow
	}
}
Write-Host ""

Write-Host "Script end"

