<#
.SYNOPSIS
This script retrieves and exports the names and versions of installed MSI applications from the Windows registry.

.DESCRIPTION
The script searches through specific registry paths to find installed MSI applications. It looks for subkeys that match the GUID pattern under the specified uninstall paths. For each matching subkey, it extracts the DisplayName and DisplayVersion properties. If a DisplayName is found, it constructs a custom object containing the GUID, DisplayName, and a category. The results are then exported to a JSON file.

.PARAMETER uninstallPaths
An array of registry paths to search for installed MSI applications. Default paths are:
- "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
- "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

.EXAMPLE
.\Get-MSI-Apps-GUIDs-and-Names.ps1
This command runs the script and exports the MSI application names and versions to the KnownGUIDs-MSI.json file.


.NOTES
Author: Petri Paavola
Date: 20241220
Version: 0.91

#>

# Define output JSON file
$outputFile = "$PSScriptRoot\KnownGUIDs-MSIApps.json"

# Define registry paths to search
$uninstallPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

# Initialize an array to store result objects
$results = @()

foreach ($path in $uninstallPaths) {
    # Get all subkeys under the uninstall paths that look like {GUID}
    $subKeys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '{[0-9A-Fa-f\-]+}' }
    
    foreach ($subKey in $subKeys) {
        # Extract properties
        $props = Get-ItemProperty -Path $subKey.PSPath -ErrorAction SilentlyContinue
        $displayName = $props.DisplayName
        $displayVersion = $props.DisplayVersion
        $guid = $subKey.Name.Split('\')[-1]

        # Remove curly brackets from GUID
        $guid = $guid -replace '[{}]', ''

        # Only proceed if we have a DisplayName
        if ($displayName) {
            # Check if displayVersion is not already in displayName
            if ($displayVersion -and (-not ($displayName -like "*$displayVersion*"))) {
                $displayName = "$displayName $displayVersion"
            }

            $results += [PSCustomObject]@{
                ID           = $guid
                displayName  = $displayName
                Category     = "MSI Application names"
            }
        }
    }
}

# Remove duplicates based on ID if any
$results = $results | Sort-Object ID -Unique

# Convert the array to JSON and write to file
$results | ConvertTo-Json -Depth 4 | Out-File $outputFile -Encoding UTF8

Write-Host "MSI GUID to DisplayName mapping saved to $outputFile"
