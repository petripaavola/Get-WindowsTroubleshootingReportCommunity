# Get-WindowsTroubleshootingReportCommunity ver 0.7

The **Ultimate Windows and Intune Troubleshooting Tool** for analyzing and visualizing Windows Event logs and log files. **Join the community** to contribute and share custom event detection rules for even better troubleshooting experiences.

### Note! This is still under development and you are the first ones to test this tool. Be nice and send feedback what is working and what is not.

Yes, there are lot's of thing to do to make this perfect but we'll get there some day :)

## Features
- **Event Log Support**: Reads Windows Event logs either from live systems or from diagnostics packages (.zip) downloaded via Intune.
- **Log File Support**: Reads any structured log file format, provided it contains dateTime and message information.
- **Unified Timeline**: Combines Event log events and log file entries into a chronological timeline for easier troubleshooting.
- **Rich HTML Reports**: Generates interactive HTML reports with real-time filtering, allowing you to search by event category, source, or custom text.
- **Scenario-Based Troubleshooting**: Focus on specific cases such as update failures, software installations, or Intune-related events.
- **Custom Event Detection**: Uses an `EventRules.json` configuration file to detect specific known events (both successes and failures) and present them visually with color-coded indicators.
- **Community-Driven**: Share your custom event rules with the IT community and help others troubleshoot effectively.

---

## Usage

```powershell
./Get-WindowsTroubleshootingReportCommunity.ps1 -LastDays 1
```
```powershell
./Get-WindowsTroubleshootingReportCommunity.ps1 -LastDays 30 -IncludeSelectedKnownRulesCategoriesOnly 'Updates - Install','Application installation - MSI','Power management - Start&Shutdown'
```
```powershell
./Get-WindowsTroubleshootingReportCommunity.ps1 -LastDays 180 -IncludeSelectedKnownRulesCategoriesOnly 'Updates - Install'
```
```powershell
./Get-WindowsTroubleshootingReportCommunity.ps1 -MinutesBeforeLastBoot 2 -MinutesAfterLastBoot 2 -AllEvents
```
```powershell
./Get-WindowsTroubleshootingReportCommunity.ps1 -LastMinutes 5 -AllEvents
```
```powershell
./Get-WindowsTroubleshootingReportCommunity.ps1 -LogFilesFolder "C:\path\to\logs" -LastDays 2
```
```powershell
./Get-WindowsTroubleshootingReportCommunity.ps1 -LogFilesFolder "C:\path\to\logs" -AllEvents -StartTime "2024-09-01 00:00:00" -EndTime "2024-09-01 00:05:00"
```


### Parameters:
- **`-LogFile`**: Path to a specific Windows Event (.evtx) or structured log file.
- **`-LogFilesFolder`**: Folder path containing multiple event or log files.
- **`-LastMinutes`**: Retrieve events/logs from the last X minutes.
- **`-LastHours`**: Retrieve events/logs from the last X hours.
- **`-LastDays`**: Retrieve events/logs from the last X days.
- **`-MinutesBeforeLastBoot`**: Retrieve logs from X minutes before the last boot.
- **`-MinutesAfterLastBoot`**: Retrieve logs from X minutes after the last boot.
- **`-StartTime`**: Specify the start time (e.g., `yyyy-MM-dd HH:mm:ss`) for retrieving events.
- **`-EndTime`**: Specify the end time for retrieving events.
- **`-AllEvents`**: Include all events in the report (not just known events).
- **`-LogViewerUI`**: Launch the tool with a user interface for reviewing logs.
- **`-RealtimeLogViewerUI`**: Launch a real-time log viewer UI.
- **`-IncludeSelectedKnownRulesCategoriesOnly`**: Filter events by selected known categories.
- **`-ExcludeSelectedKnownRulesCategories`**: Exclude events by selected known categories.
- **`-SortDescending`**: Sort the report with the most recent events first.

---

## How It Works
1. **Log Processing**: 
    - The tool reads Windows Event logs or structured log files, combining them into a single timeline.
    - It supports logs from live systems or diagnostics packages downloaded via Intune.
  
2. **Report Generation**:
    - Events are displayed chronologically in an interactive HTML report.
    - Real-time filtering allows users to search by event category, source, or free text.

3. **Known Event Detection**:
    - A separate configuration file, `EventRules.json`, contains the event rules for known issues.
    - The tool detects these events and categorizes them as either **successes (green)** or **failures (red)**.
  
4. **Customizable Event Rules**:
    - Users can create custom event rules by using the `Create-EventRules-GUI-HelperTool.ps1`.
    - These rules are shareable with the community to help others troubleshoot similar issues.
  
---

## Contributing
This tool is designed for the **IT Pro community** to share their knowledge and insights. By contributing your custom **EventRules**, you can help others troubleshoot more efficiently.

### Create Your Own Rules:
- Use the helper tool `Create-EventRules-GUI-HelperTool.ps1` to generate custom event rules.
- The tool allows you to browse available events, select the relevant ones, and categorize them with color-coded indicators (green for success, red for failure).
- Share your custom rules via GitHub so that others can benefit from your knowledge!

---

## Scenarios
- **Windows Update Troubleshooting**: See exactly when updates were installed, whether they were successful, and what errors occurred.
- **Defender for Endpoint**: Track antivirus signature updates and scheduled scans.
- **Intune Troubleshooting**: Review enrollment events, sync issues, policy successes or failures, and more.
- **Application Installations**: Track MSI and store app installations, including success/failure statuses.
- **System Events**: Check for restarts, power mode changes, and much more.

---

## Example Reports

Here are a few example reports that show the capabilities of the tool:

- **Timeline view**: A full chronological view of all events and logs.
- **Filtered report**: Displaying only known events relevant to the issue you're investigating.
  
---

## PowerShell Script Parameters

- Full parameter list available in the script comments.
  
---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE.md) file for details.

---

## Contributors

- **Petri Paavola** – *Author and Creator* (Microsoft MVP - Windows and Intune)

Feel free to contribute by creating and sharing **EventRules.json** files to help the IT community troubleshoot faster and more efficiently!

---
## Acknowledgments 🤖

A hearty thank you to **GPT-4 from OpenAI**, the ever-diligent AI co-pilot, for stepping in once again to help craft this documentation. From organizing intricate PowerShell parameters to ensuring every log file has its rightful place, GPT-4 was always on point. 

Oh, and let's not forget—this acknowledgment? Yup, 100% GPT-4. From timelines to text, AI not only builds, it celebrates too. 🚀

---
