# Get-WindowsTroubleshootingReportCommunity ver 0.9

The **Ultimate Windows and Intune Troubleshooting Tool** for analyzing and visualizing Windows Event logs and log files. **Join the community** to contribute and share custom event detection rules for even better troubleshooting experiences.

### Note! This is still under development and you are the first ones to test this tool. Be nice and send feedback what is working and what is not.

### Changelog ver 0.9
  * Maybe we will release this in the beginning of 2025 :)
  * KnownEvents, EventLogs and Logfiles are selected in Out-GridView by default
  * Support for new Intune Custom Inventory log files in folder
    * **C:\Program Files\Microsoft Device Inventory Agent\Logs**
  * Suppport for Sysinternals Process Monitor (**procmon**) csv-export files
    * Be careful with this and do not add too big traces to report
  * **Report will rename all found and known GUIDs to real names**
    * Folder KnownGUIDs has **KnownGUIDS-foo.json** files which has at least **properties Id** (which is GUID) and property **name** and/or **displayname**
    * Report includes Attack Surface Reduction (ASR) rules GUIDs file
    * In the very near future you can for example download Intune App, Powershell and Remediation scripts GUIDs easily with a tool to help Intune logs troubleshooting
  * Fixed a bug where KnownEvents for multiple .log files worked for only 1 log file for specific log. Rest of log files failed.

### Changelog ver 0.8
  * One step closer to public release :)
  * Known Events are shown in Out-GridView
  * EventLogs and log files are shown in Out-GridView
  * HTML report has lot of features for filtering and search
  * **Create-EventRules-GUI-HelperTool.ps1** creates .json to EventRules folder 


Yes, there are lot's of thing to do to make this perfect but we'll get there some day :)

**And please share EventRules you have created**

Download script package [Get-WindowsTroubleshootingReportCommunity_v0.9.zip](./Get-WindowsTroubleshootingReportCommunity_v0.9.zip)

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
Run report from local computer. Get known events from last 1 days. This is smaller report
```powershell
./Get-WindowsTroubleshootingReportCommunity.ps1 -LastDays 1
```
Run report from local computer from last 5 minutes. Capture all events and logs. This is good command to run after some specific action taken for example Intune Sync command
```powershell
./Get-WindowsTroubleshootingReportCommunity.ps1 -LastMinutes 5 -AllEvents
```
Run report from local computer. Capture logs and event from 2 minutes before last reboot and 2 minutes after last reboot. -AllEvents means that tool gathers every event found to report. This is bigger report but it shows everything
```powershell
./Get-WindowsTroubleshootingReportCommunity.ps1 -MinutesBeforeLastBoot 2 -MinutesAfterLastBoot 2 -AllEvents
```
Run report from saved log files from last 2 days. Show only known events. Remember that counting starts from current time
```powershell
./Get-WindowsTroubleshootingReportCommunity.ps1 -LogFilesFolder "C:\path\to\logs" -LastDays 2
```
Run report from saved log files. Use specified StartTime and Endtime and capture all events found. Because of -AllEvents this is limited to 5 minutes in this example command
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
- **`-ProcmonFilePath`**: Include Sysinternals Process Monitor (procmon) CSV-export file

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

## Supported log files and Event logs

- Script supports all Event Logs in **.evtx** -format either online or offline
- .log files support for now is for **CMTrace** type log files found in Intune Management Extension and ConfigMgr logs

Future support
- .etl files. These are used in few places. Get-WinEvent will not show data but there are ways to get data to clear text
- other .log files formats. Any structured .log file which has DateTime and message can be read by this tool
  - support for other log files are created by need. For example dism.log, CBS.log and other logs are probably coming in the future
  - goal is to make support for as many log file formats as possible

---

## Do I need to use Administrative rights

- Tool works without admin rights when running on local machine, **but** normal user can't access all log files
- Use Admin rights to get best result for local computer
- offline reporting for example from Diagnostics-package does **not** require admin rights

---

## Windows PowerShell or Powershell (core)

- This tool supports both Windows Powershell and Powershell (core)
- development is mainly done with PowerShell (core) but both are tested

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
