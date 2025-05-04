# 🚀 Get-WindowsTroubleshootingReportCommunity v1.0

![Version](https://img.shields.io/badge/version-1.0-blue.svg)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207.x-blue.svg)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

The **Ultimate Windows and Intune Troubleshooting Tool** for analyzing and visualizing Windows Event logs and structured log files.

Download: [📦 Get-WindowsTroubleshootingReportCommunity_v1.0.zip](./Get-WindowsTroubleshootingReportCommunity_v1.0.zip)

Join the community to contribute and share custom event detection rules for even better troubleshooting experiences.

---

## 📚 Table of Contents

- [What's New (v1.0)](#whats-new-v10)
- [Features](#features)
- [Usage](#usage)
- [Parameters](#parameters)
- [How It Works](#how-it-works)
- [Contributing](#contributing)
- [Scenarios](#scenarios)
- [Example Reports](#example-reports)
- [PowerShell Script Parameters](#powershell-script-parameters)
- [Supported Log Formats](#supported-log-files-and-event-logs)
- [Admin Rights](#do-i-need-to-use-administrative-rights)
- [PowerShell Version Support](#windows-powershell-or-powershell-core)
- [License](#license)
- [Contributors](#contributors)
- [Acknowledgments](#acknowledgments-)

---

## 📢 What's New (v1.0)

**🎉 First Public Release! (Finally 😀)**

> Please report bugs or suggestions via [GitHub Issues](../../issues).

### Changelog Highlights

- LOTS of new features throughout the script  
- New console UI look  
- Improved HTML report  
- Expanded out-of-box **KnownEvent** rules  
- Bugfix for KnownEvents in multi-log scenarios  

---

## 📸 Screenshots

> 📍 Add screenshots here to visually showcase the timeline, grid filtering, or known event highlights.

```markdown
![Timeline Example](./images/timeline-sample.png)
![Grid Filtering](./images/filter-example.png)
```

---

## 🎥 Video Demo (Coming Soon)

```markdown
[![Watch Demo](https://img.youtube.com/vi/yourvideoid/0.jpg)](https://www.youtube.com/watch?v=yourvideoid)
```

---

## 🛠️ Features

- **Event Log Support**: Read Windows Event logs from live systems or Intune diagnostics packages (.zip).
- **Log File Support**: Analyze structured log files containing dateTime and message fields.
- **Unified Timeline**: Merge event and log file data chronologically.
- **Rich HTML Reports**: Real-time filtering and searchable HTML output.
- **Scenario-Based Troubleshooting**: Focused support for updates, installs, Intune, etc.
- **Custom Event Detection**: Detect specific events using customizable `EventRules.json`.
- **Community-Driven**: Share your rules and help others troubleshoot.

---

## ▶️ Usage Examples

```powershell
# Get known events from last 1 day
./Get-WindowsTroubleshootingReportCommunity.ps1 -LastDays 1

# Full capture from last 5 minutes
./Get-WindowsTroubleshootingReportCommunity.ps1 -LastMinutes 5 -AllEvents

# Events 2 minutes before and after last reboot
./Get-WindowsTroubleshootingReportCommunity.ps1 -MinutesBeforeLastBoot 2 -MinutesAfterLastBoot 2 -AllEvents

# From folder with known events only
./Get-WindowsTroubleshootingReportCommunity.ps1 -LogFilesFolder "C:\path\to\logs" -LastDays 2

# All events between two times
./Get-WindowsTroubleshootingReportCommunity.ps1 -LogFilesFolder "C:\path\to\logs" -AllEvents -StartTime "2024-09-01 00:00:00" -EndTime "2024-09-01 00:05:00"
```

---

## ⚙️ Parameters

- `-LogFile`: Single .evtx or log file.
- `-LogFilesFolder`: Folder with multiple .evtx/log files.
- `-LastMinutes`, `-LastHours`, `-LastDays`: Choose time range.
- `-MinutesBeforeLastBoot`, `-MinutesAfterLastBoot`: Boot window capture.
- `-StartTime`, `-EndTime`: Specify exact range.
- `-AllEvents`: Include all events (not just known).
- `-LogViewerUI`: Launch log review UI.
- `-RealtimeLogViewerUI`: Real-time log monitoring UI.
- `-IncludeSelectedKnownRulesCategoriesOnly`: Filter by known rules.
- `-ExcludeSelectedKnownRulesCategories`: Exclude rule categories.
- `-SortDescending`: Sort newest-first.
- `-ProcmonFilePath`: Include Procmon CSV file.

---

## 🧠 How It Works

1. **Log Processing**: Parses event logs and .log files, combines them into timeline.
2. **Report Generation**: Builds interactive HTML report with filters and search.
3. **Known Event Detection**: Uses rules from `EventRules.json`.
4. **Customization**: Build your own rules with `Create-EventRules-GUI-HelperTool.ps1`.

---

## 🤝 Contributing

- Use `Create-EventRules-GUI-HelperTool.ps1` to create `EventRules.json`.
- Add green (success) and red (fail) markers.
- Submit via GitHub to help others.

---

## 🧪 Scenarios

- **Windows Update**: Status, success/fail, reboots.
- **Intune**: Enrollment, sync, script/app errors.
- **Defender for Endpoint**: AV definitions and scans.
- **Software Installs**: MSI/Store installs and issues.
- **System Events**: Restarts, sleep, driver changes.

---

## 📊 Example Reports

- **Timeline View**
- **Filtered Report**: Only known events or categories

---

## 📘 PowerShell Script Parameters

> Full parameter list is in the script comments or use `Get-Help`.

---

## 🗂️ Supported Log Files and Event Logs

- Full support for **.evtx** logs (online/offline)
- Support for **CMTrace-style .log** files
- Experimental support for **Procmon CSV**
- Future: CBS.log, DISM.log, .etl, etc.

---

## 🔐 Do I Need Admin Rights?

- Not required, but highly recommended for full access
- **Admin = better coverage**
- Offline logs don’t require admin

---

## 🧩 PowerShell Support

- ✅ Windows PowerShell 5.1
- ✅ PowerShell Core (7.x)

---

## 📄 License

MIT License — see [LICENSE.md](LICENSE.md)

---

## 👨‍💻 Contributors

- **Petri Paavola** – *Author* (Microsoft MVP - Windows and Intune)

---

## 🤖 Acknowledgments

Special thanks to **GPT-4 from OpenAI** for helping with documentation generation, text refactoring, and markdown polishing. AI helped us work faster, so we can troubleshoot better. 💡

---
