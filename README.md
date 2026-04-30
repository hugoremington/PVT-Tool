# PVT Tool

**PVT Tool** is a lightweight application for remote post verification testing and health monitoring of Windows servers.

It performs essential pre- and post-verification tests, including:
* **Network:** DNS, Ping, and RDP connectivity.
* **System:** OS drive status, System Uptime, and Boot Time.
* **Services:** Service state and SMBv1 status.
* **Hardware:** vCPU and Memory details.

This tool is essential for PVT (Post-Verification Testing) during server patching, disaster recovery (DR), ITIL change management, and general system testing.

---

## Instructions

1. **Single Target:** Enter a single target computer name in the first field and click **Run**.
2. **Bulk Scanning:** Click **Browse** to select a text file containing multiple computer names, then click **Run**.
3. **Results:** When the scan is complete, the PVT report will automatically pop up.
4. **Exporting:** Click `File` $\rightarrow$ `Save As` to save the report as an **Excel CSV**.
5. **Copy/Paste:** (Optional) Select desired rows and press `CTRL + C` to copy data directly into Excel.

> [!NOTE]  
> Designed for **Microsoft Windows domain systems**. May also function on Linux systems.

---

---

## To Do

1. Incorporate all system disk(s) scanning.
2. Make system disk reporting dynamic, creating columns as required. Change from static.

> [!NOTE]  
> These fixes were incorporated in the proprietary edition of PVT Tools back in the. I will need to add it in the public release.

---

## Changelog

### v1.6.8
- Codesigned compiled file using Sectigo certificate.
- Added clipboard copy/paste feature: select rows and use `CTRL + C` to paste into Excel.
- Updated OS information capture by replacing WMI calls with **CIM**.

### v1.6.5
- Fixed system boot time values; optimized by recycling existing values and `Invoke-Command` to reduce overhead.

### v1.6.4
- Fixed **About** page.
- Removed redundant 2nd form `runspace.close` calls.

### v1.6.3
- Fixed exit bug where runspaces would keep the process open; app now exits cleanly.

### v1.6.2
- Removed `$script:powershell = [powershell]::Create()` from line 1033.
- Added **System Uptime/Boot Time** feature.

### v1.5.9
- Improved reliability: Appended `| Wait-Job -Timeout 11` after every `Invoke-Command` to prevent hanging on non-responsive WinRM servers.

### v1.5.8
- Attempted to resolve potential unprotected memory errors by calling `$script:powershell.EndInvoke($script:handle)` at every exit function.
- Applied various GUI fixes, including anchoring.

### v1.5.7
- Fixed views; enabled **Search Filter** (Textbox1) to appear correctly in Tab view.
- Fixed `Datagridview` and Tab control sizing (horizontal and vertical scrollbars now appear).

### v1.5.6
- Removed **Save File Confirm-Overwrite** feature (resolved crash on Server 2012).
- Resolved missing scrollbars in both Tab control and Data Grid View.

### v1.5.4
- Fixed Runspace garbage collection on completion and exit.

### v1.5.3
- Improved logic to properly close runspaces following successful completion.

### v1.5.2
- Attempted to suppress `Test-NetConnection` progress output.

### v1.5.1
- **Compatibility:** Made runspace sessions compatible with PowerShell versions older than 5.1 (tested on PS 4.0).
- **Performance:** GUI no longer freezes; implemented multi-threading via **Runspace Pools** optimized for the system's processor count.
- **UI:** Updated UI slightly and added a manual **Save As** feature.

### v1.5.0
- Major bug fixes and feature additions.

### v1.4.0
- Implemented **Runspace Pools**.
- Minor bug fixes (missing variables, etc.).
- Added `Add-OutputBox` function in runspace code.

### v1.3.3
- Updated 2nd GUI with a new color scheme, search filter bar, and close button.
- Optimized 2nd form sizing.
- Added a 2nd form (Tab control and Data Grid View) at the end of the `RunAppCode` function.

### v1.3.1
- Switched to `New-Object System.Data.DataTable` instead of traditional PS arrays.
- Added a "Xmas" easter egg.

### v1.3.0
- Added a **Progress Bar**.
- Added **WinRM check** (continues if running, stops if not).
- Implemented **Fast PING** in `$script1`.
- Added FQDN support within runspaces.
- Fixed runspace issues via `AddScript` within `ForEach` loops.

### v1.2.9a (Beta)
- Implementation of Runspaces (Phase 1).

### v1.2.8e
- Added primitive loading progress using `.` in output windows.
- Grayed out **Browse** button during code execution.
- Improved job handling to hide windows.
- Optimized scripts by using `-AsJob` with `Invoke-Command`.

### v1.2.7
- **Major Update:**
    - Added single computer/server PVT feature via textbox.
    - Added **SMBv1 status checks** (Windows Server 2012 R2 and up).
    - Added version detection and output box colorization.
    - Added NIC DNS server configuration feature.
    - Improved RDP port logic to improve overall app performance.
    - Added vCPU and Memory features.
    - Improved Ping function for Windows Desktop OS compatibility.
    - Fixed drive free space reporting on unreachable hosts.
    - Added IP Address, Default Gateway, and Subnet Mask details.
    - Added C:\ Drive capacity and free percentage reporting.
    - Codesigned and timestamped (10-Sep-2021).

---

## Author
**Hugo Remington**