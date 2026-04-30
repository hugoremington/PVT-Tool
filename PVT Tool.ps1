$Author = "Hugo Remington"
$AuthorDate = "20/06/2022."
$Version = "1.6.8"
<#Description: A Windows verification testing utility for bulk servers from a text file.

Changelog
===
1.6.8 - 20-06-2022 - Added feature; new column Last Patch Date, which probes remote server for the last patch date.
1.6.7 - 09/05/2022 - Added feature; new columns SSH Test and Telnet Test. Now supporting Unix/Linux systems!
Added new Test-Connection methods for performing SSH and Telnet port tests. This is to enable Non-Windows testing such as Unix, Linux and other devices.
Omitted "RDP Port" column and variables as this is useless info. Value is static "3389".
1.6.6 Added feature, new columns System Manufacturer and System Model. This is to determine physical/virtual server.
Added clipboard copy/paste feature by changing runspace apartmentstate to STA. You can now select rows and perform CTRL + C if you wish to copy/paste into Excel.
Updated operating system information capture by replacing WMI calls with CIM.
1.6.5 Fixed system boot time values. Now performing this more efficiently by recycling existing values and invoke-commands, thus removed former overhead of non-working code.
1.6.4 Fixed about page.
Removed 2nd form runspace.close calls as they are irrelevant.
1.6.3 Fixed exit bug where runspaces would keep the process open. App now exits cleanly.
1.6.2 removed $script:powershell = [powershell]::Create() from line 1033.
Added system uptime/boot time feature.
#1.5.9 Appended "| Wait-Job -Timeout 11" after every invoke-command as job for data collection. This should fix the app from hanging on non-responsive WinRM servers.
#1.5.8 attempt to debug potential unprotected memory error by calling $script:powershell.EndInvoke($script:handle) at every exit function.
#Other GUI fixes including anchoring.
#1.5.7 Fixed views, enabling search filter Textbox1 to appear correctly in Tab view.
#Fixed Datagridview and tab control sizing, both horizontal and vertical scroll bars now appearing.
#1.5.6 Removed save file Confirm-Overwrite feature as this is crashing on Server 2012.
#Attempting to fix missing scroll bar in tab control take #2
#Attempting to fix missing scrollbar in data grid view.
#1.5.4 fixed Runspace garbage collection on completion and exit.
#1.5.3 Attempt to properly close runspaces following successful completion.
#1.5.2 Attempt to suppress Test-NetConnection progress on like 458.
#1.5.1 Made runspace sessions on line 954 compatible with older than PS 5.1 versions. Tested and working on PS 4.0.
#GUI is no longer frozen!
#App is now multi-threaded thanks to Runspace pools optimised for number of processorts available on your system.
#Updated UI slightly.
#Added new manual Save As feature!
#1.5.0 Lots of bugs and features added.
#1.4.0 Runspace pools are working!
#Other minor bug fixes such as missing variables, etc.
#Added Add-OutputBox function in runspace code.
#Runspace is now working correctly. Next update will be the implementation of runspace pools.
$script:runspace is working and updating the GUI by using the Run add_click button!
#Removed Out-Gridview and now using native GUI!
#1.3.3 Updated 2nd GUI, new colour scheme, new search filter bar and new close button.
#Optimised 2nd form sizing.
#Added 2nd form at the end of RunAppCode function which displays a GUI containing tab control and data grid view.
#1.3.1 Added the use of New-Object System.Data.DataTable instead of traditional PS array.
#Added Xmas easter egg.
#Added progress bar, bringing the app to v1.3.0!
#Progress bar incrementer is on line ~420.
#Add WinRM check. Continue if running, else stop the script.
#Fast PING in $script1.
#Attempt to add FQDN into runspace.
#Fixed runspace issues by add AddScript within Foreach loop.
#A good beta version. Might make final.
#Rounded off disk capacity numbers to 2 decimals.
#Improved runspaces. No more Test-NetConnection GUI. Need to monitor memory performance and address any leaks in any.
#1.2.9a (Beta) Implementing runspaces take #1.
#1.2.8e Added primitive loading progress using '.' in output windows.
#Greyed out Browse button during code execution.
#Emtpying running jobs in RunCode function, outside of foreach loop also in an attempt to hide windows.
#Removed Start-Job brackets and appended -AsJob at the end on Invoke-Command scripts.
#Other minor improvements including cosmetic.
#This version of PVT Tools is memory efficient and intended to be run on older systems.
#Fixed server input textbox watermark validation.
#Added garbage collection function RunAppCode.
#Fixed PowerShell requirement message and check from version 5.1 to 4.0.
#Minor cosmetic fixes to GUI including browse box normal colourisation upon successful file load.
#1.2.7 Added single computer/server PVT feature using textbox.
#Improved menu slightly, adding colour.
#Updated icon.
#Further tweaks with emptying variables.
#Added SMBv1 status checks for Windows Server 2012 R2 and up.
#Fixed memory leak in Foreach loop failing to reset variables.
#Added version detection.
#Added output box colourisation.
#Added extra checks to prevent unhandled exceptions.
#Added browse file path display in text box.
#Made text boxes read only.
#Codesigned and timestamped, 10-Sep-2021.
#Fixed serialisation of IP address details for successful CSV export.
#Added NIC DNS server configuration feature.
#Check RDP port logic prior performing an additional Ping test, further improving app performance.
#Empty array following successful task completion.
#Improved Ping function, bringing Windows Desktop OS compatibility.
#Fixed drive free space reporting on unreachable hosts.
#Added vCPU and memory feature.
#Minor bug fixes. Removed Ping reply in ms as was not functioning correctly with multi-threading.
#Resolved multi-thread processing using start-job functions to comply with GUI initiative.
#Major GUI improvements. Added output text window. Removed Powershell native outputs.
#Minor GUI improvements. Added Browse button.
#Reworded success criteria to Pass/FAIL.
#Fixed application termination on cancel exit code.
#Added running services count to PVT.
#Added IP Address, Default Gateway and Subnet Mask details.
#Added C:\ Drive free percentage.
#Added C:\ Drive disk capacity report.
#Adjoined dual array results into single pane of glass.
#Added major GUI feature.
#Added Export to CSV feature to GUI, simply select and click OK.
#>
#Suppress Errors
$script:ErrorActionPreference = 'SilentlyContinue'
$script:ProgressPreference = 'SilentlyContinue'
#Set-ExecutionPolicy unrestricted
$getPowerShellVersion = $PSVersionTable.PSVersion

#Hash table for runspaces
$hash = [hashtable]::Synchronized(@{})

#Collect meta data for runspace
$hash.Author = $Author
$hash.AuthorDate = $AuthorDate
$hash.Version = $Version

#Desktop path
$DesktopPath = [Environment]::GetFolderPath("Desktop")

#Working Path
$script:workingPath = Get-Location

#Date
$script:datestring = (Get-Date).ToString("s").Replace(":","-")


#FUNCIONS

#Append output to browse box display.
Function Add-BrowseBoxLine 
{
    Param ($Message)
    $browseBox.AppendText("`r`n$Message")
    $browseBox.Refresh()
    $script:browseBox.ScrollToCaret()
    $hash.Form.Refresh()
}

#Append output to text box display.
Function Add-OutputBoxLine 
{
    Param ($Message)
    $hash.outputBox.AppendText("$Message")
    $hash.outputBox.Refresh()
    $script:hash.OutputBox.ScrollToCaret()
    $script:hash.OutputBox.SelectionStart = $hash.outputBox.Text.Length
    $hash.outputBox.Selectioncolor = "WindowText"
    $hash.Form.Refresh()
}

#File browse function.
Function fileBrowser
{
    try {
        
        $serverList = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
        $serverList.Title = "Select a TEXT file containing a server list."
        $serverList.filter = "Txt (*.txt)| *.txt|Csv (*.csv)| *.csv"
        $script:inputfile = $serverList.ShowDialog()
        $script:serverList = Get-Content -Path $serverList.FileName
        $hash.serverList = $script:serverList
        $totalServers = Get-Content $serverList.FileName | Measure-Object
        $totalServers = $totalServers.Count
        
            #Empty browse box when we get file, then append file path from our selection.
            $browseBox.Text = ""
            $browseBox.ForeColor = "WindowText"
            Add-BrowseBoxLine -Message $serverList.FileName

            #Check to ensure more than 0 computer objects in text file.
            If($totalServers -le 0)
            {
                $hash.outputBox.Selectioncolor = "Red"
                Add-OutputBoxLine -Message "`r`nMust select a valid input file containing at least 1 computer. Click Browse and select a text/csv file containing a valid computer list."
            }
            else
            {
            $hash.outputBox.Selectioncolor = "Green"
            Add-OutputBoxLine -Message "`r`nTotal computers in file: $totalServers"
            $hash.outputBox.Selectioncolor = "Green"
            Add-OutputBoxLine -Message "`r`nFile loaded. Click RUN to begin."
            }

    }
    catch {
        $hash.outputBox.Selectioncolor = "Red"
        Add-OutputBoxLine -Message "`r`nMust enter a compuer name OR select a valid input file. Click Browse and select a text/csv file containing your computer list."

        }
    $hash.Form.Refresh()
}

#Function for testing runspaces and GUI updates
Function RunspaceTestAppCode{
    #$global:hash.outputBox.AppendText("This is outside of the runspace")

    $script:runspace = [runspacefactory]::CreateRunspace()
    $script:runspace.ApartmentState = "STA"
    $script:runspace.ThreadOptions = "ReuseThread"
    $script:powershell = [powershell]::Create()
    $script:powershell.Runspace = $script:runspace
    $script:runspace.Open()
    $script:runspace.SessionStateProxy.SetVariable("hash",$hash)
    
    
    $script:powershell.AddScript({

        #$global:hash.runspaceOutputbox = $hash.outputBox.Text
        $hash.outputBox.BeginInvoke([action]{$hash.outputbox.SelectionColor = "Green"})
        $hash.outputBox.BeginInvoke([action]{$hash.outputbox.AppendText("Value from runspace!!! Working son")})
    
    })
    $AsyncObject = $script:powershell.BeginInvoke()
    $script:powershell.EndInvoke($AsyncObject)


}


Function RunAppCode{

    $scriptRun = {
        #Suppress Errors
        $script:ErrorActionPreference = 'SilentlyContinue'
        $script:ProgressPreference = 'SilentlyContinue'

        #Collect metadata
        $Author = $hash.Author
        $AuthorDate = $hash.AuthorDate
        $Version = $hash.Version
        Function Add-OutputBoxLine 
        {
            Param ($Message)
            $hash.outputBox.AppendText("$Message")
            $hash.outputBox.Refresh()
            $hash.OutputBox.ScrollToCaret()
            $hash.OutputBox.SelectionStart = $hash.outputBox.Text.Length
            $hash.outputBox.Selectioncolor = "WindowText"
            $hash.Form.Refresh()
        }

        #DECLARE VARIABLES
        #Desktop path
        $DesktopPath = [Environment]::GetFolderPath("Desktop")
        #Date
        $datestring = (Get-Date).ToString("s").Replace(":","-")

        #Get the serverList file
        $serverList = $script:hash.serverList
        #Reset progress bar
        $hash.progressBar1.Value = 0


        #Grey out buttons during code execution.
        $hash.buttonRun.enabled = $false
        $hash.buttonBrowse.enabled = $false
        
        #Watermark text in single computer or FQDN search text box. This variable must match that of the one in the form.
        $WatermarkText = "Enter a computer name or FQDN."
        
        #Get a single computer to PVT via $hash.serverBox text form if a text file is not received.
        If(!$serverList)
        {
            [String]$serverList = $hash.serverBox.Text
            
            #Ensure single server input field is changed from default transperant text.
            If($serverList -eq $WatermarkText)
            {
                #Clear the text
                $serverList = ""
                #$hash.serverBox.ForeColor = 'WindowText'
            }
        }
        
        #Ensure server list text file is loaded before runnnig code.
        If(!$serverList)
        {
            $hash.outputBox.Selectioncolor = "Red"
            Add-OutputBoxLine -Message "`r`nMust enter a compuer name OR select a valid input file. Click Browse and select a text/csv file containing your computer list."
            #$hash.outputBox.BeginInvoke([action]{$hash.outputbox.SelectionColor = "Red"})
            #$hash.outputBox.BeginInvoke([action]{$hash.outputbox.AppendText("`r`nMust enter a compuer name OR select a valid input file. Click Browse and select a text/csv file containing your computer list.")})
            #$hash.outputbox.SelectionColor = "Red"
            #$hash.outputbox.AppendText("`r`nMust enter a compuer name OR select a valid input file. Click Browse and select a text/csv file containing your computer list.")
            
            $hash.buttonBrowse.enabled = $true
            $hash.buttonRun.enabled = $true
            return
        }
    
    <#Experimental new code for Data Grid View#>
    #Table
    ## - [ Section to initialize DataTable objects] - ##
    ## - Create DataTable:
    $table = New-Object System.Data.DataTable;

    ## - Defining DataTable object columns and rows properties:
    # - Column1 = "Computer Name".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = 'Computer Name';
    $table.Columns.Add($column);

    # - Column2 = "Operating System".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Operating System";
    $table.Columns.Add($column);

    # - Column3 = "CPU".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.Int32");
    $column.ColumnName = "CPU";
    $table.Columns.Add($column);

    # - Column4 = "RAM".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "RAM";
    $table.Columns.Add($column);

    # - Column5 = "DNS Test".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "DNS Test";
    $table.Columns.Add($column);

    #1.6.7 Removing RDP port from test results.
    <#
    # - Column6 = "RDP Port".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.Int32");
    $column.ColumnName = "RDP Port";
    $table.Columns.Add($column);
    #>

    # - Column7 = "RDP Test".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "RDP Test";
    $table.Columns.Add($column);

    # - Column8 = "Ping Test".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Ping Test";
    $table.Columns.Add($column);

    #1.6.7 patch code.

    # - Column9 = "SSH Test".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "SSH Test";
    $table.Columns.Add($column);

    # - Column10 = "Telnet Test".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Telnet Test";
    $table.Columns.Add($column);

    #1.6.7 patch code end.

    # - Column11 = "OS Drive".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "OS Drive";
    $table.Columns.Add($column);

    # - Column12 = "Free Space Percentage".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Free Space Percentage";
    $table.Columns.Add($column);

    # - Column13 = "Free Space in GB".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Free Space in GB";
    $table.Columns.Add($column);

    # - Column14 = "Total Size in GB".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Total Size in GB";
    $table.Columns.Add($column);

    # - Column15 = "Running Services".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.Int32");
    $column.ColumnName = "Running Services";
    $table.Columns.Add($column);

    # - Column16 = "IP Address".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "IP Address";
    $table.Columns.Add($column);

    # - Column17 = "Subnet Mask".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Subnet Mask";
    $table.Columns.Add($column);

    # - Column18 = "Default Gateway".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Default Gateway";
    $table.Columns.Add($column);

    # - Column19 = "DNS Server 1".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "DNS Server 1";
    $table.Columns.Add($column);

    # - Column20 = "DNS Server 2".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "DNS Server 2";
    $table.Columns.Add($column);

    # - Column21 = "SMBv1 Status".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "SMBv1 Status";
    $table.Columns.Add($column);

    # - Column22 = "Last Patch Date".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Last Patch Date";
    $table.Columns.Add($column);

    # - Column23 = "Boot time".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "Boot Time";
    $table.Columns.Add($column);

    # - Column24 = "System Manufacturer".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "System Manufacturer";
    $table.Columns.Add($column);

    # - Column25 = "System Model".
    $column = New-Object System.Data.DataColumn;
    $column.DataType = [System.Type]::GetType("System.String");
    $column.ColumnName = "System Model";
    $table.Columns.Add($column);



    #Declare variables.
    #$Array = @()
    $FQDN = $null
    $Object = $null
    $getPing = $null
    $getSSH = $null
    $getTelnet = $null
    $getDisk = $null
    $getNetwork = $null
    $getService = $null
    $getOS = $null
    $getSystem = $null


    #Begin foreach loop for all servers.
    Foreach($server in $serverList)
    {
        [System.Windows.Forms.Application]::DoEvents()
        
        #Calculate progress
        $i++
        [int]$progressCount = ($i/$serverList.Count)*100
        $hash.progressBar1.Value = $progressCount
        $hash.Form.Refresh()


        #Remove line spaces and white spaces from text input.
        $server = $server.Trim()
        $hash.outputBox.Selectioncolor = "Green"
        Add-OutputBoxLine -Message "`r`nProcessing $server."
        #$hash.outputbox.SelectionColor = "Green"
        #$hash.outputbox.AppendText("`r`nProcessing $server.")
        #$hash.outputBox.BeginInvoke([action]{$hash.outputbox.SelectionColor = "Green"})
        #$hash.outputBox.BeginInvoke([action]{$hash.outputbox.AppendText("`r`nProcessing $server.")})
            
        

        #Check FQDN for remote computer
        $FQDN = ([System.Net.Dns]::GetHostByName(("$server")))

        If(!$FQDN)
        {
            $hash.outputBox.Selectioncolor = "Red"
            Add-OutputBoxLine -Message "."


            $hash.outputBox.Selectioncolor = "Red"
            Add-OutputBoxLine -Message "`r`n$server does not exist or is unreachable."


            $hash.outputBox.Selectioncolor = "Red"
            Add-OutputBoxLine -Message "Fail."

        }

        
        #Probe remote computer
        Else
        {
            #Get RDP info
            $Object = Test-NetConnection -ComputerName $server -CommonTcpPort RDP -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            $getPing = $Object.PingSucceeded
            $hash.outputBox.Selectioncolor = "Green"
            Add-OutputBoxLine -Message "."


            #PING port test for legacy clients and Desktop EUC workstations if above test fails..
            If(!$getPing)
            {
                $getPing = Test-Connection $server -Quiet -Count 1
                $hash.outputBox.Selectioncolor = "Green"
                Add-OutputBoxLine -Message "."

            }

            #1.6.7 patch code.

            #SSH connection test
            $SSHTest = Test-NetConnection -ComputerName $server -Port 22 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            $getSSH = $SSHTest.TcpTestSucceeded
            $hash.outputBox.Selectioncolor = "Green"
            Add-OutputBoxLine -Message "."

            #Telnet connection test
            $TelnetTest = Test-NetConnection -ComputerName $server -Port 23 -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            $getTelnet = $TelnetTest.TcpTestSucceeded
            $hash.outputBox.Selectioncolor = "Green"
            Add-OutputBoxLine -Message "."

            #1.6.7 patch code end.
     
            #Check if WinRM is running.
            $WinRM = Get-Service "WinRM" -ComputerName $server
            $WinRMStatus = $WinRM.Status
            If($WinRMStatus -ne "Running")
            {
                $hash.outputBox.Selectioncolor = "Chocolate"
                Add-OutputBoxLine -Message "."


                $hash.outputBox.Selectioncolor = "Chocolate"
                Add-OutputBoxLine -Message "`r`nWinRM error on $server. Unable to retrieve all information. Check network or credentials."

            }
            Else
            {

                #Get C Drive health.
                $Job2 = Invoke-Command -ComputerName $server -ScriptBlock { Get-Volume -DriveLetter "C" } -AsJob | Wait-Job -Timeout 11
                $getDisk = Receive-Job $Job2 -Wait

                $hash.outputBox.Selectioncolor = "Green"
                Add-OutputBoxLine -Message "."

                
                #Get Network Config
                $Job3 = Invoke-Command -ComputerName $server -ScriptBlock { Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=TRUE | Select-Object -Property [a-z]* -ExcludeProperty IPX*,WINS*  } -AsJob | Wait-Job -Timeout 11
                $getNetwork = Receive-Job $Job3 -Wait

                $hash.outputBox.Selectioncolor = "Green"
                Add-OutputBoxLine -Message "."

                
                #Get running services.
                $Job4 = Invoke-Command -ComputerName $server -ScriptBlock { Get-Service | Where-Object {$_.Status -eq "Running"} } -AsJob | Wait-Job -Timeout 11
                $getService = Receive-Job $Job4 -Wait

                $hash.outputBox.Selectioncolor = "Green"
                Add-OutputBoxLine -Message "."


                #Get OS information.
                #EXPERIMENTAL
                $getOS = Get-CimInstance -ComputerName $server -Class CIM_OperatingSystem -ErrorAction SilentlyContinue
                #$Job5 = Invoke-Command -ComputerName $server -ScriptBlock { (Get-WMIObject win32_operatingsystem) } -AsJob | Wait-Job -Timeout 11
                #$getOS = Receive-Job $Job5 -Wait

                $hash.outputBox.Selectioncolor = "Green"
                Add-OutputBoxLine -Message "."


                #Get logical CPU count and memory.
                $Job6 = Invoke-Command -ComputerName $server -ScriptBlock { (Get-CimInstance Win32_ComputerSystem) } -AsJob | Wait-Job -Timeout 11
                $getSystem = Receive-Job $Job6 -Wait

                $hash.outputBox.Selectioncolor = "Green"
                Add-OutputBoxLine -Message "."


                #Get SMBv1 Status
                $Job7 = Invoke-Command -ComputerName $server -ScriptBlock { Get-WindowsOptionalFeature -Online -FeatureName smb1protocol } -AsJob | Wait-Job -Timeout 11
                $getSMBv1 = Receive-job $Job7 -Wait

                $hash.outputBox.Selectioncolor = "Green"
                Add-OutputBoxLine -Message "."

                #1.6.8 Get last patch date
                $getLastPatch = Get-WmiObject -ComputerName $server Win32_Quickfixengineering | Select-Object @{Name="InstalledOn";Expression={$_.InstalledOn -as [datetime]}} | Sort-Object -Property Installedon | select-object -property installedon -last 1
                $hash.outputBox.Selectioncolor = "Green"
                Add-OutputBoxLine -Message "."

                #Get boot time.
                #Commented this section out as we are retrieving these values above in the $getOS variable.
                #$Job8 = Invoke-Command -ComputerName $server -ScriptBlock { (Get-CimInstance Win32_OperatingSystem) } -AsJob | Wait-Job -Timeout 11
                #$getbootTime = Receive-Job $Job8 -Wait

                $hash.outputBox.Selectioncolor = "Green"
                Add-OutputBoxLine -Message "OK"


            } #Close WinRM Else condition.

        } #Close FQDN Else condition.

        #Get RDP connection test variables
        $RDPComputer = $Object.ComputerName
            #If DNS record is missing for the computer.
            if(!$RDPComputer)
            {
                $RDPComputer = "[DNS MISSING] $server"
            }

        #WinRM server name.
            If($WinRMStatus -ne "Running")
                {
                    $RDPComputer = "[WinRM Error] $server"
                }
        
        #1.6.7 Removing RDP port from test results.
        #$RDPPort = $Object.RemotePort
        $RDPAddress = $Object.RemoteAddress
        $RDPResult = $Object.TcpTestSucceeded
            if($RDPResult -eq $True)
            {
                $RDPResult = "Pass"
            }
            else
            {
                $RDPResult = "FAIL"
            }
        
        
        $PingResult = $getPing
            if($PingResult -eq $True)
            {
                $PingResult = "Pass"
            }
            else
            {
                $PingResult = "FAIL"
            }

        #1.6.7 patch code.
        $SSHResult = $getSSH
            if($SSHResult -eq $True)
            {
                $SSHResult = "Pass"
            }
            else
            {
                $SSHResult = "FAIL"
            }

        $TelnetResult = $getTelnet
            if($TelnetResult -eq $True)
            {
                $TelnetResult = "Pass"
            }
            else
            {
                $TelnetResult = "FAIL"
            }

        #1.6.7 patch end.


        #Drive Variables
        $DriveLetter = $getDisk.DriveLetter
        $DriveFreeSpace = $getDisk.SizeRemaining/1GB
        $DriveFreeSpace = [math]::Round($DriveFreeSpace,2)
        $DriveSize = $getDisk.Size/1GB
        $DriveSize = [math]::Round($DriveSize,2)
        #Ensure drive exists before calculating free space.
            if($DriveSize -ne 0)
            {
                $FreePercentage = $DriveFreeSpace / $DriveSize * 100
                $FreePercentage = [math]::Round($FreePercentage,2)
            }

        

        #Network Variables
        [string]$IPAddress = $getNetwork.IPAddress
        [string]$SubnetMask = $getNetwork.IPSubnet
        [string]$DefaultGateway = $getNetwork.DefaultIPGateway
        $DNSServerOrder = $getNetwork.DNSServerSearchOrder

        if(!$DNSServerOrder)
        {
            $DNSServer1 = ""
            $DNSServer2 = ""
        }
        else
        {
            $DNSServers = $DNSServerOrder.split()
            $DNSServer1 = $DNSServers[0]
            $DNSServer2 = $DNSServers[1]
        }

        #Service Variables
        $ServiceCount = $getService.Count
    
        #Get OS Version
        $OperatingSystem = $getOS.caption

        #Get logical CPU count and memory.
        $vCPU = $getSystem.NumberOfLogicalProcessors
        $vMem = $getSystem.TotalPhysicalMemory /1GB
        $vMem = [math]::Round($vMem,2)

        #Get manufacturer and model
        $Manufacturer = $getSystem.Manufacturer
        $Model = $getSystem.Model
                
        #Get SMBv1 status.
        $SMBv1Status = $getSMBv1.State

        #1.6.8 Get last patch date
        $LastPatch = Get-Date $getLastPatch.InstalledOn -Format dd/MM/yyyy

        #Get boot time
        [String]$bootTime = $getOS.LastBootUpTime
        
        #Add our results into the array.
        #$Array += [pscustomobject]@{'Computer Name'=$RDPComputer; 'Operating System'=$OperatingSystem; 'CPU'=$vCPU; 'RAM'=$vMem; 'DNS Test'=$RDPAddress; 'RDP Port'=$RDPPort; 'RDP Test'=$RDPResult; 'Ping Test'=$PingResult; 'OS Drive'=$DriveLetter; 'Free Space Percentage'=$FreePercentage; 'Free Space in GB'=$DriveFreeSpace; 'Total Size in GB'=$DriveSize; 'Running Services'=$ServiceCount; 'IP Address'=$IPAddress; 'Subnet Mask'=$SubnetMask; 'Default Gateway'=$DefaultGateway; 'DNS Server 1'=$DNSServer1; 'DNS Server 2'=$DNSServer2; 'SMBv1 Status'=$SMBv1Status}
        #Finished adding into Array.
        
        <#Experimental new code#>
        $row = $table.NewRow();
        $row["Computer Name"] = $RDPComputer;
        $row["Operating System"] = $OperatingSystem;
        $row["CPU"] = $vCPU;
        $row["RAM"] = "$vMem GB";
        $row["DNS Test"] = $RDPAddress;
        #1.6.7 Removing RDP port from test results.
        #$row["RDP Port"] = $RDPPort;
        $row["RDP Test"] = $RDPResult;
        $row["Ping Test"] = $PingResult;
        $row["SSH Test"] = $SSHResult;
        $row["Telnet Test"] = $TelnetResult;
        $row["OS Drive"] = $DriveLetter;
        $row["Free Space Percentage"] = "$FreePercentage %";
        $row["Free Space in GB"] = "$DriveFreeSpace GB";
        $row["Total Size in GB"] = "$DriveSize GB";
        $row["Running Services"] = $ServiceCount;
        $row["IP Address"] = $IPAddress;
        $row["Subnet Mask"] = $SubnetMask;
        $row["Default Gateway"] = $DefaultGateway;
        $row["DNS Server 1"] = $DNSServer1;
        $row["DNS Server 2"] = $DNSServer2;
        $row["SMBv1 Status"] = $SMBv1Status;
        $row["Last Patch Date"] = $LastPatch;
        $row["Boot Time"] = $bootTime;
        $row["System Manufacturer"] = $Manufacturer;
        $row["System Model"] = $Model;
        $table.Rows.Add($row)
        
        #$hash.outputbox.SelectionColor = "Green"
        #$hash.outputbox.AppendText("OK")

        <#Close experimental new code#>
        

        #GARBAGE COLLECTION
        $FQDN = $null
        $Object = $null
        $getPing = $null
        $getSSH = $null
        $getTelnet = $null
        $getDisk = $null
        $DriveSize = $null
        $DriveLetter = $null
        $FreePercentage = $null
        $DriveFreeSpace = $null
        $getNetwork = $null
        $getService = $null
        $getOS = $null
        $OperatingSystem = $null
        $getSystem = $null
        $ServiceCount = $null
        $getSMBv1 = $null
        $getLastPatch = $null
        $LastPatch = $null
        $SMBv1Status = $null
        $PingResult = $null
        $RDPResult = $null
        #$RDPPort = $null
        $RDPAddress = $null
        $RDPComputer = $null
        $vMem = $null
        $vCPU = $null
        $DNSServerOrder = $null
        $DNSServers = $null
        $DNSServer1 = $null
        $DNSServer2 = $null
        $DefaultGateway = $null
        $SubnetMask = $null
        $IPAddress = $null
        $bootTime = $null
        $Manufacturer = $null
        $Model = $null
        
        #Garbage collector
        Get-Job | Remove-Job -Force

    } #Closing the Foreach loop.
    


    #Textbox completion output.
    Add-OutputBoxLine -Message "`r`n=============="
    Add-OutputBoxLine -Message "`r`nPVT Complete."
    Add-OutputBoxLine -Message "`r`n=============="
    Add-OutputBoxLine -Message "`r`nCSV report saved on Desktop.`n$DesktopPath\PVT_Results - $datestring.csv"
    <#$hash.outputbox.SelectionColor = "WindowText"
    $hash.outputbox.AppendText("`r`n==============")
    $hash.outputbox.AppendText("`r`nPVT Complete.")
    $hash.outputbox.AppendText("`r`n==============")
    $hash.outputbox.AppendText("`r`nCSV report saved on Desktop.`n$DesktopPath\PVT_Results - $datestring.csv")#>

    $table | Export-Csv -Path "$DesktopPath\PVT_Results - $datestring.csv" -NoTypeInformation
    #$Array | Out-Gridview -Title "PVT Results" -PassThru
    

    <#Experimental new code for Datagridview#>
    #Display second form.
    $Form2 = New-Object system.Windows.Forms.Form
    $Form2Width = '800'
    $Form2Height = '500'
    $Form2.MinimumSize = "$Form2Width,$Form2Height"
    $Form2.StartPosition = 'CenterScreen'
    $Form2.text = "PVT Results"
    $Form2.Icon = [System.Drawing.SystemIcons]::Shield
    #Autoscaling settings
    $Form2.AutoScale = $true
    $Form2.AutoScaleMode = "Font"
    $ASsize = New-Object System.Drawing.SizeF(7,15)
    $Form2.AutoScaleDimensions = $ASsize
    $Form2.BackColor = 'SteelBlue'
    $Form2.Refresh()
    #Disable windows maximize feature.
    #$Form2.MaximizeBox = $False
    #$Form2.FormBorderStyle='FixedDialog'

    $DataGridView1 = New-Object system.Windows.Forms.DataGridView
    $DataGridView1.DataSource = $table
    $DataGridView1.Anchor = 'Top, Bottom, Left, Right'
    $DataGridView1.AutoSizeColumnsMode = 'AllCells'
    $DataGridView1.ColumnHeadersHeightSizeMode = 'AutoSize'
    $DataGridView1.width = 730
    $DataGridView1.height = 350
    #$DataGridView1.width = 710
    #$DataGridView1.height = 460
    $DataGridView1.location = New-Object System.Drawing.Point(12,14)
    $DataGridView1.AutoSize = $True
    $DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = "LightBlue"
    $DataGridView1.clipboardcopymode = "EnableAlwaysIncludeHeaderText"

    #Search filter experimental
        $textbox1_TextChanged = {
        $dataGridView1.Refresh()
        $filter = $textbox1.Text
        If($filter -eq $WatermarkText1)
        {
            $filter = ''
        }

        #$datagridview1.DataSource.DefaultView.RowFilter = "[Computer Name] LIKE '*$($textbox1.Text)*' OR [RDP Test] LIKE '*$($textbox1.Text)*' OR [Ping Test] LIKE '*$($textbox1.Text)*'"
        #$datagridview1.DataSource.DefaultView.RowFilter = "[Computer Name] LIKE '*$($textbox1.Text)*' OR [Operating System] LIKE '*$($textbox1.Text)*' OR [CPU] LIKE '*$($textbox1.Text)*' OR [RAM] LIKE '*$($textbox1.Text)*' OR [DNS Test] LIKE '*$($textbox1.Text)*' OR [RDP Port] LIKE '*$($textbox1.Text)*' OR [RDP Test] LIKE '*$($textbox1.Text)*' OR [Ping Test] LIKE '*$($textbox1.Text)*' OR [OS Drive] LIKE '*$($textbox1.Text)*' OR [Free Space in GB] LIKE '*$($textbox1.Text)*' OR [Total Size in GB] LIKE '*$($textbox1.Text)*' OR [Running Services] LIKE '*$($textbox1.Text)*' OR [IP Address] LIKE '*$($textbox1.Text)*' OR [Subnet Mask] LIKE '*$($textbox1.Text)*' OR [Default Gateway] LIKE '*$($textbox1.Text)*' OR [DNS Server 1] LIKE '*$($textbox1.Text)*' OR [DNS Server 2] LIKE '*$($textbox1.Text)*' OR [SMBv1 Status] LIKE '*$($textbox1.Text)*'"
        $datagridview1.DataSource.DefaultView.RowFilter = "[Computer Name] + [RDP Test] + [Ping Test] LIKE '*$($filter)*'"
    }
    


     

    $MainTab = New-Object System.Windows.Forms.TabControl
    $MainTab.Size = '752,370'
    #$MainTab.Size = '755,390'
    #$MainTab.Location = '15,38'
    $MainTab.Location = '15,48'
    $MainTab.Multiline = $True
    $MainTab.Name = 'Main Tab'
    $MainTab.SelectedIndex = 0
    $MainTab.Anchor = 'Top,Left,Bottom,Right'

    $TabPage1 = New-Object System.Windows.Forms.TabPage
    $Tabpage1.Name = 'PVT Results'
    #$Tabpage1.Size = '600, 370'
    #$Tabpage1.Size = '700, 370'
    $Tabpage1.Size = '740, 360'
    $Tabpage1.Padding = '5,5,5,5'
    $Tabpage1.TabIndex = 1
    $Tabpage1.Text = 'PVT Results'
    $Tabpage1.UseVisualStyleBackColor = $True
    $Tabpage1.Anchor = 'Top,Left,Bottom,Right'
    #$Tabpage1.AutoScroll = $True
    $TabPage1.HorizontalScroll = $true
    $TabPage1.VerticalScroll = $True
    #$TabPage1.Enabled = $false
    $TabPage1.Controls.AddRange(@($DataGridView1))

    <# 2nd tab, if you want to add, be sure to add it in the $MainTab Controls.
    $TabPage2 = New-Object System.Windows.Forms.TabPage
    $Tabpage2.Name = '2nd Tab'
    $Tabpage2.Size = '500, 425'
    $Tabpage2.Padding = '5,5,5,5'
    $Tabpage2.TabIndex = 1
    $Tabpage2.Text = '2nd Tab'
    $Tabpage2.UseVisualStyleBackColor = $True
    $TabPage2.Enabled = $false
    #$TabPage2.Controls.AddRange(@($DataGridView1))
    #>

    #Close Button
    $buttonClose = New-Object System.Windows.Forms.Button
    $buttonClose.text = "Close"
    #$buttonClose.Size = '80,30'
    $buttonClose.Size = '80,30'
    $buttonClose.location = '690, 420'
    $buttonClose.Font = 'Verdana,9'
    $buttonClose.Anchor = 'Bottom,Right'
    $buttonClose.Add_Click({$Form2.Close()})

    #Search filter box
    $textbox1 = New-Object System.Windows.Forms.TextBox
    $textbox1.Location = '585, 31'
    $textbox1.anchor = 'Top,Right'
    $textbox1.Name = 'textbox1'
    $textbox1.Size = '180, 15'
    $textbox1.TabIndex = 1
    $textbox1.add_TextChanged($textbox1_TextChanged)
    #Watermark
    $WatermarkText1 = "Search"
    $textbox1.ForeColor = 'Gray'
    $textbox1.Text = $WatermarkText1
    #If we have focus then clear out the text
    $textbox1.Add_GotFocus(
        {
            If($textbox1.Text -eq $WatermarkText1)
            {
                $textbox1.Text = ''
                $textbox1.ForeColor = 'WindowText'
            }
        }
    )
    #If we have lost focus and the field is empty, reset back to watermark.
    $textbox1.Add_LostFocus(
        {
            If($textbox1.Text -eq '')
            {
                $textbox1.Text = $WatermarkText1
                $textbox1.ForeColor = 'Gray'
            }
        }
    )
    

    #ABOUT FORM
    $FormAbout = New-Object system.Windows.Forms.Form
    $FormAboutWidth = '800'
    $FormAboutHeight = '500'
    $FormAbout.MinimumSize = "$FormAboutWidth,$FormAboutHeight"
    $FormAbout.StartPosition = 'CenterScreen'
    $FormAbout.text = "About"
    $FormAbout.Icon = [System.Drawing.SystemIcons]::Shield
    #Autoscaling settings
    $FormAbout.AutoScale = $true
    $FormAbout.AutoScaleMode = "Font"
    $ASsize = New-Object System.Drawing.SizeF(7,15)
    $FormAbout.AutoScaleDimensions = $ASsize
    $FormAbout.BackColor = 'SteelBlue'
    $FormAbout.Refresh()
    #Disable windows maximize feature.
    $FormAbout.MaximizeBox = $False
    #$FormAbout.FormBorderStyle ='FixedDialog'

    $AboutHeading = New-Object system.Windows.Forms.Label
    $AboutHeading.text = "PVT Tool"
    $AboutHeading.AutoSize = $false
    $AboutHeading.width = 150
    $AboutHeading.height = 30
    $AboutHeading.location = New-Object System.Drawing.Point(20,20)
    $AboutHeading.Font = 'Verdana,14'
    $AboutHeading.Anchor = 'top, left'

    $AboutDescription = New-Object system.Windows.Forms.Label
    $AboutDescription.text = "Credits.`r`nAuthor: $Author Build date: $AuthorDate Version: $Version"
    $AboutDescription.AutoSize = $false
    $AboutDescription.width = 500
    $AboutDescription.height = 70
    $AboutDescription.location = New-Object System.Drawing.Point(20,410)
    $AboutDescription.Font = 'Verdana,10'
    $AboutDescription.Anchor = 'bottom, left'

    $Description2Heading = New-Object system.Windows.Forms.Label
    $Description2Heading.text = "Instructions"
    $Description2Heading.AutoSize = $false
    $Description2Heading.width = 150
    $Description2Heading.height = 30
    $Description2Heading.location = New-Object System.Drawing.Point(20,80)
    $Description2Heading.Font = 'Verdana,12'

    $AboutDescription2 = New-Object system.Windows.Forms.Label
    $AboutDescription2.text = "1. Enter a single target computer in the first field and click Run.`r`n2. For bulk scanning; click browse to select a text file contaning multiple computers and click Run.`r`n3. When scan is complete, the PVT report will automatically pop-up.`r`n4. You can then click File -> Save As to save the report to an Excel CSV.`r`n5. By selecting desired rows and performing CTRL + C, you can copy/paste data to Excel (optional).`r`n`r`nDesigned for Microsoft Windows domain systems."    
    $AboutDescription2.AutoSize = $false
    $AboutDescription2.width = 620
    $AboutDescription2.height = 200
    $AboutDescription2.location = New-Object System.Drawing.Point(20,120)
    $AboutDescription2.Font = 'Verdana,10'

    #About Close Button
    $aboutClose = New-Object System.Windows.Forms.Button
    $aboutClose.text = "Close"
    #$buttonClose.Size = '80,30'
    $aboutClose.Size = '80,30'
    $aboutClose.location = '350, 360'
    $aboutClose.Font = 'Verdana,9'
    $aboutClose.Anchor = 'Bottom,Left'
    $aboutClose.Add_Click({$FormAbout.Close()})

    $FormAbout.Controls.AddRange(@($AboutHeading, $AboutDescription, $Description2Heading, $AboutDescription2, $aboutClose))

    #Help Menu
    $helpAbout = New-Object System.Windows.Forms.ToolStripMenuItem
    $helpAbout.Name = "About"
    $helpAbout.Text = "About"
    $helpAbout.Add_Click({$FormAbout.ShowDialog()})

    $menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuHelp.Name = "Help"
    $menuHelp.Text = "Help"
    $menuHelp.DropDownItems.AddRange(@($helpAbout))
    
    
    #Main Menu Toolbar
    #File Menu
    $menuSaveAs = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuSaveAs.Name = "Save As"
    $menuSaveAs.Text = "Save As"
    $menuSaveAs.Add_Click({
        try {
            #Save dialogue
            $saveDlg = New-Object System.Windows.Forms.SaveFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
            $saveDlg.ShowHelp = $true
            $saveDlg.CreatePrompt = $false
            $saveDlg.OverwritePrompt = $false
            $saveDlg.RestoreDirectory = $true
            $saveDlg.filter = "Csv (*.csv)| *.csv|Txt (*.txt)| *.txt"
            #$saveDlg.filter = "Txt (*.txt)| *.txt|Csv (*.csv)| *.csv"
            #$saveDlg.ShowDialog()
            if($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $SaveFileName = $saveDlg.FileName
                $table | Export-CSV -Path $SaveFileName -NoTypeInformation
            }
        }
        catch {
            $hash.outputBox.Selectioncolor = "Red"
            Add-OutputBoxLine -Message "`r`nOoops. An error occurred when saving your report $SaveFileName. Please try again."
    
            }
    })

    $menuClose2 = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuClose2.Name = "Close"
    $menuClose2.Text = "Close"
    $menuClose2.Add_Click({$Form2.Close()})

    #File Menu continued
    $menuFile2 = New-Object System.Windows.Forms.ToolStripMenuItem
    $menuFile2.Name = "File"
    $menuFile2.Text = "File"
    $menuFile2.DropDownItems.AddRange(@($menuSaveAs, $menuClose2))



    #Main menu
    $menuMain2 = New-Object System.Windows.Forms.MenuStrip
    $menuMain2.Items.AddRange(@($menuFile2, $menuHelp))
    #Display the main tab GUI.
    $MainTab.Controls.AddRange(@($TabPage1))

    #Display the 2nd GUI once the command finishes executing. Contains above tab.
    #$script:Form2.Controls.AddRange(@($MainTab, $buttonClose))
    $script:Form2.Controls.AddRange(@($menuMain2, $mainTab, $textbox1, $buttonClose))
    $Form2.ShowDialog()

    <#Close off experimental code#>
 

    
    #Make main GUI button visible again.
    $hash.buttonBrowse.enabled = $true
    $hash.buttonRun.enabled = $true



    } #Close the $scriptRun brackets for the runspace
    
    #RUNSPACE
    #$script:runspace = [runspacefactory]::CreateRunspace()
    #$maxthreads = [int]$env:NUMBER_OF_PROCESSORS+1
    $maxthreads = [int]$env:NUMBER_OF_PROCESSORS
    #$session = [initialsessionstate]::CreateDefault2()
    #$session.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new('hash', $hash,''))
    $hashVars = New-object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList 'hash',$hash,$Null
    $InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    #Add the variable to the sessionstate
    $InitialSessionState.Variables.Add($hashVars)
    $script:runspace = [runspacefactory]::CreateRunspacePool(1,$maxthreads,$InitialSessionState, $Host)
    $script:runspace.ApartmentState = "STA"
    #$script:runspace = [runspacefactory]::CreateRunspacePool(1,$maxthreads,$session, $Host)

    
    #$script:runspace.Variables.Add
    $script:powershell = [powershell]::Create()
    $script:runspace.Open()
    
      
        
        
        #Begin our main code within the runspace
        $script:powershell.AddScript($scriptRun)
        #$hash.outputBox.AppendText("`r`nAdding script")
        $script:powershell.RunspacePool = $script:runspace
        
        #$script:handle = $script:powershell.BeginInvoke()
        $script:handle = $script:powershell.BeginInvoke()
        if ($script:handle.IsCompleted)
        {
            $script:powershell.EndInvoke($script:handle)
            #$script:powershell.Close()
            $script:powershell.Dispose()
            $script:runspace.Dispose()
            $script:runspace.Close()
            [System.GC]::Collect()
        }

        

 } #Closing the function.


#Menu GUI begins.

# Install .Net Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
 
# Enable Visual Styles
[Windows.Forms.Application]::EnableVisualStyles()


#Main Form
$hash.Form = New-Object system.Windows.Forms.Form
$FormWidth = '500'
$FormHeight = '690'
$hash.Form.Size = "$FormWidth,$FormHeight"
$hash.Form.StartPosition = 'CenterScreen'
$hash.Form.text = "PVT Tool $Version"
$hash.Form.Icon = [System.Drawing.SystemIcons]::Shield
$hash.Form.BackColor = 'SteelBlue'
$hash.Form.Refresh()
#Disable windows maximize feature.
$hash.Form.MaximizeBox = $False
$hash.Form.FormBorderStyle='FixedDialog'

#Xmas easter egg
$date = (get-date)
$year = (get-date).Year
$newYear = (get-date).AddYears(1).Year
$xmas = Get-Date -Month 12 -Day 25 -Year $year
$nye = Get-Date -Month 01 -Day 02 -Year $newYear
If(($date -ge $xmas) -and ($date -le $nye))
{
    $hash.Form.text = "PVT Tool $Version - Merry Christmas and a Happy New Year!"
    $hash.Form.BackColor = 'FireBrick'
    $hash.Form.Refresh()
}


$Description = New-Object system.Windows.Forms.Label
$Description.text = "PVT Tool"
$Description.AutoSize = $false
$Description.width = 450
$Description.height = 30
$Description.location = New-Object System.Drawing.Point(20,40)
$Description.Font = 'Verdana,13'
$Description.Anchor = 'top, left'

$moto = New-Object system.Windows.Forms.Label
$moto.text = "Perform verification testing of your Windows Server workloads."
$moto.AutoSize = $false
$moto.width = 450
$moto.height = 50
$moto.location = New-Object System.Drawing.Point(20,70)
$moto.Font = 'Verdana,10'

#Input single computer or FQDN textbox
$hash.serverBox = New-Object System.Windows.Forms.TextBox 
$hash.serverBox.Location = New-Object System.Drawing.Size(90,110) 
$hash.serverBox.Size = New-Object System.Drawing.Size(300,40)
$hash.serverBox.Font = 'Verdana,9'
$hash.serverBox.ReadOnly = $False
$WatermarkText = "Enter a computer name or FQDN."
$hash.serverBox.ForeColor = 'Gray'
$hash.serverBox.Anchor = 'top, left'
$hash.serverBox.Text = $WatermarkText
#If we have focus then clear out the text
$hash.serverBox.Add_GotFocus(
    {
        If($hash.serverBox.Text -eq $WatermarkText)
        {
            $hash.serverBox.Text = ''
            $hash.serverBox.ForeColor = 'WindowText'
        }
    }
)
#If we have lost focus and the field is empty, reset back to watermark.
$hash.serverBox.Add_LostFocus(
    {
        If($hash.serverBox.Text -eq '')
        {
            $hash.serverBox.Text = $WatermarkText
            $hash.serverBox.ForeColor = 'Gray'
        }
    }
)

$hash.serverBox.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        runAppCode
    }
})

#Browse text box displaying file path.
$browseBox = New-Object System.Windows.Forms.TextBox 
$browseBox.Location = New-Object System.Drawing.Size(89,152) 
$browseBox.Size = New-Object System.Drawing.Size(192,50)
$browseBox.Font = "Verdana, 9"
$browseBox.Anchor = 'top, left'
#Must configure BackColor in a ReadOnly textbox in order for ForeColor to work.
$browseBox.BackColor = 'White'
$browseBox.ForeColor = 'Gray'
$browseBox.ReadOnly = $True
$browseWatermark = "Click Browse to select file."
$browseBox.Text = $browseWatermark



<# Output Box which is below all other buttons and displays PS Output #>
$hash.outputBox = New-Object System.Windows.Forms.RichTextBox 
$hash.outputBox.Location = New-Object System.Drawing.Size(65,350) 
$hash.outputBox.Size = New-Object System.Drawing.Size(350,180)
$hash.outputBox.Font = "Verdana, 8"
$hash.outputBox.ReadOnly = $True
$hash.outputBox.MultiLine = $True
$hash.outputBox.ScrollBars = "Vertical"
$hash.outputBox.Anchor = 'top, left'
#$hash.outputBox.AppendText("PVT Tool ready.")
#$hash.Form.Controls.Add($hash.outputBox)

#PowerShell version check.
If($getPowerShellVersion -ge "4.0")
{
        $hash.outputBox.AppendText("Your PowerShell version is $getPowerShellVersion. `nPVT Tool is ready.")
}
else
{
        $hash.outputBox.Selectioncolor = "Red"
        $hash.outputBox.AppendText("Your PowerShell version is $getPowerShellVersion. `nPVT Tool may not run correctly on your computer.")
}

#Browse Button
$hash.buttonBrowse = New-Object System.Windows.Forms.Button
$hash.buttonBrowse.text = "Browse"
$hash.buttonBrowse.Size = '80,30'
$hash.buttonBrowse.location = '310, 147'
$hash.buttonBrowse.Font = 'Verdana,9'
$hash.buttonBrowse.Anchor = 'top, left'
$hash.buttonBrowse.Add_Click({fileBrowser})
#$hash.buttonBrowse.Cursor = [System.Windows.Forms.Cursors]::Hand
#$hash.buttonBrowse.DialogResult = [System.Windows.Forms.DialogResult]::Ok

$hash.buttonRun = New-Object System.Windows.Forms.Button
$hash.buttonRun.text = "RUN"
$hash.buttonRun.Size = '300,40'
$hash.buttonRun.location = '90, 200'
$hash.buttonRun.Font = 'Verdana,11'
$hash.buttonRun.BackColor = "CornflowerBlue"
$hash.buttonRun.Cursor = [System.Windows.Forms.Cursors]::Hand
$hash.buttonRun.Anchor = 'top, left'
$hash.buttonRun.Add_Click({RunAppCode})
#$hash.buttonRun.DialogResult = [System.Windows.Forms.DialogResult]::Ok

#Close button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = '190,270'
$exitButton.Size = '100,40'
$exitButton.FlatStyle = 'Flat'
$exitButton.BackColor = 'Brown'
$exitButton.Font = 'Verdana, 9'
$exitButton.Anchor = 'top, left'

# Font styles are: Regular, Bold, Italic, Underline, Strikeout
#$exitButton.Font = $cancelFont
$hash.Form.Controls.Add($exitButton)
$exitButton.Text = 'Exit'
$exitButton.tabindex = 0
$exitButton.Add_Click({
    $hash.Form.Tag = $hash.Form.close()
    $script:powershell.EndInvoke($script:handle)
    #$script:powershell.Close()
    $script:powershell.Dispose()
    $script:runspace.Close()
    $script:runspace.Dispose()
    [System.GC]::Collect()
    })
$hash.Form.CancelButton = $exitButton


#Progress status Heading
$StatusHeading = New-Object system.Windows.Forms.Label
$StatusHeading.text = "Progress: "
$StatusHeading.AutoSize = $false
$StatusHeading.width = 80
$StatusHeading.height = 20
$StatusHeading.location = New-Object System.Drawing.Point(70,558)
$StatusHeading.Anchor = 'bottom, left'
$StatusHeading.Font = 'Verdana,10'

#Progress bar.
$hash.progressBar1 = New-Object System.Windows.Forms.ProgressBar
$hash.progressBar1.Name = 'progressBar1'
$hash.progressBar1.Value = 0
$hash.progressBar1.Style="Blocks"
$hash.progressBar1.Size = "254, 30"
$hash.progressBar1.location = '160, 550'
$hash.progressBar1.Anchor = 'bottom, left'
#$hash.progressBar1.location = '140, 550'

#menu
# Main Form .Net Objects
#$mainForm         = New-Object System.Windows.Forms.Form
#$menuView         = New-Object System.Windows.Forms.ToolStripMenuItem
#$menuTools        = New-Object System.Windows.Forms.ToolStripMenuItem
#$menuOpen         = New-Object System.Windows.Forms.ToolStripMenuItem
#$menuSave         = New-Object System.Windows.Forms.ToolStripMenuItem
#$menuSaveAs       = New-Object System.Windows.Forms.ToolStripMenuItem
<#$menuFullScr      = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOptions      = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOptions1     = New-Object System.Windows.Forms.ToolStripMenuItem
$menuOptions2     = New-Object System.Windows.Forms.ToolStripMenuItem
$menuExit         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelp         = New-Object System.Windows.Forms.ToolStripMenuItem
$menuAbout        = New-Object System.Windows.Forms.ToolStripMenuItem
$mainToolStrip    = New-Object System.Windows.Forms.ToolStrip
$toolStripOpen    = New-Object System.Windows.Forms.ToolStripButton
$toolStripSave    = New-Object System.Windows.Forms.ToolStripButton
$toolStripSaveAs  = New-Object System.Windows.Forms.ToolStripButton
$toolStripFullScr = New-Object System.Windows.Forms.ToolStripButton
$toolStripAbout   = New-Object System.Windows.Forms.ToolStripButton
$toolStripExit    = New-Object System.Windows.Forms.ToolStripButton
$statusStrip      = New-Object System.Windows.Forms.StatusStrip
$statusLabel      = New-Object System.Windows.Forms.ToolStripStatusLabel#>

#File Menu
$menuClose = New-Object System.Windows.Forms.ToolStripMenuItem
$menuClose.Name = "Close"
$menuClose.Text = "Close"
$menuClose.Add_Click({$hash.Form.Close()
    $script:powershell.EndInvoke($script:handle)
    #$script:powershell.Close()
    $script:powershell.Dispose()
    $script:runspace.Dispose()
    $script:runspace.Close()
    [System.GC]::Collect()})

#File Menu continued
$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem
$menuFile.Name = "File"
$menuFile.Text = "File"
$menuFile.DropDownItems.AddRange(@($menuClose))


#ABOUT FORM
 $FormAbout = New-Object system.Windows.Forms.Form
 $FormAboutWidth = '800'
 $FormAboutHeight = '500'
 $FormAbout.MinimumSize = "$FormAboutWidth,$FormAboutHeight"
 $FormAbout.StartPosition = 'CenterScreen'
 $FormAbout.text = "About"
 $FormAbout.Icon = [System.Drawing.SystemIcons]::Shield
 #Autoscaling settings
 $FormAbout.AutoScale = $true
 $FormAbout.AutoScaleMode = "Font"
 $ASsize = New-Object System.Drawing.SizeF(7,15)
 $FormAbout.AutoScaleDimensions = $ASsize
 $FormAbout.BackColor = 'SteelBlue'
 $FormAbout.Refresh()
 #Disable windows maximize feature.
 $FormAbout.MaximizeBox = $False
 #$FormAbout.FormBorderStyle ='FixedDialog'

 $AboutHeading = New-Object system.Windows.Forms.Label
 $AboutHeading.text = "PVT Tool"
 $AboutHeading.AutoSize = $false
 $AboutHeading.width = 150
 $AboutHeading.height = 30
 $AboutHeading.location = New-Object System.Drawing.Point(20,20)
 $AboutHeading.Font = 'Verdana,14'
 $AboutHeading.Anchor = 'top, left'

 $AboutDescription = New-Object system.Windows.Forms.Label
 $AboutDescription.text = "Credits.`r`nAuthor: $Author Build date: $AuthorDate Version: $Version"
 $AboutDescription.AutoSize = $false
 $AboutDescription.width = 500
 $AboutDescription.height = 70
 $AboutDescription.location = New-Object System.Drawing.Point(20,410)
 $AboutDescription.Font = 'Verdana,10'
 $AboutDescription.Anchor = 'bottom, left'

 $Description2Heading = New-Object system.Windows.Forms.Label
 $Description2Heading.text = "Instructions"
 $Description2Heading.AutoSize = $false
 $Description2Heading.width = 150
 $Description2Heading.height = 30
 $Description2Heading.location = New-Object System.Drawing.Point(20,80)
 $Description2Heading.Font = 'Verdana,12'

$AboutDescription2 = New-Object system.Windows.Forms.Label
$AboutDescription2.text = "1. Enter a single target computer in the first field and click Run.`r`n2. For bulk scanning; click browse to select a text file contaning multiple computers and click Run.`r`n3. When scan is complete, the PVT report will automatically pop-up.`r`n4. You can then click File -> Save As to save the report to an Excel CSV.`r`n5. By selecting desired rows and performing CTRL + C, you can copy/paste data to Excel (optional).`r`n`r`nDesigned for Microsoft Windows domain systems."
$AboutDescription2.AutoSize = $false
$AboutDescription2.width = 620
$AboutDescription2.height = 200
$AboutDescription2.location = New-Object System.Drawing.Point(20,120)
$AboutDescription2.Font = 'Verdana,10'

#About Close Button
$aboutClose = New-Object System.Windows.Forms.Button
$aboutClose.text = "Close"
#$buttonClose.Size = '80,30'
$aboutClose.Size = '80,30'
$aboutClose.location = '350, 360'
$aboutClose.Font = 'Verdana,9'
$aboutClose.Anchor = 'Bottom,Left'
$aboutClose.Add_Click({$FormAbout.Close()})

$FormAbout.Controls.AddRange(@($AboutHeading, $AboutDescription, $Description2Heading, $AboutDescription2, $aboutClose))

#Help Menu
$helpAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$helpAbout.Name = "About"
$helpAbout.Text = "About"
$helpAbout.Add_Click({$FormAbout.ShowDialog()})

$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelp.Name = "Help"
$menuHelp.Text = "Help"
$menuHelp.DropDownItems.AddRange(@($helpAbout))

$menuMain = New-Object System.Windows.Forms.MenuStrip
$menuMain.Items.AddRange(@($menuFile, $menuHelp))

#Display form
$hash.Form.Controls.AddRange(@($menuMain, $hash.serverBox, $hash.buttonBrowse, $hash.buttonRun, $Description, $moto, $hash.outputBox, $browseBox, $StatusHeading, $hash.progressBar1))
$result = $hash.Form.ShowDialog()

if($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        $hash.Form.Close()
        $script:powershell.EndInvoke($script:handle)
        #$script:powershell.Close()
        $script:powershell.Dispose()
        $script:runspace.Dispose()
        $script:runspace.Close()
        [System.GC]::Collect()
        Exit
    }

#Ensure a text file containing server list is entered.
while(!$serverList)
{
    $hash.outputBox.Selectioncolor = "Red"
    Add-OutputBoxLine -Message "`r`nMust enter a compuer name OR select a valid input file. Click Browse and select a text/csv file containing your computer list."
    $hash.buttonBrowse.enabled = $true
    $hash.buttonRun.enabled = $true
}
# SIG # Begin signature block
# MIIm1wYJKoZIhvcNAQcCoIImyDCCJsQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU2da4gof6M5k+6Cmb0WJobMBG
# 1yKggh/nMIIFbzCCBFegAwIBAgIQSPyTtGBVlI02p8mKidaUFjANBgkqhkiG9w0B
# AQwFADB7MQswCQYDVQQGEwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVzdGVy
# MRAwDgYDVQQHDAdTYWxmb3JkMRowGAYDVQQKDBFDb21vZG8gQ0EgTGltaXRlZDEh
# MB8GA1UEAwwYQUFBIENlcnRpZmljYXRlIFNlcnZpY2VzMB4XDTIxMDUyNTAwMDAw
# MFoXDTI4MTIzMTIzNTk1OVowVjELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3Rp
# Z28gTGltaXRlZDEtMCsGA1UEAxMkU2VjdGlnbyBQdWJsaWMgQ29kZSBTaWduaW5n
# IFJvb3QgUjQ2MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAjeeUEiIE
# JHQu/xYjApKKtq42haxH1CORKz7cfeIxoFFvrISR41KKteKW3tCHYySJiv/vEpM7
# fbu2ir29BX8nm2tl06UMabG8STma8W1uquSggyfamg0rUOlLW7O4ZDakfko9qXGr
# YbNzszwLDO/bM1flvjQ345cbXf0fEj2CA3bm+z9m0pQxafptszSswXp43JJQ8mTH
# qi0Eq8Nq6uAvp6fcbtfo/9ohq0C/ue4NnsbZnpnvxt4fqQx2sycgoda6/YDnAdLv
# 64IplXCN/7sVz/7RDzaiLk8ykHRGa0c1E3cFM09jLrgt4b9lpwRrGNhx+swI8m2J
# mRCxrds+LOSqGLDGBwF1Z95t6WNjHjZ/aYm+qkU+blpfj6Fby50whjDoA7NAxg0P
# OM1nqFOI+rgwZfpvx+cdsYN0aT6sxGg7seZnM5q2COCABUhA7vaCZEao9XOwBpXy
# bGWfv1VbHJxXGsd4RnxwqpQbghesh+m2yQ6BHEDWFhcp/FycGCvqRfXvvdVnTyhe
# Be6QTHrnxvTQ/PrNPjJGEyA2igTqt6oHRpwNkzoJZplYXCmjuQymMDg80EY2NXyc
# uu7D1fkKdvp+BRtAypI16dV60bV/AK6pkKrFfwGcELEW/MxuGNxvYv6mUKe4e7id
# FT/+IAx1yCJaE5UZkADpGtXChvHjjuxf9OUCAwEAAaOCARIwggEOMB8GA1UdIwQY
# MBaAFKARCiM+lvEH7OKvKe+CpX/QMKS0MB0GA1UdDgQWBBQy65Ka/zWWSC8oQEJw
# IDaRXBeF5jAOBgNVHQ8BAf8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zATBgNVHSUE
# DDAKBggrBgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEMGA1Ud
# HwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwuY29tb2RvY2EuY29tL0FBQUNlcnRpZmlj
# YXRlU2VydmljZXMuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuY29tb2RvY2EuY29tMA0GCSqGSIb3DQEBDAUAA4IBAQASv6Hvi3Sa
# mES4aUa1qyQKDKSKZ7g6gb9Fin1SB6iNH04hhTmja14tIIa/ELiueTtTzbT72ES+
# BtlcY2fUQBaHRIZyKtYyFfUSg8L54V0RQGf2QidyxSPiAjgaTCDi2wH3zUZPJqJ8
# ZsBRNraJAlTH/Fj7bADu/pimLpWhDFMpH2/YGaZPnvesCepdgsaLr4CnvYFIUoQx
# 2jLsFeSmTD1sOXPUC4U5IOCFGmjhp0g4qdE2JXfBjRkWxYhMZn0vY86Y6GnfrDyo
# XZ3JHFuu2PMvdM+4fvbXg50RlmKarkUT2n/cR/vfw1Kf5gZV6Z2M8jpiUbzsJA8p
# 1FiAhORFe1rYMIIGGjCCBAKgAwIBAgIQYh1tDFIBnjuQeRUgiSEcCjANBgkqhkiG
# 9w0BAQwFADBWMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVk
# MS0wKwYDVQQDEyRTZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYw
# HhcNMjEwMzIyMDAwMDAwWhcNMzYwMzIxMjM1OTU5WjBUMQswCQYDVQQGEwJHQjEY
# MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1Ymxp
# YyBDb2RlIFNpZ25pbmcgQ0EgUjM2MIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIB
# igKCAYEAmyudU/o1P45gBkNqwM/1f/bIU1MYyM7TbH78WAeVF3llMwsRHgBGRmxD
# eEDIArCS2VCoVk4Y/8j6stIkmYV5Gej4NgNjVQ4BYoDjGMwdjioXan1hlaGFt4Wk
# 9vT0k2oWJMJjL9G//N523hAm4jF4UjrW2pvv9+hdPX8tbbAfI3v0VdJiJPFy/7Xw
# iunD7mBxNtecM6ytIdUlh08T2z7mJEXZD9OWcJkZk5wDuf2q52PN43jc4T9OkoXZ
# 0arWZVeffvMr/iiIROSCzKoDmWABDRzV/UiQ5vqsaeFaqQdzFf4ed8peNWh1OaZX
# nYvZQgWx/SXiJDRSAolRzZEZquE6cbcH747FHncs/Kzcn0Ccv2jrOW+LPmnOyB+t
# AfiWu01TPhCr9VrkxsHC5qFNxaThTG5j4/Kc+ODD2dX/fmBECELcvzUHf9shoFvr
# n35XGf2RPaNTO2uSZ6n9otv7jElspkfK9qEATHZcodp+R4q2OIypxR//YEb3fkDn
# 3UayWW9bAgMBAAGjggFkMIIBYDAfBgNVHSMEGDAWgBQy65Ka/zWWSC8oQEJwIDaR
# XBeF5jAdBgNVHQ4EFgQUDyrLIIcouOxvSK4rVKYpqhekzQwwDgYDVR0PAQH/BAQD
# AgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwGwYD
# VR0gBBQwEjAGBgRVHSAAMAgGBmeBDAEEATBLBgNVHR8ERDBCMECgPqA8hjpodHRw
# Oi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RS
# NDYuY3JsMHsGCCsGAQUFBwEBBG8wbTBGBggrBgEFBQcwAoY6aHR0cDovL2NydC5z
# ZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdSb290UjQ2LnA3YzAj
# BggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEM
# BQADggIBAAb/guF3YzZue6EVIJsT/wT+mHVEYcNWlXHRkT+FoetAQLHI1uBy/YXK
# ZDk8+Y1LoNqHrp22AKMGxQtgCivnDHFyAQ9GXTmlk7MjcgQbDCx6mn7yIawsppWk
# vfPkKaAQsiqaT9DnMWBHVNIabGqgQSGTrQWo43MOfsPynhbz2Hyxf5XWKZpRvr3d
# MapandPfYgoZ8iDL2OR3sYztgJrbG6VZ9DoTXFm1g0Rf97Aaen1l4c+w3DC+IkwF
# kvjFV3jS49ZSc4lShKK6BrPTJYs4NG1DGzmpToTnwoqZ8fAmi2XlZnuchC4NPSZa
# PATHvNIzt+z1PHo35D/f7j2pO1S8BCysQDHCbM5Mnomnq5aYcKCsdbh0czchOm8b
# kinLrYrKpii+Tk7pwL7TjRKLXkomm5D1Umds++pip8wH2cQpf93at3VDcOK4N7Ew
# oIJB0kak6pSzEu4I64U6gZs7tS/dGNSljf2OSSnRr7KWzq03zl8l75jy+hOds9TW
# SenLbjBQUGR96cFr6lEUfAIEHVC1L68Y1GGxx4/eRI82ut83axHMViw1+sVpbPxg
# 51Tbnio1lB93079WPFnYaOvfGAA0e0zcfF/M9gXr+korwQTh2Prqooq2bYNMvUoU
# KD85gnJ+t0smrWrb8dee2CvYZXD5laGtaAxOfy/VKNmwuWuAh9kcMIIGaDCCBNCg
# AwIBAgIRAL/9KI6HeeUhYmELSlorbnEwDQYJKoZIhvcNAQEMBQAwVDELMAkGA1UE
# BhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDErMCkGA1UEAxMiU2VjdGln
# byBQdWJsaWMgQ29kZSBTaWduaW5nIENBIFIzNjAeFw0yMTA5MTAwMDAwMDBaFw0y
# MjA5MTAyMzU5NTlaMEYxCzAJBgNVBAYTAkFVMREwDwYDVQQIDAhWaWN0b3JpYTER
# MA8GA1UECgwIQW5nbGUgSVQxETAPBgNVBAMMCEFuZ2xlIElUMIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAxR2UqErPaN9XXcwEVPrABibDn9uIA+/GbnZU
# boNwvzTyFoIVaUw5/8QO/6s8j6i8TVC9szyXbxBVGpxMgKg++RZ5t4pkZJWtDU9l
# XrHErNitHDT4WMXJj5ihPM2AgkBSJUsYiSbVDuNmwV9x2nizLY3212rJeMYsVYw5
# qQ/UUNIlDx+CYogVc6esFf6gnhnf7UMlJZqDdxV/AidKtabQLoRrQK6UqGbA9CUQ
# YoJECrDQ7bRsGlByWgdOOQHXDtzDgfA3NYculKbIrm63OdIxdhE5lBSuBfc7XQqB
# l+1rPp4loYipXKFyXVolnFlSovL1LggkevQi4g1yr0aVzqAxC/NNtG9JMw/U9e4F
# rKIYTtvAbVYdG5KIx3sqteHJ56N6F1b4SpLHyB+p5zMmuIwi9gPzhCXrzlKO4GmE
# zPoVKmZhYezd1gkhlMjhlT3DkalZQIVYTYTTNgbsbuCuu7Rv6P63ioRaTfku/fyn
# QDxysZEnN8ZVzjsI14cNwGE5ILESYl831o/q5m/diTOWn3ZrC5GV4hGcVQgnkY6r
# jEiY7uo8J/ybPENDuWpyPUrd4APKHTW3w9jtkaI1HapNBQeJKyEYErTXyoVYiRjx
# NHJmbFjuISPgx9FVrYeMc5KAPDjAl8XL+RfOcjv1COlVupCeuoIhOkjt7mvo8nsk
# THuPvdECAwEAAaOCAcEwggG9MB8GA1UdIwQYMBaAFA8qyyCHKLjsb0iuK1SmKaoX
# pM0MMB0GA1UdDgQWBBS7Ym2QqvITo7sv+Bg8ny5yRnwZLTAOBgNVHQ8BAf8EBAMC
# B4AwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDAzARBglghkgBhvhC
# AQEEBAMCBBAwSgYDVR0gBEMwQTA1BgwrBgEEAbIxAQIBAwIwJTAjBggrBgEFBQcC
# ARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwCAYGZ4EMAQQBMEkGA1UdHwRCMEAw
# PqA8oDqGOGh0dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY0NvZGVT
# aWduaW5nQ0FSMzYuY3JsMHkGCCsGAQUFBwEBBG0wazBEBggrBgEFBQcwAoY4aHR0
# cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdDQVIz
# Ni5jcnQwIwYIKwYBBQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28uY29tMCMGA1Ud
# EQQcMBqBGGRldmVsb3BlckBhbmdsZWl0LmNvbS5hdTANBgkqhkiG9w0BAQwFAAOC
# AYEAUBEh0nzKQFwac0oHXEFXPdsnuygRXw7R8Q42vmN2Gney52aRRVyrQgIR7SFt
# YH2vLLDvL6U6YxAQ2PEx6unk4ng/uoS8wlZ41Dv1uHRIMEbfQ3BBPJ97aaot63LF
# +8J3dxD6lZgXsanrDQBX3fkRXo/q8E3RonsHwkMzcGskE6wIgfZj+7Qe9l2cyWBj
# Vt6TAbi31/XUf9R3Xj4CtaTwklXs9XBVkknXKVhV/3aowyVpGILQS/4Ifvu/B0v+
# KmjFP5fvBVQ4CEIFvxWnWDahwpplxyk8ILt43MMomiw306TCCfo3hO6PUDiar7Ew
# JjWP2CFfuc3eYs1bevzxHgG58MCcl5AvP6WFuD6LvMA8D0Uz0HYlXbwXbxMswGmu
# jcUNH6QY8NHws/y4KyyWYVEhme/AXLLyR94sS16EzDeez/BfHuj6LiG46YvAqiAB
# GQXn9P2jZzJ8uDioQA4+BnUTZML7p104jAoeZktK1LYoxx1sJQvxr9vemd79Zai5
# edhmMIIG7DCCBNSgAwIBAgIQMA9vrN1mmHR8qUY2p3gtuTANBgkqhkiG9w0BAQwF
# ADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCk5ldyBKZXJzZXkxFDASBgNVBAcT
# C0plcnNleSBDaXR5MR4wHAYDVQQKExVUaGUgVVNFUlRSVVNUIE5ldHdvcmsxLjAs
# BgNVBAMTJVVTRVJUcnVzdCBSU0EgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkwHhcN
# MTkwNTAyMDAwMDAwWhcNMzgwMTE4MjM1OTU5WjB9MQswCQYDVQQGEwJHQjEbMBkG
# A1UECBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdTYWxmb3JkMRgwFgYD
# VQQKEw9TZWN0aWdvIExpbWl0ZWQxJTAjBgNVBAMTHFNlY3RpZ28gUlNBIFRpbWUg
# U3RhbXBpbmcgQ0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDIGwGv
# 2Sx+iJl9AZg/IJC9nIAhVJO5z6A+U++zWsB21hoEpc5Hg7XrxMxJNMvzRWW5+adk
# FiYJ+9UyUnkuyWPCE5u2hj8BBZJmbyGr1XEQeYf0RirNxFrJ29ddSU1yVg/cyeNT
# mDoqHvzOWEnTv/M5u7mkI0Ks0BXDf56iXNc48RaycNOjxN+zxXKsLgp3/A2UUrf8
# H5VzJD0BKLwPDU+zkQGObp0ndVXRFzs0IXuXAZSvf4DP0REKV4TJf1bgvUacgr6U
# nb+0ILBgfrhN9Q0/29DqhYyKVnHRLZRMyIw80xSinL0m/9NTIMdgaZtYClT0Bef9
# Maz5yIUXx7gpGaQpL0bj3duRX58/Nj4OMGcrRrc1r5a+2kxgzKi7nw0U1BjEMJh0
# giHPYla1IXMSHv2qyghYh3ekFesZVf/QOVQtJu5FGjpvzdeE8NfwKMVPZIMC1Pvi
# 3vG8Aij0bdonigbSlofe6GsO8Ft96XZpkyAcSpcsdxkrk5WYnJee647BeFbGRCXf
# BhKaBi2fA179g6JTZ8qx+o2hZMmIklnLqEbAyfKm/31X2xJ2+opBJNQb/HKlFKLU
# rUMcpEmLQTkUAx4p+hulIq6lw02C0I3aa7fb9xhAV3PwcaP7Sn1FNsH3jYL6uckN
# U4B9+rY5WDLvbxhQiddPnTO9GrWdod6VQXqngwIDAQABo4IBWjCCAVYwHwYDVR0j
# BBgwFoAUU3m/WqorSs9UgOHYm8Cd8rIDZsswHQYDVR0OBBYEFBqh+GEZIA/DQXdF
# KI7RNV8GEgRVMA4GA1UdDwEB/wQEAwIBhjASBgNVHRMBAf8ECDAGAQH/AgEAMBMG
# A1UdJQQMMAoGCCsGAQUFBwMIMBEGA1UdIAQKMAgwBgYEVR0gADBQBgNVHR8ESTBH
# MEWgQ6BBhj9odHRwOi8vY3JsLnVzZXJ0cnVzdC5jb20vVVNFUlRydXN0UlNBQ2Vy
# dGlmaWNhdGlvbkF1dGhvcml0eS5jcmwwdgYIKwYBBQUHAQEEajBoMD8GCCsGAQUF
# BzAChjNodHRwOi8vY3J0LnVzZXJ0cnVzdC5jb20vVVNFUlRydXN0UlNBQWRkVHJ1
# c3RDQS5jcnQwJQYIKwYBBQUHMAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20w
# DQYJKoZIhvcNAQEMBQADggIBAG1UgaUzXRbhtVOBkXXfA3oyCy0lhBGysNsqfSoF
# 9bw7J/RaoLlJWZApbGHLtVDb4n35nwDvQMOt0+LkVvlYQc/xQuUQff+wdB+PxlwJ
# +TNe6qAcJlhc87QRD9XVw+K81Vh4v0h24URnbY+wQxAPjeT5OGK/EwHFhaNMxcyy
# UzCVpNb0llYIuM1cfwGWvnJSajtCN3wWeDmTk5SbsdyybUFtZ83Jb5A9f0VywRsj
# 1sJVhGbks8VmBvbz1kteraMrQoohkv6ob1olcGKBc2NeoLvY3NdK0z2vgwY4Eh0k
# hy3k/ALWPncEvAQ2ted3y5wujSMYuaPCRx3wXdahc1cFaJqnyTdlHb7qvNhCg0MF
# pYumCf/RoZSmTqo9CfUFbLfSZFrYKiLCS53xOV5M3kg9mzSWmglfjv33sVKRzj+J
# 9hyhtal1H3G/W0NdZT1QgW6r8NDT/LKzH7aZlib0PHmLXGTMze4nmuWgwAxyh8Fu
# TVrTHurwROYybxzrF06Uw3hlIDsPQaof6aFBnf6xuKBlKjTg3qj5PObBMLvAoGMs
# /FwWAKjQxH/qEZ0eBsambTJdtDgJK0kHqv3sMNrxpy/Pt/360KOE2See+wFmd7lW
# EOEgbsausfm2usg1XTN2jvF8IAwqd661ogKGuinutFoAsYyr4/kKyVRd1LlqdJ69
# SK6YMIIG9jCCBN6gAwIBAgIRAJA5f5rSSjoT8r2RXwg4qUMwDQYJKoZIhvcNAQEM
# BQAwfTELMAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQ
# MA4GA1UEBxMHU2FsZm9yZDEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSUwIwYD
# VQQDExxTZWN0aWdvIFJTQSBUaW1lIFN0YW1waW5nIENBMB4XDTIyMDUxMTAwMDAw
# MFoXDTMzMDgxMDIzNTk1OVowajELMAkGA1UEBhMCR0IxEzARBgNVBAgTCk1hbmNo
# ZXN0ZXIxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAwwjU2VjdGln
# byBSU0EgVGltZSBTdGFtcGluZyBTaWduZXIgIzMwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQCQsnE/eeHUuYoXzMOXwpCUcu1aOm8BQ39zWiifJHygNUAG
# +pSvCqGDthPkSxUGXmqKIDRxe7slrT9bCqQfL2x9LmFR0IxZNz6mXfEeXYC22B9g
# 480Saogfxv4Yy5NDVnrHzgPWAGQoViKxSxnS8JbJRB85XZywlu1aSY1+cuRDa3/J
# oD9sSq3VAE+9CriDxb2YLAd2AXBF3sPwQmnq/ybMA0QfFijhanS2nEX6tjrOlNEf
# vYxlqv38wzzoDZw4ZtX8fR6bWYyRWkJXVVAWDUt0cu6gKjH8JgI0+WQbWf3jOtTo
# uEEpdAE/DeATdysRPPs9zdDn4ZdbVfcqA23VzWLazpwe/OpwfeZ9S2jOWilh06Bc
# JbOlJ2ijWP31LWvKX2THaygM2qx4Qd6S7w/F7KvfLW8aVFFsM7ONWWDn3+gXIqN5
# QWLP/Hvzktqu4DxPD1rMbt8fvCKvtzgQmjSnC//+HV6k8+4WOCs/rHaUQZ1kHfqA
# /QDh/vg61MNeu2lNcpnl8TItUfphrU3qJo5t/KlImD7yRg1psbdu9AXbQQXGGMBQ
# 5Pit/qxjYUeRvEa1RlNsxfThhieThDlsdeAdDHpZiy7L9GQsQkf0VFiFN+XHaafS
# JYuWv8at4L2xN/cf30J7qusc6es9Wt340pDVSZo6HYMaV38cAcLOHH3M+5YVxQID
# AQABo4IBgjCCAX4wHwYDVR0jBBgwFoAUGqH4YRkgD8NBd0UojtE1XwYSBFUwHQYD
# VR0OBBYEFCUuaDxrmiskFKkfot8mOs8UpvHgMA4GA1UdDwEB/wQEAwIGwDAMBgNV
# HRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEoGA1UdIARDMEEwNQYM
# KwYBBAGyMQECAQMIMCUwIwYIKwYBBQUHAgEWF2h0dHBzOi8vc2VjdGlnby5jb20v
# Q1BTMAgGBmeBDAEEAjBEBgNVHR8EPTA7MDmgN6A1hjNodHRwOi8vY3JsLnNlY3Rp
# Z28uY29tL1NlY3RpZ29SU0FUaW1lU3RhbXBpbmdDQS5jcmwwdAYIKwYBBQUHAQEE
# aDBmMD8GCCsGAQUFBzAChjNodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29S
# U0FUaW1lU3RhbXBpbmdDQS5jcnQwIwYIKwYBBQUHMAGGF2h0dHA6Ly9vY3NwLnNl
# Y3RpZ28uY29tMA0GCSqGSIb3DQEBDAUAA4ICAQBz2u1ocsvCuUChMbu0A6MtFHsk
# 57RbFX2o6f2t0ZINfD02oGnZ85ow2qxp1nRXJD9+DzzZ9cN5JWwm6I1ok87xd4k5
# f6gEBdo0wxTqnwhUq//EfpZsK9OU67Rs4EVNLLL3OztatcH714l1bZhycvb3Byjz
# 07LQ6xm+FSx4781FoADk+AR2u1fFkL53VJB0ngtPTcSqE4+XrwE1K8ubEXjp8vmJ
# BDxO44ISYuu0RAx1QcIPNLiIncgi8RNq2xgvbnitxAW06IQIkwf5fYP+aJg05Hfl
# sc6MlGzbA20oBUd+my7wZPvbpAMxEHwa+zwZgNELcLlVX0e+OWTOt9ojVDLjRrIy
# 2NIphskVXYCVrwL7tNEunTh8NeAPHO0bR0icImpVgtnyughlA+XxKfNIigkBTKZ5
# 8qK2GpmU65co4b59G6F87VaApvQiM5DkhFP8KvrAp5eo6rWNes7k4EuhM6sLdqDV
# aRa3jma/X/ofxKh/p6FIFJENgvy9TZntyeZsNv53Q5m4aS18YS/to7BJ/lu+aSSR
# /5P8V2mSS9kFP22GctOi0MBk0jpCwRoD+9DtmiG4P6+mslFU1UzFyh8SjVfGOe1c
# /+yfJnatZGZn6Kow4NKtt32xakEnbgOKo3TgigmCbr/j9re8ngspGGiBoZw/bhZZ
# SxQJCZrmrr9gFd2G9TGCBlowggZWAgEBMGkwVDELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDErMCkGA1UEAxMiU2VjdGlnbyBQdWJsaWMgQ29k
# ZSBTaWduaW5nIENBIFIzNgIRAL/9KI6HeeUhYmELSlorbnEwCQYFKw4DAhoFAKB4
# MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkE
# MRYEFMK822cN5766Kr8ZDcaPUbr9gp/XMA0GCSqGSIb3DQEBAQUABIICALVU/AUu
# 4AZYWMnME05gec49hnT+nnawJoAyZEELzsvbk34cOsfodua5SnL/vd7CNOz7M0Gl
# 2OTyw6jsUs98ThHgpYmCToJz3LnaVBsABRpureY7kOrhpdEp4EoZo58IR1NTk46z
# k63ficdDuAzZAXIP/lE28nR94D3VkKYXC75xJ/w1AOjG1CcgHcDmVXh87fK9KO24
# ANWCCSrUhCXIFRAN8qVF5JrQ2rMsEFBh8T1YZice6HoRHQqTseOZ1FrYT8QcGGSD
# E8a6j9XK5wvRw07+wyte354+FGDvv9LGCAYurq6l90YyVGygnVSokxoISfl28zj9
# GQohIhjjkIX+/VP0vhyuEMzJhKI4kpUZgTMoVpsl1hqmghlUwMKI8Gj3Xcu3Ar4X
# yOJcqmiB+0bq1Cpar3Khvblq2rQYSKhb6ebNG/4AmuOGqN3bFxmet5VDycAcfnea
# 11EVmWYPvhM4FUEHzXiCGiAnHU5NqQE2XZKtR+bquNGGmutPGNuVpj1BK+Lstaw3
# Nnq0KfpIgs636YXrl6XmOZYAs3wlW0EmAOdP4x2Z1kQiISA4H3CXuiFdkTZ8+h54
# 0ub2JFacv6jZuq8Opa+EZltmHLBqxNf9Sf6r8ZvbjNlzZIZUMuGNtYirdpZYymjK
# K1VgqQecw35dOwKmzb6jll3KZrSvlS9EIxKWoYIDTDCCA0gGCSqGSIb3DQEJBjGC
# AzkwggM1AgEBMIGSMH0xCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1h
# bmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDElMCMGA1UEAxMcU2VjdGlnbyBSU0EgVGltZSBTdGFtcGluZyBDQQIRAJA5
# f5rSSjoT8r2RXwg4qUMwDQYJYIZIAWUDBAICBQCgeTAYBgkqhkiG9w0BCQMxCwYJ
# KoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yMjA4MTUxMDUzMjhaMD8GCSqGSIb3
# DQEJBDEyBDBSmUC1WZ8BG9CEKt6CSD1HEtLgjy+AGNnxGcLgclerkYI6KhzbZe5J
# 2fMjXxAeGEkwDQYJKoZIhvcNAQEBBQAEggIAHk8MUWTSrIsnWhZbbwTppLHZv0Bn
# 42ExzwVea8IGqZz1q+JOOkpX1hRs2LuVEJxWDmgZcK3lL39zYV7kg8Z3bNj5lTBu
# lZ65lTCpyOgSdtggClnk97OGmZJ6R2u5A9h9LyO1nJLGHZ6SNnlTB3pkdGuTCtsW
# YIzMOMr4r7/v4hS7xpJJhZO2UOCFVJJ7kJ0yK2nLz8c5xSBXYYLvYGDhPDn8U8bg
# hdBYBBPkdQIILwiJtZNl7VSyKLAkE5niEw0QnuhCZkPq+JA6suxqlxVbUXz341O3
# CV5B2xFe0l4enhDgcM/Ii3MMYavZHvkv0SnCjbCPv4EBCcggczBo9Cf2g8pzhDpG
# dTC4YnEPCDMtzwWhg4a6YipVwBimr9wigG/VwT0D0czzvVLI8OSrbESqL+lx+82S
# 66s1rsemhNZ94Hqvzd6q55pzXQ6Vtkq6gKcnC2CpoUsvw+sSvDZS1EYLw4KkKDv7
# b1rtZ+fhvqE7Qk4vRU3SS2flozZtgdT1DKIvzmlqS0n5nzs9O8sKklF7ckoogQ7q
# tOPPx4HNjIt0EsMYBj4/KezyrbIczVGHhZB15duu29c51ZUPrEitjQdJzu7jOlIU
# /m7jhVe/S5ya6p2RAo30udY+9UZ4nmFOke2Eg4wZ3Bd0jUUca2AMJFboQ3ez4/eO
# c2lPBgm66QV8qAo=
# SIG # End signature block
