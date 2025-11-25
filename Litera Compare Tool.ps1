param(
    [switch]$StartInProMode
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

[System.Windows.Forms.Application]::EnableVisualStyles()

#==============================================================================
# --- ASSEMBLY AND FORM SETUP ---
#==============================================================================

$script:Form = New-Object System.Windows.Forms.Form
$script:Form.Text = "Litera Compare Management Tool"
$script:Form.ShowInTaskbar = $true
$script:Form.Size = New-Object System.Drawing.Size(750, 650)
$script:Form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
$script:Form.StartPosition = "CenterScreen"
$script:Form.FormBorderStyle = 'FixedSingle'
$script:Form.MaximizeBox = $true
$script:Form.MinimizeBox = $true

$script:ToolTip = New-Object System.Windows.Forms.ToolTip

# --- Timer for polling the log file ---
$script:logPollTimer = New-Object System.Windows.Forms.Timer
$script:logPollTimer.Interval = 1000 # Check every 1 second

# --- Menu Strip ---
$script:MenuStrip = New-Object System.Windows.Forms.MenuStrip
$script:MenuStrip.Dock = 'Top'
$script:HelpMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Help")
$script:HelpMenuItem.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
$script:AboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("About")
$script:CheckUpdatesMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Check for Updates")
$script:SettingsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Settings")
$script:ToggleProModeItem = New-Object System.Windows.Forms.ToolStripMenuItem("Pro Mode")
$script:ToggleProModeItem.CheckOnClick = $true

$script:HelpMenuItem.DropDownItems.AddRange(@(
    $script:AboutMenuItem,
    $script:CheckUpdatesMenuItem,
    $script:SettingsMenuItem,
    $script:ToggleProModeItem
))
$script:MenuStrip.Items.Add($script:HelpMenuItem)

# --- Tab Control ---
$script:TabControl = New-Object System.Windows.Forms.TabControl
$script:TabControl.Dock = 'Fill'

# --- Create Tabs ---
$script:TabCleanup = New-Object System.Windows.Forms.TabPage
$script:TabCleanup.Text = "User File Cleanup"
$script:TabBackupRestore = New-Object System.Windows.Forms.TabPage
$script:TabBackupRestore.Text = "Backup & Restore"
$script:TabSysReq = New-Object System.Windows.Forms.TabPage
$script:TabSysReq.Text = "System Requirements/Pre-requisites"
$script:TabInstallUninstall = New-Object System.Windows.Forms.TabPage
$script:TabInstallUninstall.Text = "Install/Uninstall"
$script:TabLicense = New-Object System.Windows.Forms.TabPage
$script:TabLicense.Text = "License Info"
$script:TabCompareCustom = New-Object System.Windows.Forms.TabPage
$script:TabCompareCustom.Text = "Compare Customizations"
$script:TabAddinMgmt = New-Object System.Windows.Forms.TabPage
$script:TabAddinMgmt.Text = "Office Add-in Management"
$script:TabLogViewer = New-Object System.Windows.Forms.TabPage
$script:TabLogViewer.Text = "Log Viewer"

$script:TabControl.TabPages.AddRange(@($script:TabCleanup, $script:TabBackupRestore, $script:TabSysReq, $script:TabInstallUninstall, $script:TabLicense, $script:TabCompareCustom, $script:TabAddinMgmt, $script:TabLogViewer))

$script:Form.Controls.Add($script:TabControl)
$script:Form.Controls.Add($script:MenuStrip)
$script:Form.MainMenuStrip = $script:MenuStrip

#==============================================================================
# --- Global Variables & Paths ---
#==============================================================================
$script:selfPath = $null
$currentProcess = Get-Process -Id $PID
if ($currentProcess.ProcessName -in 'powershell', 'pwsh', 'powershell_ise') {
    $script:selfPath = $MyInvocation.MyCommand.Path
    if ($PSScriptRoot) {
        $basePath = $PSScriptRoot
    } else {
        $basePath = Split-Path -Path $script:selfPath -Parent
    }
} else {
    $script:selfPath = $currentProcess.Path
    $basePath = Split-Path -Path $script:selfPath -Parent
}
$localAppData = $env:LOCALAPPDATA
$appData = $env:APPDATA
$programData = $env:ProgramData
$literaLocalAppData = Join-Path $localAppData "Litera"
$literaAppData = Join-Path $appData "Litera"
$literaProgramData = Join-Path $programData "Litera"
$userCustomizationsPath = Join-Path $appData "Litera\UserCustomizations"
$defaultBackupRoot = Join-Path $basePath "Litera_Backup"
$defaultLogRoot = Join-Path $basePath "Logs"
$configPath = Join-Path $basePath "config"
$settingsFile = Join-Path $configPath "LC_Tool_Settings.config"
$script:backupRoot = $defaultBackupRoot
$script:logRoot = $defaultLogRoot
$script:logEnabled = $false
$script:settings = $null
$script:installedApps = @{}
$script:backupSources = @{}
$script:systemRequirementsJson = $null
$script:selectedUser = $env:USERNAME
$script:prerequisitesJson = $null
$script:lastModifiedXmlPath = $null
$sysReqUrl = "https://raw.githubusercontent.com/imumeshk/lc_tool/main/sys-req.json"
$preReqUrl = "https://raw.githubusercontent.com/imumeshk/lc_tool/main/pre-req.json"

#==============================================================================
# --- INITIALIZATION ---
#==============================================================================
if (-not (Test-Path $script:backupRoot)) {
    New-Item -ItemType Directory -Path $script:backupRoot | Out-Null
}

#==============================================================================
# --- GLOBAL FUNCTIONS ---
#==============================================================================
Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;

    public class User32 {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    }
"@ -PassThru | Out-Null

<#
.SYNOPSIS
    Sets the visual state of a ProgressBar control.
.DESCRIPTION
    Uses P/Invoke to send a message to the progress bar, changing its color to indicate normal, success (green), or error (red) states.
.PARAMETER ProgressBar
    The System.Windows.Forms.ProgressBar control to modify.
.PARAMETER State
    The desired state: 'Normal', 'Success', or 'Error'.
#>
function Set-ProgressBarState {
    param(
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [string]$State
    )
    $PBM_SETSTATE = 0x410
    $PBST_NORMAL = 1
    $PBST_ERROR = 2
    $PBST_PAUSED = 3

    switch ($State) {
        'Normal' {
            $ProgressBar.Style = 'Continuous'
            [User32]::SendMessage($ProgressBar.Handle, $PBM_SETSTATE, [IntPtr]$PBST_NORMAL, [IntPtr]::Zero) | Out-Null
            $ProgressBar.Style = 'Marquee'
        }
        'Success' {
            $ProgressBar.Style = 'Continuous'
            $ProgressBar.Value = $ProgressBar.Maximum
            [User32]::SendMessage($ProgressBar.Handle, $PBM_SETSTATE, [IntPtr]$PBST_NORMAL, [IntPtr]::Zero) | Out-Null
        }
        'Error' {
            $ProgressBar.Style = 'Continuous'
            $ProgressBar.Value = $ProgressBar.Maximum
            [User32]::SendMessage($ProgressBar.Handle, $PBM_SETSTATE, [IntPtr]$PBST_ERROR, [IntPtr]::Zero) | Out-Null
        }
    }
}

<#
.SYNOPSIS
    Gets the AppData and LocalAppData paths for the currently selected user.
.DESCRIPTION
    Returns a custom object containing the standard and Litera-specific AppData paths for the user specified in the global `$script:selectedUser` variable. It handles both the current user and other users on the system.
.RETURNS
    A PSCustomObject with properties: localAppData, appData, literaLocalAppData, literaAppData.
#>
function Get-SelectedUserPaths {
    $userName = $script:selectedUser
    $userProfilePath = "C:\Users\$userName"

    $paths = @{}
    if ($userName -eq $env:USERNAME) {
        $paths.localAppData = $env:LOCALAPPDATA
        $paths.appData = $env:APPDATA
    } else {
        if (-not (Test-Path $userProfilePath)) {
            throw "User profile path not found for '$userName'."
        }
        $paths.localAppData = Join-Path $userProfilePath "AppData\Local"
        $paths.appData = Join-Path $userProfilePath "AppData\Roaming"
    }

    $paths.literaLocalAppData = Join-Path $paths.localAppData "Litera"
    $paths.literaAppData = Join-Path $paths.appData "Litera"
    
    return [pscustomobject]$paths
}

<#
.SYNOPSIS
    Placeholder for a logging function.
.DESCRIPTION
    This function is intended for logging but is currently a no-op. It can be expanded to write messages to a file or another output stream.
.PARAMETER message
    The log message string.
#>
function Write-Log {
    param([string]$message)
}

<#
.SYNOPSIS
    Formats a PSCustomObject into a human-readable, indented string.
.DESCRIPTION
    Recursively traverses a PSCustomObject and its nested objects or arrays, formatting them into a string suitable for display in a RichTextBox or console.
.PARAMETER Item
    The PSCustomObject to format.
.PARAMETER IndentLevel
    The starting indentation level for formatting.
.RETURNS
    A formatted string representation of the object.
#>
function Format-ObjectForDisplay {
    param($Item, [int]$IndentLevel = 0)
    $output = New-Object System.Text.StringBuilder
    $indent = "  " * $IndentLevel

    foreach ($prop in $Item.PSObject.Properties | Sort-Object Name) {
        $key = $prop.Name
        $value = $prop.Value
        $formattedKey = ($key -creplace '([A-Z])', ' $1').Trim()

        if ($value -is [System.Management.Automation.PSCustomObject]) {
            $output.AppendLine("$indent${formattedKey}:") | Out-Null
            $output.Append((Format-ObjectForDisplay -Item $value -IndentLevel ($IndentLevel + 1))) | Out-Null
        } elseif ($value -is [array]) {
            $output.AppendLine("$indent${formattedKey}:") | Out-Null
            foreach ($arrayItem in $value) {
                if ($arrayItem -is [System.Management.Automation.PSCustomObject]) {
                    $output.Append((Format-ObjectForDisplay -Item $arrayItem -IndentLevel ($IndentLevel + 1))) | Out-Null
                } else {
                    $output.AppendLine("$($indent)  • $arrayItem") | Out-Null
                }
            }
        } else {
            $output.AppendLine("$indent${formattedKey}: $value") | Out-Null
        }
    }
    return $output.ToString()
}

<#
.SYNOPSIS
    Asynchronously fetches system requirement and pre-requisite data from GitHub.
.DESCRIPTION
    Starts a background job to download two JSON files from specified URLs. It updates the UI to indicate it's busy and defines a completion action to process the results (or handle errors) once the job finishes.
#>
function Fetch-RequirementsData {
    if ($script:currentJob) { [System.Windows.Forms.MessageBox]::Show("Another operation is already in progress.", "Busy", "OK", "Warning"); return }

    if (-not (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet)) {
        $script:RtbSysReq.Text = "Error: No internet connection."
        $script:RtbAppPreReq.Text = "Error: No internet connection."
        [System.Windows.Forms.MessageBox]::Show("No internet connection detected. Cannot fetch latest requirements.", "Connection Error", "OK", "Error")
        return
    }

    $script:RtbSysReq.Text = "Fetching from server..."
    $script:RtbAppPreReq.Text = "Fetching from server..."
    $script:BtnFetchRefreshData.Enabled = $false
    $script:BtnCheckMySys.Enabled = $false
    $script:BtnCheckMyPreReq.Enabled = $false

    $script:jobCompletionAction = {
        param($jobResult)
        $script:BtnFetchRefreshData.Enabled = $true
        $script:BtnCheckMySys.Enabled = $true
        $script:BtnCheckMyPreReq.Enabled = $true

        if ($jobResult -is [System.Management.Automation.ErrorRecord]) {
            $script:RtbSysReq.Text = "Failed to fetch content: $($jobResult.Exception.Message)"
            $script:RtbAppPreReq.Text = "Failed to fetch content: $($jobResult.Exception.Message)"
        } elseif ($null -eq $jobResult -or $null -eq $jobResult.SysReq -or $null -eq $jobResult.PreReq) {
            $script:RtbSysReq.Text = "Failed to process fetched content: Received empty or invalid data from server."
            $script:RtbAppPreReq.Text = "Failed to process fetched content: Received empty or invalid data from server."
        } else {
            if ($jobResult.SysReq.PSObject.Properties.Name -contains 'SystemRequirements') {
                $script:systemRequirementsJson = $jobResult.SysReq.SystemRequirements
            } else {
                $script:systemRequirementsJson = $jobResult.SysReq
            }

            if ($jobResult.PreReq.PSObject.Properties.Name -contains 'Prerequisites') {
                $script:prerequisitesJson = $jobResult.PreReq.Prerequisites
            }
            else {
                $script:prerequisitesJson = $jobResult.PreReq
            }

            $sysReqOutput = New-Object System.Text.StringBuilder
            $sysReqOutput.AppendLine("--- System Requirements ---`n") | Out-Null
            $sysReqOutput.Append((Format-ObjectForDisplay -Item $script:systemRequirementsJson)) | Out-Null
            $script:RtbSysReq.Text = $sysReqOutput.ToString()
            $preReqOutput = New-Object System.Text.StringBuilder
            $preReqOutput.AppendLine("--- Application Pre-requisites ---") | Out-Null
            foreach ($item in $script:prerequisitesJson) {
                if ($item.PSObject.Properties.Name -contains 'Category') {
                    $preReqOutput.AppendLine("`n--- $($item.Category) ---") | Out-Null
                    foreach ($req in $item.Requirements) {
                        $preReqOutput.AppendLine(" • $($req.Name)") | Out-Null
                    }
                } else {
                    $preReqOutput.AppendLine("`n • $($item.Name)") | Out-Null
                }
            }
            $script:RtbAppPreReq.Text = $preReqOutput.ToString()
        }
    }

    $jobArgs = @{ SysReqUrl = $sysReqUrl; PreReqUrl = $preReqUrl }

    try {
        $script:currentJob = Start-Job -ArgumentList $jobArgs -ScriptBlock {
            param($args)
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            try {
                $headers = @{ 'Cache-Control' = 'no-cache'; 'Pragma' = 'no-cache' }
                $sysReqJson = Invoke-RestMethod -Uri $args.SysReqUrl -UseBasicParsing -Headers $headers
                $preReqJson = Invoke-RestMethod -Uri $args.PreReqUrl -UseBasicParsing -Headers $headers
                return [PSCustomObject]@{ SysReq = $sysReqJson; PreReq = $preReqJson }
            } catch { 
                throw "Error fetching online content: $($_.Exception.Message)"
            }
        }
        $JobTimer.Start()
    } catch {
        $script:RtbSysReq.Text = "Failed to start fetch operation: $($_.Exception.Message)"
        $script:RtbAppPreReq.Text = "Failed to start fetch operation: $($_.Exception.Message)"
        $script:BtnFetchRefreshData.Enabled = $true
        $script:BtnCheckMySys.Enabled = $true
        $script:BtnCheckMyPreReq.Enabled = $true
    }
}

<#
.SYNOPSIS
    Parses a Litera customization XML file.
.DESCRIPTION
    Reads a Litera `Customize.xml`-style file and extracts all settings, returning them as a dictionary. Each setting includes its value and parent node name.
.PARAMETER FilePath
    The full path to the XML file to parse.
.RETURNS
    A dictionary of settings found in the XML, or $null on failure.
#>
function Get-CustomizationData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )

    try {
        if (-not (Test-Path -Path $FilePath)) {
            throw "File not found: $FilePath"
        }

        $settings = [System.Collections.Generic.Dictionary[string,psobject]]::new([System.StringComparer]::InvariantCultureIgnoreCase)
        $xml = [xml](Get-Content -Path $FilePath -ErrorAction Stop)

        foreach ($parentNode in $xml.Customization.ChildNodes) {
            foreach ($settingNode in $parentNode.ChildNodes) {
                $key = $settingNode.Name
                $value = $null

                if ($settingNode.HasAttribute('INT_VALUE')) {
                    $value = $settingNode.GetAttribute('INT_VALUE')
                }
                elseif ($settingNode.HasAttribute('STRING_VALUE')) {
                    $value = $settingNode.GetAttribute('STRING_VALUE')
                }

                if ($value -ne $null -and !$settings.ContainsKey($key)) {
                    $settings[$key] = [PSCustomObject]@{
                        Value  = $value
                        Parent = $parentNode.Name
                    }
                }
            }
        }
        return $settings
    }
    catch {
        $errorMessage = "An error occurred while reading '$FilePath`.`n$($_.Exception.Message)"
        Write-Log "ERROR in Get-CustomizationData: $errorMessage"
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "File Read Error", "OK", "Error")
        return $null
    }
}

<#
.SYNOPSIS
    Writes a formatted line of text to a RichTextBox.
.DESCRIPTION
    Appends a line of text to a specified RichTextBox control with a given color and optional bold or underline styles.
.PARAMETER rtb
    The RichTextBox control to write to.
.PARAMETER text
    The text to append.
.PARAMETER color
    The color of the text.
.PARAMETER Underline
    A switch to make the text underlined.
.PARAMETER Bold
    A switch to make the text bold.
#>
function Write-ColoredLine {
    param(
        [System.Windows.Forms.RichTextBox]$rtb,
        [string]$text,
        [System.Drawing.Color]$color,
        [switch]$Underline,
        [switch]$Bold
    )
    $style = [System.Drawing.FontStyle]::Regular
    if ($Underline) { $style = $style -bor [System.Drawing.FontStyle]::Underline }
    if ($Bold) { $style = $style -bor [System.Drawing.FontStyle]::Bold }

    $rtb.SelectionColor = $color
    $rtb.SelectionFont = New-Object System.Drawing.Font($rtb.Font, $style)
    $rtb.AppendText($text + "`n")
    $rtb.SelectionColor = [System.Drawing.Color]::Black
    $rtb.SelectionFont = $rtb.Font
}

<#
.SYNOPSIS
    Writes a standard header line to a RichTextBox.
.DESCRIPTION
    A helper function that calls Write-ColoredLine to append a bold, dark blue header.
.PARAMETER rtb
    The RichTextBox control to write to.
.PARAMETER text
    The header text.
#>
function Write-HeaderLine {
    param([System.Windows.Forms.RichTextBox]$rtb,[string]$text)
    Write-ColoredLine $rtb $text ([System.Drawing.Color]::DarkBlue) -Bold
}

<#
.SYNOPSIS
    Writes a standard separator line to a RichTextBox.
.DESCRIPTION
    A helper function that calls Write-ColoredLine to append a gray separator line.
.PARAMETER rtb
    The RichTextBox control to write to.
#>
function Write-Separator {
    param([System.Windows.Forms.RichTextBox]$rtb)
    Write-ColoredLine $rtb "───────────────────────────────" ([System.Drawing.Color]::Gray)
}

<#
.SYNOPSIS
    Checks if a pre-requisite application is installed.
.DESCRIPTION
    Searches the registry's Uninstall keys to determine if an application matching the given pattern is installed. Includes a special case for checking the .NET Framework 4.8 version.
.PARAMETER pattern
    A regex pattern to match against the DisplayName of installed programs.
.RETURNS
    $true if the application is found, otherwise $false.
#>
function Check-PreReqInstalled {
    param([string]$pattern)

    if ($pattern -match "NET Framework 4\.8") {
        try {
            $releaseKey = Get-ItemPropertyValue `
                -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" `
                -Name Release -ErrorAction Stop
            if ($releaseKey -ge 528040) { return $true }
        } catch { return $false }
        return $false
    }

    $keys = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'

    $installed = Get-ItemProperty $keys -ErrorAction SilentlyContinue |
                 Select-Object -ExpandProperty DisplayName -ErrorAction SilentlyContinue

    $escapedPattern = $pattern.Replace('+', '\+')

    if ($null -ne $installed -and ($installed -match $escapedPattern)) { return $true } else { return $false }
}

<#
.SYNOPSIS
    Converts plain text from the System Requirements tab into an HTML report.
.DESCRIPTION
    Takes the text content from the three RichTextBoxes on the System Requirements tab and embeds them within a pre-formatted HTML structure, then saves it to a file.
.PARAMETER sysText
    The text content of the system requirements.
.PARAMETER preText
    The text content of the pre-requisites.
.PARAMETER resultsText
    The text content of the validation results.
.PARAMETER outputFile
    The full path where the HTML report will be saved.
#>
function Convert-ToHtmlReport {
    param([string]$sysText,[string]$preText,[string]$resultsText,[string]$outputFile)

    $html = @"
<html>
<head>
<style>
body { font-family: 'Segoe UI', sans-serif; margin:20px; }
h1 { color:black; }
h2 { color:darkblue; }
.green { color:green; font-weight:bold; }
.red { color:red; font-weight:bold; }
.note { color:gray; font-style:italic; }
.link { color:blue; text-decoration:underline; }
.separator { color:gray; }
pre { font-family: 'Segoe UI Emoji', 'Segoe UI', sans-serif; font-size: 14px; }
</style>
</head>
<body>
<h1>Litera Compare - System & Pre-Requisites Report</h1>

<h2>System Requirements</h2>
<pre>$sysText</pre>

<h2>Pre-Requisites</h2>
<pre>$preText</pre>

<h2>Validation Results</h2>
<pre>$resultsText</pre>

</body>
</html>
"@

    $html | Out-File -FilePath $outputFile -Encoding utf8
}

<#
.SYNOPSIS
    Displays the application's settings dialog window.
.DESCRIPTION
    Creates and shows a modal form for configuring application settings, such as backup paths and tab visibility. If the user saves the changes, it calls `Save-Settings` and `Apply-Settings`.
#>
function Show-SettingsDialog {
    $settingsForm = New-Object System.Windows.Forms.Form
    $settingsForm.Text = "Settings"
    $settingsForm.AutoSize = $true
    $settingsForm.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $settingsForm.StartPosition = "CenterParent"
    $settingsForm.FormBorderStyle = 'FixedDialog'
    $settingsForm.MaximizeBox = $false
    $settingsForm.MinimizeBox = $false

    $mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $mainLayout.Dock = 'Fill'
    $mainLayout.Padding = (New-Object System.Windows.Forms.Padding(10))
    $mainLayout.ColumnCount = 1
    $mainLayout.RowCount = 4
    $mainLayout.AutoSize = $true
    $mainLayout.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # Admin Group
    $mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null # Buttons
    $settingsForm.Controls.Add($mainLayout)

    # --- Paths Group ---
    $gbPaths = New-Object System.Windows.Forms.GroupBox; $gbPaths.Text = "Folder Paths"; $gbPaths.Dock = 'Fill'; $gbPaths.AutoSize = $true
    $pathsLayout = New-Object System.Windows.Forms.TableLayoutPanel; $pathsLayout.Dock = 'Fill'; $pathsLayout.ColumnCount = 3; $pathsLayout.RowCount = 4; $pathsLayout.AutoSize = $true
    $pathsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $pathsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $pathsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    4 | ForEach-Object { $pathsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) }
    $gbPaths.Controls.Add($pathsLayout)
    
    $lblPath = New-Object System.Windows.Forms.Label; $lblPath.Text = "Backup Path:"; $lblPath.Dock = 'Fill'; $lblPath.TextAlign = 'MiddleLeft'
    $txtBackupPath = New-Object System.Windows.Forms.TextBox; $txtBackupPath.Dock = 'Fill'; $txtBackupPath.Text = $script:backupRoot 
    $btnBrowse = New-AppButton -Text "Browse..." -Width 80 -Height 25
    $btnRestoreDefaultPath = New-AppButton -Text "Restore Default" -AutoSize
    $lblLogPath = New-Object System.Windows.Forms.Label; $lblLogPath.Text = "Logs Path:"; $lblLogPath.Dock = 'Fill'; $lblLogPath.TextAlign = 'MiddleLeft'
    $txtLogPath = New-Object System.Windows.Forms.TextBox; $txtLogPath.Dock = 'Fill'; $txtLogPath.Text = $script:logRoot
    $btnBrowseLog = New-AppButton -Text "Browse..." -Width 80 -Height 25
    $btnRestoreDefaultLog = New-AppButton -Text "Restore Default" -AutoSize

    $pathsLayout.Controls.Add($lblPath, 0, 0); $pathsLayout.Controls.Add($txtBackupPath, 1, 0); $pathsLayout.Controls.Add($btnBrowse, 2, 0)
    $pathsLayout.Controls.Add($btnRestoreDefaultPath, 1, 1)
    $pathsLayout.Controls.Add($lblLogPath, 0, 2); $pathsLayout.Controls.Add($txtLogPath, 1, 2); $pathsLayout.Controls.Add($btnBrowseLog, 2, 2)
    $pathsLayout.Controls.Add($btnRestoreDefaultLog, 1, 3)

    # --- Tabs Group ---
    $gbTabs = New-Object System.Windows.Forms.GroupBox; $gbTabs.Text = "Visible Tabs"; $gbTabs.Dock = 'Fill'; $gbTabs.AutoSize = $true
    $gbTabsLayout = New-Object System.Windows.Forms.TableLayoutPanel; $gbTabsLayout.Dock = 'Fill'; $gbTabsLayout.AutoSize = $true; $gbTabsLayout.ColumnCount = 1
    $gbTabsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $gbTabsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $gbTabs.Controls.Add($gbTabsLayout)

    $tabGroupsContainer = New-Object System.Windows.Forms.FlowLayoutPanel; $tabGroupsContainer.Dock = 'Fill'; $tabGroupsContainer.FlowDirection = 'LeftToRight'; $tabGroupsContainer.AutoSize = $true
    $gbTabsLayout.Controls.Add($tabGroupsContainer, 0, 0)

    $gbCore = New-Object System.Windows.Forms.GroupBox; $gbCore.Text = "Core Features"; $gbCore.AutoSize = $true
    $corePanel = New-Object System.Windows.Forms.FlowLayoutPanel; $corePanel.Dock = 'Fill'; $corePanel.FlowDirection = 'TopDown'; $corePanel.AutoSize = $true
    $gbCore.Controls.Add($corePanel)
    $tabGroupsContainer.Controls.Add($gbCore)

    $gbTools = New-Object System.Windows.Forms.GroupBox; $gbTools.Text = "Tools & Info"; $gbTools.AutoSize = $true
    $toolsPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $toolsPanel.Dock = 'Fill'; $toolsPanel.FlowDirection = 'TopDown'; $toolsPanel.AutoSize = $true
    $gbTools.Controls.Add($toolsPanel)
    $tabGroupsContainer.Controls.Add($gbTools)

    $chkShowCleanup = New-Object System.Windows.Forms.CheckBox; $chkShowCleanup.Text = "User File Cleanup"; $chkShowCleanup.Tag = "ShowCleanupTab"; $chkShowCleanup.AutoSize = $true
    $chkShowBackupRestore = New-Object System.Windows.Forms.CheckBox; $chkShowBackupRestore.Text = "Backup & Restore"; $chkShowBackupRestore.Tag = "ShowBackupRestoreTab"; $chkShowBackupRestore.AutoSize = $true
    $chkShowInstallUninstall = New-Object System.Windows.Forms.CheckBox; $chkShowInstallUninstall.Text = "Install/Uninstall"; $chkShowInstallUninstall.Tag = "ShowInstallUninstallTab"; $chkShowInstallUninstall.AutoSize = $true
    $chkShowSysReq = New-Object System.Windows.Forms.CheckBox; $chkShowSysReq.Text = "System Requirements"; $chkShowSysReq.Tag = "ShowSysReqTab"; $chkShowSysReq.AutoSize = $true
    $chkShowLicense = New-Object System.Windows.Forms.CheckBox; $chkShowLicense.Text = "License Info"; $chkShowLicense.Tag = "ShowLicenseTab"; $chkShowLicense.AutoSize = $true
    $chkShowCompareCustom = New-Object System.Windows.Forms.CheckBox; $chkShowCompareCustom.Text = "Compare Customizations"; $chkShowCompareCustom.Tag = "ShowCompareCustomTab"; $chkShowCompareCustom.AutoSize = $true
    $chkShowAddinMgmt = New-Object System.Windows.Forms.CheckBox; $chkShowAddinMgmt.Text = "Office Add-in Mgmt"; $chkShowAddinMgmt.Tag = "ShowAddinMgmtTab"; $chkShowAddinMgmt.AutoSize = $true
    $chkShowLogViewer = New-Object System.Windows.Forms.CheckBox; $chkShowLogViewer.Text = "Log Viewer"; $chkShowLogViewer.Tag = "ShowLogViewerTab"; $chkShowLogViewer.AutoSize = $true

    $corePanel.Controls.AddRange(@($chkShowCleanup, $chkShowBackupRestore, $chkShowInstallUninstall))
    $toolsPanel.Controls.AddRange(@($chkShowSysReq, $chkShowLicense, $chkShowCompareCustom, $chkShowAddinMgmt, $chkShowLogViewer))

    $allTabCheckBoxes = @(
        $chkShowCleanup, $chkShowBackupRestore, $chkShowInstallUninstall,
        $chkShowSysReq, $chkShowLicense, $chkShowCompareCustom, $chkShowAddinMgmt, $chkShowLogViewer
    )

    $restoreTabsContainer = New-Object System.Windows.Forms.TableLayoutPanel
    $restoreTabsContainer.Dock = 'Fill'; $restoreTabsContainer.ColumnCount = 3; $restoreTabsContainer.RowCount = 1; $restoreTabsContainer.AutoSize = $true
    $restoreTabsContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $restoreTabsContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $restoreTabsContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

    $btnRestoreAllDefaults = New-AppButton -Text "Restore Defaults" -AutoSize

    $restoreTabsContainer.Controls.Add($btnRestoreAllDefaults, 1, 0)
    $gbTabsLayout.Controls.Add($restoreTabsContainer, 0, 1)

    $bottomButtonContainer = New-Object System.Windows.Forms.TableLayoutPanel
    $bottomButtonContainer.Dock = 'Fill'; $bottomButtonContainer.ColumnCount = 3; $bottomButtonContainer.RowCount = 1; $bottomButtonContainer.AutoSize = $true
    $bottomButtonContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $bottomButtonContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $bottomButtonContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))

    $buttonPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $buttonPanel.ColumnCount = 2; $buttonPanel.RowCount = 1; $buttonPanel.AutoSize = $true
    $buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $buttonPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))

    $btnSave = New-AppButton -Text "Save" -Width 100 -BackColor 'LightGreen'
    $btnSave.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnCancel = New-AppButton -Text "Cancel" -Width 100 -BackColor 'MistyRose'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $btnCancel.Margin = (New-Object System.Windows.Forms.Padding(10, 0, 0, 0))

    $buttonPanel.Controls.Add($btnSave, 0, 0)
    $buttonPanel.Controls.Add($btnCancel, 1, 0)
    $bottomButtonContainer.Controls.Add($buttonPanel, 1, 0)

    # --- NEW: Admin Group ---
    $gbAdmin = New-Object System.Windows.Forms.GroupBox; $gbAdmin.Text = "Administrator Options"; $gbAdmin.Dock = 'Fill'; $gbAdmin.AutoSize = $true
    $adminLayout = New-Object System.Windows.Forms.FlowLayoutPanel; $adminLayout.Dock = 'Fill'; $adminLayout.FlowDirection = 'TopDown'; $adminLayout.AutoSize = $true
    $gbAdmin.Controls.Add($adminLayout)

    $chkAllowExe = New-Object System.Windows.Forms.CheckBox; $chkAllowExe.Text = "Allow .exe file installation in Install/Uninstall tab"; $chkAllowExe.AutoSize = $true
    $chkShowAllPrograms = New-Object System.Windows.Forms.CheckBox; $chkShowAllPrograms.Text = "Show all programs in Uninstall list (not just Litera)"; $chkShowAllPrograms.AutoSize = $true

    $adminLayout.Controls.AddRange(@($chkAllowExe, $chkShowAllPrograms))

    if (-not (Test-IsAdmin)) {
        $gbAdmin.Visible = $false
    }

    # --- Add Groups to Main Layout ---
    $mainLayout.Controls.Add($gbPaths, 0, 0)
    $mainLayout.Controls.Add($gbTabs, 0, 1)
    $mainLayout.Controls.Add($gbAdmin, 0, 2)
    $mainLayout.Controls.Add($bottomButtonContainer, 0, 3)

    # Load current admin settings
    if (Test-IsAdmin) {
        $chkAllowExe.Checked = $script:settings.AllowExeInstallation
        $chkShowAllPrograms.Checked = $script:settings.ShowAllPrograms
    }

    $btnBrowse.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Select a folder for backups"
        $folderDialog.SelectedPath = $txtBackupPath.Text
        if ($folderDialog.ShowDialog() -eq 'OK') { $txtBackupPath.Text = $folderDialog.SelectedPath }
        $folderDialog.Dispose()
    })
    $btnRestoreDefaultPath.Add_Click({ $txtBackupPath.Text = $defaultBackupRoot })

    $btnBrowseLog.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Select a folder for installation logs"
        $folderDialog.SelectedPath = $txtLogPath.Text
        if ($folderDialog.ShowDialog() -eq 'OK') { $txtLogPath.Text = $folderDialog.SelectedPath }
        $folderDialog.Dispose()
    })
    $btnRestoreDefaultLog.Add_Click({ $txtLogPath.Text = $defaultLogRoot })

    $btnRestoreAllDefaults.Add_Click({
        $txtBackupPath.Text = $defaultBackupRoot
        $txtLogPath.Text = $defaultLogRoot
        foreach ($chk in $allTabCheckBoxes) { $chk.Checked = $true }
    })

    $settingsForm.AcceptButton = $btnSave
    $settingsForm.CancelButton = $btnCancel

    foreach ($chk in $allTabCheckBoxes) {
        $settingName = $chk.Tag
        $chk.Checked = $script:settings.TabVisibility[$settingName]
    }

    if ($settingsForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:settings.BackupRoot = $txtBackupPath.Text.Trim()
        $script:settings.LogRoot = $txtLogPath.Text.Trim()

        foreach ($chk in $allTabCheckBoxes) {
            $settingName = $chk.Tag
            $script:settings.TabVisibility[$settingName] = $chk.Checked
        }

        if (Test-IsAdmin) {
            $script:settings.AllowExeInstallation = $chkAllowExe.Checked
            $script:settings.ShowAllPrograms = $chkShowAllPrograms.Checked
        }

        Save-Settings
        Apply-Settings
        Populate-BackupRestoreLists
        $script:LblBackupStatus.Text = "Settings updated and applied."
        $script:LblBackupStatus.ForeColor = 'Blue'
    }
    $settingsForm.Dispose()
}

<#
.SYNOPSIS
    Displays a dialog notifying the user that a new version is available.
.DESCRIPTION
    Shows a form with the new version number and release notes, providing a "Download Now" button.
.PARAMETER Version
    The new version string (e.g., "1.1.0").
.PARAMETER Notes
    The release notes for the new version.
.RETURNS
    The DialogResult from the form ('Yes' for download, 'No' for later).
#>
function Show-UpdateDialog {
    param(
        [string]$Version,
        [string]$Notes
    )

    $updateForm = New-Object System.Windows.Forms.Form
    $updateForm.Text = "Update Available"
    $updateForm.Size = New-Object System.Drawing.Size(450, 320)
    $updateForm.StartPosition = "CenterParent"
    $updateForm.FormBorderStyle = 'FixedDialog'
    $updateForm.MaximizeBox = $false
    $updateForm.MinimizeBox = $false

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = (New-Object System.Windows.Forms.Padding(15))
    $layout.ColumnCount = 1
    $layout.RowCount = 3
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $updateForm.Controls.Add($layout)

    $lblHeader = New-Object System.Windows.Forms.Label
    $lblHeader.Text = "A new version of the tool is available!"
    $lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblHeader.Dock = 'Fill'
    $lblHeader.TextAlign = 'MiddleCenter'
    $lblHeader.Padding = (New-Object System.Windows.Forms.Padding(0, 0, 0, 10))
    $layout.Controls.Add($lblHeader, 0, 0)

    $rtbInfo = New-Object System.Windows.Forms.RichTextBox
    $rtbInfo.Dock = 'Fill'; $rtbInfo.ReadOnly = $true; $rtbInfo.BorderStyle = 'None'
    $rtbInfo.BackColor = $updateForm.BackColor
    $rtbInfo.ScrollBars = 'Vertical'

    $rtbInfo.SelectionAlignment = 'Center'
    $rtbInfo.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $rtbInfo.AppendText("Version: $Version`n`n")

    $rtbInfo.SelectionAlignment = 'Left'
    $rtbInfo.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 9)
    $rtbInfo.AppendText("Release Notes:`n$Notes")

    $layout.Controls.Add($rtbInfo, 0, 1)

    $buttonContainer = New-Object System.Windows.Forms.TableLayoutPanel
    $buttonContainer.Dock = 'Fill'; $buttonContainer.ColumnCount = 3; $buttonContainer.RowCount = 1; $buttonContainer.AutoSize = $true
    $buttonContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $buttonContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $buttonContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    
    $buttonFlowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonFlowPanel.FlowDirection = 'LeftToRight'; $buttonFlowPanel.AutoSize = $true
    $buttonFlowPanel.Padding = (New-Object System.Windows.Forms.Padding(0, 10, 0, 0))

    $btnYes = New-AppButton -Text "Download Now" -Width 120 -Height 35 -BackColor 'LightGreen'
    $btnYes.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $btnNo = New-AppButton -Text "Later" -Width 120 -Height 35 -BackColor 'LightCoral'
    $btnNo.DialogResult = [System.Windows.Forms.DialogResult]::No

    $buttonFlowPanel.Controls.AddRange(@($btnYes, $btnNo))
    $buttonContainer.Controls.Add($buttonFlowPanel, 1, 0)
    $layout.Controls.Add($buttonContainer, 0, 2)

    $updateForm.AcceptButton = $btnYes
    $updateForm.CancelButton = $btnNo

    $dialogResult = $updateForm.ShowDialog()
    $updateForm.Dispose()
    return $dialogResult
}

<#
.SYNOPSIS
    Displays a dialog informing the user they are up-to-date.
.DESCRIPTION
    Shows a simple form indicating that the user's version of the tool matches the latest version found on the server.
.PARAMETER FoundVersion
    The version string found on the server.
#>
function Show-UpToDateDialog {
    param(
        [string]$FoundVersion
    )
    $infoForm = New-Object System.Windows.Forms.Form
    $infoForm.Text = "Up to Date"
    $infoForm.Size = New-Object System.Drawing.Size(450, 220)
    $infoForm.StartPosition = "CenterParent"
    $infoForm.FormBorderStyle = 'FixedDialog'
    $infoForm.MaximizeBox = $false
    $infoForm.MinimizeBox = $false

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = (New-Object System.Windows.Forms.Padding(15))
    $layout.ColumnCount = 1
    $layout.RowCount = 3
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $infoForm.Controls.Add($layout)

    $lblHeader = New-Object System.Windows.Forms.Label
    $lblHeader.Text = "You have the latest version of the tool."
    $lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblHeader.Dock = 'Fill'
    $lblHeader.TextAlign = 'MiddleCenter'
    $layout.Controls.Add($lblHeader, 0, 0)

    $rtbInfo = New-Object System.Windows.Forms.RichTextBox
    $rtbInfo.Dock = 'Fill'; $rtbInfo.ReadOnly = $true; $rtbInfo.BorderStyle = 'None'
    $rtbInfo.BackColor = $infoForm.BackColor
    $rtbInfo.SelectionAlignment = 'Center'
    $rtbInfo.SelectionFont = New-Object System.Drawing.Font("Segoe UI", 10)
    $rtbInfo.AppendText("Your version: $currentVersion`nServer version: $FoundVersion")
    $layout.Controls.Add($rtbInfo, 0, 1)

    $buttonContainer = New-Object System.Windows.Forms.TableLayoutPanel
    $buttonContainer.Dock = 'Fill'; $buttonContainer.ColumnCount = 3; $buttonContainer.RowCount = 1; $buttonContainer.AutoSize = $true
    $buttonContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $buttonContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $buttonContainer.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    
    $btnOk = New-AppButton -Text "OK" -Width 120 -Height 35
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $buttonContainer.Controls.Add($btnOk, 1, 0)
    $layout.Controls.Add($buttonContainer, 0, 2)

    $infoForm.AcceptButton = $btnOk

    $infoForm.ShowDialog() | Out-Null
    $infoForm.Dispose()
}

<#
.SYNOPSIS
    A factory function for creating consistently styled Button controls.
.DESCRIPTION
    Simplifies the creation of `System.Windows.Forms.Button` objects with a common style.
#>
function New-AppButton {
    [CmdletBinding(DefaultParameterSetName='FixedSize')]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Text,

        [Parameter(Mandatory=$true, ParameterSetName='FixedSize')]
        [int]$Width,

        [Parameter(ParameterSetName='FixedSize')]
        [int]$Height = 30,

        [Parameter(Mandatory=$true, ParameterSetName='AutoSize')]
        [switch]$AutoSize,

        [Parameter()]
        [string]$BackColor,

        [Parameter()]
        [bool]$Enabled = $true,

        [Parameter()]
        [System.Windows.Forms.Padding]$Margin
    )
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Enabled = $Enabled
    if ($BackColor) {
        $button.BackColor = $BackColor
    }
    if ($PSBoundParameters.ContainsKey('Margin')) {
        $button.Margin = $Margin
    }
    if ($PSCmdlet.ParameterSetName -eq 'AutoSize') {
        $button.AutoSize = $true
    } else {
        $button.Size = New-Object System.Drawing.Size($Width, $Height)
    }
    return $button
}

<#
.SYNOPSIS
    A factory function for creating consistently styled Label controls for status messages.
.DESCRIPTION
    Simplifies the creation of `System.Windows.Forms.Label` objects used for status feedback throughout the UI.
#>
function New-StatusLabel {
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$Text,
        [string]$InitialColor,
        [string]$Dock,
        [string]$TextAlign = 'MiddleLeft',
        [switch]$AutoSize,
        [bool]$Visible = $true,
        [System.Windows.Forms.Padding]$Margin
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    if ($InitialColor) { $label.ForeColor = $InitialColor }
    if ($Dock) { $label.Dock = $Dock }
    $label.TextAlign = $TextAlign
    $label.AutoSize = $AutoSize.IsPresent
    $label.Visible = $Visible
    if ($PSBoundParameters.ContainsKey('Margin')) {
        $label.Margin = $Margin
    }
    return $label
}

<#
.SYNOPSIS
    Creates and configures all UI elements for the "User File Cleanup" tab.
.DESCRIPTION
    This function encapsulates the creation of all controls, panels, and layouts for the cleanup tab.
    It uses the script scope for control variables so they can be accessed by event handlers later in the script.
.RETURNS
    A `System.Windows.Forms.TableLayoutPanel` containing the complete UI for the tab.
#>
function Create-CleanupTabUI {
    $script:CleanupLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:CleanupLayout.Dock = 'Fill'
    $script:CleanupLayout.BackColor = [System.Drawing.SystemColors]::Window
    $script:CleanupLayout.ColumnCount = 2
    $script:CleanupLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $script:CleanupLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $script:CleanupLayout.RowCount = 5
    $script:CleanupLayout.RowCount = 6
    $script:CleanupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    [void]$script:CleanupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:CleanupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$script:CleanupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:CleanupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # For progress bar
    [void]$script:CleanupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) # For status label

    $script:GbUserSelection = New-Object System.Windows.Forms.GroupBox
    $script:GbUserSelection.Text = "Administrator: Select User"
    $script:GbUserSelection.Dock = 'Top'
    $script:GbUserSelection.AutoSize = $true
    $script:GbUserSelection.Padding = (New-Object System.Windows.Forms.Padding(10, 5, 10, 10))
 
    $script:UserSelectionLayout = New-Object System.Windows.Forms.FlowLayoutPanel
    $script:UserSelectionLayout.Dock = 'Top'
    $script:UserSelectionLayout.AutoSize = $true
    $script:UserSelectionLayout.WrapContents = $false

    $script:LblSelectUser = New-Object System.Windows.Forms.Label; $script:LblSelectUser.Text = "Select User Profile:"; $script:LblSelectUser.Margin = '0,5,5,0'; $script:LblSelectUser.AutoSize = $true
    $script:ComboUsers = New-Object System.Windows.Forms.ComboBox; $script:ComboUsers.DropDownStyle = 'DropDownList'; $script:ComboUsers.Width = 200; $script:ComboUsers.Margin = '0,0,5,0'
    $script:BtnLoadUserFolders = New-AppButton -Text "Load User's Folders" -AutoSize -Margin ([System.Windows.Forms.Padding]::new(0, 0, 0, 0))
    [void]$script:UserSelectionLayout.Controls.AddRange(@($script:LblSelectUser, $script:ComboUsers, $script:BtnLoadUserFolders))
    [void]$script:GbUserSelection.Controls.Add($script:UserSelectionLayout)

    $script:LocalAppPanel = New-Object System.Windows.Forms.Panel; $script:LocalAppPanel.Dock = 'Fill'
    $script:AppDataPanel = New-Object System.Windows.Forms.Panel; $script:AppDataPanel.Dock = 'Fill'

    $script:ClbLocalAppData = New-Object System.Windows.Forms.CheckedListBox; $script:ClbLocalAppData.Dock = 'Fill'
    $script:ClbAppData = New-Object System.Windows.Forms.CheckedListBox; $script:ClbAppData.Dock = 'Fill'
    $script:ChkSelectAllLocal = New-Object System.Windows.Forms.CheckBox; $script:ChkSelectAllLocal.Text = "Select All"; $script:ChkSelectAllLocal.Dock = 'Top'
    $script:ChkSelectAllRoaming = New-Object System.Windows.Forms.CheckBox; $script:ChkSelectAllRoaming.Text = "Select All"; $script:ChkSelectAllRoaming.Dock = 'Top'

    [void]$script:LocalAppPanel.Controls.Add($script:ClbLocalAppData)
    [void]$script:LocalAppPanel.Controls.Add($script:ChkSelectAllLocal)
    [void]$script:AppDataPanel.Controls.Add($script:ClbAppData)
    [void]$script:AppDataPanel.Controls.Add($script:ChkSelectAllRoaming)

    $script:LblLocalAppData = New-Object System.Windows.Forms.Label; $script:LblLocalAppData.Text = "User LocalAppData Litera"; $script:LblLocalAppData.Dock = 'Top'
    $script:LblAppData = New-Object System.Windows.Forms.Label; $script:LblAppData.Text = "User AppData Litera"; $script:LblAppData.Dock = 'Top'

    [void]$script:CleanupLayout.Controls.Add($script:GbUserSelection, 0, 0); $script:CleanupLayout.SetColumnSpan($script:GbUserSelection, 2)
    [void]$script:CleanupLayout.Controls.Add($script:LblLocalAppData, 0, 1)
    [void]$script:CleanupLayout.Controls.Add($script:LblAppData, 1, 1)
    [void]$script:CleanupLayout.Controls.Add($script:LocalAppPanel, 0, 2)
    [void]$script:CleanupLayout.Controls.Add($script:AppDataPanel, 1, 2)

    $script:CleanupButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $script:CleanupButtonsPanel.Dock = 'Fill'
    $script:CleanupButtonsPanel.FlowDirection = 'LeftToRight'
    $script:CleanupLayout.SetColumnSpan($script:CleanupButtonsPanel, 2)
    [void]$script:CleanupLayout.Controls.Add($script:CleanupButtonsPanel, 0, 3)

    $script:BtnRefreshCleanup    = New-AppButton -Text "Refresh Folders" -AutoSize
    $script:BtnCleanupSelected   = New-AppButton -Text "Cleanup Selected" -AutoSize -BackColor 'MistyRose'
    $script:ToolTip.SetToolTip($script:BtnCleanupSelected, "Deletes the entire contents of the selected folders in the lists above.")
    $script:BtnClearLogs         = New-AppButton -Text "Clear All Logs" -AutoSize -BackColor 'Red'
    $script:ToolTip.SetToolTip($script:BtnClearLogs, "Deletes all *.txt log files from Litera user folders and DMS log folders.")
    $script:BtnClearDMSLogs      = New-AppButton -Text "Clear DMS Logs" -AutoSize
    $script:ToolTip.SetToolTip($script:BtnClearDMSLogs, "Deletes logs from the DMS log folder and the dms_debug_log.txt file.")
    $script:BtnDeleteDMSConfig   = New-AppButton -Text "Delete User DMS Config" -AutoSize
    $script:ToolTip.SetToolTip($script:BtnDeleteDMSConfig, "Deletes the user-specific DMS configuration file (Litera.iManage.Work.V2.Config).")
    [void]$script:CleanupButtonsPanel.Controls.AddRange(@($script:BtnRefreshCleanup, $script:BtnCleanupSelected, $script:BtnClearLogs, $script:BtnClearDMSLogs, $script:BtnDeleteDMSConfig))

    $script:ProgressCleanup = New-Object System.Windows.Forms.ProgressBar
    $script:ProgressCleanup.Dock = 'Fill'
    $script:ProgressCleanup.Visible = $false
    $script:CleanupLayout.SetColumnSpan($script:ProgressCleanup, 2)
    [void]$script:CleanupLayout.Controls.Add($script:ProgressCleanup, 0, 4)

    $script:LblStatus = New-StatusLabel -Text "Ready." -Dock 'Fill'
    $script:CleanupLayout.SetColumnSpan($script:LblStatus, 2)
    [void]$script:CleanupLayout.Controls.Add($script:LblStatus, 0, 5)

    return $script:CleanupLayout
}

<#
.SYNOPSIS
    Defines the UI and logic for the "User File Cleanup" tab.
.DESCRIPTION
    This section builds the UI for the cleanup tab, which allows for selecting a user profile and cleaning various Litera-related files and folders, such as logs and DMS configurations.
#>
#==============================================================================
# --- UI: User File Cleanup Tab ---
#==============================================================================
$cleanupTabPanel = Create-CleanupTabUI
$script:TabCleanup.Controls.Add($cleanupTabPanel)

function Test-IsAdmin {
    return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

<#
.SYNOPSIS
    Populates a ComboBox with a list of user profiles on the system.
.DESCRIPTION
    Clears the target ComboBox and fills it with user profiles found via WMI. The current user is always listed first.
.PARAMETER ComboBox
    The `System.Windows.Forms.ComboBox` control to populate.
#>
function Populate-UserComboBox {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.ComboBox]$ComboBox
    )
    $ComboBox.Items.Clear()
    $currentUser = $env:USERNAME
    $ComboBox.Items.Add("$currentUser (Current User)")

    # Get all user profiles
    $userProfiles = Get-CimInstance -ClassName Win32_UserProfile | Where-Object { $_.Special -eq $false -and (Test-Path $_.LocalPath) }
    foreach ($profile in $userProfiles) {
        $userName = $profile.LocalPath.Split('\')[-1]
        if ($userName -ne $currentUser) {
            $ComboBox.Items.Add($userName)
        }
    }
    $ComboBox.SelectedIndex = 0
}

if (Test-IsAdmin) {
    Populate-UserComboBox -ComboBox $script:ComboUsers
} else {
    $script:GbUserSelection.Visible = $false
}

<#
.SYNOPSIS
    Populates the list boxes in the Cleanup tab with directories from the selected user's AppData folders.
#>
function Populate-CleanupLists {
    try {
        $userPaths = Get-SelectedUserPaths
        $userName = $script:selectedUser

        $script:ClbLocalAppData.Items.Clear()
        $script:ClbAppData.Items.Clear()
        if (Test-Path $userPaths.literaLocalAppData) {
            Get-ChildItem -Path $userPaths.literaLocalAppData -Directory | ForEach-Object { [void]$script:ClbLocalAppData.Items.Add($_.FullName, $false) }
        }
        if (Test-Path $userPaths.literaAppData) {
            $foldersToExclude = @('Change-Pro Styles', 'Compare', 'PersonalCleaningProfiles')
            Get-ChildItem -Path $userPaths.literaAppData -Directory |
                Where-Object { $_.Name -notin $foldersToExclude } |
                ForEach-Object { [void]$script:ClbAppData.Items.Add($_.FullName, $false) }
        }
        $script:LblStatus.Text = "Folder lists refreshed for user '$userName'."
        $script:LblStatus.ForeColor = 'Green'
    } catch {
        $script:LblStatus.Text = "Error loading folder lists: $($_.Exception.Message)"
        $script:LblStatus.ForeColor = 'Red'
    }
}

<#
.SYNOPSIS
    Enables or disables all action buttons on the Cleanup tab.
.PARAMETER Enabled
    $true to enable the buttons, $false to disable them.
#>
function Set-CleanupButtonsState {
    param([bool]$Enabled)
    $script:CleanupButtonsPanel.Controls | ForEach-Object { $_.Enabled = $Enabled }
}

<#
.SYNOPSIS
    Deletes the contents of specified folders using robocopy.
.DESCRIPTION
    This function efficiently deletes the contents of a list of folders by creating a temporary empty directory and using `robocopy /MIR` to mirror it to the target folders. This is often faster and more reliable than `Remove-Item -Recurse`.
.PARAMETER FoldersToClean
    An array of full paths to the folders whose contents should be deleted.
.RETURNS
    A PSCustomObject with a `Failures` property, which is an array of strings detailing any folders that failed to clean.
#>
function Invoke-FolderContentCleanup {
    param(
        [Parameter(Mandatory=$true)]
        [array]$FoldersToClean
    )
    $failures = [System.Collections.Generic.List[string]]::new()
    $emptyDir = $null
    try {
        $emptyDir = Join-Path $env:TEMP ([Guid]::NewGuid().ToString())
        New-Item -Path $emptyDir -ItemType Directory -Force -ErrorAction Stop | Out-Null

        foreach ($folder in $FoldersToClean) {
            if (($folder -is [string]) -and (-not [string]::IsNullOrWhiteSpace($folder)) -and (Test-Path $folder -PathType Container)) {
                Write-Log "Cleaning folder contents: $folder using robocopy."
                $process = Start-Process robocopy -ArgumentList """$emptyDir"" ""$folder"" /MIR /NJH /NJS" -Wait -PassThru -WindowStyle Hidden
                if ($process.ExitCode -ge 8) {
                    $failures.Add("'$folder' (Robocopy exit code: $($process.ExitCode))")
                }
            } else {
                Write-Log "SKIPPING cleanup for invalid path: '$folder'"
            }
        }
    }
    finally {
        if ($emptyDir -and (Test-Path $emptyDir)) {
            Remove-Item -Path $emptyDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    return [PSCustomObject]@{ Failures = $failures }
}

$script:BtnCleanupSelected.Add_Click({
    $foldersToClean = @($script:ClbLocalAppData.CheckedItems) + @($script:ClbAppData.CheckedItems)
    if ($foldersToClean.Count -eq 0) {
        $script:LblStatus.Text = "No folders selected for cleanup."; $script:LblStatus.ForeColor = 'Orange'; return
    }
    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the CONTENTS of the selected folders? This cannot be undone.", "Confirm Cleanup", "YesNo", "Warning")
    if ($confirmation -ne 'Yes') {
        $script:LblStatus.Text = "Cleanup cancelled."; $script:LblStatus.ForeColor = 'Gray'; return
    }

    try {
        Set-CleanupButtonsState -Enabled $false
        $script:LblStatus.Text = "Cleaning selected folder contents..."
        $script:LblStatus.ForeColor = 'Black'
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $script:ProgressCleanup.Style = 'Marquee'; $script:ProgressCleanup.MarqueeAnimationSpeed = 30; $script:ProgressCleanup.Visible = $true

        $result = Invoke-FolderContentCleanup -FoldersToClean $foldersToClean

        if ($result.Failures.Count -gt 0) {
            $script:LblStatus.Text = "Cleanup completed with $($result.Failures.Count) error(s)."
            $script:LblStatus.ForeColor = 'Red'
            Write-Log "ERROR during folder cleanup. Failures: $($result.Failures -join ', ')"
            [System.Windows.Forms.MessageBox]::Show("The following folders could not be cleaned:`n`n$($result.Failures -join "`n")", "Cleanup Error", "OK", "Error")
        } else {
            $script:LblStatus.Text = "Selected folder contents have been cleaned."
            $script:LblStatus.ForeColor = 'Green'
        }
        Populate-CleanupLists

    } catch {
        $script:LblStatus.Text = "A critical error occurred during cleanup: $($_.Exception.Message)"
        $script:LblStatus.ForeColor = 'Red'
        Write-Log "ERROR during folder cleanup: $($_.Exception.Message)"
    } finally {
        Set-CleanupButtonsState -Enabled $true
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
        $script:ProgressCleanup.Visible = $false
    }
})

<#
.SYNOPSIS
    Deletes the contents of specified folders using robocopy.
.DESCRIPTION
    This function efficiently deletes the contents of a list of folders by creating a temporary empty directory and using `robocopy /MIR` to mirror it to the target folders. This is often faster and more reliable than `Remove-Item -Recurse`.
.PARAMETER FoldersToClean
    An array of full paths to the folders whose contents should be deleted.
.RETURNS
    A PSCustomObject with a `Failures` property, which is an array of strings detailing any folders that failed to clean.
#>
function Invoke-FolderContentCleanup {
    param(
        [Parameter(Mandatory=$true)]
        [array]$FoldersToClean
    )
    $failures = [System.Collections.Generic.List[string]]::new()
    $emptyDir = $null
    try {
        $emptyDir = Join-Path $env:TEMP ([Guid]::NewGuid().ToString())
        New-Item -Path $emptyDir -ItemType Directory -Force -ErrorAction Stop | Out-Null

        foreach ($folder in $FoldersToClean) {
            if (($folder -is [string]) -and (-not [string]::IsNullOrWhiteSpace($folder)) -and (Test-Path $folder -PathType Container)) {
                Write-Log "Cleaning folder contents: $folder using robocopy."
                $process = Start-Process robocopy -ArgumentList """$emptyDir"" ""$folder"" /MIR /NJH /NJS" -Wait -PassThru -WindowStyle Hidden
                if ($process.ExitCode -ge 8) {
                    $failures.Add("'$folder' (Robocopy exit code: $($process.ExitCode))")
                }
            } else {
                Write-Log "SKIPPING cleanup for invalid path: '$folder'"
            }
        }
    }
    finally {
        if ($emptyDir -and (Test-Path $emptyDir)) {
            Remove-Item -Path $emptyDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    return [PSCustomObject]@{ Failures = $failures }
}

<#
.SYNOPSIS
    Gathers a list of all known Litera log files for a given user.
.DESCRIPTION
    Scans a predefined set of directories within the user's AppData and LocalAppData folders to find all potential log files.
.PARAMETER UserPaths
    A PSCustomObject containing the user's AppData paths, as returned by `Get-SelectedUserPaths`.
.RETURNS
    An array of unique `System.IO.FileInfo` objects representing the log files found.
#>
function Get-AllUserLogFiles {
    param(
        [Parameter(Mandatory=$true)]
        $UserPaths
    )
    $filesToDelete = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

    $pathsToScan = @(
        @{ Path = (Join-Path $UserPaths.appData "Litera"); Recurse = $false },
        @{ Path = (Join-Path $UserPaths.localAppData "Litera"); Recurse = $false },
        @{ Path = (Join-Path $UserPaths.appData "Litera\DMS\Log"); Recurse = $true },
        @{ Path = (Join-Path $UserPaths.localAppData "Litera\IOW\Logs"); Recurse = $true },
        @{ Path = (Join-Path $UserPaths.localAppData "Litera\ShellExtension"); Recurse = $true },
        @{ Path = (Join-Path $UserPaths.localAppData "Litera\CpClip"); Recurse = $true },
        @{ Path = (Join-Path $UserPaths.localAppData "Litera\OCR\Logs"); Recurse = $true },
        @{ Path = (Join-Path $UserPaths.localAppData "Litera\lcp_main.exe\DB\reports"); Recurse = $true },
        @{ Path = (Join-Path $UserPaths.localAppData "Litera\lcp_ppt.exe\DB\reports"); Recurse = $true },
        @{ Path = (Join-Path $UserPaths.localAppData "Litera\lcp_pdfcmp.exe\DB\reports"); Recurse = $true }
    )

    foreach ($spec in $pathsToScan) {
        if (Test-Path $spec.Path -PathType Container) {
            $foundFiles = @(Get-ChildItem -Path $spec.Path -File -Recurse:$spec.Recurse -ErrorAction SilentlyContinue)
            if ($foundFiles.Count -gt 0) {
                foreach ($file in $foundFiles) { $filesToDelete.Add($file) }
            }
        }
    }
    # Return a de-duplicated list of files
    return @($filesToDelete | Group-Object FullName | ForEach-Object { $_.Group[0] })
}

<#
.SYNOPSIS
    Deletes a list of files, prompting the user to continue or cancel if an error occurs.
.DESCRIPTION
    Iterates through a list of files, attempting to delete each one. If a file cannot be deleted, it displays a `MessageBox` asking the user whether to skip the file and continue, or cancel the entire operation.
.PARAMETER Files
    An array of `System.IO.FileInfo` objects to be deleted.
.RETURNS
    A PSCustomObject with properties `Cancelled` (boolean) and `Failures` (an array of failed file paths).
#>
function Remove-LogFilesWithPrompt {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Files
    )
    $failedFiles = [System.Collections.Generic.List[string]]::new()
    $operationCancelled = $false

    foreach ($file in $Files) {
        try {
            Remove-Item -Path $file.FullName -Force -ErrorAction Stop
        } catch {
            $msg = "Could not delete file:`n$($file.FullName)`n`nReason: $($_.Exception.Message)`n`nClick 'Yes' to skip this file and continue, or 'No' to cancel the entire operation."
            $result = [System.Windows.Forms.MessageBox]::Show($msg, "File Deletion Error", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($result -eq 'No') { $operationCancelled = $true; break } 
            else { $failedFiles.Add($file.FullName) }
        }
    }
    return [PSCustomObject]@{ Cancelled = $operationCancelled; Failures = $failedFiles }
}

$script:BtnClearLogs.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete all log files? This action is irreversible.", "Confirm Log Deletion", "YesNo", "Warning")
    if ($confirmation -ne 'Yes') {
        $script:LblStatus.Text = "Clear logs cancelled."; $script:LblStatus.ForeColor = 'Gray'; return
    }
    
    try {
        Set-CleanupButtonsState -Enabled $false
        $script:LblStatus.Text = "Clearing all log files..."
        $script:LblStatus.ForeColor = 'Black'
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        $userPaths = Get-SelectedUserPaths
        Write-Log "Gathering all user log files for user $($script:selectedUser)."
        $filesToDelete = Get-AllUserLogFiles -UserPaths $userPaths

        if ($filesToDelete.Count -eq 0) {
            $script:LblStatus.Text = "No log files found to delete."
            $script:LblStatus.ForeColor = 'Green'
            return
        }

        $result = Remove-LogFilesWithPrompt -Files $filesToDelete

        if ($result.Cancelled) {
            $script:LblStatus.Text = "Clear logs operation cancelled by user."
            $script:LblStatus.ForeColor = 'Orange'
        } elseif ($result.Failures.Count -gt 0) {
            $script:LblStatus.Text = "Log clearing complete, but $($result.Failures.Count) file(s) were skipped."
            $script:LblStatus.ForeColor = 'Orange'
            Write-Log "Log clearing complete. Skipped files: $($result.Failures -join ', ')"
        } else {
            $script:LblStatus.Text = "All log files have been cleared."
            $script:LblStatus.ForeColor = 'Green'
            Write-Log "Successfully cleared $($filesToDelete.Count) log files."
        }
    } catch {
        $script:LblStatus.Text = "Failed to clear logs: $($_.Exception.Message)"
        $script:LblStatus.ForeColor = 'Red'
        Write-Log "ERROR during log clearing: $($_.Exception.Message)"
    } finally {
        Set-CleanupButtonsState -Enabled $true
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$script:BtnClearDMSLogs.Add_Click({
    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the DMS logs? This action is irreversible.", "Confirm DMS Log Deletion", "YesNo", "Warning")
    if ($confirmation -ne 'Yes') {
        $script:LblStatus.Text = "Clear DMS logs cancelled."; $script:LblStatus.ForeColor = 'Gray'; return
    }
    
    try {
        Set-CleanupButtonsState -Enabled $false
        $script:LblStatus.Text = "Clearing DMS specific logs..."
        $script:LblStatus.ForeColor = 'Black'
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        $userPaths = Get-SelectedUserPaths
        Write-Log "Clearing DMS-specific logs."
        $dmsLogPath = Join-Path $userPaths.literaAppData "dms\log"
        if (Test-Path $dmsLogPath -PathType Container) { Get-ChildItem -Path $dmsLogPath -File | Remove-Item -Force -ErrorAction Stop }
        
        $dmsDebugLog = Join-Path $userPaths.literaAppData "dms_debug_log.txt"
        if (Test-Path $dmsDebugLog -PathType Leaf) { Remove-Item -Path $dmsDebugLog -Force -ErrorAction Stop }

        $script:LblStatus.Text = "DMS specific logs have been cleared."
        $script:LblStatus.ForeColor = 'Green'
    } catch {
        $script:LblStatus.Text = "Failed to clear DMS logs: $($_.Exception.Message)"
        $script:LblStatus.ForeColor = 'Red'
        Write-Log "ERROR during DMS log clearing: $($_.Exception.Message)"
    } finally {
        Set-CleanupButtonsState -Enabled $true
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$script:BtnDeleteDMSConfig.Add_Click({
    $userPaths = Get-SelectedUserPaths
    $dmsConfigFile = Join-Path $userPaths.appData "Litera.iManage.Work.V2.Config"
    if (-not (Test-Path $dmsConfigFile)) {
        $script:LblStatus.Text = "No user DMS config file found to delete."; $script:LblStatus.ForeColor = 'Orange'; return
    }
    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete the User DMS Config file at `n$dmsConfigFile`?", "Confirm Deletion", "YesNo", "Question")
    if ($confirmation -ne 'Yes') {
        $script:LblStatus.Text = "Delete DMS config cancelled."; $script:LblStatus.ForeColor = 'Gray'; return
    }
    
    try {
        Set-CleanupButtonsState -Enabled $false
        $script:LblStatus.Text = "Deleting user DMS config file..."
        $script:LblStatus.ForeColor = 'Black'
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        Write-Log "Deleting user DMS config file: $dmsConfigFile"
        Remove-Item -Path $dmsConfigFile -Force -ErrorAction Stop

        $script:LblStatus.Text = "User DMS config file deleted successfully."
        $script:LblStatus.ForeColor = 'Green'
    } catch {
        $script:LblStatus.Text = "Failed to delete DMS config: $($_.Exception.Message)"
        $script:LblStatus.ForeColor = 'Red'
        Write-Log "ERROR during DMS config deletion: $($_.Exception.Message)"
    } finally {
        Set-CleanupButtonsState -Enabled $true
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$script:ChkSelectAllLocal.Add_Click({ for ($i = 0; $i -lt $script:ClbLocalAppData.Items.Count; $i++) { $script:ClbLocalAppData.SetItemChecked($i, $script:ChkSelectAllLocal.Checked) } })
$script:ChkSelectAllRoaming.Add_Click({ for ($i = 0; $i -lt $script:ClbAppData.Items.Count; $i++) { $script:ClbAppData.SetItemChecked($i, $script:ChkSelectAllRoaming.Checked) } })
$script:BtnRefreshCleanup.Add_Click({ Populate-CleanupLists })
$script:BtnLoadUserFolders.Add_Click({
    $selectedItem = $script:ComboUsers.SelectedItem
    if ([string]::IsNullOrWhiteSpace($selectedItem)) { return }

    $script:selectedUser = ($selectedItem -split ' ')[0]
    Populate-CleanupLists
})

<#
.SYNOPSIS
    Creates and configures all UI elements for the "Backup & Restore" tab.
.DESCRIPTION
    This function encapsulates the creation of all controls, panels, and layouts for the backup and restore tab.
    It uses the script scope for control variables so they can be accessed by event handlers later in the script.
.RETURNS
    A `System.Windows.Forms.TableLayoutPanel` containing the complete UI for the tab.
#>
function Create-BackupRestoreTabUI {
    $script:BackupRestoreLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:BackupRestoreLayout.Dock = 'Fill'
    $script:BackupRestoreLayout.BackColor = [System.Drawing.SystemColors]::Window
    $script:BackupRestoreLayout.ColumnCount = 1
    $script:BackupRestoreLayout.RowCount = 2
    [void]$script:BackupRestoreLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$script:BackupRestoreLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))

    $script:GbBackup = New-Object System.Windows.Forms.GroupBox; $script:GbBackup.Text = "Backup from C:\ProgramData\Litera"; $script:GbBackup.Dock = 'Fill'; $script:GbBackup.Padding = (New-Object System.Windows.Forms.Padding(10))
    [void]$script:BackupRestoreLayout.Controls.Add($script:GbBackup, 0, 0)

    $script:backupGroupLayout = New-Object System.Windows.Forms.TableLayoutPanel; $script:backupGroupLayout.Dock = 'Fill'; $script:backupGroupLayout.ColumnCount = 1; $script:backupGroupLayout.RowCount = 3
    [void]$script:backupGroupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$script:backupGroupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:backupGroupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:GbBackup.Controls.Add($script:backupGroupLayout)

    $script:ClbBackup = New-Object System.Windows.Forms.CheckedListBox; $script:ClbBackup.Dock = 'Fill'
    $script:backupButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $script:backupButtonPanel.Dock = 'Fill'; $script:backupButtonPanel.FlowDirection = 'LeftToRight'
    $script:ChkSelectAllBackup = New-Object System.Windows.Forms.CheckBox; $script:ChkSelectAllBackup.Text = "Select All"
    $script:BtnBackup             = New-AppButton -Text "Backup Selected" -Width 120 -BackColor 'LightGreen'
    $script:BtnOpenBackupFolder   = New-AppButton -Text "Open Backup Folder" -Width 140
    [void]$script:backupButtonPanel.Controls.AddRange(@($script:ChkSelectAllBackup, $script:BtnBackup, $script:BtnOpenBackupFolder))
    $script:LblBackupStatus = New-StatusLabel -Text "Backup status: Idle" -InitialColor 'Gray' -Dock 'Fill'

    [void]$script:backupGroupLayout.Controls.Add($script:ClbBackup, 0, 0)
    [void]$script:backupGroupLayout.Controls.Add($script:backupButtonPanel, 0, 1)
    [void]$script:backupGroupLayout.Controls.Add($script:LblBackupStatus, 0, 2)

    $script:GbRestore = New-Object System.Windows.Forms.GroupBox; $script:GbRestore.Text = "Restore from Backup Location"; $script:GbRestore.Dock = 'Fill'; $script:GbRestore.Padding = (New-Object System.Windows.Forms.Padding(10))
    [void]$script:BackupRestoreLayout.Controls.Add($script:GbRestore, 0, 1)

    $script:restoreGroupLayout = New-Object System.Windows.Forms.TableLayoutPanel; $script:restoreGroupLayout.Dock = 'Fill'; $script:restoreGroupLayout.ColumnCount = 1; $script:restoreGroupLayout.RowCount = 3
    [void]$script:restoreGroupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$script:restoreGroupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:restoreGroupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:GbRestore.Controls.Add($script:restoreGroupLayout)

    $script:ClbRestore = New-Object System.Windows.Forms.CheckedListBox; $script:ClbRestore.Dock = 'Fill'
    $script:restoreButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $script:restoreButtonPanel.Dock = 'Fill'; $script:restoreButtonPanel.FlowDirection = 'LeftToRight'
    $script:ChkSelectAllRestore = New-Object System.Windows.Forms.CheckBox; $script:ChkSelectAllRestore.Text = "Select All"
    $script:BtnRestore        = New-AppButton -Text "Restore Selected" -Width 120 -BackColor 'LightBlue'
    $script:BtnDeleteBackup   = New-AppButton -Text "Delete Selected" -Width 120 -BackColor 'MistyRose'
    $script:BtnRefreshBackups = New-AppButton -Text "Refresh List" -Width 120
    [void]$script:restoreButtonPanel.Controls.AddRange(@($script:ChkSelectAllRestore, $script:BtnRestore, $script:BtnDeleteBackup, $script:BtnRefreshBackups))
    $script:LblRestoreStatus = New-StatusLabel -Text "Restore status: Idle" -InitialColor 'Gray' -Dock 'Fill'

    [void]$script:restoreGroupLayout.Controls.Add($script:ClbRestore, 0, 0)
    [void]$script:restoreGroupLayout.Controls.Add($script:restoreButtonPanel, 0, 1)
    [void]$script:restoreGroupLayout.Controls.Add($script:LblRestoreStatus, 0, 2)

    return $script:BackupRestoreLayout
}

<#
.SYNOPSIS
    Defines the UI and logic for the "Backup & Restore" tab.
.DESCRIPTION
    This section builds the UI for backing up and restoring application data from `C:\ProgramData\Litera`. This is useful for preserving settings before an uninstall or migration.
#>
#==============================================================================
# --- UI: Backup & Restore Tab ---
#==============================================================================
$backupRestoreTabPanel = Create-BackupRestoreTabUI
$script:TabBackupRestore.Controls.Add($backupRestoreTabPanel)

<#
.SYNOPSIS
    Enables or disables all buttons on the Backup & Restore tab.
.PARAMETER Enabled
    $true to enable the buttons, $false to disable them.
#>
function Set-BackupRestoreButtonsState {
    param([bool]$Enabled)
    $script:BtnBackup.Enabled = $Enabled
    $script:BtnOpenBackupFolder.Enabled = $Enabled
    $script:BtnRestore.Enabled = $Enabled
    $script:BtnDeleteBackup.Enabled = $Enabled
    $script:BtnRefreshBackups.Enabled = $Enabled
    $script:ChkSelectAllBackup.Enabled = $Enabled
    $script:ChkSelectAllRestore.Enabled = $Enabled
}

<#
.SYNOPSIS
    Populates the backup source list (from ProgramData) and the restore destination list (from the backup folder).
#>

function Populate-BackupRestoreLists {
    $script:ClbBackup.Items.Clear()
    $script:ClbRestore.Items.Clear()

    if (Test-Path $literaProgramData) {
        Get-ChildItem -Path $literaProgramData -Directory | ForEach-Object {
            [void]$script:ClbBackup.Items.Add($_.Name, $false)
        }
    }
    if (Test-Path $script:backupRoot) {
        Get-ChildItem -Path $script:backupRoot -Directory | ForEach-Object {
            [void]$script:ClbRestore.Items.Add($_.Name, $false)
        }
    }
}

<#
.SYNOPSIS
    Creates and configures all UI elements for the "System Requirements" tab.
.DESCRIPTION
    This function encapsulates the creation of all controls, panels, and layouts for the System Requirements tab.
    It uses the script scope for control variables so they can be accessed by event handlers later in the script.
.RETURNS
    A `System.Windows.Forms.TableLayoutPanel` containing the complete UI for the tab.
#>
function Create-SysReqTabUI {
    $script:SysReqLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:SysReqLayout.Dock = 'Fill'
    $script:SysReqLayout.BackColor = [System.Drawing.SystemColors]::Window
    $script:SysReqLayout.ColumnCount = 1
    $script:SysReqLayout.RowCount = 2
    [void]$script:SysReqLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$script:SysReqLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))

    $script:GbServerReq = New-Object System.Windows.Forms.GroupBox; $script:GbServerReq.Text = "Recommanded Requirements"; $script:GbServerReq.Dock = 'Fill'; $script:GbServerReq.Padding = (New-Object System.Windows.Forms.Padding(10))
    [void]$script:SysReqLayout.Controls.Add($script:GbServerReq, 0, 0)

    $script:serverReqLayout = New-Object System.Windows.Forms.TableLayoutPanel; $script:serverReqLayout.Dock = 'Fill'; $script:serverReqLayout.ColumnCount = 2; $script:serverReqLayout.RowCount = 2
    [void]$script:serverReqLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$script:serverReqLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$script:serverReqLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:serverReqLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$script:GbServerReq.Controls.Add($script:serverReqLayout)

    $script:lblSys = New-Object System.Windows.Forms.Label
    $script:lblSys.Text = "System Requirements"; $script:lblSys.Font = New-Object System.Drawing.Font("Segoe UI",10); $script:lblSys.AutoSize = $true
    [void]$script:serverReqLayout.Controls.Add($script:lblSys,0,0)

    $script:lblPre = New-Object System.Windows.Forms.Label
    $script:lblPre.Text = "Application Pre-Requisites"; $script:lblPre.Font = New-Object System.Drawing.Font("Segoe UI",10); $script:lblPre.AutoSize = $true
    [void]$script:serverReqLayout.Controls.Add($script:lblPre,1,0)

    $script:rtbSys = New-Object System.Windows.Forms.RichTextBox
    $script:rtbSys.ReadOnly = $true; $script:rtbSys.Dock='Fill'; $script:rtbSys.Font = New-Object System.Drawing.Font("Segoe UI Emoji",10); $script:rtbSys.DetectUrls = $true

    $script:rtbPre = New-Object System.Windows.Forms.RichTextBox
    $script:rtbPre.ReadOnly = $true; $script:rtbPre.Dock='Fill'; $script:rtbPre.Font = New-Object System.Drawing.Font("Segoe UI Emoji",10); $script:rtbPre.DetectUrls = $true

    [void]$script:serverReqLayout.Controls.Add($script:rtbSys,0,1)
    [void]$script:serverReqLayout.Controls.Add($script:rtbPre,1,1)

    $script:GbValidation = New-Object System.Windows.Forms.GroupBox; $script:GbValidation.Text = "Local System Validation & Actions"; $script:GbValidation.Dock = 'Fill'; $script:GbValidation.Padding = (New-Object System.Windows.Forms.Padding(10))
    [void]$script:SysReqLayout.Controls.Add($script:GbValidation, 0, 1)

    $script:validationLayout = New-Object System.Windows.Forms.TableLayoutPanel; $script:validationLayout.Dock = 'Fill'; $script:validationLayout.ColumnCount = 1; $script:validationLayout.RowCount = 2
    [void]$script:validationLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$script:validationLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:GbValidation.Controls.Add($script:validationLayout)

    $script:rtbResults = New-Object System.Windows.Forms.RichTextBox
    $script:rtbResults.ReadOnly = $true; $script:rtbResults.Dock='Fill'; $script:rtbResults.Font = New-Object System.Drawing.Font("Segoe UI Emoji",10); $script:rtbResults.DetectUrls = $true
    $script:rtbResults.Add_LinkClicked({ param($s,$e) Start-Process -FilePath $e.LinkText -WindowStyle Hidden })
    [void]$script:validationLayout.Controls.Add($script:rtbResults, 0, 0)

    $script:sysReqButtonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $script:sysReqButtonPanel.Dock = 'Fill'
    $script:sysReqButtonPanel.FlowDirection = 'LeftToRight'
    $script:sysReqButtonPanel.Padding = (New-Object System.Windows.Forms.Padding(0, 5, 0, 0))
    [void]$script:validationLayout.Controls.Add($script:sysReqButtonPanel, 0, 1)

    $script:btnFetch    = New-AppButton -Text "Fetch Requirements" -AutoSize
    $script:btnLocalSys = New-AppButton -Text "Check My System" -AutoSize
    $script:btnLocalPre = New-AppButton -Text "Check My Pre-Reqs" -AutoSize
    $script:btnExport   = New-AppButton -Text "Export Results" -AutoSize
    $script:btnClearSysReqOutput = New-AppButton -Text "Clear Output" -AutoSize -BackColor 'LightGray'
    [void]$script:sysReqButtonPanel.Controls.AddRange(@($script:btnFetch, $script:btnLocalSys, $script:btnLocalPre, $script:btnExport, $script:btnClearSysReqOutput))

    return $script:SysReqLayout
}

<#
.SYNOPSIS
    Defines the UI and logic for the "System Requirements / Pre-requisites" tab.
.DESCRIPTION
    This section builds the UI for fetching the latest system requirements from an online source and validating the local system against them.
#>
#==============================================================================
# --- UI: System Requirements / Pre-requisites Tab ---
#==============================================================================
$sysReqTabPanel = Create-SysReqTabUI
$script:TabSysReq.Controls.Add($sysReqTabPanel)

<#
.SYNOPSIS
    Asynchronously fetches JSON content from a URL.
.DESCRIPTION
    A helper function that starts a background job to fetch and parse JSON from a given URL with a 30-second timeout.
.RETURNS
    The deserialized PSCustomObject from the JSON, or $null on failure or timeout.
#>
function Get-TabJsonContent($url) {
    try {
        $job = Start-Job -ScriptBlock {
            param($uri)
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
            Invoke-RestMethod -Uri $uri -UseBasicParsing
        } -ArgumentList $url

        Wait-Job $job -Timeout 30 | Out-Null
        if ($job.State -eq 'Completed') {
            return Receive-Job $job
        } else {
            Write-Warning "JSON fetch job timed out or failed for $url"
            return $null
        }
    } catch {
        Write-Warning "Error fetching JSON from $url : $($_.Exception.Message)"
        return $null
    }
}

<#
.SYNOPSIS
    Gathers local system information.
.DESCRIPTION
    Collects key system details like OS, RAM, CPU, free disk space, and installed Office versions using CIM and registry queries.
.RETURNS
    A PSCustomObject containing the collected system information.
#>
function Get-TabLocalSystemInfo {
    $osInfo = Get-CimInstance Win32_OperatingSystem
    $osName = $osInfo.Caption
    $ramGB = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB,1)
    $cpu = (Get-CimInstance Win32_Processor).Name
    $diskGB = [math]::Round((Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace/1GB,0)

    $keys = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    $office = (Get-ItemProperty $keys -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName -match "Microsoft Office|Microsoft 365" } |
                Select-Object -ExpandProperty DisplayName -Unique)

    [pscustomobject]@{
        CPU   = $cpu
        RAM   = $ramGB
        Disk  = $diskGB
        OS    = $osName
        Office= $office
    }
}

$script:btnFetch.Add_Click({
    $script:rtbSys.Clear(); $script:rtbPre.Clear(); $script:rtbResults.Clear()
    $sysReq = Get-TabJsonContent $sysReqUrl
    $preReq = Get-TabJsonContent $preReqUrl

    if ($sysReq) {
        Write-HeaderLine $script:rtbSys "=== Recommended Hardware ==="
        Write-ColoredLine $script:rtbSys "🖥️ Processor: $($sysReq.SystemRequirements.RecommendedHardware.Processor)" ([System.Drawing.Color]::Black)
        Write-ColoredLine $script:rtbSys "💾 Memory: Min $($sysReq.SystemRequirements.RecommendedHardware.Memory.Minimum), Rec $($sysReq.SystemRequirements.RecommendedHardware.Memory.Recommended)" ([System.Drawing.Color]::Black)
        Write-ColoredLine $script:rtbSys "💿 Disk: $($sysReq.SystemRequirements.RecommendedHardware.DiskSpace)" ([System.Drawing.Color]::Black)
        Write-Separator $script:rtbSys

        Write-HeaderLine $script:rtbSys "=== OCR Hardware (for PDF Compare) ==="
        $ocr = $sysReq.SystemRequirements.OCRHardware
        Write-ColoredLine $script:rtbSys "🖥️ Processor: $($ocr.Processor)" ([System.Drawing.Color]::Black)
        Write-ColoredLine $script:rtbSys "💾 Memory: Min $($ocr.Memory.Minimum), Rec $($ocr.Memory.Recommended)" ([System.Drawing.Color]::Black)
        Write-ColoredLine $script:rtbSys "💿 Disk: $($ocr.DiskSpace)" ([System.Drawing.Color]::Black)
        if ($ocr.Notes) { Write-ColoredLine $script:rtbSys "📝 Note: $($ocr.Notes)" ([System.Drawing.Color]::Gray) }
        Write-Separator $script:rtbSys

        Write-HeaderLine $script:rtbSys "=== Supported Environments ==="
        $env = $sysReq.SupportedEnvironments
        if ($env.Notes) { Write-ColoredLine $script:rtbSys "📝 Note: $($env.Notes)" ([System.Drawing.Color]::Gray) }

        Write-ColoredLine $script:rtbSys "🪟 Operating Systems:" ([System.Drawing.Color]::DarkBlue) -Bold
        foreach ($os in $env.OperatingSystems) { Write-ColoredLine $script:rtbSys "  • $os" ([System.Drawing.Color]::Black) }

        Write-ColoredLine $script:rtbSys "🏢 Citrix:" ([System.Drawing.Color]::DarkBlue) -Bold
        foreach ($version in $env.Citrix.Versions) { Write-ColoredLine $script:rtbSys "  • $version" ([System.Drawing.Color]::Black) }
        if ($env.Citrix.Notes) { Write-ColoredLine $script:rtbSys "  📝 Note: $($env.Citrix.Notes)" ([System.Drawing.Color]::Gray) }

        Write-ColoredLine $script:rtbSys "📦 Microsoft Office:" ([System.Drawing.Color]::DarkBlue) -Bold
        foreach ($version in $env.MicrosoftOffice.Versions) { Write-ColoredLine $script:rtbSys "  • $version" ([System.Drawing.Color]::Black) }
        if ($env.MicrosoftOffice.Notes) {
            foreach ($note in $env.MicrosoftOffice.Notes) { Write-ColoredLine $script:rtbSys "  📝 Note: $note" ([System.Drawing.Color]::Gray) }
        }
        Write-Separator $script:rtbSys

        Write-HeaderLine $script:rtbSys "=== DMS Integration ==="
        $dms = $sysReq.DMSIntegration
        if ($dms.Notes) { Write-ColoredLine $script:rtbSys "📝 Note: $($dms.Notes)" ([System.Drawing.Color]::Gray) }

        if ($dms.NetDocuments) { Write-ColoredLine $script:rtbSys "📁 NetDocuments" ([System.Drawing.Color]::DarkBlue) -Bold; foreach($item in $dms.NetDocuments) { Write-ColoredLine $script:rtbSys "  • $item" ([System.Drawing.Color]::Black) } }
        if ($dms.iManage) {
            Write-ColoredLine $script:rtbSys "📁 iManage" ([System.Drawing.Color]::DarkBlue) -Bold
            if ($dms.iManage.Clients) {
                Write-ColoredLine $script:rtbSys "  Clients:" ([System.Drawing.Color]::Black)
                foreach($client in $dms.iManage.Clients) { Write-ColoredLine $script:rtbSys "    • $client" ([System.Drawing.Color]::Black) }
            }
            if ($dms.iManage.Notes) {
                foreach($note in $dms.iManage.Notes) { Write-ColoredLine $script:rtbSys "    📝 Note: $note" ([System.Drawing.Color]::Gray) }
            }
        }
        if ($dms.SharePoint) { Write-ColoredLine $script:rtbSys "📁 SharePoint" ([System.Drawing.Color]::DarkBlue) -Bold; foreach($item in $dms.SharePoint) { Write-ColoredLine $script:rtbSys "  • $item" ([System.Drawing.Color]::Black) } }
        if ($dms.OpenText) { Write-ColoredLine $script:rtbSys "📁 OpenText" ([System.Drawing.Color]::DarkBlue) -Bold; foreach($item in $dms.OpenText) { Write-ColoredLine $script:rtbSys "  • $item" ([System.Drawing.Color]::Black) } }
        if ($dms.Worldox) { Write-ColoredLine $script:rtbSys "📁 Worldox" ([System.Drawing.Color]::DarkBlue) -Bold; foreach($item in $dms.Worldox) { Write-ColoredLine $script:rtbSys "  • $item" ([System.Drawing.Color]::Black) } }
        if ($dms.Epona) { Write-ColoredLine $script:rtbSys "📁 Epona" ([System.Drawing.Color]::DarkBlue) -Bold; foreach($item in $dms.Epona) { Write-ColoredLine $script:rtbSys "  • $item" ([System.Drawing.Color]::Black) } }
    } else {
        Write-ColoredLine $script:rtbSys "❌ Could not load system requirements from server." ([System.Drawing.Color]::Red)
    }

    if ($preReq) {
        Write-HeaderLine $script:rtbPre "=== Pre-Requisites ==="
        foreach ($item in $preReq.Prerequisites) {
            Write-ColoredLine $script:rtbPre "📦 $($item.Name)" ([System.Drawing.Color]::Black) -Bold
            if ($item.Description) { Write-ColoredLine $script:rtbPre "$($item.Description)" ([System.Drawing.Color]::Gray) }
            if ($item.Link) { Write-ColoredLine $script:rtbPre "$($item.Link)" ([System.Drawing.Color]::Blue) -Underline }
            if ($item.Requirements) {
                foreach ($r in $item.Requirements) {
                    Write-ColoredLine $script:rtbPre "$($r.Name)" ([System.Drawing.Color]::Black)
                    if ($r.Description) { Write-ColoredLine $script:rtbPre "$($r.Description)" ([System.Drawing.Color]::Gray) }
                    if ($r.Link) { Write-ColoredLine $script:rtbPre "   $($r.Link)" ([System.Drawing.Color]::Blue) -Underline }
                }
            }
            Write-Separator $script:rtbPre
        }
    } else {
        Write-ColoredLine $script:rtbPre "❌ Could not load pre-requisites from server." ([System.Drawing.Color]::Red)
    }
})

$script:btnLocalSys.Add_Click({
    $script:rtbResults.Clear()
    Write-ColoredLine $script:rtbResults "Checking local system information..." ([System.Drawing.Color]::Gray)
    $script:rtbResults.Update()

    $local = Get-TabLocalSystemInfo

    $script:rtbResults.Clear()
    Write-HeaderLine $script:rtbResults "=== Local System Info ==="
    Write-ColoredLine $script:rtbResults "🖥️ CPU: $($local.CPU)" ([System.Drawing.Color]::Green)
    Write-ColoredLine $script:rtbResults "💾 RAM: $($local.RAM) GB" ([System.Drawing.Color]::Green)
    Write-ColoredLine $script:rtbResults "💿 Disk: $($local.Disk) GB free" ([System.Drawing.Color]::Green)
    Write-ColoredLine $script:rtbResults "🪟 OS: $($local.OS)" ([System.Drawing.Color]::Green)
    if ($local.Office) { Write-ColoredLine $script:rtbResults "📦 Office: $($local.Office)" ([System.Drawing.Color]::Green) }
    Write-Separator $script:rtbResults
})

$script:btnLocalPre.Add_Click({
    $script:rtbResults.Clear()
    Write-ColoredLine $script:rtbResults "Checking local pre-requisites..." ([System.Drawing.Color]::Gray)
    $script:rtbResults.Update()

    $preReq = Get-TabJsonContent $preReqUrl

    $script:rtbResults.Clear()
    Write-HeaderLine $script:rtbResults "=== Local Pre-Requisite Validation ==="
    if (-not $preReq) {
        Write-ColoredLine $script:rtbResults "❌ Could not load pre-requisites JSON from server to perform check." ([System.Drawing.Color]::Red)
    } else {
        foreach ($item in $preReq.Prerequisites) {
            if ($item.DetectionPattern) {
                if (Check-PreReqInstalled $item.DetectionPattern) {
                    Write-ColoredLine $script:rtbResults "✔ 📦 $($item.Name)" ([System.Drawing.Color]::Green)
                } else {
                    Write-ColoredLine $script:rtbResults "❌ 📦 $($item.Name)" ([System.Drawing.Color]::Red)
                    if ($item.Description) { Write-ColoredLine $script:rtbResults "$($item.Description)" ([System.Drawing.Color]::Gray) }
                    if ($item.Link) { Write-ColoredLine $script:rtbResults "$($item.Link)" ([System.Drawing.Color]::Blue) -Underline }
                }
            }
            if ($item.Requirements) {
                foreach ($r in $item.Requirements) {
                    if (Check-PreReqInstalled $r.DetectionPattern) {
                        Write-ColoredLine $script:rtbResults "✔ 📦 $($r.Name)" ([System.Drawing.Color]::Green)
                    } else {
                        Write-ColoredLine $script:rtbResults "❌ 📦 $($r.Name)" ([System.Drawing.Color]::Red)
                        if ($r.Description) { Write-ColoredLine $script:rtbResults "$($r.Description)" ([System.Drawing.Color]::Gray) }
                        if ($r.Link) { Write-ColoredLine $script:rtbResults "$($r.Link)" ([System.Drawing.Color]::Blue) -Underline }
                    }
                }
            }
        }
        Write-Separator $script:rtbResults
    }
})

$script:btnExport.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "HTML Files|*.html"
    $dlg.Title = "Export Report"
    if ($dlg.ShowDialog() -eq "OK") {
        $sysText = $script:rtbSys.Text
        $preText = $script:rtbPre.Text
        $resultsText = $script:rtbResults.Text
        Convert-ToHtmlReport -sysText $sysText -preText $preText -resultsText $resultsText -outputFile $dlg.FileName
        [System.Windows.Forms.MessageBox]::Show("Report exported to $($dlg.FileName)","Export Complete",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$script:btnClearSysReqOutput.Add_Click({
    $script:rtbSys.Clear()
    $script:rtbPre.Clear()
    $script:rtbResults.Clear()
})

$script:ChkSelectAllBackup.Add_Click({
    for ($i = 0; $i -lt $script:ClbBackup.Items.Count; $i++) {
        $script:ClbBackup.SetItemChecked($i, $script:ChkSelectAllBackup.Checked)
    }
})

$script:ChkSelectAllRestore.Add_Click({
    for ($i = 0; $i -lt $script:ClbRestore.Items.Count; $i++) {
        $script:ClbRestore.SetItemChecked($i, $script:ChkSelectAllRestore.Checked)
    }
})

$script:BtnBackup.Add_Click({
    $foldersToBackup = $script:ClbBackup.CheckedItems
    if ($foldersToBackup.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select folders to backup.", "No Selection", "OK", "Warning")
        return
    }

    try {
        $script:LblBackupStatus.Text = "Backing up selected folders..."
        $script:LblBackupStatus.ForeColor = [System.Drawing.Color]::Black
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        foreach ($folderName in $foldersToBackup) {
            $sourcePath = Join-Path $literaProgramData $folderName
            if (-not (Test-Path $sourcePath -PathType Container)) {
                Write-Log "SKIPPING backup for non-existent source folder: '$folderName'"
                continue
            }

            $timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
            $backupFolderName = "${folderName}_${timestamp}"
            $backupDestination = Join-Path $script:backupRoot $backupFolderName
            
            Write-Log "Backing up '$folderName' to '$backupDestination'"
            Copy-Item -Path $sourcePath -Destination $backupDestination -Recurse -Force -ErrorAction Stop
        }

        $script:LblBackupStatus.Text = "Backup completed successfully."
        $script:LblBackupStatus.ForeColor = [System.Drawing.Color]::Green
        [System.Windows.Forms.MessageBox]::Show("Selected folders backed up successfully.", "Backup Complete", "OK", "Information")
        Populate-BackupRestoreLists
    }
    catch {
        $script:LblBackupStatus.Text = "Backup failed: $($_.Exception.Message)"
        $script:LblBackupStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Log "ERROR during backup: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Backup failed: $($_.Exception.Message)", "Error", "OK", "Error")
    }
    finally {
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$script:BtnRestore.Add_Click({
    $backupSetsToRestore = $script:ClbRestore.CheckedItems
    if ($backupSetsToRestore.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a backup set to restore.", "No Selection", "OK", "Warning"); return
    }
    
    $confirmation = [System.Windows.Forms.MessageBox]::Show("This will restore the $($backupSetsToRestore.Count) selected backup set(s) and overwrite any existing files in the destination. Are you sure you want to continue?", "Confirm Restore", "YesNo", "Question")
    if ($confirmation -ne 'Yes') {
        $script:LblRestoreStatus.Text = "Restore cancelled."; $script:LblRestoreStatus.ForeColor = 'Gray'; return
    }

    try {
        $script:LblRestoreStatus.Text = "Restoring $($backupSetsToRestore.Count) backup set(s)..."
        $script:LblRestoreStatus.ForeColor = [System.Drawing.Color]::Black
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        $failures = [System.Collections.Generic.List[string]]::new()

        foreach ($selectedBackupSet in $backupSetsToRestore) {
            try {
                $backupSetPath = Join-Path $script:backupRoot $selectedBackupSet

                $match = [regex]::Match($selectedBackupSet, "(.+)_(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})")
                if (-not $match.Success) {
                    $failures.Add("'$selectedBackupSet' (Invalid Format)")
                    Write-Log "ERROR during restore: The selected backup '$selectedBackupSet' does not follow the expected 'FolderName_Timestamp' format. Cannot restore."
                    continue
                }
                
                $originalFolderName = $match.Groups[1].Value
                $restoreDestinationPath = Join-Path $literaProgramData $originalFolderName

                Write-Log "Restoring '$selectedBackupSet' to '$restoreDestinationPath'"
                # This copies the *contents* of the backup folder to the destination, overwriting existing files.
                Copy-Item -Path "$backupSetPath\*" -Destination $restoreDestinationPath -Recurse -Force -ErrorAction Stop
            } catch {
                $failures.Add("'$selectedBackupSet' ($($_.Exception.Message))")
                Write-Log "ERROR during restore of '$selectedBackupSet': $($_.Exception.Message)"
            }
        }

        if ($failures.Count -gt 0) {
            $script:LblRestoreStatus.Text = "Restore completed with $($failures.Count) error(s)."
            $script:LblRestoreStatus.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.MessageBox]::Show("Restore completed, but the following backups failed:`n`n$($failures -join "`n")", "Restore Complete with Errors", "OK", "Warning")
        } else {
            $script:LblRestoreStatus.Text = "Restore of $($backupSetsToRestore.Count) backup set(s) completed successfully."
            $script:LblRestoreStatus.ForeColor = [System.Drawing.Color]::Green
            [System.Windows.Forms.MessageBox]::Show("Restore completed successfully.", "Restore Complete", "OK", "Information")
        }
    }
    catch {
        # This is a fallback for unexpected critical errors outside the loop.
        $script:LblRestoreStatus.Text = "Restore failed: $($_.Exception.Message)"
        $script:LblRestoreStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Log "CRITICAL ERROR during restore operation: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Restore failed: $($_.Exception.Message)", "Error", "OK", "Error")
    }
    finally {
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$script:BtnDeleteBackup.Add_Click({
    $backupsToDelete = $script:ClbRestore.CheckedItems
    if ($backupsToDelete.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select backup sets to delete.", "No Selection", "OK", "Warning")
        return
    }
    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to permanently delete the selected backup sets? This cannot be undone.", "Confirm Deletion", "YesNo", "Warning")
    if ($confirmation -ne 'Yes') {
        $script:LblRestoreStatus.Text = "Delete operation cancelled."
        $script:LblRestoreStatus.ForeColor = 'Gray'
        return
    }

    try {
        $script:LblRestoreStatus.Text = "Deleting selected backups..."
        $script:LblRestoreStatus.ForeColor = [System.Drawing.Color]::Black
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        foreach ($backupName in $backupsToDelete) {
            $backupPath = Join-Path $script:backupRoot $backupName
            Write-Log "Deleting backup set: $backupPath"
            if (Test-Path $backupPath) {
                Remove-Item -Path $backupPath -Recurse -Force -ErrorAction Stop
            }
        }

        $script:LblRestoreStatus.Text = "Selected backup sets deleted successfully."
        $script:LblRestoreStatus.ForeColor = [System.Drawing.Color]::Green
        Populate-BackupRestoreLists
    }
    catch {
        $script:LblRestoreStatus.Text = "Error deleting backups: $($_.Exception.Message)"
        $script:LblRestoreStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Log "ERROR during backup deletion: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to delete backups: $($_.Exception.Message)", "Error", "OK", "Error")
    }
    finally {
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$script:BtnOpenBackupFolder.Add_Click({
    if (-not (Test-Path $script:backupRoot)) {
        New-Item -Path $script:backupRoot -ItemType Directory -Force | Out-Null
    }
    Invoke-Item $script:backupRoot
})

$script:BtnRefreshBackups.Add_Click({ Populate-BackupRestoreLists })

<#
.SYNOPSIS
    Creates and configures all UI elements for the "Install/Uninstall" tab.
.DESCRIPTION
    This function encapsulates the creation of all controls, panels, and layouts for the Install/Uninstall tab.
    It uses the script scope for control variables so they can be accessed by event handlers later in the script.
.RETURNS
    A `System.Windows.Forms.TableLayoutPanel` containing the complete UI for the tab.
#>
function Create-InstallUninstallTabUI {
    $rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $rootLayout.Dock = 'Fill'
    $rootLayout.ColumnCount = 1
    $rootLayout.RowCount = 2
    [void]$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $rootLayout.BackColor = [System.Drawing.SystemColors]::Window

    $script:InstallUninstallLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:InstallUninstallLayout.Dock = 'Fill'
    $script:InstallUninstallLayout.ColumnCount = 1
    $script:InstallUninstallLayout.RowCount = 2
    [void]$script:InstallUninstallLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:InstallUninstallLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$rootLayout.Controls.Add($script:InstallUninstallLayout, 0, 0)

    $script:GbUninstall = New-Object System.Windows.Forms.GroupBox; $script:GbUninstall.Text = "Uninstall Program (from Programs & Features)"; $script:GbUninstall.Dock = 'Top'; $script:GbUninstall.AutoSize = $true; $script:GbUninstall.Padding = (New-Object System.Windows.Forms.Padding(10))
    [void]$script:InstallUninstallLayout.Controls.Add($script:GbUninstall, 0, 0)

    $script:uninstallLayout = New-Object System.Windows.Forms.TableLayoutPanel; $script:uninstallLayout.Dock = 'Fill'; $script:uninstallLayout.AutoSize = $true; $script:uninstallLayout.ColumnCount = 3; $script:uninstallLayout.RowCount = 3
    [void]$script:uninstallLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:uninstallLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$script:uninstallLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:uninstallLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:uninstallLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:uninstallLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:GbUninstall.Controls.Add($script:uninstallLayout)

    $script:LblPrograms = New-Object System.Windows.Forms.Label; $script:LblPrograms.Text = "Select program:"; $script:LblPrograms.Dock = 'Fill'; $script:LblPrograms.TextAlign = 'MiddleLeft'
    $script:ComboPrograms = New-Object System.Windows.Forms.ComboBox; $script:ComboPrograms.Dock = 'Fill'; $script:ComboPrograms.DropDownStyle = "DropDownList"
    $script:BtnLoadPrograms = New-AppButton -Text "Load Litera Programs" -AutoSize
    $script:LblUninstallArgs = New-Object System.Windows.Forms.Label; $script:LblUninstallArgs.Text = "Command-line arguments:"; $script:LblUninstallArgs.Dock = 'Fill'; $script:LblUninstallArgs.TextAlign = 'MiddleLeft'
    $script:TxtUninstallArgs = New-Object System.Windows.Forms.TextBox; $script:TxtUninstallArgs.Dock = 'Fill'
    $script:BtnUninstall = New-AppButton -Text "Uninstall" -Width 100
    $script:uninstallProgressPanel = New-Object System.Windows.Forms.TableLayoutPanel; $script:uninstallProgressPanel.Dock = 'Fill'; $script:uninstallProgressPanel.ColumnCount = 1; $script:uninstallProgressPanel.RowCount = 3 
    [void]$script:uninstallProgressPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50))) 
    [void]$script:uninstallProgressPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) 
    [void]$script:uninstallProgressPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50))) 
    $script:ProgressUninstall = New-Object System.Windows.Forms.ProgressBar; $script:ProgressUninstall.Dock = 'Fill'; $script:ProgressUninstall.Visible = $false; $script:ProgressUninstall.MarqueeAnimationSpeed = 30 
    [void]$script:uninstallProgressPanel.Controls.Add($script:ProgressUninstall, 0, 1) 
    $script:LblUninstallResult = New-StatusLabel -Text "" -Dock 'Fill' -AutoSize -Visible $false

    [void]$script:uninstallLayout.Controls.Add($script:LblPrograms, 0, 0); [void]$script:uninstallLayout.Controls.Add($script:ComboPrograms, 1, 0); [void]$script:uninstallLayout.Controls.Add($script:BtnLoadPrograms, 2, 0)
    [void]$script:uninstallLayout.Controls.Add($script:LblUninstallArgs, 0, 1)
    [void]$script:uninstallLayout.Controls.Add($script:TxtUninstallArgs, 1, 1); $script:uninstallLayout.SetColumnSpan($script:TxtUninstallArgs, 2)
    [void]$script:uninstallLayout.Controls.Add($script:BtnUninstall, 0, 2); [void]$script:uninstallLayout.Controls.Add($script:uninstallProgressPanel, 1, 2); [void]$script:uninstallLayout.Controls.Add($script:LblUninstallResult, 2, 2)

    $script:GbInstall = New-Object System.Windows.Forms.GroupBox; $script:GbInstall.Text = "Install Program"; $script:GbInstall.Dock = 'Fill'; $script:GbInstall.Padding = (New-Object System.Windows.Forms.Padding(10))
    [void]$script:InstallUninstallLayout.Controls.Add($script:GbInstall, 0, 1)

    $script:installMainLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:installMainLayout.Dock = 'Fill'
    $script:installMainLayout.ColumnCount = 1
    $script:installMainLayout.RowCount = 4
    [void]$script:installMainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:installMainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:installMainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:installMainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$script:GbInstall.Controls.Add($script:installMainLayout)

    $script:installerPathLabelText = if ($script:settings.AllowExeInstallation) {
        "Select installer (.msi or .exe):"
    } else {
        "Select installer (.msi):"
    }

    $script:installerPathPanel = New-Object System.Windows.Forms.TableLayoutPanel; $script:installerPathPanel.Dock = 'Fill'; $script:installerPathPanel.AutoSize = $true; $script:installerPathPanel.ColumnCount = 3
    [void]$script:installerPathPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:installerPathPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$script:installerPathPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $script:LblInstaller = New-Object System.Windows.Forms.Label; $script:LblInstaller.Text = "Select installer (.msi):"; $script:LblInstaller.Dock = 'Fill'; $script:LblInstaller.TextAlign = 'MiddleLeft'
    $script:TxtInstallerPath = New-Object System.Windows.Forms.TextBox; $script:TxtInstallerPath.Dock = 'Fill'
    $script:BtnBrowseInstaller = New-AppButton -Text "Browse..." -AutoSize
    [void]$script:installerPathPanel.Controls.Add($script:LblInstaller, 0, 0); [void]$script:installerPathPanel.Controls.Add($script:TxtInstallerPath, 1, 0); [void]$script:installerPathPanel.Controls.Add($script:BtnBrowseInstaller, 2, 0)
    [void]$script:installMainLayout.Controls.Add($script:installerPathPanel, 0, 0)

    $script:installArgsPanel = New-Object System.Windows.Forms.TableLayoutPanel; $script:installArgsPanel.Dock = 'Fill'; $script:installArgsPanel.AutoSize = $true; $script:installArgsPanel.ColumnCount = 2
    [void]$script:installArgsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:installArgsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $script:LblInstallArgs = New-Object System.Windows.Forms.Label; $script:LblInstallArgs.Text = "Command-line arguments:"; $script:LblInstallArgs.Dock = 'Fill'; $script:LblInstallArgs.TextAlign = 'MiddleLeft'
    $script:TxtInstallArgs = New-Object System.Windows.Forms.TextBox; $script:TxtInstallArgs.Dock = 'Fill'
    [void]$script:installArgsPanel.Controls.Add($script:LblInstallArgs, 0, 0); [void]$script:installArgsPanel.Controls.Add($script:TxtInstallArgs, 1, 0)
    [void]$script:installMainLayout.Controls.Add($script:installArgsPanel, 0, 1)
 
    $script:installActionPanel = New-Object System.Windows.Forms.TableLayoutPanel; $script:installActionPanel.Dock = 'Top'; $script:installActionPanel.AutoSize = $true; $script:installActionPanel.ColumnCount = 3
    [void]$script:installActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:installActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$script:installActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $script:BtnInstall = New-AppButton -Text "Install" -Width 100
    $script:installProgressPanel = New-Object System.Windows.Forms.TableLayoutPanel; $script:installProgressPanel.Dock = 'Fill'; $script:installProgressPanel.ColumnCount = 1; $script:installProgressPanel.RowCount = 3 
    [void]$script:installProgressPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50))) 
    [void]$script:installProgressPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) 
    [void]$script:installProgressPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50))) 
    $script:ProgressInstall = New-Object System.Windows.Forms.ProgressBar; $script:ProgressInstall.Dock = 'Fill'; $script:ProgressInstall.Visible = $false; $script:ProgressInstall.MarqueeAnimationSpeed = 30 
    [void]$script:installProgressPanel.Controls.Add($script:ProgressInstall, 0, 1) 
    $script:LblInstallResult = New-StatusLabel -Text "" -Dock 'Fill' -AutoSize -Visible $false
    [void]$script:installActionPanel.Controls.Add($script:BtnInstall, 0, 0); [void]$script:installActionPanel.Controls.Add($script:installProgressPanel, 1, 0); [void]$script:installActionPanel.Controls.Add($script:LblInstallResult, 2, 0)
    [void]$script:installMainLayout.Controls.Add($script:installActionPanel, 0, 2)

    $script:installOptionsScrollPanel = New-Object System.Windows.Forms.Panel
    $script:installOptionsScrollPanel.Dock = 'Fill'
    $script:installOptionsScrollPanel.AutoScroll = $true
    [void]$script:installMainLayout.Controls.Add($script:installOptionsScrollPanel, 0, 3)
 
    $script:installOptionsLayout = New-Object System.Windows.Forms.TableLayoutPanel; $script:installOptionsLayout.Dock = 'Top'; $script:installOptionsLayout.AutoSize = $true; $script:installOptionsLayout.ColumnCount = 4; $script:installOptionsLayout.RowCount = 1
    [void]$script:installOptionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$script:installOptionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$script:installOptionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$script:installOptionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$script:installOptionsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:installOptionsScrollPanel.controls.add($script:installOptionsLayout)

    $script:GbDisableAddins = New-Object System.Windows.Forms.GroupBox; $script:GbDisableAddins.Text = "Disable Add-ins"; $script:GbDisableAddins.Dock = 'Fill'
    $script:addinsPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $script:addinsPanel.Dock = 'Fill'; $script:addinsPanel.FlowDirection = 'TopDown'
    $script:ChkDisableWordAddin = New-Object System.Windows.Forms.CheckBox; $script:ChkDisableWordAddin.Text = "LC4W"; $script:ChkDisableWordAddin.AutoSize = $true
    $script:ChkDisableExcelAddin = New-Object System.Windows.Forms.CheckBox; $script:ChkDisableExcelAddin.Text = "LC4XL"; $script:ChkDisableExcelAddin.AutoSize = $true
    $script:ChkDisableOutlookAddin = New-Object System.Windows.Forms.CheckBox; $script:ChkDisableOutlookAddin.Text = "LC4O"; $script:ChkDisableOutlookAddin.AutoSize = $true
    $script:ChkDisablePptAddin = New-Object System.Windows.Forms.CheckBox; $script:ChkDisablePptAddin.Text = "LC4PPT"; $script:ChkDisablePptAddin.AutoSize = $true
    [void]$script:addinsPanel.Controls.AddRange(@($script:ChkDisableWordAddin, $script:ChkDisableExcelAddin, $script:ChkDisableOutlookAddin, $script:ChkDisablePptAddin))
    [void]$script:GbDisableAddins.Controls.Add($script:addinsPanel)
    $script:GbDisableDesktopShortcuts = New-Object System.Windows.Forms.GroupBox; $script:GbDisableDesktopShortcuts.Text = "Disable Desktop Shortcuts"; $script:GbDisableDesktopShortcuts.Dock = 'Fill'
    $script:desktopShortcutsPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $script:desktopShortcutsPanel.Dock = 'Fill'; $script:desktopShortcutsPanel.FlowDirection = 'TopDown'
    $script:ChkNoDesktopWord = New-Object System.Windows.Forms.CheckBox; $script:ChkNoDesktopWord.Text = "LC4W"; $script:ChkNoDesktopWord.AutoSize = $true
    $script:ChkNoDesktopPpt = New-Object System.Windows.Forms.CheckBox; $script:ChkNoDesktopPpt.Text = "LC4PPT"; $script:ChkNoDesktopPpt.AutoSize = $true
    $script:ChkNoDesktopExcel = New-Object System.Windows.Forms.CheckBox; $script:ChkNoDesktopExcel.Text = "LC4XL"; $script:ChkNoDesktopExcel.AutoSize = $true
    [void]$script:desktopShortcutsPanel.Controls.AddRange(@($script:ChkNoDesktopWord, $script:ChkNoDesktopPpt, $script:ChkNoDesktopExcel))
    [void]$script:GbDisableDesktopShortcuts.Controls.Add($script:desktopShortcutsPanel)
    $script:GbDisableStartMenuShortcuts = New-Object System.Windows.Forms.GroupBox; $script:GbDisableStartMenuShortcuts.Text = "Disable Start Menu Shortcuts"; $script:GbDisableStartMenuShortcuts.Dock = 'Fill'
    $script:startMenuShortcutsPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $script:startMenuShortcutsPanel.Dock = 'Fill'; $script:startMenuShortcutsPanel.FlowDirection = 'TopDown'
    $script:ChkNoStartMenuWord = New-Object System.Windows.Forms.CheckBox; $script:ChkNoStartMenuWord.Text = "LC4W"; $script:ChkNoStartMenuWord.AutoSize = $true
    $script:ChkNoStartMenuPpt = New-Object System.Windows.Forms.CheckBox; $script:ChkNoStartMenuPpt.Text = "LC4PPT"; $script:ChkNoStartMenuPpt.AutoSize = $true
    $script:ChkNoStartMenuExcel = New-Object System.Windows.Forms.CheckBox; $script:ChkNoStartMenuExcel.Text = "LC4XL"; $script:ChkNoStartMenuExcel.AutoSize = $true
    [void]$script:startMenuShortcutsPanel.Controls.AddRange(@($script:ChkNoStartMenuWord, $script:ChkNoStartMenuPpt, $script:ChkNoStartMenuExcel))
    [void]$script:GbDisableStartMenuShortcuts.Controls.Add($script:startMenuShortcutsPanel)
    $script:GbInstallUiLevel = New-Object System.Windows.Forms.GroupBox; $script:GbInstallUiLevel.Text = "Install/Uninstall UI"; $script:GbInstallUiLevel.Dock = 'Fill'
    $script:uiLevelPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $script:uiLevelPanel.Dock = 'Fill'; $script:uiLevelPanel.FlowDirection = 'TopDown'
    $script:ChkInstallSilent = New-Object System.Windows.Forms.CheckBox; $script:ChkInstallSilent.Text = "Silent"; $script:ChkInstallSilent.AutoSize = $true
    $script:ChkInstallBasicUi = New-Object System.Windows.Forms.CheckBox; $script:ChkInstallBasicUi.Text = "Basic"; $script:ChkInstallBasicUi.AutoSize = $true
    $script:ChkRebootSuppress = New-Object System.Windows.Forms.CheckBox; $script:ChkRebootSuppress.Text = "REBOOT=ReallySuppress"; $script:ChkRebootSuppress.AutoSize = $true; $script:ToolTip.SetToolTip($script:ChkRebootSuppress, "Prevents the installer from forcing reboots or restarting explorer.exe.")
    [void]$script:uiLevelPanel.Controls.AddRange(@($script:ChkInstallSilent, $script:ChkInstallBasicUi, $script:ChkRebootSuppress))
    [void]$script:GbInstallUiLevel.Controls.Add($script:uiLevelPanel)
    [void]$script:installOptionsLayout.Controls.Add($script:GbDisableAddins, 0, 0)
    [void]$script:installOptionsLayout.Controls.Add($script:GbDisableDesktopShortcuts, 1, 0)
    [void]$script:installOptionsLayout.Controls.Add($script:GbDisableStartMenuShortcuts, 2, 0)
    [void]$script:installOptionsLayout.Controls.Add($script:GbInstallUiLevel, 3, 0)

    $LoggingPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $LoggingPanel.Dock = 'Fill'
    $LoggingPanel.FlowDirection = 'LeftToRight'
    $LoggingPanel.Padding = (New-Object System.Windows.Forms.Padding(10, 5, 10, 5))
    $LoggingPanel.Height = 40

    $script:ChkLogging = New-Object System.Windows.Forms.CheckBox
    $script:ChkLogging.Text = "Enable Logging"
    $script:ChkLogging.Checked = $script:logEnabled
    $script:ChkLogging.AutoSize = $true
    $script:ChkLogging.Margin = '10,8,10,8'

    $script:BtnOpenLogFolder = New-AppButton -Text "Open Log Folder" -AutoSize

    $script:LblLoggingStatus = New-StatusLabel -Text "" -AutoSize -Margin (New-Object System.Windows.Forms.Padding(20,8,10,8))

    [void]$LoggingPanel.Controls.AddRange(@($script:ChkLogging, $script:BtnOpenLogFolder, $script:LblLoggingStatus))
    [void]$rootLayout.Controls.Add($LoggingPanel, 0, 1)

    return $rootLayout
}

<#
.SYNOPSIS
    Defines the UI and logic for the "Install/Uninstall" tab.
.DESCRIPTION
    This section builds the UI for installing and uninstalling Litera Compare products, supporting silent installation, command-line arguments, and logging.
#>
#==============================================================================
# --- UI: Install/Uninstall Tab ---
#==============================================================================
$installUninstallTabPanel = Create-InstallUninstallTabUI
$script:TabInstallUninstall.Controls.Add($installUninstallTabPanel)
<#
.SYNOPSIS
    Checks and displays the system's time and time synchronization status.
.DESCRIPTION
    Queries WMI and the registry to gather information about the system's current time, time zone, time server, and the status of the Windows Time service (W32Time). The output is displayed in the License Info RichTextBox.
#>
function Check-TimeSync {
    $output = "[TIME] System Time Sync Info`r`n"
    $output += "--------------------------`r`n"
    $output += "Current System Time     : $(Get-Date)`r`n"
    $output += "Time Zone              : $([System.TimeZoneInfo]::Local.DisplayName)`r`n"

    try {
        $serviceConfig = Get-CimInstance -ClassName Win32_Service -Filter "Name='W32Time'"
        $isAuto = $serviceConfig.StartMode -eq 'Auto'
        $output += "Set Time Automatically : $($isAuto)`r`n"

        $tzSettings = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"
        $output += "Time Server            : $($tzSettings.NtpServer)`r`n"
    } catch {
        $output += "Set Time Automatically : Unknown`r`n"
        $output += "Time Server            : Unknown`r`n"
    }

    $syncStatus = "Error retrieving sync status"
    try {
        $statusKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Status"
        $lastSuccessFileTime = (Get-ItemProperty -Path $statusKeyPath -Name "LastSuccess" -ErrorAction SilentlyContinue).LastSuccess
        
        if ($lastSuccessFileTime -and $lastSuccessFileTime -ne 0) {
            $utcDateTime = [System.DateTime]::FromFileTimeUtc($lastSuccessFileTime)
            $localDateTime = $utcDateTime.ToLocalTime()
            $syncStatus = $localDateTime.ToString("dd-MM-yyyy HH:mm:ss")
        } else {
            $syncStatus = "Not available or never synced"
        }
    } catch {
        $syncStatus = "Error retrieving sync status: $($_.Exception.Message)"
    }
    $output += "Last successful time synchronization: $syncStatus`r`n"

    try {
        $service = Get-Service -Name W32Time
        $output += "Time Service Status    : $($service.Status)`r`n"
    } catch {
        $output += "Time Service Status    : Unknown`r`n"
    }

    $script:RtbLicenseInfo.Text = $output
}

<#
.SYNOPSIS
    Creates and configures all UI elements for the "License Info" tab.
.DESCRIPTION
    This function encapsulates the creation of all controls, panels, and layouts for the License Info tab.
    It uses the script scope for control variables so they can be accessed by event handlers later in the script.
.RETURNS
    A `System.Windows.Forms.TableLayoutPanel` containing the complete UI for the tab.
#>
function Create-LicenseTabUI {
    $script:LicenseLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $script:LicenseLayout.Dock = 'Fill'
    $script:LicenseLayout.ColumnCount = 1
    $script:LicenseLayout.Padding = (New-Object System.Windows.Forms.Padding(10))
    $script:LicenseLayout.BackColor = [System.Drawing.SystemColors]::Window
    $script:LicenseLayout.RowCount = 5
    [void]$script:LicenseLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:LicenseLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:LicenseLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$script:LicenseLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:LicenseLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    
    $script:LblLicenseHeader = New-Object System.Windows.Forms.Label; $script:LblLicenseHeader.Text = "License Connectivity Information"; $script:LblLicenseHeader.Dock = 'Top'
    [void]$script:LicenseLayout.Controls.Add($script:LblLicenseHeader, 0, 0)

    $script:GbLicenseUserSelection = New-Object System.Windows.Forms.GroupBox
    $script:GbLicenseUserSelection.Text = "Administrator: Select User"
    $script:GbLicenseUserSelection.Dock = 'Top'
    $script:GbLicenseUserSelection.AutoSize = $true
    $script:GbLicenseUserSelection.Padding = (New-Object System.Windows.Forms.Padding(10, 5, 10, 10))

    $script:LicenseUserSelectionLayout = New-Object System.Windows.Forms.FlowLayoutPanel
    $script:LicenseUserSelectionLayout.Dock = 'Top'
    $script:LicenseUserSelectionLayout.AutoSize = $true
    $script:LicenseUserSelectionLayout.WrapContents = $false
    $script:LblLicenseSelectUser = New-Object System.Windows.Forms.Label; $script:LblLicenseSelectUser.Text = "Check license for user:"; $script:LblLicenseSelectUser.Margin = '0,5,5,0'; $script:LblLicenseSelectUser.AutoSize = $true
    $script:ComboLicenseUsers = New-Object System.Windows.Forms.ComboBox; $script:ComboLicenseUsers.DropDownStyle = 'DropDownList'; $script:ComboLicenseUsers.Width = 200
    [void]$script:LicenseUserSelectionLayout.Controls.AddRange(@($script:LblLicenseSelectUser, $script:ComboLicenseUsers))
    [void]$script:GbLicenseUserSelection.Controls.Add($script:LicenseUserSelectionLayout)
    [void]$script:LicenseLayout.Controls.Add($script:GbLicenseUserSelection, 0, 1)

    $script:RtbLicenseInfo = New-Object System.Windows.Forms.RichTextBox
    $script:RtbLicenseInfo.Dock = 'Fill'
    $script:RtbLicenseInfo.ReadOnly = $true
    $script:RtbLicenseInfo.Font = New-Object System.Drawing.Font('Consolas', 10)
    [void]$script:LicenseLayout.Controls.Add($script:RtbLicenseInfo, 0, 2)

    $script:LicenseButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $script:LicenseButtonsPanel.Dock = 'Fill'
    $script:LicenseButtonsPanel.FlowDirection = 'LeftToRight'
    $script:LicenseButtonsPanel.Padding = (New-Object System.Windows.Forms.Padding(0, 5, 0, 0))
    [void]$script:LicenseLayout.Controls.Add($script:LicenseButtonsPanel, 0, 3)

    $script:BtnComprehensiveCheck = New-AppButton -Text "Comprehensive Check" -AutoSize -BackColor 'LightSkyBlue'
    $script:ToolTip.SetToolTip($script:BtnComprehensiveCheck, "Runs a full suite of checks for licensing issues.")
    $script:BtnCheckLicense       = New-AppButton -Text "Check License" -AutoSize
    $script:ToolTip.SetToolTip($script:BtnCheckLicense, "Reads and displays the Litera Compare license information from the local machine.")
    $script:BtnCheckTimeSync      = New-AppButton -Text "Check Time Sync" -AutoSize
    $script:ToolTip.SetToolTip($script:BtnCheckTimeSync, "Checks the system's time synchronization settings and status.")
    $script:BtnCheckServer        = New-AppButton -Text "Check Server Connectivity" -AutoSize
    $script:ToolTip.SetToolTip($script:BtnCheckServer, "Performs a connectivity check to the Litera licensing servers.")
    $script:BtnCancelCheckServer  = New-AppButton -Text "Cancel" -AutoSize -Enabled $false -BackColor 'MistyRose'
    $script:ToolTip.SetToolTip($script:BtnCancelCheckServer, "Stops the current connectivity check.")
    $script:BtnClearLicenseOutput = New-AppButton -Text "Clear Output" -AutoSize -BackColor 'LightGray'
    $script:ToolTip.SetToolTip($script:BtnClearLicenseOutput, "Clears the text from the output window above.")
    [void]$script:LicenseButtonsPanel.Controls.AddRange(@($script:BtnComprehensiveCheck, $script:BtnCheckLicense, $script:BtnCheckTimeSync, $script:BtnCheckServer, $script:BtnCancelCheckServer, $script:BtnClearLicenseOutput))

    $script:LicenseStatusPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $script:LicenseStatusPanel.Dock = 'Fill'
    $script:LicenseStatusPanel.AutoSize = $true
    $script:LicenseStatusPanel.ColumnCount = 3
    [void]$script:LicenseStatusPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:LicenseStatusPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$script:LicenseStatusPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$script:LicenseLayout.Controls.Add($script:LicenseStatusPanel, 0, 4)

    $script:ChkEnableLicenseLogging = New-Object System.Windows.Forms.CheckBox; $script:ChkEnableLicenseLogging.Text = "Enable License Check Logs"; $script:ChkEnableLicenseLogging.AutoSize = $true; $script:ChkEnableLicenseLogging.Margin = (New-Object System.Windows.Forms.Padding(0, 6, 0, 0))
    $script:BtnOpenLicenseLogFolder = New-AppButton -Text "Open Log Folder" -AutoSize; $script:ToolTip.SetToolTip($script:BtnOpenLicenseLogFolder, "Opens the folder where connectivity check logs are stored.")
    $script:LblLicenseStatus = New-StatusLabel -Text "Ready." -InitialColor 'Gray' -Dock 'Fill' -TextAlign 'MiddleRight' -Margin (New-Object System.Windows.Forms.Padding(10,3,0,3))

    [void]$script:LicenseStatusPanel.Controls.Add($script:ChkEnableLicenseLogging, 0, 0)
    [void]$script:LicenseStatusPanel.Controls.Add($script:BtnOpenLicenseLogFolder, 1, 0)
    [void]$script:LicenseStatusPanel.Controls.Add($script:LblLicenseStatus, 2, 0)

    return $script:LicenseLayout
}

$licenseTabPanel = Create-LicenseTabUI
$script:TabLicense.Controls.Add($licenseTabPanel)

if (Test-IsAdmin) {
    Populate-UserComboBox -ComboBox $script:ComboLicenseUsers
} else {
    $script:GbLicenseUserSelection.Visible = $false
}

<#
.SYNOPSIS
    Checks and displays the system's time and time synchronization status.
.DESCRIPTION
    Queries WMI and the registry to gather information about the system's current time, time zone, time server, and the status of the Windows Time service (W32Time). The output is displayed in the License Info RichTextBox.
#>
function Check-TimeSync {
    $output = "[TIME] System Time Sync Info`r`n"
    $output += "--------------------------`r`n"
    $output += "Current System Time     : $(Get-Date)`r`n"
    $output += "Time Zone              : $([System.TimeZoneInfo]::Local.DisplayName)`r`n"

    try {
        $serviceConfig = Get-CimInstance -ClassName Win32_Service -Filter "Name='W32Time'"
        $isAuto = $serviceConfig.StartMode -eq 'Auto'
        $output += "Set Time Automatically : $($isAuto)`r`n"

        $tzSettings = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"
        $output += "Time Server            : $($tzSettings.NtpServer)`r`n"
    } catch {
        $output += "Set Time Automatically : Unknown`r`n"
        $output += "Time Server            : Unknown`r`n"
    }

    $syncStatus = "Error retrieving sync status"
    try {
        $statusKeyPath = "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Status"
        $lastSuccessFileTime = (Get-ItemProperty -Path $statusKeyPath -Name "LastSuccess" -ErrorAction SilentlyContinue).LastSuccess
        
        if ($lastSuccessFileTime -and $lastSuccessFileTime -ne 0) {
            $utcDateTime = [System.DateTime]::FromFileTimeUtc($lastSuccessFileTime)
            $localDateTime = $utcDateTime.ToLocalTime()
            $syncStatus = $localDateTime.ToString("dd-MM-yyyy HH:mm:ss")
        } else {
            $syncStatus = "Not available or never synced"
        }
    } catch {
        $syncStatus = "Error retrieving sync status: $($_.Exception.Message)"
    }
    $output += "Last successful time synchronization: $syncStatus`r`n"

    try {
        $service = Get-Service -Name W32Time
        $output += "Time Service Status    : $($service.Status)`r`n"
    } catch {
        $output += "Time Service Status    : Unknown`r`n"
    }

    $script:RtbLicenseInfo.Text = $output
}

<#
.SYNOPSIS
    Decodes a Base64Url-encoded string.
.DESCRIPTION
    Converts a Base64Url-encoded string (commonly used in JWTs) to a standard Base64 string and then decodes it to a UTF8 string.
.RETURNS
    The decoded string, or $null on failure.
#>
function Decode-Base64Url {
    param([string]$base64Url)
    $base64 = $base64Url.Replace('-', '+').Replace('_', '/')
    switch ($base64.Length % 4) {
        2 { $base64 += '==' }
        3 { $base64 += '=' }
    }
    try {
        $bytes = [System.Convert]::FromBase64String($base64)
        return [System.Text.Encoding]::UTF8.GetString($bytes)
    }
    catch {
        Write-Log "Failed to decode Base64Url string. Error: $($_.Exception.Message)"
        return $null
    }
}

<#
.SYNOPSIS
    Reads, decodes, and displays the Litera Compare license token.
.DESCRIPTION
    Finds the `token.json` file for the selected user, parses the JWT access token within it, and displays the decoded license information (key, product, expiry) in the License Info RichTextBox.
#>
function Show-LicenseInfo {
    $userName = ($script:ComboLicenseUsers.SelectedItem -split ' ')[0]
    if ([string]::IsNullOrWhiteSpace($userName)) {
        $userName = $env:USERNAME
    }

    $targetAppData = if ($userName -eq $env:USERNAME) {
        $env:APPDATA
    } else {
        "C:\Users\$userName\AppData\Roaming"
    }
    $tokenFile = Join-Path $targetAppData "Litera\Compare\token.json"
    if (-not (Test-Path $tokenFile)) {
        $script:RtbLicenseInfo.Text = "❌ License file not found for user '$userName' at $tokenFile"
        return
    }

    try {
        $rawToken = Get-Content -Path $tokenFile -Raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        $script:RtbLicenseInfo.Text = "❌ Failed to parse License file"
        return
    }

    if (-not $rawToken.PSObject.Properties.Name -contains "AccessToken") {
        $script:RtbLicenseInfo.Text = "❌ License file does not contain 'AccessToken' field"
        return
    }

    $jwt = $rawToken.AccessToken
    $parts = $jwt -split '\.'
    if ($parts.Count -lt 2) {
        $script:RtbLicenseInfo.Text = "❌ Invalid JWT format"
        return
    }

    try {
        $payloadJsonText = Decode-Base64Url $parts[1]
        $payloadObj      = $payloadJsonText | ConvertFrom-Json -ErrorAction Stop

        $output  = "[KEY] License Information For User: $userName`r`n"
        $output += "--------------------------`r`n"

        if ($payloadObj.sub) {
            $output += "Key     : $($payloadObj.sub)`r`n"
        }

        if ($payloadObj.product) {
            $products = @()

            if ($payloadObj.product -is [System.Collections.IEnumerable] -and -not ($payloadObj.product -is [string])) {
                $products = $payloadObj.product
            }
            else {
                $products = @($payloadObj.product)
            }

            foreach ($prod in $products) {
                if ($prod -match "compare\.desktop") {
                    if ($prod -match "val\s*(\d+)") {
                        $unixVal = [int64]$Matches[1]
                        $expiry  = [DateTimeOffset]::FromUnixTimeSeconds($unixVal).DateTime.ToString("yyyy-MM-dd HH:mm:ss")
                    } else {
                        $expiry = "N/A"
                    }

                    $output += "Product : compare.dektop`r`n"
                    $output += "Expiry  : $expiry`r`n"
                }
            }
        }

        $script:RtbLicenseInfo.Text = $output.Trim()
    }
    catch {
        $script:RtbLicenseInfo.Text = "❌ Failed to decode JWT.`r`nPayload: $payloadJsonText"
    }
}

$script:BtnComprehensiveCheck.Add_Click({ Start-ComprehensiveLicenseCheck })

$script:BtnCheckLicense.Add_Click({ Show-LicenseInfo })
$script:BtnCheckTimeSync.Add_Click({ Check-TimeSync })
$script:BtnCheckServer.Add_Click({
    if ($script:currentJob) { [System.Windows.Forms.MessageBox]::Show("Another operation is already in progress.", "Busy", "OK", "Warning"); return }

    $script:BtnComprehensiveCheck.Enabled = $false
    $script:BtnCheckServer.Enabled = $false
    $script:BtnCheckLicense.Enabled = $false
    $script:BtnCheckTimeSync.Enabled = $false
    $script:BtnCancelCheckServer.Enabled = $true

    $script:RtbLicenseInfo.Clear()
    $script:RtbLicenseInfo.SelectionColor = 'Black'
    $script:RtbLicenseInfo.AppendText("[~] Connecting to Litera Licensing Server...`r`n")
    
    try {
        $pingResult = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet
        if (-not $pingResult) {
            $script:RtbLicenseInfo.SelectionColor = 'Red'
            $script:RtbLicenseInfo.AppendText("[X] No internet connection detected.`r`n")
            $script:BtnComprehensiveCheck.Enabled = $true; $script:BtnCheckServer.Enabled = $true; $script:BtnCheckLicense.Enabled = $true; $script:BtnCheckTimeSync.Enabled = $true
            return
        }
    } catch {
        $script:RtbLicenseInfo.SelectionColor = 'Red'
        $script:RtbLicenseInfo.AppendText("[X] Error checking internet connectivity.`r`n")
        $script:BtnComprehensiveCheck.Enabled = $true; $script:BtnCheckServer.Enabled = $true; $script:BtnCheckLicense.Enabled = $true; $script:BtnCheckTimeSync.Enabled = $true
        return
    }

    $proxy = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings" -ErrorAction SilentlyContinue).ProxyServer
    if ($proxy) {
        $script:RtbLicenseInfo.SelectionColor = 'Orange'
        $script:RtbLicenseInfo.AppendText("[!] Proxy detected: $proxy`r`n")
    }

    $vpnAdapters = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match "VPN" -and $_.Status -eq "Up" }
    if ($vpnAdapters.Count -gt 0) {
        $script:RtbLicenseInfo.SelectionColor = 'Orange'
        $script:RtbLicenseInfo.AppendText("[!] VPN connection detected.`r`n")
    }
    
    $script:RtbLicenseInfo.SelectionColor = 'Black'
    $script:RtbLicenseInfo.AppendText("[~] Attempting to establish connection in the background... This may take up to 15 seconds.`r`n")
    $script:RtbLicenseInfo.Update()

    $script:jobCompletionAction = {
        param($jobResult)
        #$CountdownTimer.Stop()
        $script:BtnComprehensiveCheck.Enabled = $true

        $script:BtnCheckServer.Enabled = $true
        $script:BtnCheckLicense.Enabled = $true
        $script:BtnCheckTimeSync.Enabled = $true
        $script:BtnCancelCheckServer.Enabled = $false

        if ($jobResult -is [System.Management.Automation.ErrorRecord]) {
            $script:RtbLicenseInfo.SelectionColor = 'Red'
            $script:RtbLicenseInfo.AppendText("An error occurred during connectivity check: $($jobResult.Exception.Message)`r`n")
        } else {
            foreach ($line in $jobResult) {
                $start = $script:RtbLicenseInfo.TextLength
                $script:RtbLicenseInfo.AppendText($line.Text)
                $script:RtbLicenseInfo.Select($start, $line.Text.Length)

                $style = [System.Drawing.FontStyle]::Regular
                if ($line.PSObject.Properties.Name -contains 'Bold' -and $line.Bold) {
                    $style = $style -bor [System.Drawing.FontStyle]::Bold
                }
                $script:RtbLicenseInfo.SelectionFont = New-Object System.Drawing.Font($script:RtbLicenseInfo.Font, $style)
                $script:RtbLicenseInfo.SelectionColor = [System.Drawing.Color]::FromName($line.Color)
            }
        }
        $script:RtbLicenseInfo.Select($script:RtbLicenseInfo.TextLength, 0)
        $script:RtbLicenseInfo.SelectionFont = $script:RtbLicenseInfo.Font
        $script:RtbLicenseInfo.SelectionColor = 'Black'

        if ($logFilePath) {
            $script:LblLicenseStatus.Text = "Primary check complete. Running tracert in background..."
            $script:LblLicenseStatus.ForeColor = 'Blue'

            $tracertJobArgs = @{
                LogFilePath = $logFilePath
                ServerName  = "lvs.core.literams.com"
            }
            Start-Job -ArgumentList $tracertJobArgs -ScriptBlock {
                param($args)
                
                function Write-Log-Append {
                    param([string]$message)
                    "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $message" | Add-Content -Path $args.LogFilePath
                }

                Write-Log-Append "--- Starting background tracert ---"
                try {
                    $tracertOutput = tracert -d $args.ServerName 2>&1 | Out-String
                    Write-Log-Append $tracertOutput
                } catch {
                    Write-Log-Append "ERROR running background tracert: $($_.Exception.Message)"
                }
                Write-Log-Append "--- Background tracert finished ---"
            } | Out-Null
        }
    }

    $logFilePath = $null
    if ($script:ChkEnableLicenseLogging.Checked) {
        try {
            $licenseLogDir = Join-Path $script:logRoot "license_logs"
            if (-not (Test-Path $licenseLogDir)) {
                New-Item -Path $licenseLogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            $logFileName = "ConnectivityCheck_$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
            $logFilePath = Join-Path $licenseLogDir $logFileName
            $script:LblLicenseStatus.Text = "Logging to $logFileName"
            $script:LblLicenseStatus.ForeColor = 'Blue'
        } catch {
            $script:LblLicenseStatus.Text = "Error creating log file: $($_.Exception.Message)"
            $script:LblLicenseStatus.ForeColor = 'Red'
        }
    }

    $jobArgs = @{
        LogFilePath = $logFilePath
    }

    $script:currentJob = Start-Job -ArgumentList $jobArgs -ScriptBlock {
        param($jobParams)

        function Write-Log {
            param([string]$message)
            if ($jobParams.LogFilePath) {
                "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $message" | Add-Content -Path $jobParams.LogFilePath
            }
        }

        $output = New-Object System.Collections.Generic.List[PSCustomObject]
        $serverName = "lvs.core.literams.com"
        $lcpPath = "C:\Program Files (x86)\Litera\Compare\lcp_main.exe"
        $lcpProcess = $null
        $connectionEstablished = $false
        $ipMatch = ""

        Write-Log "--- Starting Automated Connectivity Check for $serverName ---"

        try {
            if (-not (Test-Path $lcpPath)) {
                $output.Add([PSCustomObject]@{ Text = "[X] lcp_main.exe not found at '$lcpPath'. Cannot run automated test.`r`n"; Color = 'Red'; Bold = $true })
                Write-Log "ERROR: lcp_main.exe not found at '$lcpPath'."
                return $output
            }

            Write-Log "Attempting to resolve '$serverName'."
            try {
                $dnsResult = [System.Net.Dns]::GetHostAddresses($serverName)
                if ($dnsResult.Count -gt 0) {
                    $ipMatch = $dnsResult[0].IPAddressToString
                    Write-Log "Resolved '$serverName' to '$ipMatch'."
                }
            } catch {
                $output.Add([PSCustomObject]@{ Text = "[?] Could not resolve Litera Licensing server IP.`r`n"; Color = 'Black'; Bold = $false })
                Write-Log "ERROR: Could not resolve '$serverName'. Exception: $($_.Exception.Message)"
            }

            Write-Log "Starting lcp_main.exe in background to trigger connection."
            $lcpProcess = Start-Process -FilePath $lcpPath -WindowStyle Hidden -PassThru -ErrorAction Stop

            if ($ipMatch) {
                Write-Log "Polling netstat for connection to '$ipMatch' for up to 15 seconds..."
                foreach ($i in 1..15) {
                    Write-Log "Polling attempt $i..."
                    $netstatOutput = netstat -n | Select-String $ipMatch -ErrorAction SilentlyContinue
                    if ($netstatOutput) {
                        Write-Log "SUCCESS: Found active connection to '$ipMatch' in netstat output."
                        $connectionEstablished = $true
                        break
                    }
                    Start-Sleep -Seconds 1
                }
            }

            if ($connectionEstablished) {
                $output.Add([PSCustomObject]@{ Text = "[V] Successfully connected to Litera Licensing server ($ipMatch).`r`n"; Color = 'Green'; Bold = $false })
            } else {
                $output.Add([PSCustomObject]@{ Text = "[X] Cannot connect with Litera Licensing server.`r`n"; Color = 'Red'; Bold = $true })
                Write-Log "FAILURE: Did not find active connection to '$ipMatch' after 15 seconds."
                $firewallRules = Get-NetFirewallRule | Where-Object { $_.Direction -eq "Outbound" -and $_.Action -eq "Block" }
                if ($firewallRules.Count -gt 0) {
                    $output.Add([PSCustomObject]@{ Text = "[!] Firewall may be blocking outbound connections.`r`n"; Color = 'Orange'; Bold = $false })
                    Write-Log "INFO: Found active outbound block rules in Windows Firewall."
                }
            }

        } catch {
            $errorMessage = "An unexpected error occurred in the connectivity check job: $($_.Exception.Message)"
            Write-Log "FATAL ERROR: $errorMessage"
            $output.Add([PSCustomObject]@{ Text = "[X] $errorMessage`r`n"; Color = 'Red'; Bold = $true })
        } finally {
            if ($lcpProcess) {
                Write-Log "Stopping background lcp_main.exe process (PID: $($lcpProcess.Id))."
                Stop-Process -Id $lcpProcess.Id -Force -ErrorAction SilentlyContinue
                Write-Log "Process stopped."
            }
        }
        Write-Log "--- Primary Connectivity Check Finished ---"
        return $output
    }
    $JobTimer.Start()
})
$script:BtnCancelCheckServer.Add_Click({
    if ($script:currentJob) {
        $JobTimer.Stop()
        $CountdownTimer.Stop()
        Stop-Job -Job $script:currentJob
        Remove-Job -Job $script:currentJob -Force
        $script:currentJob = $null
        $script:jobCompletionAction = $null

        $script:BtnComprehensiveCheck.Enabled = $true
        $script:BtnCheckServer.Enabled = $true
        $script:BtnCheckLicense.Enabled = $true
        $script:BtnCheckTimeSync.Enabled = $true
        $script:BtnCancelCheckServer.Enabled = $false

        if ($script:countdownLineStart -lt $script:RtbLicenseInfo.Text.Length) {
            $script:RtbLicenseInfo.Select($script:countdownLineStart, $script:RtbLicenseInfo.Text.Length - $script:countdownLineStart)
            $script:RtbLicenseInfo.SelectedText = ""
        }
        $script:RtbLicenseInfo.SelectionColor = 'Orange'
        $script:RtbLicenseInfo.AppendText("[!] Operation cancelled by user.`r`n")
        $script:RtbLicenseInfo.SelectionColor = 'Black'

        $script:LblLicenseStatus.Text = "Operation cancelled."
        $script:LblLicenseStatus.ForeColor = 'Gray'
    }
})

$script:BtnClearLicenseOutput.Add_Click({
    $script:RtbLicenseInfo.Clear()
    $script:LblLicenseStatus.Text = "Output cleared."
    $script:LblLicenseStatus.ForeColor = 'Gray'
})

$script:BtnOpenLicenseLogFolder.Add_Click({
    $licenseLogDir = Join-Path $script:logRoot "license_logs"
    if (-not (Test-Path $licenseLogDir)) {
        New-Item -Path $licenseLogDir -ItemType Directory -Force | Out-Null
    }
    Invoke-Item $licenseLogDir
})

function Invoke-ComprehensiveLicenseCheckJob {
    param($jobParams)
    $serverName = "lvs.core.literams.com"
    $output = New-Object System.Collections.Generic.List[PSCustomObject]

    function Write-Log-Job {
        param([string]$message)
        if ($jobParams.LogFilePath) {
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $message" | Add-Content -Path $jobParams.LogFilePath
        }
    }

    function Decode-Base64Url {
        param([string]$base64Url)
        $base64 = $base64Url.Replace('-', '+').Replace('_', '/')
        switch ($base64.Length % 4) { 2 { $base64 += '==' }; 3 { $base64 += '=' } }
        try { return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64)) } catch { return $null }
    }

    function Test-LicenseServerConnectivity {
        param($output, $serverName, $ipMatch)
        Write-Log-Job "[2] Checking License Server Connectivity (continued)..."
        if ($ipMatch) {
            $netstatOutput = netstat -n | Select-String $ipMatch -ErrorAction SilentlyContinue
            if ($netstatOutput) {
                $output.Add([PSCustomObject]@{ Text = "  [✔] Successfully connected to Litera Licensing server ($ipMatch).`r`n"; Color = 'Green' })
                Write-Log-Job "  [✔] Successfully connected to Litera Licensing server ($ipMatch)."
            } else {
                $output.Add([PSCustomObject]@{ Text = "  [❌] Cannot connect with Litera Licensing server ($ipMatch).`r`n"; Color = 'Red' })
                Write-Log-Job "  [❌] Cannot connect with Litera Licensing server ($ipMatch)."
                $proxy = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings" -ErrorAction SilentlyContinue).ProxyServer
                if ($proxy) {
                    $output.Add([PSCustomObject]@{ Text = "      Note: A proxy server is configured ($proxy). This may interfere with the connection.`r`n"; Color = 'Orange' })
                    Write-Log-Job "      Note: A proxy server is configured ($proxy). This may interfere with the connection."
                }
                $vpnAdapters = Get-NetAdapter | Where-Object { $_.InterfaceDescription -match "VPN" -and $_.Status -eq "Up" }
                if ($vpnAdapters.Count -gt 0) {
                    $output.Add([PSCustomObject]@{ Text = "      Note: An active VPN connection was detected. This may interfere with the connection.`r`n"; Color = 'Orange' })
                    Write-Log-Job "      Note: An active VPN connection was detected. This may interfere with the connection."
                }
                if ((Get-NetFirewallRule | Where-Object { $_.Direction -eq "Outbound" -and $_.Action -eq "Block" }).Count -gt 0) {
                    $output.Add([PSCustomObject]@{ Text = "      Note: Firewall may be blocking outbound connections.`r`n"; Color = 'Orange' })
                    Write-Log-Job "      Note: Firewall may be blocking outbound connections."
                }
            }
        }
    }

    function Test-ComprehensivePrerequisites {
        param($output)
        $output.Add([PSCustomObject]@{ Text = "`r`n[3] Checking Pre-requisites...`r`n"; Color = 'DarkBlue'; Bold = $true })
        Write-Log-Job "[3] Checking Pre-requisites..."
        $keys = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        $installedApps = Get-ItemProperty $keys -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisplayName -ErrorAction SilentlyContinue

        $vc86 = $installedApps | Where-Object { $_ -like '*Visual C++*2015-2022*Redistributable*(x86)*' } | Select-Object -First 1
        if ($vc86) {
            $output.Add([PSCustomObject]@{ Text = "  [✔] Found: $vc86`r`n"; Color = 'Green' })
            Write-Log-Job "  [✔] Found: $vc86"
        }
        else {
            $output.Add([PSCustomObject]@{ Text = "  [❌] VC++ Redistributable (x86) is NOT installed.`r`n"; Color = 'Red' })
            Write-Log-Job "  [❌] VC++ Redistributable (x86) is NOT installed."
        }
        $vc64 = $installedApps | Where-Object { $_ -like '*Visual C++*2015-2022*Redistributable*(x64)*' } | Select-Object -First 1
        if ($vc64) {
            $output.Add([PSCustomObject]@{ Text = "  [✔] Found: $vc64`r`n"; Color = 'Green' })
            Write-Log-Job "  [✔] Found: $vc64"
        }
        else {
            $output.Add([PSCustomObject]@{ Text = "  [❌] VC++ Redistributable (x64) is NOT installed.`r`n"; Color = 'Red' })
            Write-Log-Job "  [❌] VC++ Redistributable (x64) is NOT installed."
        }
    }

    function Test-ComprehensiveDlls {
        param($output)
        $output.Add([PSCustomObject]@{ Text = "`r`n[4] Checking necessary DLLs...`r`n"; Color = 'DarkBlue'; Bold = $true })
        Write-Log-Job "[4] Checking necessary DLLs..."
        $installDir = "C:\Program Files (x86)\Litera\Compare"
        $dlls = @("Litera.Licensing.Client.dll", "Litera.Compare.License.dll", "Microsoft.IdentityModel.JsonWebTokens.dll", "Microsoft.IdentityModel.Logging.dll", "Microsoft.IdentityModel.Protocols.dll", "Microsoft.IdentityModel.Protocols.OpenIdConnect.dll", "Microsoft.IdentityModel.Tokens.dll", "Newtonsoft.Json.dll", "System.IdentityModel.Tokens.Jwt.dll")
        foreach ($dll in $dlls) {
            if (Test-Path (Join-Path $installDir $dll)) {
                $output.Add([PSCustomObject]@{ Text = "  [✔] Found: $dll`r`n"; Color = 'Green' })
                Write-Log-Job "  [✔] Found: $dll"
            }
            else {
                $output.Add([PSCustomObject]@{ Text = "  [❌] Missing: $dll in $installDir`r`n"; Color = 'Red' })
                Write-Log-Job "  [❌] Missing: $dll in $installDir"
            }
        }
    }

    function Test-ComprehensiveRegistryKeys {
        param($output)
        $output.Add([PSCustomObject]@{ Text = "`r`n[5] Checking Registry Key...`r`n"; Color = 'DarkBlue'; Bold = $true })
        Write-Log-Job "[5] Checking Registry Key..."
        $regPath = "HKLM:\SOFTWARE\WOW6432Node\Litera\ChangePro"
        if ((Get-ItemProperty -Path $regPath -Name "lk.cd.serv" -ErrorAction SilentlyContinue) -ne $null) {
            $output.Add([PSCustomObject]@{ Text = "  [✔] Registry key 'lk.cd.serv' found in '$regPath'.`r`n"; Color = 'Green' })
            Write-Log-Job "  [✔] Registry key 'lk.cd.serv' found in '$regPath'."
        }
        else {
            $output.Add([PSCustomObject]@{ Text = "  [❌] Registry key 'lk.cd.serv' NOT found in $regPath`r`n"; Color = 'Red' })
            Write-Log-Job "  [❌] Registry key 'lk.cd.serv' NOT found in $regPath"
        }
    }

    function Test-ComprehensiveTimeSync {
        param($output)
        $output.Add([PSCustomObject]@{ Text = "`r`n[6] Checking Time Synchronization...`r`n"; Color = 'DarkBlue'; Bold = $true })
        Write-Log-Job "[6] Checking Time Synchronization..."
        try {
            $timeService = Get-CimInstance -ClassName Win32_Service -Filter "Name='W32Time'"
            if ($timeService.StartMode -eq 'Auto' -and $timeService.State -eq 'Running') {
                $output.Add([PSCustomObject]@{ Text = "  [✔] Windows Time service (W32Time) is running and set to Automatic.`r`n"; Color = 'Green' })
                Write-Log-Job "  [✔] Windows Time service (W32Time) is running and set to Automatic."
            }
            else {
                $output.Add([PSCustomObject]@{ Text = "  [❌] Windows Time service (W32Time) is not running or not set to Automatic. Please sync time manually.`r`n"; Color = 'Red' })
                Write-Log-Job "  [❌] Windows Time service (W32Time) is not running or not set to Automatic."
            }
            $syncStatus = w32tm /query /status /verbose
            if ($syncStatus -match 'Leap Indicator: 0\(no warning\)') {
                $output.Add([PSCustomObject]@{ Text = "  [✔] System time appears to be synchronized.`r`n"; Color = 'Green' })
                Write-Log-Job "  [✔] System time appears to be synchronized."
            }
            else {
                $output.Add([PSCustomObject]@{ Text = "  [❌] System time may not be synchronized. Please sync time manually.`r`n"; Color = 'Red' })
                Write-Log-Job "  [❌] System time may not be synchronized."
            }
        } catch {
            $output.Add([PSCustomObject]@{ Text = "  [❌] Could not verify time service status: $($_.Exception.Message)`r`n"; Color = 'Red' })
            Write-Log-Job "  [❌] Could not verify time service status: $($_.Exception.Message)"
        }
    }

    function Test-ComprehensiveUserFiles {
        param($output, $SelectedUserName, $CurrentUser)
        $output.Add([PSCustomObject]@{ Text = "`r`n[7] Checking User Files for '$SelectedUserName'...`r`n"; Color = 'DarkBlue'; Bold = $true })
        Write-Log-Job "[7] Checking User Files for '$SelectedUserName'..."
        $targetAppData = if ($SelectedUserName -eq $CurrentUser) { $env:APPDATA } else { "C:\Users\$SelectedUserName\AppData\Roaming" }
        $compareFolder = Join-Path $targetAppData "Litera\Compare"
        $output.Add([PSCustomObject]@{ Text = "  Path: $compareFolder`r`n"; Color = 'Gray' })
        Write-Log-Job "  Path: $compareFolder"
        if (Test-Path (Join-Path $compareFolder "key")) {
            $output.Add([PSCustomObject]@{ Text = "  [✔] 'key' file found.`r`n"; Color = 'Green' })
            Write-Log-Job "  [✔] 'key' file found."
        }
        else {
            $output.Add([PSCustomObject]@{ Text = "  [❌] 'key' file NOT found in $compareFolder`r`n"; Color = 'Red' })
            Write-Log-Job "  [❌] 'key' file NOT found in $compareFolder"
        }
        $tokenFile = Join-Path $compareFolder "token.json"
        if (Test-Path $tokenFile) {
            $output.Add([PSCustomObject]@{ Text = "  [✔] 'token.json' file found.`r`n"; Color = 'Green' });
            Write-Log-Job "  [✔] 'token.json' file found."
            try {
                $rawToken = Get-Content -Path $jobParams.TokenFilePath -Raw | ConvertFrom-Json -ErrorAction Stop
                $jwt = $rawToken.AccessToken
                $parts = $jwt -split '\.'
                $payloadJsonText = Decode-Base64Url $parts[1]
                $payloadObj = $payloadJsonText | ConvertFrom-Json -ErrorAction Stop

                $output.Add([PSCustomObject]@{ Text = "    Key    : $($payloadObj.sub)`r`n"; Color = 'Black' })
                Write-Log-Job "    Key    : $($payloadObj.sub)"
                $products = @()
                if ($payloadObj.product -is [System.Collections.IEnumerable] -and -not ($payloadObj.product -is [string])) { $products = $payloadObj.product }
                else { $products = @($payloadObj.product) }
                foreach ($prod in $products) {
                    if ($prod -match "compare\.desktop") {
                        if ($prod -match "val\s*(\d+)") {
                            $unixVal = [int64]$Matches[1]
                            $expiry  = [DateTimeOffset]::FromUnixTimeSeconds($unixVal).DateTime.ToString("yyyy-MM-dd HH:mm:ss")
                            $output.Add([PSCustomObject]@{ Text = "    Product: compare.desktop`r`n"; Color = 'Black' })
                            $output.Add([PSCustomObject]@{ Text = "    Expiry : $expiry`r`n"; Color = 'Black' })
                            Write-Log-Job "    Product: compare.desktop"
                            Write-Log-Job "    Expiry : $expiry"
                        }
                    }
                }
            } catch {
                $output.Add([PSCustomObject]@{ Text = "    [❌] Failed to parse token.json: $($_.Exception.Message)`r`n"; Color = 'Red' })
                Write-Log-Job "    [❌] Failed to parse token.json: $($_.Exception.Message)"
            }
        } else {
            $output.Add([PSCustomObject]@{ Text = "  [❌] 'token.json' file NOT found in $compareFolder`r`n"; Color = 'Red' })
            Write-Log-Job "  [❌] 'token.json' file NOT found in $compareFolder"
        }
    }

    # 2. License Server Connection (Automated)
    $lcpPath = "C:\Program Files (x86)\Litera\Compare\lcp_main.exe"
    $lcpProcess = $null
    $connectionEstablished = $false
    $ipMatch = ""

    Write-Log-Job "[2] Checking License Server Connectivity (automated)..."
    try {
        if (-not (Test-Path $lcpPath)) {
            $output.Add([PSCustomObject]@{ Text = "  [❌] lcp_main.exe not found at '$lcpPath'. Cannot run automated test.`r`n"; Color = 'Red' })
            Write-Log-Job "  [❌] lcp_main.exe not found at '$lcpPath'."
        } else {
            try {
                $dnsResult = [System.Net.Dns]::GetHostAddresses($serverName)
                if ($dnsResult.Count -gt 0) { $ipMatch = $dnsResult[0].IPAddressToString }
            } catch {
                Write-Log-Job "  [!] Could not resolve '$serverName'. Exception: $($_.Exception.Message)"
            }

            Write-Log-Job "Starting lcp_main.exe in background to trigger connection."
            $lcpProcess = Start-Process -FilePath $lcpPath -WindowStyle Hidden -PassThru -ErrorAction Stop

            if ($ipMatch) {
                Write-Log-Job "Polling netstat for connection to '$ipMatch' for up to 15 seconds..."
                foreach ($i in 1..15) {
                    $netstatOutput = netstat -n | Select-String $ipMatch -ErrorAction SilentlyContinue
                    if ($netstatOutput) {
                        $connectionEstablished = $true
                        break
                    }
                    Start-Sleep -Seconds 1
                }
            }
            
            # Call the existing function with the result of our automated check
            Test-LicenseServerConnectivity -output $output -serverName $serverName -ipMatch $ipMatch -ConnectionEstablished $connectionEstablished
        }
    } catch {
        $errorMessage = "An unexpected error occurred during automated connectivity check: $($_.Exception.Message)"
        Write-Log-Job "FATAL ERROR: $errorMessage"
        $output.Add([PSCustomObject]@{ Text = "  [❌] $errorMessage`r`n"; Color = 'Red' })
    } finally {
        if ($lcpProcess) {
            Write-Log-Job "Stopping background lcp_main.exe process (PID: $($lcpProcess.Id))."
            Stop-Process -Id $lcpProcess.Id -Force -ErrorAction SilentlyContinue
            Write-Log-Job "Process stopped."
        }
    }

    Test-ComprehensivePrerequisites -output $output
    Test-ComprehensiveDlls -output $output
    Test-ComprehensiveRegistryKeys -output $output
    Test-ComprehensiveTimeSync -output $output
    Test-ComprehensiveUserFiles -output $output -SelectedUserName $jobParams.SelectedUserName -CurrentUser $jobParams.CurrentUser
    
    return $output
}

<#
.SYNOPSIS
    Runs a comprehensive suite of licensing checks.
.DESCRIPTION
    Initiates a series of synchronous and asynchronous checks to diagnose common licensing issues. This includes internet connectivity, DNS resolution, pre-requisites, necessary DLLs, registry keys, time sync, and user license files. It provides real-time feedback in the UI and runs slower checks in a background job.
#>
function Start-ComprehensiveLicenseCheck {
    if ($script:currentJob) { [System.Windows.Forms.MessageBox]::Show("Another operation is already in progress.", "Busy", "OK", "Warning"); return }

    $BtnComprehensiveCheck.Enabled = $false
    $BtnCheckServer.Enabled = $false
    $BtnCheckLicense.Enabled = $false
    $BtnCheckTimeSync.Enabled = $false
    $BtnCancelCheckServer.Enabled = $true

    $RtbLicenseInfo.Clear()
    $RtbLicenseInfo.SelectionColor = 'Black'
    $RtbLicenseInfo.AppendText("[~] Starting Comprehensive License Check...`r`n")

    $logFilePath = $null
    if ($ChkEnableLicenseLogging.Checked) {
        try {
            $licenseLogDir = Join-Path $script:logRoot "license_logs"
            if (-not (Test-Path $licenseLogDir)) {
                New-Item -Path $licenseLogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            $logFileName = "ComprehensiveCheck_$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
            $logFilePath = Join-Path $licenseLogDir $logFileName
            $LblLicenseStatus.Text = "Logging to $logFileName"
            $LblLicenseStatus.ForeColor = 'Blue'
        } catch {
            $LblLicenseStatus.Text = "Error creating log file: $($_.Exception.Message)"
            $LblLicenseStatus.ForeColor = 'Red'
        }
    }

    $logAction = if ($logFilePath) {
        { param($Message) "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $Message" | Add-Content -Path $logFilePath }
    } else {
        { param($Message) # Do nothing
        }
    }

    & $logAction "[~] Starting Comprehensive License Check..."

    $selectedUserName = ($ComboLicenseUsers.SelectedItem -split ' ')[0]
    if ([string]::IsNullOrWhiteSpace($selectedUserName)) {
        $selectedUserName = $env:USERNAME
    }

    $targetAppData = if ($selectedUserName -eq $env:USERNAME) {
        $env:APPDATA
    } else {
        "C:\Users\$selectedUserName\AppData\Roaming"
    }
    $tokenFilePath = Join-Path $targetAppData "Litera\Compare\token.json"

    & $logAction "Selected user for check: $selectedUserName"


    # --- Synchronous (fast) checks to provide immediate feedback ---
    $RtbLicenseInfo.SelectionFont = New-Object System.Drawing.Font($RtbLicenseInfo.Font, [System.Drawing.FontStyle]::Bold)
    $RtbLicenseInfo.SelectionColor = 'DarkBlue'
    $RtbLicenseInfo.AppendText("`r`n[1] Checking Internet Connection...`r`n")
    & $logAction "[1] Checking Internet Connection..."
    if (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet) {
        $RtbLicenseInfo.SelectionColor = 'Green'
        $RtbLicenseInfo.AppendText("  [✔] Internet connection successful.`r`n")
        & $logAction "  [✔] Internet connection successful."
    } else {
        $RtbLicenseInfo.SelectionColor = 'Red'
        $RtbLicenseInfo.AppendText("  [❌] No Internet connection detected. Further checks may fail.`r`n")
        & $logAction "  [❌] No Internet connection detected. Further checks may fail."
    }

    $RtbLicenseInfo.SelectionFont = New-Object System.Drawing.Font($RtbLicenseInfo.Font, [System.Drawing.FontStyle]::Bold)
    $RtbLicenseInfo.SelectionColor = 'DarkBlue'
    $RtbLicenseInfo.AppendText("`r`n[2] Verifying License Server Connection...`r`n")
    & $logAction "[2] Verifying License Server Connection..."
    $serverName = "lvs.core.literams.com"
    $ipMatch = ""
    try {
        $dnsResult = [System.Net.Dns]::GetHostAddresses($serverName)
        if ($dnsResult.Count -gt 0) {
            $ipMatch = $dnsResult[0].IPAddressToString
            $RtbLicenseInfo.SelectionColor = 'Green'
            $RtbLicenseInfo.AppendText("  [✔] Resolved '$serverName' to '$ipMatch'.`r`n")
            & $logAction "  [✔] Resolved '$serverName' to '$ipMatch'."
        }
    } catch {
        $RtbLicenseInfo.SelectionColor = 'Red'
        $RtbLicenseInfo.AppendText("  [❌] Could not resolve Litera Licensing server IP: $($_.Exception.Message)`r`n")
        & $logAction "  [❌] Could not resolve Litera Licensing server IP: $($_.Exception.Message)"
    }

    $RtbLicenseInfo.SelectionFont = $RtbLicenseInfo.Font
    $RtbLicenseInfo.SelectionColor = 'Black'
    $RtbLicenseInfo.AppendText("  [~] Attempting to establish connection and run checks in the background...`r`n")

    $RtbLicenseInfo.Update() # Force the UI to repaint now

    # --- Asynchronous (slow) checks ---
    $script:jobCompletionAction = {
        param($jobResult)
        #$CountdownTimer.Stop()
        $script:BtnComprehensiveCheck.Enabled = $true
        $script:BtnCheckServer.Enabled = $true
        $script:BtnCheckLicense.Enabled = $true
        $script:BtnCheckTimeSync.Enabled = $true
        $script:BtnCancelCheckServer.Enabled = $false

        if ($jobResult -is [System.Management.Automation.ErrorRecord]) {
            $RtbLicenseInfo.SelectionColor = 'Red'
            $RtbLicenseInfo.AppendText("An error occurred during the check: $($jobResult.Exception.Message)`r`n")
            return # Stop further processing on error
        }

        # Process and display the results from the background job
        foreach ($line in $jobResult) {
            $start = $RtbLicenseInfo.TextLength
            $RtbLicenseInfo.AppendText($line.Text)
            $RtbLicenseInfo.Select($start, $line.Text.Length)

            $style = if ($line.PSObject.Properties.Name -contains 'Bold' -and $line.Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
            $RtbLicenseInfo.SelectionFont = New-Object System.Drawing.Font($RtbLicenseInfo.Font, $style)
            $RtbLicenseInfo.SelectionColor = [System.Drawing.Color]::FromName($line.Color)
        }

        $RtbLicenseInfo.Select($RtbLicenseInfo.TextLength, 0)
        $RtbLicenseInfo.SelectionFont = $RtbLicenseInfo.Font
        $RtbLicenseInfo.SelectionColor = 'Black'
        $RtbLicenseInfo.AppendText("`r`n[~] Comprehensive Check Finished.`r`n")
    }

    $jobArgs = @{
        LogFilePath      = $logFilePath
        SelectedUserName = $selectedUserName
        CurrentUser      = $env:USERNAME
        TokenFilePath    = $tokenFilePath
    }

    $script:currentJob = Start-Job -ArgumentList $jobArgs -ScriptBlock ${function:Invoke-ComprehensiveLicenseCheckJob}
    $JobTimer.Start()
}

<#
.SYNOPSIS
    Defines logging controls for the Install/Uninstall tab.
.DESCRIPTION
    This section creates the UI elements (CheckBox, Button) for enabling and managing logging of installation and uninstallation operations.
#>
#==============================================================================
# --- UI: Logging Controls for Install/Uninstall Tab ---
#==============================================================================
$LoggingPanel = New-Object System.Windows.Forms.FlowLayoutPanel
#==============================================================================
# --- UI: Compare Customizations Tab ---
#==============================================================================
$CompareLayout = New-Object System.Windows.Forms.TableLayoutPanel
$CompareLayout.Dock = 'Fill'
$CompareLayout.BackColor = [System.Drawing.SystemColors]::Window
$CompareLayout.Padding = (New-Object System.Windows.Forms.Padding(10))
$CompareLayout.ColumnCount = 3
$CompareLayout.RowCount = 4
$CompareLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$CompareLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$CompareLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$CompareLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$CompareLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$CompareLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$CompareLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$TabCompareCustom.Controls.Add($CompareLayout)

$LblOriginalXml = New-Object System.Windows.Forms.Label; $LblOriginalXml.Text = "Original File:"; $LblOriginalXml.Dock = 'Fill'; $LblOriginalXml.TextAlign = 'MiddleLeft'
$TxtOriginalXmlPath = New-Object System.Windows.Forms.TextBox; $TxtOriginalXmlPath.Dock = 'Fill' 
$BtnBrowseOriginalXml = New-AppButton -Text "Browse..." -AutoSize

$LblModifiedXml = New-Object System.Windows.Forms.Label; $LblModifiedXml.Text = "Modified File:"; $LblModifiedXml.Dock = 'Fill'; $LblModifiedXml.TextAlign = 'MiddleLeft'
$TxtModifiedXmlPath = New-Object System.Windows.Forms.TextBox; $TxtModifiedXmlPath.Dock = 'Fill'
$BtnBrowseModifiedXml = New-AppButton -Text "Browse..." -AutoSize

$compareActionsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$compareActionsPanel.Dock = 'Fill'; $compareActionsPanel.FlowDirection = 'LeftToRight'; $compareActionsPanel.Padding = (New-Object System.Windows.Forms.Padding(0, 5, 0, 5))
$BtnCompareXml = New-AppButton -Text "Compare" -AutoSize -BackColor 'LightGreen' -Enabled $false
$compareActionsPanel.Controls.Add($BtnCompareXml)

$DgvCompareResults = New-Object System.Windows.Forms.DataGridView
$DgvCompareResults.Dock = 'Fill'
$DgvCompareResults.ReadOnly = $true
$DgvCompareResults.AllowUserToAddRows = $false; $DgvCompareResults.AllowUserToDeleteRows = $false
$DgvCompareResults.RowHeadersVisible = $false
$DgvCompareResults.AutoSizeColumnsMode = 'Fill'
$DgvCompareResults.BackgroundColor = 'White'
$DgvCompareResults.MultiSelect = $false

$CompareLayout.Controls.Add($LblOriginalXml, 0, 0)
$CompareLayout.Controls.Add($TxtOriginalXmlPath, 1, 0)
$CompareLayout.Controls.Add($BtnBrowseOriginalXml, 2, 0)

$CompareLayout.Controls.Add($LblModifiedXml, 0, 1)
$CompareLayout.Controls.Add($TxtModifiedXmlPath, 1, 1)
$CompareLayout.Controls.Add($BtnBrowseModifiedXml, 2, 1)

$CompareLayout.SetColumnSpan($compareActionsPanel, 3)
$CompareLayout.Controls.Add($compareActionsPanel, 0, 2)

$CompareLayout.SetColumnSpan($DgvCompareResults, 3)
$CompareLayout.Controls.Add($DgvCompareResults, 0, 3)

$BtnBrowseOriginalXml.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = 'XML files (*.xml)|*.xml|All files (*.*)|*.*'
    $openFileDialog.Title = "Select an Original Customization File"
    $initialDir = "C:\ProgramData\Litera\customize"
    if (Test-Path -Path $initialDir -PathType Container) {
        $openFileDialog.InitialDirectory = $initialDir
    }
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $TxtOriginalXmlPath.Text = $openFileDialog.FileName
    }
})

$BtnBrowseModifiedXml.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = 'XML files (*.xml)|*.xml|All files (*.*)|*.*'
    $openFileDialog.Title = "Select a Modified Customization File"

    # If a modified file was previously selected in this session, start in its directory.
    if ((-not [string]::IsNullOrWhiteSpace($script:lastModifiedXmlPath)) -and (Test-Path ([System.IO.Path]::GetDirectoryName($script:lastModifiedXmlPath)))) {
        $openFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($script:lastModifiedXmlPath)
    }
    # Otherwise, default to the user's Documents folder.
    else {
        $openFileDialog.InitialDirectory = [System.Environment]::GetFolderPath('MyDocuments')
    }

    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $TxtModifiedXmlPath.Text = $openFileDialog.FileName
        # Remember this location for next time.
        $script:lastModifiedXmlPath = $openFileDialog.FileName
    }
})

$checkXmlPaths = {
    $BtnCompareXml.Enabled = (-not [string]::IsNullOrWhiteSpace($TxtOriginalXmlPath.Text)) -and (-not [string]::IsNullOrWhiteSpace($TxtModifiedXmlPath.Text))
}
$TxtOriginalXmlPath.Add_TextChanged($checkXmlPaths)
$TxtModifiedXmlPath.Add_TextChanged($checkXmlPaths)

$BtnCompareXml.Add_Click({
    $Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $originalFile = $TxtOriginalXmlPath.Text
        $modifiedFile = $TxtModifiedXmlPath.Text

        $originalSettings = Get-CustomizationData -FilePath $originalFile
        $modifiedSettings = Get-CustomizationData -FilePath $modifiedFile

        if ($null -eq $originalSettings -or $null -eq $modifiedSettings) {
            return
        }

        $allKeys = ($originalSettings.Keys + $modifiedSettings.Keys) | Select-Object -Unique

        $results = [System.Collections.Generic.List[psobject]]::new()

        foreach ($key in $allKeys) {
            $original = $originalSettings[$key]
            $modified = $modifiedSettings[$key]

            $originalValue = if ($original) { $original.Value } else { '[Not Found]' }
            $modifiedValue = if ($modified) { $modified.Value } else { '[Not Found]' }

            if ($originalValue -ne $modifiedValue) {
                $parentName = if ($original) { $original.Parent } else { $modified.Parent }

                $results.Add([PSCustomObject]@{
                    'Group'                 = $parentName
                    'Customization Field'   = $key
                    'Original Value'        = $originalValue
                    'Modified Value'        = $modifiedValue
                })
            }
        }
    
        $sortedResults = $results | Sort-Object 'Group', 'Customization Field'

        $dataTable = New-Object System.Data.DataTable
        $null = $dataTable.Columns.Add('#', [int])
        $null = $dataTable.Columns.Add('Group')
        $null = $dataTable.Columns.Add('Customization Field')
        $null = $dataTable.Columns.Add('Original Value')
        $null = $dataTable.Columns.Add('Modified Value')

        if ($sortedResults.Count -gt 0) {
            $rowNum = 1
            foreach ($result in $sortedResults) {
                $row = $dataTable.NewRow()
                $row['#'] = $rowNum++
                $row['Group'] = $result.Group
                $row['Customization Field'] = $result.'Customization Field'
                $row['Original Value'] = $result.'Original Value'
                $row['Modified Value'] = $result.'Modified Value'
                $dataTable.Rows.Add($row)
            }
        }

        $DgvCompareResults.DataSource = $null
        $DgvCompareResults.DataSource = $dataTable
        
        [System.Windows.Forms.MessageBox]::Show("$($results.Count) differences found.", "Comparison Complete", "OK", "Information")
    }
    finally {
        $Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$DgvCompareResults.add_DataBindingComplete({
    param($sender, $e)
    if ($sender.Columns.Contains('#')) {
        $sender.Columns['#'].AutoSizeMode = 'AllCells'
        $sender.Columns['#'].DefaultCellStyle.Alignment = 'MiddleCenter'
        $sender.Columns['Group'].FillWeight = 25
        $sender.Columns['Customization Field'].FillWeight = 40
        $sender.Columns['Original Value'].FillWeight = 17
        $sender.Columns['Modified Value'].FillWeight = 17
    }

    foreach($row in $sender.Rows){
        if($row.Cells['Original Value'].Value -eq '[Not Found]' -or $row.Cells['Modified Value'].Value -eq '[Not Found]'){
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
        }
        else {
            $row.DefaultCellStyle.BackColor = $DgvCompareResults.DefaultCellStyle.BackColor
        }
    }
})

<#
.SYNOPSIS
    Creates and configures all UI elements for the "Office Add-in Management" tab.
.DESCRIPTION
    This function encapsulates the creation of all controls, panels, and layouts for the Office Add-in Management tab.
    It uses the script scope for control variables so they can be accessed by event handlers later in the script.
.RETURNS
    A `System.Windows.Forms.TableLayoutPanel` containing the complete UI for the tab.
#>
function Create-AddinMgmtTabUI {
    $AddinMgmtLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $AddinMgmtLayout.Dock = 'Fill'
    $AddinMgmtLayout.BackColor = [System.Drawing.SystemColors]::Window
    $AddinMgmtLayout.Padding = (New-Object System.Windows.Forms.Padding(10))
    $AddinMgmtLayout.ColumnCount = 1
    $AddinMgmtLayout.RowCount = 3
    [void]$AddinMgmtLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$AddinMgmtLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$AddinMgmtLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

    $addinActionsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $addinActionsPanel.Dock = 'Fill'; $addinActionsPanel.FlowDirection = 'LeftToRight'; $addinActionsPanel.Padding = (New-Object System.Windows.Forms.Padding(0, 0, 0, 5))
    $script:BtnRefreshAddins = New-AppButton -Text "Scan for Add-ins" -AutoSize -BackColor 'LightBlue'
    $script:ToolTip.SetToolTip($script:BtnRefreshAddins, "Scans the registry for all Litera-related Office COM add-ins.")
    $script:BtnEnableAddin = New-AppButton -Text "Enable Selected" -AutoSize -BackColor 'LightGreen'
    $script:ToolTip.SetToolTip($script:BtnEnableAddin, "Enables the selected add-in(s) by setting their LoadBehavior to 3 (Load at Startup).")
    $script:BtnDisableAddin = New-AppButton -Text "Disable Selected" -AutoSize -BackColor 'MistyRose'
    $script:ToolTip.SetToolTip($script:BtnDisableAddin, "Disables the selected add-in(s) by setting their LoadBehavior to 2 (Do not load automatically).")
    $script:BtnAddToDontDisable = New-AppButton -Text "Don't Disable" -AutoSize -BackColor 'LightYellow'
    $script:ToolTip.SetToolTip($script:BtnAddToDontDisable, "Adds the selected add-in(s) to the 'Do Not Disable' list in the registry to improve resiliency.")
    $script:BtnClearResilience = New-AppButton -Text "Clear Resilience Keys" -AutoSize -BackColor 'LightSalmon'
    $script:ToolTip.SetToolTip($script:BtnClearResilience, "Removes the selected add-in(s) from resiliency lists (DoNotDisableAddinList, DisabledItems) to reset Office's auto-disable behavior for them.")
    $script:BtnOpenAddinRegKey = New-AppButton -Text "Open Registry Key" -AutoSize -BackColor 'LightGray'
    $script:ToolTip.SetToolTip($script:BtnOpenAddinRegKey, "Opens the selected add-in's key in the Registry Editor.")
    [void]$addinActionsPanel.Controls.AddRange(@($script:BtnRefreshAddins, $script:BtnEnableAddin, $script:BtnDisableAddin, $script:BtnAddToDontDisable, $script:BtnClearResilience, $script:BtnOpenAddinRegKey))
    [void]$AddinMgmtLayout.Controls.Add($addinActionsPanel, 0, 0)

    $script:DgvAddins = New-Object System.Windows.Forms.DataGridView
    $script:DgvAddins.Dock = 'Fill'
    $script:DgvAddins.ReadOnly = $true
    $script:DgvAddins.AllowUserToAddRows = $false; $script:DgvAddins.AllowUserToDeleteRows = $false
    $script:DgvAddins.RowHeadersVisible = $false
    $script:DgvAddins.AutoSizeColumnsMode = 'Fill'
    $script:DgvAddins.BackgroundColor = 'White'
    $script:DgvAddins.SelectionMode = 'FullRowSelect'
    [void]$AddinMgmtLayout.Controls.Add($script:DgvAddins, 0, 1)

    $script:LblAddinStatus = New-StatusLabel -Text "Ready. Click 'Scan for Add-ins' to begin." -InitialColor 'Gray' -Dock 'Fill'
    [void]$AddinMgmtLayout.Controls.Add($script:LblAddinStatus, 0, 2)

    return $AddinMgmtLayout
}

<#
.SYNOPSIS
    Defines the UI and logic for the "Office Add-in Management" tab.
.DESCRIPTION
    This section builds the UI for scanning, viewing, and managing Litera Office COM Add-ins. It provides functionality to enable, disable, and manage the resiliency settings for add-ins.
#>
#==============================================================================
# --- UI: Office Add-in Management Tab ---
#==============================================================================
$addinMgmtTabPanel = Create-AddinMgmtTabUI
$script:TabAddinMgmt.Controls.Add($addinMgmtTabPanel)

<#
.SYNOPSIS
    Scans the registry for Litera-related Office COM add-ins.
.DESCRIPTION
    Queries HKCU and HKLM registry hives for Office add-ins across Word, Excel, Outlook, and PowerPoint. It filters for add-ins published by Litera and gathers details like friendly name, ProgID, and LoadBehavior.
.RETURNS
    An array of PSCustomObjects, each representing a found add-in.
#>
function Get-LiteraOfficeAddins {
    $allAddins = [System.Collections.Generic.List[PSCustomObject]]::new()
    $officeApps = @("Word", "Excel", "Outlook", "PowerPoint")
    $hives = @(
        @{ Scope = "User";            Path = "HKCU:\Software\Microsoft\Office" },
        @{ Scope = "Machine";         Path = "HKLM:\Software\Microsoft\Office" },
        @{ Scope = "Machine (Wow64)"; Path = "HKLM:\Software\Wow6432Node\Microsoft\Office" }
    )

    # 1. Gather all potential add-in keys from all locations
    foreach ($hive in $hives) {
        foreach ($app in $officeApps) {
            $registryPath = Join-Path $hive.Path "$app\Addins"
            if (-not (Test-Path $registryPath)) { continue }

            try {
                $addinKeys = Get-ChildItem -Path $registryPath -ErrorAction Stop
                foreach ($key in $addinKeys) {
                    $properties = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                    $progId = $key.PSChildName

                    # 2. Filter for Litera add-ins
                    if (($progId -like "*Litera*") -or ($properties.FriendlyName -like "*Litera*")) {
                        $allAddins.Add([PSCustomObject]@{
                            Scope        = $hive.Scope
                            Application  = $app
                            ProgID       = $progId
                            LoadBehavior = $properties.LoadBehavior
                            FriendlyName = $properties.FriendlyName # Temporary name
                            RegistryPath = $key.PSPath
                        })
                    }
                }
            } catch {
                Write-Log "Could not access add-in registry path: $registryPath. Error: $($_.Exception.Message)"
            }
        }
    }

    # 3. De-duplicate based on Application and ProgID, giving preference to User scope
    $uniqueAddins = $allAddins | Sort-Object -Property @{Expression="Application"}, @{Expression="ProgID"}, @{Expression={ $_.Scope -eq 'User' }; Descending=$true} | Group-Object -Property Application, ProgID | ForEach-Object { $_.Group[0] }

    # 4. Batch-resolve friendly names from HKEY_CLASSES_ROOT for better performance
    $progIdsToResolve = $uniqueAddins.ProgID | Select-Object -Unique
    $hkcrNames = @{}
    try {
        Get-Item -Path ($progIdsToResolve | ForEach-Object { "Registry::HKEY_CLASSES_ROOT\$_" }) -ErrorAction SilentlyContinue | ForEach-Object {
            $hkcrNames[$_.PSChildName] = $_.GetValue('')
        }
    } catch {
        Write-Log "Error during batch lookup of friendly names in HKCR. Some names may be incorrect. Error: $($_.Exception.Message)"
    }

    # 5. Finalize the objects with the best friendly name and formatted LoadBehavior
    $finalResults = foreach ($addin in $uniqueAddins) {
        $displayName = $hkcrNames[$addin.ProgID]
        if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $addin.FriendlyName }
        if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $addin.ProgID }

        $loadBehaviorText = switch ($addin.LoadBehavior) {
            0       { "0 - Disconnected" }
            1       { "1 - Connected" }
            2       { "2 - Disabled (On Demand)" }
            3       { "3 - Enabled (Startup)" }
            8       { "8 - On Demand (Unloaded)" }
            9       { "9 - On Demand (Loaded)" }
            16      { "16 - Connect First Time" }
            default { "Unknown ($($addin.LoadBehavior))" }
        }

        [PSCustomObject]@{
            Scope        = $addin.Scope
            Application  = $addin.Application
            FriendlyName = $displayName
            ProgID       = $addin.ProgID
            LoadBehavior = $loadBehaviorText
            RegistryPath = $addin.RegistryPath
        }
    }

    return $finalResults | Sort-Object Application, FriendlyName
}

<#
.SYNOPSIS
    Populates the DataGridView on the Add-in Management tab.
#>
function Populate-AddinGrid {
    $script:LblAddinStatus.Text = "Scanning for Litera add-ins..."; $script:LblAddinStatus.ForeColor = 'Black'
    $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $script:DgvAddins.DataSource = $null

    try {
        $addins = Get-LiteraOfficeAddins
        
        $dataTable = New-Object System.Data.DataTable
        $null = $dataTable.Columns.Add('Scope')
        $null = $dataTable.Columns.Add('Application')
        $null = $dataTable.Columns.Add('Friendly Name')
        $null = $dataTable.Columns.Add('Status (LoadBehavior)')
        $null = $dataTable.Columns.Add('ProgID')
        $null = $dataTable.Columns.Add('RegistryPath') # Hidden column for logic

        if ($addins.Count -gt 0) {
            foreach ($addin in $addins) {
                $row = $dataTable.NewRow()
                $row['Scope'] = $addin.Scope
                $row['Application'] = $addin.Application
                $row['Friendly Name'] = $addin.FriendlyName
                $row['Status (LoadBehavior)'] = $addin.LoadBehavior
                $row['ProgID'] = $addin.ProgID
                $row['RegistryPath'] = $addin.RegistryPath
                $dataTable.Rows.Add($row)
            }
            $script:LblAddinStatus.Text = "Scan complete. Found $($addins.Count) add-ins."
            $script:LblAddinStatus.ForeColor = 'Green'
        } else {
            $script:LblAddinStatus.Text = "No Litera Office add-ins found in the current user's registry."
            $script:LblAddinStatus.ForeColor = 'Orange'
        }
        $script:DgvAddins.DataSource = $dataTable
    } catch {
        $script:LblAddinStatus.Text = "An error occurred while scanning for add-ins: $($_.Exception.Message)"
        $script:LblAddinStatus.ForeColor = 'Red'
    } finally {
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

$script:DgvAddins.add_DataBindingComplete({
    param($sender, $e)
    if ($sender.Columns.Contains('RegistryPath')) { $sender.Columns['RegistryPath'].Visible = $false }
    if ($sender.Columns.Contains('Scope')) { $sender.Columns['Scope'].AutoSizeMode = 'AllCells' }
    if ($sender.Columns.Contains('Application')) { $sender.Columns['Application'].AutoSizeMode = 'AllCells' }
    if ($sender.Columns.Contains('Friendly Name')) { $sender.Columns['Friendly Name'].FillWeight = 50 }
    if ($sender.Columns.Contains('Status (LoadBehavior)')) { $sender.Columns['Status (LoadBehavior)'].FillWeight = 25 }
    if ($sender.Columns.Contains('ProgID')) { $sender.Columns['ProgID'].FillWeight = 25 }
})

$script:BtnRefreshAddins.Add_Click({ Populate-AddinGrid })

$script:BtnOpenAddinRegKey.Add_Click({
    if ($script:DgvAddins.SelectedRows.Count -eq 0) {
        $script:LblAddinStatus.Text = "Please select an add-in to view its registry key."; $script:LblAddinStatus.ForeColor = 'Orange'; return
    }
    if ($script:DgvAddins.SelectedRows.Count -gt 1) {
        $script:LblAddinStatus.Text = "Please select only one add-in to open its registry key."; $script:LblAddinStatus.ForeColor = 'Orange'; return
    }

    $selectedRow = $script:DgvAddins.SelectedRows[0]
    $psPath = $selectedRow.Cells['RegistryPath'].Value

    # Convert PowerShell path to a standard registry path that RegEdit understands
    $regPath = $psPath -replace '.*?Registry::', ''

    try {
        $regeditKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit"
        if (-not (Test-Path $regeditKeyPath)) { New-Item -Path $regeditKeyPath -Force | Out-Null }

        # Set the LastKey value so RegEdit opens to the correct location
        Set-ItemProperty -Path $regeditKeyPath -Name "LastKey" -Value $regPath -Force -ErrorAction Stop
        
        # Start RegEdit. The -m switch allows multiple instances.
        Start-Process "regedit.exe" -ArgumentList "-m" -ErrorAction Stop
        
        $script:LblAddinStatus.Text = "Opening Registry Editor to the selected key."; $script:LblAddinStatus.ForeColor = 'Blue'
    } catch {
        $errorMessage = "Failed to open Registry Editor: $($_.Exception.Message)"
        $script:LblAddinStatus.Text = $errorMessage; $script:LblAddinStatus.ForeColor = 'Red'
        Write-Log "REGEDIT ERROR: $errorMessage"
    }
})

$script:BtnAddToDontDisable.Add_Click({
    if ($script:DgvAddins.SelectedRows.Count -eq 0) {
        $script:LblAddinStatus.Text = "No add-ins selected."; $script:LblAddinStatus.ForeColor = 'Orange'; return
    }

    $confirmation = [System.Windows.Forms.MessageBox]::Show("This will add the selected add-in(s) to the 'Do Not Disable' resiliency list for the current user. This can prevent Office from automatically disabling them.`n`nAre you sure you want to continue?", "Confirm Action", "YesNo", "Question")
    if ($confirmation -ne 'Yes') {
        $script:LblAddinStatus.Text = "Operation cancelled."; $script:LblAddinStatus.ForeColor = 'Gray'; return
    }

    $errors = 0
    $successCount = 0

    # Find the highest installed Office version number (e.g., 16.0) from the current user's registry hive
    $officeVersion = Get-ChildItem -Path "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d{1,2}\.\d$' } |
        Sort-Object -Property @{Expression = { [version]$_.PSChildName } } -Descending |
        Select-Object -First 1 |
        Select-Object -ExpandProperty PSChildName

    if (-not $officeVersion) {
        $script:LblAddinStatus.Text = "Could not determine Office version from registry."; $script:LblAddinStatus.ForeColor = 'Red'
        Write-Log "ADD-IN MGMT ERROR: Could not determine Office version from HKCU:\Software\Microsoft\Office"
        return
    }

    foreach ($row in $script:DgvAddins.SelectedRows) {
        try {
            $progId = $row.Cells['ProgID'].Value
            $appName = $row.Cells['Application'].Value

            $regPath = "HKCU:\Software\Microsoft\Office\$officeVersion\$appName\Resiliency\DoNotDisableAddinList"

            # Ensure the path exists
            if (-not (Test-Path $regPath)) {
                New-Item -Path $regPath -Force -ErrorAction Stop | Out-Null
                Write-Log "ADD-IN MGMT: Created resiliency key: $regPath"
            }

            # Set the DWORD value
            Set-ItemProperty -Path $regPath -Name $progId -Value 1 -Type DWord -Force -ErrorAction Stop
            $successCount++
            Write-Log "ADD-IN MGMT: Added '$progId' to DoNotDisableAddinList for $appName."
        } catch {
            $errors++
            $script:LblAddinStatus.Text = "Error adding '$($row.Cells['Friendly Name'].Value)' to list: $($_.Exception.Message)"; $script:LblAddinStatus.ForeColor = 'Red'
            Write-Log "ADD-IN MGMT ERROR: Failed to add '$($progId)' to DoNotDisableAddinList. $($_.Exception.Message)"
        }
    }

    if ($errors -eq 0) {
        $script:LblAddinStatus.Text = "Successfully added $successCount add-in(s) to the 'Do Not Disable' list."; $script:LblAddinStatus.ForeColor = 'Green'
    }
})

$script:BtnClearResilience.Add_Click({
    if ($script:DgvAddins.SelectedRows.Count -eq 0) {
        $script:LblAddinStatus.Text = "No add-ins selected."; $script:LblAddinStatus.ForeColor = 'Orange'; return
    }

    $confirmation = [System.Windows.Forms.MessageBox]::Show("This will attempt to clear resiliency keys for the selected add-in(s). This can resolve issues where Office automatically disables an add-in.`n`nSpecifically, it will:`n • Remove entries from the 'Do Not Disable' list.`n • Clear the entire 'Disabled Items' list for the relevant Office application(s).`n`nAre you sure you want to continue?", "Confirm Clear Resilience", "YesNo", "Question")
    if ($confirmation -ne 'Yes') {
        $script:LblAddinStatus.Text = "Operation cancelled."; $script:LblAddinStatus.ForeColor = 'Gray'; return
    }

    $errors = 0
    $processedApps = [System.Collections.Generic.HashSet[string]]::new()

    # Find the highest installed Office version number (e.g., 16.0) from the current user's registry hive
    $officeVersion = Get-ChildItem -Path "HKCU:\Software\Microsoft\Office" -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d{1,2}\.\d$' } |
        Sort-Object -Property @{Expression = { [version]$_.PSChildName } } -Descending |
        Select-Object -First 1 |
        Select-Object -ExpandProperty PSChildName

    if (-not $officeVersion) {
        $script:LblAddinStatus.Text = "Could not determine Office version from registry."; $script:LblAddinStatus.ForeColor = 'Red'
        Write-Log "ADD-IN MGMT ERROR: Could not determine Office version from HKCU:\Software\Microsoft\Office"
        return
    }

    foreach ($row in $script:DgvAddins.SelectedRows) {
        try {
            $progId = $row.Cells['ProgID'].Value
            $appName = $row.Cells['Application'].Value

            # 1. Clear from DoNotDisableAddinList
            $doNotDisablePath = "HKCU:\Software\Microsoft\Office\$officeVersion\$appName\Resiliency\DoNotDisableAddinList"
            if (Test-Path $doNotDisablePath) {
                if (Get-ItemProperty -Path $doNotDisablePath -Name $progId -ErrorAction SilentlyContinue) {
                    Remove-ItemProperty -Path $doNotDisablePath -Name $progId -Force -ErrorAction Stop
                    Write-Log "ADD-IN MGMT: Removed '$progId' from DoNotDisableAddinList for $appName."
                }
            }

            # 2. Clear the entire DisabledItems key for the application
            if (-not $processedApps.Contains($appName)) {
                $disabledItemsPath = "HKCU:\Software\Microsoft\Office\$officeVersion\$appName\Resiliency\DisabledItems"
                if (Test-Path $disabledItemsPath) {
                    Remove-Item -Path $disabledItemsPath -Recurse -Force -ErrorAction Stop
                    Write-Log "ADD-IN MGMT: Cleared DisabledItems list for $appName."
                }
                $processedApps.Add($appName) | Out-Null
            }
        } catch {
            $errors++
            $script:LblAddinStatus.Text = "Error clearing resiliency for '$($row.Cells['Friendly Name'].Value)': $($_.Exception.Message)"; $script:LblAddinStatus.ForeColor = 'Red'
            Write-Log "ADD-IN MGMT ERROR: Failed to clear resiliency for '$($progId)'. $($_.Exception.Message)"
        }
    }

    if ($errors -eq 0) {
        $script:LblAddinStatus.Text = "Successfully cleared resiliency keys. Please restart the Office application(s)."; $script:LblAddinStatus.ForeColor = 'Green'
    }
})
$script:JobTimer = New-Object System.Windows.Forms.Timer; $script:JobTimer.Interval = 200
$script:CountdownTimer = New-Object System.Windows.Forms.Timer; $script:CountdownTimer.Interval = 1000
$script:currentJob = $null
$script:jobCompletionAction = $null
$script:countdownValue = 0
$script:countdownLineStart = 0

<#
.SYNOPSIS
    Defines global event handlers for asynchronous operations.
.DESCRIPTION
    This section sets up timers and global variables (`$JobTimer`, `$CountdownTimer`, `$script:currentJob`) to manage background jobs started from various tabs, ensuring the UI remains responsive.
#>
    $script:ChkEnableLicenseLogging.Add_CheckedChanged({
        $script:settings.EnableLicenseCheckLogging = $script:ChkEnableLicenseLogging.Checked
        Save-Settings
        $script:LblLicenseStatus.Text = if ($script:ChkEnableLicenseLogging.Checked) { "License check logging is enabled." } else { "License check logging is disabled." }
        $script:LblLicenseStatus.ForeColor = 'Gray'
    })

<#
    The main timer tick event handler for processing completed background jobs.
#>
$script:JobTimer.Add_Tick({
    if ($script:currentJob -and $script:currentJob.State -in ('Completed', 'Failed', 'Stopped')) {
        $script:JobTimer.Stop()

        $completedJob = $script:currentJob
        $actionToRun = $script:jobCompletionAction
        $script:currentJob = $null
        $script:jobCompletionAction = $null

        $jobResult = if ($completedJob.State -eq 'Completed') {
            try {
                Receive-Job $completedJob
            } catch {
                $completedJob.ChildJobs[0].Error
            }
        } else {
            $completedJob.ChildJobs[0].Error
        }

        # Ensure we always have a result, even if it's an error record
        if ($null -eq $jobResult) { $jobResult = $completedJob.ChildJobs[0].Error }

        if ($actionToRun) {
            & $actionToRun $jobResult
        }

        Remove-Job $completedJob -Force
    }
})

<#
    The timer tick event handler for the countdown displayed during license checks.
#>
$script:CountdownTimer.Add_Tick({
    $script:countdownValue--
    if ($script:countdownValue -ge 0) {
        $script:RtbLicenseInfo.Select($script:countdownLineStart, $script:RtbLicenseInfo.Text.Length - $script:countdownLineStart)
        $script:RtbLicenseInfo.SelectionFont = $script:RtbLicenseInfo.Font
        $script:RtbLicenseInfo.SelectionColor = 'Black'
        $script:RtbLicenseInfo.SelectedText = "  (o) Waiting for connection: $($script:countdownValue) seconds remaining...`r`n"
    } else {
        $script:CountdownTimer.Stop()
    }
})

<#
    Event handler for the "Load Litera Programs" button click. Starts a job to find installed programs.
#>
$script:BtnLoadPrograms.Add_Click({
    if ($script:currentJob) { [System.Windows.Forms.MessageBox]::Show("Another operation is already in progress.", "Busy", "OK", "Warning"); return }

    $script:BtnLoadPrograms.Enabled = $false
    $script:ProgressUninstall.Style = 'Marquee'; $script:ProgressUninstall.Visible = $true

    $script:jobCompletionAction = {
        param($jobResult)
        $script:ProgressUninstall.Visible = $false
        $script:BtnLoadPrograms.Enabled = $true

        if ($jobResult -is [System.Management.Automation.ErrorRecord]) {
            [System.Windows.Forms.MessageBox]::Show("Failed to load programs: $($jobResult.Exception.Message)", "Error", "OK", "Error")
            return
        }

        $script:ComboPrograms.Items.Clear()
        $script:installedApps.Clear()
        foreach ($app in $jobResult) {
            $ComboPrograms.Items.Add($app.DisplayText)
            $script:installedApps[$app.DisplayText] = $app.AppObject
        }

        if ($ComboPrograms.Items.Count -gt 0) {
            $script:ComboPrograms.SelectedIndex = 0
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("No Litera products found.", "Information", "OK", "Information")
        }
    }

    $jobArgs = @{ ShowAll = $script:settings.ShowAllPrograms }
    $script:currentJob = Start-Job -ArgumentList $jobArgs -ScriptBlock {
        param($arguments)
        $registryPaths = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        
        $programs = Get-ItemProperty $registryPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }

        if (-not $arguments.ShowAll) {
            $programs = $programs | Where-Object { $_.Publisher -like '*Litera*' }
        }

        $programs | Sort-Object DisplayName | ForEach-Object {
                $displayText = $_.DisplayName
                if ($_.DisplayVersion) { $displayText += " (v$($_.DisplayVersion))" }
                [PSCustomObject]@{ DisplayText = $displayText; AppObject = $_ }
            }
    }
    $script:JobTimer.Start()
})

<#
    Event handler for the "Uninstall" button click. Starts a job to run the uninstaller.
#>
$script:BtnUninstall.Add_Click({
    if ($script:currentJob) { [System.Windows.Forms.MessageBox]::Show("Another operation is already in progress.", "Busy", "OK", "Warning"); return }

    $selectedApp = $script:ComboPrograms.SelectedItem
    if (-not $selectedApp) { [System.Windows.Forms.MessageBox]::Show("Please select a program to uninstall.", "No Selection", "OK", "Warning"); return }
    
    $app = $script:installedApps[$selectedApp]
    if (-not $app -or [string]::IsNullOrWhiteSpace($app.UninstallString)) {
        $script:LblUninstallResult.Text = "Uninstall string not found!"; $script:LblUninstallResult.ForeColor = 'Red'; $script:LblUninstallResult.Visible = $true
        Write-Log "Uninstall FAILED for '$selectedApp'. Application uninstall string not found or is empty in registry."
        return
    }

    $script:BtnUninstall.Enabled = $false; $script:BtnLoadPrograms.Enabled = $false
    $script:ProgressUninstall.Visible = $true
    Set-ProgressBarState -ProgressBar $script:ProgressUninstall -State 'Normal'
    $script:LblUninstallResult.Visible = $true; $script:LblUninstallResult.Text = "Uninstalling..."; $script:LblUninstallResult.ForeColor = 'Blue'

    $script:jobCompletionAction = {
        param($jobResult)
        $script:BtnUninstall.Enabled = $true; $script:BtnLoadPrograms.Enabled = $true

        if ($jobResult -is [System.Management.Automation.PSCustomObject]) {
            if ($jobResult.ExitCode -eq 0 -or $jobResult.ExitCode -eq 3010) {
                Set-ProgressBarState -ProgressBar $script:ProgressUninstall -State 'Success'
                $script:LblUninstallResult.Text = "Success! Refresh list to see changes."; $script:LblUninstallResult.ForeColor = 'Green'
                Write-Log "Uninstall successful for '$selectedApp'. Exit Code: $($jobResult.ExitCode)"
            } else {
                Set-ProgressBarState -ProgressBar $script:ProgressUninstall -State 'Error'
                $script:LblUninstallResult.Text = "Failed! (Code: $($jobResult.ExitCode))"; $script:LblUninstallResult.ForeColor = 'Red'
                Write-Log "Uninstall FAILED for '$selectedApp'. Exit Code: $($jobResult.ExitCode)"
            }
        } else {
            Set-ProgressBarState -ProgressBar $script:ProgressUninstall -State 'Error'
            $errorMessage = if ($jobResult -is [System.Collections.IList] -and $jobResult.Count -gt 0) { $jobResult[0].Exception.Message } else { "Unknown error" }
            $script:LblUninstallResult.Text = "Error starting process!"; $script:LblUninstallResult.ForeColor = 'Red'
            Write-Log "Uninstall FAILED for '$selectedApp'. Error executing process. Message: $errorMessage"
        }
    }

    $jobArgs = @{
        UninstallString = $app.UninstallString
        UserArguments   = $script:TxtUninstallArgs.Text
        Silent          = $script:ChkInstallSilent.Checked
        BasicUi         = $script:ChkInstallBasicUi.Checked
        RebootSuppress  = $script:ChkRebootSuppress.Checked
        LogRoot         = $script:logRoot
        LogEnabled      = $script:logEnabled
        SelectedApp     = $selectedApp
    }

    $script:currentJob = Start-Job -ArgumentList $jobArgs -ScriptBlock {
        param($jobParams)
        function Write-Log {
            param([string]$message)
        }

        try {
            $executable = $null
            $finalArguments = $null
            $guidMatch = [regex]::Match($jobParams.UninstallString, '(\{([a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12})\})')

            if ($guidMatch.Success -and ($jobParams.Silent -or $jobParams.BasicUi)) {
                $executable = "msiexec.exe"
                $productCode = $guidMatch.Groups[1].Value
                
                $argList = New-Object System.Collections.Generic.List[string]
                $argList.Add("/X$productCode")
                
                if ($jobParams.Silent) { $argList.Add("/qn") }
                if ($jobParams.BasicUi) { $argList.Add("/qb") }
                if (-not [string]::IsNullOrWhiteSpace($jobParams.UserArguments)) {
                    $argList.Add($jobParams.UserArguments)
                }
                $finalArguments = ($argList -join " ").Trim()
            } else {
                $commandMatch = [regex]::Match($jobParams.UninstallString, '^"([^"]+)"|^(\S+)')
                $executable = $commandMatch.Value.Replace('"','')

                if (-not $commandMatch.Success -or [string]::IsNullOrWhiteSpace($executable)) {
                    throw "Could not parse executable from uninstall string: $($jobParams.UninstallString)"
                }

                $baseArguments = $jobParams.UninstallString.Substring($commandMatch.Length).Trim()
                $finalArguments = "$baseArguments $($jobParams.UserArguments)".Trim()
            }
            
            if ($jobParams.LogEnabled -and $executable -like '*msiexec.exe*') {
                $uninstallLogPath = Join-Path $jobParams.LogRoot "uninstall_logs"
                if (-not (Test-Path $uninstallLogPath)) { New-Item -Path $uninstallLogPath -ItemType Directory -Force | Out-Null }
                $safeAppName = $jobParams.SelectedApp -replace '[^a-zA-Z0-9]', '_'
                $logFileName = "Uninstall_$($safeAppName)_$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
                $logFilePath = Join-Path $uninstallLogPath $logFileName
                $finalArguments += " /L*v `"$logFilePath`""
            }
            
            Write-Log "Starting uninstall for '$($jobParams.SelectedApp)'. Executable: '$executable', Arguments: '$finalArguments'"
            $uninstallProcess = Start-Process $executable -ArgumentList $finalArguments -Wait -PassThru -ErrorAction Stop
            return [PSCustomObject]@{ ExitCode = $uninstallProcess.ExitCode }
        }
        catch {
            Write-Log "ERROR during uninstall job setup: $($_.Exception.Message)"
            throw
        }
    }
    $script:JobTimer.Start()
})

<#
    Event handlers to prevent mutually exclusive "Silent" and "Basic UI" checkboxes from being selected simultaneously.
#>
$script:ChkInstallSilent.Add_CheckedChanged({
    if ($script:ChkInstallSilent.Checked -and $script:ChkInstallBasicUi.Checked) {
        [System.Windows.Forms.MessageBox]::Show("Silent and Basic UI are mutually exclusive. Please select only one.", "Installation Option Conflict", "OK", "Error") | Out-Null
        $script:ChkInstallBasicUi.Checked = $false
    }
})

<#
    Event handler for the "Browse" button to select an MSI installer.
#>
$script:BtnInstall.Add_Click({
    if ($script:currentJob) { [System.Windows.Forms.MessageBox]::Show("Another operation is already in progress.", "Busy", "OK", "Warning"); return }

    $installerPath = $script:TxtInstallerPath.Text
    if (-not (Test-Path $installerPath)) { [System.Windows.Forms.MessageBox]::Show("Installer file not found.", "File Not Found", "OK", "Error"); return }

    $script:BtnInstall.Enabled = $false; $script:BtnBrowseInstaller.Enabled = $false
    $script:ProgressInstall.Visible = $true
    Set-ProgressBarState -ProgressBar $script:ProgressInstall -State 'Normal'
    $script:LblInstallResult.Visible = $true; $script:LblInstallResult.Text = "Installing..."; $script:LblInstallResult.ForeColor = 'Blue'

    $script:jobCompletionAction = {
        param($jobResult)
        $script:BtnInstall.Enabled = $true; $script:BtnBrowseInstaller.Enabled = $true

        if ($jobResult -is [System.Management.Automation.PSCustomObject]) {
            if ($jobResult.ExitCode -eq 0 -or $jobResult.ExitCode -eq 3010) {
                Set-ProgressBarState -ProgressBar $script:ProgressInstall -State 'Success'
                $script:LblInstallResult.Text = "Success!"; $script:LblInstallResult.ForeColor = 'Green'
                Write-Log "Install successful for '$installerPath'. Exit Code: $($jobResult.ExitCode)"
                $script:BtnLoadPrograms.PerformClick()
            } else {
                Set-ProgressBarState -ProgressBar $script:ProgressInstall -State 'Error'
                $script:LblInstallResult.Text = "Failed! (Code: $($jobResult.ExitCode))"; $script:LblInstallResult.ForeColor = 'Red'
                Write-Log "Install FAILED for '$installerPath'. Exit Code: $($jobResult.ExitCode)"
            }
        } else {
            Set-ProgressBarState -ProgressBar $script:ProgressInstall -State 'Error'
            $errorMessage = if ($jobResult -is [System.Collections.IList] -and $jobResult.Count -gt 0) { $jobResult[0].Exception.Message } else { "Unknown error" }
            $script:LblInstallResult.Text = "Error starting process!"; $script:LblInstallResult.ForeColor = 'Red'
            Write-Log "Install FAILED for '$installerPath'. Error executing process. Message: $errorMessage"
        }
    }

    $jobArgs = @{
        InstallerPath = $installerPath
        UserArguments = $script:TxtInstallArgs.Text
        Silent = $script:ChkInstallSilent.Checked
        BasicUi = $script:ChkInstallBasicUi.Checked
        RebootSuppress = $script:ChkRebootSuppress.Checked
        DisableWord = $script:ChkDisableWordAddin.Checked
        DisableExcel = $script:ChkDisableExcelAddin.Checked
        DisableOutlook = $script:ChkDisableOutlookAddin.Checked
        DisablePpt = $script:ChkDisablePptAddin.Checked
        NoDesktopWord = $script:ChkNoDesktopWord.Checked
        NoDesktopPpt = $script:ChkNoDesktopPpt.Checked
        NoDesktopExcel = $script:ChkNoDesktopExcel.Checked
        NoStartMenuWord = $script:ChkNoStartMenuWord.Checked
        NoStartMenuPpt = $script:ChkNoStartMenuPpt.Checked
        NoStartMenuExcel = $script:ChkNoStartMenuExcel.Checked
        LogEnabled = $script:logEnabled
        LogRoot = $script:logRoot
    }

    $script:currentJob = Start-Job -ArgumentList $jobArgs -ScriptBlock {
        param($jobParams)
        function Write-Log {
            param([string]$message)
            # This function is a placeholder inside the job. 
            # If real-time logging from the job to the main script's log file is needed,
            # it would require more complex inter-process communication.
            # For now, the final command is logged before the job starts.
        }

        $msiArguments = New-Object System.Collections.Generic.List[string]
        if ($jobParams.Silent) { $msiArguments.Add("/qn") }
        if ($jobParams.BasicUi) { $msiArguments.Add("/qb") }
        if ($jobParams.DisableWord) { $msiArguments.Add("WORDADDIN=0") }
        if ($jobParams.DisableExcel) { $msiArguments.Add("EXCELADDIN=0") }
        if ($jobParams.DisableOutlook) { $msiArguments.Add("OUTLOOKADDIN=0") }
        if ($jobParams.DisablePpt) { $msiArguments.Add("PPTADDIN=0") }
        if ($jobParams.NoDesktopWord) { $msiArguments.Add("DESKTOP_SHORTCUT_LC_FOR_WORD=0") }
        if ($jobParams.NoDesktopPpt) { $msiArguments.Add("DESKTOP_SHORTCUT_LC_FOR_PPT=0") }
        if ($jobParams.NoDesktopExcel) { $msiArguments.Add("DESKTOP_SHORTCUT_LC_FOR_EXCEL=0") }
        if ($jobParams.NoStartMenuWord) { $msiArguments.Add("START_MENU_SHORTCUT_LC_FOR_WORD=0") }
        if ($jobParams.NoStartMenuPpt) { $msiArguments.Add("START_MENU_SHORTCUT_LC_FOR_PPT=0") }
        if ($jobParams.NoStartMenuExcel) { $msiArguments.Add("START_MENU_SHORTCUT_LC_FOR_EXCEL=0") }
        if ($jobParams.RebootSuppress) {
            $msiArguments.Add("REBOOT=ReallySuppress")
        }

        $executable = "msiexec.exe"
        $finalArguments = ""
        $isMsiInstall = $false

        if ($jobParams.InstallerPath.EndsWith(".msi")) {
            $isMsiInstall = $true
            $msiArgsString = $msiArguments -join " "
            $finalArguments = "/i `"$($jobParams.InstallerPath)`" $msiArgsString $($jobParams.UserArguments)".Trim()
        } else { # Assume .exe
            $executable = $jobParams.InstallerPath
            $finalArguments = $jobParams.UserArguments
            # If there are any MSI-specific arguments, assume the .exe is a wrapper and pass them via /v
            if ($msiArguments.Count -gt 0) {
                $isMsiInstall = $true # Treat as MSI for logging purposes
                $msiArgsString = ($msiArguments -join " ").Replace('"', '\"')
                # Prepend the MSI arguments, wrapped for the .exe bootstrapper
                $finalArguments = "/v`"$msiArgsString`" " + $finalArguments
            }
            $finalArguments = $finalArguments.Trim()
        }

        # Add logging only if it's an MSI install or an EXE that we assume is wrapping an MSI
        if ($jobParams.LogEnabled -and $isMsiInstall) {
            $installLogPath = Join-Path $jobParams.LogRoot "install_logs"
            if (-not (Test-Path $installLogPath)) { New-Item -Path $installLogPath -ItemType Directory -Force | Out-Null }
            $installerName = (Split-Path $jobParams.InstallerPath -Leaf) -replace '[^a-zA-Z0-9.]', '_'
            $logFileName = "Install_Log_$($installerName)_$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
            $logFilePath = Join-Path $installLogPath $logFileName
            # If it's an exe, logging needs to be passed inside the /v switch
            if ($jobParams.InstallerPath.EndsWith(".exe")) {
                $finalArguments = $finalArguments.Insert($finalArguments.IndexOf('"')+1, "/L*v `"$logFilePath`" ")
            } else {
                $finalArguments += " /L*v `"$logFilePath`""
            }
        }
        
        Write-Log "Starting install for '$($jobParams.InstallerPath)'. Command: $executable $finalArguments"

        # For .exe installers, they often act as wrappers. We need to find and wait for the child msiexec process.
        if ($jobParams.InstallerPath.EndsWith(".exe", [System.StringComparison]::OrdinalIgnoreCase)) {
            # Get a baseline of current msiexec processes
            $baselineMsiProcs = Get-Process -Name "msiexec" -ErrorAction SilentlyContinue

            # Start the .exe installer but do NOT wait for it
            if (-not [string]::IsNullOrWhiteSpace($finalArguments)) {
                Start-Process $executable -ArgumentList $finalArguments -ErrorAction Stop
            } else {
                Start-Process $executable -ErrorAction Stop
            }

            # Poll for the new msiexec process spawned by the .exe
            $actualInstallerProcess = $null
            foreach ($i in 1..10) { # Poll for up to 10 seconds
                $currentMsiProcs = Get-Process -Name "msiexec" -ErrorAction SilentlyContinue
                $actualInstallerProcess = Compare-Object -ReferenceObject $baselineMsiProcs -DifferenceObject $currentMsiProcs -PassThru
                if ($actualInstallerProcess) { break }
                Start-Sleep -Seconds 1
            }

            if ($actualInstallerProcess) {
                Write-Log "Wrapper .exe detected. Waiting for actual installer process (PID: $($actualInstallerProcess.Id))."
                $actualInstallerProcess | Wait-Process
                return [PSCustomObject]@{ ExitCode = $actualInstallerProcess.ExitCode }
            } else {
                # Fallback: If no new msiexec process was found, assume it's a self-contained .exe and wait for it directly.
                # This might still give a false positive, but it's the best we can do.
                Write-Log "No child msiexec process found. Waiting for the main .exe to complete."
                $exeProcess = Get-Process -Name ($executable | Split-Path -LeafBase) | Sort-Object StartTime -Descending | Select-Object -First 1
                if ($exeProcess) { $exeProcess | Wait-Process }
                return [PSCustomObject]@{ ExitCode = $exeProcess.ExitCode }
            }
        } else { # Standard handling for .msi files
            $installProcess = if (-not [string]::IsNullOrWhiteSpace($finalArguments)) {
                Start-Process $executable -ArgumentList $finalArguments -Wait -PassThru -ErrorAction Stop
            } else {
                Start-Process $executable -Wait -PassThru -ErrorAction Stop
            }
            return [PSCustomObject]@{ ExitCode = $installProcess.ExitCode }
        }
    }
    $script:JobTimer.Start()
})

<#
    Event handler for the "Open Log Folder" button on the Install/Uninstall tab.
#>
$script:BtnOpenLogFolder.Add_Click({
    if (-not (Test-Path $script:logRoot)) {
        New-Item -Path $script:logRoot -ItemType Directory -Force | Out-Null
    }
    Invoke-Item $script:logRoot
})

$script:ChkInstallSilent.Add_CheckedChanged({
    if ($script:ChkInstallSilent.Checked -and $script:ChkInstallBasicUi.Checked) {
        [System.Windows.Forms.MessageBox]::Show("Silent and Basic UI are mutually exclusive. Please select only one.", "Installation Option Conflict", "OK", "Error") | Out-Null
        $script:ChkInstallSilent.Checked = $false
    }
})

$script:ChkInstallBasicUi.Add_CheckedChanged({
    if ($script:ChkInstallBasicUi.Checked -and $script:ChkInstallSilent.Checked) {
        [System.Windows.Forms.MessageBox]::Show("Silent and Basic UI are mutually exclusive. Please select only one.", "Installation Option Conflict", "OK", "Error") | Out-Null
        $script:ChkInstallBasicUi.Checked = $false
    }
})

$script:BtnBrowseInstaller.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    if ($script:settings.AllowExeInstallation) {
        $fileDialog.Filter = "Installers (*.msi, *.exe)|*.msi;*.exe|All Files (*.*)|*.*"
    } else {
        $fileDialog.Filter = "MSI Installer (*.msi)|*.msi|All Files (*.*)|*.*"
    }
    $fileDialog.Title = "Select Installer File"
    if ($fileDialog.ShowDialog() -eq 'OK') { $script:TxtInstallerPath.Text = $fileDialog.FileName }
})

$script:ChkLogging.Add_CheckedChanged({
    if ($script:ChkLogging.Checked) {
        try {
            if (-not (Test-Path $script:logRoot)) {
                New-Item -Path $script:logRoot -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            $tempFile = Join-Path $script:logRoot ([System.IO.Path]::GetRandomFileName())
            Set-Content -Path $tempFile -Value "permission_check" -ErrorAction Stop
            Remove-Item -Path $tempFile -Force -ErrorAction Stop

            $script:LblLoggingStatus.Text = "Logging enabled for installs/uninstalls."
            $script:LblLoggingStatus.ForeColor = 'Green'
            $script:logEnabled = $true
            $script:settings.EnableInstallLogging = $true
            Save-Settings
        }
        catch {
            $script:LblLoggingStatus.Text = "Error: Failed to access log directory. Check permissions."
            $script:LblLoggingStatus.ForeColor = 'Red'
            $script:logEnabled = $false
            $script:ChkLogging.Checked = $false
            Write-Warning "Failed to enable logging: $($_.Exception.Message)"
        }
    }
    else {
        $script:LblLoggingStatus.Text = "Logging disabled."
        $script:LblLoggingStatus.ForeColor = 'Gray'
        $script:logEnabled = $false
        $script:settings.EnableInstallLogging = $false
        Save-Settings
    }
})

<#
.SYNOPSIS
    Defines functions for loading, saving, and applying application settings.
.DESCRIPTION
    This section contains the core logic for settings persistence. Settings are stored in a JSON file in the script's directory.
#>

<#
.SYNOPSIS
    Loads application settings from a JSON config file.
.DESCRIPTION
    Reads `LC_Tool_Settings.config` from the script's directory. If the file doesn't exist, it creates one with default settings. It also handles backward compatibility by adding any missing settings to the loaded object.
#>
function Load-Settings {
    if (-not (Test-Path $settingsFile)) {
        Write-Log "Settings file not found. Creating default settings at '$settingsFile'."
        $script:settings = @{
            BackupRoot           = $defaultBackupRoot
            LogRoot              = $defaultLogRoot
            EnableInstallLogging = $false
            EnableLicenseCheckLogging = $false
            AllowExeInstallation = $false
            ShowAllPrograms      = $false
            TabVisibility        = @{
                ShowCleanupTab          = $true
                ShowBackupRestoreTab    = $true
                ShowSysReqTab           = $true
                ShowInstallUninstallTab = $true
                ShowLicenseTab          = $true
                ShowCompareCustomTab    = $true
                ShowAddinMgmtTab        = $true
                ShowLogViewerTab        = $true
            }
        }
        Save-Settings
    }
    else {
        try {
            $script:settings = Get-Content -Path $settingsFile -Raw | ConvertFrom-Json -ErrorAction Stop

            # Convert PSCustomObject to Hashtable for easier manipulation
            if ($script:settings.TabVisibility -is [System.Management.Automation.PSCustomObject]) {
                $tabVisibilityHashtable = @{}
                $script:settings.TabVisibility.PSObject.Properties | ForEach-Object { $tabVisibilityHashtable[$_.Name] = $_.Value }
                $script:settings.TabVisibility = $tabVisibilityHashtable
            }

            # For backward compatibility, add any missing settings with default values
            $defaultSettings = @{
                BackupRoot           = $defaultBackupRoot
                LogRoot              = $defaultLogRoot 
                EnableInstallLogging = $false
                EnableLicenseCheckLogging = $false
                AllowExeInstallation = $false
                ShowAllPrograms      = $false
                TabVisibility        = @{ ShowCleanupTab=$true; ShowBackupRestoreTab=$true; ShowSysReqTab=$true; ShowInstallUninstallTab=$true; ShowLicenseTab=$true; ShowCompareCustomTab=$true; ShowAddinMgmtTab=$true; ShowLogViewerTab=$true }
            }

            foreach ($key in $defaultSettings.Keys) {
                if (-not $script:settings.PSObject.Properties.Name.Contains($key)) { Add-Member -InputObject $script:settings -MemberType NoteProperty -Name $key -Value $defaultSettings.$key }
            }
            foreach ($key in $defaultSettings.TabVisibility.Keys) { if (-not $script:settings.TabVisibility.ContainsKey($key)) { $script:settings.TabVisibility[$key] = $defaultSettings.TabVisibility[$key] } }
            if (-not $script:settings.PSObject.Properties.Name.Contains('AllowExeInstallation')) { Add-Member -InputObject $script:settings -MemberType NoteProperty -Name 'AllowExeInstallation' -Value $false }
            if (-not $script:settings.PSObject.Properties.Name.Contains('ShowAllPrograms')) { Add-Member -InputObject $script:settings -MemberType NoteProperty -Name 'ShowAllPrograms' -Value $false }

        }
        catch {
            Write-Warning "Failed to load or parse '$settingsFile'. Using default settings. Error: $($_.Exception.Message)"
            $script:settings = @{
                BackupRoot           = $defaultBackupRoot
                LogRoot              = $defaultLogRoot
                EnableInstallLogging = $false
                EnableLicenseCheckLogging = $false
                AllowExeInstallation = $false
                ShowAllPrograms      = $false
                TabVisibility        = @{
                    ShowCleanupTab          = $true
                    ShowBackupRestoreTab    = $true
                    ShowSysReqTab           = $true
                    ShowInstallUninstallTab = $true
                    ShowLicenseTab          = $true
                    ShowCompareCustomTab    = $true
                    ShowAddinMgmtTab        = $true
                    ShowLogViewerTab        = $true
                }
            }
        }
    }
}

<#
.SYNOPSIS
    Saves the current application settings to the JSON config file.
.DESCRIPTION
    Converts the `$script:settings` object to JSON and overwrites the `LC_Tool_Settings.config` file.
#>
function Save-Settings {
    try {
        $settingsDir = Split-Path -Path $settingsFile -Parent
        if (-not (Test-Path $settingsDir)) {
            New-Item -ItemType Directory -Path $settingsDir -Force | Out-Null
        }
        $script:settings | ConvertTo-Json -Depth 5 | Out-File -FilePath $settingsFile -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to save settings to '$settingsFile'.`nError: $($_.Exception.Message)", "Save Error", "OK", "Error")
    }
}

<#
.SYNOPSIS
    Applies the loaded settings to the application's current state.
.DESCRIPTION
    Updates global variables and UI elements (like tab visibility and checkbox states) based on the values in the `$script:settings` object.
#>
function Apply-Settings {
    $script:backupRoot = $script:settings.BackupRoot
    $script:logRoot = $script:settings.LogRoot
    $script:logEnabled = $script:settings.EnableInstallLogging
    $script:ChkLogging.Checked = $script:logEnabled
    $script:ChkEnableLicenseLogging.Checked = $script:settings.EnableLicenseCheckLogging

    # Update Install/Uninstall tab UI based on settings
    if ($script:LblInstaller) {
        $script:LblInstaller.Text = if ($script:settings.AllowExeInstallation) { "Select installer (.msi or .exe):" } else { "Select installer (.msi):" }
    }
    if ($script:BtnLoadPrograms) {
        $script:BtnLoadPrograms.Text = if ($script:settings.ShowAllPrograms) { "Load All Programs" } else { "Load Litera Programs" }
    }

    $script:TabControl.TabPages.Clear()
    if ($script:settings.TabVisibility.ShowCleanupTab) { $script:TabControl.TabPages.Add($script:TabMap.ShowCleanupTab) }
    if ($script:settings.TabVisibility.ShowBackupRestoreTab) { $script:TabControl.TabPages.Add($script:TabMap.ShowBackupRestoreTab) }
    if ($script:settings.TabVisibility.ShowSysReqTab) { $script:TabControl.TabPages.Add($script:TabMap.ShowSysReqTab) }
    if ($script:settings.TabVisibility.ShowInstallUninstallTab) { $script:TabControl.TabPages.Add($script:TabMap.ShowInstallUninstallTab) }
    if ($script:settings.TabVisibility.ShowLicenseTab) { $script:TabControl.TabPages.Add($script:TabMap.ShowLicenseTab) }
    if ($script:settings.TabVisibility.ShowCompareCustomTab) { $script:TabControl.TabPages.Add($script:TabMap.ShowCompareCustomTab) }
    if ($script:settings.TabVisibility.ShowAddinMgmtTab) { $script:TabControl.TabPages.Add($script:TabMap.ShowAddinMgmtTab) }
    if ($script:settings.TabVisibility.ShowLogViewerTab) { $script:TabControl.TabPages.Add($script:TabMap.ShowLogViewerTab) }
}

<#
    A hashtable mapping setting names to their corresponding TabPage objects for easy access.
#>
$script:TabMap = @{
    ShowCleanupTab          = $TabCleanup
    ShowBackupRestoreTab    = $TabBackupRestore
    ShowSysReqTab           = $TabSysReq
    ShowInstallUninstallTab = $TabInstallUninstall
    ShowLicenseTab          = $TabLicense
    ShowCompareCustomTab    = $TabCompareCustom
    ShowAddinMgmtTab        = $TabAddinMgmt
    ShowLogViewerTab        = $TabLogViewer
}

Load-Settings
Apply-Settings

<#
    Event handler for the main form's `Shown` event to populate initial data.
#>
$script:Form.Add_Shown({
    Populate-AddinGrid
    Populate-CleanupLists
    Populate-BackupRestoreLists
})

<#
    Event handler for the main form's `FormClosing` event to perform cleanup.
#>
$script:Form.Add_FormClosing({
    Get-Job | Remove-Job -Force
    $script:JobTimer.Stop()
    $script:CountdownTimer.Stop()
    $script:logPollTimer.Stop()
    if ($script:TabLiteraLogging) { $script:TabLiteraLogging.Dispose() }
})

<#
.SYNOPSIS
    Prompts the user to restart the tool as an administrator if required.
.DESCRIPTION
    Checks if an operation requires admin rights and, if so, shows a confirmation dialog. If the user agrees, it restarts the script with elevated privileges.
.RETURNS
    $true if the restart was initiated, $false otherwise.
#>
function Invoke-RestartAsAdmin {
    $ans = [System.Windows.Forms.MessageBox]::Show("This action requires administrator privileges. Restart the tool as an administrator?", "Admin Required", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
        $currentProcess = Get-Process -Id $PID
        if ($currentProcess.ProcessName -in 'powershell', 'pwsh', 'powershell_ise') {
            Start-Process powershell.exe -Verb RunAs -ArgumentList "-File `"$script:selfPath`""
        } else {
            Start-Process -FilePath $script:selfPath -Verb RunAs
        }
        $script:Form.Close()
        return $true
    }
    return $false
}

<#
.SYNOPSIS
    A helper function to set the LoadBehavior for selected add-ins.
.DESCRIPTION
    This function consolidates the logic for enabling and disabling add-ins. It checks for required permissions, prompts for confirmation, and then iterates through the selected add-ins to set their `LoadBehavior` registry value.
.PARAMETER LoadBehavior
    The integer value to set for LoadBehavior (e.g., 3 for Enabled, 2 for Disabled).
.PARAMETER ActionName
    A string describing the action (e.g., "ENABLE") for use in confirmation dialogs.
#>
function Set-AddinLoadBehavior {
    param([int]$LoadBehavior, [string]$ActionName)

    if ($script:DgvAddins.SelectedRows.Count -eq 0) {
        $script:LblAddinStatus.Text = "No add-ins selected."
        $script:LblAddinStatus.ForeColor = 'Orange'
        return
    }
    
    $requiresAdmin = $false
    foreach ($row in $script:DgvAddins.SelectedRows) {
        if ($row.Cells['RegistryPath'].Value -like 'HKLM:*') { $requiresAdmin = $true; break }
    }

    if ($requiresAdmin -and -not (Test-IsAdmin)) {
        if (Invoke-RestartAsAdmin) { return }
        $script:LblAddinStatus.Text = "Operation cancelled. Administrator privileges required."
        $script:LblAddinStatus.ForeColor = 'Orange'
        return
    }

    $confirmation = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to $($ActionName.ToUpper()) the $($script:DgvAddins.SelectedRows.Count) selected add-in(s)?", "Confirm Action", "YesNo", "Warning")
    if ($confirmation -ne 'Yes') {
        $script:LblAddinStatus.Text = "$ActionName action cancelled."
        $script:LblAddinStatus.ForeColor = 'Gray'
        return
    }

    $script:BtnEnableAddin.Enabled = $false
    $script:BtnDisableAddin.Enabled = $false
    $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    $successes = [System.Collections.Generic.List[string]]::new()
    $failures = [System.Collections.Generic.List[string]]::new()

    try {
        foreach ($row in $script:DgvAddins.SelectedRows) {
            $addinName = $row.Cells['Friendly Name'].Value
            $regPath = $row.Cells['RegistryPath'].Value
            try {
                Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value $LoadBehavior -Force -ErrorAction Stop
                $successes.Add($addinName)
            } catch {
                $errorMessage = $_.Exception.Message
                $failures.Add("'$addinName': $errorMessage")
                Write-Log "ADD-IN MGMT ERROR: Failed to set LoadBehavior for $regPath. $errorMessage"
            }
        }

        if ($failures.Count -gt 0) {
            $script:LblAddinStatus.Text = "Completed with $($failures.Count) error(s). See details."
            $script:LblAddinStatus.ForeColor = 'Red'
            $errorDetails = $failures -join "`r`n"
            [System.Windows.Forms.MessageBox]::Show("The following errors occurred:`r`n`r`n$errorDetails", "Operation Failed for Some Add-ins", "OK", "Error")
        } else {
            $script:LblAddinStatus.Text = "Successfully $($ActionName.ToLower())ed $($successes.Count) add-in(s). Please restart the Office application(s)."
            $script:LblAddinStatus.ForeColor = 'Green'
        }
    } finally {
        $script:BtnEnableAddin.Enabled = $true
        $script:BtnDisableAddin.Enabled = $true
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
        Populate-AddinGrid
    }
}

$script:BtnEnableAddin.Add_Click({ Set-AddinLoadBehavior -LoadBehavior 3 -ActionName "Enable" })
$script:BtnDisableAddin.Add_Click({ Set-AddinLoadBehavior -LoadBehavior 2 -ActionName "Disable" })

<#
.SYNOPSIS
    Creates and configures all UI elements for the "Log Viewer" tab.
.DESCRIPTION
    This function encapsulates the creation of all controls, panels, and layouts for the Log Viewer tab.
    It uses the script scope for control variables so they can be accessed by event handlers later in the script.
.RETURNS
    A `System.Windows.Forms.TableLayoutPanel` containing the complete UI for the tab.
#>
function Create-LogViewerTabUI {
    $logViewerLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $logViewerLayout.Dock = 'Fill'
    $logViewerLayout.BackColor = [System.Drawing.SystemColors]::Window
    $logViewerLayout.Padding = (New-Object System.Windows.Forms.Padding(10))
    $logViewerLayout.ColumnCount = 2
    $logViewerLayout.RowCount = 3
    [void]$logViewerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$logViewerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$logViewerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$logViewerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$logViewerLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 0))) # Hidden column for splitter

    # --- Action Panel ---
    $actionsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $actionsPanel.Dock = 'Fill'; $actionsPanel.FlowDirection = 'LeftToRight'; $actionsPanel.AutoSize = $true; $actionsPanel.WrapContents = $true
    [void]$logViewerLayout.Controls.Add($actionsPanel, 0, 0)

    $script:BtnOpenExternally = New-AppButton -Text "Open Externally" -AutoSize -BackColor 'LightGray' -Enabled $false; $script:ToolTip.SetToolTip($script:BtnOpenExternally, "Show the 'Open With' dialog to choose an application for the current log file.")
    $script:BtnOpenLogFile = New-AppButton -Text "Open Log..." -AutoSize -BackColor 'LightBlue'
    $script:ToolTip.SetToolTip($script:BtnOpenLogFile, "Open a .log, .txt, or .csv file to view its contents.")

    $script:LblFilterText = New-Object System.Windows.Forms.Label; $script:LblFilterText.Text = "Filter Text:"; $script:LblFilterText.Margin = '10,5,0,0'; $script:LblFilterText.AutoSize = $true

    $script:TxtFilterLogText = New-Object System.Windows.Forms.TextBox; $script:TxtFilterLogText.Width = 150; $script:TxtFilterLogText.Enabled = $false

    $script:BtnFindNextLog = New-AppButton -Text "Find Next" -AutoSize -Enabled $false
    $script:ToolTip.SetToolTip($script:BtnFindNextLog, "Find the next occurrence of the filter text.")

    $script:ChkShowErrorsOnly = New-Object System.Windows.Forms.CheckBox; $script:ChkShowErrorsOnly.Text = "Show Errors Only"; $script:ChkShowErrorsOnly.AutoSize = $true; $script:ChkShowErrorsOnly.Margin = '10,5,0,0'; $script:ChkShowErrorsOnly.Enabled = $false
    $script:ToolTip.SetToolTip($script:ChkShowErrorsOnly, "Filter for lines containing 'error', 'fail', or 'exception'.")

    $script:BtnApplyLogFilter = New-AppButton -Text "Apply Filter" -AutoSize -Enabled $false
    $script:ToolTip.SetToolTip($script:BtnApplyLogFilter, "Filter the view to only show lines matching the criteria.")

    $script:BtnClearFilter = New-AppButton -Text "Clear Filter" -AutoSize -Enabled $false
    $script:ToolTip.SetToolTip($script:BtnClearFilter, "Remove all active filters and show the full log.")

    $script:BtnGroupCount = New-AppButton -Text "Group & Count" -AutoSize -Enabled $false
    $script:ToolTip.SetToolTip($script:BtnGroupCount, "Group identical lines in the current view and show their counts.")

    $script:ChkTailLog = New-Object System.Windows.Forms.CheckBox; $script:ChkTailLog.Text = "Tail"; $script:ChkTailLog.AutoSize = $true; $script:ChkTailLog.Margin = '10,5,0,0'; $script:ChkTailLog.Enabled = $false
    $script:ToolTip.SetToolTip($script:ChkTailLog, "Automatically scroll and show new content as the log file is updated.")

    $script:BtnExportToCsv = New-AppButton -Text "Export to CSV" -AutoSize -BackColor 'LightCyan' -Enabled $false
    $script:ToolTip.SetToolTip($script:BtnExportToCsv, "Export the current view (raw or grouped) to a CSV file.")
    $script:BtnClearLogViewer = New-AppButton -Text "Clear Output" -AutoSize -BackColor 'LightGray' -Enabled $false
    $script:ToolTip.SetToolTip($script:BtnClearLogViewer, "Clear the log viewer and reset all filters.")
    [void]$actionsPanel.Controls.AddRange(@($script:BtnOpenLogFile, $script:BtnOpenExternally, $script:LblFilterText, $script:TxtFilterLogText, $script:BtnFindNextLog, $script:ChkShowErrorsOnly, $script:BtnApplyLogFilter, $script:BtnClearFilter, $script:BtnGroupCount, $script:ChkTailLog, $script:BtnExportToCsv, $script:BtnClearLogViewer))

    # --- RichTextBox for Log Content ---
    $script:RtbLogViewer = New-Object System.Windows.Forms.RichTextBox
    $script:RtbLogViewer.Dock = 'Fill'
    $script:RtbLogViewer.ReadOnly = $true
    $script:RtbLogViewer.Font = New-Object System.Drawing.Font('Consolas', 9.75)
    $script:RtbLogViewer.WordWrap = $false
    $script:RtbLogViewer.ScrollBars = 'Both'
    [void]$logViewerLayout.Controls.Add($script:RtbLogViewer, 0, 1)

    # --- Status Label ---
    $script:LblLogViewerStatus = New-StatusLabel -Text "Ready. Open a log file to begin." -InitialColor 'Gray' -Dock 'Fill'
    [void]$logViewerLayout.Controls.Add($script:LblLogViewerStatus, 0, 2)
    $logViewerLayout.SetColumnSpan($script:LblLogViewerStatus, 1)

    return $logViewerLayout
}

<#
.SYNOPSIS
    Defines the UI and logic for the "Log Viewer" tab.
.DESCRIPTION
    This section builds the UI for viewing and analyzing log files, with features for filtering and pattern highlighting.
#>
#==============================================================================
# --- UI: Log Viewer Tab ---
#==============================================================================
$logViewerTabPanel = Create-LogViewerTabUI
$script:TabLogViewer.Controls.Add($logViewerTabPanel)
$script:fullLogContent = @() # Stores the original, unfiltered log lines
$script:currentLogFilePath = $null # Stores the path of the currently open log file
$script:currentLogView = @() # Stores the current filtered (but not grouped) view

function Update-LogViewerDisplay {
    param([array]$LinesToShow, [switch]$Reverse)
    if ($null -eq $LinesToShow) {
        $script:RtbLogViewer.Lines = @()
        return
    }
    $lines = $LinesToShow.Clone()
    if ($Reverse.IsPresent) { [array]::Reverse($lines) }
    $script:RtbLogViewer.Lines = $lines
    Highlight-LogText -rtb $script:RtbLogViewer -textToFind "error" -color "MistyRose"
    Highlight-LogText -rtb $script:RtbLogViewer -textToFind "fail" -color "LightCoral"
    Highlight-LogText -rtb $script:RtbLogViewer -textToFind "exception" -color "IndianRed"
}

function Check-For-LogUpdates {
    if (-not $script:currentLogFilePath -or -not (Test-Path $script:currentLogFilePath)) {
        return
    }

    try {
        $currentSize = (Get-Item -LiteralPath $script:currentLogFilePath).Length
        if ($currentSize -gt $script:lastLogFileSize) {
            $stream = New-Object System.IO.FileStream($script:currentLogFilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $reader = New-Object System.IO.StreamReader($stream)
            try {
                $stream.Seek($script:lastLogFileSize, [System.IO.SeekOrigin]::Begin) | Out-Null
                $newContentString = $reader.ReadToEnd()
                $newLines = $newContentString -split '\r?\n' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

                if ($newLines.Count -gt 0) {
                    [array]::Reverse($newLines)
                    $textToPrepend = ($newLines -join "`r`n") + "`r`n"
                    $script:RtbLogViewer.Select(0, 0)
                    $script:RtbLogViewer.SelectedText = $textToPrepend
                }
            }
            finally {
                if ($reader) { $reader.Dispose() }
                if ($stream) { $stream.Dispose() }
            }
            $script:lastLogFileSize = $currentSize
        }
    } catch { }
}

function Highlight-LogText {
    param(
        [System.Windows.Forms.RichTextBox]$rtb,
        [string]$textToFind,
        [System.Drawing.Color]$color,
        [bool]$matchCase = $false
    )
    $rtb.SuspendLayout()
    $comparison = if ($matchCase) { [System.StringComparison]::InvariantCulture } else { [System.StringComparison]::InvariantCultureIgnoreCase }
    $startIndex = 0
    while ($startIndex -lt $rtb.TextLength) {
        $foundIndex = $rtb.Text.IndexOf($textToFind, $startIndex, $comparison)
        if ($foundIndex -eq -1) { break }

        $lineIndex = $rtb.GetLineFromCharIndex($foundIndex)
        $lineStart = $rtb.GetFirstCharIndexFromLine($lineIndex)
        $lineLength = $rtb.Lines[$lineIndex].Length

        $rtb.Select($lineStart, $lineLength); $rtb.SelectionBackColor = $color
        $startIndex = $lineStart + $lineLength
    }
    $rtb.DeselectAll()
    $rtb.ResumeLayout()
}

$script:BtnOpenLogFile.Add_Click({
    if ($script:ChkTailLog.Checked) { $script:ChkTailLog.PerformClick() }
    $script:ChkTailLog.Checked = $false
    $script:BtnOpenExternally.Enabled = $false
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Log files (*.log, *.txt, *.csv)|*.log;*.txt;*.csv|All files (*.*)|*.*"
    $openFileDialog.Title = "Select a Log File"
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        if ($script:currentJob) { [System.Windows.Forms.MessageBox]::Show("Another operation is already in progress.", "Busy", "OK", "Warning"); return }

        $script:LblLogViewerStatus.Text = "Loading file..."
        $script:LblLogViewerStatus.ForeColor = 'Blue'
        $script:RtbLogViewer.Clear()

        # Reset filters for the new file
        $script:TxtFilterLogText.Clear()
        $script:ChkShowErrorsOnly.Checked = $false
        $script:fullLogContent = @()
        $script:currentLogView = @()
        $script:currentLogFilePath = $null
        $script:lastLogFileSize = 0

        $script:jobCompletionAction = {
            param($jobResult)
            if ($jobResult -is [System.Management.Automation.ErrorRecord]) {
                $script:LblLogViewerStatus.Text = "Error opening file: $($jobResult.Exception.Message)"
                $script:LblLogViewerStatus.ForeColor = 'Red'
                return
            }

            $script:fullLogContent = $jobResult.Content
            $script:currentLogView = $script:fullLogContent # Initially, the view is the full content
            $script:lastLogFileSize = (Get-Item $jobResult.FilePath).Length
            $script:currentLogFilePath = $jobResult.FilePath
            $script:BtnOpenExternally.Enabled = $true

            $script:TxtFilterLogText.Enabled = $true
            $script:BtnFindNextLog.Enabled = $true
            $script:ChkShowErrorsOnly.Enabled = $true
            $script:BtnApplyLogFilter.Enabled = $true
            $script:BtnExportToCsv.Enabled = $true
            $script:BtnClearLogViewer.Enabled = $true
            $script:BtnGroupCount.Enabled = $true
            $script:BtnClearFilter.Enabled = $true
            $script:BtnClearFilter.Enabled = $true
            $script:ChkTailLog.Enabled = $true

            Update-LogViewerDisplay -LinesToShow $script:fullLogContent
            $script:LblLogViewerStatus.Text = "Loaded $($script:fullLogContent.Count) lines from $(Split-Path $jobResult.FilePath -Leaf)"
            $script:LblLogViewerStatus.ForeColor = 'Green'
        }

        $jobArgs = @{ FilePath = $openFileDialog.FileName }
        $script:currentJob = Start-Job -ArgumentList $jobArgs -ScriptBlock {
            param($jobArgs)
            $content = Get-Content -Path $jobArgs.FilePath

            return [PSCustomObject]@{
                Content  = $content
                FilePath = $jobArgs.FilePath
            }
        }
        $script:JobTimer.Start()
    }
})

$script:BtnFindNextLog.Add_Click({
    $searchText = $script:TxtFilterLogText.Text
    if ([string]::IsNullOrWhiteSpace($searchText)) {
        $script:LblLogViewerStatus.Text = "Please enter text to find."; $script:LblLogViewerStatus.ForeColor = 'Orange'
        return
    }

    $start = $script:RtbLogViewer.SelectionStart + $script:RtbLogViewer.SelectionLength
    if ($start -ge $script:RtbLogViewer.TextLength) { $start = 0 }

    $foundIndex = $script:RtbLogViewer.Find($searchText, $start, [System.Windows.Forms.RichTextBoxFinds]::None)
    if ($foundIndex -eq -1 -and $start -ne 0) { # If not found, wrap around and search from the beginning
        $foundIndex = $script:RtbLogViewer.Find($searchText, 0, [System.Windows.Forms.RichTextBoxFinds]::None)
    }

    if ($foundIndex -ne -1) {
        # Clear previous highlight
        $script:RtbLogViewer.Select(0, $script:RtbLogViewer.TextLength)
        $script:RtbLogViewer.SelectionBackColor = $script:RtbLogViewer.BackColor

        # Highlight the new line
        $lineIndex = $script:RtbLogViewer.GetLineFromCharIndex($foundIndex)
        $lineStart = $script:RtbLogViewer.GetFirstCharIndexFromLine($lineIndex)
        $lineLength = $script:RtbLogViewer.Lines[$lineIndex].Length
        $script:RtbLogViewer.Select($lineStart, $lineLength)
        $script:RtbLogViewer.SelectionBackColor = [System.Drawing.Color]::LightGoldenrodYellow

        $script:RtbLogViewer.ScrollToCaret()
        $script:LblLogViewerStatus.Text = "Found text."; $script:LblLogViewerStatus.ForeColor = 'Green'
    } else {
        $script:LblLogViewerStatus.Text = "Text not found: '$searchText'"; $script:LblLogViewerStatus.ForeColor = 'Orange'
    }
})

$script:BtnApplyLogFilter.Add_Click({
    if ($script:fullLogContent.Count -eq 0) { return }

    try {
        $filteredContent = $script:fullLogContent
        $filterText = $script:TxtFilterLogText.Text
        $errorsOnly = $script:ChkShowErrorsOnly.Checked

        if (-not [string]::IsNullOrWhiteSpace($filterText)) {
            $filteredContent = @($filteredContent | Where-Object { $_ -like "*$filterText*" })
        }

        if ($errorsOnly) {
            $filteredContent = @($filteredContent | Where-Object { $_ -match "error|fail|exception" })
        }

        $script:currentLogView = $filteredContent        
        Update-LogViewerDisplay -LinesToShow $filteredContent -Reverse:$script:ChkTailLog.Checked
        $script:BtnExportToCsv.Enabled = ($script:RtbLogViewer.Lines.Count -gt 0)
        $script:LblLogViewerStatus.Text = "Filter applied. Showing $($filteredContent.Count) of $($script:fullLogContent.Count) lines."
        $script:LblLogViewerStatus.ForeColor = 'Green'
    } finally {
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$script:BtnClearFilter.Add_Click({
    if ($script:fullLogContent.Count -eq 0) { return }

    $script:TxtFilterLogText.Clear()
    $script:ChkShowErrorsOnly.Checked = $false

    $script:currentLogView = $script:fullLogContent
    Update-LogViewerDisplay -LinesToShow $script:fullLogContent -Reverse:$script:ChkTailLog.Checked
    $script:LblLogViewerStatus.Text = "Filter cleared. Showing all $($script:fullLogContent.Count) lines."
    $script:LblLogViewerStatus.ForeColor = 'Green'
})

$script:BtnGroupCount.Add_Click({
    if ($script:currentLogView.Count -eq 0) { return }

    $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    if ($script:ChkTailLog.Checked) {
        [System.Windows.Forms.MessageBox]::Show("Cannot group lines while Tail mode is active.", "Info", "OK", "Information")
        return
    }
    try {
        $groupedLines = $script:currentLogView | Group-Object | Sort-Object Count -Descending
        $formattedLines = $groupedLines | ForEach-Object { "($($_.Count)) $($_.Name)" }

        $script:RtbLogViewer.Lines = $formattedLines
        $script:LblLogViewerStatus.Text = "Grouped $($groupedLines.Count) unique lines."
        $script:LblLogViewerStatus.ForeColor = 'Green'
    } finally {
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$script:BtnOpenExternally.Add_Click({
    if ($script:currentLogFilePath -and (Test-Path $script:currentLogFilePath)) {
        try {
            Start-Process -FilePath "rundll32.exe" -ArgumentList "shell32.dll,OpenAs_RunDLL $($script:currentLogFilePath)" -ErrorAction Stop
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Could not show the 'Open With' dialog. Error: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    }
})

$script:ChkShowErrorsOnly.Add_Click({ $script:BtnApplyLogFilter.PerformClick() })

$script:BtnExportToCsv.Add_Click({
    if ($script:RtbLogViewer.Lines.Count -eq 0) { return }

    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveFileDialog.Title = "Export Log Data"
    $saveFileDialog.FileName = "LogExport_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        $filePath = $saveFileDialog.FileName
        $script:LblLogViewerStatus.Text = "Exporting to CSV..."
        $script:LblLogViewerStatus.ForeColor = 'Blue'
        $script:Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        try {
            $rawLines = $script:RtbLogViewer.Lines | ForEach-Object { [PSCustomObject]@{ LogEntry = $_ } }
            $rawLines | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8

            $script:LblLogViewerStatus.Text = "Successfully exported to $filePath"
            $script:LblLogViewerStatus.ForeColor = 'Green'
            [System.Windows.Forms.MessageBox]::Show("Export complete!", "Success", "OK", "Information")
        } catch {
            $script:LblLogViewerStatus.Text = "Export failed: $($_.Exception.Message)"
            $script:LblLogViewerStatus.ForeColor = 'Red'
            [System.Windows.Forms.MessageBox]::Show("Export failed:`n$($_.Exception.Message)", "Error", "OK", "Error")
        } finally {
            $script:Form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
})

$script:BtnClearLogViewer.Add_Click({
    if ($script:ChkTailLog.Checked) { $script:ChkTailLog.PerformClick() }
    $script:fullLogContent = @()
    $script:currentLogView = @()
    $script:currentLogFilePath = $null
    $script:RtbLogViewer.Clear()
    $script:TxtFilterLogText.Clear()
    $script:ChkShowErrorsOnly.Checked = $false
    $script:BtnOpenLogFile.Enabled = $true # Ensure this is always re-enabled
    @($script:TxtFilterLogText, $script:BtnFindNextLog, $script:ChkShowErrorsOnly, $script:BtnApplyLogFilter,
      $script:BtnExportToCsv, $script:BtnClearLogViewer, $script:BtnGroupCount, $script:BtnClearFilter,
      $script:BtnOpenExternally, $script:ChkTailLog) | ForEach-Object { $_.Enabled = $false }

    $script:LblLogViewerStatus.Text = "Ready. Open a log file to begin."
    $script:LblLogViewerStatus.ForeColor = 'Gray'
})

$script:logPollTimer.Add_Tick({ Check-For-LogUpdates })

$script:ChkTailLog.Add_Click({
    if ($script:ChkTailLog.Checked) {
        Update-LogViewerDisplay -LinesToShow $script:currentLogView -Reverse
        $script:lastLogFileSize = (Get-Item $script:currentLogFilePath).Length # Ensure size is current
        $script:logPollTimer.Start() # Start the timer
        $script:LblLogViewerStatus.Text = "Tail mode enabled. Polling for changes... (newest on top)"
        $script:LblLogViewerStatus.ForeColor = 'Blue'
        @($script:BtnOpenLogFile, $script:BtnFindNextLog, $script:BtnApplyLogFilter, $script:BtnClearFilter, $script:BtnGroupCount, $script:TxtFilterLogText, $script:ChkShowErrorsOnly, $script:BtnExportToCsv) | ForEach-Object { $_.Enabled = $false }
    } else {
        $script:logPollTimer.Stop() # Stop the timer
        Update-LogViewerDisplay -LinesToShow $script:currentLogView # No -Reverse
        $script:LblLogViewerStatus.Text = "Tail mode disabled."; $script:LblLogViewerStatus.ForeColor = 'Gray'
        @($script:BtnOpenLogFile, $script:BtnFindNextLog, $script:BtnApplyLogFilter, $script:BtnClearFilter, $script:BtnGroupCount, $script:TxtFilterLogText, $script:ChkShowErrorsOnly, $script:BtnExportToCsv) | ForEach-Object { $_.Enabled = $true }
    }
})

<#
.SYNOPSIS
    Defines the logic for the main application menu strip.
.DESCRIPTION
    This section contains the event handlers for the "Help" menu items, including "Check for Updates", "About", and "Settings".
#>
$latestVersionUrl = "https://raw.githubusercontent.com/imumeshk/lc_tool/refs/heads/main/updater.json"
$currentVersion = "1.0.0"

$script:CheckUpdatesMenuItem.Add_Click({
    try {
        $internet = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet
        if (-not $internet) {
            [System.Windows.Forms.MessageBox]::Show("No Internet connection detected.", "Update Check", "OK", "Error")
            return
        }
        $cacheBuster = Get-Random
        $urlWithBuster = "$latestVersionUrl`?$cacheBuster"

        $webClient = New-Object System.Net.WebClient
        $webClient.CachePolicy = New-Object System.Net.Cache.RequestCachePolicy('NoCacheNoStore')
        $webClient.Headers.Add("Cache-Control", "no-cache")
        $webClient.Headers.Add("Pragma", "no-cache")
        $jsonContent = $webClient.DownloadString($urlWithBuster)
        $json = $jsonContent | ConvertFrom-Json
        $webClient.Dispose()

        $serverVersion = [version]$json.version
        $localVersion = [version]$currentVersion

        if ($serverVersion -gt $localVersion) {
            if ((Show-UpdateDialog -Version $json.version -Notes $json.notes) -eq [System.Windows.Forms.DialogResult]::Yes) {
                Start-Process $json.downloadUrl
            }
        } else {
            Show-UpToDateDialog -FoundVersion $json.version
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error checking for updates: $($_.Exception.Message)", "Update Check", "OK", "Error")
    }
})

$script:AboutMenuItem.Add_Click({
    $aboutText = @"
Litera Compare Management Tool
Version $currentVersion

This tool provides a centralized interface for managing Litera Compare installations, including cleanup, backup/restore, and troubleshooting.

Developed to simplify common support tasks.
"@
    [System.Windows.Forms.MessageBox]::Show($aboutText, "About This Tool", "OK", "Information")
})

$script:SettingsMenuItem.Add_Click({ Show-SettingsDialog })

<#
.SYNOPSIS
    Defines the UI and logic for the advanced "Pro Mode" feature.
.DESCRIPTION
    This section contains the functions and UI components for the "Litera Logging" tab, which is revealed in Pro Mode. It provides advanced controls for enabling detailed application logging and crash dump generation for deep troubleshooting. This feature requires administrator privileges.
#>
#==============================================================================
# --- Pro Mode: Litera Logging Tab ---
#==============================================================================
$script:CrashDumpSubFolderMap = @{
    'lcp_main.exe'   = 'LC4W Dump'
    'lcx_main.exe'   = 'LC4XL Dump'
    'lcp_ppt.exe'    = 'LC4PPT Dump'
    'lcp_pdfcmp.exe' = 'LC4PDF Dump'
}

<#
.SYNOPSIS
    Sets a specific integer value in a Litera customization XML file.
.DESCRIPTION
    Finds or creates a specified XML node and sets its `INT_VALUE` attribute. This is used to control logging flags in `Customize.xml` and related files.
#>
function Set-XmlValue {
    param([string]$filePath, [string]$xpath, [string]$newValue)
    try {
        $dir = Split-Path $filePath -Parent
        if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
        if (-not (Test-Path $filePath)) {
            '<Customize/>' | Out-File -FilePath $filePath -Encoding UTF8
            Write-Log "Set-XmlValue: Created new file: $filePath"
        }
        [xml]$xml = Get-Content -Path $filePath -Encoding UTF8
        $node = $xml.SelectSingleNode($xpath)

        if (-not $node) {
            $nodeName = $xpath.Substring($xpath.LastIndexOf('/') + 1)
            if ($nodeName) { $node = $xml.SelectSingleNode("//$nodeName") }
        }

        if (-not $node) {
            Write-Log "Set-XmlValue: Node not found for '$xpath', creating it."
            $parts = $xpath.Split('/') | Where-Object { $_ -ne '' }
            $parent = $xml.DocumentElement
            for ($i = 1; $i -lt $parts.Length; $i++) {
                $p = $parts[$i]
                $child = $parent.SelectSingleNode($p)
                if (-not $child) {
                    $child = $xml.CreateElement($p)
                    $parent.AppendChild($child) | Out-Null
                }
                $parent = $child
            }
            $node = $parent
        }
        $oldValue = if ($node.Attributes['INT_VALUE']) { $node.Attributes['INT_VALUE'].Value } else { "[not set]" }
        $node.SetAttribute("INT_VALUE", "$newValue")
        $xml.Save($filePath)
        Write-Log "Set-XmlValue: Updated $filePath : $xpath. INT_VALUE from '$oldValue' to '$newValue'"
    }
    catch {
        Write-Log "Set-XmlValue ERROR for $filePath : $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Retrieves a specific integer value from a Litera customization XML file.
.DESCRIPTION
    Reads the `INT_VALUE` attribute from a specified XML node.
#>
function Get-XmlValue {
    param([string]$filePath, [string]$xpath)
    if (-not (Test-Path $filePath)) { Write-Log "Get-XmlValue: File not found: $filePath"; return $null }
    try {
        [xml]$xml = Get-Content -Path $filePath -Encoding UTF8
        $node = $xml.SelectSingleNode($xpath)
        if (-not $node) {
            $nodeName = $xpath.Substring($xpath.LastIndexOf('/') + 1)
            if ($nodeName) { $node = $xml.SelectSingleNode("//$nodeName") }
        }
        if ($node -and $node.Attributes['INT_VALUE']) {
            return $node.Attributes['INT_VALUE'].Value
        }
        else {
            Write-Log "Get-XmlValue: Node or INT_VALUE attribute not found at $xpath in $filePath"
        }
    }
    catch {
        Write-Log "Get-XmlValue ERROR for $filePath : $($_.Exception.Message)"
    }
    return $null
}

<#
.SYNOPSIS
    Checks if crash dumps are enabled for a specific executable.
.DESCRIPTION
    Tests for the existence of the required registry key under Windows Error Reporting's `LocalDumps` key.
#>
function Get-CrashDumpStatus {
    param([string]$exeName)
    Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps\$exeName"
}

<#
.SYNOPSIS
    Enables crash dump generation for a specific executable.
#>
function Enable-CrashDump {
    param([string]$exeName, [string]$folderName)
    try {
        $key = "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps\$exeName"
        if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
        New-ItemProperty -Path $key -Name "DumpType" -Value 2 -PropertyType DWord -Force | Out-Null
        $dumpFolder = Join-Path $script:logRoot "crash_dump\$folderName"
        if (-not (Test-Path $dumpFolder)) { New-Item -ItemType Directory -Path $dumpFolder -Force | Out-Null }
        New-ItemProperty -Path $key -Name "DumpFolder" -Value $dumpFolder -PropertyType ExpandString -Force | Out-Null

        Write-Log "Enable-CrashDump: Enabled $exeName => $dumpFolder"
    }
    catch {
        Write-Log "Enable-CrashDump ERROR: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Disables crash dump generation for a specific executable.
#>
function Disable-CrashDump {
    param([string]$exeName)
    try {
        $key = "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps\$exeName"
        if (Test-Path $key) {
            Remove-Item -Path $key -Recurse -Force
            Write-Log "Disable-CrashDump: Disabled $exeName"
        }
        else {
            Write-Log "Disable-CrashDump: Key not found for $exeName"
        }
    }
    catch {
        Write-Log "Disable-CrashDump ERROR: $($_.Exception.Message)"
        throw
    }
}

<#
.SYNOPSIS
    Refreshes the status of all controls on the Pro Mode tab.
.DESCRIPTION
    Reads the current logging and crash dump settings from the XML files and registry, and updates the state of the corresponding checkboxes in the UI.
#>
function Update-ProModeStatus {
    if (-not $script:xmlLoggingCheckBoxes) { return }

    try {
        $customizeXmlPath = "C:\ProgramData\Litera\Customize\Customize.xml"
        $pptCustomizeXmlPath = "C:\ProgramData\Litera\Customize\PPTCustomize.xml"
        $pdfCustomizeXmlPath = "C:\ProgramData\Litera\Customize\PDFCustomize.xml"

        $updateXmlCheckbox = {
            param($checkbox, $filePath, $xpath)
            
            $originalText = $checkbox.Tag
            $checkbox.Text = $originalText
            $checkbox.ForeColor = [System.Drawing.Color]::Black
            $checkbox.Enabled = $true

            if (Test-Path $filePath) {
                $checkbox.Checked = (Get-XmlValue $filePath $xpath) -eq '1'
            } else {
                $checkbox.Checked = $false
                $checkbox.Enabled = $false
                $checkbox.Text = "$originalText (File not found)"
                $checkbox.ForeColor = [System.Drawing.Color]::Gray
            }
        }

        & $updateXmlCheckbox $script:xmlLoggingCheckBoxes['DMS']   $customizeXmlPath    "/Customize/EnableDMSLogging"
        & $updateXmlCheckbox $script:xmlLoggingCheckBoxes['Addin'] $customizeXmlPath    "/Customize/EnableCPOfficeAddinLogging"
        & $updateXmlCheckbox $script:xmlLoggingCheckBoxes['Word']  $customizeXmlPath    "/Customize/MainEnableCPLogging"
        & $updateXmlCheckbox $script:xmlLoggingCheckBoxes['PPT']   $pptCustomizeXmlPath "/Customize/MainEnableCPLoggingPPT"
        & $updateXmlCheckbox $script:xmlLoggingCheckBoxes['PDF']   $pdfCustomizeXmlPath "/Customize/PDFEnableCPLogging"

        $enabledXmlCheckboxes = $script:xmlLoggingCheckBoxes.Values | Where-Object { $_.Enabled }
        $script:ChkSelectAllLc.Checked = $enabledXmlCheckboxes.Count -gt 0 -and ($enabledXmlCheckboxes | Where-Object { -not $_.Checked }).Count -eq 0

        foreach ($chk in $script:crashLoggingCheckBoxes.Values) {
            $exeName = $chk.Tag
            $isChecked = Get-CrashDumpStatus $exeName
            $chk.Checked = $isChecked
        }
        $script:ChkSelectAllCrash.Checked = ($script:crashLoggingCheckBoxes.Values | Where-Object { -not $_.Checked }).Count -eq 0

        $script:LblLcLoggingStatus.Text = "Status refreshed."
        $script:LblLcLoggingStatus.ForeColor = [System.Drawing.Color]::Blue
        $script:LblCrashLoggingStatus.Text = "Status refreshed."
        $script:LblCrashLoggingStatus.ForeColor = [System.Drawing.Color]::Blue
        Write-Log "Update-ProModeStatus: refreshed"
    }
    catch {
        $errorMessage = "Error refreshing: $($_.Exception.Message)"
        if ($script:LblLcLoggingStatus) {
            $script:LblLcLoggingStatus.Text = $errorMessage
            $script:LblLcLoggingStatus.ForeColor = [System.Drawing.Color]::Red
        }
        if ($script:LblCrashLoggingStatus) {
            $script:LblCrashLoggingStatus.Text = $errorMessage
            $script:LblCrashLoggingStatus.ForeColor = [System.Drawing.Color]::Red
        }
        Write-Log "Update-ProModeStatus ERROR: $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
    Creates and configures all UI elements for the "Litera Logging" (Pro Mode) tab.
.DESCRIPTION
    This function encapsulates the creation of all controls, panels, and layouts for the advanced logging tab.
    It uses the script scope for control variables so they can be accessed by event handlers later in the script.
    This tab is only shown when "Pro Mode" is enabled.
.RETURNS
    A `System.Windows.Forms.TabPage` containing the complete UI for the tab.
#>
function Create-LiteraLoggingTabUI {
    $script:TabLiteraLogging = New-Object System.Windows.Forms.TabPage
    $script:TabLiteraLogging.Text = "Litera Logging"

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'; $layout.ColumnCount = 1; $layout.RowCount = 3
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $script:TabLiteraLogging.Controls.Add($layout)

    $grpXml = New-Object System.Windows.Forms.GroupBox; $grpXml.Text = "Litera Compare Logging (XML)"; $grpXml.Dock = 'Fill'
    $layout.Controls.Add($grpXml, 0, 0)
    
    $xmlLayout = New-Object System.Windows.Forms.TableLayoutPanel; $xmlLayout.Dock = 'Fill'; $xmlLayout.RowCount = 3
    $xmlLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $xmlLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $xmlLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $grpXml.Controls.Add($xmlLayout)

    $script:ChkSelectAllLc = New-Object System.Windows.Forms.CheckBox; $script:ChkSelectAllLc.Text = "Select All"
    $xmlLoggingPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $xmlLoggingPanel.Dock = 'Fill'; $xmlLoggingPanel.FlowDirection = 'TopDown'
    $script:xmlLoggingCheckBoxes = @{
        'DMS'   = (New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Enable DMS logs"; AutoSize = $true; Tag = "Enable DMS logs" })
        'Addin' = (New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Enable Add-ins logs"; AutoSize = $true; Tag = "Enable Add-ins logs" })
        'Word'  = (New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Enable Litera Compare for Word logs"; AutoSize = $true; Tag = "Enable Litera Compare for Word logs" })
        'PPT'   = (New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Enable Litera Compare for PPT logs"; AutoSize = $true; Tag = "Enable Litera Compare for PPT logs" })
        'PDF'   = (New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Enable Litera Compare for PDF logs"; AutoSize = $true; Tag = "Enable Litera Compare for PDF logs" })
    }
    $xmlLoggingPanel.Controls.AddRange($script:xmlLoggingCheckBoxes.Values)
    
    $xmlActionPanel = New-Object System.Windows.Forms.TableLayoutPanel; $xmlActionPanel.Dock = 'Fill'; $xmlActionPanel.ColumnCount = 4; $xmlActionPanel.RowCount = 1; $xmlActionPanel.AutoSize = $true
    $xmlActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $xmlActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $xmlActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $xmlActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $script:BtnEnableLcLogging = New-AppButton -Text "Enable Logging" -AutoSize -BackColor 'LightGreen'
    $script:BtnDisableLcLogging = New-AppButton -Text "Disable Logging" -AutoSize -BackColor 'LightCoral'
    $script:BtnRefreshLcLogging = New-AppButton -Text "Refresh Status" -AutoSize -BackColor 'LightGray'
    $script:LblLcLoggingStatus = New-StatusLabel -Text "Logging status: Idle" -InitialColor 'Gray' -Dock 'Fill' -TextAlign 'MiddleRight'

    $xmlActionPanel.Controls.Add($script:BtnEnableLcLogging, 0, 0)
    $xmlActionPanel.Controls.Add($script:BtnDisableLcLogging, 1, 0)
    $xmlActionPanel.Controls.Add($script:BtnRefreshLcLogging, 2, 0)
    $xmlActionPanel.Controls.Add($script:LblLcLoggingStatus, 3, 0)

    $xmlLayout.Controls.Add($script:ChkSelectAllLc, 0, 0)
    $xmlLayout.Controls.Add($xmlLoggingPanel, 0, 1)
    $xmlLayout.Controls.Add($xmlActionPanel, 0, 2)

    $grpCrash = New-Object System.Windows.Forms.GroupBox; $grpCrash.Text = "Litera Compare Crash Dumps (Registry)"; $grpCrash.Dock = 'Fill'
    $layout.Controls.Add($grpCrash, 0, 1)
    
    $crashLayout = New-Object System.Windows.Forms.TableLayoutPanel; $crashLayout.Dock = 'Fill'; $crashLayout.RowCount = 3
    $crashLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $crashLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $crashLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $grpCrash.Controls.Add($crashLayout)

    $script:ChkSelectAllCrash = New-Object System.Windows.Forms.CheckBox; $script:ChkSelectAllCrash.Text = "Select All"
    $crashLoggingPanel = New-Object System.Windows.Forms.FlowLayoutPanel; $crashLoggingPanel.Dock = 'Fill'; $crashLoggingPanel.FlowDirection = 'TopDown'
    $script:crashLoggingCheckBoxes = @{
        'Word'  = (New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Enable Litera Compare for Word Crash Dump"; AutoSize = $true; Tag = "lcp_main.exe" })
        'Excel' = (New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Enable Litera Compare for Excel Crash Dump"; AutoSize = $true; Tag = "lcx_main.exe" })
        'PPT'   = (New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Enable Litera Compare for PPT Crash Dump"; AutoSize = $true; Tag = "lcp_ppt.exe" })
        'PDF'   = (New-Object System.Windows.Forms.CheckBox -Property @{ Text = "Enable Litera Compare for PDF Crash Dump"; AutoSize = $true; Tag = "lcp_pdfcmp.exe" })
    }
    $crashLoggingPanel.Controls.AddRange($script:crashLoggingCheckBoxes.Values)
    
    $crashActionPanel = New-Object System.Windows.Forms.TableLayoutPanel; $crashActionPanel.Dock = 'Fill'; $crashActionPanel.ColumnCount = 5; $crashActionPanel.RowCount = 1; $crashActionPanel.AutoSize = $true
    $crashActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $crashActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $crashActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $crashActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $crashActionPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $script:BtnEnableCrashLogging = New-AppButton -Text "Enable Crash Dumps" -AutoSize -BackColor 'LightGreen'
    $script:BtnDisableCrashLogging = New-AppButton -Text "Disable Crash Dumps" -AutoSize -BackColor 'LightCoral'
    $script:BtnRefreshCrashLogging = New-AppButton -Text "Refresh Status" -AutoSize -BackColor 'LightGray'
    $script:BtnOpenCrashRegKey = New-AppButton -Text "Open Registry Key" -AutoSize -BackColor 'LightGray'
    $script:ToolTip.SetToolTip($script:BtnOpenCrashRegKey, "Opens the selected crash dump's key in the Registry Editor.")
    $script:LblCrashLoggingStatus = New-StatusLabel -Text "Crash dump status: Idle" -InitialColor 'Gray' -Dock 'Fill' -TextAlign 'MiddleRight'

    $crashActionPanel.Controls.Add($script:BtnEnableCrashLogging, 0, 0)
    $crashActionPanel.Controls.Add($script:BtnDisableCrashLogging, 1, 0)
    $crashActionPanel.Controls.Add($script:BtnRefreshCrashLogging, 2, 0)
    $crashActionPanel.Controls.Add($script:BtnOpenCrashRegKey, 3, 0)
    $crashActionPanel.Controls.Add($script:LblCrashLoggingStatus, 4, 0)

    $crashLayout.Controls.Add($script:ChkSelectAllCrash, 0, 0)
    $crashLayout.Controls.Add($crashLoggingPanel, 0, 1)
    $crashLayout.Controls.Add($crashActionPanel, 0, 2)

    $restorePanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $restorePanel.Dock = 'Fill'; $restorePanel.FlowDirection = 'LeftToRight'; $restorePanel.WrapContents = $false
    $script:BtnRestoreProModeDefaults = New-AppButton -Text "Restore Defaults" -AutoSize -BackColor 'LightGray'
    $restorePanel.Controls.Add($script:BtnRestoreProModeDefaults)
    $layout.Controls.Add($restorePanel, 0, 2)

    return $script:TabLiteraLogging
}

$script:ToggleProModeItem.Add_Click({
    if ($ToggleProModeItem.Checked) {
        if (-not (Test-IsAdmin)) {
            if (Invoke-RestartAsAdmin) {
                # A restart was initiated, so we stop execution here.
                return
            }
            $ToggleProModeItem.Checked = $false; return # User cancelled the admin prompt
        }

        if ($null -eq $script:TabLiteraLogging -or $script:TabLiteraLogging.IsDisposed) {
            # Create the tab UI by calling the new function
            Create-LiteraLoggingTabUI | Out-Null

            $script:ChkSelectAllLc.Add_Click({
                foreach ($chk in ($script:xmlLoggingCheckBoxes.Values | Where-Object { $_.Enabled })) { $chk.Checked = $script:ChkSelectAllLc.Checked }
            })

            $script:ChkSelectAllCrash.Add_Click({
                foreach ($chk in $script:crashLoggingCheckBoxes.Values) { $chk.Checked = $script:ChkSelectAllCrash.Checked }
            })

            $script:BtnEnableLcLogging.Add_Click({
                $checkedBoxes = $script:xmlLoggingCheckBoxes.Values | Where-Object { $_.Checked }
                if ($checkedBoxes.Count -eq 0) {
                    $script:LblLcLoggingStatus.Text = "No logging options selected."; $script:LblLcLoggingStatus.ForeColor = 'Orange'
                    return
                }
                $ans = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to enable the selected logging options?", "Confirm Enable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($ans -ne 'Yes') { $script:LblLcLoggingStatus.Text = "Enable operation cancelled."; $script:LblLcLoggingStatus.ForeColor = 'Gray'; return }
                try {
                    if ($script:xmlLoggingCheckBoxes['DMS'].Checked)   { Set-XmlValue "C:\ProgramData\Litera\Customize\Customize.xml" "/Customize/EnableDMSLogging" 1 }
                    if ($script:xmlLoggingCheckBoxes['Addin'].Checked) { Set-XmlValue "C:\ProgramData\Litera\Customize\Customize.xml" "/Customize/EnableCPOfficeAddinLogging" 1 }
                    if ($script:xmlLoggingCheckBoxes['Word'].Checked)  { Set-XmlValue "C:\ProgramData\Litera\Customize\Customize.xml" "/Customize/MainEnableCPLogging" 1 }
                    if ($script:xmlLoggingCheckBoxes['PPT'].Checked)   { Set-XmlValue "C:\ProgramData\Litera\Customize\PPTCustomize.xml" "/Customize/MainEnableCPLoggingPPT" 1 }
                    if ($script:xmlLoggingCheckBoxes['PDF'].Checked)   { Set-XmlValue "C:\ProgramData\Litera\Customize\PDFCustomize.xml" "/Customize/PDFEnableCPLogging" 1 }

                    $script:LblLcLoggingStatus.Text = "Selected logging enabled successfully."; $script:LblLcLoggingStatus.ForeColor = 'Green'
                    Write-Log "Enabled selected XML logging options"
                    Update-ProModeStatus
                }
                catch {
                    $script:LblLcLoggingStatus.Text = "Failed to enable logging: $($_.Exception.Message)"; $script:LblLcLoggingStatus.ForeColor = 'Red'
                    Write-Log "ERROR enabling logging: $($_.Exception.Message)"
                }
            })

            $script:BtnDisableLcLogging.Add_Click({
                $checkedBoxes = $script:xmlLoggingCheckBoxes.Values | Where-Object { $_.Checked }
                if ($checkedBoxes.Count -eq 0) {
                    $script:LblLcLoggingStatus.Text = "No logging options selected."; $script:LblLcLoggingStatus.ForeColor = 'Orange'
                    return
                }
                $ans = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to disable the selected logging options (restoring defaults)?", "Confirm Disable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($ans -ne 'Yes') { $script:LblLcLoggingStatus.Text = "Disable operation cancelled."; $script:LblLcLoggingStatus.ForeColor = 'Gray'; return }
                try {
                    if ($script:xmlLoggingCheckBoxes['DMS'].Checked)   { Set-XmlValue "C:\ProgramData\Litera\Customize\Customize.xml" "/Customize/EnableDMSLogging" 0 }
                    if ($script:xmlLoggingCheckBoxes['Addin'].Checked) { Set-XmlValue "C:\ProgramData\Litera\Customize\Customize.xml" "/Customize/EnableCPOfficeAddinLogging" 0 }
                    if ($script:xmlLoggingCheckBoxes['Word'].Checked)  { Set-XmlValue "C:\ProgramData\Litera\Customize\Customize.xml" "/Customize/MainEnableCPLogging" 5 }
                    if ($script:xmlLoggingCheckBoxes['PPT'].Checked)   { Set-XmlValue "C:\ProgramData\Litera\Customize\PPTCustomize.xml" "/Customize/MainEnableCPLoggingPPT" 5 }
                    if ($script:xmlLoggingCheckBoxes['PDF'].Checked)   { Set-XmlValue "C:\ProgramData\Litera\Customize\PDFCustomize.xml" "/Customize/PDFEnableCPLogging" 5 }

                    $script:LblLcLoggingStatus.Text = "Selected logging disabled successfully."; $script:LblLcLoggingStatus.ForeColor = 'Green'
                    Write-Log "Disabled selected XML logging options"
                    Update-ProModeStatus
                }
                catch {
                    $script:LblLcLoggingStatus.Text = "Failed to disable logging: $($_.Exception.Message)"; $script:LblLcLoggingStatus.ForeColor = 'Red'
                    Write-Log "ERROR disabling logging: $($_.Exception.Message)"
                }
            })

            $script:BtnEnableCrashLogging.Add_Click({
                $checkedBoxes = $script:crashLoggingCheckBoxes.Values | Where-Object { $_.Checked }
                if ($checkedBoxes.Count -eq 0) {
                    $script:LblCrashLoggingStatus.Text = "No crash dump options selected."; $script:LblCrashLoggingStatus.ForeColor = 'Orange'
                    return
                }
                $ans = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to enable the selected crash dumps?", "Confirm Enable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                if ($ans -ne 'Yes') { $script:LblCrashLoggingStatus.Text = "Enable operation cancelled."; $script:LblCrashLoggingStatus.ForeColor = 'Gray'; return }
                try {
                    foreach ($chk in $checkedBoxes) {
                        $exe = $chk.Tag
                        $subFolder = $script:CrashDumpSubFolderMap[$exe]
                        if (-not $subFolder) { Write-Log "No subfolder mapping for $exe"; continue }
                        Enable-CrashDump -exeName $exe -folderName $subFolder
                    }
                    $script:LblCrashLoggingStatus.Text = "Selected crash dumps enabled successfully."; $script:LblCrashLoggingStatus.ForeColor = 'Green'
                    Write-Log "Enabled selected crash dumps"
                    Update-ProModeStatus
                }
                catch {
                    $script:LblCrashLoggingStatus.Text = "Failed to enable crash dumps: $($_.Exception.Message)"; $script:LblCrashLoggingStatus.ForeColor = 'Red'
                    Write-Log "ERROR enabling crash dumps: $($_.Exception.Message)"
                }
            })
            $script:BtnDisableCrashLogging.Add_Click({
                $checkedBoxes = $script:crashLoggingCheckBoxes.Values | Where-Object { $_.Checked }
                if ($checkedBoxes.Count -eq 0) {
                    $script:LblCrashLoggingStatus.Text = "No crash dump options selected."; $script:LblCrashLoggingStatus.ForeColor = 'Orange'
                    return
                }
                $ans = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to disable the selected crash dumps?", "Confirm Disable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($ans -ne 'Yes') { $script:LblCrashLoggingStatus.Text = "Disable operation cancelled."; $script:LblCrashLoggingStatus.ForeColor = 'Gray'; return }
                try {
                    foreach ($chk in $checkedBoxes) {
                        $exe = $chk.Tag
                        Disable-CrashDump -exeName $exe
                    }
                    $script:LblCrashLoggingStatus.Text = "Selected crash dumps disabled successfully."; $script:LblCrashLoggingStatus.ForeColor = 'Green'
                    Write-Log "Disabled selected crash dumps"
                    Update-ProModeStatus
                }
                catch {
                    $script:LblCrashLoggingStatus.Text = "Failed to disable crash dumps: $($_.Exception.Message)"; $script:LblCrashLoggingStatus.ForeColor = 'Red'
                    Write-Log "ERROR disabling crash dumps: $($_.Exception.Message)"
                }
            })

            $script:BtnOpenCrashRegKey.Add_Click({
                $selectedCheckboxes = $script:crashLoggingCheckBoxes.Values | Where-Object { $_.Checked }
                if ($selectedCheckboxes.Count -eq 0) {
                    $script:LblCrashLoggingStatus.Text = "Please check a crash dump option to view its registry key."; $script:LblCrashLoggingStatus.ForeColor = 'Orange'; return
                }
                if ($selectedCheckboxes.Count -gt 1) {
                    $script:LblCrashLoggingStatus.Text = "Please check only one crash dump option to open its registry key."; $script:LblCrashLoggingStatus.ForeColor = 'Orange'; return
                }

                $exeName = $selectedCheckboxes[0].Tag
                $psPath = "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps\$exeName"
                
                # Convert PowerShell path to a standard registry path that RegEdit understands
                $regPath = $psPath -replace '.*?Registry::', ''

                try {
                    $regeditKeyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit"
                    if (-not (Test-Path $regeditKeyPath)) { New-Item -Path $regeditKeyPath -Force | Out-Null }

                    # Set the LastKey value so RegEdit opens to the correct location
                    Set-ItemProperty -Path $regeditKeyPath -Name "LastKey" -Value $regPath -Force -ErrorAction Stop
                    
                    # Start RegEdit. The -m switch allows multiple instances.
                    Start-Process "regedit.exe" -ArgumentList "-m" -ErrorAction Stop
                    
                    $script:LblCrashLoggingStatus.Text = "Opening Registry Editor to the selected key."; $script:LblCrashLoggingStatus.ForeColor = 'Blue'
                } catch {
                    $errorMessage = "Failed to open Registry Editor: $($_.Exception.Message)"
                    $script:LblCrashLoggingStatus.Text = $errorMessage; $script:LblCrashLoggingStatus.ForeColor = 'Red'
                    Write-Log "REGEDIT ERROR: $errorMessage"
                }
            })

            $script:BtnRefreshLcLogging.Add_Click({ Update-ProModeStatus })
            $script:BtnRefreshCrashLogging.Add_Click({ Update-ProModeStatus })

            $script:BtnRestoreProModeDefaults.Add_Click({
                $ans = [System.Windows.Forms.MessageBox]::Show("This will reset XML logging INT_VALUEs and remove crash dump keys. Continue?", "Restore Defaults", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
                    try {
                        Set-XmlValue "C:\ProgramData\Litera\Customize\Customize.xml" "/Customize/EnableDMSLogging" 0
                        Set-XmlValue "C:\ProgramData\Litera\Customize\Customize.xml" "/Customize/EnableCPOfficeAddinLogging" 0
                        Set-XmlValue "C:\ProgramData\Litera\Customize\Customize.xml" "/Customize/MainEnableCPLogging" 5
                        Set-XmlValue "C:\ProgramData\Litera\Customize\PPTCustomize.xml" "/Customize/MainEnableCPLoggingPPT" 5
                        Set-XmlValue "C:\ProgramData\Litera\Customize\PDFCustomize.xml" "/Customize/PDFEnableCPLogging" 5
                        foreach ($chk in $script:crashLoggingCheckBoxes.Values) { Disable-CrashDump -exeName $chk.Tag }
                        Update-ProModeStatus
                        $script:LblLcLoggingStatus.Text = "All logging reset to defaults."; $script:LblLcLoggingStatus.ForeColor = 'Green'
                        $script:LblCrashLoggingStatus.Text = "All crash dumps reset to defaults."; $script:LblCrashLoggingStatus.ForeColor = 'Green'
                        Write-Log "Restore Defaults executed for XML and crash dump settings"
                    }
                    catch {
                        $script:LblLcLoggingStatus.Text = "Failed to restore defaults: $($_.Exception.Message)"; $script:LblLcLoggingStatus.ForeColor = 'Red'
                        $script:LblCrashLoggingStatus.Text = "Failed to restore defaults: $($_.Exception.Message)"; $script:LblCrashLoggingStatus.ForeColor = 'Red'
                        Write-Log "ERROR restoring defaults: $($_.Exception.Message)"
                    }
                }
            })
        }

        if (-not $script:TabControl.TabPages.Contains($script:TabLiteraLogging)) {
            $script:TabControl.TabPages.Add($script:TabLiteraLogging)
        }

        Update-ProModeStatus
    }
    else {
        if ($script:TabLiteraLogging -and $script:TabControl.TabPages.Contains($script:TabLiteraLogging)) {
            $script:TabControl.TabPages.Remove($script:TabLiteraLogging)
        }
    }
})

<#
.SYNOPSIS
    Final form display and cleanup logic.
#>
if ($StartInProMode) {
    $script:ToggleProModeItem.PerformClick()
}

$script:Form.ShowDialog()

$script:JobTimer.Dispose()
$script:logPollTimer.Dispose()
$script:Form.Dispose()
