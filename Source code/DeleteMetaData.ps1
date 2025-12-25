<#
DeleteMetaData-UI.ps1 (Upgraded: Remember + Tooltips + Loaded Display)
By NekoJonez - v0.2 BETA
Release build date: 25/12/2025
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- CONFIG ----------
$AllowedExtensions = @(".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".heic", ".heif", ".gif", ".bmp", ".jfif", ".avif")

# ---------- Config persistence ----------
$AppName = "ExifToolUI"
$ConfigDir = Join-Path $Env:APPDATA $AppName
$ConfigPath = Join-Path $ConfigDir "config.json"

# If the user stored a config, let's load it.
function Import-Config {
    if (Test-Path -LiteralPath $ConfigPath) { try { return (Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json) } catch {} }
    return $Null
}

# If the user wants to save a config, let's let them.
function Save-Config($Obj) {
    if (!(Test-Path -LiteralPath $ConfigDir)) { New-Item -ItemType Directory -Path $ConfigDir -Force | Out-Null }
    ($Obj | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
}

# ---------- Helpers & functions ----------
# ! TODO: comments
# With this function we are going to make sure that if the folder doesn't exist, we create it first.
function Set-Ensure-Dir([string]$Path) {
    if (!(Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Get-RelativePath([string]$Base, [string]$Full) {
    $BaseTrim = $Base.TrimEnd('\')
    if ($Full.Length -le $BaseTrim.Length) { return [IO.Path]::GetFileName($Full) }
    return $Full.Substring($BaseTrim.Length).TrimStart('\')
}

# With this function we are going to add to the log with a datetime stamp.
function Add-To-Log([System.Windows.Forms.TextBox]$InputTextbox, [string]$Logline) {
    $Timestamp_Get = (Get-Date).ToString("HH:mm:ss")
    $InputTextbox.AppendText("[$Timestamp_Get] $Logline`r`n")
    $InputTextbox.SelectionStart = $InputTextbox.TextLength
    $InputTextbox.ScrollToCaret()
}

# We are going to test if the tool isn't corrupted/an actual ExifTool and get the version number.
function Test-ExifToolRunnable {
    param([Parameter(Mandatory)][string]$Path)

    if (!(Test-Path -LiteralPath $Path)) { return $Null }

    # Fast filename sanity check (allows versioned filenames)
    $Name = [IO.Path]::GetFileName($Path)
    if ($Name -notmatch '(?i)exiftool') { return $Null }

    # Try runtime check first (authoritative)
    try {
        $Ver = & $Path -ver 2>$Null
        if ($Ver -and $Ver.Trim() -match '^\d+(\.\d+)*$') {
            return $Ver.Trim()
        }
    }
    catch { return $Null }

    # Fallback: metadata check (best-effort)
    try {
        $Info = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($Path)
        if ($Info.ProductName -and $Info.ProductName -match '(?i)exiftool') {
            # If ProductName says ExifTool but -ver failed, treat as not runnable
            return $Null
        }
    }
    catch {}

    return $Null
}

function Expand-ExeCandidates([string[]]$Patterns) {
    $Out = New-Object System.Collections.Generic.List[string]

    foreach ($Pattern in $Patterns) {
        try {
            if ($Pattern -match '[\*\?]') {
                Get-ChildItem -Path $Pattern -ErrorAction SilentlyContinue | ForEach-Object { $Out.Add($_.FullName) }
            }
            else {
                if (Test-Path -LiteralPath $Pattern) { $Out.Add($Pattern) }
            }
        }
        catch {}
    }

    $Out | Select-Object -Unique
}


function Get-ExifToolSearchInfo {
    <#
      Single source of truth for:
      - Candidate path patterns to scan (Resolve-ExifToolPath)
      - Good starting directories for file picker (Set-ExifToolPath)
    #>

    $ChocoTools = "C:\ProgramData\chocolatey\lib\exiftool\tools"
    $WgLinks = Join-Path $Env:LOCALAPPDATA "Microsoft\WinGet\Links"
    $WgPackages = Join-Path $Env:LOCALAPPDATA "Microsoft\WinGet\Packages"

    $User_Downloads_Folder = Join-Path $Env:USERPROFILE "Downloads"
    $User_Documents_Folder = Join-Path $Env:USERPROFILE "Documents"
    $User_Desktop_Folder = Join-Path $Env:USERPROFILE "Desktop"
    $User_Pictures_Folder = Join-Path $Env:USERPROFILE "Pictures"

    # Candidate PATH PATTERNS (used for scanning)
    $CandidatePatterns = @(
        # Chocolatey
        "C:\ProgramData\chocolatey\bin\exiftool.exe",
        "C:\ProgramData\chocolatey\lib\exiftool\tools\exiftool.exe",
        "C:\ProgramData\chocolatey\lib\exiftool\tools\exiftool-*\exiftool.exe",

        # Winget links + packages
        (Join-Path $WgLinks    "exiftool.exe"),
        (Join-Path $WgPackages "*\*\exiftool.exe"),
        (Join-Path $WgPackages "*\exiftool.exe"),
        (Join-Path $WgPackages "*\*\*exiftool*.exe"),
        (Join-Path $WgPackages "*\*exiftool*.exe"),

        # User folders (manual installs)
        (Join-Path $User_Downloads_Folder  "exiftool.exe"),
        (Join-Path $User_Documents_Folder "exiftool.exe"),
        (Join-Path $User_Desktop_Folder "exiftool.exe"),
        (Join-Path $User_Downloads_Folder  "exiftool-*\exiftool.exe"),
        (Join-Path $User_Documents_Folder "exiftool-*\exiftool.exe"),
        (Join-Path $User_Desktop_Folder "exiftool-*\exiftool.exe"),
        (Join-Path $User_Downloads_Folder  "*exiftool*\exiftool.exe"),
        (Join-Path $User_Documents_Folder "*exiftool*\exiftool.exe"),
        (Join-Path $User_Desktop_Folder "*exiftool*\exiftool.exe"),
        (Join-Path $User_Pictures_Folder "*exiftool*\exiftool.exe")
    )

    # Preferred INITIAL DIRECTORIES (used for OpenFileDialog starting folder)
    $InitialDirectories = @( $WgLinks, $WgPackages, $ChocoTools, $User_Downloads_Folder, $User_Documents_Folder, $User_Desktop_Folder, $User_Pictures_Folder, $Env:USERPROFILE )

    return [pscustomobject]@{
        ChocoTools         = $ChocoTools
        WgLinks            = $WgLinks
        WgPackages         = $WgPackages
        Downloads          = $User_Downloads_Folder
        Documents          = $User_Documents_Folder
        Desktop            = $User_Desktop_Folder
        CandidatePatterns  = $CandidatePatterns
        InitialDirectories = $InitialDirectories
    }
}

function Resolve-ExifToolPath {
    param( [string]$PreferredPath = $Null, [switch]$OfferWingetInstall, [switch]$OpenDownloadIfMissing )

    $Info = Get-ExifToolSearchInfo

    # 1) Preferred path
    if ($PreferredPath) {
        $Ver = Test-ExifToolRunnable -Path $PreferredPath
        if ($Ver) { return [pscustomobject]@{ Path = $PreferredPath; Version = $Ver } }
    }

    # 2) In PATH
    $Cmd = Get-Command exiftool.exe -ErrorAction SilentlyContinue
    if ($Cmd -and $Cmd.Source) {
        $Ver = Test-ExifToolRunnable -Path $Cmd.Source
        if ($Ver) { return [pscustomobject]@{ Path = $Cmd.Source; Version = $Ver } }
    }

    # 3) Scan all known candidate patterns (single source of truth)
    foreach ($P in (Expand-ExeCandidates $Info.CandidatePatterns)) {
        $Ver = Test-ExifToolRunnable -Path $P
        if ($Ver) { return [pscustomobject]@{ Path = $P; Version = $Ver } }
    }

    # 4) Ask user to locate exiftool.exe
    $Ofd = New-Object System.Windows.Forms.OpenFileDialog
    $Ofd.Title = "Locate exiftool.exe"
    $Ofd.Filter = "ExifTool (exiftool.exe)|exiftool.exe|Executable (*.exe)|*.exe|All files (*.*)|*.*"
    $Ofd.FileName = "exiftool.exe"
    $Ofd.CheckFileExists = $True
    $Ofd.CheckPathExists = $True

    # Choose best initial directory from same sources as resolver
    $Ofd.InitialDirectory = $Env:USERPROFILE
    foreach ($Dir in $Info.InitialDirectories) {
        if ($Dir -and (Test-Path -LiteralPath $Dir)) { $Ofd.InitialDirectory = $Dir; break }
    }

    if ($Ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Picked = $Ofd.FileName
        $Ver = Test-ExifToolRunnable -Path $Picked
        if ($Ver) { return [pscustomobject]@{ Path = $Picked; Version = $Ver } }

        [System.Windows.Forms.MessageBox]::Show(
            "That file doesn't seem to be ExifTool or it won't run.",
            "Invalid selection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }

    # 5) Offer winget install (optional) — unchanged logic
    if ($OfferWingetInstall) {
        $Winget = Get-Command winget.exe -ErrorAction SilentlyContinue
        if ($Winget) {
            $Choice = [System.Windows.Forms.MessageBox]::Show(
                "ExifTool was not found.`r`n`r`nInstall ExifTool using winget now?",
                "Install ExifTool",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )

            if ($Choice -eq [System.Windows.Forms.DialogResult]::Yes) {
                try { & $Winget.Source update 2>$Null | Out-Null } catch {}

                $Search = ""
                try { $Search = (& $Winget.Path search exiftool 2>$Null | Out-String) } catch {}

                if ($Search -match "ExifTool" -or $Search -match "Harvey" -or $Search -match "Phil") {
                    try { & $Winget.Path install -e --id "PhilHarvey.ExifTool" 2>$Null | Out-Null } catch {}
                    if ($LASTEXITCODE -ne 0) { try { & $Winget.Path install "ExifTool" 2>$Null | Out-Null } catch {} }

                    Start-Sleep -Milliseconds 300
                    $Cmd2 = Get-Command exiftool.exe -ErrorAction SilentlyContinue
                    if ($Cmd2 -and $Cmd2.Source) {
                        $Ver2 = Test-ExifToolRunnable -Path $Cmd2.Source
                        if ($Ver2) { return [pscustomobject]@{ Path = $Cmd2.Source; Version = $Ver2 } }
                    }
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Winget couldn't find an ExifTool package in your sources.",
                        "Winget: package not found",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    ) | Out-Null
                }
            }
        }
    }

    # 6) Offer official download page (optional)
    if ($OpenDownloadIfMissing) {
        $Choice = [System.Windows.Forms.MessageBox]::Show( "ExifTool was not found on this system.`r`n`r`nOpen the official ExifTool download page?", "ExifTool missing", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

        if ($Choice -eq [System.Windows.Forms.DialogResult]::Yes) {
            Start-Process "https://exiftool.org/"
            Start-Process "https://exiftool.org/install.html"
        }
    }

    return $Null
}

function Set-ExifToolPath {
    param(
        [string]$InitialDirectory = $Null,
        [string]$CurrentExifToolPath = $Null
    )

    $Info = Get-ExifToolSearchInfo

    $Ofd = New-Object System.Windows.Forms.OpenFileDialog
    $Ofd.Title = "Select exiftool.exe"
    $Ofd.Filter = "ExifTool (exiftool.exe)|exiftool.exe|Executable (*.exe)|*.exe|All files (*.*)|*.*"
    $Ofd.FileName = "exiftool.exe"
    $Ofd.CheckFileExists = $True
    $Ofd.CheckPathExists = $True

    # Priority for picker start folder:
    # 1) caller-provided InitialDirectory
    # 2) folder of currently-loaded ExifTool
    # 3) first existing directory from the SAME sources as resolver
    # 4) user profile
    if ($InitialDirectory -and (Test-Path -LiteralPath $InitialDirectory)) {
        $Ofd.InitialDirectory = $InitialDirectory
    }
    elseif ($CurrentExifToolPath -and (Test-Path -LiteralPath $CurrentExifToolPath)) {
        $Ofd.InitialDirectory = Split-Path -Parent $CurrentExifToolPath
    }
    else {
        $Ofd.InitialDirectory = $Env:USERPROFILE
        foreach ($Dir in $Info.InitialDirectories) {
            if ($Dir -and (Test-Path -LiteralPath $Dir)) { $Ofd.InitialDirectory = $Dir; break }
        }
    }

    if ($Ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $Null }

    $Picked = $Ofd.FileName
    $Ver = Test-ExifToolRunnable -Path $Picked
    if (-not $Ver) {
        [System.Windows.Forms.MessageBox]::Show("That file doesn't seem to be ExifTool or it won't run.", "Invalid selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return $Null
    }

    return [pscustomobject]@{ Path = $Picked; Version = $Ver }
}

function Update-ExifLabel {
    if ($Script:Exif -and $Script:Exif.Path) {
        $Label_Chosen_ExifTool.Text = "ExifTool: $($Script:Exif.Path) - (v$($Script:Exif.Version))"
    }
    else {
        $Label_Chosen_ExifTool.Text = "ExifTool: (not loaded)"
    }
}

# This is a helper function to avoid a possible bug where trailing slashes duplicate.
function Watch-Valid-Folders {
    $InputPath = $Text_InputFolder.Text.Trim()
    $OutputPath = $Text_OutputFolder.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($InputPath) -or -not (Test-Path -LiteralPath $InputPath)) { return $False }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) { return $False }
    if ($InputPath.TrimEnd('\') -ieq $OutputPath.TrimEnd('\')) { return $False }
    return $True
}

# PowerShell doesn't always update the UI cleanly, so we need to refresh the UI.
function Reset-UIState {
    $HasExif = $False
    if ($Script:Exif -and $Script:Exif.Path) { $HasExif = [bool](Test-ExifToolRunnable -Path $Script:Exif.Path) }

    $Button_Remember_Exif_Location.Enabled = $HasExif
    $Button_Start_Process.Enabled = ($HasExif -and (Watch-Valid-Folders))

    Update-ExifLabel
}

# With this function, we don't have to rewrite the UI per new option that we add. This makes the UI drawing dynamic for that section.
function Add-OptionRow {
    param( [System.Windows.Forms.TableLayoutPanel]$Panel, [System.Windows.Forms.Control[]]$Controls )

    $Row = $Panel.RowCount
    $Panel.RowCount++
    $Panel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle 'AutoSize')) | Out-Null

    for ($Col = 0; $Col -lt $Controls.Count; $Col++) { if ($Controls[$Col]) { $Panel.Controls.Add($Controls[$Col], $Col, $Row) | Out-Null } }
}

# ---------- UI ----------
# --- Form ---
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Bulk EXIF/Metadata Stripper (ExifTool)"
$Form.StartPosition = "CenterScreen"
$Form.Size = New-Object System.Drawing.Size(920, 610)
$Form.MinimumSize = New-Object System.Drawing.Size(920, 700)

# --- TableLayoutPanel ---
$Table_Input_Output = New-Object System.Windows.Forms.TableLayoutPanel
$Table_Input_Output.Dock = 'Top'
$Table_Input_Output.AutoSize = $True
$Table_Input_Output.AutoSizeMode = 'GrowAndShrink'
$Table_Input_Output.ColumnCount = 3
$Table_Input_Output.RowCount = 2

# Column styles
$Table_Input_Output.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle 'AutoSize')) | Out-Null
$Table_Input_Output.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle 'Percent', 100)) | Out-Null
$Table_Input_Output.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle 'Absolute', 160)) | Out-Null

# --- Input row ---
$Label_Input_Folder = New-Object System.Windows.Forms.Label
$Label_Input_Folder.Text = "Input folder:"
$Label_Input_Folder.AutoSize = $True
$Label_Input_Folder.Anchor = 'Left'

$Text_InputFolder = New-Object System.Windows.Forms.TextBox
$Text_InputFolder.Dock = 'Fill'

$Button_Input_Folder_Choose = New-Object System.Windows.Forms.Button
$Button_Input_Folder_Choose.Text = "Browse…"
$Button_Input_Folder_Choose.Dock = 'Fill'
$Button_Input_Folder_Choose.Cursor = [System.Windows.Forms.Cursors]::Hand

# --- Output row ---
$Label_Output_Folder = New-Object System.Windows.Forms.Label
$Label_Output_Folder.Text = "Output folder:"
$Label_Output_Folder.AutoSize = $True
$Label_Output_Folder.Anchor = 'Left'

$Text_OutputFolder = New-Object System.Windows.Forms.TextBox
$Text_OutputFolder.Dock = 'Fill'

$Button_Output_Folder_Choose = New-Object System.Windows.Forms.Button
$Button_Output_Folder_Choose.Text = "Browse…"
$Button_Output_Folder_Choose.Dock = 'Fill'
$Button_Output_Folder_Choose.Cursor = [System.Windows.Forms.Cursors]::Hand

# --- Input & output adding to the table_input_output ---
$Table_Input_Output.Controls.Add($Label_Input_Folder, 0, 0) | Out-Null
$Table_Input_Output.Controls.Add($Text_InputFolder, 1, 0) | Out-Null
$Table_Input_Output.Controls.Add($Button_Input_Folder_Choose, 2, 0) | Out-Null

$Table_Input_Output.Controls.Add($Label_Output_Folder, 0, 1) | Out-Null
$Table_Input_Output.Controls.Add($Text_OutputFolder, 1, 1) | Out-Null
$Table_Input_Output.Controls.Add($Button_Output_Folder_Choose, 2, 1) | Out-Null

# --- Options TableLayoutPanel (under the first table) ---
$OptTable = New-Object System.Windows.Forms.TableLayoutPanel
$OptTable.Dock = 'Top'
$OptTable.Padding = '10,0,10,10'
$OptTable.AutoSize = $True
$OptTable.AutoSizeMode = 'GrowAndShrink'
$OptTable.ColumnCount = 1
$OptTable.RowCount = 0

$OptTable.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle 'Percent', 100) ) | Out-Null

$Label_Options_Section = New-Object System.Windows.Forms.Label
$Label_Options_Section.Text = "Options"
$Label_Options_Section.AutoSize = $True
$Label_Options_Section.Font = New-Object System.Drawing.Font( $Form.Font, [System.Drawing.FontStyle]::Bold )
$Label_Options_Section.Margin = '0,8,0,6'

$Checkbox_Include_Subfolders = New-Object System.Windows.Forms.CheckBox
$Checkbox_Include_Subfolders.Text = "Include subfolders (recursive)"
$Checkbox_Include_Subfolders.AutoSize = $True
$Checkbox_Include_Subfolders.Checked = $True
$Checkbox_Include_Subfolders.Cursor = [System.Windows.Forms.Cursors]::Hand

$Checkbox_Remove_Thumbs = New-Object System.Windows.Forms.CheckBox
$Checkbox_Remove_Thumbs.Text = "Remove embedded previews/thumbnails too"
$Checkbox_Remove_Thumbs.AutoSize = $True
$Checkbox_Remove_Thumbs.Checked = $True
$Checkbox_Remove_Thumbs.Cursor = [System.Windows.Forms.Cursors]::Hand

$Checkbox_Overwrite_Filenames = New-Object System.Windows.Forms.CheckBox
$Checkbox_Overwrite_Filenames.Text = "Overwrite existing filenames"
$Checkbox_Overwrite_Filenames.AutoSize = $True
$Checkbox_Overwrite_Filenames.Checked = $False
$Checkbox_Overwrite_Filenames.Cursor = [System.Windows.Forms.Cursors]::Hand

$Checkbox_Open_Output_Folder = New-Object System.Windows.Forms.CheckBox
$Checkbox_Open_Output_Folder.Text = "Open output folder after conversion"
$Checkbox_Open_Output_Folder.AutoSize = $True
$Checkbox_Open_Output_Folder.Checked = $False
$Checkbox_Open_Output_Folder.Cursor = [System.Windows.Forms.Cursors]::Hand

Add-OptionRow -Panel $OptTable -Controls @($Label_Options_Section)
Add-OptionRow -Panel $OptTable -Controls @($Checkbox_Include_Subfolders)
Add-OptionRow -Panel $OptTable -Controls @($Checkbox_Remove_Thumbs)
Add-OptionRow -Panel $OptTable -Controls @($Checkbox_Overwrite_Filenames)
Add-OptionRow -Panel $OptTable -Controls @($Checkbox_Open_Output_Folder)

# --- ExifTool Row Panel (FIX: prevents it being hidden behind Dock panels) ---
$PanelExif = New-Object System.Windows.Forms.TableLayoutPanel
$PanelExif.Dock = 'Top'
$PanelExif.AutoSize = $True
$PanelExif.AutoSizeMode = 'GrowAndShrink'
$PanelExif.ColumnCount = 3

$PanelExif.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle 'Percent', 100)) | Out-Null
$PanelExif.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle 'Absolute', 160)) | Out-Null
$PanelExif.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle 'Absolute', 160)) | Out-Null

$Label_Chosen_ExifTool = New-Object System.Windows.Forms.Label
$Label_Chosen_ExifTool.Text = "ExifTool: (not loaded)"
$Label_Chosen_ExifTool.Size = New-Object System.Drawing.Size(710, 18)
$Label_Chosen_ExifTool.Anchor = "Top,Left,Right"
$Label_Chosen_ExifTool.Dock = 'Fill'

$Button_Select_ExifTool = New-Object System.Windows.Forms.Button
$Button_Select_ExifTool.Text = "Select ExifTool…"
$Button_Select_ExifTool.Size = New-Object System.Drawing.Size(160, 28)
$Button_Select_ExifTool.Anchor = "Top,Right"
$Button_Select_ExifTool.Cursor = [System.Windows.Forms.Cursors]::Hand
$Button_Select_ExifTool.Dock = 'Fill'

$Button_Remember_Exif_Location = New-Object System.Windows.Forms.Button
$Button_Remember_Exif_Location.Text = "Remember"
$Button_Remember_Exif_Location.Size = New-Object System.Drawing.Size(160, 28)
$Button_Remember_Exif_Location.Anchor = "Top,Right"
$Button_Remember_Exif_Location.Cursor = [System.Windows.Forms.Cursors]::Hand
$Button_Remember_Exif_Location.Dock = 'Fill'

$PanelExif.Controls.Add($Label_Chosen_ExifTool, 0, 0)
$PanelExif.Controls.Add($Button_Select_ExifTool, 1, 0)
$PanelExif.Controls.Add($Button_Remember_Exif_Location, 2, 0)

$PanelAction = New-Object System.Windows.Forms.TableLayoutPanel
$PanelAction.Dock = 'Top'
$PanelAction.AutoSize = $True
$PanelAction.AutoSizeMode = 'GrowAndShrink'
$PanelAction.ColumnCount = 2
$PanelAction.RowCount = 1

$PanelAction.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$PanelAction.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 170))) | Out-Null

$Progress_Bar_Process = New-Object System.Windows.Forms.ProgressBar
$Progress_Bar_Process.Dock = 'Fill'
$Progress_Bar_Process.Minimum = 0
$Progress_Bar_Process.Maximum = 100
$Progress_Bar_Process.Value = 0

$Button_Start_Process = New-Object System.Windows.Forms.Button
$Button_Start_Process.Text = "GO"
$Button_Start_Process.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$Button_Start_Process.Dock = 'Fill'
$Button_Start_Process.Cursor = [System.Windows.Forms.Cursors]::Hand

$PanelAction.Controls.Add($Progress_Bar_Process, 0, 0) | Out-Null
$PanelAction.Controls.Add($Button_Start_Process, 1, 0) | Out-Null

# ---------------------------
# STATUS + LOG PANEL
# ---------------------------
$PanelLog = New-Object System.Windows.Forms.TableLayoutPanel
$PanelLog.Dock = 'Fill'
$PanelLog.ColumnCount = 1
$PanelLog.RowCount = 2

$PanelLog.RowStyles.Add(( New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$PanelLog.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

# Status label (THIS was missing)
$Label_Status = New-Object System.Windows.Forms.Label
$Label_Status.Text = "Status: idle"
$Label_Status.AutoSize = $True
$Label_Status.Dock = 'Fill'
$Label_Status.TextAlign = 'MiddleLeft'

# Log textbox
$Logbox_Log = New-Object System.Windows.Forms.TextBox
$Logbox_Log.Dock = 'Fill'
$Logbox_Log.Multiline = $True
$Logbox_Log.ScrollBars = 'Vertical'
$Logbox_Log.ReadOnly = $True
$Logbox_Log.Font = New-Object System.Drawing.Font("Consolas", 9)

$PanelLog.Controls.Add($Label_Status, 0, 0)
$PanelLog.Controls.Add($Logbox_Log, 0, 1)

$PanelFooter = New-Object System.Windows.Forms.TableLayoutPanel
$PanelFooter.Dock = 'Bottom'
$PanelFooter.AutoSize = $True
$PanelFooter.AutoSizeMode = 'GrowAndShrink'
$PanelFooter.ColumnCount = 3
$PanelFooter.RowCount = 2

$PanelFooter.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
$PanelFooter.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$PanelFooter.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null

$PanelFooter.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
$PanelFooter.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

$Label_LinkGithub = New-Object System.Windows.Forms.LinkLabel
$Label_LinkGithub.Text = "Bulk Image Metadata Removal – GitHub - Developed by NekoJonez - 25/12/2025 - Build 0.2 BETA"
$Label_LinkGithub.AutoSize = $True
$Label_LinkGithub.LinkColor = [System.Drawing.Color]::DodgerBlue
$Label_LinkGithub.ActiveLinkColor = [System.Drawing.Color]::RoyalBlue
$Label_LinkGithub.VisitedLinkColor = [System.Drawing.Color]::Purple
$Label_LinkGithub.Add_LinkClicked({ Start-Process "https://github.com/NekoJonez/Bulk-Image-Metadata-Removal/" })

$Button_Exit_Tool = New-Object System.Windows.Forms.Button
$Button_Exit_Tool.Text = "Exit"
$Button_Exit_Tool.AutoSize = $True
$Button_Exit_Tool.Cursor = [System.Windows.Forms.Cursors]::Hand
$Button_Exit_Tool.Add_Click({ $Form.Close() })

# Center them by putting in the middle column
$PanelFooter.Controls.Add($Label_LinkGithub, 1, 0) | Out-Null
$PanelFooter.Controls.Add($Button_Exit_Tool, 1, 1) | Out-Null

# --- FIX DOCK ORDER (prevents Top panels overlapping the Fill log) ---
$Form.SuspendLayout()
$Form.Controls.Clear()
$Form.Controls.Add($PanelLog)      # Dock = Fill
$Form.Controls.Add($PanelFooter)   # Dock = Bottom
$Form.Controls.Add($PanelAction)        # Dock = Top (progress + GO)
$Form.Controls.Add($PanelExif)          # Dock = Top
$Form.Controls.Add($OptTable)           # Dock = Top
$Form.Controls.Add($Table_Input_Output) # Dock = Top (should be the very top)

$Form.ResumeLayout($True)

# Tooltips
$Tooltip_Display = New-Object System.Windows.Forms.ToolTip
$Tooltip_Display.AutoPopDelay = 12000
$Tooltip_Display.InitialDelay = 250
$Tooltip_Display.ReshowDelay = 150
$Tooltip_Display.ShowAlways = $True

$Tooltip_Display.SetToolTip($Button_Input_Folder_Choose, "Pick the folder that contains the images you want to clean.")
$Tooltip_Display.SetToolTip($Button_Output_Folder_Choose, "Pick where the cleaned copies should be written.")
$Tooltip_Display.SetToolTip($Checkbox_Include_Subfolders, "If enabled, includes images in subfolders too.")
$Tooltip_Display.SetToolTip($Checkbox_Remove_Thumbs, "If enabled, removes embedded preview images/thumbnails if present.")
$Tooltip_Display.SetToolTip($Checkbox_Overwrite_Filenames, "If enabled, this overwrites the filenames of files existing in the output folder.")
$Tooltip_Display.SetToolTip($Checkbox_Open_Output_Folder, "If enabled, the output folder will be opened after conversion.")
$Tooltip_Display.SetToolTip($Button_Select_ExifTool, "Select a different exiftool.exe.")
$Tooltip_Display.SetToolTip($Button_Remember_Exif_Location, "Save the ExifTool path so you don’t need to re-select it next run.")
$Tooltip_Display.SetToolTip($Button_Start_Process, "Copy images to Output and strip all metadata from the copies.")
$Tooltip_Display.SetToolTip($Button_Exit_Tool, "This will exit the tool.")

$FolderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderDlg.ShowNewFolderButton = $True

# Script-scoped state
$Script:Exif = $Null

$Text_InputFolder.Add_TextChanged({ Reset-UIState })
$Text_OutputFolder.Add_TextChanged({ Reset-UIState })

$Button_Input_Folder_Choose.Add_Click({
        $FolderDlg.Description = "Select INPUT folder containing images"
        if ($FolderDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $Text_InputFolder.Text = $FolderDlg.SelectedPath }
    })

$Button_Output_Folder_Choose.Add_Click({
        $FolderDlg.Description = "Select OUTPUT folder for cleaned images"
        if ($FolderDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $Text_OutputFolder.Text = $FolderDlg.SelectedPath }
    })

# Load config preferred path, then resolve exiftool
$ConfigFile = Import-Config
$Preferred = $Null
if ($ConfigFile -and $ConfigFile.ExifToolPath) { $Preferred = [string]$ConfigFile.ExifToolPath }

$Script:Exif = Resolve-ExifToolPath -PreferredPath $Preferred -OfferWingetInstall -OpenDownloadIfMissing
Reset-UIState

$Button_Select_ExifTool.Add_Click({
        # If ExifTool is already loaded and runnable, confirm before changing it
        $HasValidExif = $False
        if ($Script:Exif -and $Script:Exif.Path) { $HasValidExif = [bool](Test-ExifToolRunnable -Path $Script:Exif.Path) }

        if ($HasValidExif) {
            $Choice = [System.Windows.Forms.MessageBox]::Show( $Form, "ExifTool is already loaded:`r`n$($Script:Exif.Path)`r`n`r`nAre you sure you want to change the tool?", "Confirm change", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($Choice -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }

        # Force picker; it will start in the best directory automatically:
        # - InitialDirectory (if you pass it)
        # - else folder of current ExifTool
        # - else common resolver locations
        $Picked = Set-ExifToolPath -CurrentExifToolPath ($Script:Exif.Path)
        if ($Picked) {
            $Script:Exif = $Picked
            Reset-UIState
        }
    })

$Button_Remember_Exif_Location.Add_Click({
        if ($Script:Exif -and $Script:Exif.Path) {
            Save-Config ([pscustomobject]@{ ExifToolPath = $Script:Exif.Path })
            [System.Windows.Forms.MessageBox]::Show($Form, "Saved ExifTool path for next time.", "Saved", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    })

$Button_Start_Process.Add_Click({
        if (!($Script:Exif -and $Script:Exif.Path)) { return }
        if (-not (Watch-Valid-Folders)) { return }

        $InputPath = $Text_InputFolder.Text.Trim()
        $OutputPath = $Text_OutputFolder.Text.Trim()

        # Disable controls during run
        $Button_Start_Process.Enabled = $False
        $Button_Input_Folder_Choose.Enabled = $False
        $Button_Output_Folder_Choose.Enabled = $False
        $Button_Select_ExifTool.Enabled = $False
        $Button_Remember_Exif_Location.Enabled = $False
        $Button_Exit_Tool.Enabled = $False
        $Checkbox_Include_Subfolders.Enabled = $False
        $Checkbox_Remove_Thumbs.Enabled = $False
        $Checkbox_Overwrite_Filenames.Enabled = $False
        $Checkbox_Open_Output_Folder.Enabled = $False

        try {
            $Logbox_Log.Clear()
            $Progress_Bar_Process.Value = 0
            $Label_Status.Text = "Status: scanning…"

            Add-To-Log $Logbox_Log "ExifTool: $($Script:Exif.Path) (v$($Script:Exif.Version))"
            Add-To-Log $Logbox_Log "Input:   $InputPath"
            Add-To-Log $Logbox_Log "Output:  $OutputPath"
            Add-To-Log $Logbox_Log ("Include subfolders: " + $Checkbox_Include_Subfolders.Checked)
            Add-To-Log $Logbox_Log ("Remove thumbs: " + $Checkbox_Remove_Thumbs.Checked)
            Add-To-Log $Logbox_Log ("Overwriting files in output folder: " + $Checkbox_Overwrite_Filenames.Checked)
            Add-To-Log $Logbox_Log ("Remove thumbs: " + $Checkbox_Remove_Thumbs.Checked)

            Set-Ensure-Dir $OutputPath

            $Gci = @{ LiteralPath = $InputPath; File = $True; ErrorAction = 'SilentlyContinue' }
            if ($Checkbox_Include_Subfolders.Checked) { $Gci.Recurse = $True }

            $Files = Get-ChildItem @Gci | Where-Object { $AllowedExtensions -contains $_.Extension.ToLowerInvariant() }
            $TotalAmountOfFiles = $Files.Count

            Add-To-Log $Logbox_Log "Found $TotalAmountOfFiles image(s)."
            if ($TotalAmountOfFiles -eq 0) {
                $Label_Status.Text = "Status: nothing to do"
                return
            }

            $IterationCounter = 0
            foreach ($File in $Files) {
                $IterationCounter++
                $Label_Status.Text = "Status: processing $IterationCounter / $TotalAmountOfFiles"
                $Progress_Bar_Process.Value = [Math]::Min(100, [int](($IterationCounter / $TotalAmountOfFiles) * 100))

                $RelativePath = Get-RelativePath -Base $InputPath -Full $File.FullName
                $DestinationFolder = Join-Path $OutputPath $RelativePath
                Set-Ensure-Dir (Split-Path -Parent $DestinationFolder)

                $DestExists = Test-Path -LiteralPath $DestinationFolder

                if ($DestExists -and -not $Checkbox_Overwrite_Filenames.Checked) {
                    Add-To-Log $Logbox_Log "SKIP: Can't convert, file with same name exists in the output folder: $RelativePath"
                    continue
                }

                if ($DestExists -and $Checkbox_Overwrite_Filenames.Checked) {
                    Add-To-Log $Logbox_Log "OVERWRITE: $RelativePath"
                    Copy-Item -LiteralPath $File.FullName -Destination $DestinationFolder -Force
                }
                else {
                    Copy-Item -LiteralPath $File.FullName -Destination $DestinationFolder
                }

                # Strip metadata on copy
                $ArgsForTool = @("-all=", "-overwrite_original", "-q", "-q")
                if ($Checkbox_Remove_Thumbs.Checked) { $ArgsForTool += @("-ThumbnailImage=", "-PreviewImage=") }
                $QuotedDest = '"' + $DestinationFolder + '"'
                $ArgsForTool += @("--", $QuotedDest)

                $Proc = Start-Process -FilePath $Script:Exif.Path -ArgumentList $ArgsForTool -NoNewWindow -Wait -PassThru
                if ($Proc.ExitCode -eq 0) {
                    Add-To-Log $Logbox_Log "OK: $RelativePath"
                }
                else {
                    Add-To-Log $Logbox_Log "FAIL ($($Proc.ExitCode)): $RelativePath"
                }

                [System.Windows.Forms.Application]::DoEvents()
            }

            $Label_Status.Text = "Status: done"
            $Progress_Bar_Process.Value = 100
            Add-To-Log $Logbox_Log "Done. Clean images written to: $OutputPath"

            [System.Windows.Forms.MessageBox]::Show($Form, "Finished stripping metadata.`r`nCheck the output folder.", "Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null

            if ($Checkbox_Open_Output_Folder.Checked -and (Test-Path -LiteralPath $OutputPath)) { Start-Process explorer.exe $OutputPath }
        }
        catch {
            $Detail = ($_ | Format-List * -Force | Out-String)
            Add-To-Log $Logbox_Log "ERROR: $($_.Exception.Message)"
            Add-To-Log $Logbox_Log $Detail

            [System.Windows.Forms.MessageBox]::Show($Form, $Detail, "Error details", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
        finally {
            # Re-enable
            $Button_Input_Folder_Choose.Enabled = $True
            $Button_Output_Folder_Choose.Enabled = $True
            $Button_Select_ExifTool.Enabled = $True
            $Button_Remember_Exif_Location.Enabled = $True
            $Button_Exit_Tool.Enabled = $True
            $Checkbox_Include_Subfolders.Enabled = $True
            $Checkbox_Remove_Thumbs.Enabled = $True
            $Checkbox_Overwrite_Filenames.Enabled = $True
            $Checkbox_Open_Output_Folder.Enabled = $True
            Reset-UIState
        }
    })

[void]$Form.ShowDialog()