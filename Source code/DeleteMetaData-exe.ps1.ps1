##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uK1
##Nc3NCtDXThU=
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4NI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+Vs1Q=
##M9jHFoeYB2Hc8u+Vs1Q=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWI0g==
##OsfOAYaPHGbQvbyVvnQnqxqO
##LNzNAIWJGmPcoKHc7Do3uAu+DDlL
##LNzNAIWJGnvYv7eVvnRV8EflAk0TR+PWlLOzxYW96+usiDHLQIoETEYX
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnRA7F/mAnwjZ9bbrbmoyM385uX5+ybYR9oGTFl4gCy8CU6pWPYTQfRVssMeWF0rI/5L67/RHOK6BasE0uI/beqCorc7VVPWo5Hh3xLy
##P8HPFJGEFzWE8pzQ6Dtd6kXrWAg=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+VsWQlsAujAkYuZcqxtrii3uE=
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlaDjofG5iZk2U/rTG0yUuGeqr2zy5GA3f/gsGX8cbMyBHlygij4BV+8GcEGR/wFpNQDGxgyKpI=
##Kc/BRM3KXxU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
<#
DeleteMetaData-UI.ps1 (Upgraded: Remember + Tooltips + Loaded Display)
By NekoJonez - v0.2 BETA
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- CONFIG ----------
$AllowedExtensions = @(".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".heic", ".heif", ".gif", ".bmp", ".jfif", ".avif")

# ---------- Config persistence ----------
$AppName = "ExifToolUI"
$ConfigDir = Join-Path $Env:APPDATA $AppName
$ConfigPath = Join-Path $ConfigDir "config.json"

function Import-Config {
    if (Test-Path -LiteralPath $ConfigPath) { try { return (Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json) } catch {} }
    return $Null
}

function Save-Config($Obj) {
    if (!(Test-Path -LiteralPath $ConfigDir)) { New-Item -ItemType Directory -Path $ConfigDir -Force | Out-Null }
    ($Obj | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $ConfigPath -Encoding UTF8
}

# ---------- Helpers & functions ----------
# ! TODO: comments
function Set-Ensure-Dir([string]$Path) {
    if (!(Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Get-RelativePath([string]$Base, [string]$Full) {
    $BaseTrim = $Base.TrimEnd('\')
    if ($Full.Length -le $BaseTrim.Length) { return [IO.Path]::GetFileName($Full) }
    return $Full.Substring($BaseTrim.Length).TrimStart('\')
}

function Add-To-Log([System.Windows.Forms.TextBox]$Box, [string]$Text) {
    $Ts = (Get-Date).ToString("HH:mm:ss")
    $Box.AppendText("[$Ts] $Text`r`n")
    $Box.SelectionStart = $Box.TextLength
    $Box.ScrollToCaret()
}

function Test-ExifToolRunnable {
    param([Parameter(Mandatory)][string]$Path)

    if (!(Test-Path -LiteralPath $Path)) { return $Null }
    try {
        $Ver = & $Path -ver 2>$Null
        if ($Ver) { return $Ver.Trim() }
    }
    catch {}
    return $Null
}

function Expand-ExeCandidates([string[]]$Patterns) {
    $Out = New-Object System.Collections.Generic.List[string]
    foreach ($Pattern in $Patterns) {
        try {
            if ($Pattern -like "*`**") {
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

function Resolve-ExifToolPath {
    param( [string]$PreferredPath = $Null, [switch]$OfferWingetInstall, [switch]$OpenDownloadIfMissing )

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

    # 3) Chocolatey + version subfolders
    $ChocoTools = "C:\ProgramData\chocolatey\lib\exiftool\tools"
    $Candidates = @(
        "C:\ProgramData\chocolatey\bin\exiftool.exe",
        "C:\ProgramData\chocolatey\lib\exiftool\tools\exiftool.exe",
        "C:\ProgramData\chocolatey\lib\exiftool\tools\exiftool-*\exiftool.exe"
    )

    if (Test-Path -LiteralPath $ChocoTools) {
        try {
            $More = Get-ChildItem -LiteralPath $ChocoTools -Recurse -File -Filter "exiftool.exe" -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty FullName
            $Candidates += $More
        }
        catch {}
    }

    foreach ($P in (Expand-ExeCandidates $Candidates)) {
        $Ver = Test-ExifToolRunnable -Path $P
        if ($Ver) { return [pscustomobject]@{ Path = $P; Version = $Ver } }
    }

    # 4) Winget usual spots + links
    $WingetCandidates = @(
        "$Env:LOCALAPPDATA\Microsoft\WinGet\Links\exiftool.exe",
        "$Env:LOCALAPPDATA\Microsoft\WinGet\Packages\*\*\exiftool.exe",
        "$Env:LOCALAPPDATA\Microsoft\WinGet\Packages\*\exiftool.exe",
        "$Env:LOCALAPPDATA\Microsoft\WinGet\Packages\*\*\*exiftool*.exe",
        "$Env:LOCALAPPDATA\Microsoft\WinGet\Packages\*\*exiftool*.exe"
    )

    foreach ($P in (Expand-ExeCandidates $WingetCandidates)) {
        $Ver = Test-ExifToolRunnable -Path $P
        if ($Ver) { return [pscustomobject]@{ Path = $P; Version = $Ver } }
    }

    # 5) Downloads / Documents / Desktop (common manual extract locations)
    $Dl = Join-Path $Env:USERPROFILE "Downloads"
    $Doc = Join-Path $Env:USERPROFILE "Documents"
    $Dsk = Join-Path $Env:USERPROFILE "Desktop"

    $UserCandidates = @(
        (Join-Path $Dl  "exiftool.exe"),
        (Join-Path $Doc "exiftool.exe"),
        (Join-Path $Dsk "exiftool.exe"),
        (Join-Path $Dl  "exiftool-*\exiftool.exe"),
        (Join-Path $Doc "exiftool-*\exiftool.exe"),
        (Join-Path $Dsk "exiftool-*\exiftool.exe"),
        (Join-Path $Dl  "*exiftool*\exiftool.exe"),
        (Join-Path $Doc "*exiftool*\exiftool.exe"),
        (Join-Path $Dsk "*exiftool*\exiftool.exe")
    )

    foreach ($P in (Expand-ExeCandidates $UserCandidates)) {
        $Ver = Test-ExifToolRunnable -Path $P
        if ($Ver) { return [pscustomobject]@{ Path = $P; Version = $Ver } }
    }

    # 6) Ask user to locate exiftool.exe
    $Ofd = New-Object System.Windows.Forms.OpenFileDialog
    $Ofd.Title = "Locate exiftool.exe"
    $Ofd.Filter = "ExifTool (exiftool.exe)|exiftool.exe|Executable (*.exe)|*.exe|All files (*.*)|*.*"
    $Ofd.FileName = "exiftool.exe"
    $Ofd.CheckFileExists = $True
    $Ofd.CheckPathExists = $True

    $WgLinks = Join-Path $Env:LOCALAPPDATA "Microsoft\WinGet\Links"
    if (Test-Path -LiteralPath $WgLinks) { $Ofd.InitialDirectory = $WgLinks }
    elseif (Test-Path -LiteralPath $ChocoTools) { $Ofd.InitialDirectory = $ChocoTools }
    elseif (Test-Path -LiteralPath $Dl) { $Ofd.InitialDirectory = $Dl }
    else { $Ofd.InitialDirectory = $Env:USERPROFILE }

    if ($Ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Picked = $Ofd.FileName
        $Ver = Test-ExifToolRunnable -Path $Picked
        if ($Ver) { return [pscustomobject]@{ Path = $Picked; Version = $Ver } }

        [System.Windows.Forms.MessageBox]::Show("That file doesn't seem to be ExifTool or it won't run.", "Invalid selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }

    # 7) Offer winget install (optional)
    if ($OfferWingetInstall) {
        $Winget = Get-Command winget.exe -ErrorAction SilentlyContinue
        if ($Winget) {
            $Choice = [System.Windows.Forms.MessageBox]::Show("ExifTool was not found.`r`n`r`nInstall ExifTool using winget now?", "Install ExifTool", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

            # TODO: Check here correctly to install since --accept is a thing.
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
                    [System.Windows.Forms.MessageBox]::Show("Winget couldn't find an ExifTool package in your sources.", "Winget: package not found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning ) | Out-Null
                }
            }
        }
    }

    # 8) Offer official download page (optional)
    if ($OpenDownloadIfMissing) {
        $Choice = [System.Windows.Forms.MessageBox]::Show("ExifTool was not found on this system.`r`n`r`nOpen the official ExifTool download page?", "ExifTool missing", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($Choice -eq [System.Windows.Forms.DialogResult]::Yes) {
            Start-Process "https://exiftool.org/"
            Start-Process "https://exiftool.org/install.html"
        }
    }

    return $Null
}

function Set-ExifToolPath {
    param([string]$InitialDirectory = $Null)

    $Ofd = New-Object System.Windows.Forms.OpenFileDialog
    $Ofd.Title = "Select exiftool.exe"
    $Ofd.Filter = "ExifTool (exiftool.exe)|exiftool.exe|Executable (*.exe)|*.exe|All files (*.*)|*.*"
    $Ofd.FileName = "exiftool.exe"
    $Ofd.CheckFileExists = $True
    $Ofd.CheckPathExists = $True

    if ($InitialDirectory -and (Test-Path -LiteralPath $InitialDirectory)) {
        $Ofd.InitialDirectory = $InitialDirectory
    }
    else {
        $WgLinks = Join-Path $Env:LOCALAPPDATA "Microsoft\WinGet\Links"
        $ChocoTools = "C:\ProgramData\chocolatey\lib\exiftool\tools"
        $Dl = Join-Path $Env:USERPROFILE "Downloads"

        if (Test-Path -LiteralPath $WgLinks) { $Ofd.InitialDirectory = $WgLinks }
        elseif (Test-Path -LiteralPath $ChocoTools) { $Ofd.InitialDirectory = $ChocoTools }
        elseif (Test-Path -LiteralPath $Dl) { $Ofd.InitialDirectory = $Dl }
        else { $Ofd.InitialDirectory = $Env:USERPROFILE }
    }

    if ($Ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $Null }

    $Picked = $Ofd.FileName
    $Ver = Test-ExifToolRunnable -Path $Picked
    if (-not $Ver) {
        [System.Windows.Forms.MessageBox]::Show( "That file doesn't seem to be ExifTool or it won't run.", "Invalid selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return $Null
    }

    return [pscustomobject]@{ Path = $Picked; Version = $Ver }
}

function Update-ExifLabel {
    if ($Script:Exif -and $Script:Exif.Path) {
        $LblExif.Text = "ExifTool: $($Script:Exif.Path)  (v$($Script:Exif.Version))"
    }
    else {
        $LblExif.Text = "ExifTool: (not loaded)"
    }
}

function Watch-Valid-Folders {
    $InPath = $TxtIn.Text.Trim()
    $OutPath = $TxtOut.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($InPath) -or -not (Test-Path -LiteralPath $InPath)) { return $False }
    if ([string]::IsNullOrWhiteSpace($OutPath)) { return $False }
    if ($InPath.TrimEnd('\') -ieq $OutPath.TrimEnd('\')) { return $False }
    return $True
}

function Reset-UIState {
    $HasExif = $False
    if ($Script:Exif -and $Script:Exif.Path) { $HasExif = [bool](Test-ExifToolRunnable -Path $Script:Exif.Path) }

    $BtnRemember.Enabled = $HasExif
    $BtnGo.Enabled = ($HasExif -and (Watch-Valid-Folders))

    Update-ExifLabel
}

# ---------- UI ----------
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Bulk EXIF/Metadata Stripper (ExifTool)"
$Form.StartPosition = "CenterScreen"
$Form.Size = New-Object System.Drawing.Size(920, 610)
$Form.MinimumSize = New-Object System.Drawing.Size(920, 700)

$LblIn = New-Object System.Windows.Forms.Label
$LblIn.Text = "Input folder:"
$LblIn.Location = New-Object System.Drawing.Point(12, 16)
$LblIn.AutoSize = $True

$TxtIn = New-Object System.Windows.Forms.TextBox
$TxtIn.Location = New-Object System.Drawing.Point(110, 12)
$TxtIn.Size = New-Object System.Drawing.Size(610, 24)
$TxtIn.Anchor = "Top,Left,Right"

$BtnIn = New-Object System.Windows.Forms.Button
$BtnIn.Text = "Browse…"
$BtnIn.Location = New-Object System.Drawing.Point(735, 10)
$BtnIn.Size = New-Object System.Drawing.Size(160, 28)
$BtnIn.Anchor = "Top,Right"
$BtnIn.Cursor = [System.Windows.Forms.Cursors]::Hand

$LblOut = New-Object System.Windows.Forms.Label
$LblOut.Text = "Output folder:"
$LblOut.Location = New-Object System.Drawing.Point(12, 52)
$LblOut.AutoSize = $True

$TxtOut = New-Object System.Windows.Forms.TextBox
$TxtOut.Location = New-Object System.Drawing.Point(110, 48)
$TxtOut.Size = New-Object System.Drawing.Size(610, 24)
$TxtOut.Anchor = "Top,Left,Right"

$BtnOut = New-Object System.Windows.Forms.Button
$BtnOut.Text = "Browse…"
$BtnOut.Location = New-Object System.Drawing.Point(735, 46)
$BtnOut.Size = New-Object System.Drawing.Size(160, 28)
$BtnOut.Anchor = "Top,Right"
$BtnOut.Cursor = [System.Windows.Forms.Cursors]::Hand

$ChkSub = New-Object System.Windows.Forms.CheckBox
$ChkSub.Text = "Include subfolders (recursive)"
$ChkSub.Location = New-Object System.Drawing.Point(110, 82)
$ChkSub.AutoSize = $True
$ChkSub.Checked = $True

$ChkThumbs = New-Object System.Windows.Forms.CheckBox
$ChkThumbs.Text = "Remove embedded previews/thumbnails too"
$ChkThumbs.Location = New-Object System.Drawing.Point(360, 82)
$ChkThumbs.AutoSize = $True
$ChkThumbs.Checked = $True

$LblExif = New-Object System.Windows.Forms.Label
$LblExif.Text = "ExifTool: (not loaded)"
$LblExif.Location = New-Object System.Drawing.Point(12, 110)
$LblExif.Size = New-Object System.Drawing.Size(710, 18)
$LblExif.Anchor = "Top,Left,Right"

$BtnExif = New-Object System.Windows.Forms.Button
$BtnExif.Text = "Select ExifTool…"
$BtnExif.Location = New-Object System.Drawing.Point(735, 104)
$BtnExif.Size = New-Object System.Drawing.Size(160, 28)
$BtnExif.Anchor = "Top,Right"
$BtnExif.Cursor = [System.Windows.Forms.Cursors]::Hand

$BtnRemember = New-Object System.Windows.Forms.Button
$BtnRemember.Text = "Remember"
$BtnRemember.Location = New-Object System.Drawing.Point(735, 136)
$BtnRemember.Size = New-Object System.Drawing.Size(160, 28)
$BtnRemember.Anchor = "Top,Right"
$BtnRemember.Cursor = [System.Windows.Forms.Cursors]::Hand

$BtnGo = New-Object System.Windows.Forms.Button
$BtnGo.Text = "GO"
$BtnGo.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$BtnGo.Location = New-Object System.Drawing.Point(735, 202)
$BtnGo.Size = New-Object System.Drawing.Size(160, 34)
$BtnGo.Anchor = "Top,Right"
$BtnGo.Cursor = [System.Windows.Forms.Cursors]::Hand

$Progress = New-Object System.Windows.Forms.ProgressBar
$Progress.Location = New-Object System.Drawing.Point(12, 212)
$Progress.Size = New-Object System.Drawing.Size(710, 18)
$Progress.Anchor = "Top,Left,Right"
$Progress.Minimum = 0
$Progress.Maximum = 100
$Progress.Value = 0

$LblStatus = New-Object System.Windows.Forms.Label
$LblStatus.Text = "Status: idle"
$LblStatus.Location = New-Object System.Drawing.Point(12, 238)
$LblStatus.Size = New-Object System.Drawing.Size(883, 18)
$LblStatus.Anchor = "Top,Left,Right"

$TxtLog = New-Object System.Windows.Forms.TextBox
$TxtLog.Location = New-Object System.Drawing.Point(12, 265)
$TxtLog.Size = New-Object System.Drawing.Size(883, 300)
$TxtLog.Anchor = "Top,Bottom,Left,Right"
$TxtLog.Multiline = $True
$TxtLog.ScrollBars = "Vertical"
$TxtLog.ReadOnly = $True
$TxtLog.Font = New-Object System.Drawing.Font("Consolas", 9)

$LinkGithub = New-Object System.Windows.Forms.LinkLabel
$LinkGithub.Text = "Bulk Image Metadata Removal – GitHub - Developed by NekoJonez - 22/12/2025 - Build 0.1 BETA"
$LinkGithub.AutoSize = $True
$LinkGithub.LinkColor = [System.Drawing.Color]::DodgerBlue
$LinkGithub.ActiveLinkColor = [System.Drawing.Color]::RoyalBlue
$LinkGithub.VisitedLinkColor = [System.Drawing.Color]::Purple
$LinkGithub.Location = New-Object System.Drawing.Point(12, 575)
$LinkGithub.Add_LinkClicked({ Start-Process "https://github.com/NekoJonez/Bulk-Image-Metadata-Removal/" })

$BtnExit = New-Object System.Windows.Forms.Button
$BtnExit.Text = "Exit"
$BtnExit.Size = New-Object System.Drawing.Size(120, 30)
$BtnExit.Cursor = [System.Windows.Forms.Cursors]::Hand
$BtnExit.AutoSize = $True
$BtnExit.Location = New-Object System.Drawing.Point(12, 600)
$BtnExit.Add_Click({ $Form.Close() })

$Form.Controls.AddRange(@(
        $LblIn, $TxtIn, $BtnIn,
        $LblOut, $TxtOut, $BtnOut,
        $ChkSub, $ChkThumbs,
        $LblExif, $BtnExif, $BtnRemember,
        $Progress, $BtnGo,
        $LblStatus, $TxtLog,
        $LinkGithub, $BtnExit
    ))

# Tooltips
$Tip = New-Object System.Windows.Forms.ToolTip
$Tip.AutoPopDelay = 12000
$Tip.InitialDelay = 250
$Tip.ReshowDelay = 150
$Tip.ShowAlways = $True

$Tip.SetToolTip($BtnIn, "Pick the folder that contains the images you want to clean.")
$Tip.SetToolTip($BtnOut, "Pick where the cleaned copies should be written.")
$Tip.SetToolTip($ChkSub, "If enabled, includes images in subfolders too.")
$Tip.SetToolTip($ChkThumbs, "Also removes embedded preview images/thumbnails if present.")
$Tip.SetToolTip($BtnExif, "Select a different exiftool.exe.")
$Tip.SetToolTip($BtnRemember, "Save the ExifTool path so you don’t need to re-select it next run.")
$Tip.SetToolTip($BtnGo, "Copy images to Output and strip all metadata from the copies.")
$Tip.SetToolTip($BtnExit, "This will exit the tool.")

$FolderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
$FolderDlg.ShowNewFolderButton = $True

# Script-scoped state
$Script:Exif = $Null

$TxtIn.Add_TextChanged({ Reset-UIState })
$TxtOut.Add_TextChanged({ Reset-UIState })

$BtnIn.Add_Click({
        $FolderDlg.Description = "Select INPUT folder containing images"
        if ($FolderDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $TxtIn.Text = $FolderDlg.SelectedPath }
    })

$BtnOut.Add_Click({
        $FolderDlg.Description = "Select OUTPUT folder for cleaned images"
        if ($FolderDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $TxtOut.Text = $FolderDlg.SelectedPath }
    })

# Load config preferred path, then resolve exiftool
$Cfg = Import-Config
$Preferred = $Null
if ($Cfg -and $Cfg.ExifToolPath) { $Preferred = [string]$Cfg.ExifToolPath }

$Script:Exif = Resolve-ExifToolPath -PreferredPath $Preferred -OfferWingetInstall -OpenDownloadIfMissing
Reset-UIState

$BtnExif.Add_Click({
        # If ExifTool is already loaded and runnable, confirm before changing it
        $HasValidExif = $False
        if ($Script:Exif -and $Script:Exif.Path) {
            $HasValidExif = [bool](Test-ExifToolRunnable -Path $Script:Exif.Path)
        }

        if ($HasValidExif) {
            $Choice = [System.Windows.Forms.MessageBox]::Show($Form, "ExifTool is already loaded:`r`n$($Script:Exif.Path)`r`n`r`nAre you sure you want to change the tool?", "Confirm change", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($Choice -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }

        # Browse for a new exiftool.exe (force picker)
        $Initial = $Null
        if ($Script:Exif -and $Script:Exif.Path) { $Initial = (Split-Path -Parent $Script:Exif.Path) }

        $Picked = Set-ExifToolPath -InitialDirectory $Initial
        if ($Picked) {
            $Script:Exif = $Picked
            Reset-UIState
        }
    })

$BtnRemember.Add_Click({
        if ($Script:Exif -and $Script:Exif.Path) {
            Save-Config ([pscustomobject]@{ ExifToolPath = $Script:Exif.Path })
            [System.Windows.Forms.MessageBox]::Show($Form, "Saved ExifTool path for next time.", "Saved", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
    })

$BtnGo.Add_Click({
        if (!($Script:Exif -and $Script:Exif.Path)) { return }
        if (-not (Watch-Valid-Folders)) { return }

        $InPath = $TxtIn.Text.Trim()
        $OutPath = $TxtOut.Text.Trim()

        # Disable controls during run
        $BtnGo.Enabled = $False
        $BtnIn.Enabled = $False
        $BtnOut.Enabled = $False
        $BtnExif.Enabled = $False
        $BtnRemember.Enabled = $False
        $ChkSub.Enabled = $False
        $ChkThumbs.Enabled = $False

        try {
            $TxtLog.Clear()
            $Progress.Value = 0
            $LblStatus.Text = "Status: scanning…"

            Add-To-Log $TxtLog "ExifTool: $($Script:Exif.Path) (v$($Script:Exif.Version))"
            Add-To-Log $TxtLog "Input:   $InPath"
            Add-To-Log $TxtLog "Output:  $OutPath"
            Add-To-Log $TxtLog ("Recursive: " + $ChkSub.Checked)
            Add-To-Log $TxtLog ("Remove thumbs: " + $ChkThumbs.Checked)

            Set-Ensure-Dir $OutPath

            $Gci = @{ LiteralPath = $InPath; File = $True; ErrorAction = 'SilentlyContinue' }
            if ($ChkSub.Checked) { $Gci.Recurse = $True }

            $Files = Get-ChildItem @Gci | Where-Object { $AllowedExtensions -contains $_.Extension.ToLowerInvariant() }
            $Total = $Files.Count

            Add-To-Log $TxtLog "Found $Total image(s)."
            if ($Total -eq 0) {
                $LblStatus.Text = "Status: nothing to do"
                return
            }

            $I = 0
            foreach ($F in $Files) {
                $I++
                $LblStatus.Text = "Status: processing $I / $Total"
                $Progress.Value = [Math]::Min(100, [int](($I / $Total) * 100))

                $Rel = Get-RelativePath -Base $InPath -Full $F.FullName
                $Dest = Join-Path $OutPath $Rel
                Set-Ensure-Dir (Split-Path -Parent $Dest)

                # Copy first
                Copy-Item -LiteralPath $F.FullName -Destination $Dest -Force

                # Strip metadata on copy
                $ArgsForTool = @("-all=", "-overwrite_original", "-q", "-q")
                if ($ChkThumbs.Checked) { $ArgsForTool += @("-ThumbnailImage=", "-PreviewImage=") }
                $QuotedDest = '"' + $Dest + '"'
                $ArgsForTool += @("--", $QuotedDest)

                $Proc = Start-Process -FilePath $Script:Exif.Path -ArgumentList $ArgsForTool -NoNewWindow -Wait -PassThru
                if ($Proc.ExitCode -eq 0) {
                    Add-To-Log $TxtLog "OK: $Rel"
                }
                else {
                    Add-To-Log $TxtLog "FAIL ($($Proc.ExitCode)): $Rel"
                }

                [System.Windows.Forms.Application]::DoEvents()
            }

            $LblStatus.Text = "Status: done"
            $Progress.Value = 100
            Add-To-Log $TxtLog "Done. Clean images written to: $OutPath"

            [System.Windows.Forms.MessageBox]::Show($Form, "Finished stripping metadata.`r`nCheck the output folder.", "Done", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        catch {
            $Detail = ($_ | Format-List * -Force | Out-String)
            Add-To-Log $TxtLog "ERROR: $($_.Exception.Message)"
            Add-To-Log $TxtLog $Detail

            [System.Windows.Forms.MessageBox]::Show($Form, $Detail, "Error details", [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
        finally {
            # Re-enable
            $BtnIn.Enabled = $True
            $BtnOut.Enabled = $True
            $BtnExif.Enabled = $True
            $ChkSub.Enabled = $True
            $ChkThumbs.Enabled = $True
            Reset-UIState
        }
    })

[void]$Form.ShowDialog()