<#!
.SYNOPSIS
    Lightweight Windows Forms frontend for Graph_api_User_add_remove.ps1.

.DESCRIPTION
    Provides a simple GUI to run View/Add/Remove actions against
    Graph_api_User_add_remove.ps1 without typing command-line arguments.
#>

[CmdletBinding()]
param ()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $IsWindows) {
    throw 'This frontend uses Windows Forms and only runs on Windows PowerShell/Windows hosts.'
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$backendScript = Join-Path -Path $PSScriptRoot -ChildPath 'Graph_api_User_add_remove.ps1'
if (-not (Test-Path -LiteralPath $backendScript)) {
    throw "Could not find backend script: $backendScript"
}

# ---------------------------------------------------------------------------
# Theme constants
# ---------------------------------------------------------------------------
$clrBackground  = [System.Drawing.ColorTranslator]::FromHtml('#1E1E2E')
$clrPanel       = [System.Drawing.ColorTranslator]::FromHtml('#2A2A3E')
$clrAccent      = [System.Drawing.ColorTranslator]::FromHtml('#7C9EFF')
$clrAccentHover = [System.Drawing.ColorTranslator]::FromHtml('#A3BEFF')
$clrText        = [System.Drawing.ColorTranslator]::FromHtml('#CDD6F4')
$clrInput       = [System.Drawing.ColorTranslator]::FromHtml('#313244')
$clrOutputBg    = [System.Drawing.ColorTranslator]::FromHtml('#181825')
$clrOutputText  = [System.Drawing.ColorTranslator]::FromHtml('#A6E3A1')
$clrClearNormal = [System.Drawing.ColorTranslator]::FromHtml('#45475A')
$clrClearHover  = [System.Drawing.ColorTranslator]::FromHtml('#585B70')

$fontUI     = New-Object System.Drawing.Font('Segoe UI', 10)
$fontLabel  = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$fontOutput = New-Object System.Drawing.Font('Consolas', 10)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Set-DarkControl {
    param(
        [System.Windows.Forms.Control]$Control,
        [System.Drawing.Color]$BackColor = $clrInput,
        [System.Drawing.Color]$ForeColor = $clrText
    )
    $Control.BackColor = $BackColor
    $Control.ForeColor = $ForeColor
}

function Set-DarkButton {
    param(
        [System.Windows.Forms.Button]$Button,
        [System.Drawing.Color]$NormalColor = $clrAccent,
        [System.Drawing.Color]$HoverColor  = $clrAccentHover,
        [System.Drawing.Color]$TextColor   = $clrBackground
    )
    $Button.FlatStyle = 'Flat'
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $NormalColor
    $Button.ForeColor = $TextColor
    $Button.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $Button.Add_MouseEnter({ $this.BackColor = $HoverColor }.GetNewClosure())
    $Button.Add_MouseLeave({ $this.BackColor = $NormalColor }.GetNewClosure())
}

# ---------------------------------------------------------------------------
# Form
# ---------------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Microsoft Graph API — User Grant Manager'
$form.Size          = New-Object System.Drawing.Size(960, 700)
$form.StartPosition = 'CenterScreen'
$form.Font          = $fontUI
$form.BackColor     = $clrBackground
$form.ForeColor     = $clrText

# ---------------------------------------------------------------------------
# GroupBox — Parameters
# ---------------------------------------------------------------------------
$grpInput           = New-Object System.Windows.Forms.GroupBox
$grpInput.Location  = New-Object System.Drawing.Point(16, 12)
$grpInput.Size      = New-Object System.Drawing.Size(912, 262)
$grpInput.Text      = 'Parameters'
$grpInput.ForeColor = $clrAccent
$grpInput.BackColor = $clrPanel
$form.Controls.Add($grpInput)

# ---------------------------------------------------------------------------
# Labels
# ---------------------------------------------------------------------------
$lblAction          = New-Object System.Windows.Forms.Label
$lblAction.Location = New-Object System.Drawing.Point(12, 22)
$lblAction.Size     = New-Object System.Drawing.Size(100, 24)
$lblAction.Text     = 'Action'
$lblAction.Font     = $fontLabel
$lblAction.ForeColor = $clrAccent
$lblAction.BackColor = $clrPanel
$grpInput.Controls.Add($lblAction)

$lblScopes          = New-Object System.Windows.Forms.Label
$lblScopes.Location = New-Object System.Drawing.Point(12, 62)
$lblScopes.Size     = New-Object System.Drawing.Size(100, 24)
$lblScopes.Text     = 'Scopes'
$lblScopes.Font     = $fontLabel
$lblScopes.ForeColor = $clrAccent
$lblScopes.BackColor = $clrPanel
$grpInput.Controls.Add($lblScopes)

$lblConsent          = New-Object System.Windows.Forms.Label
$lblConsent.Location = New-Object System.Drawing.Point(12, 102)
$lblConsent.Size     = New-Object System.Drawing.Size(100, 24)
$lblConsent.Text     = 'Consent'
$lblConsent.Font     = $fontLabel
$lblConsent.ForeColor = $clrAccent
$lblConsent.BackColor = $clrPanel
$grpInput.Controls.Add($lblConsent)

$lblPrincipal          = New-Object System.Windows.Forms.Label
$lblPrincipal.Location = New-Object System.Drawing.Point(12, 142)
$lblPrincipal.Size     = New-Object System.Drawing.Size(100, 24)
$lblPrincipal.Text     = 'UPN'
$lblPrincipal.Font     = $fontLabel
$lblPrincipal.ForeColor = $clrAccent
$lblPrincipal.BackColor = $clrPanel
$grpInput.Controls.Add($lblPrincipal)

$lblResolvedId          = New-Object System.Windows.Forms.Label
$lblResolvedId.Location = New-Object System.Drawing.Point(12, 184)
$lblResolvedId.Size     = New-Object System.Drawing.Size(100, 24)
$lblResolvedId.Text     = 'Object ID'
$lblResolvedId.Font     = $fontLabel
$lblResolvedId.ForeColor = $clrAccent
$lblResolvedId.BackColor = $clrPanel
$grpInput.Controls.Add($lblResolvedId)

# ---------------------------------------------------------------------------
# Input controls
# ---------------------------------------------------------------------------
$cmbAction              = New-Object System.Windows.Forms.ComboBox
$cmbAction.Location     = New-Object System.Drawing.Point(120, 20)
$cmbAction.Size         = New-Object System.Drawing.Size(180, 26)
$cmbAction.DropDownStyle = 'DropDownList'
[void]$cmbAction.Items.AddRange(@('View', 'Add', 'Remove'))
$cmbAction.SelectedItem = 'View'
Set-DarkControl -Control $cmbAction
$grpInput.Controls.Add($cmbAction)

$txtScopes                 = New-Object System.Windows.Forms.TextBox
$txtScopes.Location        = New-Object System.Drawing.Point(120, 60)
$txtScopes.Size            = New-Object System.Drawing.Size(772, 26)
$txtScopes.PlaceholderText = 'User.Read Mail.Read or User.Read,Mail.Read'
$txtScopes.BorderStyle     = 'FixedSingle'
Set-DarkControl -Control $txtScopes
$grpInput.Controls.Add($txtScopes)

$cmbConsent              = New-Object System.Windows.Forms.ComboBox
$cmbConsent.Location     = New-Object System.Drawing.Point(120, 100)
$cmbConsent.Size         = New-Object System.Drawing.Size(180, 26)
$cmbConsent.DropDownStyle = 'DropDownList'
[void]$cmbConsent.Items.AddRange(@('AllPrincipals', 'Principal'))
$cmbConsent.SelectedItem = 'AllPrincipals'
Set-DarkControl -Control $cmbConsent
$grpInput.Controls.Add($cmbConsent)

$chkWhatIf           = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Location  = New-Object System.Drawing.Point(316, 102)
$chkWhatIf.Size      = New-Object System.Drawing.Size(170, 24)
$chkWhatIf.Text      = 'Preview only (-WhatIf)'
$chkWhatIf.ForeColor = $clrText
$chkWhatIf.BackColor = $clrPanel
$grpInput.Controls.Add($chkWhatIf)

$txtUpn                 = New-Object System.Windows.Forms.TextBox
$txtUpn.Location        = New-Object System.Drawing.Point(120, 140)
$txtUpn.Size            = New-Object System.Drawing.Size(614, 26)
$txtUpn.PlaceholderText = 'user@domain.com'
$txtUpn.BorderStyle     = 'FixedSingle'
Set-DarkControl -Control $txtUpn
$grpInput.Controls.Add($txtUpn)

$btnLookup          = New-Object System.Windows.Forms.Button
$btnLookup.Location = New-Object System.Drawing.Point(742, 140)
$btnLookup.Size     = New-Object System.Drawing.Size(150, 26)
$btnLookup.Text     = 'Lookup ID'
Set-DarkButton -Button $btnLookup -NormalColor $clrClearNormal -HoverColor $clrClearHover -TextColor $clrText
$grpInput.Controls.Add($btnLookup)

$txtResolvedId              = New-Object System.Windows.Forms.TextBox
$txtResolvedId.Location     = New-Object System.Drawing.Point(120, 184)
$txtResolvedId.Size         = New-Object System.Drawing.Size(772, 26)
$txtResolvedId.ReadOnly     = $true
$txtResolvedId.BorderStyle  = 'FixedSingle'
$txtResolvedId.PlaceholderText = '(resolved object ID will appear here after Lookup)'
Set-DarkControl -Control $txtResolvedId -BackColor $clrPanel
$grpInput.Controls.Add($txtResolvedId)

# ---------------------------------------------------------------------------
# Buttons (inside GroupBox)
# ---------------------------------------------------------------------------
$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Location = New-Object System.Drawing.Point(120, 220)
$btnRun.Size     = New-Object System.Drawing.Size(140, 36)
$btnRun.Text     = 'Run'
$btnRun.Font     = $fontLabel
Set-DarkButton -Button $btnRun -NormalColor $clrAccent -HoverColor $clrAccentHover -TextColor $clrBackground
$grpInput.Controls.Add($btnRun)

$btnClear          = New-Object System.Windows.Forms.Button
$btnClear.Location = New-Object System.Drawing.Point(270, 224)
$btnClear.Size     = New-Object System.Drawing.Size(110, 28)
$btnClear.Text     = 'Clear Output'
Set-DarkButton -Button $btnClear -NormalColor $clrClearNormal -HoverColor $clrClearHover -TextColor $clrText
$grpInput.Controls.Add($btnClear)

# ---------------------------------------------------------------------------
# Output TextBox
# ---------------------------------------------------------------------------
$txtOutput             = New-Object System.Windows.Forms.TextBox
$txtOutput.Location    = New-Object System.Drawing.Point(16, 286)
$txtOutput.Size        = New-Object System.Drawing.Size(912, 346)
$txtOutput.Multiline   = $true
$txtOutput.ScrollBars  = 'Both'
$txtOutput.WordWrap    = $false
$txtOutput.Font        = $fontOutput
$txtOutput.BackColor   = $clrOutputBg
$txtOutput.ForeColor   = $clrOutputText
$txtOutput.BorderStyle = 'None'
$form.Controls.Add($txtOutput)

# ---------------------------------------------------------------------------
# StatusStrip
# ---------------------------------------------------------------------------
$statusStrip           = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $clrPanel
$statusLabel           = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text      = 'Idle'
$statusLabel.ForeColor = $clrText
[void]$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)
$statusStrip.Dock = 'Bottom'

# ---------------------------------------------------------------------------
# Logic
# ---------------------------------------------------------------------------
function Set-InputState {
    param ([string]$Action)

    $requiresScopes = $Action -in @('Add', 'Remove')
    $txtScopes.Enabled = $requiresScopes

    if (-not $requiresScopes) {
        $txtScopes.Text = ''
    }
    $txtScopes.BackColor = if ($requiresScopes) { $clrInput } else { $clrPanel }
}

$cmbAction.Add_SelectedIndexChanged({
    Set-InputState -Action $cmbAction.SelectedItem
})

$btnClear.Add_Click({
    $txtOutput.Clear()
})

$btnLookup.Add_Click({
    $upn = $txtUpn.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($upn)) {
        $txtResolvedId.Text = ''
        $statusLabel.Text   = 'Enter a UPN before looking up.'
        return
    }

    $statusLabel.Text       = "Looking up '$upn'..."
    $txtResolvedId.Text     = ''
    $btnLookup.Enabled      = $false

    try {
        $lookupArgs = @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass',
            '-Command', "Import-Module Microsoft.Graph.Users -EA Stop; Connect-MgGraph -Scopes User.Read.All -NoWelcome; (Get-MgUser -UserId '$upn' -Property Id -EA Stop).Id"
        )
        $result = (& pwsh @lookupArgs 2>&1) | Where-Object { $_ -match '^[0-9a-fA-F\-]{36}$' } | Select-Object -Last 1

        if ($result) {
            $txtResolvedId.Text = $result.Trim()
            $statusLabel.Text   = "Resolved: $upn  →  $($txtResolvedId.Text)"
        }
        else {
            $statusLabel.Text = "Could not resolve '$upn' — user not found or not connected."
        }
    }
    catch {
        $statusLabel.Text = "Lookup error: $($_.Exception.Message)"
    }
    finally {
        $btnLookup.Enabled = $true
    }
})

$btnRun.Add_Click({
    try {
        $action = [string]$cmbAction.SelectedItem
        if ([string]::IsNullOrWhiteSpace($action)) {
            throw 'Please select an action.'
        }

        $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $backendScript, '-Action', $action)

        if ($action -in @('Add', 'Remove')) {
            if ([string]::IsNullOrWhiteSpace($txtScopes.Text)) {
                throw 'Scopes are required for Add/Remove.'
            }
            $arguments += @('-Scopes', $txtScopes.Text)
        }

        if ($cmbConsent.SelectedItem) {
            $arguments += @('-ConsentType', [string]$cmbConsent.SelectedItem)
        }

        if (-not [string]::IsNullOrWhiteSpace($txtResolvedId.Text)) {
            $arguments += @('-PrincipalId', $txtResolvedId.Text.Trim())
        }

        if ($chkWhatIf.Checked) {
            $arguments += '-WhatIf'
        }

        $txtOutput.AppendText("`r`n> pwsh $($arguments -join ' ')`r`n")

        $result = & pwsh @arguments 2>&1
        if ($result) {
            $txtOutput.AppendText(($result | Out-String))
        }

        if ($LASTEXITCODE -ne 0) {
            $txtOutput.AppendText("Exited with code $LASTEXITCODE`r`n")
        }

        $statusLabel.Text = "Last run: $(Get-Date -Format 'HH:mm:ss')  |  Action: $action  |  Exit: $LASTEXITCODE"
    }
    catch {
        $txtOutput.AppendText("ERROR: $($_.Exception.Message)`r`n")
        $statusLabel.Text = "Error at $(Get-Date -Format 'HH:mm:ss') — $($_.Exception.Message)"
    }
})

Set-InputState -Action 'View'
[void]$form.ShowDialog()
