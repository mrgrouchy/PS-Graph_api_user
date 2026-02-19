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

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Microsoft Graph API User Grant Manager'
$form.Size = New-Object System.Drawing.Size(920, 650)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$lblAction = New-Object System.Windows.Forms.Label
$lblAction.Location = New-Object System.Drawing.Point(20, 20)
$lblAction.Size = New-Object System.Drawing.Size(90, 24)
$lblAction.Text = 'Action'
$form.Controls.Add($lblAction)

$cmbAction = New-Object System.Windows.Forms.ComboBox
$cmbAction.Location = New-Object System.Drawing.Point(120, 18)
$cmbAction.Size = New-Object System.Drawing.Size(180, 24)
$cmbAction.DropDownStyle = 'DropDownList'
[void]$cmbAction.Items.AddRange(@('View', 'Add', 'Remove'))
$cmbAction.SelectedItem = 'View'
$form.Controls.Add($cmbAction)

$lblScopes = New-Object System.Windows.Forms.Label
$lblScopes.Location = New-Object System.Drawing.Point(20, 60)
$lblScopes.Size = New-Object System.Drawing.Size(90, 24)
$lblScopes.Text = 'Scopes'
$form.Controls.Add($lblScopes)

$txtScopes = New-Object System.Windows.Forms.TextBox
$txtScopes.Location = New-Object System.Drawing.Point(120, 58)
$txtScopes.Size = New-Object System.Drawing.Size(760, 24)
$txtScopes.PlaceholderText = 'User.Read Mail.Read or User.Read,Mail.Read'
$form.Controls.Add($txtScopes)

$lblConsent = New-Object System.Windows.Forms.Label
$lblConsent.Location = New-Object System.Drawing.Point(20, 100)
$lblConsent.Size = New-Object System.Drawing.Size(90, 24)
$lblConsent.Text = 'Consent'
$form.Controls.Add($lblConsent)

$cmbConsent = New-Object System.Windows.Forms.ComboBox
$cmbConsent.Location = New-Object System.Drawing.Point(120, 98)
$cmbConsent.Size = New-Object System.Drawing.Size(180, 24)
$cmbConsent.DropDownStyle = 'DropDownList'
[void]$cmbConsent.Items.AddRange(@('AllPrincipals', 'Principal'))
$cmbConsent.SelectedItem = 'AllPrincipals'
$form.Controls.Add($cmbConsent)

$lblPrincipal = New-Object System.Windows.Forms.Label
$lblPrincipal.Location = New-Object System.Drawing.Point(20, 140)
$lblPrincipal.Size = New-Object System.Drawing.Size(90, 24)
$lblPrincipal.Text = 'PrincipalId'
$form.Controls.Add($lblPrincipal)

$txtPrincipal = New-Object System.Windows.Forms.TextBox
$txtPrincipal.Location = New-Object System.Drawing.Point(120, 138)
$txtPrincipal.Size = New-Object System.Drawing.Size(500, 24)
$form.Controls.Add($txtPrincipal)

$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Location = New-Object System.Drawing.Point(320, 100)
$chkWhatIf.Size = New-Object System.Drawing.Size(140, 24)
$chkWhatIf.Text = 'Preview only (-WhatIf)'
$form.Controls.Add($chkWhatIf)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Location = New-Object System.Drawing.Point(20, 180)
$btnRun.Size = New-Object System.Drawing.Size(120, 34)
$btnRun.Text = 'Run'
$form.Controls.Add($btnRun)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Location = New-Object System.Drawing.Point(150, 180)
$btnClear.Size = New-Object System.Drawing.Size(120, 34)
$btnClear.Text = 'Clear Output'
$form.Controls.Add($btnClear)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(20, 230)
$txtOutput.Size = New-Object System.Drawing.Size(860, 360)
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = 'Both'
$txtOutput.WordWrap = $false
$txtOutput.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOutput)

function Set-InputState {
    param ([string]$Action)

    $requiresScopes = $Action -in @('Add', 'Remove')
    $txtScopes.Enabled = $requiresScopes

    if (-not $requiresScopes) {
        $txtScopes.Text = ''
    }
}

$cmbAction.Add_SelectedIndexChanged({
    Set-InputState -Action $cmbAction.SelectedItem
})

$btnClear.Add_Click({
    $txtOutput.Clear()
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

        if (-not [string]::IsNullOrWhiteSpace($txtPrincipal.Text)) {
            $arguments += @('-PrincipalId', $txtPrincipal.Text.Trim())
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
    }
    catch {
        $txtOutput.AppendText("ERROR: $($_.Exception.Message)`r`n")
    }
})

Set-InputState -Action 'View'
[void]$form.ShowDialog()
