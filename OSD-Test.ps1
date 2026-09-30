<#
.SYNOPSIS
    Defines the Run-OSDGUI command: pick and run a script from the Justin-Swets/OSD repo.

.DESCRIPTION
    Dot-source or iex this file to create Run-OSDGUI in the current session:

        iex (irm 'https://raw.githubusercontent.com/Justin-Swets/OSD/refs/heads/main/OSD-Test.ps1')
        Run-OSDGUI

    raw.githubusercontent.com cannot list a directory, so the script list comes from the
    GitHub contents API (unauthenticated: 60 requests/hour per IP, plenty for this).
    Each selected script is run in its own PowerShell window so the GUI stays open and a
    script that calls exit does not close it.

    Uses WinForms rather than WPF: it needs only the .NET Framework that WinPE PowerShell
    already requires, and works on Windows PowerShell 5.1 and PowerShell 7.
#>

$Script:OSDRepoOwner  = 'Justin-Swets'
$Script:OSDRepoName   = 'OSD'
$Script:OSDRepoBranch = 'main'

function Get-OSDRepoScript {
    <# Returns [pscustomobject] Name/Url for every .ps1 in the repo root, sorted by name. #>
    [CmdletBinding()]
    param()

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $api = "https://api.github.com/repos/$($Script:OSDRepoOwner)/$($Script:OSDRepoName)/contents?ref=$($Script:OSDRepoBranch)"
    # The API rejects requests that have no User-Agent.
    $items = Invoke-RestMethod -Uri $api -UseBasicParsing -TimeoutSec 60 -Headers @{
        'User-Agent' = 'OSD-Run-OSDGUI'
        'Accept'     = 'application/vnd.github+json'
    }

    $rawBase = "https://raw.githubusercontent.com/$($Script:OSDRepoOwner)/$($Script:OSDRepoName)/refs/heads/$($Script:OSDRepoBranch)"

    $items |
        Where-Object { $_.type -eq 'file' -and $_.name -like '*.ps1' } |
        Sort-Object name |
        ForEach-Object {
            [pscustomobject]@{
                Name = $_.name
                Url  = "$rawBase/$([uri]::EscapeDataString($_.name))"
            }
        }
}

function Run-OSDGUI {
    <#
    .SYNOPSIS
        Shows a list of the scripts in the OSD GitHub repo; Run executes the selected one.
    #>
    [CmdletBinding()]
    param()

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $form = New-Object System.Windows.Forms.Form
    $form.Text            = 'OSD Script Launcher'
    $form.Size            = New-Object System.Drawing.Size(560, 520)
    $form.MinimumSize     = New-Object System.Drawing.Size(420, 360)
    $form.StartPosition   = 'CenterScreen'

    $label = New-Object System.Windows.Forms.Label
    $label.Text     = "Scripts in $($Script:OSDRepoOwner)/$($Script:OSDRepoName) ($($Script:OSDRepoBranch))"
    $label.Location = New-Object System.Drawing.Point(12, 10)
    $label.AutoSize = $true
    $form.Controls.Add($label)

    $list = New-Object System.Windows.Forms.ListBox
    $list.Location            = New-Object System.Drawing.Point(12, 34)
    $list.Size                = New-Object System.Drawing.Size(520, 380)
    $list.Anchor              = 'Top,Bottom,Left,Right'
    $list.DisplayMember       = 'Name'
    $list.HorizontalScrollbar = $true
    $form.Controls.Add($list)

    $status = New-Object System.Windows.Forms.Label
    $status.Location = New-Object System.Drawing.Point(12, 424)
    $status.Size     = New-Object System.Drawing.Size(520, 20)
    $status.Anchor   = 'Bottom,Left,Right'
    $form.Controls.Add($status)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text     = 'Run'
    $btnRun.Size     = New-Object System.Drawing.Size(90, 30)
    $btnRun.Location = New-Object System.Drawing.Point(342, 448)
    $btnRun.Anchor   = 'Bottom,Right'
    $btnRun.Enabled  = $false
    $form.Controls.Add($btnRun)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text     = 'Refresh'
    $btnRefresh.Size     = New-Object System.Drawing.Size(90, 30)
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 448)
    $btnRefresh.Anchor   = 'Bottom,Left'
    $form.Controls.Add($btnRefresh)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text     = 'Close'
    $btnClose.Size     = New-Object System.Drawing.Size(90, 30)
    $btnClose.Location = New-Object System.Drawing.Point(442, 448)
    $btnClose.Anchor   = 'Bottom,Right'
    $form.Controls.Add($btnClose)

    $form.AcceptButton = $btnRun
    $form.CancelButton = $btnClose

    $loadScripts = {
        $status.Text = 'Loading script list from GitHub...'
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $form.Refresh()
        try {
            $scripts = @(Get-OSDRepoScript)
            $list.Items.Clear()
            foreach ($s in $scripts) { [void]$list.Items.Add($s) }
            $status.Text = "$($scripts.Count) script(s) found."
        }
        catch {
            $list.Items.Clear()
            $status.Text = "Failed to load list: $($_.Exception.Message)"
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $runSelected = {
        $sel = $list.SelectedItem
        if (-not $sel) { return }

        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Run $($sel.Name)?`r`n`r`n$($sel.Url)", 'Confirm', 'YesNo', 'Question')
        if ($answer -ne 'Yes') { return }

        $status.Text = "Started $($sel.Name) in a new window."
        # -NoExit keeps the window open so the script's output can be read.
        $cmd = "iex (irm '$($sel.Url)')"
        Start-Process -FilePath 'powershell.exe' -ArgumentList @(
            '-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-Command', $cmd
        )
    }

    $list.Add_SelectedIndexChanged({ $btnRun.Enabled = ($null -ne $list.SelectedItem) })
    $list.Add_DoubleClick($runSelected)
    $btnRun.Add_Click($runSelected)
    $btnRefresh.Add_Click($loadScripts)
    $btnClose.Add_Click({ $form.Close() })
    $form.Add_Shown($loadScripts)

    [void]$form.ShowDialog()
    $form.Dispose()
}

function Install-OSDGUICommand {
    <#
        Writes Run-OSDGUI to a module under the machine-wide module path so it is
        auto-loaded by EVERY new PowerShell session (e.g. the prompt you open in WinPE
        after Cloud-Provision.ps1 has finished). A function defined by iex only lives in
        the session that ran it, which is why it was "not recognized" before.
    #>
    [CmdletBinding()]
    param(
        [string]$ModuleRoot = (Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules')
    )

    $dir = Join-Path $ModuleRoot 'OSDGUI'
    try {
        New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null

        $psm1 = @"
`$Script:OSDRepoOwner  = '$($Script:OSDRepoOwner)'
`$Script:OSDRepoName   = '$($Script:OSDRepoName)'
`$Script:OSDRepoBranch = '$($Script:OSDRepoBranch)'

function Get-OSDRepoScript {
$((Get-Command Get-OSDRepoScript -CommandType Function).Definition)
}

function Run-OSDGUI {
$((Get-Command Run-OSDGUI -CommandType Function).Definition)
}

Export-ModuleMember -Function Run-OSDGUI, Get-OSDRepoScript
"@
        Set-Content -Path (Join-Path $dir 'OSDGUI.psm1') -Value $psm1 -Encoding UTF8 -Force -ErrorAction Stop

        # A manifest lets command auto-discovery find Run-OSDGUI without importing first.
        New-ModuleManifest -Path (Join-Path $dir 'OSDGUI.psd1') -RootModule 'OSDGUI.psm1' `
            -ModuleVersion '1.0.0' -Description 'Run-OSDGUI: pick and run scripts from Justin-Swets/OSD' `
            -FunctionsToExport @('Run-OSDGUI', 'Get-OSDRepoScript') -CmdletsToExport @() `
            -VariablesToExport @() -AliasesToExport @()

        Write-Host "Run-OSDGUI installed: $dir" -ForegroundColor Green
    }
    catch {
        Write-Host "Could not install the OSDGUI module ($($_.Exception.Message)); Run-OSDGUI is available in this session only." -ForegroundColor Yellow
    }
}

# Loading this script (dot-source or iex) also makes the command permanent for the machine.
Install-OSDGUICommand
