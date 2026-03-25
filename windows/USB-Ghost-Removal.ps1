#Requires -RunAsAdministrator

<#
.SYNOPSIS
    GUI tool to list and remove ghost (phantom) USB devices on Windows 10/11.

.DESCRIPTION
    Ghost (phantom) USB devices are entries that Windows retains in its device
    database after the physical hardware has been unplugged or is no longer
    present. Over time these orphaned records can accumulate, causing issues
    such as driver conflicts, resource exhaustion, and duplicate COM port
    assignments.

    This script opens an interactive WinForms GUI that:
      - Enumerates all non-present (ghost) USB-related devices using the
        Windows Setup API (SetupDiGetClassDevs / SetupDiEnumDeviceInfo).
      - Compares them against currently present devices to identify phantoms.
      - Displays full device details: name, class, class GUID, enumerator,
        hardware ID, manufacturer, location, and instance ID.
      - Supports live text filtering, multi-row selection (Shift / Ctrl),
        and individual or bulk removal via SetupDiRemoveDevice.
      - Optionally overlays currently active USB devices in the same list
        (highlighted in green) for side-by-side comparison.

    Active (present) devices are shown read-only and cannot be removed.

.NOTES
    Author  : Robert Andersson Jarl / GitHub Copilot
    Version : 1.0
    Tested  : Windows 10, Windows 11
    Requires: Run as Administrator (enforced via #Requires)

.EXAMPLE
    .\USB-Ghost-Removal.ps1
    Launches the GUI. Ghost devices are listed immediately on startup.
    Use the filter box to narrow results, select rows with the mouse
    (Shift for range, Ctrl for individual picks), then click
    "Remove Selected" to delete the chosen phantom entries.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

#region ── Setup API P/Invoke ──────────────────────────────────────────────────
if (-not ([System.Management.Automation.PSTypeName]'GhostApi').Type) {
    Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class GhostApi {

    [DllImport("setupapi.dll", SetLastError = true)]
    public static extern IntPtr SetupDiGetClassDevs(
        ref Guid ClassGuid, IntPtr Enumerator, IntPtr hwndParent, uint Flags);

    [DllImport("setupapi.dll", SetLastError = true)]
    public static extern bool SetupDiEnumDeviceInfo(
        IntPtr DeviceInfoSet, uint MemberIndex, ref SP_DEVINFO_DATA DeviceInfoData);

    [DllImport("setupapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool SetupDiGetDeviceRegistryProperty(
        IntPtr DeviceInfoSet, ref SP_DEVINFO_DATA DeviceInfoData,
        uint Property, out uint PropertyRegDataType,
        StringBuilder PropertyBuffer, uint PropertyBufferSize, out uint RequiredSize);

    [DllImport("setupapi.dll", SetLastError = true)]
    public static extern bool SetupDiRemoveDevice(
        IntPtr DeviceInfoSet, ref SP_DEVINFO_DATA DeviceInfoData);

    [DllImport("setupapi.dll", SetLastError = true)]
    public static extern bool SetupDiDestroyDeviceInfoList(IntPtr DeviceInfoSet);

    [DllImport("cfgmgr32.dll", CharSet = CharSet.Unicode)]
    public static extern uint CM_Get_Device_IDW(
        uint dnDevInst, StringBuilder Buffer, int BufferLen, uint ulFlags);

    public const uint DIGCF_ALLCLASSES      = 0x00000004;
    public const uint DIGCF_PRESENT         = 0x00000002;
    public const uint SPDRP_DEVICEDESC      = 0x00000000;
    public const uint SPDRP_HARDWAREID      = 0x00000001;
    public const uint SPDRP_CLASS           = 0x00000007;
    public const uint SPDRP_MFG             = 0x0000000B;
    public const uint SPDRP_FRIENDLYNAME    = 0x0000000C;
    public const uint SPDRP_LOCATION_INFO   = 0x0000000D;
    public const uint SPDRP_ENUMERATOR_NAME = 0x00000016;

    public static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);
}

[StructLayout(LayoutKind.Sequential)]
public struct SP_DEVINFO_DATA {
    public uint   cbSize;
    public Guid   ClassGuid;
    public uint   DevInst;
    public IntPtr Reserved;
}
'@
}
#endregion

#region ── Device Enumeration ─────────────────────────────────────────────────
$script:allDis  = [GhostApi]::INVALID_HANDLE_VALUE
$script:lastScan = $null

function Get-DevProp([IntPtr]$Dis, [ref]$Did, [uint32]$Prop) {
    $buf = New-Object System.Text.StringBuilder 512
    $rt = [uint32]0; $rs = [uint32]0
    if ([GhostApi]::SetupDiGetDeviceRegistryProperty(
            $Dis, $Did, $Prop, [ref]$rt, $buf, [uint32]$buf.Capacity, [ref]$rs)) {
        return $buf.ToString()
    }
    return ''
}

function Get-InstanceId([uint32]$DevInst) {
    $buf = New-Object System.Text.StringBuilder 512
    if ([GhostApi]::CM_Get_Device_IDW($DevInst, $buf, $buf.Capacity, 0) -eq 0) {
        return $buf.ToString()
    }
    return ''
}

function Read-DeviceSet([IntPtr]$Dis) {
    $list = [System.Collections.Generic.List[hashtable]]::new()
    $did  = New-Object SP_DEVINFO_DATA
    $did.cbSize = [uint32][System.Runtime.InteropServices.Marshal]::SizeOf($did)
    $i = [uint32]0
    while ([GhostApi]::SetupDiEnumDeviceInfo($Dis, $i, [ref]$did)) {
        $fn   = Get-DevProp $Dis ([ref]$did) ([GhostApi]::SPDRP_FRIENDLYNAME)
        $desc = Get-DevProp $Dis ([ref]$did) ([GhostApi]::SPDRP_DEVICEDESC)
        $hwid = Get-DevProp $Dis ([ref]$did) ([GhostApi]::SPDRP_HARDWAREID)
        if ($hwid -match "`0") { $hwid = $hwid.Split("`0")[0] }
        $list.Add(@{
            DevInfo    = $did                       # SP_DEVINFO_DATA value copy
            DevInst    = $did.DevInst
            Name       = if ($fn) { $fn } else { $desc }
            Class      = Get-DevProp $Dis ([ref]$did) ([GhostApi]::SPDRP_CLASS)
            ClassGuid  = $did.ClassGuid.ToString('B').ToUpper()
            Enumerator = Get-DevProp $Dis ([ref]$did) ([GhostApi]::SPDRP_ENUMERATOR_NAME)
            HardwareId = $hwid
            Mfg        = Get-DevProp $Dis ([ref]$did) ([GhostApi]::SPDRP_MFG)
            Location   = Get-DevProp $Dis ([ref]$did) ([GhostApi]::SPDRP_LOCATION_INFO)
            InstanceId = Get-InstanceId $did.DevInst
        })
        $i++
    }
    return $list
}

function Reload-Devices {
    if ($script:allDis -ne [GhostApi]::INVALID_HANDLE_VALUE) {
        [GhostApi]::SetupDiDestroyDeviceInfoList($script:allDis) | Out-Null
        $script:allDis = [GhostApi]::INVALID_HANDLE_VALUE
    }

    $g = [Guid]::Empty
    $script:allDis = [GhostApi]::SetupDiGetClassDevs(
        [ref]$g, [IntPtr]::Zero, [IntPtr]::Zero, [GhostApi]::DIGCF_ALLCLASSES)
    if ($script:allDis -eq [GhostApi]::INVALID_HANDLE_VALUE) {
        throw "SetupDiGetClassDevs (all) failed. Win32 error: $([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())"
    }

    $g2 = [Guid]::Empty
    $presDis = [GhostApi]::SetupDiGetClassDevs(
        [ref]$g2, [IntPtr]::Zero, [IntPtr]::Zero,
        ([GhostApi]::DIGCF_ALLCLASSES -bor [GhostApi]::DIGCF_PRESENT))
    if ($presDis -eq [GhostApi]::INVALID_HANDLE_VALUE) {
        throw "SetupDiGetClassDevs (present) failed. Win32 error: $([System.Runtime.InteropServices.Marshal]::GetLastWin32Error())"
    }

    try {
        $allList  = Read-DeviceSet $script:allDis
        $presList = Read-DeviceSet $presDis

        $presSet = [System.Collections.Generic.HashSet[uint32]]::new()
        foreach ($d in $presList) { [void]$presSet.Add($d.DevInst) }

        # Index allDis entries by DevInst — active devices must use DevInfo from allDis for SetupDiRemoveDevice
        $allByInst = @{}
        foreach ($d in $allList) { $allByInst[$d.DevInst] = $d }

        $usbFilter = {
            $_.Enumerator -match '^USB$' -or
            $_.Class      -match 'Net|USB|Bluetooth' -or
            $_.Name       -match 'USB|Wireless|Wi-Fi|WiFi|Ethernet|Network|Adapter|Bluetooth'
        }

        $ghosts = @($allList | Where-Object { -not $presSet.Contains($_.DevInst) } |
            Where-Object $usbFilter |
            ForEach-Object { $_.IsGhost = $true; $_ })

        $active = @($presList | Where-Object $usbFilter | ForEach-Object {
            if ($allByInst.ContainsKey($_.DevInst)) {
                $d = $allByInst[$_.DevInst]
                $d.IsGhost = $false
                $d
            }
        } | Where-Object { $null -ne $_ })

        return @{ Ghosts = $ghosts; Active = $active }
    }
    finally {
        [GhostApi]::SetupDiDestroyDeviceInfoList($presDis) | Out-Null
    }
}
#endregion

#region ── GUI Helpers ─────────────────────────────────────────────────────────
function Update-StatusBar {
    $selected  = $script:dgv.SelectedRows.Count
    $ghostVis  = @($script:dgv.Rows | Where-Object { $_.Visible -and $_.Tag -and  $_.Tag.IsGhost }).Count
    $activeVis = @($script:dgv.Rows | Where-Object { $_.Visible -and $_.Tag -and -not $_.Tag.IsGhost }).Count

    if (($ghostVis + $activeVis) -eq 0) {
        $script:statusLabel.Text = 'No USB devices found.'
        return
    }
    $parts = [System.Collections.Generic.List[string]]::new()
    if ($ghostVis  -gt 0) { $parts.Add("$ghostVis ghost")  }
    if ($activeVis -gt 0) { $parts.Add("$activeVis active") }
    $txt = ($parts -join '  |  ') + '  device(s)'
    if ($selected -gt 0) { $txt += "  |  $selected selected" }
    $script:statusLabel.Text = $txt
}

function Populate-Grid([hashtable]$Scan) {
    $ghosts = if ($Scan.Ghosts -and $script:cmbShowMode.SelectedIndex -ne 1) { $Scan.Ghosts } else { @() }
    $active = if ($script:cmbShowMode.SelectedIndex -ge 1 -and $Scan.Active) { $Scan.Active } else { @() }

    $colActiveBg   = [System.Drawing.Color]::FromArgb(220, 245, 215)
    $colActiveSel  = [System.Drawing.Color]::FromArgb(130, 195, 130)

    $script:dgv.SuspendLayout()
    $script:dgv.Rows.Clear()

    foreach ($g in $ghosts) {
        $idx = $script:dgv.Rows.Add(
            $g.Name, $g.Class, $g.ClassGuid, $g.Enumerator,
            $g.HardwareId, $g.Mfg, $g.Location, $g.InstanceId)
        $row = $script:dgv.Rows[$idx]
        $row.Tag = $g
    }

    foreach ($g in $active) {
        $idx = $script:dgv.Rows.Add(
            $g.Name, $g.Class, $g.ClassGuid, $g.Enumerator,
            $g.HardwareId, $g.Mfg, $g.Location, $g.InstanceId)
        $row = $script:dgv.Rows[$idx]
        $row.Tag = $g
        $row.DefaultCellStyle.BackColor          = $colActiveBg
        $row.DefaultCellStyle.SelectionBackColor = $colActiveSel
        $row.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black
    }

    $script:dgv.ResumeLayout($false)
    $script:dgv.Sort($script:dgv.Columns['Name'], [System.ComponentModel.ListSortDirection]::Ascending)
    Apply-Filter
}

function Apply-Filter {
    $filter = $script:filterBox.Text.Trim()
    $script:dgv.SuspendLayout()
    foreach ($row in $script:dgv.Rows) {
        if ([string]::IsNullOrEmpty($filter)) {
            $row.Visible = $true
        } else {
            $match = $false
            foreach ($cell in $row.Cells) {
                if ($cell.Value -and
                        $cell.Value.ToString() -match [regex]::Escape($filter)) {
                    $match = $true; break
                }
            }
            $row.Visible = $match
        }
    }
    $script:dgv.ResumeLayout($false)
    Update-StatusBar
}

function Remove-GhostRow([System.Windows.Forms.DataGridViewRow]$Row) {
    $ghost       = $Row.Tag
    $devInfoCopy = $ghost.DevInfo   # value copy — safe for P/Invoke pass-by-ref
    $ok = [GhostApi]::SetupDiRemoveDevice($script:allDis, [ref]$devInfoCopy)
    if ($ok) {
        $script:dgv.Rows.Remove($Row)
        Update-StatusBar
        return $true
    }
    $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to remove:`n$($ghost.Name)`n`nWin32 error: $err",
        'Remove Failed',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    return $false
}
#endregion

#region ── Build Form ──────────────────────────────────────────────────────────
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'USB Ghost Device Remover'
$form.Size          = New-Object System.Drawing.Size(1280, 680)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize   = New-Object System.Drawing.Size(900, 420)
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor     = [System.Drawing.Color]::FromArgb(240, 240, 240)

# ─ Toolbar ─
$toolbar            = New-Object System.Windows.Forms.Panel
$toolbar.Dock       = 'Top'
$toolbar.Height     = 48
$toolbar.BackColor  = [System.Drawing.Color]::FromArgb(228, 228, 228)
$toolbar.Padding    = New-Object System.Windows.Forms.Padding(10, 9, 10, 5)

$lblFilter          = New-Object System.Windows.Forms.Label
$lblFilter.Text     = 'Search:'
$lblFilter.AutoSize = $true
$lblFilter.Location = New-Object System.Drawing.Point(10, 15)

$script:filterBox          = New-Object System.Windows.Forms.TextBox
$script:filterBox.Location = New-Object System.Drawing.Point(54, 11)
$script:filterBox.Width    = 260
$script:filterBox.Font     = New-Object System.Drawing.Font('Segoe UI', 9)

$btnRefresh            = New-Object System.Windows.Forms.Button
$btnRefresh.Text       = 'Refresh'
$btnRefresh.Location   = New-Object System.Drawing.Point(326, 9)
$btnRefresh.Size       = New-Object System.Drawing.Size(88, 28)
$btnRefresh.FlatStyle  = 'Flat'
$btnRefresh.BackColor  = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnRefresh.ForeColor  = [System.Drawing.Color]::White
$btnRefresh.Font       = New-Object System.Drawing.Font('Segoe UI', 9)

$btnRemoveSel           = New-Object System.Windows.Forms.Button
$btnRemoveSel.Text      = 'Remove Selected'
$btnRemoveSel.Location  = New-Object System.Drawing.Point(424, 9)
$btnRemoveSel.Size      = New-Object System.Drawing.Size(130, 28)
$btnRemoveSel.FlatStyle = 'Flat'
$btnRemoveSel.BackColor = [System.Drawing.Color]::FromArgb(196, 43, 28)
$btnRemoveSel.ForeColor = [System.Drawing.Color]::White
$btnRemoveSel.Font      = New-Object System.Drawing.Font('Segoe UI', 9)
$btnRemoveSel.Enabled   = $false

$script:cmbShowMode          = New-Object System.Windows.Forms.ComboBox
$script:cmbShowMode.Location = New-Object System.Drawing.Point(564, 11)
$script:cmbShowMode.Width    = 130
$script:cmbShowMode.Font     = New-Object System.Drawing.Font('Segoe UI', 9)
$script:cmbShowMode.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$script:cmbShowMode.Items.AddRange(@('Show ghosts', 'Show active', 'Show all')) | Out-Null
$script:cmbShowMode.SelectedIndex = 0

$btnAbout           = New-Object System.Windows.Forms.Button
$btnAbout.Text      = 'About'
$btnAbout.Location  = New-Object System.Drawing.Point(730, 9)
$btnAbout.Size      = New-Object System.Drawing.Size(70, 28)
$btnAbout.FlatStyle = 'Flat'
$btnAbout.BackColor = [System.Drawing.Color]::FromArgb(228, 228, 228)
$btnAbout.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$btnAbout.Font      = New-Object System.Drawing.Font('Segoe UI', 9)

$btnSaveCsv           = New-Object System.Windows.Forms.Button
$btnSaveCsv.Text      = 'Save list as CSV'
$btnSaveCsv.Location  = New-Object System.Drawing.Point(812, 9)
$btnSaveCsv.Size      = New-Object System.Drawing.Size(130, 28)
$btnSaveCsv.FlatStyle = 'Flat'
$btnSaveCsv.BackColor = [System.Drawing.Color]::FromArgb(0, 140, 70)
$btnSaveCsv.ForeColor = [System.Drawing.Color]::White
$btnSaveCsv.Font      = New-Object System.Drawing.Font('Segoe UI', 9)

$toolbar.Controls.AddRange(@($lblFilter, $script:filterBox, $btnRefresh, $btnRemoveSel, $script:cmbShowMode, $btnAbout, $btnSaveCsv))

# ─ DataGridView ─
$script:dgv = New-Object System.Windows.Forms.DataGridView
$script:dgv.Dock                          = 'Fill'
$script:dgv.AllowUserToAddRows            = $false
$script:dgv.AllowUserToDeleteRows         = $false
$script:dgv.ReadOnly                      = $false
$script:dgv.SelectionMode                 = 'FullRowSelect'
$script:dgv.MultiSelect                   = $true
$script:dgv.RowHeadersVisible             = $false
$script:dgv.AutoSizeColumnsMode           = 'None'
$script:dgv.ColumnHeadersHeightSizeMode   = 'AutoSize'
$script:dgv.ScrollBars                    = 'Both'
$script:dgv.BackgroundColor               = [System.Drawing.Color]::White
$script:dgv.BorderStyle                   = 'None'
$script:dgv.GridColor                     = [System.Drawing.Color]::FromArgb(218, 218, 218)
$script:dgv.DefaultCellStyle.Padding      = New-Object System.Windows.Forms.Padding(4, 2, 4, 2)
$script:dgv.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$script:dgv.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$script:dgv.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(246, 249, 255)
$script:dgv.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$script:dgv.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(228, 228, 228)
$script:dgv.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$script:dgv.EnableHeadersVisualStyles     = $false

# Text columns
$textCols = @(
    @{ Name = 'Name';       Header = 'Name / Description'; Width = 260 }
    @{ Name = 'Class';      Header = 'Class';               Width = 105 }
    @{ Name = 'ClassGuid';  Header = 'Class GUID';          Width = 220 }
    @{ Name = 'Enumerator'; Header = 'Enumerator';          Width = 85  }
    @{ Name = 'HardwareId'; Header = 'Hardware ID';         Width = 220 }
    @{ Name = 'Mfg';        Header = 'Manufacturer';        Width = 130 }
    @{ Name = 'Location';   Header = 'Location';            Width = 110 }
    @{ Name = 'InstanceId'; Header = 'Instance ID';         Width = 220 }
)
foreach ($c in $textCols) {
    $col            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name       = $c.Name
    $col.HeaderText = $c.Header
    $col.Width      = $c.Width
    $col.ReadOnly   = $true
    $col.SortMode   = 'Automatic'
    $script:dgv.Columns.Add($col) | Out-Null
}

# Remove button column
# (removed — use the toolbar buttons instead)

# ─ Status strip ─
$statusStrip                = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor      = [System.Drawing.Color]::FromArgb(228, 228, 228)
$statusStrip.SizingGrip     = $false
$script:statusLabel         = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:statusLabel.Text    = 'Loading...'
$script:statusLabel.Spring  = $true
$script:statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statusStrip.Items.Add($script:statusLabel) | Out-Null

# Assemble
$form.Controls.Add($script:dgv)
$form.Controls.Add($toolbar)
$form.Controls.Add($statusStrip)
#endregion

#region ── Event Handlers ──────────────────────────────────────────────────────
$script:filterBox.Add_TextChanged({ Apply-Filter })

$script:cmbShowMode.Add_SelectedIndexChanged({
    if ($script:lastScan) { Populate-Grid $script:lastScan }
})

$btnSaveCsv.Add_Click({
    $modeNames = @('USB-GhostsOnly', 'USB-ActiveOnly', 'USB-All')
    $modeName  = $modeNames[$script:cmbShowMode.SelectedIndex]
    $timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
    $suggested = "${modeName}_${timestamp}.csv"

    $dlg                  = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title            = 'Save list as CSV'
    $dlg.Filter           = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dlg.FileName         = $suggested
    $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')

    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $headers = @('Name','Class','ClassGuid','Enumerator','HardwareId','Mfg','Location','InstanceId')
        $lines   = [System.Collections.Generic.List[string]]::new()
        $lines.Add('"' + ($headers -join '","') + '"')
        foreach ($row in $script:dgv.Rows) {
            if ($row.Visible) {
                $cells = foreach ($h in $headers) {
                    if ($row.Cells[$h].Value) { $row.Cells[$h].Value.ToString() -replace '"','""' } else { '' }
                }
                $lines.Add('"' + ($cells -join '","') + '"')
            }
        }
        [System.IO.File]::WriteAllLines($dlg.FileName, $lines, [System.Text.Encoding]::UTF8)
        [System.Windows.Forms.MessageBox]::Show(
            "Saved $($lines.Count - 1) row(s) to:`n$($dlg.FileName)",
            'CSV Saved',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
})

$btnAbout.Add_Click({
    [System.Windows.Forms.MessageBox]::Show(
        "USB Ghost Device Remover`n" +
        "----------------------------`n" +
        "Author  : Robert Andersson Jarl / GitHub Copilot`n" +
        "Version : 1.0`n" +
        "Tested  : Windows 10, Windows 11`n" +
        "Requires: Run as Administrator`n`n" +
        "Identifies and removes ghost (phantom) USB devices`n" +
        "left behind in the Windows device database after`n" +
        "hardware has been disconnected.",
        'About USB Ghost Device Remover',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
})

# Update Remove Selected button state and status bar when selection changes
$script:dgv.Add_SelectionChanged({
    $sel = $script:dgv.SelectedRows.Count
    $btnRemoveSel.Enabled = ($sel -gt 0)
    $btnRemoveSel.Text    = if ($sel -gt 1) { "Remove Selected ($sel)" } else { 'Remove Selected' }
    Update-StatusBar
})

$btnRefresh.Add_Click({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $script:statusLabel.Text = 'Scanning...'
    $form.Refresh()
    try {
        $script:lastScan = Reload-Devices
        Populate-Grid $script:lastScan
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message, 'Scan Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Per-row Remove button handler removed (column no longer exists)

# Remove Selected (ghosts only — active rows are silently skipped)
$btnRemoveSel.Add_Click({
    $ghostRows = @($script:dgv.SelectedRows | Where-Object { $_.Tag.IsGhost })
    if ($ghostRows.Count -eq 0) { return }

    if ($ghostRows.Count -eq 1) {
        $ghost = $ghostRows[0].Tag
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "Remove this ghost device?`n`n$($ghost.Name)",
            'Confirm Remove',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    } else {
        $names = ($ghostRows | ForEach-Object { "  • $($_.Tag.Name)" }) -join "`n"
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "Remove $($ghostRows.Count) selected ghost device(s)?`n`n$names`n`nThis action cannot be undone.",
            'Confirm Remove Selected',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $removed = 0; $failed = 0
    foreach ($row in $ghostRows) {
        $ghost       = $row.Tag
        $devInfoCopy = $ghost.DevInfo
        if ([GhostApi]::SetupDiRemoveDevice($script:allDis, [ref]$devInfoCopy)) {
            $script:dgv.Rows.Remove($row)
            $removed++
        } else {
            $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            Write-Warning "Failed to remove: $($ghost.Name) (Win32 error $err)"
            $failed++
        }
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
    Update-StatusBar
    if ($failed -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Removed: $removed`nFailed:  $failed", 'Done',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
})

# Initial load on form show
$form.Add_Load({
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $script:statusLabel.Text = 'Scanning for ghost devices...'
    $form.Refresh()
    try {
        $script:lastScan = Reload-Devices
        Populate-Grid $script:lastScan
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message, 'Scan Error',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        $script:statusLabel.Text = 'Error scanning devices.'
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

# Cleanup on close
$form.Add_FormClosed({
    if ($script:allDis -ne [GhostApi]::INVALID_HANDLE_VALUE) {
        [GhostApi]::SetupDiDestroyDeviceInfoList($script:allDis) | Out-Null
        $script:allDis = [GhostApi]::INVALID_HANDLE_VALUE
    }
})
#endregion

# ── Launch ─────────────────────────────────────────────────────────────────────
$form.ShowDialog() | Out-Null
