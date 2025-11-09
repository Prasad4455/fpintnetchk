function Show-IntuneJsonGuiAndUpload {
    <#
    .SYNOPSIS
      Show a WinForms GUI for an Intune connectivity JSON file with filters and SharePoint upload.

    .PARAMETER JsonPath
      Path to the JSON file to display.

    .PARAMETER SharePointSiteUrl
      Target SharePoint site URL (e.g. https://contoso.sharepoint.com/sites/IntuneReports). Optional for local export.

    .PARAMETER SharePointLibrary
      Document library name (e.g. 'Shared Documents' or 'Documents'). Default 'Shared Documents'.

    .PARAMETER SharePointFolder
      Optional folder inside the library to upload file to (e.g. 'Intune/Reports').

    .PARAMETER UsePnP
      Switch to use PnP.PowerShell for upload (recommended). If not present and switch passed, script offers to install.

    .EXAMPLE
      Show-IntuneJsonGuiAndUpload -JsonPath ".\Deepe_HPWORKSTATION99_20251109_002758.json" -SharePointSiteUrl "https://contoso.sharepoint.com/sites/IntuneReports" -UsePnP
    #>

    param(
        [Parameter(Mandatory = $true)][string]$JsonPath,
        [string]$SharePointSiteUrl = "",
        [string]$SharePointLibrary = "Shared Documents",
        [string]$SharePointFolder = "",
        [switch]$UsePnP
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    if (-not (Test-Path -Path $JsonPath)) {
        [System.Windows.Forms.MessageBox]::Show("JSON file not found: $JsonPath","File missing",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    try {
        $jsonRaw = Get-Content -Path $JsonPath -Raw
        $jsonObj = $jsonRaw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to parse JSON: $($_.Exception.Message)","JSON parse error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return
    }

    # Build a flattened list of endpoints for display: include Category, URL, Status, StatusCode, Mandatory, TestTimestamp, RegionNote, SourceGroup (M365CommonEndpoints/IntuneEndpoints)
    $items = @()
    if ($jsonObj.ConnectivityResults.M365CommonEndpoints) {
        foreach ($e in $jsonObj.ConnectivityResults.M365CommonEndpoints) {
            $items += [PSCustomObject]@{
                SourceGroup   = "M365Common"
                Category      = $e.Category
                URL           = $e.URL
                Status        = $e.Status
                StatusCode    = $e.StatusCode
                Mandatory     = [string]$e.Mandatory
                TestTimestamp = $e.TestTimestamp
                RegionNote    = $e.RegionNote
            }
        }
    }
    if ($jsonObj.ConnectivityResults.IntuneEndpoints) {
        foreach ($e in $jsonObj.ConnectivityResults.IntuneEndpoints) {
            $items += [PSCustomObject]@{
                SourceGroup   = "Intune"
                Category      = $e.Category
                URL           = $e.URL
                Status        = $e.Status
                StatusCode    = $e.StatusCode
                Mandatory     = [string]$e.Mandatory
                TestTimestamp = $e.TestTimestamp
                RegionNote    = $e.RegionNote
            }
        }
    }

    # convert timestamps to [datetime] where possible for filtering
    foreach ($it in $items) {
        try { $it.TestDateTime = [datetime]::Parse($it.TestTimestamp) } catch { $it.TestDateTime = $null }
    }

    # Build UI
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Intune Connectivity — JSON Viewer & SharePoint Uploader"
    $form.Size = New-Object System.Drawing.Size(1200,760)
    $form.StartPosition = "CenterScreen"

    # Summary / Metadata panel (top)
    $panelTop = New-Object System.Windows.Forms.Panel
    $panelTop.Dock = 'Top'
    $panelTop.Height = 120
    $panelTop.BorderStyle = 'FixedSingle'
    $form.Controls.Add($panelTop)

    $lblSummary = New-Object System.Windows.Forms.Label
    $lblSummary.AutoSize = $true
    $lblSummary.Location = New-Object System.Drawing.Point(10,8)
    $lblSummary.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    # Compose summary text
    $summaryText = "Test Status: {0}    Total Endpoints: {1}    Failed: {2}    Passed: {3}`r`nTester: {4}    Computer: {5}    TestDateTime: {6}" -f `
        ($jsonObj.Summary.TestStatus), ($jsonObj.Summary.TotalEndpointsTested), ($jsonObj.Summary.FailedConnections), ($jsonObj.Summary.SuccessfulConnections), `
        ($jsonObj.TestMetadata.Username), ($jsonObj.TestMetadata.ComputerName), ($jsonObj.TestMetadata.TestDateTime)
    $lblSummary.Text = $summaryText
    $lblSummary.Size = New-Object System.Drawing.Size(1100,40)
    $panelTop.Controls.Add($lblSummary)

    # Filter controls
    $x = 10; $y = 60; $gap = 160
    # Category filter
    $lblCat = New-Object System.Windows.Forms.Label
    $lblCat.Location = New-Object System.Drawing.Point($x,$y)
    $lblCat.Size = New-Object System.Drawing.Size(60,20)
    $lblCat.Text = "Category"
    $panelTop.Controls.Add($lblCat)

    $cmbCategory = New-Object System.Windows.Forms.ComboBox
    $cmbCategory.DropDownStyle = 'DropDownList'
    $cmbCategory.Location = New-Object System.Drawing.Point($x + 64, $y - 3)
    $cmbCategory.Size = New-Object System.Drawing.Size(140,22)
    $panelTop.Controls.Add($cmbCategory)

    # Status filter
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point($x + $gap, $y)
    $lblStatus.Size = New-Object System.Drawing.Size(50,20)
    $lblStatus.Text = "Status"
    $panelTop.Controls.Add($lblStatus)

    $cmbStatus = New-Object System.Windows.Forms.ComboBox
    $cmbStatus.DropDownStyle = 'DropDownList'
    $cmbStatus.Location = New-Object System.Drawing.Point($x + $gap + 64, $y - 3)
    $cmbStatus.Size = New-Object System.Drawing.Size(120,22)
    $panelTop.Controls.Add($cmbStatus)

    # Mandatory filter
    $lblMand = New-Object System.Windows.Forms.Label
    $lblMand.Location = New-Object System.Drawing.Point($x + ($gap*2), $y)
    $lblMand.Size = New-Object System.Drawing.Size(70,20)
    $lblMand.Text = "Mandatory"
    $panelTop.Controls.Add($lblMand)

    $cmbMandatory = New-Object System.Windows.Forms.ComboBox
    $cmbMandatory.DropDownStyle = 'DropDownList'
    $cmbMandatory.Location = New-Object System.Drawing.Point($x + ($gap*2) + 80, $y - 3)
    $cmbMandatory.Size = New-Object System.Drawing.Size(120,22)
    $panelTop.Controls.Add($cmbMandatory)

    # Tester / Username filter
    $lblTester = New-Object System.Windows.Forms.Label
    $lblTester.Location = New-Object System.Drawing.Point($x + ($gap*3), $y)
    $lblTester.Size = New-Object System.Drawing.Size(50,20)
    $lblTester.Text = "Tester"
    $panelTop.Controls.Add($lblTester)

    $txtTester = New-Object System.Windows.Forms.TextBox
    $txtTester.Location = New-Object System.Drawing.Point($x + ($gap*3) + 56, $y - 3)
    $txtTester.Size = New-Object System.Drawing.Size(140,22)
    $txtTester.Text = $jsonObj.TestMetadata.Username
    $panelTop.Controls.Add($txtTester)

    # Date range
    $lblFrom = New-Object System.Windows.Forms.Label
    $lblFrom.Location = New-Object System.Drawing.Point(10,92)
    $lblFrom.Size = New-Object System.Drawing.Size(30,20)
    $lblFrom.Text = "From"
    $panelTop.Controls.Add($lblFrom)

    $dtFrom = New-Object System.Windows.Forms.DateTimePicker
    $dtFrom.Format = 'Custom'
    $dtFrom.CustomFormat = 'yyyy-MM-dd'
    $dtFrom.Location = New-Object System.Drawing.Point(50,88)
    $dtFrom.Size = New-Object System.Drawing.Size(110,22)
    $panelTop.Controls.Add($dtFrom)

    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Location = New-Object System.Drawing.Point(170,92)
    $lblTo.Size = New-Object System.Drawing.Size(20,20)
    $lblTo.Text = "To"
    $panelTop.Controls.Add($lblTo)

    $dtTo = New-Object System.Windows.Forms.DateTimePicker
    $dtTo.Format = 'Custom'
    $dtTo.CustomFormat = 'yyyy-MM-dd'
    $dtTo.Location = New-Object System.Drawing.Point(200,88)
    $dtTo.Size = New-Object System.Drawing.Size(110,22)
    $panelTop.Controls.Add($dtTo)

    # URL search
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location = New-Object System.Drawing.Point(330,92)
    $lblSearch.Size = New-Object System.Drawing.Size(30,20)
    $lblSearch.Text = "URL"
    $panelTop.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(366,88)
    $txtSearch.Size = New-Object System.Drawing.Size(240,22)
    $panelTop.Controls.Add($txtSearch)

    # Buttons: Apply Filters, Reset, Save JSON, Upload to SharePoint, Close
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = "Apply Filters"
    $btnApply.Location = New-Object System.Drawing.Point(620,86)
    $btnApply.Size = New-Object System.Drawing.Size(100,26)
    $panelTop.Controls.Add($btnApply)

    $btnReset = New-Object System.Windows.Forms.Button
    $btnReset.Text = "Reset Filters"
    $btnReset.Location = New-Object System.Drawing.Point(730,86)
    $btnReset.Size = New-Object System.Drawing.Size(100,26)
    $panelTop.Controls.Add($btnReset)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save Filtered JSON"
    $btnSave.Location = New-Object System.Drawing.Point(840,86)
    $btnSave.Size = New-Object System.Drawing.Size(140,26)
    $panelTop.Controls.Add($btnSave)

    $btnUpload = New-Object System.Windows.Forms.Button
    $btnUpload.Text = "Upload to SharePoint"
    $btnUpload.Location = New-Object System.Drawing.Point(990,86)
    $btnUpload.Size = New-Object System.Drawing.Size(150,26)
    $panelTop.Controls.Add($btnUpload)

    # Data grid (center)
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AutoSizeColumnsMode = 'AllCells'
    $form.Controls.Add($grid)
    $grid.BringToFront()

    # Populate filter choices
    $catVals = $items | Select-Object -ExpandProperty Category -Unique | Sort-Object
    $cmbCategory.Items.Add("All") | Out-Null
    foreach ($c in $catVals) { $cmbCategory.Items.Add($c) | Out-Null }
    $cmbCategory.SelectedIndex = 0

    $statusVals = $items | Select-Object -ExpandProperty Status -Unique | Sort-Object
    $cmbStatus.Items.Add("All") | Out-Null
    foreach ($s in $statusVals) { $cmbStatus.Items.Add($s) | Out-Null }
    $cmbStatus.SelectedIndex = 0

    $cmbMandatory.Items.Add("All") | Out-Null
    $cmbMandatory.Items.Add("True") | Out-Null
    $cmbMandatory.Items.Add("False") | Out-Null
    $cmbMandatory.SelectedIndex = 0

    # create binding source
    $binding = New-Object System.Windows.Forms.BindingSource
    $binding.DataSource = ($items | Sort-Object TestDateTime)
    $grid.DataSource = $binding

    # Filter function
    $applyFilter = {
        $filtered = $items

        # Category
        if ($cmbCategory.SelectedItem -and $cmbCategory.SelectedItem -ne "All") {
            $filtered = $filtered | Where-Object { $_.Category -eq $cmbCategory.SelectedItem }
        }
        # Status
        if ($cmbStatus.SelectedItem -and $cmbStatus.SelectedItem -ne "All") {
            $filtered = $filtered | Where-Object { $_.Status -eq $cmbStatus.SelectedItem }
        }
        # Mandatory
        if ($cmbMandatory.SelectedItem -and $cmbMandatory.SelectedItem -ne "All") {
            $filtered = $filtered | Where-Object { $_.Mandatory -eq $cmbMandatory.SelectedItem }
        }
        # Tester match (search in TestMetadata Username or TesterName)
        if ($txtTester.Text -and $txtTester.Text.Trim().Length -gt 0) {
            $tester = $txtTester.Text.Trim()
            # JSON TestMetadata Username stored separately; we simply filter if tester matches the metadata or is substring of TesterName in EnhancedSystemInfo
            # But for endpoints, we include rows if metadata username matches tester
            if ($jsonObj.TestMetadata.Username -and ($jsonObj.TestMetadata.Username -like "*$tester*")) {
                # no-op (metadata applies to whole set)
            } else {
                $filtered = $filtered | Where-Object { ($_.URL -match [regex]::Escape($tester)) -or ($_.Category -match [regex]::Escape($tester)) }
            }
        }
        # Date range
        if ($dtFrom.Value -and $dtTo.Value) {
            $from = $dtFrom.Value.Date
            $to   = $dtTo.Value.Date.AddDays(1).AddSeconds(-1)
            $filtered = $filtered | Where-Object { $_.TestDateTime -ne $null -and $_.TestDateTime -ge $from -and $_.TestDateTime -le $to }
        }
        # URL search
        if ($txtSearch.Text -and $txtSearch.Text.Trim().Length -gt 0) {
            $q = $txtSearch.Text.Trim()
            $filtered = $filtered | Where-Object { $_.URL -like "*$q*" -or $_.RegionNote -like "*$q*" }
        }

        $binding.DataSource = ($filtered | Sort-Object TestDateTime)
        $grid.Refresh()
    }

    $btnApply.Add_Click({ & $applyFilter })
    $btnReset.Add_Click({
        $cmbCategory.SelectedIndex = 0
        $cmbStatus.SelectedIndex = 0
        $cmbMandatory.SelectedIndex = 0
        $txtTester.Text = $jsonObj.TestMetadata.Username
        $txtSearch.Text = ""
        # reset date range to TestDate if present
        try {
            $td = [datetime]::Parse($jsonObj.TestMetadata.TestDateTime)
            $dtFrom.Value = $td.Date
            $dtTo.Value   = $td.Date
        } catch {
            $dtFrom.Value = (Get-Date).Date
            $dtTo.Value   = (Get-Date).Date
        }
        & $applyFilter
    })

    # Save filtered JSON button click: create JSON structure containing TestMetadata, Summary, NetworkConfiguration, EnhancedSystemInfo and Filtered endpoints
    $btnSave.Add_Click({
        try {
            $filteredRows = $binding.DataSource
            $exportObj = [PSCustomObject]@{
                TestMetadata = $jsonObj.TestMetadata
                Summary = $jsonObj.Summary
                EnhancedSystemInfo = $jsonObj.EnhancedSystemInfo
                NetworkConfiguration = $jsonObj.NetworkConfiguration
                FilteredEndpoints = $filteredRows
            }
            # filename: username_YYYYMMDD_HHMMSS.json in same directory as source file
            $dir = Split-Path -Path $JsonPath -Parent
            $username = $jsonObj.TestMetadata.Username -replace '\s+',''
            if (-not $username) { $username = "unknown" }
            $ts = Get-Date -Format "yyyyMMdd_HHmmss"
            $outPath = Join-Path $dir ("{0}_{1}.json" -f $username, $ts)
            $exportObj | ConvertTo-Json -Depth 10 | Out-File -FilePath $outPath -Encoding UTF8

            [System.Windows.Forms.MessageBox]::Show("Saved filtered JSON to `n$outPath","Save complete",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to save JSON: $($_.Exception.Message)","Save error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    # Upload to SharePoint (uses PnP.PowerShell)
    $btnUpload.Add_Click({
        # Ensure there's a filtered file saved (or create a temporary export first)
        try {
            $filteredRows = $binding.DataSource
            $exportObj = [PSCustomObject]@{
                TestMetadata = $jsonObj.TestMetadata
                Summary = $jsonObj.Summary
                EnhancedSystemInfo = $jsonObj.EnhancedSystemInfo
                NetworkConfiguration = $jsonObj.NetworkConfiguration
                FilteredEndpoints = $filteredRows
            }
            $dir = Split-Path -Path $JsonPath -Parent
            $username = $jsonObj.TestMetadata.Username -replace '\s+',''
            if (-not $username) { $username = "unknown" }
            $ts = Get-Date -Format "yyyyMMdd_HHmmss"
            $outPath = Join-Path $dir ("{0}_{1}.json" -f $username, $ts)
            $exportObj | ConvertTo-Json -Depth 10 | Out-File -FilePath $outPath -Encoding UTF8
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to prepare JSON for upload: $($_.Exception.Message)","Export error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            return
        }

        if (-not $SharePointSiteUrl -or $SharePointSiteUrl.Trim().Length -eq 0) {
            $res = [System.Windows.Forms.MessageBox]::Show("No SharePointSiteUrl provided. Do you want to copy the exported file path to clipboard instead?","No SharePoint", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($res -eq [System.Windows.Forms.DialogResult]::Yes) {
                Set-Clipboard -Value $outPath
                [System.Windows.Forms.MessageBox]::Show("Path copied to clipboard: `n$outPath","Copied",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            }
            return
        }

        # Ensure PnP.PowerShell module if using it
        if ($UsePnP) {
            if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
                $install = [System.Windows.Forms.MessageBox]::Show("PnP.PowerShell module is not installed. Install from PSGallery now? (Requires admin rights)","Module missing",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
                if ($install -eq [System.Windows.Forms.DialogResult]::Yes) {
                    try {
                        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show("Failed to install PnP.PowerShell: $($_.Exception.Message)","Install error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                        return
                    }
                } else {
                    return
                }
            }

            try {
                Import-Module PnP.PowerShell -ErrorAction Stop
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to import PnP.PowerShell: $($_.Exception.Message)","Import error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
                return
            }

            # Connect (interactive) and upload
            try {
                # Interactive sign-in to allow MFA and modern auth
                Connect-PnPOnline -Url $SharePointSiteUrl -Interactive -ErrorAction Stop

                # Build destination folder path
                $destFolder = $SharePointLibrary.TrimEnd('/') 
                if ($SharePointFolder -and $SharePointFolder.Trim().Length -gt 0) {
                    $destFolder = "$destFolder/$SharePointFolder"
                }

                # Create folder if not exists
                try {
                    Ensure-PnPFolder -SiteRelativePath $destFolder -ErrorAction SilentlyContinue
                } catch {
                    # If Ensure-PnPFolder not present, fallback to Add-PnPFolder -Folder or create via web
                    try { Add-PnPFolder -Name $SharePointFolder -Folder $SharePointLibrary -ErrorAction SilentlyContinue } catch {}
                }

                # Upload file - overwrite if exists
                Add-PnPFile -Path $outPath -Folder $destFolder -Overwrite -ErrorAction Stop

                [System.Windows.Forms.MessageBox]::Show("Uploaded $([IO.Path]::GetFileName($outPath)) to $SharePointSiteUrl / $destFolder","Upload successful",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            } catch {
                [System.Windows.Forms.MessageBox]::Show("SharePoint upload failed: $($_.Exception.Message)","Upload error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            } finally {
                try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
            }

        } else {
            # Without PnP: provide guidance and copy path to clipboard
            [System.Windows.Forms.MessageBox]::Show("UsePnP switch not specified. The filtered JSON was exported to:`n$outPath`nTo upload to SharePoint please either supply the -UsePnP switch, or upload manually to $SharePointSiteUrl/$SharePointLibrary/$SharePointFolder.","Local export only",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            Set-Clipboard -Value $outPath
        }
    })

    # Double-click row to copy URL to clipboard
    $grid.add_CellDoubleClick({
        param($s,$e)
        try {
            $row = $grid.Rows[$e.RowIndex].DataBoundItem
            if ($row -and $row.URL) {
                Set-Clipboard -Value $row.URL
                [System.Windows.Forms.MessageBox]::Show("Copied URL to clipboard:`n$($row.URL)","Copied",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            }
        } catch {}
    })

    # Initial date defaults
    try {
        $td = [datetime]::Parse($jsonObj.TestMetadata.TestDateTime)
        $dtFrom.Value = $td.Date
        $dtTo.Value   = $td.Date
    } catch {
        $dtFrom.Value = (Get-Date).Date
        $dtTo.Value   = (Get-Date).Date
    }

    # Start with filter applied
    & $applyFilter

    # Show form
    $form.Add_Shown({ $form.Activate() })
    [void]$form.ShowDialog()
}

# ------------------------
# Example usage (uncomment and adapt when ready)
# ------------------------
# $jsonPath = "C:\Path\To\Deepe_HPWORKSTATION99_20251109_002758.json"
# Show-IntuneJsonGuiAndUpload -JsonPath $jsonPath -SharePointSiteUrl "https://yourtenant.sharepoint.com/sites/IntuneReports" -SharePointLibrary "Shared Documents" -SharePointFolder "Intune/Reports" -UsePnP
