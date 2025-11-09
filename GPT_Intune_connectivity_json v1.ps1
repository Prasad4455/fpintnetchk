# -------------------------
# GUI: Parse & Edit JSON results (append to end of your script)
# -------------------------

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-IntuneResultsGUI {
    param(
        [string]$JsonPath
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Helper: if no path provided, try to find the most recent JSON in script directory
    if (-not $JsonPath) {
        try {
            $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
        } catch {
            $scriptDir = Get-Location
        }
        $candidates = Get-ChildItem -Path $scriptDir -Filter *.json -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
        if ($candidates -and $candidates.Count -gt 0) {
            $JsonPath = $candidates[0].FullName
        } else {
            $JsonPath = (Get-ChildItem -Path (Get-Location) -Filter *.json -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
        }
    }

    if (-not (Test-Path $JsonPath)) {
        [System.Windows.Forms.MessageBox]::Show("JSON file not found: $JsonPath","File not found",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Read and parse JSON
    $raw = Get-Content -Raw -Path $JsonPath
    try {
        $json = $raw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to parse JSON file: $JsonPath`n$($_.Exception.Message)","Parse error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Build a combined table from known endpoint arrays inside ConnectivityResults
    $rows = @()
    if ($json.ConnectivityResults -ne $null) {
        $conn = $json.ConnectivityResults
        # iterate known arrays
        $possibleArrays = @('M365CommonEndpoints','IntuneEndpoints')
        foreach ($arrName in $possibleArrays) {
            if ($conn.PSObject.Properties.Name -contains $arrName) {
                $arr = $conn.$arrName
                if ($arr) {
                    foreach ($item in $arr) {
                        $rows += [PSCustomObject]@{
                            SourceArray    = $arrName
                            Category       = $item.Category
                            URL            = $item.URL
                            TestTimestamp  = $item.TestTimestamp
                            Mandatory      = $item.Mandatory
                            StatusCode     = $item.StatusCode
                            Status         = $item.Status
                            RegionNote     = $item.RegionNote
                            RawObject      = $item
                        }
                    }
                }
            }
        }
    }

    # Create a DataTable for DataGridView binding (editable)
    $dt = New-Object System.Data.DataTable
    $cols = @(
        @{n='SourceArray';t='string'},
        @{n='Category';t='string'},
        @{n='URL';t='string'},
        @{n='TestTimestamp';t='string'},
        @{n='Mandatory';t='string'},
        @{n='StatusCode';t='string'},
        @{n='Status';t='string'},
        @{n='RegionNote';t='string'}
    )
    foreach ($c in $cols) { $dt.Columns.Add($c.n,[string]) | Out-Null }
    foreach ($r in $rows) {
        $dr = $dt.NewRow()
        $dr['SourceArray'] = $r.SourceArray
        $dr['Category'] = $r.Category
        $dr['URL'] = $r.URL
        $dr['TestTimestamp'] = $r.TestTimestamp
        $dr['Mandatory'] = [string]$r.Mandatory
        $dr['StatusCode'] = [string]$r.StatusCode
        $dr['Status'] = $r.Status
        $dr['RegionNote'] = $r.RegionNote
        $dt.Rows.Add($dr)
    }

    # WinForms construction
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Intune Connectivity Results - Edit & Save"
    $form.Size = New-Object System.Drawing.Size(1100,700)
    $form.StartPosition = "CenterScreen"

    # Search box
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Text = "Search (URL / Category):"
    $lblSearch.Location = New-Object System.Drawing.Point(10,12)
    $lblSearch.AutoSize = $true
    $form.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location = New-Object System.Drawing.Point(150,8)
    $txtSearch.Size = New-Object System.Drawing.Size(300,22)
    $form.Controls.Add($txtSearch)

    # Filter combobox (All / Success / Failed)
    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = "Status Filter:"
    $lblFilter.Location = New-Object System.Drawing.Point(470,12)
    $lblFilter.AutoSize = $true
    $form.Controls.Add($lblFilter)

    $cmbFilter = New-Object System.Windows.Forms.ComboBox
    $cmbFilter.Location = New-Object System.Drawing.Point(550,8)
    $cmbFilter.Size = New-Object System.Drawing.Size(120,22)
    $cmbFilter.DropDownStyle = 'DropDownList'
    $cmbFilter.Items.AddRange(@('All','Success','Failed','Warning'))
    $cmbFilter.SelectedIndex = 0
    $form.Controls.Add($cmbFilter)

    # Show only source arrays checkbox (optional)
    $chkShowSource = New-Object System.Windows.Forms.CheckedListBox
    $chkShowSource.Location = New-Object System.Drawing.Point(690,8)
    $chkShowSource.Size = New-Object System.Drawing.Size(200,50)
    $chkShowSource.CheckOnClick = $true
    $chkShowSource.Items.AddRange(@('M365CommonEndpoints','IntuneEndpoints')) | Out-Null
    $chkShowSource.SetItemChecked(0,$true)
    $chkShowSource.SetItemChecked(1,$true)
    $form.Controls.Add($chkShowSource)

    # DataGridView
    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Location = New-Object System.Drawing.Point(10,65)
    $dgv.Size = New-Object System.Drawing.Size(1060,540)
    $dgv.AutoGenerateColumns = $true
    $dgv.ReadOnly = $false
    $dgv.AllowUserToAddRows = $false
    $dgv.AllowUserToDeleteRows = $false
    $dgv.DataSource = $dt
    $dgv.SelectionMode = 'FullRowSelect'
    $dgv.MultiSelect = $false
    $dgv.AutoSizeColumnsMode = 'Fill'
    $form.Controls.Add($dgv)

    # Buttons: Save, Save As, Refresh, Close
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save (overwrite)"
    $btnSave.Location = New-Object System.Drawing.Point(10,620)
    $btnSave.Size = New-Object System.Drawing.Size(140,30)
    $form.Controls.Add($btnSave)

    $btnSaveAs = New-Object System.Windows.Forms.Button
    $btnSaveAs.Text = "Save As..."
    $btnSaveAs.Location = New-Object System.Drawing.Point(160,620)
    $btnSaveAs.Size = New-Object System.Drawing.Size(100,30)
    $form.Controls.Add($btnSaveAs)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh (re-read file)"
    $btnRefresh.Location = New-Object System.Drawing.Point(270,620)
    $btnRefresh.Size = New-Object System.Drawing.Size(140,30)
    $form.Controls.Add($btnRefresh)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "Close"
    $btnClose.Location = New-Object System.Drawing.Point(420,620)
    $btnClose.Size = New-Object System.Drawing.Size(100,30)
    $form.Controls.Add($btnClose)

    # Info label
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(540,620)
    $lblInfo.Size = New-Object System.Drawing.Size(530,30)
    $lblInfo.Text = "Tip: edit the Status/StatusCode/RegionNote/Mandatory fields and click Save."
    $form.Controls.Add($lblInfo)

    # Filtering function
    function Apply-Filters {
        param($dtSource)
        $view = New-Object System.Data.DataView($dtSource)

        $filters = @()

        # status filter
        $sel = $cmbFilter.SelectedItem
        if ($sel -and $sel -ne 'All') {
            $escaped = $sel.Replace("'", "''")
            $filters += "Status = '$escaped'"
        }

        # search filter
        $q = $txtSearch.Text.Trim()
        if ($q -ne '') {
            $qEsc = $q.Replace("'", "''")
            $filters += "URL LIKE '%$qEsc%' OR Category LIKE '%$qEsc%'"
        }

        # source filter - checked items
        $checked = @()
        for ($i=0; $i -lt $chkShowSource.Items.Count; $i++) {
            if ($chkShowSource.GetItemChecked($i)) {
                $checked += $chkShowSource.Items[$i]
            }
        }
        if ($checked.Count -gt 0 -and $checked.Count -lt $chkShowSource.Items.Count) {
            $srcList = ($checked | ForEach-Object { "'$_'" }) -join ","
            $filters += "SourceArray IN ($srcList)"
        }

        if ($filters.Count -gt 0) {
            $view.RowFilter = ($filters -join " AND ")
        } else {
            $view.RowFilter = ""
        }

        $dgv.DataSource = $view
    }

    # Attach events
    $cmbFilter.Add_SelectedIndexChanged({ Apply-Filters -dtSource $dt })
    $txtSearch.Add_TextChanged({ Apply-Filters -dtSource $dt })
    $chkShowSource.Add_ItemCheck({
        Start-Sleep -Milliseconds 50
        Apply-Filters -dtSource $dt
    })

    $btnRefresh.Add_Click({
        try {
            $newRaw = Get-Content -Raw -Path $JsonPath
            $newJson = $newRaw | ConvertFrom-Json
            [void][System.Windows.Forms.MessageBox]::Show("Refreshing requires re-opening the GUI to ensure full fidelity. Click OK to re-open.","Refresh", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            $form.Close()
            Show-IntuneResultsGUI -JsonPath $JsonPath
            return
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to refresh from file: $($_.Exception.Message)","Refresh failed",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # Save logic: read current grid rows and update json object arrays
    $saveAction = {
        param($targetPath)
        try {
            # get the binding source (DataView or DataTable)
            $binding = $dgv.DataSource
            if ($binding -is [System.Data.DataView]) {
                $tableToSave = $binding.ToTable()
            } else {
                $tableToSave = $binding
            }

            # Reconstruct arrays grouped by SourceArray
            $grouped = @{}
            foreach ($r in $tableToSave.Rows) {
                # DataRow: use indexer
                $src = $r['SourceArray']
                if (-not $grouped.ContainsKey($src)) { $grouped[$src] = @() }

                # parse StatusCode safely
                $parsedInt = $null
                $successParse = $false
                try {
                    $tmp = 0
                    $successParse = [int]::TryParse([string]$r['StatusCode'], [ref]$tmp)
                    if ($successParse) { $parsedInt = [int]$tmp }
                } catch {
                    $parsedInt = $null
                }

                $obj = [ordered]@{
                    Category = $r['Category']
                    TestTimestamp = $r['TestTimestamp']
                    Mandatory = $r['Mandatory']
                    URL = $r['URL']
                    StatusCode = $parsedInt
                    Status = $r['Status']
                    RegionNote = $r['RegionNote']
                }

                $grouped[$src] += (New-Object PSObject -Property $obj)
            }

            # Update json object while preserving other fields
            $out = $json.PSObject | Select-Object -Property * | ForEach-Object { $_ }

            if ($out.PSObject.Properties.Name -contains 'ConnectivityResults') {
                foreach ($k in $grouped.Keys) {
                    $out.ConnectivityResults.$k = $grouped[$k]
                }

                # Optionally recalc Summary (safe)
                try {
                    $allEndpoints = @()
                    foreach ($prop in $out.ConnectivityResults.PSObject.Properties.Name) {
                        $val = $out.ConnectivityResults.$prop
                        if ($val -is [System.Array]) { $allEndpoints += $val }
                    }
                    $total = $allEndpoints.Count
                    $failed = ($allEndpoints | Where-Object { $_.Status -ne 'Success' }).Count
                    $success = $total - $failed
                    if ($out.PSObject.Properties.Name -contains 'Summary') {
                        $out.Summary.TestStatus = if ($failed -gt 0) { 'Failed' } else { 'Passed' }
                        $out.Summary.TotalEndpointsTested = $total
                        $out.Summary.FailedConnections = $failed
                        $out.Summary.SuccessfulConnections = $success
                    } else {
                        $summaryObj = [PSCustomObject]@{
                            TestStatus = if ($failed -gt 0) { 'Failed' } else { 'Passed' }
                            TotalEndpointsTested = $total
                            FailedConnections = $failed
                            SuccessfulConnections = $success
                        }
                        $out | Add-Member -MemberType NoteProperty -Name Summary -Value $summaryObj
                    }
                } catch {
                    # ignore recalculation errors
                }
            } else {
                # If no ConnectivityResults object in original file, create one
                $connObj = [PSCustomObject]@{}
                foreach ($k in $grouped.Keys) {
                    $connObj | Add-Member -MemberType NoteProperty -Name $k -Value $grouped[$k]
                }
                $out | Add-Member -MemberType NoteProperty -Name ConnectivityResults -Value $connObj
            }

            # Convert back to JSON and write
            $jsonText = $out | ConvertTo-Json -Depth 8 -Compress
            Set-Content -Path $targetPath -Value $jsonText -Force -Encoding UTF8

            [System.Windows.Forms.MessageBox]::Show("Saved: $targetPath","Success",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to save JSON: $($_.Exception.Message)","Save error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # Save button handler (overwrite)
    $btnSave.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show("Overwrite JSON file`n$JsonPath ?","Confirm Overwrite",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            & $saveAction $JsonPath
        }
    })

    # SaveAs button handler
    $btnSaveAs.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $sfd.FileName = (Split-Path -Leaf $JsonPath)
        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            & $saveAction $sfd.FileName
        }
    })

    $btnClose.Add_Click({ $form.Close() })

    # Initial filter apply
    Apply-Filters -dtSource $dt

    # Show form
    $form.Topmost = $true
    $form.Add_Shown({$form.Activate()})
    [void]$form.ShowDialog()
}


# Example usage:
# If your existing script wrote $JsonOutputPath, call:
# Show-IntuneResultsGUI -JsonPath $JsonOutputPath
#
# If you did not store the path, simply call:
# Show-IntuneResultsGUI
#
# Append one of the calls below at the end of your script (uncomment as appropriate):

# If your script variable is $JsonOutputPath or $outfile or $outFile, pass it here:
# Show-IntuneResultsGUI -JsonPath $JsonOutputPath
# Or auto-detect:
Show-IntuneResultsGUI
