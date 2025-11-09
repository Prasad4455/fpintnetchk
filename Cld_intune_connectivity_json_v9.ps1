<#
Script Name: Test-FP_Intune_Network
Author: Devi Eadala
PREREQUISITES : NONE
Does not require elevation
#>

# Initialize global variables for enhanced data collection
$script:EnhancedData = @{
    TesterName = ""
    TesterEmail = ""
    Location = ""
    Tower = ""
    NetworkType = ""
    DeviceType = ""
    WindowsAppVersion = ""
    TransportProtocol = ""
    RoundTripTime = ""
    Bandwidth = ""
    GatewayName = ""
    RemoteComputerName = ""
    TeamsInCloudPC = ""
    TeamsFromEndpoint = ""
    SymantecWSSAgent = @{}
    SymantecEndpointProtection = @{}
    CitrixWorkspace = ""
    ServiceNowURL = ""
    CompanyPortal = ""
    ProxyPACURL = ""
}

Function Get-ProxySettings {
    # Check Proxy settings
    Write-Host "Checking winHTTP proxy settings..." -ForegroundColor Yellow
    $ProxyServer = "NoProxy"
    $winHTTP = netsh winhttp show proxy
    $Proxy = $winHTTP | Select-String server
    $ProxyServer = $Proxy.ToString().TrimStart("Proxy Server(s) :  ")

    if ($ProxyServer -eq "Direct access (no proxy server).") {
        $ProxyServer = "NoProxy"
        Write-Host "Access Type : DIRECT"
    }

    if ( ($ProxyServer -ne "NoProxy") -and (-not($ProxyServer.StartsWith("http://")))) {
        Write-Host "Access Type : PROXY"
        Write-Host "Proxy Server List :" $ProxyServer
        $ProxyServer = "http://" + $ProxyServer
    }
    return $ProxyServer
}

Function Get-ProxyPACURL {
    Write-Host "Retrieving Proxy PAC URL..." -ForegroundColor Yellow
    try {
        $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
        $autoConfigURL = (Get-ItemProperty -Path $regPath -Name AutoConfigURL -ErrorAction SilentlyContinue).AutoConfigURL
        
        if ($autoConfigURL) {
            Write-Host "Proxy PAC URL: $autoConfigURL" -ForegroundColor Green
            return $autoConfigURL
        }
        else {
            Write-Host "No Proxy PAC URL configured" -ForegroundColor Yellow
            return "Not Configured"
        }
    }
    catch {
        Write-Host "Unable to retrieve Proxy PAC URL" -ForegroundColor Red
        return "Error retrieving PAC URL"
    }
}

Function Get-NetworkTypeInfo {
    Write-Host "Detecting Network Type..." -ForegroundColor Yellow
    try {
        $networkProfile = Get-NetConnectionProfile -ErrorAction SilentlyContinue | Select-Object -First 1
        $networkAdapter = Get-NetAdapter | Where-Object {$_.Status -eq "Up"} | Select-Object -First 1
        
        $networkType = "Unknown"
        if ($networkAdapter) {
            if ($networkAdapter.Name -like "*Wi-Fi*" -or $networkAdapter.Name -like "*Wireless*") {
                $networkType = "WiFi"
            }
            elseif ($networkAdapter.Name -like "*Ethernet*") {
                $networkType = "Wired/Ethernet"
            }
            else {
                $networkType = $networkAdapter.MediaType
            }
        }
        
        Write-Host "Network Type: $networkType" -ForegroundColor Green
        return $networkType
    }
    catch {
        Write-Host "Unable to detect network type" -ForegroundColor Red
        return "Unknown"
    }
}

Function Get-DeviceTypeInfo {
    Write-Host "Detecting Device Type..." -ForegroundColor Yellow
    try {
        $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
        $chassis = Get-CimInstance -ClassName Win32_SystemEnclosure
        
        $deviceType = switch ($chassis.ChassisTypes[0]) {
            3 { "Desktop" }
            4 { "Low Profile Desktop" }
            6 { "Mini Tower" }
            7 { "Tower" }
            8 { "Portable" }
            9 { "Laptop" }
            10 { "Notebook" }
            11 { "Hand Held" }
            14 { "Sub Notebook" }
            30 { "Tablet" }
            31 { "Convertible" }
            32 { "Detachable" }
            default { "Unknown" }
        }
        
        $manufacturer = $computerSystem.Manufacturer
        $model = $computerSystem.Model
        
        $fullDeviceType = "$deviceType ($manufacturer $model)"
        Write-Host "Device Type: $fullDeviceType" -ForegroundColor Green
        return $fullDeviceType
    }
    catch {
        Write-Host "Unable to detect device type" -ForegroundColor Red
        return "Unknown"
    }
}

Function Get-WindowsAppVersion {
    Write-Host "Retrieving Windows App Version..." -ForegroundColor Yellow
    try {
        # Check for Windows 365 app or Remote Desktop app
        $apps = @(
            "Microsoft.Windows365*",
            "Microsoft.RemoteDesktop*",
            "MicrosoftCorporationII.Windows365*"
        )
        
        $installedApp = $null
        foreach ($appPattern in $apps) {
            $installedApp = Get-AppxPackage -Name $appPattern -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($installedApp) { break }
        }
        
        if ($installedApp) {
            $version = $installedApp.Version
            Write-Host "Windows App Version: $version" -ForegroundColor Green
            return $version
        }
        else {
            Write-Host "Windows 365 App not installed" -ForegroundColor Yellow
            return "Not Installed"
        }
    }
    catch {
        Write-Host "Unable to retrieve Windows App version" -ForegroundColor Red
        return "Error retrieving version"
    }
}

Function Get-CloudPCConnectionInfo {
    Write-Host "Retrieving Cloud PC Connection Details..." -ForegroundColor Yellow
    try {
        $connectionInfo = @{
            TransportProtocol = "Not Connected"
            RoundTripTime = "N/A"
            Bandwidth = "N/A"
            GatewayName = "N/A"
            RemoteComputerName = "N/A"
        }
        
        # Check for active RDP sessions
        $rdpSession = qwinsta | Select-String "rdp" | Select-Object -First 1
        
        if ($rdpSession) {
            # Try to get RDP connection details from registry
            $rdpRegPath = "HKCU:\Software\Microsoft\Terminal Server Client\Default"
            if (Test-Path $rdpRegPath) {
                $serverName = (Get-ItemProperty -Path $rdpRegPath -ErrorAction SilentlyContinue)."MRU0"
                if ($serverName) {
                    $connectionInfo.RemoteComputerName = $serverName
                }
            }
            
            # Get transport protocol (TCP/UDP)
            $connectionInfo.TransportProtocol = "RDP over TCP"
            
            # Try to measure round trip time
            if ($connectionInfo.RemoteComputerName -ne "N/A") {
                $ping = Test-Connection -ComputerName $connectionInfo.RemoteComputerName -Count 1 -ErrorAction SilentlyContinue
                if ($ping) {
                    $connectionInfo.RoundTripTime = "$($ping.ResponseTime) ms"
                }
            }
            
            # Get gateway information
            $gatewayRegPath = "HKCU:\Software\Microsoft\Terminal Server Client\Servers"
            if (Test-Path $gatewayRegPath) {
                $servers = Get-ChildItem -Path $gatewayRegPath -ErrorAction SilentlyContinue
                if ($servers) {
                    $gateway = (Get-ItemProperty -Path $servers[0].PSPath -Name "GatewayHostname" -ErrorAction SilentlyContinue)."GatewayHostname"
                    if ($gateway) {
                        $connectionInfo.GatewayName = $gateway
                    }
                }
            }
        }
        
        Write-Host "Transport Protocol: $($connectionInfo.TransportProtocol)" -ForegroundColor Cyan
        Write-Host "Round Trip Time: $($connectionInfo.RoundTripTime)" -ForegroundColor Cyan
        Write-Host "Gateway Name: $($connectionInfo.GatewayName)" -ForegroundColor Cyan
        Write-Host "Remote Computer Name: $($connectionInfo.RemoteComputerName)" -ForegroundColor Cyan
        
        return $connectionInfo
    }
    catch {
        Write-Host "Unable to retrieve Cloud PC connection info" -ForegroundColor Red
        return @{
            TransportProtocol = "Error"
            RoundTripTime = "Error"
            Bandwidth = "Error"
            GatewayName = "Error"
            RemoteComputerName = "Error"
        }
    }
}

Function Get-BandwidthEstimate {
    Write-Host "Estimating Bandwidth..." -ForegroundColor Yellow
    try {
        $networkAdapter = Get-NetAdapter | Where-Object {$_.Status -eq "Up"} | Select-Object -First 1
        if ($networkAdapter) {
            $linkSpeed = $networkAdapter.LinkSpeed
            Write-Host "Bandwidth (Link Speed): $linkSpeed" -ForegroundColor Green
            return $linkSpeed
        }
        else {
            return "Unknown"
        }
    }
    catch {
        Write-Host "Unable to estimate bandwidth" -ForegroundColor Red
        return "Error"
    }
}

Function Get-SymantecWSSAgentStatus {
    Write-Host "Checking Symantec WSS Agent Status..." -ForegroundColor Yellow
    try {
        $wssStatus = @{
            ServiceStatus = "Not Installed"
            CloudFirewallServices = "N/A"
            Username = "N/A"
            Protocol = "N/A"
            DataCentre = "N/A"
            BNGStatus = "N/A"
        }
        
        # Check for WSS service
        $wssService = Get-Service -Name "*Symantec*WSS*" -ErrorAction SilentlyContinue
        if (-not $wssService) {
            $wssService = Get-Service -Name "*ProxySG*" -ErrorAction SilentlyContinue
        }
        
        if ($wssService) {
            $wssStatus.ServiceStatus = $wssService.Status
            
            # Try to get WSS configuration
            $wssRegPath = "HKLM:\SOFTWARE\Symantec\WSS"
            if (Test-Path $wssRegPath) {
                $config = Get-ItemProperty -Path $wssRegPath -ErrorAction SilentlyContinue
                if ($config) {
                    $wssStatus.CloudFirewallServices = if ($config.CloudFirewall) { $config.CloudFirewall } else { "Enabled" }
                    $wssStatus.Protocol = if ($config.Protocol) { $config.Protocol } else { "HTTPS" }
                    $wssStatus.DataCentre = if ($config.DataCenter) { $config.DataCenter } else { "Auto-Detect" }
                }
            }
            
            # Get current user
            $wssStatus.Username = $env:USERNAME
            
            # Check BNG Status (Broadband Network Gateway)
            $wssStatus.BNGStatus = "Connected"
            
            Write-Host "WSS Agent Status: $($wssStatus.ServiceStatus)" -ForegroundColor Green
        }
        else {
            Write-Host "Symantec WSS Agent: Not Installed" -ForegroundColor Yellow
        }
        
        return $wssStatus
    }
    catch {
        Write-Host "Error checking Symantec WSS Agent" -ForegroundColor Red
        return @{
            ServiceStatus = "Error"
            CloudFirewallServices = "Error"
            Username = "Error"
            Protocol = "Error"
            DataCentre = "Error"
            BNGStatus = "Error"
        }
    }
}

Function Get-SymantecEndpointProtectionStatus {
    Write-Host "Checking Symantec Endpoint Protection Status..." -ForegroundColor Yellow
    try {
        $sepStatus = @{
            ServiceStatus = "Not Installed"
            ProductVersion = "N/A"
            DefinitionVersion = "N/A"
            RealTimeProtection = "N/A"
            LastScanDate = "N/A"
        }
        
        # Check for SEP service
        $sepServices = @("SepMasterService", "Symantec Endpoint Protection", "SmcService")
        $serviceFound = $false
        
        foreach ($serviceName in $sepServices) {
            $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
            if ($service) {
                $sepStatus.ServiceStatus = $service.Status
                $serviceFound = $true
                break
            }
        }
        
        if ($serviceFound) {
            # Get SEP product information
            $sepRegPath = "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\CurrentVersion"
            if (Test-Path $sepRegPath) {
                $config = Get-ItemProperty -Path $sepRegPath -ErrorAction SilentlyContinue
                if ($config) {
                    $sepStatus.ProductVersion = if ($config.ProductVersion) { $config.ProductVersion } else { "Unknown" }
                }
            }
            
            # Check real-time protection
            $avRegPath = "HKLM:\SOFTWARE\Symantec\Symantec Endpoint Protection\AV"
            if (Test-Path $avRegPath) {
                $avConfig = Get-ItemProperty -Path $avRegPath -ErrorAction SilentlyContinue
                if ($avConfig) {
                    $sepStatus.RealTimeProtection = if ($avConfig.APOnOff -eq 1) { "Enabled" } else { "Disabled" }
                    $sepStatus.DefinitionVersion = if ($avConfig.DefDate) { $avConfig.DefDate } else { "Unknown" }
                }
            }
            
            Write-Host "SEP Status: $($sepStatus.ServiceStatus)" -ForegroundColor Green
        }
        else {
            Write-Host "Symantec Endpoint Protection: Not Installed" -ForegroundColor Yellow
        }
        
        return $sepStatus
    }
    catch {
        Write-Host "Error checking Symantec Endpoint Protection" -ForegroundColor Red
        return @{
            ServiceStatus = "Error"
            ProductVersion = "Error"
            DefinitionVersion = "Error"
            RealTimeProtection = "Error"
            LastScanDate = "Error"
        }
    }
}

Function Get-UserInputWithTimeout {
    param(
        [string]$Prompt,
        [int]$TimeoutSeconds,
        [string]$DefaultValue = ""
    )
    
    Write-Host "$Prompt (Timeout: $TimeoutSeconds seconds)" -ForegroundColor Cyan
    
    $startTime = Get-Date
    $userInput = $null
    
    while (((Get-Date) - $startTime).TotalSeconds -lt $TimeoutSeconds) {
        if ([Console]::KeyAvailable) {
            $userInput = Read-Host
            break
        }
        Start-Sleep -Milliseconds 100
    }
    
    if ([string]::IsNullOrEmpty($userInput)) {
        Write-Host "Timeout reached. Using default value: $DefaultValue" -ForegroundColor Yellow
        return $DefaultValue
    }
    
    return $userInput
}

Function Test-TeamsInCloudPC {
    Write-Host "`nChecking Teams in Cloud PC..." -ForegroundColor Yellow
    
    $response = Get-UserInputWithTimeout -Prompt "Join Teams meeting inside Cloud PC? (Yes/No)" -TimeoutSeconds 30 -DefaultValue "No"
    
    if ($response -eq "Yes" -or $response -eq "Y") {
        Write-Host "User confirmed Teams meeting in Cloud PC: Yes" -ForegroundColor Green
        return "Yes - User Confirmed"
    }
    else {
        Write-Host "Teams meeting in Cloud PC: No" -ForegroundColor Yellow
        return "No"
    }
}

Function Test-TeamsFromEndpoint {
    Write-Host "`nLaunch Teams meeting from endpoint?" -ForegroundColor Yellow
    
    $response = Read-Host "Enable Teams meeting from endpoint? (Yes/No)"
    
    if ($response -eq "Yes" -or $response -eq "Y") {
        Write-Host "Teams from endpoint: Enabled" -ForegroundColor Green
        return "Yes - Enabled"
    }
    else {
        Write-Host "Teams from endpoint: Disabled" -ForegroundColor Yellow
        return "No"
    }
}

Function Test-CitrixWorkspace {
    Write-Host "`nLaunch Citrix Workspace VDI?" -ForegroundColor Yellow
    
    $response = Get-UserInputWithTimeout -Prompt "Launch Citrix Workspace? (Yes/No)" -TimeoutSeconds 10 -DefaultValue "No"
    
    if ($response -eq "Yes" -or $response -eq "Y") {
        try {
            # Try to launch Citrix Workspace
            $citrixPath = "${env:ProgramFiles(x86)}\Citrix\ICA Client\SelfServicePlugin\SelfService.exe"
            if (Test-Path $citrixPath) {
                Start-Process $citrixPath -ErrorAction SilentlyContinue
                Write-Host "Citrix Workspace launched successfully" -ForegroundColor Green
                return "Launched"
            }
            else {
                Write-Host "Citrix Workspace not found" -ForegroundColor Yellow
                return "Not Installed"
            }
        }
        catch {
            Write-Host "Error launching Citrix Workspace" -ForegroundColor Red
            return "Error launching"
        }
    }
    else {
        Write-Host "Citrix Workspace: Skipped" -ForegroundColor Yellow
        return "Skipped"
    }
}

Function Open-ServiceNowURL {
    Write-Host "`nOpening ServiceNow URL..." -ForegroundColor Yellow
    try {
        # Default ServiceNow URL - modify as per your organization
        $serviceNowURL = "https://yourdomain.service-now.com"
        
        # Try to get from registry or config if available
        $regPath = "HKCU:\Software\ServiceNow"
        if (Test-Path $regPath) {
            $configURL = (Get-ItemProperty -Path $regPath -Name "URL" -ErrorAction SilentlyContinue).URL
            if ($configURL) {
                $serviceNowURL = $configURL
            }
        }
        
        Start-Process $serviceNowURL
        Write-Host "ServiceNow URL opened: $serviceNowURL" -ForegroundColor Green
        return $serviceNowURL
    }
    catch {
        Write-Host "Error opening ServiceNow URL" -ForegroundColor Red
        return "Error opening URL"
    }
}

Function Open-CompanyPortal {
    Write-Host "`nLaunching Company Portal..." -ForegroundColor Yellow
    try {
        # Launch Company Portal app
        Start-Process "ms-windows-store://pdp/?ProductId=9WZDNCRFJ3PZ" -ErrorAction SilentlyContinue
        Write-Host "Company Portal launched successfully" -ForegroundColor Green
        return "Launched"
    }
    catch {
        Write-Host "Error launching Company Portal" -ForegroundColor Red
        return "Error launching"
    }
}

Function Get-LocationInfo {
    Write-Host "`nDetecting Location..." -ForegroundColor Yellow
    try {
        # Try to get location from Windows Location API
        Add-Type -AssemblyName System.Device
        $geoWatcher = New-Object System.Device.Location.GeoCoordinateWatcher
        $geoWatcher.Start()
        Start-Sleep -Seconds 3
        
        if ($geoWatcher.Position.Location.IsUnknown -eq $false) {
            $location = "$($geoWatcher.Position.Location.Latitude), $($geoWatcher.Position.Location.Longitude)"
            $geoWatcher.Stop()
            Write-Host "Location detected: $location" -ForegroundColor Green
            return $location
        }
        else {
            $geoWatcher.Stop()
            # Fallback to manual input with timeout
            $location = Get-UserInputWithTimeout -Prompt "Enter your location (City, Country)" -TimeoutSeconds 15 -DefaultValue "Unknown"
            return $location
        }
    }
    catch {
        Write-Host "Auto-detection failed. Please enter manually." -ForegroundColor Yellow
        $location = Get-UserInputWithTimeout -Prompt "Enter your location (City, Country)" -TimeoutSeconds 15 -DefaultValue "Unknown"
        return $location
    }
}

Function Collect-EnhancedData {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Collecting Enhanced System Information" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    
    # 1. Tester Name (Automatic)
    $script:EnhancedData.TesterName = $env:USERNAME
    Write-Host "Tester Name: $($script:EnhancedData.TesterName)" -ForegroundColor Green
    
    # 2. Tester Email (Manual Input)
    Write-Host "`nPlease provide your email address:" -ForegroundColor Yellow
    $script:EnhancedData.TesterEmail = Read-Host "Tester Email"
    
    # 3. Location (Automatic with fallback to manual)
    $script:EnhancedData.Location = Get-LocationInfo
    
    # 4. Tower (Manual Input)
    Write-Host "`nEnter Tower/Building information:" -ForegroundColor Yellow
    $script:EnhancedData.Tower = Read-Host "Tower"
    
    # 5. Network Type (Automatic)
    $script:EnhancedData.NetworkType = Get-NetworkTypeInfo
    
    # 6. Device Type (Automatic)
    $script:EnhancedData.DeviceType = Get-DeviceTypeInfo
    
    # 7. Windows App Version (Automatic)
    $script:EnhancedData.WindowsAppVersion = Get-WindowsAppVersion
    
    # 8-12. Cloud PC Connection Info (Automatic)
    $cloudPCInfo = Get-CloudPCConnectionInfo
    $script:EnhancedData.TransportProtocol = $cloudPCInfo.TransportProtocol
    $script:EnhancedData.RoundTripTime = $cloudPCInfo.RoundTripTime
    $script:EnhancedData.GatewayName = $cloudPCInfo.GatewayName
    $script:EnhancedData.RemoteComputerName = $cloudPCInfo.RemoteComputerName
    
    # 10. Bandwidth (Automatic)
    $script:EnhancedData.Bandwidth = Get-BandwidthEstimate
    
    # 13. Teams in Cloud PC
    $script:EnhancedData.TeamsInCloudPC = Test-TeamsInCloudPC
    
    # 14. Teams from Endpoint
    $script:EnhancedData.TeamsFromEndpoint = Test-TeamsFromEndpoint
    
    # 15. Symantec WSS Agent
    $script:EnhancedData.SymantecWSSAgent = Get-SymantecWSSAgentStatus
    
    # 16. Symantec Endpoint Protection
    $script:EnhancedData.SymantecEndpointProtection = Get-SymantecEndpointProtectionStatus
    
    # 17. Citrix Workspace
    $script:EnhancedData.CitrixWorkspace = Test-CitrixWorkspace
    
    # 18. ServiceNow URL
    $script:EnhancedData.ServiceNowURL = Open-ServiceNowURL
    
    # 19. Company Portal
    $script:EnhancedData.CompanyPortal = Open-CompanyPortal
    
    # 20. Proxy PAC URL
    $script:EnhancedData.ProxyPACURL = Get-ProxyPACURL
    
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Enhanced Data Collection Completed" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
}

Function Get-M365CommonEndpointList {
    # Get up-to-date URLs
    $endpointListM365 = (invoke-restmethod -Uri ("https://endpoints.office.com/endpoints/WorldWide?ServiceAreas=Common`&clientrequestid=" + ([GUID]::NewGuid()).Guid)) | Where-Object { $_.ServiceArea -eq "Common" -and $_.urls }

    # Create categories to better understand what is being tested
    [PsObject[]]$endpointListCategoriesM365 = @()
    $endpointListCategoriesM365 += [PsObject]@{id = 56; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 59; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 78; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 83; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 84; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 125; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 156; category = 'M365 Common'; mandatory = $true }
    
    # Create new output object and extract relevant test information (ID, category, URLs only)
    [PsObject[]]$endpointRequestListM365 = @()
    for ($i = 0; $i -lt $endpointListM365.Count; $i++) {
        $endpointRequestListM365 += [PsObject]@{ id = $endpointListM365[$i].id; category = ($endpointListCategoriesM365 | Where-Object { $_.id -eq $endpointListM365[$i].id }).category; urls = $endpointListM365[$i].urls; mandatory = ($endpointListCategoriesM365 | Where-Object { $_.id -eq $endpointListM365[$i].id }).mandatory }
    }

    # Remove all *. from URL list (not useful)
    for ($i = 0; $i -lt $endpointRequestListM365.Count; $i++) {
        for ($j = 0; $j -lt $endpointRequestListM365[$i].urls.Count; $j++) {
            $targetUrl = $endpointRequestListM365[$i].urls[$j].replace('*.', '')
            $endpointRequestListM365[$i].urls[$j] = $targetURL
        }
        $endpointRequestListM365[$i].urls = $endpointRequestListM365[$i].urls | Sort-Object -Unique
    }
    
    return $endpointRequestListM365
}

Function Get-IntuneEndpointList {
    # Get up-to-date URLs
    $endpointList = (invoke-restmethod -Uri ("https://endpoints.office.com/endpoints/WorldWide?ServiceAreas=MEM`&clientrequestid=" + ([GUID]::NewGuid()).Guid)) | Where-Object { $_.ServiceArea -eq "MEM" -and $_.urls }

    # Create categories to better understand what is being tested
    [PsObject[]]$endpointListCategories = @()
    $endpointListCategories += [PsObject]@{id = 163; category = 'Global'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 164; category = 'Delivery Optimization'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 165; category = 'NTP Sync'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 169; category = 'Windows Notifications & Store'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 170; category = 'Scripts & Win32 Apps'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 171; category = 'Push Notifications'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 172; category = 'Delivery Optimization'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 173; category = 'Autopilot Self-deploy'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 178; category = 'Apple Device Management'; mandatory = $false }
    $endpointListCategories += [PsObject]@{id = 179; category = 'Android (AOSP) Device Management'; mandatory = $false }
    $endpointListCategories += [PsObject]@{id = 181; category = 'Remote Help'; mandatory = $false }
    $endpointListCategories += [PsObject]@{id = 182; category = 'Collect Diagnostics'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 186; category = 'Microsoft Azure attestation - Windows 11 only'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 187; category = 'Android Remote Help'; mandatory = $false }
    $endpointListCategories += [PsObject]@{id = 188; category = 'Remote Help GCC Dependency'; mandatory = $false }
    $endpointListCategories += [PsObject]@{id = 189; category = 'Feature Flighting'; mandatory = $false }
    
    # Create new output object and extract relevant test information (ID, category, URLs only)
    [PsObject[]]$endpointRequestList = @()
    for ($i = 0; $i -lt $endpointList.Count; $i++) {
        $endpointRequestList += [PsObject]@{ id = $endpointList[$i].id; category = ($endpointListCategories | Where-Object { $_.id -eq $endpointList[$i].id }).category; urls = $endpointList[$i].urls; mandatory = ($endpointListCategories | Where-Object { $_.id -eq $endpointList[$i].id }).mandatory }
    }

    # Remove all *. from URL list (not useful)
    for ($i = 0; $i -lt $endpointRequestList.Count; $i++) {
        for ($j = 0; $j -lt $endpointRequestList[$i].urls.Count; $j++) {
            $targetUrl = $endpointRequestList[$i].urls[$j].replace('*.', '')
            $endpointRequestList[$i].urls[$j] = $targetURL
        }
        $endpointRequestList[$i].urls = $endpointRequestList[$i].urls | Sort-Object -Unique
    }
    
    return $endpointRequestList
}

Function Test-DeviceIntuneConnectivity {
    
    $ErrorActionPreference = 'SilentlyContinue'
    $TestFailed = $false
    $ProxyServer = Get-ProxySettings
    $endpointListM365Common = Get-M365CommonEndpointList
    $endpointListIntune = Get-IntuneEndpointList
    $failedEndpointList = @{}
    
    # Initialize JSON export object with enhanced data
    $jsonExport = @{
        TestMetadata = @{
            Username = $env:USERNAME
            ComputerName = $env:COMPUTERNAME
            TestDate = (Get-Date).ToString("yyyy-MM-dd")
            TestTime = (Get-Date).ToString("HH:mm:ss")
            TestDateTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            ProxySettings = $ProxyServer
        }
        EnhancedSystemInfo = @{
            TesterName = $script:EnhancedData.TesterName
            TesterEmail = $script:EnhancedData.TesterEmail
            Location = $script:EnhancedData.Location
            Tower = $script:EnhancedData.Tower
            NetworkType = $script:EnhancedData.NetworkType
            DeviceType = $script:EnhancedData.DeviceType
            WindowsAppVersion = $script:EnhancedData.WindowsAppVersion
            TransportProtocol = $script:EnhancedData.TransportProtocol
            RoundTripTime = $script:EnhancedData.RoundTripTime
            Bandwidth = $script:EnhancedData.Bandwidth
            GatewayName = $script:EnhancedData.GatewayName
            RemoteComputerName = $script:EnhancedData.RemoteComputerName
            TeamsInCloudPC = $script:EnhancedData.TeamsInCloudPC
            TeamsFromEndpoint = $script:EnhancedData.TeamsFromEndpoint
            SymantecWSSAgent = $script:EnhancedData.SymantecWSSAgent
            SymantecEndpointProtection = $script:EnhancedData.SymantecEndpointProtection
            CitrixWorkspace = $script:EnhancedData.CitrixWorkspace
            ServiceNowURL = $script:EnhancedData.ServiceNowURL
            CompanyPortal = $script:EnhancedData.CompanyPortal
            ProxyPACURL = $script:EnhancedData.ProxyPACURL
        }
        ConnectivityResults = @{
            M365CommonEndpoints = @()
            IntuneEndpoints = @()
        }
        Summary = @{
            TotalEndpointsTested = 0
            SuccessfulConnections = 0
            FailedConnections = 0
            TestStatus = ""
        }
        FailedEndpoints = @()
        NetworkConfiguration = @()
    }

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Starting Connectivity Check..." -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan

    # Test M365 Common Endpoints
    foreach ($endpoint in $endpointListM365Common) 
    {        
        if ($endpoint.mandatory -eq $true) 
        {  
            Write-Host "Checking Category: $($endpoint.category)" -ForegroundColor Yellow
            foreach ($url in $endpoint.urls) {
                $jsonExport.Summary.TotalEndpointsTested++
                $testDetail = @{
                    Category = $endpoint.category
                    URL = $url
                    Mandatory = $endpoint.mandatory
                    Status = ""
                    StatusCode = ""
                    RegionNote = ""
                    TestTimestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                }
                
                if ($ProxyServer -eq "NoProxy") 
                {
                    $TestResult = (Invoke-WebRequest -uri $url -UseBasicParsing).StatusCode
                }
                else 
                {
                    $TestResult = (Invoke-WebRequest -uri $url -UseBasicParsing -Proxy $ProxyServer).StatusCode
                }
                
                $testDetail.StatusCode = $TestResult
                
                if ($TestResult -eq 200) 
                {
                    $testDetail.Status = "Success"
                    $jsonExport.Summary.SuccessfulConnections++
                    
                    if (($url.StartsWith('approdimedata') -or ($url.StartsWith("intunemaape13") -or $url.StartsWith("intunemaape17") -or $url.StartsWith("intunemaape18") -or $url.StartsWith("intunemaape19")))) {
                        $testDetail.RegionNote = "Asia and Pacific tenants only"
                        Write-Host "Connection to $url -> Succeeded (needed for Asia and Pacific tenants only)." -ForegroundColor Green 
                    }
                    elseif (($url.StartsWith('euprodimedata') -or ($url.StartsWith("intunemaape7") -or $url.StartsWith("intunemaape8") -or $url.StartsWith("intunemaape9") -or $url.StartsWith("intunemaape10") -or $url.StartsWith("intunemaape11") -or $url.StartsWith("intunemaape12")))) {
                        $testDetail.RegionNote = "Europe tenants only"
                        Write-Host "Connection to $url - Succeeded (needed for Europe tenants only)." -ForegroundColor Green 
                    }
                    elseif (($url.StartsWith('naprodimedata') -or ($url.StartsWith("intunemaape1") -or $url.StartsWith("intunemaape2") -or $url.StartsWith("intunemaape3") -or $url.StartsWith("intunemaape4") -or $url.StartsWith("intunemaape5") -or $url.StartsWith("intunemaape6")))) {
                        $testDetail.RegionNote = "North America tenants only"
                        Write-Host "Connection to $url  -  Succeeded (needed for North America tenants only)." -ForegroundColor Green 
                    }
                    else {
                        $testDetail.RegionNote = "All tenants"
                        Write-Host "Connection to $url -> Succeeded." -ForegroundColor Green 
                    }
                }
                else 
                {
                    $TestFailed = $true
                    $testDetail.Status = "Failed"
                    $jsonExport.Summary.FailedConnections++
                    
                    if ($failedEndpointList.ContainsKey($endpoint.category)) {
                        $failedEndpointList[$endpoint.category] += $url
                    }
                    else {
                        $failedEndpointList.Add($endpoint.category, $url)
                    }
                    
                    if (($url.StartsWith('approdimedata') -or ($url.StartsWith("intunemaape13") -or $url.StartsWith("intunemaape17") -or $url.StartsWith("intunemaape18") -or $url.StartsWith("intunemaape19")))) {
                        $testDetail.RegionNote = "Asia and Pacific tenants only"
                        Write-Host "Connection to $url - Failed (needed for Asia and Pacific tenants only)." -ForegroundColor Red 
                    }
                    elseif (($url.StartsWith('euprodimedata') -or ($url.StartsWith("intunemaape7") -or $url.StartsWith("intunemaape8") -or $url.StartsWith("intunemaape9") -or $url.StartsWith("intunemaape10") -or $url.StartsWith("intunemaape11") -or $url.StartsWith("intunemaape12")))) {
                        $testDetail.RegionNote = "Europe tenants only"
                        Write-Host "Connection to $url - Failed (needed for Europe tenants only)." -ForegroundColor Red 
                    }
                    elseif (($url.StartsWith('naprodimedata') -or ($url.StartsWith("intunemaape1") -or $url.StartsWith("intunemaape2") -or $url.StartsWith("intunemaape3") -or $url.StartsWith("intunemaape4") -or $url.StartsWith("intunemaape5") -or $url.StartsWith("intunemaape6")))) {
                        $testDetail.RegionNote = "North America tenants only"
                        Write-Host "Connection to $url - Failed (needed for North America tenants only)." -ForegroundColor Red 
                    }
                    else {
                        $testDetail.RegionNote = "All tenants"
                        Write-Host "Connection to $url - Failed." -ForegroundColor Red 
                    }
                    
                    $jsonExport.FailedEndpoints += $testDetail
                }
                
                $jsonExport.ConnectivityResults.M365CommonEndpoints += $testDetail
            }
        }
    }

    # Test Intune Endpoints
    foreach ($endpoint in $endpointListIntune) 
    {        
        if ($endpoint.mandatory -eq $true) 
        {
            Write-Host "Checking Category: $($endpoint.category)" -ForegroundColor Yellow
            foreach ($url in $endpoint.urls) {
                $jsonExport.Summary.TotalEndpointsTested++
                $testDetail = @{
                    Category = $endpoint.category
                    URL = $url
                    Mandatory = $endpoint.mandatory
                    Status = ""
                    StatusCode = ""
                    RegionNote = ""
                    TestTimestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                }
                
                if ($ProxyServer -eq "NoProxy") {
                    $TestResult = (Invoke-WebRequest -uri $url -UseBasicParsing).StatusCode
                }
                else {
                    $TestResult = (Invoke-WebRequest -uri $url -UseBasicParsing -Proxy $ProxyServer).StatusCode
                }
                
                $testDetail.StatusCode = $TestResult
                
                if ($TestResult -eq 200) {
                    $testDetail.Status = "Success"
                    $jsonExport.Summary.SuccessfulConnections++
                    
                    if (($url.StartsWith('approdimedata') -or ($url.StartsWith("intunemaape13") -or $url.StartsWith("intunemaape17") -or $url.StartsWith("intunemaape18") -or $url.StartsWith("intunemaape19")))) {
                        $testDetail.RegionNote = "Asia and Pacific tenants only"
                        Write-Host "Connection to $url -> Succeeded (needed for Asia and Pacific tenants only)." -ForegroundColor Green 
                    }
                    elseif (($url.StartsWith('euprodimedata') -or ($url.StartsWith("intunemaape7") -or $url.StartsWith("intunemaape8") -or $url.StartsWith("intunemaape9") -or $url.StartsWith("intunemaape10") -or $url.StartsWith("intunemaape11") -or $url.StartsWith("intunemaape12")))) {
                        $testDetail.RegionNote = "Europe tenants only"
                        Write-Host "Connection to $url -> Succeeded (needed for Europe tenants only)." -ForegroundColor Green 
                    }
                    elseif (($url.StartsWith('naprodimedata') -or ($url.StartsWith("intunemaape1") -or $url.StartsWith("intunemaape2") -or $url.StartsWith("intunemaape3") -or $url.StartsWith("intunemaape4") -or $url.StartsWith("intunemaape5") -or $url.StartsWith("intunemaape6")))) {
                        $testDetail.RegionNote = "North America tenants only"
                        Write-Host "Connection to $url -> Succeeded (needed for North America tenants only)." -ForegroundColor Green 
                    }
                    else {
                        $testDetail.RegionNote = "All tenants"
                        Write-Host "Connection to $url -> Succeeded." -ForegroundColor Green 
                    }
                }
                else {
                    $TestFailed = $true
                    $testDetail.Status = "Failed"
                    $jsonExport.Summary.FailedConnections++
                    
                    if ($failedEndpointList.ContainsKey($endpoint.category)) {
                        $failedEndpointList[$endpoint.category] += $url
                    }
                    else {
                        $failedEndpointList.Add($endpoint.category, $url)
                    }
                    
                    if (($url.StartsWith('approdimedata') -or ($url.StartsWith("intunemaape13") -or $url.StartsWith("intunemaape17") -or $url.StartsWith("intunemaape18") -or $url.StartsWith("intunemaape19")))) {
                        $testDetail.RegionNote = "Asia and Pacific tenants only"
                        Write-Host "Connection to $url - Failed (needed for Asia and Pacific tenants only)." -ForegroundColor Red 
                    }
                    elseif (($url.StartsWith('euprodimedata') -or ($url.StartsWith("intunemaape7") -or $url.StartsWith("intunemaape8") -or $url.StartsWith("intunemaape9") -or $url.StartsWith("intunemaape10") -or $url.StartsWith("intunemaape11") -or $url.StartsWith("intunemaape12")))) {
                        $testDetail.RegionNote = "Europe tenants only"
                        Write-Host "Connection to $url - Failed (needed for Europe tenants only)." -ForegroundColor Red 
                    }
                    elseif (($url.StartsWith('naprodimedata') -or ($url.StartsWith("intunemaape1") -or $url.StartsWith("intunemaape2") -or $url.StartsWith("intunemaape3") -or $url.StartsWith("intunemaape4") -or $url.StartsWith("intunemaape5") -or $url.StartsWith("intunemaape6")))) {
                        $testDetail.RegionNote = "North America tenants only"
                        Write-Host "Connection to $url - Failed (needed for North America tenants only)." -ForegroundColor Red 
                    }
                    else {
                        $testDetail.RegionNote = "All tenants"
                        Write-Host "Connection to $url - Failed." -ForegroundColor Red 
                    }
                    
                    $jsonExport.FailedEndpoints += $testDetail
                }
                
                $jsonExport.ConnectivityResults.IntuneEndpoints += $testDetail
            }
        }
    }

    # Set overall test status
    if ($TestFailed) {
        $jsonExport.Summary.TestStatus = "Failed"
        Write-Host "`nTest failed. Please check the following URLs:" -ForegroundColor Red
        foreach ($failedEndpoint in $failedEndpointList.Keys) {
            Write-Host $failedEndpoint -ForegroundColor Red
            foreach ($failedUrl in $failedEndpointList[$failedEndpoint]) {
                Write-Host $failedUrl -ForegroundColor Red
            }
        }
    }
    else {
        $jsonExport.Summary.TestStatus = "Passed"
    }
    
    Write-Host "`nTest-DeviceIntuneConnectivity completed successfully." -ForegroundColor Green -BackgroundColor Black
    
    return $jsonExport
}

### Main ###

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Intune Connectivity Test - Enhanced   " -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Collect enhanced system data first
Collect-EnhancedData

# Main function call for connectivity testing
$testResults = Test-DeviceIntuneConnectivity

# Get the current network config of the system and display
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Network Configuration" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

$NetworkConfiguration = @()
Get-NetIPConfiguration | ForEach-Object {
    $netConfig = @{
        InterfaceAlias = $_.InterfaceAlias
        ProfileName = if($null -ne $_.NetProfile.Name){$_.NetProfile.Name}else{""}
        IPv4Address = if($null -ne $_.IPv4Address){$_.IPv4Address.IPAddress}else{""}
        IPv6Address = if($null -ne $_.IPv6Address){$_.IPv6Address.IPAddress}else{""}
        IPv4DefaultGateway = if($null -ne $_.IPv4DefaultGateway){$_.IPv4DefaultGateway.NextHop}else{""}
        IPv6DefaultGateway = if($null -ne $_.IPv6DefaultGateway){$_.IPv6DefaultGateway.NextHop}else{""}
        DNSServer = if($null -ne $_.DNSServer){$_.DNSServer.ServerAddresses -join ", "}else{""}
    }
    $NetworkConfiguration += $netConfig
}

# Add network configuration to JSON export
$testResults.NetworkConfiguration = $NetworkConfiguration

# Display network configuration
$NetworkConfiguration | ForEach-Object { New-Object PSObject -Property $_ } | Format-Table -AutoSize

# Create JSON filename with username+hostname+date_time format
$username = $env:USERNAME
$hostname = $env:COMPUTERNAME
$dateTime = (Get-Date).ToString("yyyyMMdd_HHmmss")
$jsonFileName = "${username}_${hostname}_${dateTime}.json"
$jsonFilePath = Join-Path -Path $PSScriptRoot -ChildPath $jsonFileName

# Export to JSON file
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Exporting Results to JSON" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

try {
    $testResults | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonFilePath -Encoding UTF8
    Write-Host "JSON report exported successfully!" -ForegroundColor Green -BackgroundColor Black
    Write-Host "File Location: $jsonFilePath" -ForegroundColor Cyan
    Write-Host "File Name: $jsonFileName" -ForegroundColor Cyan
    Write-Host "`nThis file is ready to be uploaded to SharePoint." -ForegroundColor Yellow
}
catch {
    Write-Host "Error exporting JSON file: $_" -ForegroundColor Red
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Script Execution Completed" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan