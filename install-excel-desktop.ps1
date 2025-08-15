# ==============================================================================
#  IPG Taxonomy Extractor - Excel Desktop Installation Script
#  Self-installation script for individual users
#  Version: 1.0.0
#  Author: IPG Mediabrands Technology Team
# ==============================================================================

param(
    [switch]$Uninstall,
    [switch]$Quiet,
    [string]$ManifestUrl = "https://ipg-taxonomy-extractor-addin.wookstar.com/manifest.xml"
)

# Script Configuration
$Script:Config = @{
    Version = "1.0.0"
    AddInName = "IPG Taxonomy Extractor"
    AddInId = "f3b3c5d7-8a2b-4c9e-9f1a-2d3e4f5a6b7c"
    ProductionUrl = "https://ipg-taxonomy-extractor-addin.wookstar.com"
    SupportUrl = "https://github.com/henkisdabro/taxonomy-extractor-officeaddin"
    RequiredOfficeVersion = "16.0"
}

# ==============================================================================
#  UI FUNCTIONS - Enhanced User Experience
# ==============================================================================

function Write-Header {
    if (-not $Quiet) {
        Clear-Host
        Write-Host ""
        Write-Host "╔═══════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║                    IPG TAXONOMY EXTRACTOR - EXCEL DESKTOP                    ║" -ForegroundColor Cyan
        Write-Host "║                           Installation Assistant v$($Script:Config.Version)                          ║" -ForegroundColor Cyan
        Write-Host "╚═══════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
    }
}

function Write-Step {
    param(
        [string]$Message,
        [string]$Status = "INFO",
        [switch]$NoNewLine
    )
    
    if ($Quiet) { return }
    
    $colors = @{
        "INFO" = "Blue"
        "SUCCESS" = "Green"
        "WARNING" = "Yellow"
        "ERROR" = "Red"
        "WORKING" = "Cyan"
    }
    
    $icons = @{
        "INFO" = "ℹ️"
        "SUCCESS" = "✅"
        "WARNING" = "⚠️"
        "ERROR" = "❌"
        "WORKING" = "⚡"
    }
    
    $color = $colors[$Status]
    $icon = $icons[$Status]
    
    if ($NoNewLine) {
        Write-Host "$icon $Message" -ForegroundColor $color -NoNewline
    } else {
        Write-Host "$icon $Message" -ForegroundColor $color
    }
}

function Show-Progress {
    param(
        [string]$Activity,
        [int]$PercentComplete,
        [string]$Status = "Processing..."
    )
    
    if (-not $Quiet) {
        Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete
    }
}

function Write-Spinner {
    param([string]$Message)
    
    if ($Quiet) { return }
    
    $spinChars = @('|', '/', '-', '\')
    for ($i = 0; $i -lt 20; $i++) {
        $char = $spinChars[$i % 4]
        Write-Host "`r$char $Message" -NoNewline -ForegroundColor Cyan
        Start-Sleep -Milliseconds 100
    }
    Write-Host "`r✅ $Message - Complete" -ForegroundColor Green
}

# ==============================================================================
#  SYSTEM DETECTION FUNCTIONS
# ==============================================================================

function Test-Prerequisites {
    Write-Step "Checking system prerequisites..." -Status "WORKING"
    
    $issues = @()
    
    # Check PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        $issues += "PowerShell 5.0 or higher required (current: $($PSVersionTable.PSVersion))"
    }
    
    # Check Windows version
    $osVersion = [System.Environment]::OSVersion.Version
    if ($osVersion.Major -lt 10 -and ($osVersion.Major -eq 6 -and $osVersion.Minor -lt 1)) {
        $issues += "Windows 7 SP1 or higher required"
    }
    
    # Check internet connectivity
    try {
        $null = Invoke-WebRequest -Uri "https://www.microsoft.com" -UseBasicParsing -TimeoutSec 10
    } catch {
        $issues += "Internet connection required for installation"
    }
    
    if ($issues.Count -eq 0) {
        Write-Step "All prerequisites met" -Status "SUCCESS"
        return $true
    } else {
        Write-Step "Prerequisites check failed:" -Status "ERROR"
        foreach ($issue in $issues) {
            Write-Host "   • $issue" -ForegroundColor Red
        }
        return $false
    }
}

function Get-OfficeInstallation {
    Write-Step "Detecting Office installation..." -Status "WORKING"
    
    $installations = @()
    
    # Check Click-to-Run installations
    try {
        $ctrPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        if (Test-Path $ctrPath) {
            $ctrInfo = Get-ItemProperty -Path $ctrPath -ErrorAction SilentlyContinue
            if ($ctrInfo.InstallationPath) {
                $installations += @{
                    Type = "Click-to-Run"
                    Path = $ctrInfo.InstallationPath
                    Version = $ctrInfo.VersionToReport
                    DisplayVersion = $ctrInfo.DisplayVersion
                    Channel = $ctrInfo.CDNBaseUrl -replace ".*/(.*)/office/.*", '$1'
                }
            }
        }
    } catch {
        Write-Step "Could not detect Click-to-Run installation" -Status "WARNING"
    }
    
    # Check MSI installations
    try {
        $versions = @("16.0", "15.0", "14.0")
        foreach ($version in $versions) {
            $msiPath = "HKLM:\SOFTWARE\Microsoft\Office\$version\Common\InstallRoot"
            if (Test-Path $msiPath) {
                $msiInfo = Get-ItemProperty -Path $msiPath -ErrorAction SilentlyContinue
                if ($msiInfo.Path) {
                    $installations += @{
                        Type = "MSI"
                        Path = $msiInfo.Path
                        Version = $version
                        DisplayVersion = $version
                        Channel = "MSI"
                    }
                }
            }
        }
    } catch {
        Write-Step "Could not detect MSI installation" -Status "WARNING"
    }
    
    if ($installations.Count -eq 0) {
        Write-Step "No Office installation found" -Status "ERROR"
        return $null
    }
    
    # Return the most recent installation
    $latest = $installations | Sort-Object Version -Descending | Select-Object -First 1
    Write-Step "Found Office $($latest.DisplayVersion) ($($latest.Type))" -Status "SUCCESS"
    return $latest
}

function Test-ExcelProcess {
    $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    return $excelProcesses.Count -gt 0
}

# ==============================================================================
#  MANIFEST MANAGEMENT FUNCTIONS
# ==============================================================================

function Get-RemoteManifest {
    param([string]$Url)
    
    Write-Step "Downloading add-in manifest..." -Status "WORKING"
    Show-Progress -Activity "Downloading Manifest" -PercentComplete 25 -Status "Connecting to server..."
    
    try {
        $tempPath = [System.IO.Path]::GetTempPath()
        $manifestFile = Join-Path $tempPath "ipg-taxonomy-extractor-manifest.xml"
        
        Show-Progress -Activity "Downloading Manifest" -PercentComplete 50 -Status "Downloading file..."
        
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "IPG-TaxonomyExtractor-Installer/$($Script:Config.Version)")
        $webClient.DownloadFile($Url, $manifestFile)
        
        Show-Progress -Activity "Downloading Manifest" -PercentComplete 100 -Status "Download complete"
        
        if (Test-Path $manifestFile) {
            Write-Step "Manifest downloaded successfully" -Status "SUCCESS"
            return $manifestFile
        } else {
            throw "File not found after download"
        }
    }
    catch {
        Write-Step "Failed to download manifest: $($_.Exception.Message)" -Status "ERROR"
        return $null
    }
    finally {
        Write-Progress -Activity "Downloading Manifest" -Completed
    }
}

function Test-ManifestFile {
    param([string]$FilePath)
    
    Write-Step "Validating manifest file..." -Status "WORKING"
    
    try {
        if (-not (Test-Path $FilePath)) {
            throw "Manifest file not found: $FilePath"
        }
        
        [xml]$manifest = Get-Content $FilePath -ErrorAction Stop
        
        # Validate required elements
        if (-not $manifest.OfficeApp) {
            throw "Invalid manifest: Missing OfficeApp element"
        }
        
        $addInInfo = @{
            Id = $manifest.OfficeApp.Id
            Version = $manifest.OfficeApp.Version
            DisplayName = $manifest.OfficeApp.DisplayName.DefaultValue
            Description = $manifest.OfficeApp.Description.DefaultValue
            Provider = $manifest.OfficeApp.ProviderName
        }
        
        # Validate Office App ID matches expected
        if ($addInInfo.Id -ne $Script:Config.AddInId) {
            throw "Manifest ID mismatch. Expected: $($Script:Config.AddInId), Found: $($addInInfo.Id)"
        }
        
        Write-Step "Manifest validation successful" -Status "SUCCESS"
        
        if (-not $Quiet) {
            Write-Host ""
            Write-Host "📋 Add-in Information:" -ForegroundColor Blue
            Write-Host "   Name: $($addInInfo.DisplayName)" -ForegroundColor White
            Write-Host "   Version: $($addInInfo.Version)" -ForegroundColor White
            Write-Host "   Provider: $($addInInfo.Provider)" -ForegroundColor White
            Write-Host "   ID: $($addInInfo.Id)" -ForegroundColor Gray
            Write-Host ""
        }
        
        return $addInInfo
    }
    catch {
        Write-Step "Manifest validation failed: $($_.Exception.Message)" -Status "ERROR"
        return $null
    }
}

# ==============================================================================
#  REGISTRY MANAGEMENT FUNCTIONS
# ==============================================================================

function Add-TrustedCatalog {
    param(
        [string]$ManifestUrl,
        [hashtable]$AddInInfo
    )
    
    Write-Step "Adding add-in to trusted catalogs..." -Status "WORKING"
    
    try {
        $registryBase = "HKCU:\Software\Microsoft\Office\16.0\WEF"
        $catalogsPath = "$registryBase\TrustedCatalogs"
        
        # Ensure registry path exists
        if (-not (Test-Path $registryBase)) {
            New-Item -Path $registryBase -Force | Out-Null
        }
        
        if (-not (Test-Path $catalogsPath)) {
            New-Item -Path $catalogsPath -Force | Out-Null
        }
        
        # Create unique catalog entry
        $catalogId = "{$([System.Guid]::NewGuid().ToString())}"
        $catalogPath = "$catalogsPath\$catalogId"
        
        New-Item -Path $catalogPath -Force | Out-Null
        Set-ItemProperty -Path $catalogPath -Name "Id" -Value $ManifestUrl
        Set-ItemProperty -Path $catalogPath -Name "Url" -Value $ManifestUrl
        Set-ItemProperty -Path $catalogPath -Name "Flags" -Value 1
        
        # Store installation metadata
        $metaPath = "$registryBase\IPGTaxonomyExtractor"
        if (-not (Test-Path $metaPath)) {
            New-Item -Path $metaPath -Force | Out-Null
        }
        
        Set-ItemProperty -Path $metaPath -Name "CatalogId" -Value $catalogId
        Set-ItemProperty -Path $metaPath -Name "InstallDate" -Value (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        Set-ItemProperty -Path $metaPath -Name "Version" -Value $AddInInfo.Version
        Set-ItemProperty -Path $metaPath -Name "ManifestUrl" -Value $ManifestUrl
        
        Write-Step "Trusted catalog entry created successfully" -Status "SUCCESS"
        return $catalogId
    }
    catch {
        Write-Step "Failed to create trusted catalog: $($_.Exception.Message)" -Status "ERROR"
        return $null
    }
}

function Remove-TrustedCatalogs {
    Write-Step "Removing existing add-in installations..." -Status "WORKING"
    
    try {
        $registryBase = "HKCU:\Software\Microsoft\Office\16.0\WEF"
        $catalogsPath = "$registryBase\TrustedCatalogs"
        $removed = 0
        
        if (Test-Path $catalogsPath) {
            $catalogs = Get-ChildItem -Path $catalogsPath -ErrorAction SilentlyContinue
            
            foreach ($catalog in $catalogs) {
                try {
                    $url = Get-ItemProperty -Path $catalog.PSPath -Name "Url" -ErrorAction SilentlyContinue
                    if ($url -and $url.Url -like "*ipg-taxonomy-extractor*") {
                        Remove-Item -Path $catalog.PSPath -Recurse -Force
                        $removed++
                    }
                } catch {
                    # Continue on individual catalog errors
                }
            }
        }
        
        # Remove installation metadata
        $metaPath = "$registryBase\IPGTaxonomyExtractor"
        if (Test-Path $metaPath) {
            Remove-Item -Path $metaPath -Recurse -Force
            $removed++
        }
        
        if ($removed -gt 0) {
            Write-Step "Removed $removed registry entries" -Status "SUCCESS"
        } else {
            Write-Step "No existing installations found" -Status "INFO"
        }
        
        return $true
    }
    catch {
        Write-Step "Error during cleanup: $($_.Exception.Message)" -Status "WARNING"
        return $false
    }
}

# ==============================================================================
#  INSTALLATION FUNCTIONS
# ==============================================================================

function Install-AddIn {
    Write-Header
    
    # Step 1: Prerequisites
    if (-not (Test-Prerequisites)) {
        Write-Step "Installation aborted due to prerequisite failures" -Status "ERROR"
        return $false
    }
    
    # Step 2: Office Detection
    $officeInfo = Get-OfficeInstallation
    if (-not $officeInfo) {
        Write-Step "Microsoft Excel not found. Please install Microsoft Office." -Status "ERROR"
        return $false
    }
    
    # Step 3: Process Check
    if (Test-ExcelProcess) {
        Write-Step "Excel is currently running" -Status "WARNING"
        Write-Host ""
        Write-Host "⚠️  Please close all Excel windows and try again." -ForegroundColor Yellow
        Write-Host "   This ensures the add-in registers properly with Excel." -ForegroundColor Yellow
        Write-Host ""
        return $false
    }
    
    # Step 4: Download Manifest
    $manifestPath = Get-RemoteManifest -Url $ManifestUrl
    if (-not $manifestPath) {
        return $false
    }
    
    # Step 5: Validate Manifest
    $addInInfo = Test-ManifestFile -FilePath $manifestPath
    if (-not $addInInfo) {
        return $false
    }
    
    # Step 6: Clean Previous Installations
    Remove-TrustedCatalogs | Out-Null
    
    # Step 7: Install Add-in
    $catalogId = Add-TrustedCatalog -ManifestUrl $ManifestUrl -AddInInfo $addInInfo
    if (-not $catalogId) {
        return $false
    }
    
    # Installation Complete
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║                            INSTALLATION SUCCESSFUL!                          ║" -ForegroundColor Green
    Write-Host "╚═══════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "🎉 IPG Taxonomy Extractor has been installed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "📋 Next Steps:" -ForegroundColor Blue
    Write-Host "   1. Open Microsoft Excel" -ForegroundColor White
    Write-Host "   2. Go to Home → Add-ins → More Add-ins" -ForegroundColor White
    Write-Host "   3. Look for 'IPG Taxonomy Extractor' in the available add-ins" -ForegroundColor White
    Write-Host "   4. Click 'Add' to activate the add-in" -ForegroundColor White
    Write-Host ""
    Write-Host "💡 Troubleshooting:" -ForegroundColor Yellow
    Write-Host "   • If the add-in doesn't appear, restart Excel and try again" -ForegroundColor White
    Write-Host "   • Check if your organization allows custom add-ins" -ForegroundColor White
    Write-Host "   • Contact IT support if you encounter permission issues" -ForegroundColor White
    Write-Host ""
    Write-Host "🔗 Support: $($Script:Config.SupportUrl)" -ForegroundColor Cyan
    Write-Host ""
    
    return $true
}

function Uninstall-AddIn {
    Write-Header
    
    Write-Step "Starting uninstallation process..." -Status "WORKING"
    
    if (Test-ExcelProcess) {
        Write-Step "Excel is currently running" -Status "WARNING"
        Write-Host ""
        Write-Host "⚠️  Please close all Excel windows and try again." -ForegroundColor Yellow
        Write-Host "   This ensures complete removal of the add-in." -ForegroundColor Yellow
        Write-Host ""
        return $false
    }
    
    $success = Remove-TrustedCatalogs
    
    if ($success) {
        Write-Host ""
        Write-Host "╔═══════════════════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║                           UNINSTALLATION COMPLETE!                           ║" -ForegroundColor Green
        Write-Host "╚═══════════════════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "✅ IPG Taxonomy Extractor has been removed successfully!" -ForegroundColor Green
        Write-Host ""
        Write-Host "ℹ️  You may need to restart Excel to see the changes." -ForegroundColor Blue
        Write-Host ""
    } else {
        Write-Step "Uninstallation completed with warnings" -Status "WARNING"
    }
    
    return $success
}

# ==============================================================================
#  MAIN EXECUTION
# ==============================================================================

try {
    if ($Uninstall) {
        $result = Uninstall-AddIn
    } else {
        $result = Install-AddIn
    }
    
    if ($result) {
        exit 0
    } else {
        exit 1
    }
}
catch {
    Write-Step "Unexpected error: $($_.Exception.Message)" -Status "ERROR"
    Write-Host ""
    Write-Host "🐛 If this error persists, please report it at:" -ForegroundColor Yellow
    Write-Host "   $($Script:Config.SupportUrl)/issues" -ForegroundColor Yellow
    Write-Host ""
    exit 1
}
finally {
    if (-not $Quiet) {
        Write-Host "Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}