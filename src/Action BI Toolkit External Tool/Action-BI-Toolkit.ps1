#region config

# Run this once to prevent the script from prompting warning each time
#Get-ChildItem C:\Users\BrianMather\AppData\Local\ActionBIToolkit\Action-BI-Toolkit.ps1 | Unblock-File

$global:ActionBIToolkitDependenciesPath = Join-Path ${env:LOCALAPPDATA} "ActionBIToolkit"

# Override the version of Power BI Desktop that pbi-tools uses
# $env:PBITOOLS_PbiInstallDir = "C:\Program Files\Microsoft Power BI Desktop\"
$env:PBITOOLS_PbiInstallDir = "C:\Users\BrianMather\Downloads\2.94.921.0\2.94.921.0\Microsoft Power BI Desktop\"

# Set the path to the pbi-tools executable within the Action BI Toolkit directory
$PbiToolsExePath = Join-Path $ActionBIToolkitDependenciesPath "pbi-tools\pbi-tools.exe"

# Test is pbi-tools.exe installed
[bool] $global:UsePBITools = Test-Path $PbiToolsExePath

if ($UsePBITools) {
    # Define new alias for pbi-tools.exe to ensure the External Tool uses it's version, not a version from the $env:Path
    New-Alias pbitools $PbiToolsExePath -Force
}

#endregion config

#region toolkit_startup

#region fetch_current_PBIXfileDetails
function Get-PowerBIDesktopSessions {
    param
    (
        [Parameter(Mandatory = $false, Position = 0)] [int] $LocalPbixNextWindowProcessId 
    )

    $id = Get-CimInstance -Class Win32_Process -Filter "Name LIKE 'msmdsrv.exe'" ##| Select-Object -ExpandProperty ProcessId
    
    if (0 -ne $LocalPbixNextWindowProcessId) {
        $id = $id | Where-Object { $_.ParentProcessId -eq $LocalPbixNextWindowProcessId }
    }
    else {}

    $localsessions = @()

    foreach ($i in $id) {
        # $AnalysisServicesWorkspacePath = (($i | Select-Object -ExpandProperty Commandline) -split " -s " -split "-c -n ")[2].Replace("`"","")
        # $PortNumberInMSMDSRVPORTTXT = Get-Content (Join-Path $AnalysisServicesWorkspacePath msmdsrv.port.txt) -encoding unicode
        $LocalPort = (Get-NetTCPConnection -OwningProcess $i.ProcessId)[0].LocalPort
        
        $PowerBIDesktopProcessId = $i.ParentProcessId
        $PowerBIDesktopProcess = (Get-CimInstance -Class Win32_Process -Filter "ProcessId = $($PowerBIDesktopProcessId)" )
        
        [string]$PBIXFilePath = (($PowerBIDesktopProcess | Select-Object -ExpandProperty Commandline) -split '\"')[3]
        [string]$LocalPBIXName = ((Get-Process -id $PowerBIDesktopProcessId).MainWindowTitle -split " - Power BI Desktop")[0]
        
        # Fall back to split path of commandline if MainWindowTitle does not return name
        if ($LocalPBIXName -eq "") {
            $LocalPBIXName = (Split-Path -Path $PBIXFilePath -Leaf).Replace(".pbix", "")
        } 
        else {}
        
        $as = New-Object Microsoft.AnalysisServices.Tabular.Server
        try {
            $as.Connect("localhost:$LocalPort")
            $LocalDatabase = ($as.Databases)[0].ID
            $as.Disconnect()
            if ($LocalDatabase.Count -eq 0) {
                $LocalDatsetType = "Thin report: connected to Local PBIX"
            }
            else {
                $LocalDatsetType = "Thick Report: Model embedded within Local PBIX"

            }
        }
        catch {
            $LocalDatsetType = "Thin report: connected to Local PBIX"
        }

        $localsessions += [pscustomobject]@{
            Port                 = $LocalPort;
            Server               = "localhost:$LocalPort";
            Database             = $LocalDatabase;
            PBIXName             = $LocalPBIXName;
            DatasetType          = $LocalDatsetType;
            PbixPath             = $PBIXFilePath;
            PbixDesktopProcessId = $PowerBIDesktopProcessId
        }
    }

    return , $localsessions
}
function Get-UserSelectLocalPBISession {
    
    $LocalSessions = Get-PowerBIDesktopSessions
    
    switch ($LocalSessions.count) {
        0 { Exit-ActionBIToolkit "No PBI Desktop sessions found...  " }
        1 { $ThisSession = $LocalSessions[0] }
        default {
            Write-Host "Please select one of the running pbix instances:"
            $i = 1
            foreach ($session in $LocalSessions ) {
                if ("" -eq $session.PbixPath) {
                    Write-Host "$i) $($session.PBIXName)"    
                }
                else {
                    Write-Host "$i) $($session.PbixPath)"
                }
                $i++
            }
            $selection = Read-Host -Prompt "`nEnter the number of the file to select"
            $ThisSession = $LocalSessions[$selection - 1]
        }
    }
    return $ThisSession
}
function Get-ConnectionDetails {
    param
    (
        [Parameter(Mandatory = $false, Position = 1)] [string] $LocalExternalToolServer,
        [Parameter(Mandatory = $false, Position = 2)] [string] $LocalExternalToolDatabase
    )

    # Define outputs
    $LocalServerOut = ""
    $LocalDatabaseOut = ""
    $LocalPBIXFileType = ""
    
    $LocalServerOut = $LocalExternalToolServer
    $LocalDatabaseOut = $LocalExternalToolDatabase
    $PbixProcessId = Get-PbixProcessIdFromNextWindow
    $LocalSelectedSession = Get-PowerBIDesktopSessions $PbixProcessId
    
    if ($null -eq $LocalSelectedSession) {
        # Catch where Power BI Desktop ends up in a temporary save TempSaves state.
        Write-Host "Power BI Desktop session is in a temporary save state" -ForegroundColor Red
        Write-Host "You will need to chose >File>Save in Power BI Desktop and try the external tool again" -ForegroundColor Red
        Exit-ActionBIToolkit "`n`nExiting";
    }
    else {
        
        $OutFilePath = $LocalSelectedSession.PbixPath
        $OutPBIXName = $LocalSelectedSession.PBIXName

        if ($LocalExternalToolServer -eq "pbiazure://api.powerbi.com") {
            $LocalPBIXFileType = "Thin Report: Power BI Service"
        }

        if ($LocalExternalToolServer -match "asazure") {
            $LocalPBIXFileType = "Thin Report: Azure Analysis Services"
        }

        # Differentiate between golden dataset (hub) and spoke (thin report) 
        if (($LocalExternalToolServer -match "local") -or ($LocalServerOut -match "local") ) {
            if ("localhost:$($LocalSelectedSession.Port)" -eq $LocalExternalToolServer) {
                $LocalPBIXFileType = "Thick Report: Model embedded within Local PBIX"
            }
            else {
                $LocalPBIXFileType = "Thin report: connected to Local PBIX"
            }            
        }
    }
    
    return $LocalServerOut, $LocalDatabaseOut, $LocalPBIXFileType, $OutFilePath, $OutPBIXName

} #close Get-ConnectionDetails
function Get-PbixProcessIdFromNextWindow {
    
    #########################################################################
    
    # Find the location and name of the .pbix file being locked by the window
    # that the External Tool was called from
    # and return the name of the file, the folder and the name of the model 
    #https://docs.microsoft.com/en-us/sysinternals/downloads/handle
    
    $signature = @"
[DllImport("user32.dll")]
public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")]
public static extern IntPtr GetWindow(IntPtr hWnd, Int32 wCmd);
[DllImport("user32.dll")]
public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
"@

    # Add Win32 functions
    #--if (("Win32Functions.Win32GetForegroundWindow").IsPublic -as [type] -or ($null -eq $FindForegroundWindow)) {} else {
    $FindForegroundWindow = Add-Type -memberDefinition $signature -name "Win32GetForegroundWindow" -namespace Win32Functions -passThru
    #--}
    #--if (("Win32Functions.Win32GetWindow" -as [type]).IsPublic -or ($null -eq $FindWindow)) {} else {
    $FindWindow = Add-Type -memberDefinition $signature -name "Win32GetWindow" -namespace Win32Functions -passThru
    #}

    #if (("Win32Functions.Win32GetWindowThreadProcessId" -as [type]).IsPublic -or ($null -eq $FindWindowThreadProcessId)) {} else {
    $FindWindowThreadProcessId = Add-Type -memberDefinition $signature -name "Win32GetWindowThreadProcessId" -namespace Win32Functions -passThru
    #}
    # Get window handle for PowerShell Window
    $hPowerShell = $FindForegroundWindow::GetForegroundWindow()

    # Get window handle for the Window below PowerShell which will be the window handle for the PBI Desktop instance the External Tool has been called from
    $hPBI = $FindWindow::GetWindow($hPowerShell, 2)

    # Initialise $pbiProcessID 
    $pbiProcessId = 0

    # Find the process ID for the Power BI Desktop instance that owns that window handle
    $FindWindowThreadProcessId::GetWindowThreadProcessId($hPBI, [ref]$pbiProcessId) | Out-Null
    
    if ((Get-Process -Id $pbiProcessId).ProcessName -ne "PBIDesktop") {
        Exit-ActionBIToolkit "Next window not a Power BI Desktop session. Please try again"
    }
    
    return $pbiProcessId
}
#endregion fetch_current_PBIXfileDetails
function Initialize-Toolkit {
    param
    (
        [Parameter(Mandatory = $false, Position = 0)] [string]$LocalArgs0,
        [Parameter(Mandatory = $false, Position = 1)] [string]$LocalArgs1
    )

    Get-TabularPackages

    # Get connection and pbix file details in VS Code mode
    if ("" -eq $LocalArgs0) { 
        $SelectedSession = Get-UserSelectLocalPBISession
        $LocalServer, $LocalDatabase, $LocalPBIXFileType, $LocalPbixFilePath, $LocalPbixFileName = $SelectedSession.Server, $SelectedSession.Database, $SelectedSession.DatasetType, $SelectedSession.PbixPath, $SelectedSession.PBIXName
    }

    # Get connection and pbix file details in External Tool mode
    else { 
        $LocalServer, $LocalDatabase, $LocalPBIXFileType, $LocalPbixFilePath, $LocalPbixFileName = Get-ConnectionDetails $LocalArgs0 $LocalArgs1 
    }

    # Exit if pbix file not found
    if (!(Test-Path $LocalPbixFilePath)) {
        Exit-ActionBIToolkit "The pbix file for $($LocalPbixFileName) could not be found. You may need to save, close and reopen your pbix file and try again"
    }
    
    $LocalPbixRootFolder = Split-Path -Path $LocalPbixFilePath
    
    Write-Host $("`nAction BI Toolkit:") -ForegroundColor Yellow
    Write-Host "`Powershell Version:" $PSVersionTable.PSVersion

    Write-Host $("`nModel name:  $LocalPbixFileName")
    Write-Host $("File Path:   $LocalPbixFilePath")
    Write-Host $("File type:   $LocalPBIXFileType") -NoNewline

    Write-Host $("Server:      $LocalServer") 
    Write-Host $("Database:    $LocalDatabase") 

    # Define Output folders
    $thinReportsSubFolder = "_Thin Reports"
    $toolkitSubFolder = "__Toolkit"
    $deploymentSubFolder = "__Deployment"
    $dependenciesSubFolder = "_Dependencies"
    $performanceDataSubFolder = "_Page Tests"
    $exportDaxQueriesSubFolder = "_ExportQueries"
    $pbixQuickBackupSubFolder = "_QuickBackups"
    $pbixBIMSubFolder = "_Bim"

    $toolkitFolder = Join-Path $LocalPbixRootFolder $toolkitSubFolder
    $deploymentFolder = Join-Path $LocalPbixRootFolder $deploymentSubFolder
    $LocalthinReportsFolder = Join-Path $LocalPbixRootFolder $thinReportsSubFolder
    
    $LocalToolkitSettings = New-Object -TypeName psobject
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name toolkitFolder -Value $toolkitFolder
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name deploymentFolder -Value $deploymentFolder
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name dependenciesOutFolder -Value ( Join-Path $toolkitFolder $dependenciesSubFolder)
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name performanceDataFolder -Value (Join-Path $toolkitFolder $performanceDataSubFolder)
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name exportDaxQueriesFolder -Value (Join-Path $toolkitFolder $exportDaxQueriesSubFolder)
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name PbixQuickBackupFolder -Value (Join-Path $toolkitFolder $pbixQuickBackupSubFolder)
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name PbixBIMDeployFolder -Value (Join-Path $toolkitFolder $pbixBIMSubFolder)
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name thinReportsFolder -Value $LocalthinReportsFolder
    
    # Create toolkit folders if they don't exist
    $LocalToolkitSettings | Get-Member -MemberType NoteProperty | foreach-object {
        $directory = $LocalToolkitSettings."$($_.Name)"
        $directory
        New-Item -ItemType Directory -Force -Path $directory | Out-Null
    }

    if ($LocalPBIXFileType -ne "Thick Report: Model embedded within Local PBIX") {
        $LocalPBIXExportFolder = Join-Path $LocalthinReportsFolder $LocalPbixFileName
    }
    else {
        $LocalPBIXExportFolder = Join-Path $LocalPbixRootFolder $LocalPbixFileName
    }

    # Prepare the remaining settings
    $LocalPbixProjSettings = Get-PbixProjSettings $LocalPBIXExportFolder

    switch ( $LocalPBIXFileType ) {
        "Thin Report: Power BI Service" { 
            Write-Host " (Thin report - connected to dataset in Power BI Service)"; 
            $LocalConString = ($LocalPbixProjSettings.workspaceConnectionString) 
        }
        "Thin Report: Azure Analysis Services" { 
            Write-Host " (Thin report - connected to Azure Analyis Services instance)"; 
            $LocalConString = $LocalPbixProjSettings.workspaceConnectionString 
        }
        "Thick Report: Model embedded within Local PBIX" { 
            Write-Host " (Report with embedded model)"; 
            $LocalConString = "Datasource=$($LocalServer); Initial Catalog=$($LocalDatabase);timeout=0; connect timeout =0" 
        }
        "Thin report: connected to Local PBIX" { 
            Write-Host " (Thin report - connected to local pbix instance)"; 
            $LocalConString = "Datasource=$($LocalServer); Initial Catalog=$($LocalDatabase);timeout=0; connect timeout =0" 
        }
        default { $LocalConString = $null; } 
    }
    
    # Add the pbix file details to the settings object
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name PbixRootFolder -Value $LocalPbixRootFolder
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name PbixFileName -Value $LocalPbixFileName
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name PbixFilePath -Value $LocalPbixFilePath
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name PbixFileType -Value $LocalPBIXFileType
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name PbixExportFolder -Value $LocalPBIXExportFolder
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name Server -Value $LocalServer
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name Database -Value $LocalDatabase
    $LocalToolkitSettings | Add-Member -MemberType NoteProperty -Name ConString -Value $LocalConString
    
    Get-GitIgnoreTemplate $LocalToolkitSettings

    return $LocalToolkitSettings
}
function Get-RunOption {
    param ( 
        [string] $LocalPbixFileType 
    )

    if ($LocalPbixFileType -eq "Thick Report: Model embedded within Local PBIX") {
        Write-Host "`nExecution options: " -Foregroundcolor Yellow
        Write-Host " 1) Export PBIX for source control"
        Write-Host " 2) Run page tests" -Foregroundcolor DarkGray
        Write-Host " 3) Save backup of PBIX" -Foregroundcolor DarkGray
        Write-Host " 4) Execute DAX Queries and export to csv" -Foregroundcolor DarkGray
        Write-Host " 5) Open in VS Code" -Foregroundcolor DarkGray
        Write-Host " 6) Compile Thick Report .pbit from Source Control"
        Write-Host " 7) Compile Thin Report(s) .pbix from Source Control"
        Write-Host " 8) Deploy to Power BI environment"
        Write-Host "         (NB: default = 1) " -Foregroundcolor DarkGray
        
        Write-Host "`nChoose option: " -ForegroundColor Yellow -NoNewline
        
        $option = Read-Host
        if ($option -eq "") { $Chosenoption = "Export PBIX for Source Control"; $option = 1 }
        else {
            switch ($option) {
                1 { $Chosenoption = "Export PBIX for source control" }
                2 { $Chosenoption = "Report Page Tests" }
                3 { $Chosenoption = "Backup PBIX" }
                4 { $Chosenoption = "Export DAX Queries" }
                5 { $Chosenoption = "Open VSCode" }
                6 { $Chosenoption = "Compile Thick Report .pbit from Source Control" }
                7 { $Chosenoption = "Compile Thin Report(s) .pbix from Source Control" }
                8 { $Chosenoption = "Deploy to Power BI environment" }
                9 { $Chosenoption = "All" }
            }
        }
        Clear-Host
        Write-host "`nExecution option chosen: " -NoNewline
        Write-Host "$option) $Chosenoption" -ForegroundColor Yellow
    }
    else {
        Write-host "`nDefault execution option: " -NoNewline
        Write-host "Export PBIX for source control" -ForegroundColor Yellow
        $Chosenoption = "Export PBIX for source control"
    }
    
    return $Chosenoption
}
function Get-DeploymentFolders {
    param
    (
        [Parameter(Mandatory = $false, Position = 0)] [psobject] $ls
    )

    $LocalDeploymentFolderName = $ls.deploymentFolder
    $LocalFolders = Get-ChildItem $ls.deploymentFolder -Recurse | Where-Object { $_.Mode -match "d" }
    $LocalDeploymentWorkspaceFolder = $LocalFolders | Where-Object {
        ($_.Name -notmatch "Local") -and 
        ($_.Parent.Fullname -eq $LocalDeploymentFolderName)
    }
    $LocalDeploymentEnvironmentFolders = $LocalFolders | Where-Object {
        ($_.Name -notmatch "Local") -and 
        ($_.Parent.Fullname -ne $LocalDeploymentFolderName)
    }
    return $LocalDeploymentWorkspaceFolder, $LocalDeploymentEnvironmentFolders
}
function Get-GitIgnoreTemplate {

param
    (
        [Parameter(Mandatory = $true, Position = 0)] [psobject] $ls
    )

    $GitIgnorePath = Join-Path $ls.PbixRootFolder ".gitignore"
    $GitIgnoreTemplatePath = Join-Path $ActionBIToolkitDependenciesPath "sample gitignore.txt"
    if (Test-Path $GitIgnorePath) {
        # do nothing if .gitignore already exists
    }
    else {
        Copy-Item -Path $GitIgnoreTemplatePath -Destination $GitIgnorePath 
    }
}
#endregion toolkit_startup

#region general_utilities
#requires -Version 2
function Show-Process($Process, [Switch]$Maximize)
{
  $sig = '
    [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern int SetForegroundWindow(IntPtr hwnd);
  '
  
  if ($Maximize) { $Mode = 3 } else { $Mode = 4 }
  $type = Add-Type -MemberDefinition $sig -Name WindowAPI -PassThru
  $hwnd = $process.MainWindowHandle
  $null = $type::ShowWindowAsync($hwnd, $Mode)
  $null = $type::SetForegroundWindow($hwnd) 
}
function Exit-ActionBIToolkit ($LocalMessage, $LocalPbixProjFolder) {
    #param {[Parameter(Mandatory = $true, Position = 0)] $LocalMessage,[Parameter(Mandatory = $true, Position = 0)] $LocalPbixProjFolder}
    # Check if running Powershell ISE
    if ($psISE) {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show($this, "$LocalMessage")
    }
    else {
        
        Write-Host "$LocalMessage" -ForegroundColor Yellow

        # Open or switch to project folder in VS Code
        if ($null -ne $LocalPbixProjFolder) {
            Write-Host "`rOpening in VS Code..." -ForegroundColor Yellow
            code $LocalPbixProjFolder
        }

        #$host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
        $counter = 6
        while (!$Host.UI.RawUI.KeyAvailable -and ($counter-- -gt 1)) {
            Write-Host -NoNewline $("`rClosing in $counter... click to pause exit")
            [Threading.Thread]::Sleep( 1000 )
        }
        Write-Host -NoNewline $("`r                                   ")
        Write-Host "`n"
        exit;
    }
} #close Exit-ActionBIToolkit
function Open-FolderInVSCode {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] $LocalPbixProjFolder    
    )
    Write-Host "`n`nOpening folder in VSCode..."
    code $LocalPbixProjFolder
} #close Open-FolderInVSCode
function Get-TabularPackages {

    if ($PSVersionTable.PSEdition -eq "Desktop") {
        $AdmomdClientAssemblyString = "$ActionBIToolkitDependenciesPath\Microsoft.AnalysisServices.AdomdClient.retail.amd64.19.18.0\lib\net45\Microsoft.AnalysisServices.AdomdClient.dll"
    }
    else { 
        $AdmomdClientAssemblyString = "$ActionBIToolkitDependenciesPath\Microsoft.AnalysisServices.AdomdClient.NetCore.retail.amd64.19.18.0\lib\netcoreapp3.0\Microsoft.AnalysisServices.AdomdClient.dll"
    }
    if ( Test-Path $AdmomdClientAssemblyString) { 
        $AdmomdClientAssemblyPath = Resolve-Path $AdmomdClientAssemblyString #| Out-Null 
    }
    else {
        if (-not $(Get-PackageSource -ProviderName NuGet -ErrorAction Ignore)) {
            # add the packagesource
            Find-PackageProvider -Name Nuget | Install-PackageProvider -Scope CurrentUser -Force
            Register-PackageSource -Name nuget.org -Location https://www.nuget.org/api/v2 -ProviderName NuGet
        }
        
        if ($PSVersionTable.PSEdition -eq "Desktop") {
            Write-Host "Installing Microsoft.AnalysisServices.AdomdClient.retail.amd64 package`n$ActionBIToolkitDependenciesPath"
            Install-Package -Name Microsoft.AnalysisServices.AdomdClient.retail.amd64 -ProviderName NuGet -Scope CurrentUser -RequiredVersion 19.18.0 -SkipDependencies -Destination $ActionBIToolkitDependenciesPath -Force;
            $AdmomdClientAssemblyPath = Resolve-Path $AdmomdClientAssemblyString #| Out-Null
        }
        else {
            Write-Host "Installing Microsoft.AnalysisServices.AdomdClient.NetCore.retail.amd64 package`n$ActionBIToolkitDependenciesPath"
            Install-Package -Name Microsoft.AnalysisServices.AdomdClient.NetCore.retail.amd64 -ProviderName NuGet -Scope CurrentUser -RequiredVersion 19.18.0 -SkipDependencies -Destination $ActionBIToolkitDependenciesPath -Force;
            $AdmomdClientAssemblyPath = Resolve-Path $AdmomdClientAssemblyString #| Out-Null
        }
    }
    
    try {
        Add-Type -Path $AdmomdClientAssemblyPath
    }
    catch { 
        $_.Exception.LoaderExceptions 
    }

    if ($PSVersionTable.PSEdition -eq "Desktop") {
        $TabularAssemblyString = "$ActionBIToolkitDependenciesPath\Microsoft.AnalysisServices.retail.amd64.19.18.0\lib\net45\Microsoft.AnalysisServices.Tabular.dll"
    }
    else { 
        $TabularAssemblyString = "$ActionBIToolkitDependenciesPath\Microsoft.AnalysisServices.NetCore.retail.amd64.19.18.0\lib\netcoreapp3.0\Microsoft.AnalysisServices.Tabular.dll"
    }
    if ( Test-Path $TabularAssemblyString) { 
        $TabularAssemblyPath = Resolve-Path $TabularAssemblyString # | Out-Null
    }
    else {
        if ($PSVersionTable.PSEdition -eq "Desktop") {
            Write-Host "Installing Microsoft.AnalysisServices.retail.amd64 package`n $ActionBIToolkitDependenciesPath"
            Install-Package -Name Microsoft.AnalysisServices.retail.amd64 -ProviderName NuGet -Scope CurrentUser -RequiredVersion 19.18.0 -SkipDependencies -Destination $ActionBIToolkitDependenciesPath -Force;
            $TabularAssemblyPath = Resolve-Path $TabularAssemblyString #| Out-Null
        }
        else {
            Write-Host "Installing Microsoft.AnalysisServices.NetCore.retail.amd64 package`n $ActionBIToolkitDependenciesPath"
            Install-Package -Name Microsoft.AnalysisServices.NetCore.retail.amd64 -ProviderName NuGet -Scope CurrentUser -RequiredVersion 19.18.0 -SkipDependencies -Destination $ActionBIToolkitDependenciesPath -Force;
            $TabularAssemblyPath = Resolve-Path $TabularAssemblyString #| Out-Null
        }
    }
    try {
        Add-Type -Path $TabularAssemblyPath
    }
    catch { 
        $_.Exception.LoaderExceptions 
    }
} #close Get-TabularPackages
function Test-FileLock {
    param(
        [parameter(Mandatory = $True)]
        [string]$Path
    )
    $OFile = New-Object System.IO.FileInfo $Path
    if ((Test-Path -Path $Path -PathType Leaf) -eq $False) { return $False }
    else {
        try {
            $OStream = $OFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            if ($OStream) { $OStream.Close() }
            return $False
        } 
        catch { return $True }
    }
}
#endregion general_utilities

#region toolkit_mainfunctions
function Export-PBIX {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] [psobject] $ls,
        [Parameter(Mandatory = $true, Position = 2)] [string] $LocalReportJson
    )

    if ($UsePBITools) {
        if ( $ls.PbixFileType -ne "Thick Report: Model embedded within Local PBIX") {
            $LocalServer = ''
        }
        else {
            $LocalServer = $ls.Server
        }

        Invoke-PbiToolsExtract $ls.PbixFilePath $ls.PbixExportFolder $LocalServer
    }
    else {
        Write-Host "Export pages & visuals using PowerShell (without pbi-tools)..." -ForegroundColor Cyan -NoNewline
        Export-ReportPagesToFolders $LocalReportJson $ls.PbixExportFolder

        # Optional reformat files
        #Re-format xml and json files to be optimised for readability and source control
        ###??? Format-FilesForSourceControl $pbixTempUnzipFolder $PbixProjFolder
    }

} # end of function Export PBIX
function Get-PbixProjSettings {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] $LocalPbixProjFolder
    )

    # Define settings object
    $LocalPbixProjSettings = New-Object -TypeName psobject

    
    $configFilePath = Join-Path $LocalPbixProjFolder ".pbixproj.json"
    if (Test-Path $configFilePath) {
        $configFile = Get-Content (Join-Path $LocalPbixProjFolder ".pbixproj.json")
        $configFile = $configFile -replace '(?m)(?<=^([^"]|"[^"]*")*)//.*' -replace '(?ms)/\*.*?\*/'
        
        try {
            # Read settings json to object
            $LocalPbixProjSettings = $configFile | ConvertFrom-Json
            $LocalPbixProjSettings | Add-Member -MemberType NoteProperty -Name settingsRead -Value 'ValidSettings'
        } 
        catch { 
            $LocalPbixProjSettings | Add-Member -MemberType NoteProperty -Name settingsRead -Value 'InvalidSettings'
            Write-Host ".pbixproj.json  invalid json content" -ForegroundColor Red
            $configFile | Write-Host -ForegroundColor Red  
        }
    }
    else {
        Write-Host ".pbixproj.json missing" -ForegroundColor Red
        $LocalPbixProjSettings | Add-Member -MemberType NoteProperty -Name settingsRead -Value 'MissingSettings'
    }
    
    return $LocalPbixProjSettings
} #close Get-PbixProjSettings
function Invoke-PbiToolsExportBim {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] $LocalPbixProjFolder,
        [Parameter(Mandatory = $false, Position = 1)] [string] $LocalBIMExportSetting
    )

    Write-Host $("Extracting bim for $LocalBIMExportSetting...") -ForegroundColor Cyan -NoNewline

    switch ( $LocalBIMExportSetting ) {
        "Power BI Premium" { $response = pbitools export-bim $LocalPbixProjFolder skipDataSources }
        "Azure Analysis Services" { $response = pbitools export-bim $LocalPbixProjFolder RemovePBIDataSourceVersion }
        "SQL Server Analysis Services" { $response = pbitools export-bim $LocalPbixProjFolder RemovePBIDataSourceVersion }
        default { $response = $null } 
    }
    
    if ($LASTEXITCODE -ne 0 ) {
        Write-Host '  Failed.' -ForegroundColor Red
        $response | Write-Host -ForegroundColor Red
        
        Write-Host "pbi-tools export-bim failed..." -ForegroundColor Red
        Exit-ActionBIToolkit "`nAction BI Toolkit complete. Press any key to close..."
    }
    else {
        if ($null -ne $response -and $response[-1] -ne "A BIM file could not be exported.") { 
            Write-Host "  Done." 
            #$response | Write-Host -ForegroundColor Yellow
        }
        else { 
            Write-Host "   Failed." -ForegroundColor Red
            Write-Host $("BIM not exported due to incorrect .pbixproj.json setting ""$LocalBIMExportSetting""") -ForegroundColor Red 
            $response | Write-Host 
        }
    }
    return
} #close Invoke-PbiToolsExportBim
function Invoke-PbiToolsExtract {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] $LocalPbixFilePath,
        [Parameter(Mandatory = $true, Position = 1)] [string] $LocalExtractFolderPath,
        [Parameter(Mandatory = $false, Position = 2)] [string] $LocalServer
    )
    
    Write-Host "`nSTARTING .pbix EXTRACTION"
    Write-Host "Extracting pbix $($LocalExtractFolderPath)..." -ForegroundColor Cyan -NoNewline

    if ('' -eq $LocalServer) {
        $response = (pbitools extract -pbixPath $LocalPbixFilePath -extractFolder $LocalExtractFolderPath )
    }
    
    else {
        $LocalPort = $LocalServer.Replace('localhost:', '')
        $response = (pbitools extract -pbixPath $LocalPbixFilePath -extractFolder $LocalExtractFolderPath -pbiPort $LocalPort)
    }

    if ($LASTEXITCODE -ne 0 ) {
        $response | Write-Host -ForegroundColor Red
        pause "pbi-tools extract failed..."
        break;
    }
    else {
        Write-Host "  Done.`n`n"
        #$response | Write-Host -ForegroundColor Yellow
    }
    return
} #close Invoke-PbiToolsExtract
#endregion toolkit_mainfunctions

#region export_reportpages_without_pbitools
function Get-ReportJson {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] [psobject] $ls
    )
    
    ##Create .zip copy of PBIX to work with
    $zipcopyFolder = $ls.PbixRootFolder + "____zipcopy"
    $ZipFile = (Copy-PbixFileToZip $ls.PbixFilePath $zipcopyFolder)
    
    ## Export enitre .zip to folder for source control
    $LocalpbixTempUnzipFolder = Join-Path $ls.PbixRootFolder "____pbix file contents"
    Copy-ZipToFolders $ZipFile $LocalpbixTempUnzipFolder
    
    try { $LocalReportJson = Get-Content -Encoding Unicode -Path (Join-Path $LocalpbixTempUnzipFolder "Report\Layout") | ConvertFrom-Json }
    catch { $LocalReportJson = Get-Content -Encoding Utf8 -Path (Join-Path $LocalpbixTempUnzipFolder "Report\Layout") | ConvertFrom-Json }
    
    Remove-Item -LiteralPath $LocalpbixTempUnzipFolder -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    
    return $LocalReportJson
}
function Format-FilesForSourceControl {
    param ( [string] $LocalPbixContentFolder, [String] $LocalMainFolder )
        
    $filestoAddXmlExtension = Get-ChildItem $LocalPbixContentFolder -Recurse | Where-Object { $_.Mode -notmatch "d" -and ($_.Name -eq "LinguisticSchema") }
    foreach ($file in $filestoAddXmlExtension ) {
        $newFileName = ($file.FullName + ".xml")
        Rename-Item -Path $file.FullName -NewName $newFileName    
    }

    # Re-formats xml and json files for readability

    $filestoFormatXml = Get-ChildItem $LocalPbixContentFolder -Recurse | Where-Object { $_.Mode -notmatch "d" -and ($_.Extension -eq ".xml") }
    foreach ($xmlFile in $filestoFormatXml) {
        $xml = [xml](Get-Content -encoding unicode -Raw -Path $xmlFile.FullName)
        Format-XML $xml 2 | Out-File (Join-Path $LocalMainFolder $xmlFile.Name)
    }

    $filestoTxt = Get-ChildItem $LocalPbixContentFolder -Recurse | Where-Object { $_.Mode -notmatch "d" -and !($_.Extension) -and $_.Name -match "Version" }
        
    foreach ($file in $filestoTxt  ) {
        $newFileName = ($file.FullName + ".txt")
        Rename-Item -Path $file.FullName -NewName $newFileName    
    }

    $filestoJson = Get-ChildItem $LocalPbixContentFolder -Recurse | Where-Object { $_.Mode -notmatch "d" -and !($_.Extension) -and $_.Name -notmatch "DataModel" -and $_.Name -notmatch "LinguisticSchema" -and $_.Name -notmatch ".pbiviz" -and $_.Name -notmatch "SecurityBindings" -and $_.Name -notmatch "Version" }

    foreach ($file in $filestoJson  ) {
        $newFileName = ($file.FullName + ".json")
        Rename-Item -Path $file.FullName -NewName $newFileName    
    }
    $txtfiles = Get-ChildItem $LocalPbixContentFolder | Where-Object { $_.Mode -notmatch "d" -and $_.Name -match "txt" }
    foreach ($tfile in $txtfiles) { 
        Get-Content $tfile | Out-File  (Join-Path $LocalMainFolder $tfile.Name)
    }

    $jsonfiles = Get-ChildItem $LocalPbixContentFolder | Where-Object { $_.Mode -notmatch "d" -and $_.Name -match "json" }
    foreach ($jfile in $jsonfiles) { 
        try {
            ( (Get-Content -encoding unicode -Path $jFile.FullName).Replace("`\`"", "`"").Replace("]`"", "]").Replace("}`"", "}").Replace("`"[", "[").Replace("`"{", "{") | ConvertFrom-Json) | ConvertTo-Json -Depth 100 | Out-File  (Join-Path $LocalMainFolder $jFile.Name)
        }
        catch { 
            try {
                (Get-Content -encoding utf8 -Path $jFile.FullName.Replace("`\`"", "`"").Replace("]`"", "]").Replace("}`"", "}").Replace("`"[", "[").Replace("`"{", "{") | ConvertFrom-Json) | ConvertTo-Json -Depth 100 | Out-File (Join-Path $LocalMainFolder $jFile.Name)
            }
            catch {
                Continue
                pause "$jFile.FullName json file did not decode, continuing without formatting"
            } 
        }
    }

    # Process custom visuals
    $zipCustomVisualsFolder = (Join-Path $LocalPbixContentFolder "Report\CustomVisuals")
    $outCustomVisualsFolder = (Join-Path $LocalMainFolder "CustomVisuals")
    if ( Test-Path $zipCustomVisualsFolder ) {
        Remove-Item -LiteralPath $outCustomVisualsFolder -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

        Copy-Item -Path $zipCustomVisualsFolder -Destination $outCustomVisualsFolder -recurse -Force
        $customVisuals = Get-ChildItem $zipCustomVisualsFolder -Recurse | Where-Object { $_.Mode -notmatch "d" -and ($_.Name -eq "package.json") }
        foreach ($customVisual in $customVisuals) {
            $content = Get-Content $customVisual.FullName | ConvertFrom-Json
            $newFilePath = $customVisual.FullName.Replace($zipCustomVisualsFolder, $outCustomVisualsFolder)
            ConvertTo-Json -InputObject $content -Depth 100 | Format-Json | Set-Content $newFilePath
        }
    }
        
    # Process Static Resources     
    $zipStaticResourcesFolder = (Join-Path $LocalPbixContentFolder "Report\StaticResources")
    $outStaticResourcesFolder = (Join-Path $LocalMainFolder "StaticResources")
        
    Remove-Item -LiteralPath $outStaticResourcesFolder -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

    Copy-Item -Path $zipStaticResourcesFolder -Destination $outStaticResourcesFolder -recurse -Force
    $files = Get-ChildItem $zipStaticResourcesFolder -Recurse | Where-Object { $_.Mode -notmatch "d" -and ($_.Extension -eq ".json") }
    foreach ($file in $files) {
        $content = Get-Content $file.FullName -Raw | ConvertFrom-Json
        $newFilePath = $file.FullName.Replace($zipStaticResourcesFolder, $outStaticResourcesFolder)
        ConvertTo-Json -InputObject $content -Depth 100 | Format-Json | Set-Content $newFilePath
    }

}
function Export-ReportPagesToFolders {
    
    param ( [PSCustomObject] $LocalReportJson, [String]$LocalReportOut )

    # Delete existing report folder
    Remove-Item -LiteralPath $LocalReportOut -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    
    foreach ($ReportPage in $LocalReportJson.sections) {
        $PageFolder = "00" + ($ReportPage.ordinal + 1).ToString().PadLeft(1, "0") + "_" + (Encode-PathString $ReportPage.displayName)
        $PageOutPath = (Join-Path $LocalReportOut $PageFolder)
        
        New-Item -ItemType Directory -Force -Path $PageOutPath | Out-Null
        
        foreach ($Visual in $ReportPage.visualContainers) {
            $VisualConfig = $Visual.config | ConvertFrom-Json

            $DefaultVisualName = $VisualConfig.singleVisual.visualType + "_" + $VisualConfig.name
            $vcObjectTitle = $VisualConfig.singleVisual.vcObjects.title.properties.text.expr.Literal.Value
            $VisualTitle = $VisualConfig.singleVisual.title.properties.text.expr.Literal.Value
        
            $VisualName = 
            if ($null -eq $vcObjectTitle) {
                if ($null -eq $VisualTitle) {
                    $DefaultVisualName + "x"
                }
                else {
                    $VisualTitle + "y"
                }
            }
            else {
                $vcObjectTitle.Substring(1, $vcObjectTitle.Length - 2)
            }

        
            # Determine the visual display order from various places the tabOrder can be set

            $LayoutsTabOrder = $VisualConfig.layouts.tabOrder
        
            $LayoutsPositionTabOrder = $VisualConfig.layouts.position.tabOrder
            [string]$VisualTabOrder = 
            if ($null -eq $LayoutsPositionTabOrder) {
                if ($null -eq $LayoutsTabOrder) { 0000 } 
                else { $LayoutsTabOrder }
            } 
            else {
                $LayoutsPositionTabOrder
            }

            $VisualFolder = $VisualTabOrder.ToString().Trim().PadLeft(4, "0") + "_" + (Encode-PathString $VisualName)
        
            # Create folder for each Visual
            $VisualFolderPath = Join-Path $PageOutPath $VisualFolder 
            ####$VisualFileName = ((Encode-PathString $VisualName) + ".json") #.Replace("_","")

            New-Item -ItemType Directory -Force -Path $VisualFolderPath | Out-Null
        
            # Save entire Visual definition to json
            # Remove noisy query property
            $Visual.PSObject.Properties.Remove('query')
            $Visual.PSObject.Properties.Remove('x')
            $Visual.PSObject.Properties.Remove('y')
            $Visual.PSObject.Properties.Remove('z')
            $Visual.PSObject.Properties.Remove('width')
            $Visual.PSObject.Properties.Remove('height')
            $Visual.PSObject.Properties.Remove('id')
            
            # Remove properties duplicated in config 
            
            # $Visual | ConvertTo-JSON -Depth 100 | Out-File (Join-Path $VisualFolderPath $VisualFileName)
            $Visual.config | ConvertFrom-Json | ConvertTo-JSON -Depth 100 | Format-Json | Out-File (Join-Path $VisualFolderPath "visual config.json")
            if ($Visual.dataTransforms.length -eq 0 -or $null -eq $Visual.dataTransforms.length) {} else { 
                $Visual.dataTransforms | ConvertFrom-Json | ConvertTo-JSON -Depth 100 | Select-Object -Property * -ExcludeProperty @('queryMetadata') | Format-Json | Out-File (Join-Path $VisualFolderPath "visual dataTransforms.json") 
            }
            if ($Visual.filters -eq "[]" -or $null -eq $Visual.filters) {} else { $Visual.filters | ConvertFrom-Json | ConvertTo-JSON -Depth 100 | Format-Json | Out-File (Join-Path $VisualFolderPath "visual filters.json") }
            if ($Visual.query -eq "[]" -or $null -eq $Visual.query) {} else { $Visual.query | ConvertFrom-Json | ConvertTo-JSON -Depth 100 | Format-Json | Out-File (Join-Path $VisualFolderPath "visual query.json") }
            $Visual | Select-Object -Property * -ExcludeProperty @('config', 'filters', 'dataTransforms', 'query')  | ConvertTo-JSON -Depth 100 | Format-Json |  Out-File (Join-Path $VisualFolderPath "visual.json")
        }

        # Save report page json after removing the visual containers
        $page = $ReportPage | Select-Object -Property * -ExcludeProperty visualContainers 
        if ($page.config -eq "[]" -or $null -eq $page.config) {} else { $page.config | ConvertFrom-Json | ConvertTo-JSON -Depth 100 | Format-Json | Out-File (Join-Path $PageOutPath "page config.json") }
        if ($page.filters -eq "[]" -or $null -eq $page.filters) {} else { $page.filters | ConvertFrom-Json | ConvertTo-JSON -Depth 100 | Format-Json | Out-File (Join-Path $PageOutPath "page filter.json") }
        $page | Select-Object -Property * -ExcludeProperty @('config', 'filters') | ConvertTo-JSON -Depth 100 | Format-Json |  Out-File (Join-Path $PageOutPath "page.json") 
    }

    $report = $LocalReportJson | Select-Object -Property * -ExcludeProperty sections
    if ($report.config -eq "[]" -or $null -eq $report.config) {} else { $report.config | ConvertFrom-Json | ConvertTo-JSON -Depth 100 | Format-Json | Out-File (Join-Path $LocalReportOut "config.json") }
    if ($report.filters -eq "[]" -or $null -eq $report.filters) {} else { $report.filters | ConvertFrom-Json | ConvertTo-JSON -Depth 100 | Format-Json | Out-File (Join-Path $LocalReportOut "report filter.json") }
    $LocalReportJson | Select-Object -Property * -ExcludeProperty @('config', 'filters', 'sections') | ConvertTo-JSON -Depth 100 | Format-Json | Out-File (Join-Path $LocalReportOut "report.json")
    Write-Host "  Done."
}
function Copy-PbixFileToZip {

    param ( [String] $FullFilePath, [String] $LocalZipFolder )

    # Write-Host "Copying .pbix to a .zip file"

    # Check the file exists
    if (-not(Test-Path $FullFilePath)) { break }

    New-Item -ItemType Directory -Force -Path $LocalZipFolder | Out-Null

    $FileName = Split-Path -Path $FullFilePath -Leaf
    $FileWithoutExtension = $FileName.Replace(".pbix", "")
    
    ########## # Separate part to add metadata to filename - ##TODO   
    #$Comment = Read-Host 'What is the file description'
    #$FileLastSaved = (Get-Item $FullFilePath).LastWriteTime
    #$DateStamp = " "+$FileLastSaved.ToString("yyyy-MM-dd_HHmm")
    ##########

    $Comment = "" # don't implimement this yet ##TODO
    $DateStamp = "" # don't impliment timestamp for now ##TODO

    #####delete $fileObj = Get-Item $FullFilePath

    $FileNewName = $FileWithoutExtension + $DateStamp + $Comment + ".zip"
    $FileNewPath = (Join-Path $LocalZipFolder $FileNewName)

    Copy-Item -Path $FullFilePath -destination $FileNewPath | Out-Null
    
    return $FileNewPath
}
function Copy-ZipToFolders {
    
    param ( [String]$Localzipfile, [String]$Localoutdir )

    # Write-Host "Unzipping .zip file to folder $outdir"
    Remove-Item -LiteralPath $Localoutdir -Recurse -Force -ErrorAction SilentlyContinue | Out-Null


    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Localzipfile)
    try {
        foreach ($entry in $archive.Entries) {
            $entryTargetFilePath = [System.IO.Path]::Combine($Localoutdir, $entry.FullName)
            $entryDir = [System.IO.Path]::GetDirectoryName($entryTargetFilePath)

            #Ensure the directory of the archive entry exists
            if (!(Test-Path $entryDir )) {
                New-Item -ItemType Directory -Path $entryDir | Out-Null 
            }

            #If the entry is not a directory entry, then extract entry
            if (!$entryTargetFilePath.EndsWith("\")) {
                [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $entryTargetFilePath.Replace("`[", "").Replace("`]", ""), $true);
            }
        }
    }
    finally {
        $archive.Dispose()

        # Remove zipfile and zipcopy Folder
        Remove-Item (Split-Path -Path $Localzipfile) -Recurse -Force | Out-Null
    }
}
function Format-Json {
    <#
    .SYNOPSIS
        Prettifies JSON output.
    .DESCRIPTION
        Reformats a JSON string so the output looks better than what ConvertTo-Json outputs.
    .PARAMETER Json
        Required: [string] The JSON text to prettify.
    .PARAMETER Minify
        Optional: Returns the json string compressed.
    .PARAMETER Indentation
        Optional: The number of spaces (1..1024) to use for indentation. Defaults to 4.
    .PARAMETER AsArray
        Optional: If set, the output will be in the form of a string array, otherwise a single string is output.
    .EXAMPLE
        $json | ConvertTo-Json  | Format-Json -Indentation 2
    #>
    [CmdletBinding(DefaultParameterSetName = 'Prettify')]
    param (
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [string]$Json,

        [Parameter(ParameterSetName = 'Minify')]
        [switch]$Minify,

        [Parameter(ParameterSetName = 'Prettify')]
        [ValidateRange(1, 1024)]
        [int]$Indentation = 2,

        [Parameter(ParameterSetName = 'Prettify')]
        [switch]$AsArray
    )

    # Don't reformat if using Powershell version 7+
    if ($PSVersionTable.PSVersion.Major -eq 7) { 
        return $Json 
    }
    if ($PSCmdlet.ParameterSetName -eq 'Minify') {
        return ($Json | ConvertFrom-Json) | ConvertTo-Json -Depth 100 -Compress
    }

    # If the input JSON text has been created with ConvertTo-Json -Compress
    # then we first need to reconvert it without compression
    if ($Json -notmatch '\r?\n') {
        $Json = try { ($Json | ConvertFrom-Json) | ConvertTo-Json -Depth 100 } catch { $Json }
    }

    $indent = 0
    $regexUnlessQuoted = '(?=([^"]*"[^"]*")*[^"]*$)'

    $result = $Json -split '\r?\n' |
    ForEach-Object {
        # If the line contains a ] or } character, 
        # we need to decrement the indentation level unless it is inside quotes.
        if ($_ -match "[}\]]$regexUnlessQuoted") {
            $indent = [Math]::Max($indent - $Indentation, 0)
        }

        # Replace all colon-space combinations by ": " unless it is inside quotes.
        $line = (' ' * $indent) + ($_.TrimStart() -replace ":\s+$regexUnlessQuoted", ': ')

        # If the line contains a [ or { character, 
        # we need to increment the indentation level unless it is inside quotes.
        if ($_ -match "[\{\[]$regexUnlessQuoted") {
            $indent += $Indentation
        }

        $line
    }

    if ($AsArray) { return $result }
    return $result -Join [Environment]::NewLine
}
function Format-XML ([xml]$xml, $indent = 2) {
    $StringWriter = New-Object System.IO.StringWriter
    $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter
    $xmlWriter.Formatting = “indented”
    $xmlWriter.Indentation = $Indent
    $xml.WriteContentTo($XmlWriter)
    $XmlWriter.Flush()
    $StringWriter.Flush()
    Write-Output $StringWriter.ToString()
}
function Encode-PathString { 
    param ( [String] $string )
    #todo ask why ... doesn't get encoded...
    $outstring = [System.Web.HttpUtility]::UrlEncode($string).Replace("*", "_").Replace("...", "") #$string.ToString().Replace("\", "_").Replace("/", "_").Replace("%", "Perc").Replace(",", "").Replace("""", "").Replace(":", "").Replace(".", "_").Replace("&", "_").Replace("*", "_")
    return $outstring
}
function Decode-PathString { 
    param ( [String] $string )

    $outstring = [System.Web.HttpUtility]::UrlDecode($string)
    return $outstring
}
#endregion export_reportpages_without_pbitools

#region dependencyanalysis
function Export-CalculationDependencies {
    param
    (
        [string] $LocalPbixFolder, [string]$LocalServer, [string]$LocalDatabase, [string] $LocalPBIXName
    )
    
    #########################################################################
    # Load the AnalysisServices client
    
    [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.AnalysisServices.AdomdClient") 
    #[Microsoft.AnalysisServices.AdomdClient.AdomdConnection]  
        
    # Create the first connection object  
    $con = new-object Microsoft.AnalysisServices.AdomdClient.AdomdConnection 
    $con.ConnectionString = "Datasource=$LocalServer; Initial Catalog=$LocalDatabase;timeout=0; connect timeout =0" 
    $con.Open()
    
    # Create a command and send a query 
    $command = $con.CreateCommand()
    
            
    # Build output filename    
    $outFileFullPath = "$($LocalPbixFolder)\$($LocalPBIXName)_CalculationDependencies.csv"
        
    $query = "select 
        [Referenced_Table] AS [ObjectTable],
        [Referenced_Object] AS [Object],
        [Referenced_Object_Type] as [ObjectType],
        'Model' as [DependencyType],
        '$($LocalPBIXName)' as [PBIX],
        [Table] AS [DependentObjectHome],
        [Object] AS [DependentObject],
        [Object_type] AS [DependentObjectType],
        [Expression] as [DependentObjectExpression]
        from `$SYSTEM.DISCOVER_CALC_DEPENDENCY
        where [Referenced_Object_Type] <> 'Table' and [Referenced_Object_Type] <> 'Calc_Table' 
        and [Referenced_Object_Type] <> 'Active_Relationship'
        and [Referenced_Object_Type] <> 'Relationship'
        and [Referenced_Object_Type] <> 'RowNumber' "

    $command.CommandText = $query
    $adapter = New-Object -TypeName Microsoft.AnalysisServices.AdomdClient.AdomdDataAdapter $command
    $dataset = New-Object -TypeName System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null

    # DataTable definition
    $dtable = New-Object System.Data.DataTable
    #$dtable.Columns.Add("Table", "System.String") | Out-Null
    #$dtable.Columns.Add("Field", "System.String") | Out-Null
    #$dtable.Columns.Add("Object", "System.String") | Out-Null
    $dtable.Columns.Add("PBIX", "System.String") | Out-Null
    $dtable.Columns.Add("ObjectKey", "System.String") | Out-Null
    $dtable.Columns.Add("ObjectType", "System.String") | Out-Null
    $dtable.Columns.Add("DependencyType", "System.String") | Out-Null
    $dtable.Columns.Add("DependentObjectType", "System.String") | Out-Null
    $dtable.Columns.Add("DependentObjectHome", "System.String") | Out-Null
    $dtable.Columns.Add("DependentObject", "System.String") | Out-Null
    $dtable.Columns.Add("DependentObjectExpression", "System.String") | Out-Null
    
    # Set HostFile name once
    $dtable.Columns["PBIX"].DefaultValue = $LocalPBIXName
    
    # Instantiate required object
    $textInfo = (Get-Culture).TextInfo

    # Convert String
        
    foreach ($o in $dataset.Tables[0]) {
        if ($o.Object.StartsWith('RowNumber')) {}
        else {
            $nRow = $dtable.NewRow()
            #$nRow.Object = $o.Object
            $nRow.ObjectKey = Join-TableAndField $o.ObjectTable $o.Object $o.ObjectType
            $nRow.ObjectType = $textInfo.ToTitleCase($o.ObjectType.ToLower())
            $nRow.DependencyType = 
            if ($o.ObjectType -eq "ATTRIBUTE_HIERARCHY" -and $o.DependentObjectType -eq "ATTRIBUTE_HIERARCHY") { "Model Sort By" }
            else { "Model " + $textInfo.ToTitleCase($o.DependentObjectType.ToLower()) }
            $nRow.DependentObjectType = $textInfo.ToTitleCase($o.DependentObjectType.ToLower())
            $nRow.DependentObjectHome = $o.DependentObjectHome
            $nRow.DependentObject = Join-TableAndField $o.DependentObjectHome $o.DependentObject $o.DependentObjectType
            $nRow.DependentObjectExpression = $o.DependentObjectExpression
            #$nRow.Table = $filter.Expression.Column.Expression.SourceRef.Entity
            #$nRow.Field = $filter.Expression.Column.Property
            $dtable.Rows.Add($nRow)
        }
    } # End of function Export-CalculationDependencies
        
    while (Test-FileLock $outFileFullPath) {
        Read-Host "`nPlease close file and press key to try again: $outFileFullPath"
    }

    try {
        $dtable | Sort-Object -Property ObjectType, Object | Export-csv -UseQuotes Always -path $outFileFullPath -UseCulture -notypeinformation  
    }
    catch {
        $dtable | Sort-Object -Property ObjectType, Object | Export-csv -path $outFileFullPath -UseCulture -notypeinformation  

    }
    # Open the created csv, replace unwanted characters and resave
    #(Get-Content $outFileFullPath) -replace '[\[\]\"]+' | Out-File $outFileFullPath
            
    # Close the connection 
    $con.Close() 
    
}
function Export-ReportFieldDependencies {
    param (
        [Parameter(Mandatory = $true, Position = 0)] [string] $LocalFolder,
        [Parameter(Mandatory = $true, Position = 1)] [pscustomobject] $LocalReportJson,
        [Parameter(Mandatory = $true, Position = 2)] [string] $LocalPBIXName
    )

    $ArrayOfDependencies = @()
    
    # Add report filters
    if ($null -eq $LocalReportJson.filters ) {}
    else {
        foreach ($filter in ($LocalReportJson.filters | ConvertFrom-Json)) {
            if ($null -eq $filter.Expression.Column) { }
            else {

                $ArrayOfDependencies += [pscustomobject]@{
                    'PBIX'                      = $LocalPBIXName
                    'ObjectKey'                 = Join-TableAndField $filter.Expression.Column.Expression.SourceRef.Entity $filter.Expression.Column.Property
                    'ObjectType'                = "Column"
                    'DependencyType'            = "Filter"
                    'DependentObjectType'       = "Report Filter"
                    'DependentObjectHome'       = "All Pages"
                    'DependentObject'           = "All Visuals"
                    'DependentObjectExpression' = ""
                }
            }
            # Add GroupRef used in filter
            if ($null -eq $filter.Expression.GroupRef) { }
            else {

                $ArrayOfDependencies += [pscustomobject]@{
                    'PBIX'                      = $LocalPBIXName
                    'ObjectKey'                 = Join-TableAndField $filter.Expression.GroupRef.Expression.SourceRef.Entity $filter.Expression.GroupRef.Property
                    'ObjectType'                = "GroupRef"
                    'DependencyType'            = "Filter"
                    'DependentObjectType'       = "Report Filter"
                    'DependentObjectHome'       = "All Pages"
                    'DependentObject'           = "All Visuals"
                    'DependentObjectExpression' = ""
                }
            }

            # Add Hierarchy used in filter
            if ($null -eq $filter.Expression.HierarchyLevel) { }
            else {

                $ArrayOfDependencies += [pscustomobject]@{
                    'PBIX'                      = $LocalPBIXName
                    'ObjectKey'                 = Join-TableAndField $filter.Expression.HierarchyLevel.Expression.Hierarchy.Expression.SourceRef.Entity $filter.Expression.HierarchyLevel.Expression.Hierarchy.Hierarchy
                    'ObjectType'                = "HierarchyLevel"
                    'DependencyType'            = "Filter"
                    'DependentObjectType'       = "Report Filter"
                    'DependentObjectHome'       = "All Pages"
                    'DependentObject'           = "All Visuals"
                    'DependentObjectExpression' = ""
                }
            }
        }
    }

    # Report Pages
    foreach ($ReportPage in $LocalReportJson.sections) {

        $PageFolder = Get-PageNameProperties($ReportPage)
        $PageFolder = Decode-PathString ($PageFolder)

        $reportpagefilters = $reportpage.filters 
        if ('[]' -eq $reportpagefilters -or $null -eq $reportpagefilters) {}
        else {
            foreach ($filter in $reportpagefilters | ConvertFrom-Json) {
                if ($null -eq $filter.Expression.Column) { }
                else {

                    $ArrayOfDependencies += [pscustomobject]@{
                        'PBIX'                      = $LocalPBIXName
                        'ObjectKey'                 = Join-TableAndField $filter.Expression.Column.Expression.SourceRef.Entity $filter.Expression.Column.Property
                        'ObjectType'                = "Column"
                        'DependencyType'            = "Filter"
                        'DependentObjectType'       = "Page Filter"
                        'DependentObjectHome'       = $PageFolder
                        'DependentObject'           = "All Visuals on Page"
                        'DependentObjectExpression' = ""
                    }
                }
                if ($null -eq $filter.Expression.GroupRef ) { }
                else {
                    
                    $ArrayOfDependencies += [pscustomobject]@{
                        'PBIX'                      = $LocalPBIXName
                        'ObjectKey'                 = Join-TableAndField $filter.Expression.GroupRef.Expression.SourceRef.Entity $filter.Expression.GroupRef.Property
                        'ObjectType'                = "GroupRef"
                        'DependencyType'            = "Filter"
                        'DependentObjectType'       = "Page Filter"
                        'DependentObjectHome'       = $PageFolder
                        'DependentObject'           = "All Visuals on Page"
                        'DependentObjectExpression' = ""
                    }
                }
                if ($null -eq $filter.Expression.HierarchyLevel) { }
                else {

                    $ArrayOfDependencies += [pscustomobject]@{
                        'PBIX'                      = $LocalPBIXName
                        'ObjectKey'                 = Join-TableAndField $filter.Expression.HierarchyLevel.Expression.Hierarchy.Expression.SourceRef.Entity $filter.Expression.HierarchyLevel.Expression.Hierarchy.Hierarchy
                        'ObjectType'                = "HierarchyLevel"
                        'DependencyType'            = "Filter"
                        'DependentObjectType'       = "Page Filter"
                        'DependentObjectHome'       = $PageFolder
                        'DependentObject'           = "All Visuals on Page"
                        'DependentObjectExpression' = ""
                    }
                }
            }
        }

        # Add Visual dependencies
        $visualContainers = $ReportPage.visualcontainers 
        if ('[]' -eq $visualContainers -or $null -eq $visualContainers) {}
        else {
            foreach ($visual in $visualContainers) {
                $vis = ($visual.config | ConvertFrom-Json)

                $VisualFolder, $VisualName, $LayoutsPositionTabOrder = Get-VisualProperties($vis)
                $VisualFolder = Decode-PathString $VisualFolder
                $select = $vis.singleVisual.prototypequery.Select
                $from = $vis.singleVisual.prototypequery.From
                
                $ttable = New-Object System.Data.DataTable
                $ttable.Columns.Add("Name", "System.String") | Out-Null
                $ttable.Columns.Add("Entity", "System.String") | Out-Null
                #$ttable.Columns.Add("Type", "System.Int32") | Out-Null

                if ($null -eq $visual.filters) {}
                else {
                    foreach ($filter in ($visual.filters | ConvertFrom-Json )) {
                        
                        # Add visual level filters from Measure
                        if ($null -eq $filter.Expression.Measure) { }
                        else {
                            
                            $ArrayOfDependencies += [pscustomobject]@{
                                'PBIX'                      = $LocalPBIXName
                                'ObjectKey'                 = Join-TableAndField $filter.Expression.Measure.Expression.SourceRef.Entity $filter.Expression.Measure.Property "Measure"
                                'ObjectType'                = "Measure"
                                'DependencyType'            = "Filter"
                                'DependentObjectType'       = "Visual Filter"
                                'DependentObjectHome'       = $PageFolder
                                'DependentObject'           = $VisualFolder
                                'DependentObjectExpression' = ""
                            }
                        }

                        # Add visual level filters from Column
                        if ($null -eq $filter.Expression.Column) { }
                        else {
                            
                            $ArrayOfDependencies += [pscustomobject]@{
                                'PBIX'                      = $LocalPBIXName
                                'ObjectKey'                 = Join-TableAndField $filter.Expression.Column.Expression.SourceRef.Entity $filter.Expression.Column.Property
                                'ObjectType'                = "Column"
                                'DependencyType'            = "Filter"
                                'DependentObjectType'       = "Visual Filter"
                                'DependentObjectHome'       = $PageFolder
                                'DependentObject'           = $VisualFolder
                                'DependentObjectExpression' = ""
                            }
                        }
                    }
                }

                if ($null -eq $from) {
                    Write-Host $vis
                }
                else {
                    foreach ($t in $from) {
                        $nRow = $ttable.NewRow()
                        $nRow.Name = $t.name
                        $nRow.Entity = $t.Entity
                        #$nRow.Type = $t.Type ?? ""
                        $ttable.Rows.Add($nRow)
                    }
                }

                foreach ($item in $select) {
                    if ($null -eq $item.Measure) { }
                    else {
                        
                        $ArrayOfDependencies += [pscustomobject]@{
                            'PBIX'                      = $LocalPBIXName
                            'ObjectKey'                 = Join-TableAndField $ttable.Select($expression)[0].Entity $item.Measure.Property "Measure"
                            'ObjectType'                = "Measure"
                            'DependencyType'            = "Visual"
                            'DependentObjectType'       = $vis.singleVisual.visualType
                            'DependentObjectHome'       = $PageFolder
                            'DependentObject'           = $VisualFolder
                            'DependentObjectExpression' = ""
                        }
                    }
                    if ( $null -eq $item.Column) { }
                    else {
                        $expression = "Name = '" + $item.Column.Expression.SourceRef.Source + "'"
                        
                        $ArrayOfDependencies += [pscustomobject]@{
                            'PBIX'                      = $LocalPBIXName
                            'ObjectKey'                 = Join-TableAndField $ttable.Select($expression)[0].Entity $item.Column.Property
                            'ObjectType'                = "Column"
                            'DependencyType'            = "Visual"
                            'DependentObjectType'       = $vis.singleVisual.visualType
                            'DependentObjectHome'       = $PageFolder
                            'DependentObject'           = $VisualFolder
                            'DependentObjectExpression' = ""
                        }
                    }
                    if ($null -eq $item.Aggregation) { }
                    else {
                        $expression = "Name = '" + $item.Aggregation.Expression.Column.Expression.SourceRef.Source + "'"
                        
                        $ArrayOfDependencies += [pscustomobject]@{
                            'PBIX'                      = $LocalPBIXName
                            'ObjectKey'                 = Join-TableAndField $ttable.Select($expression)[0].Entity $item.Aggregation.Expression.Column.Property #$string
                            'ObjectType'                = "Aggregation"
                            'DependencyType'            = "Visual"
                            'DependentObjectType'       = $vis.singleVisual.visualType
                            'DependentObjectHome'       = $PageFolder
                            'DependentObject'           = $VisualFolder
                            'DependentObjectExpression' = ""
                        }
                    }
                }
            }
        }
    }
    $newfile = Join-Path $LocalFolder ($LocalPBIXName + "_ReportDependencies.csv")
    $ArrayOfDependencies | Sort-Object -Property ObjectKey, ObjectType, Object | export-csv $newfile -notypeinformation

}
function Export-ReportFieldDependenciesByRegex {
    param (
        [Parameter(Mandatory = $true, Position = 0)] [string] $LocalOutFolder,
        [Parameter(Mandatory = $true, Position = 1)] [string] $LocalReportFolder,
        [Parameter(Mandatory = $true, Position = 2)] [string] $LocalPBIXName
    )

    $MatchType = '((?<MatchType>"webURL":|"databars":|"Icon":|"title":|"color":).*?)'
    $Entity = '((?<FieldType>Column|Measure?).*?)?:{"Expression":{"SourceRef":{"Entity":(?<Entity>.+?)}}?(,"Property":)(?<Property>.*?)}}'
    $ArrayOfDependencies = @()

    #foreach ($FieldType in $FieldTypes) {
    $RegexMainnnString = ($MatchType + $Entity)
    $r = [regex] ($RegexMainnnString)
    Get-ChildItem -Path $LocalReportFolder -Include "config.json" -Recurse -File  | 
    ForEach-Object {
        $content = (Get-Content $_.FullName -Raw).Replace("`r`n", "").Replace(" ", "")
        #$contentJson = $content | ConvertFrom-Json
        $t = $r.Matches($content);
        $fileFullName = $_.FullName
        if ($null -ne $t) {
            foreach ($res in $t) {
                ####$MatchType = $res.Groups['MatchType'].Value.Replace("`"", "'")
                $TableRef = $res.Groups['Entity'].Value.Replace("`"", "'")
                $FieldRef = $res.Groups['Property'].Value.Replace("`"", "")
                $FieldType = $res.Groups['FieldType'].Value.Replace("`"", "")
                    
                $ArrayOfDependencies += [pscustomobject]@{
                    'PBIX'                      = $LocalPBIXName
                    'ObjectKey'                 = Join-TableAndField $TableRef $FieldRef $FieldType
                    'ObjectType'                = $res.Groups['FieldType'].Value
                    'DependencyType'            = if ($res.Groups['MatchType'].Success) { $res.Groups['MatchType'].Value.Replace("`"", "") } 
                    else { "Unmapped" }
                    'DependentObjectType'       = ( $content | ConvertFrom-Json ).singleVisual.visualType
                    'DependentObjectHome'       = Split-Path $fileFullName | Split-Path | Split-Path | Split-Path -Leaf
                    'DependentObject'           = Split-Path ( Split-Path $_ ) -Leaf
                    'DependentObjectExpression' = ""
                }
            }
        }
        # }
    }

    $DeduplicatedDtable = ($ArrayOfDependencies | Sort-Object -Unique -Property { $_.PBIX + $_.ObjectKey + $_.ObjectType + $_.DependencyType + $_.DepentObjectType + $_.DependentObjectHome + $_.DependentObject })

    $newfile = Join-Path $LocalOutFolder ($LocalPBIXName + "_ReportDependenciesByRegex.csv")
    $DeduplicatedDtable | export-csv $newfile -notypeinformation
}
function Export-UnusedFieldScript {
    param (
        [Parameter(Mandatory = $true, Position = 1)] [string] $LocalDependenciesFolder,
        [Parameter(Mandatory = $true, Position = 2)] [string] $LocalPBIXName
    )

    # DataTable definition
    $dependencies = New-Object System.Data.DataTable
    $allFields = New-Object System.Data.DataTable

    # Load all model fields from file
    $AllFieldsFile = Join-Path $LocalDependenciesFolder "$($LocalPBIXName)_AllFields.csv"
    $allFields = Import-Csv $AllFieldsFile
    
    $ListOfFilesToLoad = @()

    # 'Normal' report dependencies such as measures and columns in barcharts or table/matrix
    $ListOfFilesToLoad += (Get-ChildItem $LocalDependenciesFolder -Filter *_ReportDependencies.csv) | Select-Object -Property Fullname

    # 'Special case' report dependencies such as WebUrl, fx Text, conditional formatting etc.
    $ListOfFilesToLoad += (Get-ChildItem $LocalDependenciesFolder -Filter *_ReportDependenciesByRegex.csv) | Select-Object -Property Fullname
    
    # Calculation dependencies
    $ListOfFilesToLoad += (Get-ChildItem $LocalDependenciesFolder -Filter "$($LocalPBIXName)_CalculationDependencies.csv") | Select-Object -Property Fullname

    # Load all dependencies from file and make unique list of ObjectKey
    $dependencies = Import-Csv $ListOfFilesToLoad.FullName
    $DependentObjectList = $dependencies | Select-Object ObjectKey | sort-object -Property ObjectKey -Unique
    
    # Create scripts folder if it doesn't exist
    $LocalScriptsFolder = Join-path $LocalDependenciesFolder "Scripts to hide unused fields"
    if (Test-Path $LocalScriptsFolder) {} 
    else { New-Item -ItemType Directory -Force -Path $LocalScriptsFolder | Out-Null }

    $OrganiseUnusedFieldScript = @()
    $OrganiseUnusedFieldScriptFilePath = Join-Path $LocalScriptsFolder "$($LocalPBIXName)_UnusedFieldsDisplayFolderScript.cs"
    
    # File new unused fields by pre-pending Display Folder with "_Unused\"
    $unusedFields = $allFields | 
    Where-Object { 
        ( $DependentObjectList.ObjectKey -notcontains $_.ObjectKey ) -and 
        ( $_.ObjectType -ne "Row Number" ) -and 
        ( $_.DisplayFolder -notmatch "_Unused" ) -and
        ( $_.Table -notmatch "LocalDateTable_" ) # Ignore temporary date tables, we don't care if they are unused
    }

    if ($unusedFields.Count -ne 0) { 
        Write-Host "`n             $($unusedFields.Count) new unused fields found...."  -Foregroundcolor Cyan -NoNewLine 
    }

    foreach ($field in $unusedFields) {
        switch ($field.ObjectType) {
            "Measure" { $TOM_objType = "Measures" }
            "Calculated Column" { $TOM_objType = "Columns" }
            "Column" { $TOM_objType = "Columns" }
            "Hierarchy" { $TOM_objType = "Hierarchies" }
            "Calculated Table Column" { $TOM_objType = "Columns" }
            default {}
        }

        # Generate script lines to organise unused fields into _Unused display folder and hide from report
        $FieldHandle = "Model.Tables[""$($field.Table)""].$TOM_objType[""$($field.Object)""]"
        $NewDisplayFolder = "$($FieldHandle).DisplayFolder.Replace(""_Unused\\"","""").Replace(""_Unused"","""")"
        $OrganiseUnusedFieldScript += "$($FieldHandle).DisplayFolder = ""_Unused\\"" + $($NewDisplayFolder);"
        $OrganiseUnusedFieldScript += "$($FieldHandle).IsHidden = true;"
    }

    # Take any newly used fields out of _Unused\ dispaly folder (retain any existing other subfolder)
    $UsedFieldsCurrentlyInUnusedFolder = $allFields | 
    Where-Object { 
        ( $DependentObjectList.ObjectKey -contains $_.ObjectKey ) -and 
        ( $_.ObjectType -ne "Row Number" ) -and 
        ( $_.DisplayFolder -match "_Unused" ) 
    }

    if ($UsedFieldsCurrentlyInUnusedFolder.Count -ne 0) { 
        Write-Host "`n             $($UsedFieldsCurrentlyInUnusedFolder.Count) new used fields found." -Foregroundcolor Cyan -NoNewLine 
    } 

    foreach ($field in $UsedFieldsCurrentlyInUnusedFolder) {
        switch ($field.ObjectType) {
            "Measure" { $TOM_objType = "Measures" }
            "Calculated Column" { $TOM_objType = "Columns" }
            "Column" { $TOM_objType = "Columns" }
            "Hierarchy" { $TOM_objType = "Hierarchies" }
            "Calculated Table Column" { $TOM_objType = "Columns" }
            default {}
        }

        # Generate script lines to take newly used fields out of _Unused display folder
        $FieldHandle = "Model.Tables[""$($field.Table)""].$TOM_objType[""$($field.Object)""]"
        $NewDisplayFolder = "$($FieldHandle).DisplayFolder.Replace(""_Unused\\"","""").Replace(""_Unused"","""")"
        $OrganiseUnusedFieldScript += "$($FieldHandle).DisplayFolder = $($NewDisplayFolder);"
    }
    
    # Add one blank line to script to ensure file is 'emptied' if there are no fields to organise
    $OrganiseUnusedFieldScript += ""

    # Write the script to file
    $OrganiseUnusedFieldScript | Set-Content $OrganiseUnusedFieldScriptFilePath
} 
function Export-ModelFields {
    param ( 
        [string]$LocalPbixFolder, 
        [string]$LocalServer, 
        [string]$LocalDatabase, 
        [string]$LocalPBIXName
    )
    #TODO
        
    $saveasfile = "$($LocalPBIXName)_AllFields.csv";
    $saveas = Join-Path $LocalPbixFolder $saveasfile
    
    if (Test-Path $LocalPbixFolder) {} 
    else { New-Item -ItemType Directory -Force -Path $LocalPbixFolder | Out-Null }

    $as = New-Object Microsoft.AnalysisServices.Tabular.Server;
    $as.Connect($LocalServer);
    $db = $as.Databases[$LocalDatabase];
    

    # DataTable definition
    $dtable = New-Object System.Data.DataTable
    #$dtable.Columns.Add("Table", "System.String") | Out-Null
    #$dtable.Columns.Add("Field", "System.String") | Out-Null
    $dtable.Columns.Add("Table", "System.String") | Out-Null
    $dtable.Columns.Add("ObjectType", "System.String") | Out-Null
    $dtable.Columns.Add("Object", "System.String") | Out-Null
    $dtable.Columns.Add("ObjectKey", "System.String") | Out-Null
    $dtable.Columns.Add("Expression", "System.String") | Out-Null
    $dtable.Columns.Add("DisplayFolder", "System.String") | Out-Null
    $dtable.Columns.Add("Description", "System.String") | Out-Null
    $dtable.Columns.Add("IsHidden", "System.String") | Out-Null
    $dtable.Columns.Add("DataType", "System.String") | Out-Null
    $dtable.Columns.Add("FormatString", "System.String") | Out-Null
    $dtable.Columns.Add("TOMType", "System.String") | Out-Null
    $dtable.Columns.Add("TOMAs", "System.String") | Out-Null
    
    #$MeasureResult = @()
    foreach ($t in $db.Model.Tables) {
        foreach ($M in $t.Measures) {
            #$MeasureResult +=  "[$($M.Name)]" #,$M.Expression;
            $nRow = $dtable.NewRow()
            $nRow.Table = $t.Name
            $nRow.ObjectType = "Measure"
            $nRow.Object = $M.Name
            $nRow.ObjectKey = Join-TableAndField $t.Name $M.Name "Measure"
            $nRow.Expression = $M.Expression
            $nRow.DisplayFolder = $M.DisplayFolder
            $nRow.Description = $M.Description
            $nRow.IsHidden = $M.IsHidden
            $nRow.DataType = $M.DataType.ToString()
            $nRow.FormatString = $M.FormatString
            $nRow.TOMType = "Measures"
            $nRow.TOMAs = "Measure"
            $dtable.Rows.Add($nRow)   
        
        }
        foreach ($C in $t.Columns) {
            #$MeasureResult +=  "[$($M.Name)]" #,$M.Expression;
            $nRow = $dtable.NewRow()
            $nRow.Table = $t.Name
            $nRow.ObjectType = switch ($C.type) { 
                "Calculated" { "Calculated Column" } 
                "CalculatedTableColumn" { "Calculated Table Column" } 
                "RowNumber" { "Row Number" }
                "Data" { "Column" }
                default { $C.type }
            }
            $nRow.Object = $C.Name
            $nRow.ObjectKey = Join-TableAndField $t.Name $C.Name "Column"
            $nRow.Expression = $C.Expression
            $nRow.DisplayFolder = $C.DisplayFolder
            $nRow.Description = $C.Description
            $nRow.IsHidden = $C.IsHidden
            $nRow.DataType = $C.DataType.ToString()
            $nRow.FormatString = $C.FormatString
            $nRow.TOMType = "Columns"
            $nRow.TOMAs = "Column"
            $dtable.Rows.Add($nRow)   
        
        }
        foreach ($H in $t.Hierarchies) {
            #$MeasureResult +=  "[$($M.Name)]" #,$M.Expression;
            $nRow = $dtable.NewRow()
            $nRow.Table = $t.Name
            $nRow.ObjectType = "Hierarchy"
            $nRow.Object = $H.Name
            $nRow.ObjectKey = Join-TableAndField $t.Name $H.Name "Column"
            $nRow.Expression = $H.Expression
            $nRow.DisplayFolder = $H.DisplayFolder
            $nRow.Description = $H.Description
            $nRow.IsHidden = $H.IsHidden
            $nRow.DataType = ""
            $nRow.FormatString = ""
            $nRow.TOMType = "Hierarchies"
            $nRow.TOMAs = "Hierarchy"
            $dtable.Rows.Add($nRow) 
        }
        
    }

    $as.Disconnect();
    
    ####$out = $out.Replace("`t","  "); # I prefer spaces over tabs :-)
    while (Test-FileLock $saveas) {
        Read-Host "Please close file and try again: $saveas"
    }
    try {
        $dtable | Sort-Object -Property ObjectType, Object | Export-csv -UseQuotes Always -path $saveas -UseCulture -notypeinformation 
    }
    catch {
        $dtable | Sort-Object -Property ObjectType, Object | Export-csv -path $saveas -UseCulture -notypeinformation 
    }
}
function Get-VisualProperties($visConfig) {
    
    $DefaultVisualName = $visConfig.singleVisual.visualType + "_" + $visConfig.name
    $vcObjectTitle = $visConfig.singleVisual.vcObjects.title.properties.text.expr.Literal.Value
    $VisualTitle = $visConfig.singleVisual.title.properties.text.expr.Literal.Value

    $VisualName = 
    if ($null -eq $vcObjectTitle) {
        if ($null -eq $VisualTitle) {
            $DefaultVisualName + "x"
        }
        else {
            $VisualTitle + "y"
        }
    }
    else {
        $vcObjectTitle.Substring(1, $vcObjectTitle.Length - 2)
    }


    # Determine the visual display order from various places the tabOrder can be set

    $LayoutsTabOrder = $visConfig.layouts.tabOrder

    $LayoutsPositionTabOrder = $visConfig.layouts.position.tabOrder
    [string]$VisualTabOrder = 
    if ($null -eq $LayoutsPositionTabOrder) {
        if ($null -eq $LayoutsTabOrder) { 0000 } 
        else { $LayoutsTabOrder }
    } 
    else {
        $LayoutsPositionTabOrder
    }

    $VisualFolder = $VisualTabOrder.ToString().Trim().PadLeft(4, "0") + "_" + (Encode-PathString $VisualName)
    return $VisualFolder, $VisualName, $LayoutsPositionTabOrder
}
function Join-TableAndField {
    param ([String] $table, [String] $field, [String] $fieldType )
    
    $output =
    switch ( $fieldType.ToLower() ) {
        "measure" { "[" + $field + "]"; Break }
        "calc_table" { "'" + $field + "'"; Break }
        default { "'" + $table + "'[" + $field + "]" } 
    }
    $output = $output.Replace("`'`'", "")
    return $output
}
function Get-PageNameProperties ($page) {
    $PageFolder = "0" + ($page.ordinal + 1).ToString().PadLeft(1, "0") + "_" + (Encode-PathString $page.displayName)
    return $PageFolder
}

#endregion dependencyanalysis

#region toolkit_ExtendedFeatures
function Copy-FileWithTimestamp {

    param
    ( 
        [String] $FullFilePath,
        [String] $LocalQuickBackupFolder
    )

    if (Test-Path $LocalQuickBackupFolder) {} 
    else { 
        New-Item -ItemType Directory -Force -Path $LocalQuickBackupFolder | Out-Null 
    }
    
    # Check the pbix file exists
    if (-not(Test-Path $FullFilePath)) {
        pause "Cannot find the pbix to copy"; 
        break
    }

    $FileName = Split-Path -Path $FullFilePath -Leaf
    $FileWithoutExtension = $FileName.Replace(".pbix", "")

    # Get file last saved timestamp
    $FileLastSaved = (Get-Item $FullFilePath).LastWriteTime
    $DateStamp = $FileLastSaved.ToString("yyyy-MM-ddThhmmss")
    
    #Ask user for comment to append
    Write-Host 'What is the file description:' -ForegroundColor Yellow -NoNewline
    $Comment = Read-Host

    $FileNewName = "$FileWithoutExtension $DateStamp $Comment.pbix"
    $FileNewPath = Join-Path $LocalQuickBackupFolder $FileNewName
    Copy-Item -Path $FullFilePath -destination $FileNewPath
    Write-Host "`nBackup PBIX copy created in folder $LocalQuickBackupFolder" -ForegroundColor Cyan
    Write-Host "  `"$FileNewName`""
}
function Test-ReportPages {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] [string] $LocalPerformanceFolder,
        [Parameter(Mandatory = $true, Position = 1)] [string] $LocalConnectionString
    )
    
    #########################################################################
    # Load the AnalysisServices client

    # Create the first connection object  
    try { 
        $con = new-object Microsoft.AnalysisServices.AdomdClient.AdomdConnection 
        $con.ConnectionString = $LocalConnectionString
        $con.Open()
        $connectionDriver = "ADMOMD"
    }
    catch {
        $con = New-Object -TypeName System.Data.OleDb.OleDbConnection
        $con.ConnectionString = $LocalConnectionString
        $con.Open()
        $connectionDriver = "OLEDB"
    }
            
    # Create a command and send a query 
    $command = $con.CreateCommand()

    if (Test-Path $LocalPerformanceFolder) {} 
    else { 
        New-Item -ItemType Directory -Force -Path $LocalPerformanceFolder | Out-Null 
    }
        
    $PerformanceFiles = Get-ChildItem $LocalPerformanceFolder -Recurse | Where-Object { $_.Mode -notmatch "d" -and ($_.Name.EndsWith("PowerBIPerformanceData.json")) }
    if ( $PerformanceFiles.Length -eq 0 ) { 
        Write-Host "Please save PowerBIPerformanceData.json files into folder" -ForegroundColor Red
        Write-Host $LocalPerformanceFolder
    }
    else {
        foreach ($File in $PerformanceFiles) {
                
            $i = 0
            Write-Host "`nTesting page $($File.Name)..." -ForegroundColor Cyan
            $content = Get-Content $File.FullName | ConvertFrom-Json
            foreach ($DaxQueryEvent in ($content.events | Where-Object { $_.name -eq "Execute DAX Query" })) {
                $i += 1
                $query = $DaxQueryEvent.metrics.querytext
                $command.CommandText = $query
                switch ($connectionDriver) {
                    "ADMOMD" { $adapter = New-Object -TypeName Microsoft.AnalysisServices.AdomdClient.AdomdDataAdapter $command }
                    "OLEDB" { $adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $query, $con }
                }
                $dataset = New-Object -TypeName System.Data.DataSet
                try {
                    $numberOfRows = $adapter.Fill($dataset)
                    Write-Host "Visual $i succeeded with $numberOfRows rows" -NoNewline 
                }
                catch {    
                    $ErrorMessage = $_.Exception.Message
                    Write-Host "Visual $i failed on page $($File.Name) with error:" -ForegroundColor Red
                    Write-Host $ErrorMessage -ForegroundColor Red
                }
                Write-Host
            }
            #ConvertTo-Json -InputObject $content -Depth 100 | Format-Json
        }
    }
    $con.Close() 

    Write-Host "Page testing complete`n" -ForegroundColor Yellow
}
function Export-DaxQueries {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] [string] $LocalFolder,
        [Parameter(Mandatory = $true, Position = 1)] [string] $LocalConnectionString
    )
        
    #########################################################################
    # Load the AnalysisServices client

        
    # Search for .dax query files in the same directory as the pbix file
    Write-Host("`nSearching for .dax query files at this location:") -ForegroundColor Cyan
    Write-Host($LocalFolder)
        
    if (Test-Path $LocalFolder) {} 
    else { 
        New-Item -ItemType Directory -Force -Path $LocalFolder | Out-Null 
    }
        
    $DaxQueryFileNames = Get-ChildItem $LocalFolder -Filter *.dax
        
    if ($null -eq $DaxQueryFileNames) {
        Write-Host "Please save the *.dax queries you wish to export into folder" -ForegroundColor Red
        Write-Host "$LocalFolder"
    }
    else {
            
        # Create the first connection object  
        try { 
            $con = new-object Microsoft.AnalysisServices.AdomdClient.AdomdConnection 
            $con.ConnectionString = $LocalConnectionString
            $con.Open()
            $connectionDriver = "ADMOMD"
        }
        catch {
            $con = New-Object -TypeName System.Data.OleDb.OleDbConnection
            $con.ConnectionString = $LocalConnectionString
            $con.Open()
            $connectionDriver = "OLEDB"
        }
        # Create a command and send a query 
        $command = $con.CreateCommand()
        
        foreach ($File in $DaxQueryFileNames) {
            
            # Build output filename    
            $outFileName = $File.Name.Replace(".dax", ".csv")
            $outFileFullPath = "$($LocalFolder)\$(get-date -f yyyy-MM-dd) $($outFileName)"
                
            $query = (Get-Content $File.FullName)
            $command.CommandText = $query
            
            switch ($connectionDriver) {
                "ADMOMD" { $adapter = New-Object -TypeName Microsoft.AnalysisServices.AdomdClient.AdomdDataAdapter $command }
                "OLEDB" { $adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $query, $con }
            }

            $dataset = New-Object -TypeName System.Data.DataSet
                
            $numberOfRows = $adapter.Fill($dataset)
            Write-Host "$numberOfRows rows exported from $($File.Name)"
        
            while (Test-FileLock $outFileFullPath) {
                Read-Host "Please close file and try again: $outFileFullPath"
            }
            $dataset.Tables[0] | Export-csv -UseQuotes Always -path $outFileFullPath -UseCulture -notypeinformation
        
            # Replace the square brackets in the csv column headers e.g.  "[Column]" to "Column" 
            $content = Get-Content $outFileFullPath ###[System.IO.File]::ReadAllLines( $outFileFullPath )
            $content[0] = $content[0] -replace '[\[\]]' 
            $content | Out-File $outFileFullPath
        }
        
        # Close the connection 
        $con.Close() 
    }
}

function Compile-ThickReport {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] [psobject] $ls
    )

    $NewPbixFilePath = (Join-Path $ls.PbixRootFolder $ls.PbixFilePath).replace(".pbix",".pbit")

    if ((Test-Path $NewPbixFilePath) -and !($ProceedToOverwriteAll) ) {
        Write-Host "`nAre you sure you wish to overwrite the existing file(s)?" -ForegroundColor Cyan -NoNewLine
        Write-Host "`n`n$NewPbixFilePath " -ForegroundColor Cyan
        $ClearToProceed = Read-Host "(Y)Yes to confirm for this file `n(N)No to cancel for this file `n(A)To confirm for this and all further files"
    }
    else {
        $ClearToProceed = "Y"
    }

        if ( ($ClearToProceed.ToUpper() -eq "Y") ) {
            
            Write-Host "`Compiling thick report template `"$(Split-Path $NewPbixFilePath -Leaf)`" from source control..." -ForegroundColor Cyan -NoNewline

            while (Test-FileLock $NewPbixFilePath) {
                Read-Host "Please close file and press key to try again: $($NewPbixFilePath)" -ForegroundColor Red
            }
            
            $myfolder = $ls.PbixExportFolder + "\"

            # Compile pbix from Source Control using pbi-tools
            $pbitoolsReponse = pbitools compile-pbix $myfolder -outpath $ls.PbixRootFolder -format 'pbit' -overwrite
            if (!$?) {
                throw $pbitoolsReponse
            }
            else {
                Write-Host "Done."
            }
    }
}
function Compile-ThinReportWithLocalConnection {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] [psobject] $ls
    )

    $LocalThinReportDirectories = Get-ChildItem  $ls.ThinReportsFolder | Where-Object { $_.Mode -eq "d----" }
    Write-Host "`nPlease select one of the thin report folders:"
    $i = 1
    foreach ($f in $LocalThinReportDirectories  ) {
        Write-Host "$i) $($f.Name)" -ForegroundColor Cyan
        $i++
    }

    $selection = Read-Host -Prompt "`nEnter the number of the file to select (or hit enter for All)"
    if ($selection -ne "" -and $selection -le $LocalThinReportDirectories.Count) {
        $SelectedReportSourceFolders = $LocalThinReportDirectories[$selection - 1]
    }
    else {
        if ($selection -eq "") {
            $SelectedReportSourceFolders = $LocalThinReportDirectories
        }
        else {
            Exit-ActionBIToolkit "Invalid selection $($selection))... Closing"
        }
    }

    # Set flag to proceed to overwrite all thin report pbix files - initial state $False
    $ProceedToOverwriteAll = $False

    foreach ($d in $SelectedReportSourceFolders) {
        $LocalConnections = @"
{
  "Version": 1,
  "Connections": [
    {
      "Name": "EntityDataSource",
      "ConnectionString": "Data Source=$($ls.Server);Initial Catalog=$($ls.Database);Cube=Model",
      "ConnectionType": "analysisServicesDatabaseLive"
      }
  ]
}

"@

        $newLocalConnections = $LocalConnections -replace 'Data Source=.+?;', "Data Source=$($ls.Server);"
        $newLocalConnections = $newLocalConnections -replace 'Initial Catalog=.+?;', "Initial Catalog=$($ls.Database);"

        $ConnectionsJsonPath = Join-Path $d "Connections.json"
        $newLocalConnections | Out-File $ConnectionsJsonPath

        $NewPbixFilePath = Join-Path $ls.PbixRootFolder "$($d.Name).pbix"

        if ((Test-Path $NewPbixFilePath) -and !($ProceedToOverwriteAll) ) {
            Write-Host "`nAre you sure you wish to overwrite the existing file(s)?" -ForegroundColor Cyan -NoNewLine
            Write-Host "`n`n$NewPbixFilePath " -ForegroundColor Cyan
            $ClearToProceed = Read-Host "(Y)Yes to confirm for this file `n(N)No to cancel for this file `n(A)To confirm for this and all further files"
        }
        else {
            $ClearToProceed = "Y"
        }

        if ( $ClearToProceed.ToUpper() -eq "A") {
            $ProceedToOverwriteAll = $true
            Write-Host "Overwrite -All- selected"
        }

        if ( ($ClearToProceed.ToUpper() -eq "Y") -or $ProceedToOverwriteAll ) {
            
            Write-Host "`Compiling thin report `"$($d.Name).pbix`" from source control..." -ForegroundColor Cyan -NoNewline

            while (Test-FileLock $NewPbixFilePath) {
                Read-Host "Please close file and press key to try again: $($NewPbixFilePath)" -ForegroundColor Red
            }
            
            # Compile pbix from Source Control using pbi-tools
            $pbitoolsReponse = pbitools compile-pbix $d $ls.PbixRootFolder -overwrite
            if (!$?) {
                throw $pbitoolsReponse
            }
            else {
                Write-Host "Done."
            }

            # # Launch new pbix file
            # Write-Host "`nLaunching thin report `"$($d.Name).pbix`"..." -ForegroundColor Cyan -NoNewline
            # $pbitoolsReponse = pbitools launch-pbi $NewPbixFilePath | Out-Null
            # if (!$?) {
            #     throw $pbitoolsReponse
            # }
            # else {
            #     Write-Host "Done."
            # }
        }
        else {
            Write-Host "Skipping file $NewPbixFilePath" -ForegroundColor Cyan
        }
    }
}
#endregion toolkit_ExtendedFeatures

#region deployment
function Initialize-DeploymentWorkspaces {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] $DeploymentWorkspaceFolder,
        [Parameter(Mandatory = $true, Position = 1)] $DeploymentEnvironmentFolders
    )
    try {
        Get-PowerBIAccessToken | Out-Null
    }
    catch {
        Write-Host "Opening browser to authenticate to Power BI service" -ForegroundColor Cyan -NoNewline
        Connect-PowerBIServiceAccount | Out-Null
        Write-Host "Done."
        Show-Process -Process (Get-Process -Id $PID)
    }


    $workspaces = @()

    foreach ($d in $DeploymentEnvironmentFolders) {
        $wsn = $DeploymentWorkspaceFolder.Name + ($d.Name.Replace("Production",""))
        Write-Host "Retrieveing Power BI Workspace $wsn..." -ForegroundColor Cyan -NoNewline
        $ws = Get-PowerBIWorkspace -Name $wsn
        if ($ws.Count -eq 0 ) {
            Write-Host "`n`nPower BI Workspace $wsn not found..."
            Write-Host "Creating new Power BI Workspace $wsn..." -ForegroundColor Cyan -NoNewline
            $w = New-PowerBIWorkspace -Name $wsn
            $workspaces += $w
            Write-Host "Done.`n"
        }
        else {
            Write-Host "Done."
            $workspaces += $ws
        }
    }
    return $workspaces
}
function Deploy-ThickReport {
    param
    (
        [Parameter(Mandatory = $true, Position = 0)] [psobject] $ls,
        [Parameter(Mandatory = $true, Position = 1)] $Workspace,
        [Parameter(Mandatory = $true, Position = 2)] $WorkspaceFolder
    )
    try {
        Get-PowerBIAccessToken | Out-Null
    }
    catch {
        #Connect using User account
        Write-Host "Opening browser to authenticate to Power BI service" -ForegroundColor -NoNewline
        Connect-PowerBIServiceAccount | Out-Null # don't end up with the connection in the pipeline!
        Write-Host "Done."

        Show-Process -Process (Get-Process -Id $PID)
        #Connect using Service Principal
        #Connect-PowerBIServiceAccount -ServicePrincipal ####
    }
    
    $DeploymentEnvironmentPath = $WorkspaceFolder.FullName
    $DeploymentDatasetName =  $ls.PbixFileName + ($WorkspaceFolder.Name).Replace("Production","")

    Write-Host "Publishing pbix file to Power BI Workspace $DeploymentDatasetName..." -ForegroundColor Cyan -NoNewline
    # ConflictAction overwrite will fail if there are already more than one report/datasets in the workspace with the same name
    # Need to check that there is at most one
    $Report = New-PowerBIReport -Path $ls.PbixFilePath -Name $DeploymentDatasetName -WorkspaceId $Workspace.Id -ConflictAction CreateOrOverwrite
    Write-Host "Done."

    # Fetch latest import
    $Import = (Get-PowerBIImport -Scope Organization -Filter "name eq '$($Report.Name)'")[-1]
    
    $Dataset = Get-PowerBIDataset -Scope Organization -DataSetId $Import.Datasets[0].Id.Guid
    $WorkspaceConnectionString = "powerbi://api.powerbi.com/v1.0/myorg/$([uri]::EscapeDataString($WorkspaceFolder.Name))"
    
    # Save workspace and dataset information to deployment folder
    $Datasetfile = Join-Path $DeploymentEnvironmentPath "$($Dataset.Name) Dataset.json"
    $WorkspaceSettingsFile = Join-Path $DeploymentEnvironmentPath "$($WorkspaceFolder.Name) Workspace Settings.json"
    $WorkspaceConnectionStringFile = Join-Path $DeploymentEnvironmentPath "$($WorkspaceFolder.Name) Workspace ConnectionString.json"
    $Workspace | ConvertTo-Json | Set-Content $WorkspaceSettingsFile
    $Dataset | ConvertTo-Json | Set-Content $DatasetFile
    $WorkspaceConnectionString | Set-Content $WorkspaceConnectionStringFile
    
    # Launch Workspace
    
    Start-Process "https://app.powerbi.com/groups/$($Workspace.Id)"
}
#endregion deployment

# Replace server and database details here for local VS Code testing#
#[psobject]$TKS = Initialize-Toolkit "localhost:53278" "015f3b90-b5ad-4cea-8f14-f6d69704a260"

# Get the Toolkit Session settings ($TKS) for this session
[psobject]$TKS = Initialize-Toolkit $args[0] $args[1]

# Fetch selected user run option
[string]$RunOption = Get-RunOption $TKS.PbixFileType

# Add selected user run option to Toolkit Session settings ($TKS)
$TKS | Add-Member -MemberType NoteProperty -Name runOption -Value $RunOption

# Fetch the json definition for the report of the selected .pbix file
[string]$reportJson = Get-ReportJson $TKS


if ($RunOption -eq "Export PBIX for source control" -or $RunOption -eq "All" ) {
    Export-PBIX $TKS $reportJson

    # Start of dependencies analysis
    #####Write-Host "Extracting fields used in report..." -ForegroundColor Cyan -NoNewline
    Export-ReportFieldDependencies $TKS.dependenciesOutFolder $reportJson $TKS.PbixFileName
    Export-ReportFieldDependenciesByRegex $TKS.dependenciesOutFolder $TKS.PbixExportFolder $TKS.PbixFileName
    #####Write-Host " Done."
    # end of dependencies analysis

    # Extra steps if we have a thick report
    if ( $TKS.PbixFileType -eq "Thick Report: Model embedded within Local PBIX") {

        Write-Host "STARTING DEPENDENCY ANALYSIS"
        # Export a registry of all model fields for dependency analysis
        #####Write-Host "Extracting all model fields for dependency analysis..." -ForegroundColor Cyan -NoNewline
        Export-ModelFields $TKS.dependenciesOutFolder $TKS.Server $TKS.Database $TKS.PbixFileName
        #####Write-Host " Done."

        # Export all model calculation dependencies from DMVs
        #####Write-Host "Extracting calculation dependencies from DMVs..." -ForegroundColor Cyan -NoNewline
        Export-CalculationDependencies $TKS.dependenciesOutFolder $TKS.Server $TKS.Database $TKS.PbixFileName     
        #####Write-Host " Done."

        # Export unused field scripts
        #####Write-Host "Analysing dependencies to find unused fields..." -ForegroundColor Cyan
        Write-Host "Generating helper script to organise unused fields..." -ForegroundColor Cyan -NoNewline
        Export-UnusedFieldScript $TKS.dependenciesOutFolder $TKS.PbixFileName
        Write-Host " Done." -NoNewLine
        Write-Host

        # Export bim if required
        if ( ( $TKS.PbixProjSettings.settingsRead -eq "ValidSettings") -and ( $null -ne $TKS.PbixProjSettings.bimDeployType ) ) {
            Invoke-PbiToolsExportBim $TKS.pbixBIMDeployFolder $TKS.PbixProjSettings.bimDeployType
        }
    }
    Exit-ActionBIToolkit "`n`nAction BI Toolkit complete: Launching pbix repo in VS Code" $TKS.PbixRootFolder
}
    
if ($RunOption -eq "Export DAX Queries" -or $RunOption -eq "All") {
    Export-DaxQueries $TKS.exportDaxQueriesFolder $TKS.ConString
}

if ($RunOption -eq "Backup PBIX" ) {
    Copy-FileWithTimestamp $TKS.PbixFilePath $TKS.pbixQuickBackupFolder
}
if ($RunOption -eq "Report Page Tests" -or $RunOption -eq "All") {
    Test-ReportPages $TKS.performanceDataFolder $TKS.ConString
}

if ($RunOption -eq "Open VSCode" -or $RunOption -eq "All") {
    #Open and/or switch to in VSCode
    Exit-ActionBIToolkit "`n`nAction BI Toolkit complete: Launching pbix repo in VS Code" $TKS.PbixRootFolder;
}

if ($RunOption -eq "Compile Thin Report(s) .pbix from Source Control") {
    #Hot swap connection to running desktop instance
    Compile-ThinReportWithLocalConnection $TKS
    Exit-ActionBIToolkit "`n`nAction BI Toolkit complete: Compile Thin Reports from Source Control"
}

if ($RunOption -eq "Compile Thick Report .pbit from Source Control") {
    #Hot swap connection to running desktop instance
    Compile-ThickReport $TKS
    Exit-ActionBIToolkit "`n`nAction BI Toolkit complete: Compile Thick Report template .pbit from Source Control"
}

if ($RunOption -eq "Deploy to Power BI environment") {
    $DeploymentWorkspaceFolder, $DeploymentEnvironmentFolders = Get-DeploymentFolders $TKS
    $PowerBIWorkspaces = Initialize-DeploymentWorkspaces $DeploymentWorkspaceFolder $DeploymentEnvironmentFolders   
    switch ($PowerBIWorkspaces.count) {
        0 { Exit-ActionBIToolkit "No PowerBI Workspaces found...  " }
        1 { $ThisWorkSpace = $PowerBIWorkspaces[0] }
        default {
            Write-Host "`nPlease select one of the Power BI Workspace environments:"
            $i = 1
            foreach ($w in $PowerBIWorkspaces ) {
                Write-Host "$i) $($w.Name)"    
                $i++
            }
        }
    }
    $selection = Read-Host -Prompt "`nEnter the number of the environment to deploy to:"
    
    $ThisWorkspace = $PowerBIWorkspaces[$selection - 1]
    $ThisWorkspaceFolder = $DeploymentEnvironmentFolders[$selection - 1]

    if ($TKS.PbixFileType  -eq "Thick Report: Model embedded within Local PBIX" ) {
        Deploy-ThickReport $TKS $ThisWorkspace $ThisWorkspaceFolder
    }
    else {
        Deploy-ThickReport $TKS $ThisWorkspace $ThisWorkspaceFolder
    }

}


if ($null -ne $args[0]) {
    Exit-ActionBIToolkit "`n`nAction BI Toolkit operation complete."
}
