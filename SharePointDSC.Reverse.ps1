<##############################################################
 # This script is used to analyze an existing SharePoint (2013, 2016 or greater), and to produce the resulting PowerShell DSC Configuration Script representing it. Its purpose is to help SharePoint Admins and Devs replicate an existing SharePoint farm in an isolated area in order to troubleshoot an issue. This script needs to be executed directly on one of the SharePoint server in the far we wish to replicate. Upon finishing its execution, this Powershell script will prompt the user to specify a path to a FOLDER where the resulting PowerShell DSC Configuraton (.ps1) script will be generated. The resulting script will be named "SP-Farm.DSC.ps1" and will contain an exact description, in DSC notation, of the various components and configuration settings of the current SharePoint Farm. This script can then be used in an isolated environment to replicate the SharePoint server farm. The script could also be used as a simple textual (while in a DSC notation format) description of what the configuraton of the SharePoint farm looks like. This script is meant to be community driven, and everyone is encourage to participate and help improve and mature it. It is not officially endorsed by Microsoft, and support is 'offered' on a best effort basis by its contributors. Bugs suggestions should be reported through the issue system on GitHub. They will be looked at as time permits.
 # v1.0.0.37 - Nik Charlebois
 ##############################################################>

<## Script Settings #>
$VerbosePreference = "SilentlyContinue"

<## Scripts Variables #>
$Script:dscConfigContent = ""
$SPDSCSource = "C:\Program Files\WindowsPowerShell\Modules\SharePointDSC\"
$SPDSCVersion = "1.5.0.0"
$Script:SPDSCPath = $SPDSCSource + $SPDSCVersion
$Global:spFarmAccount = ""

<## This is the main function for this script. It acts as a call dispatcher, calling the various functions required in the proper order to get the full farm picture. #>
function Orchestrator
{
    Check-Prerequisites
    $ReverseDSCModule = "ReverseDSC.Core.psm1"
    $module = (Join-Path -Path $PSScriptRoot -ChildPath $ReverseDSCModule -Resolve -ErrorAction SilentlyContinue)
    if($module -eq $null)
    {
        $module = (Join-Path -Path $PSScriptRoot -ChildPath "..\ReverseDSC\$ReverseDSCModule" -Resolve)
    }
    
    Import-Module -Name $module -Force

    $Global:spFarmAccount = Get-Credential -Message "Credentials with Farm Admin Rights" -UserName $env:USERDOMAIN\$env:USERNAME
    Save-Credentials $Global:spFarmAccount

    $Global:spCentralAdmin = Get-SPWebApplication -IncludeCentralAdministration | Where{$_.DisplayName -like '*Central Administration*'}
    $spFarm = Get-SPFarm
    $spServers = $spFarm.Servers

    $totalSteps = 46 + $spServers.Count
    $currentStep = 1

    Write-Progress -Activity "Scanning Operating System Version..." -PercentComplete ($currentStep/$totalSteps*100)
    Read-OperatingSystemVersion
    $currentStep++

    Write-Progress -Activity "Scanning SQL Server Version..." -PercentComplete ($currentStep/$totalSteps*100)
    Read-SQLVersion
    $currentStep++

    Write-Progress -Activity "Scanning Patch Levels..." -PercentComplete ($currentStep/$totalSteps*100)
    Read-SPProductVersions
    $currentStep++

    $Script:dscConfigContent += "Configuration SharePointFarm`r`n"
    $Script:dscConfigContent += "{`r`n"

    Write-Progress -Activity "Configuring Credentials..." -PercentComplete ($currentStep/$totalSteps*100)
    Set-ObtainRequiredCredentials
    $currentStep++

    Write-Progress -Activity "Configuring Dependencies..." -PercentComplete ($currentStep/$totalSteps*100)
    Set-Imports
    $currentStep++

    Write-Progress -Activity "Configuring Variables..." -PercentComplete ($currentStep/$totalSteps*100)
    Set-VariableSection
    $currentStep++

    $serverNumber = 1
    foreach($spServer in $spServers)
    {
        <## SQL servers are returned by Get-SPServer but they have a Role of 'Invalid'. Therefore we need to ignore these. The resulting PowerShell DSC Configuration script does not take into account the configuration of the SQL server for the SharePoint Farm at this point in time. We are activaly working on giving our users an experience that is as painless as possible, and are planning on integrating the SQL DSC Configuration as part of our feature set. #>
        if($spServer.Role -ne "Invalid")
        {
            $Script:dscConfigContent += "`r`n    node " + $spServer.Name + "`r`n    {`r`n"
            
            <# If this is the first server in the farm, then generate the SPCreateFarm config section. Otherwise, generate the 
               SPJoinFarm one. #>
            if($serverNumber -eq 1)
            {
                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning the SharePoint Farm...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-SPFarm
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Web Application(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-SPWebApplications
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Alternate Url(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-SPAlternateUrl
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Managed Path(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-SPManagedPaths
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Managed Account(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-SPManagedAccounts
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Application Pool(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-SPServiceApplicationPools
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Content Database(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-SPContentDatabase
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Site Collection(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-SPSitesAndWebs
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Diagnostic Logging Settings...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-DiagnosticLoggingSettings
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Usage Service Application...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-UsageServiceApplication
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning State Service Application...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-StateServiceApplication
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning User Profile Service Application(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-UserProfileServiceapplication
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Cache Account(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-CacheAccounts
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Secure Store Service Application(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-SecureStoreServiceApplication
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Business Connectivity Service Application(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-BCSServiceApplication
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Search Service Application(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-SearchServiceApplication
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Managed Metadata Service Application(s)...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-ManagedMetadataServiceApplication
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Access Service Application(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPAccessServiceApp
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Antivirus Settings(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPAntivirusSettings
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning App Catalog Settings(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPAppCatalog
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning App Domain Settings(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPAppDomain
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning App Management Service App Settings(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPAppManagementServiceApp
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning App Store Settings(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPAppStoreSettings
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Blob Cache Settings(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPBlobCacheSettings
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Configuration Wizard Settings(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPConfigWizard
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Database(s) Availability Group Settings(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPDatabaseAAG
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Distributed Cache Settings(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPDistributedCacheService
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Excel Services Application Settings(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPExcelServiceApp
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Farm Administrator(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPFarmAdministrators
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Farm Solution(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPFarmSolution
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Health Rule(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPHealthAnalyzerRuleState
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning IRM Settings(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPIrmSettings
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Office Online Binding(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPOfficeOnlineServerBinding
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Crawl Rules(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPSearchCrawlRule
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Search File Type(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPSearchFileType
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Search Index Partition(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPSearchIndexPartition
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Search Result Source(s)...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPSearchResultSource
                $currentStep++

                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Search Topology...") -PercentComplete ($currentStep/$totalSteps*100)                
                Read-SPSearchTopology
                $currentStep++
            }
            else
            {
                Write-Progress -Activity ("[" + $spServer.Name + "] Scanning the SharePoint Farm...") -PercentComplete ($currentStep/$totalSteps*100)
                Read-SPJoinFarm
                $currentStep++
            }

            Write-Progress -Activity ("[" + $spServer.Name + "] Scanning Service Instance(s)...") -PercentComplete ($currentStep/$totalSteps*100)
            Read-SPServiceInstance -Server $spServer.Name
            $currentStep++

            Write-Progress -Activity ("[" + $spServer.Name + "] Configuring Local Configuration Manager (LCM)...") -PercentComplete ($currentStep/$totalSteps*100)
            Set-LCM
            $currentStep++

            $Script:dscConfigContent += "`r`n    }`r`n"
            $serverNumber++
        }
    }    
    $Script:dscConfigContent += "`r`n}`r`n"
    Write-Progress -Activity "[$spServer.Name] Setting Configuration Data..." -PercentComplete ($currentStep/$totalSteps*100)
    Set-ConfigurationData
    $currentStep++
    $Script:dscConfigContent += "SharePointFarm -ConfigurationData `$ConfigData"
}

function Check-Prerequisites
{
    <# Validate the PowerShell Version #>
    if($psVersionTable.PSVersion.Major -eq 4)
    {
        Write-Host "`r`nI100"  -BackgroundColor Cyan -ForegroundColor Black -NoNewline
        Write-Host "    - PowerShell v4 detected. While this script will work just fine with v4, it is highly recommended you upgrade to PowerShell v5 to get the most out of DSC"
    }
    elseif($psVersionTable.PSVersion.Major -lt 4)
    {
        Write-Host "`r`nE100"  -BackgroundColor Yellow -ForegroundColor Black -NoNewline
        Write-Host "    - We are sorry, PowerShell v3 or lower is not supported by the Reverse DSC Engine"
        exit
    }

    <# Check to see if the SharePointDSC module is installed on the machine #>
    if(Get-Command "Get-DSCModule" -EA SilentlyContinue)
    {
        $spDSCCheck = Get-DSCResource -Module "SharePointDSC" | ?{$_.Version -eq $SPDSCVersion}
        <# Because the SkipPublisherCheck parameter doesn't seem to be supported on Win2012R2 / PowerShell prior to 5.1, let's set whether the parameters are specified here. #>
        if (Get-Command -Name Install-Module -ParameterName SkipPublisherCheck -ErrorAction SilentlyContinue)
        {
            $skipPublisherCheckParameter = @{SkipPublisherCheck = $true}
        }
        else {$skipPublisherCheckParameter = @{}}
        if($spDSCCheck.Length -eq 0)
        {        
            $cmd = Get-Command Install-Module
            if($psVersionTable.PSVersion.Major -ge 5 -or $cmd)
            {
                $shouldInstall = Read-Host "The SharePointDSC module could not be found on the machine. Do you wish to download and install it (y/n)?"
                if($shouldInstall.ToLower() -eq "y")
                {
                    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
                    Install-Module SharePointDSC -RequiredVersion $SPDSCVersion -Confirm:$false @skipPublisherCheckParameter
                }
                else
                {
                    Write-Host "`r`nE101"  -BackgroundColor Yellow -ForegroundColor Black -NoNewline
                    Write-Host "   - We are sorry, but the script cannot continue without the SharePoint DSC module installed."
                    exit
                }
            }
            else
            {
                Write-Host "`r`nW101"  -BackgroundColor Yellow -ForegroundColor Black -NoNewline
                Write-Host "   - We could not find the PackageManagement modules on the machine. Please make sure you download and install it at https://www.microsoft.com/en-us/download/details.aspx?id=51451 before executing this script"
                $Script:SPDSCPath = $moduleObject[0].Module.Path.ToLower().Replace("sharepointdsc.psd1", "").Replace("\", "/")
            }
        }
    }
    else
    {
        <# PowerShell v4 is most likely present, without the PackageManagement module. We need to manually check to see if the SharePoint
           DSC Module is present on the machine. #>
        $cmd = Get-Command Install-Module -EA SilentlyContinue
        if(!$cmd)
        {
            Write-Host "`r`nW102"  -BackgroundColor Yellow -ForegroundColor Black -NoNewline
            Write-Host "   - We could not find the PackageManagement modules on the machine. Please make sure you download and install it at https://www.microsoft.com/en-us/download/details.aspx?id=51451 before executing this script"
        }
        $moduleObject = Get-DSCResource | ?{$_.Module -like "SharePointDsc"}
        if(!$moduleObject)
        {
            Write-Host "`r`nE103"  -BackgroundColor Red -ForegroundColor Black -NoNewline
            Write-Host "    - Could not find the SharePointDSC Module Resource on the current server."
            exit;
        }
        $Script:SPDSCPath = $moduleObject[0].Module.Path.ToLower().Replace("sharepointdsc.psd1", "").Replace("\", "/")
    }
}

function Read-OperatingSystemVersion
{
    $servers = Get-SPServer
    $Script:dscConfigContent += "<#`r`n    Operating Systems in this Farm`r`n-------------------------------------------`r`n"
    $Script:dscConfigContent += "    Products and Language Packs`r`n"
    $Script:dscConfigContent += "-------------------------------------------`r`n"
    foreach($spServer in $servers)
    {
        $serverName = $spServer.Name
        try{
            $osInfo = Get-CimInstance Win32_OperatingSystem  -ComputerName $serverName -ErrorAction SilentlyContinue| Select-Object @{Label="OSName"; Expression={$_.Name.Substring($_.Name.indexof("W"),$_.Name.indexof("|")-$_.Name.indexof("W"))}} , Version ,OSArchitecture -ErrorAction SilentlyContinue
            $Script:dscConfigContent += "    [" + $serverName + "]: " + $osInfo.OSName + "(" + $osInfo.OSArchitecture + ")    ----    " + $osInfo.Version + "`r`n"
        }
        catch{}
    }    
    $Script:dscConfigContent += "#>`r`n`r`n"
}

function Read-SQLVersion
{
    $uniqueServers = @()
    $sqlServers = Get-SPDatabase | select Server -Unique
    foreach($sqlServer in $sqlServers)
    {
        $serverName = $sqlServer.Server.Name

        if($serverName -eq $null)
        {
            $serverName = $sqlServer.Server
        }
        
        if(!($uniqueServers -contains $serverName))
        {
            $sqlVersionInfo = Invoke-SQL -Server $serverName -dbName "Master" -sqlQuery "SELECT @@VERSION AS 'SQLVersion'"
            $uniqueServers += $serverName.ToString()
            $Script:dscConfigContent += "<#`r`n    SQL Server Product Versions Installed on this Farm`r`n-------------------------------------------`r`n"
            $Script:dscConfigContent += "    Products and Language Packs`r`n"
            $Script:dscConfigContent += "-------------------------------------------`r`n"
            $Script:dscConfigContent += "    [" + $serverName + "]: " + $sqlVersionInfo.SQLversion + "`r`n#>`r`n`r`n"
        }
    }
}

function Set-VariableSection
{
    $Script:dscConfigContent += "    `$Script:passphrase = Read-Host `"Farm Passphrase`" -AsSecureString;`r`n"
}

function Set-ConfigurationData
{
    $Script:dscConfigContent += "`$ConfigData = @{`r`n"
    $Script:dscConfigContent += "    AllNodes = @(`r`n"

    $spFarm = Get-SPFarm
    $spServers = $spFarm.Servers

    $tempConfigDataContent = ""
    foreach($spServer in $spServers)
    {
        $tempConfigDataContent += "    @{`r`n"
        $tempConfigDataContent += "        NodeName = `"" + $spServer.Name + "`";`r`n"
        $tempConfigDataContent += "        PSDscAllowPlainTextPassword =`$true;`r`n"
        $tempConfigDataContent += "    },`r`n"
    }

    # Remove the last ',' in the array
    $tempConfigDataContent = $tempConfigDataContent.Remove($tempConfigDataContent.LastIndexOf(","), 1)
    $Script:dscConfigContent += $tempConfigDataContent
    $Script:dscConfigContent += ")}`r`n"
}

<## This function ensures all required DSC Modules are properly loaded into the current PowerShell session. #>
function Set-Imports
{
    $Script:dscConfigContent += "    Import-DscResource -ModuleName PSDesiredStateConfiguration`r`n"
    $Script:dscConfigContent += "    Import-DscResource -ModuleName SharePointDSC`r`n"
}

<## This function really is optional, but helps provide valuable information about the various software components installed in the current SharePoint farm (i.e. Cummulative Updates, Language Packs, etc.). #>
function Read-SPProductVersions
{    
    $Script:dscConfigContent += "<#`r`n    SharePoint Product Versions Installed on this Farm`r`n-------------------------------------------`r`n"
    $Script:dscConfigContent += "    Products and Language Packs`r`n"
    $Script:dscConfigContent += "-------------------------------------------`r`n"

    if($PSVersionTable.PSVersion -like "2.*")
    {
        $RegLoc = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall
        $Programs = $RegLoc | where-object { $_.PsPath -like "*\Office*" } | foreach {Get-ItemProperty $_.PsPath}        

        foreach($program in $Programs)
        {
            $Script:dscConfigContent += "    " +  $program.DisplayName + " -- " + $program.DisplayVersion + "`r`n"
        }
    }
    else
    {
        $regLoc = Get-ChildItem HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall
        $programs = $regLoc | where-object { $_.PsPath -like "*\Office*" } | foreach {Get-ItemProperty $_.PsPath} 
        $components = $regLoc | where-object { $_.PsPath -like "*1000-0000000FF1CE}" } | foreach {Get-ItemProperty $_.PsPath} 

        foreach($program in $programs)
        { 
            $productCodes = $_.ProductCodes
            $component = @() + ($components |     where-object { $_.PSChildName -in $productCodes } | foreach {Get-ItemProperty $_.PsPath})
            foreach($component in $components)
            {
                $Script:dscConfigContent += "    " + $component.DisplayName + " -- " + $component.DisplayVersion + "`r`n"
            }        
        }
    }
    $Script:dscConfigContent += "#>`r`n"
}

<## This function declares the xSPCreateFarm object required to create the config and admin database for the resulting SharePoint Farm. #>
function Read-SPFarm (){
    $spMajorVersion = (Get-SPDSCInstalledProductVersion).FileMajorPart
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPCreateFarm\MSFT_SPCreateFarm.psm1")
    Import-Module $module

    $Script:dscConfigContent += "        SPCreateFarm CreateSPFarm`r`n        {`r`n"
    $params = Get-DSCFakeParameters -ModulePath $module
    <# If not SP2016, remove the server role param. #>
    if ($spMajorVersion -ne 16) {
        $params.Remove("ServerRole")
    }

    <# Can't have both the InstallAccount and PsDscRunAsCredential variables present. Remove InstallAccount if both are there. #>
    if($params.Contains("InstallAccount"))
    {
        $params.Remove("InstallAccount")
    }
    $params.FarmAccount = $Global:spFarmAccount
    $params.Passphrase = $Global:spFarmAccount
    $results = Get-TargetResource @params
    <# Remove the default generated PassPhrase and ensure the resulting Configuration Script will prompt user for it. #>
    $results.Remove("Passphrase");
    $Script:dscConfigContent += "            Passphrase = New-Object System.Management.Automation.PSCredential ('Passphrase', `$passphrase);`r`n"
    if($spMajorVersion -ge 16)
    {
        $currentServer = Get-SPServer | ?{$_.Address -eq $env:COMPUTERNAME}
        $results.Add("ServerRole", $currentServer.Role)
    }
    $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
    $Script:dscConfigContent += "        }`r`n"

    <# SPFarm Feature Section #>
    $versionFilter = $spMajorVersion.ToString() + "*"
    $farmFeatures = Get-SPFeature | ?{$_.Scope -eq "Farm" -and $_.Version -like $versionFilter}
    $moduleFeature = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFeature\MSFT_SPFeature.psm1")
    Import-Module $moduleFeature
    $paramsFeature = Get-DSCFakeParameters -ModulePath $moduleFeature

    $featuresAlreadyAdded = @()
    foreach($farmFeature in $farmFeatures)
    {
        if(!$featuresAlreadyAdded.Contains($farmFeature.DisplayName))
        {
            $featuresAlreadyAdded += $farmFeature.DisplayName
            $paramsFeature.Name = $farmFeature.DisplayName
            $paramsFeature.FeatureScope = "Farm"
            $resultsFeature = Get-TargetResource @paramsFeature

            if($resultsFeature.Get_Item("Ensure").ToLower() -eq "present")
            {
                $Script:dscConfigContent += "        SPFeature " + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"

                <# Manually add the InstallAccount param due to a bug in 1.5.0.0 that returns a param named InstalAcount (typo) instead.
                   https://github.com/PowerShell/SharePointDsc/issues/481 #>
                if(!$resultsFeature.ContainsKey("InstallAccount"))
                {
                    $resultsFeature.Add("InstallAccount", "`$CredsFarmAccount")
                }
                if($resultsFeature.ContainsKey("InstalAcount"))
                {
                    $resultsFeature.Remove("InstalAcount")
                }
                #$resultsFeature["Version"] = $farmFeature.Version.

                $Script:dscConfigContent += Get-DSCBlock -Params $resultsFeature -ModulePath $moduleFeature
                $Script:dscConfigContent += "        }`r`n"
            }
            }
    }
}

<## This function declares the xSPCreateFarm object required to create the config and admin database for the resulting SharePoint Farm. #>
function Read-SPJoinFarm (){
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPJoinFarm\MSFT_SPJoinFarm.psm1")
    Import-Module $module
        
    $Script:dscConfigContent += "        SPJoinFarm JoinFarm`r`n        {`r`n"
    $params = Get-DSCFakeParameters -ModulePath $module

    <# If not SP2016, remove the server role param. #>
    if ((Get-SPDSCInstalledProductVersion).FileMajorPart -ne 16 -and $params.ContainsKey("ServerRole")) {
        $params.Remove("ServerRole")
    }

    <# Can't have both the InstallAccount and PsDscRunAsCredential variables present. Remove InstallAccount if both are there. #>
    if($params.ContainsKey("InstallAccount"))
    {
        $params.Remove("InstallAccount")
    }

    $results = Get-TargetResource @params

    <# Remove the default generated PassPhrase and ensure the resulting Configuration Script will prompt user for it. #>
    $results.Remove("Passphrase");    
    $Script:dscConfigContent += "            Passphrase = New-Object System.Management.Automation.PSCredential ('Passphrase', `$passphrase);`r`n"

    $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
    $Script:dscConfigContent += "`r`n        }`r`n"
}


<## This function obtains a reference to every Web Application in the farm and declares their properties (i.e. Port, Associated IIS Application Pool, etc.). #>
function Read-SPWebApplications (){
    Write-Verbose "Reading Information about all Web Applications..."
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWebApplication\MSFT_SPWebApplication.psm1")
    Import-Module $module
    $spWebApplications = Get-SPWebApplication | Sort-Object -Property Name
    $params = Get-DSCFakeParameters -ModulePath $module
    
    foreach($spWebApp in $spWebApplications)
    {
        Import-Module $module
        $Script:dscConfigContent += "        SPWebApplication " + $spWebApp.Name.Replace(" ", "") + "`r`n        {`r`n"      

        $params.Name = $spWebApp.Name
        $results = Get-TargetResource @params

        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "`r`n        }`r`n"
        Read-SPDesignerSettings($spWebApplications.Url.ToString(), "WebApplication", $spWebApp.Name.Replace(" ", ""))

        <# SPWebApplication Feature Section #>
        $spMajorVersion = (Get-SPDSCInstalledProductVersion).FileMajorPart
        $versionFilter = $spMajorVersion.ToString() + "*"
        $webAppFeatures = Get-SPFeature | ?{$_.Scope -eq "WebApplication" -and $_.Version -like $versionFilter}
        $moduleFeature = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFeature\MSFT_SPFeature.psm1")
        Import-Module $moduleFeature
        $paramsFeature = Get-DSCFakeParameters -ModulePath $moduleFeature
        foreach($webAppFeature in $webAppFeatures)
        {
            $paramsFeature.Name = $webAppFeature.DisplayName
            $paramsFeature.FeatureScope = "WebApplication"
            $paramsFeature.Url = $spWebApp.Url
            $resultsFeature = Get-TargetResource @paramsFeature

            if($resultsFeature.Get_Item("Ensure").ToLower() -eq "present")
            {
                $Script:dscConfigContent += "        SPFeature " + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
            
                <# Manually add the InstallAccount param due to a bug in 1.5.0.0 that returns a param named InstalAcount (typo) instead.
                   https://github.com/PowerShell/SharePointDsc/issues/481 #>
                if(!$resultsFeature.ContainsKey("InstallAccount"))
                {
                    $resultsFeature.Add("InstallAccount", "`$CredsFarmAccount")
                }
                if($resultsFeature.ContainsKey("InstalAcount"))
                {
                    $resultsFeature.Remove("InstalAcount")
                }
                $Script:dscConfigContent += Get-DSCBlock -Params $resultsFeature -ModulePath $moduleFeature
                $Script:dscConfigContent += "            DependsOn = `"[SPWebApplication]" + $spWebApp.Name.Replace(" ", "") + "`";`r`n"
                $Script:dscConfigContent += "        }`r`n"
            }
        }

        <# Outgoing Email Setting Region #>
        $moduleEmail = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPOutgoingEmailSettings\MSFT_SPOutgoingEmailSettings.psm1")
        Import-Module $moduleEmail
        $paramsEmail = Get-DSCFakeParameters -ModulePath $moduleEmail

        $paramsEmail.WebAppUrl = $spWebApp.Url
        $resultsEmail = Get-TargetResource @paramsEmail
        if($null -ne $resultsEmail["SMTPServer"] -and "" -ne $resultsEmail["SMTPServer"])
        {
            if(!$resultsEmail.ContainsKey("InstallAccount"))
            {
                $resultsEmail.Add("InstallAccount", "`$CredsFarmAccount")
            }
            $Script:dscConfigContent += "        SPOutgoingEmailSettings " + [System.Guid]::NewGuid().ToString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $Script:dscConfigContent += Get-DSCBlock -Params $resultsEmail -ModulePath $moduleEmail
            $Script:dscConfigContent += "            DependsOn = `"[SPWebApplication]" + $spWebApp.Name.Replace(" ", "") + "`";`r`n"
            $Script:dscConfigContent += "        }`r`n"
        }
    }
}

<## This function loops through every IIS Application Pool that are associated with the various existing Service Applications in the SharePoint farm. ##>
function Read-SPServiceApplicationPools ($modulePath, $params){
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPServiceAppPool\MSFT_SPServiceAppPool.psm1")
        Import-Module $module
    }
    
    $spServiceAppPools = Get-SPServiceApplicationPool | Sort-Object -Property Name

    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    foreach($spServiceAppPool in $spServiceAppPools)
    {
        $Script:dscConfigContent += "        SPServiceAppPool " + $spServiceAppPool.Name.Replace(" ", "") + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Name = $spServiceAppPool.Name
        $results = Get-TargetResource @params    
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

<## This function retrieves a list of all site collections, no matter what Web Application they belong to. The Url attribute helps the xSharePoint DSC Resource determine what Web Application they belong to. #>
function Read-SPSitesAndWebs (){
    
    $spSites = Get-SPSite -Limit All
    $siteGuid = $null
    $siteTitle = $null
    foreach($spsite in $spSites)
    {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSite\MSFT_SPSite.psm1")
        Import-Module $module
        $params = Get-DSCFakeParameters -ModulePath $module
        $siteGuid = [System.Guid]::NewGuid().toString()
        $siteTitle = $spSite.RootWeb.Title
        if($siteTitle -eq $null)
        {
            $siteTitle = "SiteCollection"
        }
        $Script:dscConfigContent += "        SPSite " + $siteGuid + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Url = $spSite.Url
        $results = Get-TargetResource @params

        <# If the current Quota ID is 0, it means no quota templates were used. Remove param in that case. #>
        if($spSite.Quota.QuotaID -eq 0)
        {
            $results.Remove("QuotaTemplate")
        }
        if($null -eq $results.Get_Item("SecondaryOwnerAlias"))
        {
            $results.Remove("SecondaryOwnerAlias")
        }
        if($null -eq $results.Get_Item("SecondaryEmail"))
        {
            $results.Remove("SecondaryEmail")
        }
        if($null -eq $results.Get_Item("OwnerEmail") -or "" -eq $results.Get_Item("OwnerEmail"))
        {
            $results.Remove("OwnerEmail")
        }
        if($null -eq $results.Get_Item("HostHeaderWebApplication"))
        {
            $results.Remove("HostHeaderWebApplication")
        }
        if($null -eq $results.Get_Item("Name") -or "" -eq $results.Get_Item("Name"))
        {
            $results.Remove("Name")
        }
        if($null -eq $results.Get_Item("Description") -or "" -eq $results.Get_Item("Description"))
        {
            $results.Remove("Description")
        }
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "            DependsOn =  `"[SPWebApplication]" + $spSite.WebApplication.Name.Replace(" ", "") + "`"`r`n"
        $Script:dscConfigContent += "        }`r`n"

        <# Nik20170112 - There are restrictions preventing this setting from being applied if the PsDscRunAsCredential parameter is not used.
                         Since this is only available in WMF 5, we check to see if the node farm we are extracting the configuration from is
                         running at least PowerShell v5 before reading the Site Collection level SPDesigner settings. #>
        if($PSVersionTable.PSVersion.Major -ge 5)
        {
            Read-SPDesignerSettings($spSite.Url, "SiteCollection")
        }
        
        $webs = Get-SPWeb -Limit All -Site $spsite
        foreach($spweb in $webs)
        {
            $moduleWeb = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPWeb\MSFT_SPWeb.psm1")
            Import-Module $moduleWeb
            $paramsWeb = Get-DSCFakeParameters -ModulePath $moduleWeb
            $paramsWeb.Url = $spweb.Url            
            $resultsWeb = Get-TargetResource @paramsWeb
            if(!$resultsWeb.ContainsKey("InstallAccount"))
            {
                $resultsWeb.Add("InstallAccount", "`$CredsFarmAccount")
            }
            if($resultsWeb.ContainsKey("PsDscRunAsCredential"))
            {
                $resultsWeb.Remove("PsDscRunAsCredential")
            }
            $Script:dscConfigContent += "        SPWeb " + [System.Guid]::NewGuid().toString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $Script:dscConfigContent += Get-DSCBlock -Params $resultsWeb -ModulePath $moduleWeb
            $Script:dscConfigContent += "            DependsOn = `"[SPSite]" + $siteGuid + "`";`r`n"
            $Script:dscConfigContent += "        }`r`n"

            <# SPWeb Feature Section #>
            $spMajorVersion = (Get-SPDSCInstalledProductVersion).FileMajorPart
            $versionFilter = $spMajorVersion.ToString() + "*"
            $webFeatures = Get-SPFeature | ?{$_.Scope -eq "Web" -and $_.Version -like $versionFilter}
            $moduleFeature = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFeature\MSFT_SPFeature.psm1")
            Import-Module $moduleFeature
            $paramsFeature = Get-DSCFakeParameters -ModulePath $moduleFeature

            foreach($webFeature in $webFeatures)
            {
                $paramsFeature.Name = $webFeature.DisplayName
                $paramsFeature.FeatureScope = "Web"
                $paramsFeature.Url = $spWeb.Url
                $resultsFeature = Get-TargetResource @paramsFeature

                if($resultsFeature.Get_Item("Ensure").ToLower() -eq "present")
                {
                    $Script:dscConfigContent += "        SPFeature " + [System.Guid]::NewGuid().ToString() + "`r`n"
                    $Script:dscConfigContent += "        {`r`n"
                
                    <# Manually add the InstallAccount param due to a bug in 1.5.0.0 that returns a param named InstalAcount (typo) instead.
                       https://github.com/PowerShell/SharePointDsc/issues/481 #>
                    if(!$resultsFeature.ContainsKey("InstallAccount"))
                    {
                        $resultsFeature.Add("InstallAccount", "`$CredsFarmAccount")
                    }
                    if($resultsFeature.ContainsKey("InstalAcount"))
                    {
                        $resultsFeature.Remove("InstalAcount")
                    }

                    $Script:dscConfigContent += Get-DSCBlock -Params $resultsFeature -ModulePath $moduleFeature
                    $Script:dscConfigContent += "            DependsOn = `"[SPSite]" + $siteGuid + "`";`r`n"
                    $Script:dscConfigContent += "        }`r`n"
                }
            }
        }

        <# SPSite Feature Section #>
        $spMajorVersion = (Get-SPDSCInstalledProductVersion).FileMajorPart
        $versionFilter = $spMajorVersion.ToString() + "*"
        $siteFeatures = Get-SPFeature | ?{$_.Scope -eq "Site" -and $_.Version -like $versionFilter}
        $moduleFeature = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFeature\MSFT_SPFeature.psm1")
        Import-Module $moduleFeature
        $paramsFeature = Get-DSCFakeParameters -ModulePath $moduleFeature
        foreach($siteFeature in $siteFeatures)
        {
            $paramsFeature.Name = $siteFeature.DisplayName
            $paramsFeature.FeatureScope = "Site"
            $paramsFeature.Url = $spSite.Url
            $resultsFeature = Get-TargetResource @paramsFeature

            if($resultsFeature.Get_Item("Ensure").ToLower() -eq "present")
            {
                $Script:dscConfigContent += "        SPFeature " + $siteTitle.Replace(" ", "") + "-" + $siteFeature.DisplayName + "-" + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
            
                <# Manually add the InstallAccount param due to a bug in 1.5.0.0 that returns a param named InstalAcount (typo) instead.
                   https://github.com/PowerShell/SharePointDsc/issues/481 #>
                if(!$resultsFeature.ContainsKey("InstallAccount"))
                {
                    $resultsFeature.Add("InstallAccount", "`$CredsFarmAccount")
                }
                if($resultsFeature.ContainsKey("InstalAcount"))
                {
                    $resultsFeature.Remove("InstalAcount")
                }
                $Script:dscConfigContent += Get-DSCBlock -Params $resultsFeature -ModulePath $moduleFeature
                $Script:dscConfigContent += "            DependsOn = `"[SPSite]" + $siteGuid + "`";`r`n"
                $Script:dscConfigContent += "        }`r`n"
            }
        }
    }
}

<## This function generates a list of all Managed Paths, no matter what their associated Web Application is. The xSharePoint DSC Resource uses the WebAppUrl attribute to identify what Web Applicaton they belong to. #>
function Read-SPManagedPaths ($modulePath, $params){
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPManagedPath\MSFT_SPManagedPath.psm1")
        Import-Module $module
    }   

    $spWebApps = Get-SPWebApplication

    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }
    foreach($spWebApp in $spWebApps)
    {
        $spManagedPaths = Get-SPManagedPath -WebApplication $spWebApp.Url | Sort-Object -Property Name

        foreach($spManagedPath in $spManagedPaths)
        {
            if($spManagedPath.Name.Length -gt 0 -and $spManagedPath.Name -ne "sites")
            {
                $Script:dscConfigContent += "        SPManagedPath " + $spWebApp.Name.Replace(" ", "") + "Path" + $spManagedPath.Name + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                if($spManagedPath.Name -ne $null)
                {
                    $params.RelativeUrl = $spManagedPath.Name
                }                
                $params.WebAppUrl = $spWebApp.Url
                $params.HostHeader = $false;
                $results = Get-TargetResource @params    
                $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }            
        }

        $spManagedPaths = Get-SPManagedPath -HostHeader | Sort-Object -Property Name
        foreach($spManagedPath in $spManagedPaths)
        {
            if($spManagedPath.Name.Length -gt 0 -and $spManagedPath.Name -ne "sites")
            {
                $Script:dscConfigContent += "        SPManagedPath " + $spWebApp.Name.Replace(" ", "") + "Path" + $spManagedPath.Name + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                
                if($spManagedPath.Name -ne $null)
                {
                    $params.RelativeUrl = $spManagedPath.Name
                } 
                if($params.ContainsKey("Explicit"))
                {
                    $params.Explicit = ($spManagedPath.Type -eq "ExplicitInclusion")
                }
                else
                {
                    $params.Add("Explicit", ($spManagedPath.Type -eq "ExplicitInclusion"))
                }
                $params.HostHeader = $true;
                $params.WebAppUrl = $spWebApp.Url
                $results = Get-TargetResource @params
                $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }            
        }
    }
}

<## This function retrieves all Managed Accounts in the SharePoint Farm. The Account attribute sets the associated credential variable (each managed account is declared as a variable and the user is prompted to Manually enter the credentials when first executing the script. See function "Set-ObtainRequiredCredentials" for more details on how these variales are set. #>
function Read-SPManagedAccounts (){
    $managedAccounts = Get-SPManagedAccount

    foreach($managedAccount in $managedAccounts)
    {
        $managedCreds = Get-Credentials $managedAccount.UserName
        $params = @{Account = $managedCreds; AccountName = $managedAccount.UserName}
        $exportContent = Export-TargetResource -ResourceName "MSFT_SPManagedAccount" -MandatoryParameters $params
        $Script:dscConfigContent += $exportContent
    }
}

<## This function retrieves all Services in the SharePoint farm. It does not care if the service is enabled or not. It lists them all, and simply sets the "Ensure" attribute of those that are disabled to "Absent". #>
function Read-SPServiceInstance ($Server){

    $serviceInstancesOnCurrentServer = Get-SPServiceInstance | Where{$_.Server.Name -eq $Server} | Sort-Object -Property TypeName

    $ensureValue = "Present"
    foreach($serviceInstance in $serviceInstancesOnCurrentServer)
    {
        if($serviceInstance.Status -eq "Online")
        {
            $ensureValue = "Present"
        }
        else
        {
            $ensureValue = "Absent"
        }
        Write-Verbose $serviceInstance.TypeName
        if($serviceInstance.TypeName -eq "Distributed Cache")
        {
            $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPDistributedCacheService\MSFT_SPDistributedCacheService.psm1")
            Import-Module $module
            $params = Get-DSCFakeParameters -ModulePath $module
            $params.Ensure = $ensureValue
            if($params.ServerProvisionOrder -eq $null)
            {
                $params.ServerProvisionOrder = "@()"
            }
            $results = Get-TargetResource @params
            $Script:dscConfigContent += "        SPDistributedCacheService " + $serviceInstance.TypeName.Replace(" ", "") + "Instance`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
        elseif($serviceInstance.TypeName -eq "User Profile Synchronization Service")
        {
            $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPUserProfileSyncService\MSFT_SPUserProfileSyncService.psm1")
            Import-Module $module
            $params = Get-DSCFakeParameters -ModulePath $module
            $params.Ensure = $ensureValue
            $params.FarmAccount = $Global:spFarmAccount            
            $results = Get-TargetResource @params
            if($ensureValue -eq "Present")
            {            
                $Script:dscConfigContent += "        SPUserProfileSyncService " + $serviceInstance.TypeName.Replace(" ", "") + "Instance`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
        else
        {
            $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPServiceInstance\MSFT_SPServiceInstance.psm1")
            Import-Module $module
            $params = Get-DSCFakeParameters -ModulePath $module
            if($params.ContainsKey("Name"))
            {
                $params.Name = $serviceInstance.TypeName
            }
            if($params.ContainsKey("FarmAccount"))
            {
                $params.FarmAccount = $Global:spFarmAccount
            }
            $params.Ensure = $ensureValue
            $results = Get-TargetResource @params
            $Script:dscConfigContent += "        SPServiceInstance " + $serviceInstance.TypeName.Replace(" ", "") + "Instance`r`n"
            $Script:dscConfigContent += "        {`r`n"            
            $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
    }
}

<## This function retrieves all settings related to Diagnostic Logging (ULS logs) on the SharePoint farm. #>
function Read-DiagnosticLoggingSettings ($modulePath, $params){
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPDiagnosticLoggingSettings\MSFT_SPDiagnosticLoggingSettings.psm1")
        Import-Module $module
    }
   
    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }
    $diagConfig = Get-SPDiagnosticConfig    

    $Script:dscConfigContent += "        SPDiagnosticLoggingSettings ApplyDiagnosticLogSettings`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $results = Get-TargetResource @params
    $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
    $Script:dscConfigContent += "        }`r`n"
}

<## This function retrieves all settings related to the SharePoint Usage Service Application, assuming it exists. #>
function Read-UsageServiceApplication ($modulePath, $params){
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPUsageApplication\MSFT_SPUsageApplication.psm1")
        Import-Module $module
    }
    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $usageApplication = Get-SPUsageApplication
    if($usageApplication.Length -gt 0)
    {
        $Script:dscConfigContent += "        SPUsageApplication " + $usageApplication.TypeName.Replace(" ", "") + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Get-TargetResource @params
        $results.Add("InstallAccount", "`$CredsFarmAccount")
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

<## This function retrieves settings associated with the State Service Application, assuming it exists. #>
function Read-StateServiceApplication ($modulePath, $params){
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPStateServiceApp\MSFT_SPStateServiceApp.psm1")
        Import-Module $module
    }
    
    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $stateApplications = Get-SPStateServiceApplication
    foreach($stateApp in $stateApplications)
    {
        if($stateApp -ne $null)
        {
            $params.Name = $stateApp.DisplayName
            $Script:dscConfigContent += "        SPStateServiceApp " + $stateApp.DisplayName.Replace(" ", "") + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $results = Get-TargetResource @params
            $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
    }
}

<## This function retrieves information about all the "Super" accounts (Super Reader & Super User) used for caching. #>
function Read-CacheAccounts ($modulePath, $params){
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPCacheAccounts\MSFT_SPCacheAccounts.psm1")
        Import-Module $module
    }
    
    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $webApps = Get-SPWebApplication

    foreach($webApp in $webApps)
    {
        $params.WebAppUrl = $webApp.Url
        $results = Get-TargetResource @params

        $accountsMissing = 0
        if($results.SuperReaderAlias -ne "" -and $results.SuperUserAlias -ne "")
        {
            $Script:dscConfigContent += "        SPCacheAccounts " + $webApp.DisplayName.Replace(" ", "") + "CacheAccounts`r`n"
            $Script:dscConfigContent += "        {`r`n"        
            $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
    }
}

<## This function retrieves settings related to the User Profile Service Application. #>
function Read-UserProfileServiceapplication ($modulePath, $params){
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPUserProfileServiceApp\MSFT_SPUserProfileServiceApp.psm1")
        Import-Module $module
    }
    
    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $ups = Get-SPServiceApplication | Where{$_.TypeName -eq "User Profile Service Application"}

    $sites = Get-SPSite
    if($sites.Length -gt 0)
    {
        $context = Get-SPServiceContext $sites[0]
        try
        {
            $pm = new-object Microsoft.Office.Server.UserProfiles.UserProfileManager($context)
        }
        catch{
                Write-Host "`r`nW102"  -BackgroundColor Yellow -ForegroundColor Black -NoNewline
                Write-Host "   - Farm Account does not have Full Control on the User Profile Service Application."
        }

        if($ups -ne $null)
        {
            $params.Name = $ups.DisplayName
            $Script:dscConfigContent += "        SPUserProfileServiceApp UserProfileServiceApp`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $results = Get-TargetResource @params
            $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
    }
}

<## This function retrieves all settings related to the Secure Store Service Application. Currently this function makes a direct call to the Secure Store database on the farm's SQL server to retrieve information about the logging details. There are currently no publicly available hooks in the SharePoint/Office Server Object Model that allow us to do it. This forces the user executing this reverse DSC script to have to install the SQL Server Client components on the server on which they execute the script, which is not a "best practice". #>
<# TODO: Change the logic to extract information about the logging from being a direct SQL call to something that uses the Object Model. #>
function Read-SecureStoreServiceApplication ($modulePath, $params){
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSecureStoreServiceApp\MSFT_SPSecureStoreServiceApp.psm1")
        Import-Module $module
    }
    
    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $ssa = Get-SPServiceApplication | Where{$_.TypeName -eq "Secure Store Service Application"}
    for($i = 0; $i -lt $ssa.Length; $i++)
    {
        $params.Name = $ssa.DisplayName
        $Script:dscConfigContent += "        SPSecureStoreServiceApp " + $ssa[$i].Name.Replace(" ", "") + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Get-TargetResource @params

        # HACK: Can't dynamically retrieve value from the Secure Store at the moment #>
        $results.Add("AuditingEnabled", $true)

        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"        
    }
}

<## This function retrieves settings related to the Managed Metadata Service Application. #>
function Read-ManagedMetadataServiceApplication ($modulePath, $params){
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPManagedMetadataServiceApp\MSFT_SPManagedMetadataServiceApp.psm1")
        Import-Module $module
    }
    
    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $mms = Get-SPServiceApplication | Where{$_.TypeName -eq "Managed Metadata Service"}
    if (Get-Command "Get-SPMetadataServiceApplication" -errorAction SilentlyContinue)
    {
        foreach($mmsInstance in $mms)
        {
            if($mmsInstance -ne $null)
            {
                $params.Name = $mmsInstance.Name
                $Script:dscConfigContent += "        SPManagedMetaDataServiceApp " + $mmsInstance.Name.Replace(" ", "") + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $results = Get-TargetResource @params
                $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
    }
}

<## This function retrieves settings related to the Business Connectivity Service Application. #>
function Read-BCSServiceApplication ($modulePath, $params){
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPBCSServiceApp\MSFT_SPBCSServiceApp.psm1")
        Import-Module $module
    }
    
    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $bcsa = Get-SPServiceApplication | Where{$_.TypeName -eq "Business Data Connectivity Service Application"}
    
    foreach($bcsaInstance in $bcsa)
    {
        if($bcsaInstance -ne $null)
        {
            $Script:dscConfigContent += "        SPBCSServiceApp " + $bcsaInstance.Name.Replace(" ", "") + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.Name = $bcsa.DisplayName
            $results = Get-TargetResource @params
            $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"        
        }
    }
}

<## This function retrieves settings related to the Search Service Application. #>
function Read-SearchServiceApplication ($modulePath, $params){
    if($modulePath -ne $null)
    {
        $module = Resolve-Path $modulePath
    }
    else {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchServiceApp\MSFT_SPSearchServiceApp.psm1")
        Import-Module $module
    }
    
    if($params -eq $null)
    {
        $params = Get-DSCFakeParameters -ModulePath $module
    }

    $searchSA = Get-SPServiceApplication | Where{$_.TypeName -eq "Search Service Application"}
    
    foreach($searchSAInstance in $searchSA)
    {
        if($searchSAInstance -ne $null)
        {
            $Script:dscConfigContent += "        SPSearchServiceApp " + $searchSAInstance.Name.Replace(" ", "") + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.Name = $searchSAInstance.Name
            $results = Get-TargetResource @params
            if($results.Get_Item("CloudIndex") -eq $false)
            {
                $results.Remove("CloudIndex")
            }
            <# Nik20170111 - Fix a bug in 1.5.0.0 where DatabaseName and DatabaseServer is not properly returned #>
            $results["DatabaseName"] = $searchSAInstance.SearchAdminDatabase.Name
            $results["DatabaseServer"] = $searchSAInstance.SearchAdminDatabase.Server.Name
            $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"

            #region Search Content Sources
            $moduleContentSource = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchContentSource\MSFT_SPSearchContentSource.psm1")
            Import-Module $moduleContentSource
            $paramsContentSource = Get-DSCFakeParameters -ModulePath $moduleContentSource
            $contentSources = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $searchSAInstance.Name

            foreach($contentSource in $contentSources)
            {
                $sscsGuid = [System.Guid]::NewGuid().toString()
                $Script:dscConfigContent += "        SPSearchContentSource " + $contentSource.Name.Replace(" ", "") + $sscsGuid + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $paramsContentSource.Name = $contentSource.Name
                $paramsContentSource.ServiceAppName  = $searchSAInstance.Name
                $resultsContentSource = Get-TargetResource @paramsContentSource
                

                $searchScheduleModulePath = Resolve-Path ($Script:SPDSCPath + "\Modules\SharePointDsc.Search\SPSearchContentSource.Schedules.psm1")            
                Import-Module -Name $searchScheduleModulePath
                # TODO: Figure out way to properly pass CimInstance objects and then add the schedules back;
                $incremental = Get-SPDSCSearchCrawlSchedule -Schedule $contentSource.IncrementalCrawlSchedule
                $full = Get-SPDSCSearchCrawlSchedule -Schedule $contentSource.FullCrawlSchedule

                
                $resultsContentSource.IncrementalSchedule = Get-SPCrawlSchedule $incremental
                $resultsContentSource.FullSchedule = Get-SPCrawlSchedule $full
                $Script:dscConfigContent += Get-DSCBlock -Params $resultsContentSource -ModulePath $moduleContentSource
                $Script:dscConfigContent += "        }`r`n"
            }
            #endregion
        }
        else {
            $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchServiceApp\MSFT_SPSearchServiceApp.psm1")
            Import-Module $module
        }        
    }
}

function Get-SPCrawlSchedule($params)
{
    $currentSchedule = "MSFT_SPSearchCrawlSchedule{`r`n"
    foreach($key in $params.Keys)
    {
        $currentSchedule += "                " + $key + " = `"" + $params[$key] + "`"`r`n"
    }
    $currentSchedule += "            }"
    return $currentSchedule
}

function Read-SPContentDatabase
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPContentDatabase\MSFT_SPContentDatabase.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $spContentDBs = Get-SPContentDatabase

    foreach($spContentDB in $spContentDBs)
    {
        $Script:dscConfigContent += "        SPContentDatabase " + $spContentDB.Name.Replace(" ", "") + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Name = $spContentDB.Name
        $params.WebAppUrl = $spContentDB.WebApplication.Url
        $results = Get-TargetResource @params
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"  
    }
}

function Read-SPAccessServiceApp
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAccessServiceApp\MSFT_SPAccessServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $serviceApps = Get-SPServiceApplication
    $serviceApps = $serviceApps | Where-Object -FilterScript { 
            $_.GetType().FullName -eq "Microsoft.Office.Access.Services.MossHost.AccessServicesWebServiceApplication"}

    foreach($spAccessService in $serviceApps)
    {
        $Script:dscConfigContent += "        SPAccessServiceApp " + $spAccessService.Name.Replace(" ", "") + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Name = $spAccessService.Name
        $dbServer = $spAccessService.GetDatabaseServers(1).ServerName
        $params.DatabaseServer = $dbServer
        $results = Get-TargetResource @params
        if(!$results.ContainsKey("InstallAccount"))
        {
            $results.Add("InstallAccount", "`$CredsFarmAccount")
        }
        if(!$results.ContainsKey("DatabaseServer"))
        {
            $results.Add("DatabaseServer", $dbServer)
        }
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"  
    }
}

function Read-SPAppCatalog
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAppCatalog\MSFT_SPAppCatalog.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $webApps = Get-SPWebApplication

    foreach($webApp in $webApps)
    {
        $feature = $webApp.Features.Item([Guid]::Parse("f8bea737-255e-4758-ab82-e34bb46f5828"))
        if($null -ne $feature)
        {
            $appCatalogSiteId = $feature.Properties["__AppCatSiteId"].Value
            $appCatalogSite = $webApp.Sites | ?{$_.ID -eq $appCatalogSiteId}

            if($null -ne $appCatalogSite)
            {
                $Script:dscConfigContent += "        SPAppCatalog " + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $params.SiteUrl = $appCatalogSite.Url
                $results = Get-TargetResource @params
                $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
    }
}

function Read-SPAppDomain
{
    $serviceApp = Get-SPServiceApplication | ?{$_.TypeName -eq "Microsoft SharePoint Foundation Subscription Settings Service Application"}
    $appDomain =  Get-SPAppDomain
    if($serviceApp.Length -ge 1 -and $appDomain.Length -ge 1)
    {
        $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAppDomain\MSFT_SPAppDomain.psm1")
        Import-Module $module
        $params = Get-DSCFakeParameters -ModulePath $module
        $Script:dscConfigContent += "        SPAppDomain " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Get-TargetResource @params
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPSearchFileType
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchFileType\MSFT_SPSearchFileType.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssa = Get-SPServiceApplication | Where-Object -FilterScript { 
            $_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication" 
    }

    if($null -ne $ssa)
    {
        $fileFormats = Get-SPEnterpriseSearchFileFormat -SearchApplication $ssa

        foreach($fileFormat in $fileFormats)
        {
          $Script:dscConfigContent += "        SPSearchFileType " + [System.Guid]::NewGuid().ToString() + "`r`n"
          $Script:dscConfigContent += "        {`r`n"
          $params.ServiceAppName = $ssa.DisplayName
          $params.FileType = $fileFormat.Identity
          $results = Get-TargetResource @params
          $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
          $Script:dscConfigContent += "        }`r`n"
       }
    }
}

function Read-SPSearchIndexPartition
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchIndexPartition\MSFT_SPSearchIndexPartition.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssa = Get-SPServiceApplication | Where-Object -FilterScript { 
            $_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication" 
    }
    if($null -ne $ssa)
    {
        $ssa = Get-SPEnterpriseSearchServiceApplication -Identity $ssa
        $currentTopology = $ssa.ActiveTopology
        $indexComponents = Get-SPEnterpriseSearchComponent -SearchTopology $currentTopology | `
                                    Where-Object -FilterScript { 
                                        $_.GetType().Name -eq "IndexComponent" 
                                    }

        foreach($indexComponent in $indexComponents)
        {
            $Script:dscConfigContent += "        SPSearchIndexPartition " + [System.Guid]::NewGuid().ToString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.ServiceAppName = $ssa.DisplayName
            $params.Index = $indexComponent.IndexPartitionOrdinal
            $params.Servers = $indexComponent.ServerName
            $results = Get-TargetResource @params
            $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
    }
}

function Read-SPSearchTopology
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchTopology\MSFT_SPSearchTopology.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssa = Get-SPServiceApplication | Where-Object -FilterScript { 
            $_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication" 
    }

    if($null -ne $ssa)
    {
        $Script:dscConfigContent += "        SPSearchTopology " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.ServiceAppName = $ssa.DisplayName
        $results = Get-TargetResource @params
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPSearchResultSource
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchResultSource\MSFT_SPSearchResultSource.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssa = Get-SPServiceApplication | Where-Object -FilterScript { 
            $_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication" 
    }
    if($null -ne $ssa)
    {
        $ssa = Get-SPEnterpriseSearchServiceApplication -Identity $ssa
        $searchSiteUrl = $ssa.SearchCenterUrl -replace "/pages"
        $searchSite = Get-SPWeb -Identity $searchSiteUrl -ErrorAction SilentlyContinue

        if(!$null -eq $searchSite)
        {
            $adminNamespace = "Microsoft.Office.Server.Search.Administration"
            $queryNamespace = "Microsoft.Office.Server.Search.Administration.Query"
            $objectLevel = [Microsoft.Office.Server.Search.Administration.SearchObjectLevel]
            $fedManager = New-Object -TypeName "$queryNamespace.FederationManager" `
                                     -ArgumentList $ssa
            $searchOwner = New-Object -TypeName "$adminNamespace.SearchObjectOwner" `
                                      -ArgumentList @(
                                          $objectLevel::Ssa, 
                                          $searchSite
                                      )
            $resultSources = Get-SPEnterpriseSearchResultSource -SearchApplication $ssa -Owner $searchOwner
            foreach($resultSource in $resultSources)
            {
                $Script:dscConfigContent += "        SPSearchResultSource " + [System.Guid]::NewGuid().ToString() + "`r`n"
                $Script:dscConfigContent += "        {`r`n"
                $params.SearchServiceAppName = $ssa.DisplayName
                $params.Name = $resultSource.Name
                $results = Get-TargetResource @params
                if($null -eq $results.Get_Item("ConnectionUrl"))
                {
                    $results.Remove("ConnectionUrl")
                }
                $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
                $Script:dscConfigContent += "        }`r`n"
            }
        }
    }
}

function Read-SPSearchCrawlRule
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPSearchCrawlRule\MSFT_SPSearchCrawlRule.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $ssa = Get-SPServiceApplication | Where-Object -FilterScript { 
            $_.GetType().FullName -eq "Microsoft.Office.Server.Search.Administration.SearchServiceApplication" 
    }
    if($null -ne $ssa)
    {
        $crawlRules = Get-SPEnterpriseSearchCrawlRule -SearchApplication $ssa

        foreach($crawlRule in $crawlRules)
        {
            $Script:dscConfigContent += "        SPSearchCrawlRule " + [System.Guid]::NewGuid().ToString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.ServiceAppName = $ssa.DisplayName
            $params.Path = $crawlRule.Path
            $params.Remove("CertificateName")
            $results = Get-TargetResource @params
            $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
    }
}

function Read-SPOfficeOnlineServerBinding
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPOfficeOnlineServerBinding\MSFT_SPOfficeOnlineServerBinding.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $Script:dscConfigContent += "        SPOfficeOnlineServerBinding " + [System.Guid]::NewGuid().ToString() + "`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $results = Get-TargetResource @params
    $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
    $Script:dscConfigContent += "        }`r`n"
}

function Read-SPIrmSettings
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPIrmSettings\MSFT_SPIrmSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $Script:dscConfigContent += "        SPIrmSettings " + [System.Guid]::NewGuid().ToString() + "`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $results = Get-TargetResource @params
    if(!$results.ContainsKey("InstallAccount"))
    {
        $results.Add("InstallAccount", "`$CredsFarmAccount")
    }
    $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
    $Script:dscConfigContent += "        }`r`n"
}

function Read-SPHealthAnalyzerRuleState
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPHealthAnalyzerRuleState\MSFT_SPHealthAnalyzerRuleState.psm1")
    $caWebapp = Get-SPWebApplication -IncludeCentralAdministration `
            | Where-Object -FilterScript {
                $_.IsAdministrationWebApplication
            }
    $caWeb = Get-SPWeb($caWebapp.Url)
    $healthRulesList = $caWeb.Lists | Where-Object -FilterScript { 
        $_.BaseTemplate -eq "HealthRules"
    }

    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    foreach($healthRule in $healthRulesList.Items)
    {
        $params.Name = $healthRule.Title
        $Script:dscConfigContent += "        SPHealthAnalyzerRuleState " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $results = Get-TargetResource @params
        if($results.Get_Item("Schedule") -eq "On Demand")
        {
            $results.Schedule = "OnDemandOnly"    
        }
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPFarmSolution
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFarmSolution\MSFT_SPFarmSolution.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $solutions = Get-SPSolution
    $farm = Get-SPFarm
    foreach($solution in $solutions)
    {
        $file = $farm.Solutions.Item($solution.Name).SolutionFile
        $filePath = (Get-Item -Path ".\" -Verbose).FullName + "\" + $solution.Name
        $file.SaveAs($filePath)
        $Script:dscConfigContent += "        SPFarmSolution " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Name = $solution.Name
        $results = Get-TargetResource @params
        if($results.ContainsKey("ContainsGlobalAssembly"))
        {
            $results.Remove("ContainsGlobalAssembly")
        }
        if(!$results.ContainsKey("InstallAccount"))
        {
            $results.Add("InstallAccount", "`$CredsFarmAccount")
        }
        $results["LiteralPath"] = $filePath
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPFeature
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFeature\MSFT_SPFeature.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $features = Get-SPFeature
    foreach($feature in $features)
    {
        $Script:dscConfigContent += "        SPFeature " + $feature.DisplayName.Replace(" ", "") + "-" + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Name = $solution.Name
        $results = Get-TargetResource @params
        if($results.ContainsKey("ContainsGlobalAssembly"))
        {
            $results.Remove("ContainsGlobalAssembly")
        }
        if(!$results.ContainsKey("InstallAccount"))
        {
            $results.Add("InstallAccount", "`$CredsFarmAccount")
        }
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPFarmAdministrators
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPFarmAdministrators\MSFT_SPFarmAdministrators.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $params.Remove("MembersToInclude")
    $params.Remove("MembersToExclude")
    $Script:dscConfigContent += "        SPFarmAdministrators " + [System.Guid]::NewGuid().ToString() + "`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $results = Get-TargetResource @params
    $results.Name = "SPFarmAdministrators"
    $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
    $Script:dscConfigContent += "        }`r`n"
}

function Read-SPExcelServiceApp
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPExcelServiceApp\MSFT_SPExcelServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $excelSSA = Get-SPServiceApplication | ?{$_.TypeName -eq "Excel Services Application Web Service Application"}

    if($null -ne $excelSSA)
    {
        $Script:dscConfigContent += "        SPExcelServiceApp " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Name = $excelSSA.DisplayName
        $results = Get-TargetResource @params
        $privateK = $results.Get_Item("PrivateBytesMax")
        $unusedMax = $results.Get_Item("UnusedObjectAgeMax")
        <# Nik20170106 - Temporary fix while waiting to hear back from Brian F. on how to properly pass these params. #>
        if($results.ContainsKey("TrustedFileLocations"))
        {
            $results.Remove("TrustedFileLocations")
        }
        if($results.ContainsKey("PrivateBytesMax") -and $privateK -eq "-1")
        {
            $results.Remove("PrivateBytesMax")
        }
        if($results.ContainsKey("UnusedObjectAgeMax") -and $unusedMax -eq "-1")
        {
            $results.Remove("UnusedObjectAgeMax")
        }
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

<# Nik20170106 - Read the Designer Settings of either the Site Collection or the Web Application #>
function Read-SPDesignerSettings($receiver)
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPDesignerSettings\MSFT_SPDesignerSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $params.Url = $receiver[0]
    $params.SettingsScope = $receiver[1]
    $results = Get-TargetResource @params

    <# Nik20170106 - The logic here differs from other Read functions due to a bug in the Designer Resource that doesn't properly obtains a reference to the Site Collection. #>
    if($null -ne $results)
    {
        $Script:dscConfigContent += "        SPDesignerSettings " + $receiver[1] + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        <#if($receiver[1] -eq "SiteCollection" -and $results.ContainsKey("InstallAccount"))
        {
            $results.Add("PsDscRunAsAccount", $results.Get_Item("InstallAccount"))
            $results.Remove("InstallAccount")
        }#>
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        if($receiver.Length -eq 3)
        {
            $Script:dscConfigContent += "            DependsOn = `"[SP" + $receiver[1].Replace("Collection", "") + "]" + $receiver[2] + "`";`r`n"
        }
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPDatabaseAAG
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPDatabaseAAG\MSFT_SPDatabaseAAG.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $databases = Get-SPDatabase
    foreach($database in $databases)
    {
        if($null -ne $database.AvailabilityGroup)
        {
            $Script:dscConfigContent += "        SPDatabaseAAG " + [System.Guid]::NewGuid().ToString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.DatabaseName = $database.Name
            $params.AGName = $configDatabase.AvailabilityGroup
            $results = Get-TargetResource @params
            $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
    }
}

function Read-SPConfigWizard
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPConfigWizard\MSFT_SPConfigWizard.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $Script:dscConfigContent += "        SPConfigWizard " + [System.Guid]::NewGuid().ToString() + "`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $results = Get-TargetResource @params
    if(!$results.ContainsKey("InstallAccount"))
    {
        $results.Add("InstallAccount", "`$CredsFarmAccount")
    }
    $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
    $Script:dscConfigContent += "        }`r`n"
}

function Read-SPBlobCacheSettings
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPBlobCacheSettings\MSFT_SPBlobCacheSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module

    $webApps = Get-SPWebApplication
    foreach($webApp in $webApps)
    {
        $alternateUrls = $webApp.AlternateUrls
        foreach($alternateUrl in $alternateurls)
        {
            $Script:dscConfigContent += "        SPBlobCacheSettings " + [System.Guid]::NewGuid().ToString() + "`r`n"
            $Script:dscConfigContent += "        {`r`n"
            $params.WebAppUrl = $webApp.Url
            $params.Zone = $alternateUrls.Zone
            $results = Get-TargetResource @params
            $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
            $Script:dscConfigContent += "        }`r`n"
        }
    }
}

function Read-SPAppManagementServiceApp
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAppManagementServiceApp\MSFT_SPAppManagementServiceApp.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $serviceApps = Get-SPServiceApplication | ? {$_.TypeName -eq "App Management Service Application"}

    foreach($appManagement in $serviceApps)
    {
        $Script:dscConfigContent += "        SPAppManagementServiceApp " + $appManagement.Name.Replace(" ", "") + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.Name = $appManagement.Name

        $results = Get-TargetResource @params
        <# Nik20170111 - Fixes a bug in 1.5.0.0 where the Database Name and Server is not properly returned; #>
        $results.DatabaseName = $appManagement.Databases.Name
        $results.DatabaseServer = $appManagement.Databases.Server.Name

        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPAppStoreSettings
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAppStoreSettings\MSFT_SPAppStoreSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $webApps = Get-SPWebApplication

    foreach($webApp in $webApps)
    {
        $Script:dscConfigContent += "        SPAppStoreSettings " + $webApp.Name.Replace(" ", "") + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.WebAppUrl = $webApp.Url
        $results = Get-TargetResource @params
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPAntivirusSettings
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAntivirusSettings\MSFT_SPAntivirusSettings.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $Script:dscConfigContent += "        SPAntivirusSettings AntivirusSettings`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $results = Get-TargetResource @params
    $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
    $Script:dscConfigContent += "        }`r`n"    
}

function Read-SPDistributedCacheService
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPDistributedCacheService\MSFT_SPDistributedCacheService.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $params.Name = "DistributedCache"
    $results = Get-TargetResource @params
    if($results.Get_Item("Ensure").ToLower() -eq "present")
    {
        $Script:dscConfigContent += "        SPDistributedCacheService " + [System.Guid]::NewGuid().ToString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"
    }
}

function Read-SPAlternateUrl
{
    $module = Resolve-Path ($Script:SPDSCPath + "\DSCResources\MSFT_SPAlternateUrl\MSFT_SPAlternateUrl.psm1")
    Import-Module $module
    $params = Get-DSCFakeParameters -ModulePath $module
    $alternateUrls = Get-SPAlternateUrl

    foreach($alternateUrl in $alternateUrls)
    {
        $Script:dscConfigContent += "        SPAlternateUrl " + [System.Guid]::NewGuid().toString() + "`r`n"
        $Script:dscConfigContent += "        {`r`n"
        $params.WebAppUrl = $alternateUrl.Uri.AbsoluteUri
        $params.Zone = $alternateUrl.UrlZone
        $results = Get-TargetResource @params
        $Script:dscConfigContent += Get-DSCBlock -Params $results -ModulePath $module
        $Script:dscConfigContent += "        }`r`n"  
    }
}

<## This function sets the settings for the Local Configuration Manager (LCM) component on the server we will be configuring using our resulting DSC Configuration script. The LCM component is the one responsible for orchestrating all DSC configuration related activities and processes on a server. This method specifies settings telling the LCM to not hesitate rebooting the server we are configurating automatically if it requires a reboot (i.e. During the SharePoint Prerequisites installation). Setting this value helps reduce the amount of manual interaction that is required to automate the configuration of our SharePoint farm using our resulting DSC Configuration script. #>
function Set-LCM
{
    $Script:dscConfigContent += "        LocalConfigurationManager"  + "`r`n"
    $Script:dscConfigContent += "        {`r`n"
    $Script:dscConfigContent += "            RebootNodeIfNeeded = `$True`r`n"
    $Script:dscConfigContent += "        }`r`n"
}

function Invoke-SQL {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Server,
        [Parameter(Mandatory=$true)]
        [string]$dbName,
        [Parameter(Mandatory=$true)]
        [string]$sqlQuery
    )
 
    $ConnectString="Data Source=${Server}; Integrated Security=SSPI; Initial Catalog=${dbName}"
 
    $Conn= New-Object System.Data.SqlClient.SQLConnection($ConnectString)
    $Command = New-Object System.Data.SqlClient.SqlCommand($sqlQuery,$Conn)
    $Conn.Open()
 
    $Adapter = New-Object System.Data.SqlClient.SqlDataAdapter $Command
    $DataSet = New-Object System.Data.DataSet
    $Adapter.Fill($DataSet) | Out-Null
 
    $Conn.Close()
    $DataSet.Tables
}


<## This method is used to determine if a specific PowerShell cmdlet is available in the current Powershell Session. It is currently used to determine wheter or not the user has access to call the Invoke-SqlCmd cmdlet or if he needs to install the SQL Client coponent first. It simply returns $true if the cmdlet is available to the user, or $false if it is not. #>
function Test-CommandExists
{
    param ($command)

    $errorActionPreference = "stop"
    try {
        if(Get-Command $command)
        {
            return $true
        }
    }
    catch
    {
        return $false
    }
}

function Get-SPReverseDSC()
{
    <## Call into our main function that is responsible for extracting all the information about our SharePoint farm. #>
    Orchestrator

    <## Prompts the user to specify the FOLDER path where the resulting PowerShell DSC Configuration Script will be saved. #>
    $OutputDSCPath = Read-Host "Please enter the full path of the output folder for DSC Configuration (will be created as necessary)"

    <## Ensures the specified output folder path actually exists; if not, tries to create it and throws an exception if we can't. ##>
    while (!(Test-Path -Path $OutputDSCPath -PathType Container -ErrorAction SilentlyContinue))
    {
        try
        {
            Write-Output "Directory `"$OutputDSCPath`" doesn't exist; creating..."
            New-Item -Path $OutputDSCPath -ItemType Directory | Out-Null
            if ($?) {break}
        }
        catch
        {
            Write-Warning "$($_.Exception.Message)"
            Write-Warning "Could not create folder $OutputDSCPath!"
        }
        $OutputDSCPath = Read-Host "Please Enter Output Folder for DSC Configuration (Will be Created as Necessary)"
    }
    <## Ensures the path we specify ends with a Slash, in order to make sure the resulting file path is properly structured. #>
    if(!$OutputDSCPath.EndsWith("\") -and !$OutputDSCPath.EndsWith("/"))
    {
        $OutputDSCPath += "\"
    }

    <## Save the content of the resulting DSC Configuration file into a file at the specified path. #>
    $outputDSCFile = $OutputDSCPath + "SP-Farm.DSC.ps1"
    $Script:dscConfigContent | Out-File $outputDSCFile
    Write-Output "Done."
    <## Wait a couple of seconds, then open our $outputDSCPath in Windows Explorer so we can review the glorious output. ##>
    Start-Sleep 2
    Invoke-Item -Path $OutputDSCPath
}

<## This function defines variables of type Credential for the resulting DSC Configuraton Script. Each variable declared in this method will result in the user being prompted to manually input credentials when executing the resulting script. #>
function Set-ObtainRequiredCredentials
{
    # Farm Account
    $localspFarmAccount = $Global:spCentralAdmin.ApplicationPool.ProcessAccount.Name
    $requiredCredentials = @($localspFarmAccount)
    $managedAccounts = Get-SPManagedAccount
    foreach($managedAccount in $managedAccounts)
    {
        $requiredCredentials += $managedAccounts.UserName
    }

    $spServiceAppPools = Get-SPServiceApplicationPool
    foreach($spServiceAppPool in $spServiceAppPools)
    {
        $requiredCredentials += $spServiceAppPools.ProcessAccount.Name
    }

    $requiredCredentials = $requiredCredentials | Select -Unique

    foreach($account in $requiredCredentials)
    {
        $accountName = $account
        if($account -eq $localspFarmAccount)
        {
            $accountName = "FarmAccount"
        }
        else
        {
            $accountParts = $accountName.Split('\')
            if($accountParts.Length -gt 1)
            {
                $accountName = $accountParts[1]
            }
        }
        $Script:dscConfigContent += "    `$Creds" + $accountName + "= Get-Credential -UserName `"" + $account + "`" -Message `"Credentials for " + $account + "`"`r`n"
    }

    $Script:dscConfigContent += "`r`n"
}

function Get-RequiredModules
{
    try
    {
        if ($null -eq (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue))
        {
            # Install NuGet
            Write-Host " - Installing NuGet..."
            Install-PackageProvider -Name NuGet -Force -ForceBootstrap | Out-Null
        }
        [array]$requiredModules = "PowerShellGet",`
                                  "PackageManagement",`
                                  "SharePointDSC"
        Write-Host " - Checking for required PowerShell modules..."
        # Because SkipPublisherCheck and AllowClobber parameters don't seem to be supported on Win2012R2 let's set whether the parameters are specified here
        if (Get-Command -Name Install-Module -ParameterName AllowClobber -ErrorAction SilentlyContinue)
        {
            $allowClobberParameter = @{AllowClobber = $true}
        }
        else {$allowClobberParameter = @{}}
        if (Get-Command -Name Install-Module -ParameterName SkipPublisherCheck -ErrorAction SilentlyContinue)
        {
            $skipPublisherCheckParameter = @{SkipPublisherCheck = $true}
        }
        else {$skipPublisherCheckParameter = @{}}
        foreach ($requiredModule in $requiredModules)
        {
            [array]$installedModules = Get-Module -ListAvailable -FullyQualifiedName $requiredModule
            if ($null -eq $installedModules)
            {
                # Install required module since it wasn't detected
                $onlineModule = Find-Module -Name $requiredModule -ErrorAction SilentlyContinue
                if ($onlineModule)
                {
                    Write-Host -ForegroundColor DarkYellow  "  - Module $requiredModule not present. Installing version $($onlineModule.Version)..." -NoNewline
                    Install-Module -Name $requiredModule -ErrorAction Inquire -Force @allowClobberParameter @skipPublisherCheckParameter
                    if ($?) {Write-Host -ForegroundColor Green "Done."}
                }
                else
                {
                    Write-Host -ForegroundColor Yellow "  - Module $requiredModule not present, and was not found in the PowerShell Gallery for installation/update."
                }
            }
            else
            {
                $installedModule = Get-InstalledModule -Name $requiredModule -ErrorAction SilentlyContinue
                if ($installedModule)
                {
                    # If we were successful in querying the module this way it was probably originally installed from the Gallery
                    $installedModuleWasFromGallery = $true
                }
                else # Was probably pre-installed or installed manually
                {
                    # Grab the newest version in case there are multiple
                    $installedModule = ($installedModules | Sort-Object Version -Descending)[0]
                    $installedModuleWasFromGallery = $false
                }
                # Look for online updates to already-installed required module
                Write-Host "  - Module $requiredModule version $($installedModule.Version) is already installed. Looking for updates..." -NoNewline
                $onlineModule = Find-Module -Name $requiredModule -ErrorAction SilentlyContinue
                if ($null -eq $onlineModule)
                {
                    Write-Host -ForegroundColor Yellow "Not found in the PowerShell Gallery!"
                }
                else
                {
                    # Get the last module
                    if ($installedModule.Version -eq $onlineModule.version)
                    {
                        # Online and local versions match; no action required
                        Write-Host -ForegroundColor Gray "Already up-to-date ($($installedModule.Version))."
                    }
                    else
                    {
                        Write-Host -ForegroundColor Magenta "Newer version $($onlineModule.Version) found!"
                        if ($installedModule -and $installedModuleWasFromGallery)
                        {
                            # Update to newest online version using PowerShellGet
                            Write-Host "  - Updating module $requiredModule..." -NoNewline
                            Update-Module -Name $requiredModule -Force -ErrorAction Continue
                            if ($?) {Write-Host -ForegroundColor Green "Done."}
                        }
                        else
                        {
                            # Update won't work as it appears the module wasn't installed using the PS Gallery initially, so let's try a straight install
                            Write-Host "  - Installing $requiredModule..." -NoNewline
                            Install-Module -Name $requiredModule -Confirm:$false -Force @allowClobberParameter @skipPublisherCheckParameter
                            if ($?) {Write-Host -ForegroundColor Green "Done."}
                        }
                    }
                    # Now check if we have more than one version installed
                    [array]$installedModules = Get-Module -ListAvailable -FullyQualifiedName $requiredModule
                    if ($installedModules.Count -gt 1)
                    {
                        # Remove all non-current module versions including ones that weren't put there via the PowerShell Gallery
                        [array]$oldModules = $installedModules | Where-Object {$_.Version -ne $onlineModule.Version}
                        foreach ($oldModule in $oldModules)
                        {
                            Write-Host "  - Uninstalling old version $($oldModule.Version) of $($oldModule.Name)..." -NoNewline
                            Uninstall-Module -Name $oldModule.Name -RequiredVersion $oldModule.Version -Force -ErrorAction SilentlyContinue
                            if ($?) {Write-Host -ForegroundColor Green "Done."}
                            # Unload the old module in case it was automatically loaded in this console
                            if (Get-Module -Name $oldModule.Name -ErrorAction SilentlyContinue)
                            {
                                Write-Host "  - Unloading prior loaded version $($oldModule.Version) of $($oldModule.Name)..." -NoNewline
                                Remove-Module -Name $oldModule.Name -Force -ErrorAction Inquire
                                if ($?) {Write-Host -ForegroundColor Green "Done."}
                            }
                            Write-Host "  - Removing old module files from $($oldModule.ModuleBase)..." -NoNewline
                            Remove-Item -Path $oldModule.ModuleBase -Recurse -Confirm:$true -ErrorAction SilentlyContinue
                            if ($?) {Write-Host -ForegroundColor Green "Done."}
                        }
                    }
                }
            }
            $installedModule = Get-InstalledModule -Name $requiredModule -ErrorAction SilentlyContinue
            if ($null -eq $installedModule)
            {
                # Module was not installed from the Gallery, so we look for it an alternate way
                $installedModule = Get-Module -Name $requiredModule -ListAvailable | Sort-Object Version | Select-Object -Last 1
            }
            Write-Output ""
            Write-Output "  --"
            # Clean up the variables
            Remove-Variable -Name installedModules -ErrorAction SilentlyContinue
            Remove-Variable -Name installedModule -ErrorAction SilentlyContinue
            Remove-Variable -Name oldModules -ErrorAction SilentlyContinue
            Remove-Variable -Name oldModule -ErrorAction SilentlyContinue
            Remove-Variable -Name onlineModule -ErrorAction SilentlyContinue
        }
        Write-Host " - Done checking required modules."
    }
    catch
    {
        <# Nik20161220 - Not sure we want to throw an error if the server doesn't have internet connectivity as this will mostly
                         represent close to 90% of the scenarios; #>
        #Write-Output $_.Exception
        #Write-Error "Unable to download/install $requiredModule - check Internet access etc."
    }
}

Add-PSSnapin Microsoft.SharePoint.PowerShell -EA SilentlyContinue
$sharePointSnapin = Get-PSSnapin | ?{$_.Name -eq "Microsoft.SharePoint.PowerShell"}
if($null -ne $sharePointSnapin)
{
    Get-RequiredModules
    Get-SPReverseDSC
}
else
{
    Write-Host "`r`nE101"  -BackgroundColor Red -ForegroundColor Black -NoNewline
    Write-Host "    - We couldn't detect a SharePoint instalation on this machine. Please execute the SharePoint Reverse DSC script on an existing SharePoint server."
}
