param ([string]$XmlPath="C:\Users\nicharl\Documents\My Received Files\AutoSPInstallerInput-Azure-SingleServer-2010.xml", [string]$OutputPath) 

#region Script Variables
[xml]$ASPIXml = Get-Content $XmlPath
$DSCContent = ""
$ASPIVersion = $ASPIXml.Configuration.Version
$ASPIEnvironment = $ASPIXml.Configuration.Environment
$nl = [Environment]::NewLine
#endregion

#region Header
$DSCContent += "Configuration " + $ASPIEnvironment + $nl
$DSCContent += "{" + $nl
$DSCContent += "    `$FarmUserName = `"" + $ASPIXml.Configuration.Farm.Account.Username + "`"" + $nl
$DSCContent += "    `$FarmPassword = `"" + $ASPIXml.Configuration.Farm.Account.Password + "`" | ConvertTo-SecureString -AsPlainText -Force" + $nl
$DSCContent += "    `$FarmAccount = New-Object System.Management.Automation.PSCredential -ArgumentList `$FarmUserName, `$FarmPassword" + $nl
$DSCContent += "    `$Passphrase = `"" + $ASPIXml.Configuration.Farm.Passphrase + "`" | ConvertTo-SecureString -AsPlainText -Force" + $nl
#endregion

#region SPFarm
$DSCContent += "    SPCreateFarm SPFarm" + $nl
$DSCContent += "    {" + $nl
$DSCContent += "        FarmConfigDatabaseName = `"" + $ASPIXml.Configuration.Farm.Database.ConfigDB + "`"" + $nl
$DSCContent += "        DatabaseServer = `"" + $ASPIXml.Configuration.Farm.Database.DBServer + "`"" + $nl
$DSCContent += "        FarmAccount = `$FarmAccount" + $nl
$DSCContent += "        Passphrase = `$Passphrase" + $nl
$DSCContent += "        AdminContentDatabase = `"" + $ASPIXml.Configuration.Farm.CentralAdmin.Database + "`"" + $nl
$DSCContent += "        CentralAdministrationPort = " + $ASPIXml.Configuration.Farm.CentralAdmin.Port + $nl
$DSCContent += "    }" + $nl
#endregion

#region SPServiceInstance
if(![string]::IsNullOrEmpty($ASPIXml.Configuration.Farm.Services.SandboxedCodeService))
{
	$DSCContent += "    SPServiceInstance SandboxCodeService" + $nl
	$DSCContent += "    {" + $nl
	$DSCContent += "        Name = `"SandboxCodeService`"" + $nl
	$DSCContent += "        Ensure = `"Present`"" + $nl
	$DSCContent += "    }" + $nl
}
#endregion

#region SPWebApplications
foreach($webApp in $ASPIXml.Configuration.WebApplications.WebApplication)
{
	[boolean]$useSSL = $false;
	if($webApp.Url -like 'https://*')
	{
		$useSSL = $true
	}
	$DSCContent += "    SPWebApplication " + $webApp.Name + $nl
	$DSCContent += "    {" + $nl
	$DSCContent += "        Url = `"" + $webApp.Url + "`"" + $nl
	if(![string]::IsNullOrEmpty($webApp.ApplicationPool))
	{
		$DSCContent += "        ApplicationPool = `"" + $webApp.ApplicationPool + "`"" + $nl
	}	
	if($webApp.Database -ne $null)
	{
		if(![string]::IsNullOrEmpty($webApp.Database.DBServer))
		{
			$DSCContent += "        DatabaseServer = `"" + $webApp.Database.DBServer + "`"" + $nl
		}
		if(![string]::IsNullOrEmpty($webApp.Database.Name))
		{
			$DSCContent += "        DatabaseName = `"" + $webApp.Database.Name + "`"" + $nl
		}
	}
	if(![string]::IsNullOrEmpty($webApp.Port))
	{
		$DSCContent += "        Port = " + $webApp.Port + $nl
	}
	$DSCContent += "        UseSSL = `$$useSSL" + $nl
	if(![string]::IsNullOrEmpty($webApp.UseHostHeader) -and $webApp.UseHostHeader.ToLower() -eq "true")
	{
		$DSCContent += "        HostHeader = `"" + $webApp.Url.ToLower().Replace("https://", "").Replace("http://", "").Split('/')[0] + "`"" + $nl
	}
	
	$DSCContent += "    }" + $nl

	#region SPSite
	foreach($spSite in $webApp.SiteCollections.SiteCollection)
	{
		$DSCContent += "    SPSite " + $spSite.Name + $nl
		$DSCContent += "    {" + $nl
		$DSCContent += "        Name = `"" + $spSite.Name + "`"" + $nl
		if(![string]::IsNullOrEmpty($spSite.Owner))
		{
			$DSCContent += "        OwnerAlias = `"" + $spSite.Owner + "`"" + $nl
		}
		if(![string]::IsNullOrEmpty($spSite.Description))
		{	
			$DSCContent += "        Description = `"" + $spSite.Description + "`"" + $nl
		}
		$DSCContent += "        Url = `"" + $spSite.siteUrl + "`"" + $nl
		if(![string]::IsNullOrEmpty($spSite.siteUrl))
		{
			$DSCContent += "        Language = " + $spSite.LCID + $nl
		}
		if(![string]::IsNullOrEmpty($spSite.Template))
		{
			$DSCContent += "        Template = `"" + $spSite.Template + "`"" + $nl;
		}
		if(![string]::IsNullOrEmpty($spSite.CustomDatabase))
		{
			$DSCContent += "        ContentDatabase = `"" + $spSite.CustomDatabase + "`"" + $nl
		}
		if(![string]::IsNullOrEmpty($spSite.HostNamedSiteCollection) -and $spSite.HostNamedSiteCollection.ToLower() -eq "true")
		{
			$DSCContent += "        HostHeaderWebApplication = `"" + $webApp.Url + "`"" + $nl
		}
		$DSCContent += "        DependsOn = `"[SPWebApplication]" + $webApp.Name + "`"" + $nl
		$DSCContent += "    }" + $nl
	}
	#endregion

	#region SPCacheAccount
	if(![string]::IsNullOrEmpty($ASPIXml.Configuration.Farm.ObjectCacheAccounts.SuperReader))
	{
		$DSCContent += "    SPCacheAccount CacheAccounts" + $webApp.Name + $nl
		$DSCContent += "    {" + $nl
		$DSCContent += "        WebAppUrl = `"" + $webApp.Url + "`"" + $nl
		$DSCContent += "        SuperUserAlias = `"" + $ASPIXml.Configuration.Farm.ObjectCacheAccounts.SuperUser + "`"" + $nl
		$DSCContent += "        SuperReaderAlias = `"" + $ASPIXml.Configuration.Farm.ObjectCacheAccounts.SuperReader + "`"" + $nl
		$DSCContent += "    }" + $nl
	}
#endregion
}
#endregion
$DSCContent += "}" + $nl
$outputFilePrefix = ($XmlPath -split "\\" | Select-Object -Last 1).TrimEnd(".xml")
Write-Host $DSCContent
Out-File -FilePath "$OutputPath\$outputFilePrefix.ps1" -InputObject $DSCContent
