########################################################################
# Detection Script for Office 365 Apps Install
# Created by Matt Lavine [@mattlavine](https://github.com/mattlavine)
# Contains code forked from @JankeSkanke
########################################################################
#DetectionScript
function Write-LogEntry {
	param (
		[parameter(Mandatory = $true, HelpMessage = "Value added to the log file.")]
		[ValidateNotNullOrEmpty()]
		[string]$Value,
		[parameter(Mandatory = $true, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning and 3 for Error.")]
		[ValidateNotNullOrEmpty()]
		[ValidateSet("1", "2", "3")]
		[string]$Severity,
		[parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
		[ValidateNotNullOrEmpty()]
		[string]$FileName = $LogFileName
	)
	# Determine log file location
	$LogFilePath = Join-Path -Path $env:SystemRoot -ChildPath $("Temp\$FileName")
	
	# Construct time stamp for log entry
	$Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
	
	# Construct date for log entry
	$Date = (Get-Date -Format "MM-dd-yyyy")
	
	# Construct context for log entry
	$Context = $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
	
	# Construct final log entry
	$LogText = "<![LOG[$($Value)]LOG]!><time=""$($Time)"" date=""$($Date)"" component=""$($LogFileName)"" context=""$($Context)"" type=""$($Severity)"" thread=""$($PID)"" file="""">"
	
	# Add value to log file
	try {
		Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFilePath -ErrorAction Stop
		if ($Severity -eq 1) {
			Write-Verbose -Message $Value
		} elseif ($Severity -eq 3) {
			Write-Warning -Message $Value
		}
	} catch [System.Exception] {
		Write-Warning -Message "Unable to append log entry to $LogFileName.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
	}
}
#--------------------------------------------------
#--------------------------------------------------
#--------------------------------------------------
#--------------------------------------------------

# List of Product IDs that are supported by the Office Deployment Tool for Click-to-Run
# https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/installation/product-ids-supported-office-deployment-click-to-run

#--------------------------------------------------

# (REQUIRED) The Display Name for the App Install to match against. This is a CONTAINS match, not an EXACT match, to account for the culture code (ex. "en-us") being appended to the display name.
$RequiredM365AppsDisplayName = "Microsoft 365 Apps for enterprise"

# (REQUIRED) The Bit Version/Architecture of the App Install ("x64" or "x86")
$RequiredPlatform = "x64"

# (REQUIRED) The Download method of the install (ex. "Local", "CDN")
$RequiredMediaType = "CDN"

# (REQUIRED) The Required ProductReleaseID to check
$RequiredProductReleaseID = "O365ProPlusRetail"

# (OPTIONAL) List of Required O365 Excluded Apps to check. The list is a CONTAINS match, not an EXACT match. As long as the list found contains AT LEAST these items then it will pass.
$RequiredM365ExcludedApps = @("groove", "lync", "bing")

#--------------------------------------------------

$RegistryUninstallKeys = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$ClickToRunConfigurationRegistryPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

#--------------------------------------------------
#--------------------------------------------------
#--------------------------------------------------
#--------------------------------------------------

$LogFileName = "M365AppsSetup.log"
Write-LogEntry -Value "Starting `"$RequiredM365AppsDisplayName`" install detection logic" -Severity 1

#--------------------------------------------------

# Get the Display Name for the current Microsoft 365 Apps install
$M365AppsDisplayNameCheck = $RegistryUninstallKeys | Get-ItemProperty | Where-Object { $_.DisplayName -match $RequiredM365AppsDisplayName }

if ($RequiredM365AppsDisplayName) {
	Write-LogEntry -Value "The following Microsoft 365 App install is REQUIRED: `"$RequiredM365AppsDisplayName`"" -Severity 1
}

# add line for if key does not exist
if ($M365AppsDisplayNameCheck) {
	Write-LogEntry -Value "The following Microsoft 365 App install was FOUND: `"$(($M365AppsDisplayNameCheck).DisplayName)`"" -Severity 1
} else {
	Write-LogEntry -Value "A matching Microsoft 365 App install was NOT FOUND." -Severity 2
}

#--------------------------------------------------

# Get the Platform registry key
$Platform = (Get-ItemProperty -Path $ClickToRunConfigurationRegistryPath -Name "Platform" -ErrorAction SilentlyContinue).Platform

if ($RequiredPlatform) {
	Write-LogEntry -Value "The following Platform is REQUIRED: `"$RequiredPlatform`"" -Severity 1
}

# add line for if key does not exist
if ($Platform) {
	Write-LogEntry -Value "The following Platform was FOUND: `"$Platform`"" -Severity 1
} else {
	Write-LogEntry -Value "The Platform key was empty or didn't exist." -Severity 2
}

#--------------------------------------------------

# Check the MediaType registry key (formatted as "[ProductReleaseID].MediaType" - e.g. "O365ProPlusRetail.MediaType") for the download method used for install. Example: "Local", "CDN"
$MediaType = (Get-ItemProperty -Path $ClickToRunConfigurationRegistryPath -Name "$RequiredProductReleaseID.MediaType" -ErrorAction SilentlyContinue)."$RequiredProductReleaseID.MediaType"

if ($RequiredMediaType) {
	Write-LogEntry -Value "The following MediaType (`"$RequiredProductReleaseID.MediaType`") is REQUIRED: `"$RequiredMediaType`"" -Severity 1
}

# add line for if key does not exist
if ($MediaType) {
	Write-LogEntry -Value "The following MediaType (`"$RequiredProductReleaseID.MediaType`") was FOUND: `"$MediaType`"" -Severity 1
} else {
	Write-LogEntry -Value "The MediaType key was empty or didn't exist." -Severity 2
}

#--------------------------------------------------

# Check the ProductReleaseIds registry key for a list of items
$ProductReleaseIDs = (Get-ItemProperty -Path $ClickToRunConfigurationRegistryPath -Name "ProductReleaseIds" -ErrorAction SilentlyContinue).ProductReleaseIds
$ProductReleaseIDList = $ProductReleaseIDs -split ','

if ($RequiredProductReleaseID) {
	Write-LogEntry -Value "The following ProductReleaseId is REQUIRED: `"$RequiredProductReleaseID`"" -Severity 1
}

# add line for if key is not found
if ($ProductReleaseIDs) {
	Write-LogEntry -Value "The following ProductReleaseIds were FOUND: `"$ProductReleaseIDs`"" -Severity 1
} else {
	Write-LogEntry -Value "The ProductReleaseIds key was empty or didn't exist." -Severity 2
}

# Assume the required ProductReleaseID isn't present until proven otherwise. Compare it to the list of installed ProductReleaseIDs
$RequiredProductReleaseIDPresent = $false
if ($ProductReleaseIDList -contains $RequiredProductReleaseID) {
	$RequiredProductReleaseIDPresent = $true
}

#--------------------------------------------------

# Skip this section if $RequiredM365ExcludedApps was left empty
if (-not $RequiredM365ExcludedApps) {
	Write-LogEntry -Value "(NOT APPLICABLE) Required Excluded Apps were not configured. Skipping..." -Severity 1
	$AllRequiredM365ExcludedAppsPresent = $true
} else {
	# Proceed with this section if $RequiredM365ExcludedApps was configured
	# Check the ExcludedApps registry key (formatted as "[ProductReleaseID].ExcludedApps" - e.g. "O365ProPlusRetail.ExcludedApps") for a list of excluded apps
	$M365ExcludedApps = (Get-ItemProperty -Path $ClickToRunConfigurationRegistryPath -Name "$RequiredProductReleaseID.ExcludedApps" -ErrorAction SilentlyContinue)."$RequiredProductReleaseID.ExcludedApps"
	$M365ExcludedAppsList = $M365ExcludedApps -split ','
	
	if ($RequiredM365ExcludedApps) {
		Write-LogEntry -Value "The following ExcludedApps (`"$RequiredProductReleaseID.ExcludedApps`") are REQUIRED: `"$RequiredM365ExcludedApps`"" -Severity 1
	}
	
	# add line for if key does not exist
	if ($M365ExcludedApps) {
		Write-LogEntry -Value "The following ExcludedApps (`"$RequiredProductReleaseID.ExcludedApps`") were FOUND: `"$M365ExcludedApps`"" -Severity 1
	} else {
		Write-LogEntry -Value "The ExcludedApps key was empty or didn't exist." -Severity 2
	}
	
	$AllRequiredM365ExcludedAppsPresent = $true
	# Loop through the list of required Excluded Apps in the registry and mark the variable as false if any of them aren't found in the list
	foreach ($M365ExcludedApp in $RequiredM365ExcludedApps) {
		if (-not ($M365ExcludedAppsList -contains $M365ExcludedApp)) {
			$AllRequiredM365ExcludedAppsPresent = $false
			break
		}
	}
}

#--------------------------------------------------

if ($M365AppsDisplayNameCheck -and ($Platform -eq $RequiredPlatform) -and ($MediaType -eq $RequiredMediaType) -and $RequiredProductReleaseIDPresent -and $AllRequiredM365ExcludedAppsPresent) {
	Write-LogEntry -Value "All detection conditions were met successfully." -Severity 1
	Write-Output "All detection conditions were met successfully."
	exit 0
} else {
	if (-not $M365AppsDisplayNameCheck) {
		Write-LogEntry -Value "`"$RequiredM365AppsDisplayName`" was not detected. This is not the desired state." -Severity 2
	}
	if ($Platform -ne $RequiredPlatform) {
		Write-LogEntry -Value "Platform was not set to `"$RequiredPlatform`". This is not the desired state." -Severity 2
	}
	if ($MediaType -ne $RequiredMediaType) {
		Write-LogEntry -Value "MediaType was not set to `"$RequiredMediaType`". This is not the desired state." -Severity 2
	}
	if (-not $RequiredProductReleaseIDPresent) {
		Write-LogEntry -Value "The required ProductReleaseId (`"$RequiredProductReleaseID`") was not found. This is not the desired state." -Severity 2
	}
	if (-not $AllRequiredM365ExcludedAppsPresent) {
		Write-LogEntry -Value "Not all required ExcludedApps (`"$RequiredM365ExcludedApps`") were listed. This is not the desired state." -Severity 2
	}
	Write-LogEntry -Value "One or more detections conditions have not been met. Detection has failed." -Severity 3
	#Write-Output "One or more detections conditions have not been met. Detection has failed."
	exit 1
}