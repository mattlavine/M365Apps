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

# The Display Name for the App Install to match against. This is an EXACT match.
#$ClassicTeamsDisplayName = "Teams Machine-Wide Installer"

$ClassicTeamsMachineWideInstaller64BitMSICode = "{731F6BAA-A986-45A4-8936-7C3AAAAA760B}"
#$ClassicTeamsMachineWideInstaller32BitMSICode = "{39AF0813-FA7B-4860-ADBE-93B9B214B914}"

#--------------------------------------------------

#$RegistryUninstallKeys64Bit = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$RegistryUninstallKeys32Bit = Get-ChildItem -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

#--------------------------------------------------

$LogFileName = "TeamsSetup.log"
Write-LogEntry -Value "Start Classic Teams Machine-Wide Installer Uninstall detection logic" -Severity 1

#--------------------------------------------------

# Check if the Classic Teams Machine-Wide Installer (64-bit) is installed on the system
# Both the 64-bit and 32-bit versions of the Classic Teams Machine-Wide Installer are listed in the WOW6432Node registry section on a 64-bit system.
# I'm only checking for the 64-bit Teams Machine-Wide Installer because, as of 10/8/24, the Teams Bootstrapper doesn't seem to detect the 32-bit version
#$ClassicTeamsMachineWideInstaller64BitCheck = $RegistryUninstallKeys32Bit | Get-ItemProperty | Where-Object { $_.DisplayName -eq $ClassicTeamsDisplayName }

$ClassicTeamsMachineWideInstaller64BitCheck = Test-Path -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$ClassicTeamsMachineWideInstaller64BitMSICode"

#if (-not $ClassicTeamsMachineWideInstaller64BitCheck) {
#	Write-LogEntry -Value "The Classic Teams Machine-Wide Installer (64-bit) was not detected. This is the desired state." -Severity 1
#}

#--------------------------------------------------

if (-not $ClassicTeamsMachineWideInstaller64BitCheck) {
	Write-LogEntry -Value "The Classic Teams Machine-Wide Installer (64-bit) was not detected. This is the desired state. All detection conditions were met successfully." -Severity 1
	Write-Output "The Classic Teams Machine-Wide Installer (64-bit) was not detected. This is the desired state. All detection conditions were met successfully."
	exit 0
} else {
	if ($ClassicTeamsMachineWideInstaller64BitCheck) {
		Write-LogEntry -Value "The Classic Teams Machine-Wide Installer (64-bit) was detected. This is not the desired state." -Severity 2
	}
	Write-LogEntry -Value "One or more detections conditions have not been met. Detection has failed." -Severity 3
	#Write-Output "One or more detections conditions have not been met. Detection has failed."
	exit 1
}