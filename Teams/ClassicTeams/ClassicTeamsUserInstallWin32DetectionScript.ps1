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

# The path of a Classic Teams install in a user's home directory
$ClassicTeamsUserInstallPath = "C:\Users\*\AppData\Local\Microsoft\Teams\current\Teams.exe"

#--------------------------------------------------

$LogFileName = "TeamsSetup.log"
Write-LogEntry -Value "Start Classic Teams Per-User App Uninstall detection logic" -Severity 1

#--------------------------------------------------

#Commenting this out because the Teams Bootstrapper, in my testing, as of 10/8/24, doesn't seem to remove the Teams install from from every user, even though it is suppose to

# Check if the Classic Teams app is already in a user's home directory
$ClassicTeamsUserInstallCheck = Test-Path $ClassicTeamsUserInstallPath

#if (-not $ClassicTeamsUserInstallCheck) {
#	Write-LogEntry -Value "The Classic Teams app was not detected in any user's home directory. This is the desired state." -Severity 1
#}

#--------------------------------------------------

if (-not $ClassicTeamsUserInstallCheck) {
	Write-LogEntry -Value "The Classic Teams app was not detected in any user's home directory. This is the desired state. All detection conditions were met successfully." -Severity 1
	Write-Output "The Classic Teams app was not detected in any user's home directory. This is the desired state. All detection conditions were met successfully."
	exit 0
} else {
	if ($ClassicTeamsUserInstallCheck) {
		Write-LogEntry -Value "The Classic Teams app was detected in a user's home directory. This is not the desired state." -Severity 2
	}
	Write-LogEntry -Value "One or more detections conditions have not been met. Detection has failed." -Severity 3
	#Write-Output "One or more detections conditions have not been met. Detection has failed."
	exit 1
}