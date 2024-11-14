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

# The Name of the package. This is an exact match.
$PackageName = "Microsoft.OutlookForWindows"
$AppName = "New Outlook"
$LogFileName = "NewOutlookRemoval.log"

#--------------------------------------------------

Write-LogEntry -Value "Starting $AppName uninstall process" -Severity 1

#--------------------------------------------------

# Starting App Removal
try {
	#Running office installer
	Write-LogEntry -Value "Starting $AppName uninstall with Remove-AppxProvisionedPackage method" -Severity 1
	Remove-AppxProvisionedPackage -AllUsers -Online -PackageName (Get-AppxPackage $PackageName).PackageFullName
	Write-LogEntry -Value "Finished $AppName uninstall" -Severity 1
}
catch [System.Exception] {
	Write-LogEntry -Value  "Error running the $AppName uninstall. Error Message: $($_.Exception.Message)" -Severity 3
	exit 1
}

Write-LogEntry -Value "Finished $AppName uninstall process" -Severity 1