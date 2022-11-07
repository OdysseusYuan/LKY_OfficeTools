Function Execute-SaraCMD($exeFilePath)
{
	$processInfo = new-Object System.Diagnostics.ProcessStartInfo($exeFilePath);
	$processInfo.Arguments = "-s OfficeScrubScenario -AcceptEULA";
	$processInfo.CreateNoWindow = $true;     # Runs SaRA CMD in a quite mode, $false to show its console
	$processInfo.UseShellExecute = $false;
	$processInfo.RedirectStandardOutput = $true;

	$process = [System.Diagnostics.Process]::Start($processInfo);
	$process.StandardOutput.ReadToEnd();     # Displays SaRA CMD's output in the PowerShell window, Comment if not needed
	$process.WaitForExit();
	$success = $process.HasExited -and $process.ExitCode -eq 0;
	$process.Dispose();

	return $success;                         # $true if the scenario's execution PASSED otherwise $false 
}

# Execute SaRA CMD by passing its exe file path
Execute-SaraCMD -exeFilePath "FullPath\SaRACmd.exe";