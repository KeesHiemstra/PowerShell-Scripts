#
# AnalyseEnvironment.ps1
#

## Capture the evironment variables on the Clipboard
Start-Process -FilePath "Cmd.exe" -ArgumentList "/C Set | Clip"

## Dump the capture in $SetDump
$SetDump = Get-Clipboard

$SysVars = @{}
foreach ($V in $SetDump) {
	$W = $V.Split('=')
	if($W -ne '') { 
		$SysVars.Add($W[0], $W[1])
		#"$($W[0]) ==> '$($W[1])'"
	}
}

$SysVars[0]