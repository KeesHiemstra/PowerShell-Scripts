$Properties = [ordered]@{Path=$null;LogLength=$null;ConsoleOutput=$false}

Get-ChildItem -Path $PSScriptRoot\*.ps1 -Recurse | ForEach-Object { . $_.FullName }