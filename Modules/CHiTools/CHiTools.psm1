﻿Get-ChildItem -Path $PSScriptRoot\*.ps1 -Recurse | ForEach-Object { . $_.FullName }