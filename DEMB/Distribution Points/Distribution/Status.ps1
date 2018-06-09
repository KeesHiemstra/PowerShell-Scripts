#$SOEToolsPath = (Get-WmiObject -Class Win32_Share | Where-Object { $_.Name -eq 'SOETools$' }).Path
#$SOEImagePath = (Get-WmiObject -Class Win32_Share | Where-Object { $_.Name -eq 'SOE_Image$' }).Path
#$SOEO365Path = (Get-WmiObject -Class Win32_Share | Where-Object { $_.Name -eq 'SOE$' }).Path

$Properties = [ordered]@{ComputerName=$env:COMPUTERNAME; 
                CountryCode="NL";
                Location="Utrecht";
                Description="Development Distribution Point";
                TimeZone="unknown";
                InOfficeHours=$false;
                StartStatus=$null;
                LastStatus=(Get-Date);
                Download=@();
                }

$Me = New-Object -TypeName PSObject -Property $Properties

$Me | Export-Clixml "$($PSScriptRoot)\Status.xml"
