$ComputerName = 'FR5CG6082MJV'

Invoke-Command -ComputerName $ComputerName -ScriptBlock { Get-ChildItem -Path 'C:\Bonus calculation JDE 2017 France Band F and Below*.xlsx' -Recurse -ErrorAction SilentlyContinue }

break

Remove-CheckComputerOnline $ComputerName
Notepad "\\HPNLDev05.corp.demb.com\d$\Data\ExcelDoc\$ComputerName.txt"

break
Invoke-Command -ComputerName $ComputerName -ScriptBlock { Remove-Item -Path 'C:\Users\angele.lalungbarty\Desktop\ARCHIVES PERSO A LALUNG\ANGELE PROFESSIONNEL\Bonus calculation JDE 2017 France Band F and Below*.xlsx' }