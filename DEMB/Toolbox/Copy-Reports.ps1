$ReportFolder = "B:\Reports"

Copy-Item -Path "\\DEMBMCAPS032SQ2.corp.demb.com\H$\ITAM\Usr\Report\ITAM19\Customer\Computer_in_contract.xlsx" -Destination $ReportFolder -Force
Copy-Item -Path "\\DEMBMCAPS032SQ2.corp.demb.com\H$\ITAM\Usr\Report\ITAM19\Customer\Computer_in_stock.xlsx" -Destination $ReportFolder -Force
Copy-Item -Path "\\DEMBMCAPS032SQ2.corp.demb.com\H$\ITAM\Usr\Report\ITAM19\Customer\Computer_in_contract_changes.xlsx" -Destination $ReportFolder -Force
Copy-Item -Path "\\DEMBMCAPS032SQ2.corp.demb.com\H$\ITAM\Usr\Report\ITAM19\HP\Computer_not_in_contract.xlsx" -Destination $ReportFolder -Force
Copy-Item -Path "\\DEMBMCAPS032SQ2.corp.demb.com\H$\ITAM\Usr\Report\ITAM19\Customer\Computer_in_contract_OpCo_Difference.xlsx" -Destination $ReportFolder -Force
Copy-Item -Path "\\DEMBMCAPS032SQ2.corp.demb.com\H$\ITAM\Usr\Report\ITAM19\Customer\Logon_history.xlsx" -Destination $ReportFolder -Force
Copy-Item -Path "\\DEMBMCAPS032SQ2.corp.demb.com\H$\ITAM\Usr\Report\ITAM19\Customer\Application_count.xlsx" -Destination $ReportFolder -Force

Get-ChildItem $ReportFolder | ForEach-Object { $_.Attributes = 'ReadOnly' }