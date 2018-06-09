<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.149
	 Created on:   	15/3/2018 11:29
	 Created by:   	Lex van der Horst
	 Organization: 	TechEdge
	 Filename:     	Diskspace-Report.ps1
	===========================================================================
	.DESCRIPTION
		Calculates the free space of each drive and write the results to HDReport.html
#>
 
$Computers = "HPNLDev07"
$Path = "C:\Temp\Diskspace-Report.html"
$Title = "HD Report to HTML"
 
# Embed the stylesheet in HTML header 
$Head = @" 
<mce:style><!-- 
mce:0 
--></mce:style><!-- 
mce:0 
--></style> 
<Title>$Title</Title> 
<br> 
"@
 
# Define array for html fragments 
$fragments = @()
# Get the disk data 
$Data = Get-WmiObject -Class Win32_logicaldisk -filter "drivetype=3" -computer $Computers
# Group data by computername 
$Groups = $Data | Group-Object -Property SystemName
 
# Graph character 
[string]$g = [char]9608
 
# Create html fragments for each computer and iterate through each group object 
ForEach ($Computer in $Groups)
{
	
	$Fragments += "<H2>$($Computer.Name)</H2>"
	
	# Define a collection of drives from the group object 
	$Drives = $Computer.group
	
	# Create HTML fragment 
	$html = $drives | Select @{ Name = "Drive"; Expression = { $_.DeviceID } },
							 @{ Name = "Size (GB)"; Expression = { $_.Size/1GB -as [int] } },
							 @{ Name = "Used (GB)"; Expression = { "{0:N2}" -f (($_.Size - $_.Freespace)/1GB) } },
							 @{ Name = "Free (GB)"; Expression = { "{0:N2}" -f ($_.FreeSpace/1GB) } },
							 @{
		Name	  = "Usage"; Expression = {
			$UsedPer = (($_.Size - $_.Freespace)/$_.Size) * 100
			$UsedGraph = $g * ($UsedPer/2)
			$FreeGraph = $g * ((100 - $UsedPer)/2)
			# Using place holders for the < and > characters 
			"xopenFont color=Redxclose{0}xopen/FontxclosexopenFont Color=Bluexclose{1}xopen/fontxclose" -f $usedGraph, $FreeGraph
		}
	} | ConvertTo-Html -Fragment
	
	# Replace the tag place holders 
	$html = $html -replace "xopen", "<"
	$html = $html -replace "xclose", ">"
	# Add to fragments 
	$Fragments += $html
	# Insert a return between each computer 
	$Fragments += "<br>"
	
} # Foreach computer 
 
# Add a footer to the HD Report
$Footer = ("<br><I>Diskspace Report run on {0} by {1}\{2}<I>" -f (Get-Date -DisplayHint date), $Env:UserDomain, $Env:Username)
$Fragments += $Footer
 
# Write results to HD Report
ConvertTo-Html -Head $Head -Body $Fragments | Out-File -FilePath $Path
