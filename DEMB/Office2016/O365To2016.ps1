<#
    Create installation xml for Office 2016 for migration or initial installation.

    === Version history
    Version 1.91 (2017-03-19, Kees Hiemstra)
    - Bug fix: Avoid empty LanguageID to be added to the XML.
    Version 1.90 (2017-02-20, Kees Hiemstra)
    - Changed the Registry path to detect if Office 2016 is already installed.
    - Add Exit code if the script fails to install because Office is already installed.
    Version 1.80 (2017-01-21, Kees Hiemstra)
    - Include FORCEAPPSHUTDOWN = TRUE.
    Version 1.70 (2016-12-09, Kees Hiemstra)
    - Exclude the OneDrive and Groove in the XML file.
    Version 1.60 (2016-12-05, Kees Hiemstra)
    - Exclude the Groove in the XML file.
    Version 1.50 (2016-11-28, Kees Hiemstra)
    - Exclude the OneDrive in the XML file.
    Version 1.40 (2016-11-26, Kees Hiemstra)
    - Change the SourcePath and UpdatePath to \\corp.demb.com\DFS\Projects\HP_SOE\O365\2016\DC\
    - Remove the Version in the XML.
    Version 1.30 (2016-11-23, Kees Hiemstra)
    - Replace the Display Level from "Full" to "None"
    Version 1.20 (2016-11-21, Kees Hiemstra)
    - Replace the Deferred channel with version to avoid that setup.exe is using the newest version.
    Version 1.10 (2016-11-07, Kees Hiemstra)
    - Add 'Deferred' as channel to overwrite the default for Project and Visio (which is current).
    - Deal with old naming convesion of the computer name.
    Version 1.00 (2016-10-25, Kees Hiemstra)
    - Initial version.
#>

#region Definitions

$XMLPath = 'c:\hp\installO365.xml'

$Registry2Language = @{1026 = 'bg-BG';
    1029 = 'cs-CZ';
    1030 = 'da-DK';
    1031 = 'de-DE';
    1032 = 'el-GR';
    1033 = 'en-US';
    3082 = 'es-ES';
    1036 = 'fr-FR';
    1038 = 'hu-HU';
    1057 = 'id-ID';
    1040 = 'it-IT';
    1041 = 'ja-JP';
    1087 = 'kk-KZ';
    1042 = 'ko-KR';
    1063 = 'lt-LT';
    1062 = 'lv-LV';
    1086 = 'ms-MY';
    1044 = 'nb-NO';
    1043 = 'nl-NL';
    1045 = 'pl-PL';
    1046 = 'pt-BR';
    2070 = 'pt-PT';
    1048 = 'ro-RO';
    1049 = 'ru-RU';
    1051 = 'sk-SK';
    1053 = 'sv-SE';
    1054 = 'th-TH';
    1055 = 'tr-TR';
    1058 = 'uk-UA';
    1066 = 'vi-VN';
    2052 = 'zh-CN'
    }

[xml]$BaseXML = @"
<Configuration>
  <Add SourcePath="\\corp.demb.com\DFS\Projects\HP_SOE\O365\2016\DC\" OfficeClientEdition="32" Channel="Deferred" >
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="Groove" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" UpdatePath="\\corp.demb.com\DFS\Projects\HP_SOE\O365\2016\DC\" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE"/>
  <Display Level="None" AcceptEULA="TRUE" />
  <Logging Path="%windir%\Logs\Office365\" />
  <Property Name="AUTOACTIVATE" Value="1" />
</Configuration>
"@

[array]$OfficeLanguage = @('en-US')

#endregion

#region Pre checks

#Remove existing install xml
if ( (Test-Path -Path $XMLPath) )
{
    Remove-Item -Path $XMLPath -Force
}

#Check if Office 2016 is already installed
if ( (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Office\16.0\ClickToRun\ProductReleaseIDs\Active\O365ProPlusRetail') )
{
    #Office 2016 is already install, exit script
    exit 1604
}

#endregion

#Check if Office 2013 is installed
if ( (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\O365ProPlusRetail') )
{
    #Collect the languages
    foreach ( $RegPath in Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\O365ProPlusRetail' )
    {
        $RegValue = Get-ItemProperty $RegPath.PSPath
        if ( $RegValue.PSChildName -ne 'x-none' )
        {
            foreach ( $Item in ($RegValue.PSChildName -split ';') ) { if ( $Item -notin $OfficeLanguage ) { $OfficeLanguage += $Item } }
        }
    }

    #Check if Project 2013 is installed
    if ( (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\ProjectProRetail') )
    {
        #Add Project
        $NodeProduct = $BaseXML.Configuration.Add

        $NewProductNode = $BaseXML.CreateNode('element', 'Product', '')
        $NewProductAttr = $BaseXML.CreateAttribute('ID')
        $NewProductAttr.Value = 'ProjectProRetail'
        $NewProductNode.SetAttributeNode($NewProductAttr) | Out-Null
        $NodeProduct.AppendChild($NewProductNode) | Out-Null

        $NodeProduct = $BaseXML.Configuration.Add.ChildNodes

        $NodeLanguage = ($NodeProduct | Where-Object { $_.ID -eq 'ProjectProRetail' })
        $NewLanguageNode = $BaseXML.CreateNode('element', 'Language', '')
        $NewLanguageAttr = $BaseXML.CreateAttribute('ID')
        $NewLanguageAttr.Value = 'en-US'
        $NewLanguageNode.SetAttributeNode($NewLanguageAttr) | Out-Null
        $NodeLanguage.AppendChild($NewLanguageNode) | Out-Null
    }

    #Check if Visio 2013 is installed
    if ( (Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Office\15.0\ClickToRun\ProductReleaseIDs\Active\VisioProRetail') )
    {
        #Add Visio
        $NodeProduct = $BaseXML.Configuration.Add

        $NewProductNode = $BaseXML.CreateNode('element', 'Product', '')
        $NewProductAttr = $BaseXML.CreateAttribute('ID')
        $NewProductAttr.Value = 'VisioProRetail'
        $NewProductNode.SetAttributeNode($NewProductAttr) | Out-Null
        $NodeProduct.AppendChild($NewProductNode) | Out-Null

        $NodeProduct = $BaseXML.Configuration.Add.ChildNodes

        $NodeLanguage = ($NodeProduct | Where-Object { $_.ID -eq 'VisioProRetail' })
        $NewLanguageNode = $BaseXML.CreateNode('element', 'Language', '')
        $NewLanguageAttr = $BaseXML.CreateAttribute('ID')
        $NewLanguageAttr.Value = 'en-US'
        $NewLanguageNode.SetAttributeNode($NewLanguageAttr) | Out-Null
        $NodeLanguage.AppendChild($NewLanguageNode) | Out-Null
    }
}
else
{
    #Determine the languages based on the first 2 character of the computer name
    if ( ($env:COMPUTERNAME).Length -eq 14 )
    {
        #Old naming convention (DExxnnnnnnnnnn)
        $Computer = ($env:COMPUTERNAME).Substring(2,2)
    }
    else
    {
        $Computer = ($env:COMPUTERNAME).Substring(0,2)
    }

    switch ( $Computer )
    {
        'AT' { $Language = 'de-DE' }
        'BE' { $Language = 'nl-NL;fr-FR' }
        'BG' { $Language = 'bg-BG' }
        'BR' { $Language = 'pt-BR' }
        'CH' { $Language = 'de-DE' }
        'CN' { $Language = 'zh-CN' }
        'CZ' { $Language = 'cs-CZ' }
        'DE' { $Language = 'de-DE' }
        'DK' { $Language = 'da-DK' }
        'ES' { $Language = 'es-ES' }
        'FR' { $Language = 'fr-FR' }
        'GR' { $Language = 'el-GR' }
        'HU' { $Language = 'hu-HU' }
        'ID' { $Language = 'id-ID' }
        'IT' { $Language = 'it-IT' }
        'JP' { $Language = 'ja-JP' }
        'KR' { $Language = 'ko-KR' }
        'KZ' { $Language = 'kk-KZ' }
        'LT' { $Language = 'lt-LT' }
        'LV' { $Language = 'lv-LV' }
        'MA' { $Language = 'fr-FR' }
        'MY' { $Language = 'ms-MY' }
        'NO' { $Language = 'nb-NO' }
        'NL' { $Language = 'nl-NL' }
        'PL' { $Language = 'pl-PL' }
        'PT' { $Language = 'pt-PT' }
        'RO' { $Language = 'ro-RO' }
        'RU' { $Language = 'ru-RU' }
        'SE' { $Language = 'sv-SE' }
        'SK' { $Language = 'sk-SK' }
        'TH' { $Language = 'th-TH' }
        'TR' { $Language = 'tr-TR' }
        'UA' { $Language = 'uk-UA' }
        'VN' { $Language = 'vi-VN' }
    }

    foreach ( $Item in ($Language -split ';') ) { if ( $Item -notin $OfficeLanguage ) { $OfficeLanguage += $Item } }
}


#Process the collected languages (either Office 2013 installation or from the image)
$NodeProduct = $BaseXML.Configuration.Add.ChildNodes

$NodeLanguage = ($NodeProduct | Where-Object { $_.ID -eq 'O365ProPlusRetail' })

foreach ( $Item in ($OfficeLanguage | Where-Object { $_ -ne 'en-US' -and -not [string]::IsNullOrEmpty($_) }) )
{
    $NewLanguageNode = $BaseXML.CreateNode('element', 'Language', '')
    $NewLanguageAttr = $BaseXML.CreateAttribute('ID')
    $NewLanguageAttr.Value = $Item
    $NewLanguageNode.SetAttributeNode($NewLanguageAttr) | Out-Null
    $NodeLanguage.AppendChild($NewLanguageNode) | Out-Null
}

#Save the install xml
$BaseXML.Save($XMLPath)
