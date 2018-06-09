Get-ADGroup -Filter * -Properties SAMAccountName, Description, Info, Manager, WhenCreated, WhenChanged, Members, MemberOf, DistinguishedName |
    Select-Object SAMAccountName,
        Description,
        Info,
        @{n='Manager'; e={ $_.Manager | Convert-DistinguishedNameToName }},
        @{n='Members'; e={ $_.Members.Count }},
        WhenCreated,
        WhenChanged,
        @{n='MemberOf'; e={ $_.MemberOf | Convert-DistinguishedNameToName -join '; ' }},
        DistinguishedName |
    ConvertTo-Csv | Clip