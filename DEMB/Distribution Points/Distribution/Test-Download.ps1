<#
    Test-Download.ps1

    Handle download.

    --- Version history
	Version 1.01 (2016-12-19, Kees Hiemstra)
	- Added JobState 'Error' with description.
    Version 1.00 (2016-05-03, Kees Hiemstra)
    - Initial version.
#>

. "$($PSScriptRoot)\CommonEnvironment.ps1"

$ScriptName = "Test-Download"

if ( $Status.Download.Count -eq 0 )
{
    #Write-Log -Message "No downloads"
    Stop-Script -Silent
}


foreach ($Download in $Status.Download)
{
    $Bits = Get-BitsTransfer -JobId $Download.JobId -ErrorAction SilentlyContinue
    if ( $Bits -eq $null )
    {
        Write-Log -Message "$($Download.Description) doesn't exits anylonger"
        $Download = $null
        continue
    }

    #Write-Log -Message "JobState for $($Bits.Description) is $($Bits.JobState)"
    switch ($Bits.JobState)
    {
        'Suspended'
            {
                if ( (Test-InOfficeHours) )
                {
                }
                else
                {
                    $null = Resume-BitsTransfer -BitsJob $Download.JobId -Asynchronous
                    Write-Log -Message ("Resume at {0:P} for '$($Bits.Description)'" -f ($Bits.BytesTransferred / $Bits.BytesTotal))
                }
            }
        'Transferring'
            {
                if ( (Test-InOfficeHours) )
                {
                    $Download = Suspend-BitsTransfer -BitsJob $Bits.JobId
                    Write-Log -Message ("Suspended at {0:P} for '$($Bits.Description)'" -f ($Bits.BytesTransferred / $Bits.BytesTotal))
                }
                else
                {
                    Write-Log -Message ("Download at {0:P} for '$($Bits.Description)'" -f ($Bits.BytesTransferred / $Bits.BytesTotal))
                }
            }
        'Transferred'
            {
                $Download.TransferCompletionTime = $Bits.TransferCompletionTime
                Complete-BitsTransfer -BitsJob $Download.JobId
                Write-Log -Message "Download of '$($Download.Description)' complete at $((Get-CorrectedDate -Date $Download.TransferCompletionTime).ToString('yyyy-MM-dd HH:mm:ss'))"
                Write-Log -Message "Duration of download $((New-TimeSpan -Start (Get-CorrectedDate -Date $Download.CreationTime) -End (Get-CorrectedDate -Date $Download.TransferCompletionTime)).ToString())"
                $DeletedJobs += "$($Download.JobId);"
                $SendCompleted = $true
            }
        'Connecting'
            {
                Write-Log -Message "$($Bits.Description) is connecting"
            }
        'Error'
            {
                Write-Log -Message "Error JobState: $($Bits.ErrorDescription)"
            }
        default
            {
                Write-Log -Message "Unknown JobState: $($Bits.JobState)"
            }
    }#switch
}# foreach download

$Status.Download = $Status.Download | Where-Object { $_.JobId -notin ($DeletedJobs -split ";") }
Update-Status

Stop-Script -Silent
