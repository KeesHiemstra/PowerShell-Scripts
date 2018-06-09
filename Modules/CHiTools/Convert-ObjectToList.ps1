#region Convert-ObjectToList

<#
.SYNOPSIS
    Rotate an object as a list.

.DESCRIPTION
    This function works like Format-Table, but the output is an object that can be used by for example ConvertTo-HTML.

.EXAMPLE
    Get-Service | Where-Object { $_.Name -like 'BITS' } | Convert-ObjectToList

    >>---
    Property            Value1                                 
    --------            ------                                 
    Name                BITS                                   
    RequiredServices    RpcSs; EventSystem                     
    CanPauseAndContinue False                                  
    CanShutdown         False                                  
    CanStop             True                                   
    DisplayName         Background Intelligent Transfer Service
    DependentServices                                          
    MachineName         .                                      
    ServiceName         BITS                                   
    ServicesDependedOn  RpcSs; EventSystem                     
    ServiceHandle                                              
    Status              Running                                
    ServiceType         Win32ShareProcess                      
    StartType           Automatic                              
    Site                                                       
    Container                                                  

.EXAMPLE
    Get-Service | Where-Object { $_.Name -like 'Lan*' } | Convert-ObjectToList 

    >>The first column will be the list of all properties, the following columns (Value1, Value2) will show the content.
    Property            Value1            Value2                         
    --------            ------            ------                         
    Name                LanmanServer      LanmanWorkstation              
    RequiredServices    SamSS; Srv        NSI; MRxSmb20; MRxSmb10; Bowser
    CanPauseAndContinue True              True                           
    CanShutdown         False             False                          
    CanStop             True              True                           
    DisplayName         Server            Workstation                    
    DependentServices   Browser           SessionEnv; Netlogon; Browser  
    MachineName         .                 .                              
    ServiceName         LanmanServer      LanmanWorkstation              
    ServicesDependedOn  SamSS; Srv        NSI; MRxSmb20; MRxSmb10; Bowser
    ServiceHandle                                                        
    Status              Running           Running                        
    ServiceType         Win32ShareProcess Win32ShareProcess              
    StartType           Automatic         Automatic                      
    Site                                                                 
    Container                                                            

.EXAMPLE
    Get-Service | Where-Object { $_.Name -like 'Lan*' } | Convert-ObjectToList | ConvertTo-Html

    >> Will convert the data in an HTML table where the properties are listes as rows.

.EXAMPLE
    Get-Service | Convert-ObjectToList -MaximumSize 3 -ExcludeProperty @('RequiredServices', 'DependentServices','ServicesDependedOn', 'ServiceHandle', 'Site', 'Container')

    >> Maximizes the number of columns to 4 (first column contains the property name) and exclude unwanted properties.
    Property            Value1                       Value2                 Value3                           
    --------            ------                       ------                 ------                           
    Name                AdobeARMservice              AeLookupSvc            ALG                              
    CanPauseAndContinue False                        False                  False                            
    CanShutdown         False                        False                  False                            
    CanStop             True                         False                  False                            
    DisplayName         Adobe Acrobat Update Service Application Experience Application Layer Gateway Service
    MachineName         .                            .                      .                                
    ServiceName         AdobeARMservice              AeLookupSvc            ALG                              
    Status              Running                      Stopped                Stopped                          
    ServiceType         Win32OwnProcess              Win32ShareProcess      Win32OwnProcess                  
    StartType           Automatic                    Manual                 Manual                           
.INPUTS
    [System.Management.Automation.PSObject]
.OUTPUTS
    [System.Management.Automation.PSObject]
.NOTES
    === Version history
    Version 1.00 (2017-01-06, Kees Hiemstra)
    - Initial version
.COMPONENT
    Tools
.ROLE

.FUNCTIONALITY
    Conversion
#>
function Convert-ObjectToList
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$false,
                  PositionalBinding=$false,
                  ConfirmImpact='Low')]
    [OutputType([System.Management.Automation.PSObject])]
    Param
    (
        #Object to be turned into a list
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [System.Management.Automation.PSObject[]]
        $InputObject,

        #Properties from InputObject to exclude from output
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
        [string[]]
        $ExcludeProperty,

        #Maximum number of InputObjects in output
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$false,
                   ValueFromPipelineByPropertyName=$false,
                   ValueFromRemainingArguments=$false,
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        [int]
        $MaximumSize
    )

    Begin
    {
        $OutputObject = @(New-Object -TypeName PSObject)
        $Count = 0
    }
    Process
    {

        foreach ( $Object in $Input )
        {
            $Count++
            if ( $MyInvocation.BoundParameters.ContainsKey('MaximumSize') -and $Count -gt $MaximumSize ) { break }

            foreach ( $Property in $Object.PSObject.Properties )
            {
                if ( $Property.Name -in @('PropertyNames', 'PropertyCount' ) + $ExcludeProperty ) { continue }

                if ( $OutputObject.Property -notcontains $Property.Name)
                {
                    $Obj = New-Object -TypeName PSObject
                    Add-Member -InputObject $Obj -NotePropertyName 'Property' -NotePropertyValue $Property.Name -Force
                    $OutputObject += $Obj
                }
                $Obj = $OutputObject | Where-Object { $_.Property -eq $Property.Name }

                Add-Member -InputObject $Obj -NotePropertyName "Value$Count" -NotePropertyValue ($Property.Value -join '; ').Trim()
            }
        }
    }
    End
    {
        Write-Output $OutputObject
    }
}

#endregion
