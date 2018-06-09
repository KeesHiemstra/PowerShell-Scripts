# Only list output fields with content

function Remove-EmptyProperty  {
    param (
        [Parameter(Mandatory,ValueFromPipeline)]
            $InputObject,
            
            [Switch]
            $AsHashTable
    )


    begin
    {
      $props = @()
  
    }
    
    process {
        if ($props.Count -eq 0)
        {
          $props = $InputObject | 
            Get-Member -MemberType *Property | 
            Select-Object -ExpandProperty Name
        }
    
        $notEmpty = $props | Where-Object { 
          !($InputObject.$_ -eq $null -or
             $InputObject.$_ -eq '' -or 
             $InputObject.$_.Count -eq 0) |
          Sort-Object
        
        }
        
        if ($AsHashTable)
        {
          $notEmpty | 
            ForEach-Object { 
                $h = [Ordered]@{}} { 
                    $h.$_ = $InputObject.$_ 
                    } { 
                    $h 
                    }
        }
        else
        {
          $InputObject | 
            Select-Object -Property $notEmpty
        }
    }
}