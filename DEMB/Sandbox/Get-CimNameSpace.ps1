#requires -version 4.0

Function Get-CimNameSpace {
[cmdletbinding(DefaultParameterSetName = 'computer')]
Param(
    [Parameter(Position=0)]
    [ValidateNotNullorEmpty()]
    [string]$Namespace = "Root",
    [Parameter(ParameterSetName = 'computer')]
    [ValidateNotNullorEmpty()]
    [string]$Computername = $env:computername,
    [Parameter(ParameterSetName = 'computer')]
    [PSCredential]$Credential,
    [switch]$Recurse,
    [Parameter(ParameterSetName = 'session')]
    [CimSession]$CimSession
    )

Write-Verbose "[$((Get-Date).timeOfDay)] Starting $($myinvocation.MyCommand)"    

#private function to do the actual enumeration
Function _EnumNamespace {
[cmdletbinding(DefaultParameterSetName = 'computer')]
Param(
    [Parameter(Position=0)]
    [ValidateNotNullorEmpty()]
    [string]$Namespace = "Root",
    [Parameter(ParameterSetName = 'computer')]
    [ValidateNotNullorEmpty()]
    [string]$Computername = $env:computername,
    [Parameter(ParameterSetName = 'computer')]
    [PSCredential]$Credential,
    [switch]$Recurse,
    [Parameter(ParameterSetName = 'session')]
    [CimSession]$CimSession
    )
    Write-Verbose "[$((Get-Date).timeOfDay)] Starting $($myinvocation.MyCommand)"    

#define parameter hashtable for Get-CimInstance
$cimInst = @{
    Namespace = $Namespace
    ClassName = '__Namespace'
}

if ($pscmdlet.ParameterSetName -eq 'computer' -AND $Computername -ne $env:computername) {
    Write-Verbose "[$((Get-Date).timeOfDay)] Creating a CIMSession to $Computername" 
    $cimParams=@{
        Erroraction = 'stop'
        Computername = $computername
    }   
    if ($Credential) {
        Write-Verbose "[$((Get-Date).timeOfDay)] Using alternate credential for $($credential.username)"    
        $cimParams.Add('Credential',$Credential)
    }
    #create a CIM Session
    Try {
        $CimSession = New-CimSession @cimParams
        #add to CimInstance hashtable
        $cimInst.Add("cimsession",$CimSession)
        #save a reference to this session so it can be removed later
        $script:cs = $cimSession
    }
    Catch {
        Throw $_
    }
}
elseif ($pscmdlet.ParameterSetName -eq 'session') {
    Write-Verbose "[$((Get-Date).timeOfDay)] Using CimSession" 
    $cimInst.Add("cimsession",$CimSession)
    #save a reference to this session so it can be removed later
    $script:cs = $cimSession
}

Write-Verbose "[$((Get-Date).timeOfDay)] Getting Namespaces"    

$nspaces = Get-CimInstance @cimInst

Write-Verbose "[$((Get-Date).timeOfDay)] Found $($nspaces.count) namespaces under $namespace"
if ($nspaces.count -gt 0) {
    foreach ($nspace in $nspaces) {
        Write-Verbose "[$((Get-Date).timeOfDay)] Processing $($nspace.name)"
        $child = Join-Path -Path $Namespace -ChildPath $nspace.Name
        [pscustomobject]@{
            Computername = $nspace.cimsystemproperties.servername
            Namespace = $child
        }
        if ($recurse -and $CimSession) {
            Write-Verbose "[$((Get-Date).timeOfDay)] Recursing and re-using cimsession" 
            _EnumNamespace -namespace $child -Recurse -CimSession $cimSession
        }
        elseif ($recurse) {
            Write-Verbose "[$((Get-Date).timeOfDay)] Recursing" 
            _EnumNamespace -namespace $child -Recurse
        }
    } #foreach
} #if
    Write-Verbose "[$((Get-Date).timeOfDay)] Ending $($myinvocation.MyCommand)"
} #close _EnumNamespace function

#initiate the process
_EnumNamespace @PSBoundParameters

#clean-up
if ($script:cs) {
    Write-Verbose "[$((Get-Date).timeOfDay)] Removing CIMSession"    
    $script:cs | Remove-CimSession
}

Write-Verbose "[$((Get-Date).timeOfDay)] Ending $($myinvocation.MyCommand)"    
} #close main function


break
### Examples
Get-CimNameSpace

Get-CimNameSpace -Namespace root\Microsoft -Recurse

