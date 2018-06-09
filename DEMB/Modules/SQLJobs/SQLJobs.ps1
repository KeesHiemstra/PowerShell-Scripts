Function Start-SQLJob
{
<#
.Synopsis
    Start an SQL job.
.DESCRIPTION
    Start the named SQL job at the the SQL Agent of the mentioned SQL server. The outcome of 
    this cmdlet inclused a simplyfied representation of the SQL job status.

    If the parameter StartAtJobStep is not applied, the cmdlet will query the job and determine
    the default starting step.

    After the job is started, the cmdlet will poll the SQL server at least one time to see if the
    job is really running. This polling mechanism can be regulated with the parameters MaxPolling
    and PollingPause.
    
    The server instance supports the default installed instance and named instances. The jobname
    must match the name of the SQL job.

    The local server (with default instance) can be addressed by '(Local)' or '.'.

    By default the authentication is done by the Windows authentication, but will switch to
    SQL authentication if Login and password is provided. 
.OUTPUTS
    Object():
    - SQLInstance
    - JobName
    - IsEnabled
    - HasSchedule
    - CurrentRunStatus
    - LastRunDate
    - LastRunOutcome
    - SimplifiedResult

    Simplyfied
    result           Description
    ------           --------------------------------------------------------------
    Running          The SQL job is currently running.
    Succeeded        The SQL job finished successful.
    Failed           The SQL job finished, but reported a failure or was cancelled.
    Starting failed  The SQL job is unsuccessfully started.
.NOTES
    Author         : Kees Hiemstra (Kees.Hiemstra@hp.com)
    PSVerion       : 4.0 (3.0 should work)
    Assemblies     : Microsoft.SqlServer.SMO and Microsoft.SqlServer.Management.Common
                     Tested with SQL Server 2008, SQL Server 2008 R2

    Version history:

    Version 1.01 (2014-11-16, Kees Hiemstra)
    - Bug fix: Assemblies where not loaded properly.
    Version 1.00 (2014-11-10, Kees Hiemstra)
    - Inital version
.EXAMPLE
    Start-SQLJob -SQLInstance "DEMBMCAPS032SQ2\Prod_2" -JobName "ITAM - ! Import AssetCenter data"


    SQLInstance      : DEMBMCAPS032SQ2\Prod_2
    JobName          : ITAM - ! Import AssetCenter data
    IsEnabled        : True
    HasSchedule      : False
    CurrentRunStatus : Executing
    LastRunDate      : 2014-11-15 05:26:04
    LastRunOutcome   : Cancelled
    SimplifiedResult : Running
.EXAMPLE
    Stop-SQLJob -SQLInstance "(Local)" -JobName "MaintenancePlan.Transaction log backup" -Login "sa" -Password "It$MeUno"

#>
    [CmdletBinding()]
    #[OutputType(Object)]

    param
    (
        #The server name is used when only the named instanced is installed.
        #Servers with named (a) instance(s) are written like <Server name>\<Instance name>
        #e.g. DEMBMCAPS032SQ2\Prod_2
        [Parameter(Mandatory=$true,
                   Position=0,
                   HelpMessage = "<Server name> or <Sever name>\<Instance name>",
                   ParameterSetName = "WindowsAuthentication")]
        [Parameter(Mandatory=$true,
                   Position=0,
                   HelpMessage = "<Server name> or <Sever name>\<Instance name>",
                   ParameterSetName="SQLAuthentication")]
        [Alias('ComputerName')]
        [string]$SQLInstance,


        #Job name that needs to be stopped. The name must match the name in the SQL agent.
        [Parameter(Mandatory=$true,
                   Position=1,
                   HelpMessage = "Name of the SQL job (wild cards are allowed)",
                   ParameterSetName = "WindowsAuthentication")]
        [Parameter(Mandatory=$true,
                   Position=1,
                   HelpMessage = "Name of the SQL job (wild cards are allowed)",
                   ParameterSetName = "SQLAuthentication")]
        [string]$JobName,

        #A job can have more then 1 step. This parameter indicates at which step the job should
        #begin. The name of the step must match the name in the job.
        #If this parameter is not provided, the job will start at the default step provided in 
        #the job definition.
        [Parameter(Mandatory = $false,
                   Position=2)]
        [string]$StartAtJobStep,

        #Maximum number of polling attempts. If the maximum number is reached and the job hasn't
        #started yet, the SimplifiedResult field will indicate 'Starting failed'. If the job
        #started within the maximum number of polling attempts, the SimpliFiedResult field will
        #indicate 'Running'.

        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 15)]
        [int]$MaxPolling = 1,

        #Time between polling steps measured in seconds.
        [ValidateRange(1, 10)]
        [int]$PollingPause = 1,

        #Provide the login name if the access need to be granted through SQL authentication.
        #The login name always goes to with password.
        [Parameter(Mandatory=$true,
                   HelpMessage="Logon name for SQL authentication.",
                   ParameterSetName = "SQLAuthentication")]
        [string]$Login,

        #Povide the password if the access need to be granted through SQL authentication.
        #The password always goes to with login name.
        [Parameter(Mandatory=$true,
                   HelpMessage="Password for SQL authentication.",
                   ParameterSetName = "SQLAuthentication")]
        [string]$Password
    )

    begin
    {
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo")
        [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.Management.Common')
        [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
    }

    process
    {
        $Connection = new-object Microsoft.SqlServer.Management.Common.ServerConnection

        if ( -not (([string]::IsNullOrWhiteSpace($Login)) -and ([string]::IsNullOrWhiteSpace($Password))) )
        {
            Write-Verbose "Use SQL authentication"
            $Connection.LoginSecure = $false
            $Connection.Login = $Login
            $Connection.Password = $Password
        } #SQL authentication

        $Connection.ServerInstance = $SQLInstance
        $Server = New-Object "Microsoft.SqlServer.Management.Smo.Server" $Connection

        $JobStatus = Get-SQLJobStatus -SQLInstance $SQLInstance -JobName $JobName -Login $Login -Password $Password
        if ($JobStatus -eq $null) { break }

        If ($JobStatus.SimplifiedResult -eq "Running")
        {
            Write-Error ("The job '{0}' already started on '{1}'" -f $JobName, $SQLInstance)
            $JobStatus.SimplifiedResult = "Already running"
            Write-Output $JobStatus
            break
        }

        $LastStartTime = $JobStatus.LastRunDate
        Write-Verbose ("Last start time: {0}" -f $LastStartTime)

        $Job = $Server.JobServer.Jobs | Where-Object { $_.Name -eq $JobName }

        if (-not [string]::IsNullOrWhiteSpace($StartAtJobStep))
        {
            #Get default starting step
            $StartAtJobStepName = ($Job.JobSteps | Where-Object { $_.ID -eq $Job.StartStepID }).Name
    
        }
        else
        {
            $StartAtJobStepName = $StartJobStep
        }

        Write-Verbose ("Start job at step: {0}" -f $StartAtJobStepName)

        $Job.Start($StartAtJobStepName)

        #It takes time to start the job, this delay will give the server the oppertunity to do so
        do
        {
            sleep $PollingPause
            $MaxPolling--

            $JobStatus = Get-SQLJobStatus -SQLInstance $SQLInstance -JobName $JobName -Login $Login -Password $Password
        } until ( ($MaxPolling -le 0) -or ($JobStatus.LastRunTime -ne $LastStartTime) )

        if ($JobStatus.SimplifiedResult -ne "Running") { $JobStatus.SimplifiedResult = "Starting failed"}

        Write-Output $JobStatus
    }#Process

    end {}
}

Function Stop-SQLJob
{
<#
.Synopsis
    Stop an SQL job.
.DESCRIPTION
    Stop the named SQL job at the the SQL Agent of the mentioned SQL server. The outcome of 
    this cmdlet inclused a simplyfied representation of the SQL job status.
    
    The server instance supports the default installed instance and named instances. The jobname
    must match the name of the SQL job.

    The local server (with default instance) can be addressed by '(Local)' or '.'.

    By default the authentication is done by the Windows authentication, but will switch to
    SQL authentication if Login and password is provided. 
.OUTPUTS
    Object():
    - SQLInstance
    - JobName
    - IsEnabled
    - HasSchedule
    - CurrentRunStatus
    - LastRunDate
    - LastRunOutcome
    - SimplifiedResult

    Simplyfied
    result           Description
    ------           --------------------------------------------------------------
    Running          The SQL job is currently running.
    Succeeded        The SQL job finished successful.
    Failed           The SQL job finished, but reported a failure or was cancelled.
.NOTES
    Author         : Kees Hiemstra (Kees.Hiemstra@hp.com)
    PSVerion       : 4.0 (3.0 should work)
    Assemblies     : Microsoft.SqlServer.SMO and Microsoft.SqlServer.Management.Common
                     Tested with SQL Server 2008, SQL Server 2008 R2

    Version history:

    Version 1.01 (2014-11-16, Kees Hiemstra)
    - Bug fix: Assemblies where not loaded properly.
    Version 1.00 (2014-11-10, Kees Hiemstra)
    - Inital version
.EXAMPLE
    Stop-SQLJob -SQLInstance "DEMBMCAPS032SQ2\Prod_2" -JobName "ITAM - ! Import AssetCenter data"


    SQLInstance      : DEMBMCAPS032SQ2\Prod_2
    JobName          : ITAM - ! Import AssetCenter data
    IsEnabled        : True
    HasSchedule      : False
    CurrentRunStatus : Cancelled
    LastRunDate      : 2014-11-15 05:26:04
    LastRunOutcome   : Cancelled
    SimplifiedResult : Failed
.EXAMPLE
    Stop-SQLJob -SQLInstance "(Local)" -JobName "MaintenancePlan.Transaction log backup" -Login "sa" -Password "It$MeUno"

#>
    [CmdletBinding()]
    [OutputType([string])]

    param
    (
        #The server name is used when only the named instanced is installed.
        #Servers with named (a) instance(s) are written like <Server name>\<Instance name>
        #e.g. DEMBMCAPS032SQ2\Prod_2
        [Parameter(Mandatory=$true,
                   Position=0,
                   HelpMessage = "<Server name> or <Sever name>\<Instance name>",
                   ParameterSetName = "WindowsAuthentication")]
        [Parameter(Mandatory=$true,
                   Position=0,
                   HelpMessage = "<Server name> or <Sever name>\<Instance name>",
                   ParameterSetName="SQLAuthentication")]
        [Alias('ComputerName')]
        [string]$SQLInstance,

        #Job name that needs to be stopped. The name must match the name in the SQL agent.
        [Parameter(Mandatory=$true,
                   Position=1,
                   HelpMessage = "Name of the SQL job (wild cards are allowed)",
                   ParameterSetName = "WindowsAuthentication")]
        [Parameter(Mandatory=$true,
                   Position=1,
                   HelpMessage = "Name of the SQL job (wild cards are allowed)",
                   ParameterSetName = "SQLAuthentication")]
        [string]$JobName,

        #Provide the login name if the access need to be granted through SQL authentication.
        #The login name always goes to with password.
        [Parameter(Mandatory=$true,
                   HelpMessage="Logon name for SQL authentication.",
                   ParameterSetName = "SQLAuthentication")]
        [string]$Login,

        #Povide the password if the access need to be granted through SQL authentication.
        #The password always goes to with login name.
        [Parameter(Mandatory=$true,
                   HelpMessage="Password for SQL authentication.",
                   ParameterSetName = "SQLAuthentication")]
        [string]$Password
    )

    begin
    {
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo")
        [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.Management.Common')
        [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
    }

    process
    {
        $Connection = new-object Microsoft.SqlServer.Management.Common.ServerConnection

        if ( -not (([string]::IsNullOrWhiteSpace($Login)) -and ([string]::IsNullOrWhiteSpace($Password))) )
        {
            Write-Verbose "Use SQL authentication"
            $Connection.LoginSecure = $false
            $Connection.Login = $Login
            $Connection.Password = $Password
        } #SQL authentication

        $Connection.ServerInstance = $SQLInstance
        $Server = New-Object "Microsoft.SqlServer.Management.Smo.Server" $Connection

        $Job = $Server.JobServer.Jobs | Where-Object { $_.Name -eq $JobName }

        $Job.Stop()
    }#Process

    end {}
}

Function Get-SQLJobStatus
{
<#
.Synopsis
    Get the current status of an SQL job.
.DESCRIPTION
    Get the current status of an SQL job in the the SQL Agent of the mentioned server. The 
    outcome of this cmdlet inclused a simplyfied representation of the SQL job status.
    
    The server instance supports the default installed instance and named instances. The jobname
    supports wildcards.

    The local server (with default instance) can be addressed by '(Local)' or '.'.

    By default the authentication is done by the Windows authentication, but will switch to
    SQL authentication if Login and password is provided. 
.OUTPUTS
    Object():
    - SQLInstance
    - JobName
    - IsEnabled
    - HasSchedule
    - CurrentRunStatus
    - LastRunDate
    - LastRunOutcome
    - SimplifiedResult

    Simplyfied
    result           Description
    ------           --------------------------------------------------------------
    Running          The SQL job is currently running.
    Succeeded        The SQL job finished successful.
    Failed           The SQL job finished, but reported a failure or was cancelled.
.NOTES
    Author         : Kees Hiemstra (Kees.Hiemstra@hp.com)
    PSVerion       : 4.0 (3.0 should work)
    Assemblies     : Microsoft.SqlServer.SMO and Microsoft.SqlServer.Management.Common
                     Tested with SQL Server 2008, SQL Server 2008 R2

    Version history:

    Version 1.01 (2014-11-16, Kees Hiemstra)
    - Bug fix: Assemblies where not loaded properly.
    Version 1.00 (2014-11-10, Kees Hiemstra)
    - Inital version
.EXAMPLE
    Get-SQLJobStatus -SQLInstance "DEMBMCAPS032SQ2\Prod_2" -JobName "ITAM - ! Import AssetCenter data"


    SQLInstance      : DEMBMCAPS032SQ2\Prod_2
    JobName          : ITAM - ! Import AssetCenter data
    IsEnabled        : True
    HasSchedule      : False
    CurrentRunStatus : Idle
    LastRunDate      : 2014-11-15 05:26:04
    LastRunOutcome   : Succeeded
    SimplifiedResult : Succeeded
.EXAMPLE
    Get-SQLJobStatus -SQLInstance "DEMBMCAPS032SQ2\Prod_2" -JobName "ITAM - *" | Format-Table SimpliFiedResult, JobName -AutoSize


SimplifiedResult JobName                                               
---------------- -------                                               
Succeeded        ITAM - ! Import AssetCenter data                      
Succeeded        ITAM - Clean audits                                   
Succeeded        ITAM - Collect SCCM online devices                    
Succeeded        ITAM - Export IMACD                                   
Succeeded        ITAM - Export online devices                          
Succeeded        ITAM - Import AD (Diff)                               
Succeeded        ITAM - Import AD (Full)                               
Succeeded        ITAM - Import AntiVirus data                          
Succeeded        ITAM - Import audits                                  
Succeeded        ITAM - Import OSReport                                
Succeeded        ITAM - Invoice (don't start!)                         
Succeeded        ITAM - Monitoring database changes                    
Succeeded        ITAM - Monitoring imports                             
Succeeded        ITAM - Radia cleanup and Software processing/reporting
Succeeded        ITAM - Reporting general                              
Succeeded        ITAM - Reporting refresh                              
Succeeded        ITAM - SCCM Synchronization                           

.EXAMPLE
    Get-SQLJobStatus -SQLInstance "(Local)" -JobName * -Login "sa" -Password "It$MeUno" | Format-Table JobName, SimplifiedResult -AutoSize


    JobName                                SimplifiedResult
    -------                                ----------------
    MaintenancePlan.Full database backup   Failed          
    MaintenancePlan.Transaction log backup Failed          
    syspolicy_purge_history                Succeeded       

#>
    [CmdletBinding()]
    [OutputType([Object])]

    param
    (
        #The server name is used when only the named instanced is installed.
        #Servers with named (a) instance(s) are written like <Server name>\<Instance name>
        #e.g. DEMBMCAPS032SQ2\Prod_2
        [Parameter(Mandatory=$true,
                   Position=0,
                   HelpMessage = "<Server name> or <Sever name>\<Instance name>",
                   ParameterSetName = "WindowsAuthentication")]
        [Parameter(Mandatory=$true,
                   Position=0,
                   HelpMessage = "<Server name> or <Sever name>\<Instance name>",
                   ParameterSetName="SQLAuthentication")]
        [Alias('ComputerName')]
        [string]$SQLInstance,

        #Job name that needs to be stopped. The name must match the name in the SQL agent.
        [Parameter(Mandatory=$true,
                   Position=1,
                   HelpMessage = "Name of the SQL job (wild cards are allowed)",
                   ParameterSetName = "WindowsAuthentication")]
        [Parameter(Mandatory=$true,
                   Position=1,
                   HelpMessage = "Name of the SQL job (wild cards are allowed)",
                   ParameterSetName = "SQLAuthentication")]
        [string]$JobName,

        #Provide the login name if the access need to be granted through SQL authentication.
        #The login name always goes to with password.
        [Parameter(Mandatory=$true,
                   HelpMessage="Logon name for SQL authentication.",
                   ParameterSetName = "SQLAuthentication")]
        [string]$Login,

        #Povide the password if the access need to be granted through SQL authentication.
        #The password always goes to with login name.
        [Parameter(Mandatory=$true,
                   HelpMessage="Password for SQL authentication.",
                   ParameterSetName = "SQLAuthentication")]
        [string]$Password
    )

    begin
    {
        [void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SqlServer.ConnectionInfo")
        [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.Management.Common')
        [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO')
    }

    process
    {
        $Connection = new-object Microsoft.SqlServer.Management.Common.ServerConnection

        if ( -not (([string]::IsNullOrWhiteSpace($Login)) -and ([string]::IsNullOrWhiteSpace($Password))) )
        {
            Write-Verbose "Use SQL authentication"
            $Connection.LoginSecure = $false
            $Connection.Login = $Login
            $Connection.Password = $Password
        } #SQL authentication

        foreach ($SQLInstance1 in $SQLInstance)
        {
            $Connection.ServerInstance = $SQLInstance1
            $Server = New-Object "Microsoft.SqlServer.Management.Smo.Server" $Connection

            try
            {
                foreach ($JobName1 in $JobName)
                {

                    $Job = $Server.JobServer.Jobs | Where-Object {$_.Name -like $JobName1}

                    if($Job.Count -eq 0)
                    {
                        Write-Error ("Job '{0}' does not exist on '{1}'" -f $JobName1, $SQLInstance1)
                        break
                    }

                    foreach ($Job1 in $Job)
                    {
                        Write-Verbose ("Job name: {0}" -f $Job1.Name)
                        Write-Verbose ("Current run status / Last run outcome: {0}/{1}" -f $Job1.CurrentRunStatus, $Job1.LastRunOutcome)

                        if ($Job1.CurrentRunStatus -eq "Executing")
                        {
                            $SimplifiedResult = "Running"
                        }
                        elseif ($Job1.LastRunOutcome -eq "Succeeded")
                        {
                            $SimplifiedResult = "Succeeded"
                        }
                        else
                        {
                            $SimplifiedResult = "Failed"
                        }

                        Write-Verbose ("Simplified result: {0}" -f $SimplifiedResult)

                        $Properties = [ordered]@{'SQLInstance'      = $SQLInstance1;
                                                 'JobName'          = $Job1.Name;
                                                 'IsEnabled'        = $Job1.IsEnabled;
                                                 'HasSchedule'      = $Job1.HasSchedule;
                                                 'CurrentRunStatus' = $Job1.CurrentRunStatus;
                                                 'LastRunDate'      = $Job1.LastRundate;
                                                 'LastRunOutcome'   = $Job1.LastRunOutcome;
                                                 'SimplifiedResult' = $SimplifiedResult}
         
                        $Result = New-Object -TypeName PSObject -Property $Properties
                        Write-Output $Result

                    } #All jobs
                } #All job names
            } #try
            catch
            {
                Write-Error ("Error occurred on connecting to '{0}': " -f $SQLInstance1)
                Break
            }
        } #All instances
    }#Process

    end {}
    }
