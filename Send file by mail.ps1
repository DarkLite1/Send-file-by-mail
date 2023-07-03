#Requires -Version 5.1
#Requires -Modules Toolbox.EventLog, Toolbox.HTML

<# 
    .SYNOPSIS   
        Send specific files to users by mail.

    .DESCRIPTION
        All parameters are defined in the input file. Which file on which 
        computer, and where to send the e-mail too with the file in attachment.

    .PARAMETER ImportFile
        The file that contains all the parameters.

    .PARAMETER Tasks
        Each child represents one e-mail that will be sent.

    .PARAMETER Tasks.Mail
        To who the e-mail will be sent.

    .PARAMETER Tasks.ComputerName
        On which computer(s) the file(s) are stored.

    .PARAMETER Tasks.File
        Which file(s) to sent
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Send file by mail\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Function ConvertTo-UncPathHC {
            Param (
                [Parameter(Mandatory)]
                [String]$ComputerName,
                [Parameter(Mandatory)]
                [String]$LocalPath
            )
            '\\{0}\{1}' -f $ComputerName, $LocalPath.Replace(':', '$')
        }

        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion

        #region Test .json file properties
        $M = "Test .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        if (-not ($Tasks = $file.Tasks)) {
            throw "Input file '$ImportFile': Property 'Tasks' not found."
        }

        foreach ($task in $Tasks) {
            @('Mail', 'ComputerName', 'File') | 
            ForEach-Object {
                if (-not $task.$_) {
                    throw "Input file '$ImportFile': Property 'Tasks.$_' not found."
                }    
            }

            @('Header', 'To', 'Body', 'Priority', 'Subject') | 
            ForEach-Object {
                if (-not $task.Mail.$_) {
                    throw "Input file '$ImportFile': Property 'Tasks.Mail.$_' not found."
                }    
            }

            if (-not ($task.Mail.Priority -match '^high$|^low$|^normal$')) {
                throw "Input file '$ImportFile': Property 'Tasks.Mail.Priority' is not 'High', 'Low' or 'Normal'."
            }
            
            if (-not ($task.File -match ':')) {
                throw "Input file '$ImportFile': File '$($task.File)' is not supported, only local file paths are supported."
            }
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        #region Send mail for each task
        $i = 0

        foreach ($task in $Tasks) {
            $i++

            $mailParams = @{
                To        = $task.Mail.To
                Subject   = $task.Mail.Subject
                Header    = $task.Mail.Header
                Priority  = $task.Mail.Priority
                LogFolder = $logParams.LogFolder
                Save      = '{0} - {1} - Mail.html' -f $logFile, $i
            }

            #region Find files
            $files = @{
                Found    = @()
                NotFound = @()
            }

            foreach ($computer in $task.ComputerName) {
                foreach ($file in $task.File) {
                    $convertParams = @{
                        ComputerName = $computer
                        LocalPath    = $file
                    }
                    $path = ConvertTo-UncPathHC @convertParams

                    $testPathParams = @{
                        LiteralPath = $path 
                        PathType    = 'Leaf' 
                        ErrorAction = 'Stop'
                    }
                    if (-not (Test-Path @testPathParams)) {
                        $M = "File '$path' not found"
                        Write-Warning $M
                        Write-EventLog @EventErrorParams -Message $M

                        $files.NotFound += $path
                        Continue
                    }

                    $M = "Found file '$path'"
                    Write-Verbose $M
                    Write-EventLog @EventVerboseParams -Message $M

                    $files.Found += $path
                }
            }
            #endregion

            #region Create HTML list for files not found
            $filesNotFoundHtml = if ($files.NotFound) {
                $mailParams.Priority = 'High'
                $params = @{
                    Message = $files.NotFound
                    Header  = 'Files not found'
                }
                ConvertTo-HtmlListHC @params
            }
            #endregion

            #region Add found files in attachment to mail
            if ($files.Found) {
                $mailParams.Attachments = $files.Found
            }
            #endregion
            
            #region Mail cc
            if (
                $mailCc = $task.Mail.Cc | 
                Where-Object { $ScriptAdmin -notContains $_ }
            ) {
                $mailParams.Cc = $mailCc
            }
            #endregion

            #region Mail bcc
            $mailParams.Bcc = $ScriptAdmin

            if ($task.Bcc) {
                foreach (
                    $mailBcc in 
                    $task.Mail.Bcc | 
                    Where-Object { $ScriptAdmin -notContains $_ }
                ) {
                    $mailParams.Bcc += $mailBcc
                }
            }
            #endregion

            #region Send mail
            Write-Verbose 'Send mail'

            $mailParams.BodyHtml = "{0}{1}" -f 
            $task.Mail.Body, $filesNotFoundHtml

            Send-MailHC @mailParams
            #endregion
        }
        #endregion

        $M = "Sent $I mails"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

        Get-ScriptRuntimeHC -Stop
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}