#Requires -Version 7
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
            @('Mail', 'Option') |
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

            #region Mail attachment
            $attachment = @{
                Found    = @()
                NotFound = @()
            }

            foreach ($file in $task.Attachment) {
                $testPathParams = @{
                    LiteralPath = $file
                    PathType    = 'Leaf'
                    ErrorAction = 'Stop'
                }
                if (-not (Test-Path @testPathParams)) {
                    $M = "Attachment file '$file' not found"
                    Write-Warning $M
                    Write-EventLog @EventErrorParams -Message $M

                    $attachment.NotFound += $file
                    Continue
                }

                $M = "Attachment file '$file' found"
                Write-Verbose $M
                Write-EventLog @EventVerboseParams -Message $M

                $attachment.Found += $file
            }

            if ($attachment.Found) {
                $mailParams.Attachments = $attachment.Found
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

            #region Mail From
            if ($task.Mail.From) {
                $mailParams.From = $task.Mail.From
            }
            #endregion

            #region Send mail
            if (
                $task.Option.ErrorWhen.AttachmentNotFound -and
                $attachment.NotFound
            ) {
                $mailParams.Priority = 'High'
                $mailParams.To = $ScriptAdmin
                $mailParams.Subject = '{0} attachment{1} not found' -f
                $($attachment.NotFound.Count),
                $(if ($attachment.NotFound.Count -ne 1) { 's' })

                #region Create HTML list for files not found
                $attachmentNotFoundHtml = if ($attachment.NotFound) {
                    $params = @{
                        Message = $attachment.NotFound
                        Header  = 'Attachments not found'
                    }
                    ConvertTo-HtmlListHC @params
                }
                #endregion

                Write-Verbose 'Send mail to admin'

                $mailParams.BodyHtml = "{0}{1}{2}" -f
                'No e-mail sent to the users because not all attachments were found',
                $task.Mail.Body, $attachmentNotFoundHtml

                Send-MailHC @mailParams

                $M = "No e-mail sent to the users because not all attachments were found"
                Write-Verbose $M
                Write-EventLog @EventErrorParams -Message $M
            }
            else {
                Write-Verbose 'Send mail to user'

                $mailParams.BodyHtml = $task.Mail.Body

                Send-MailHC @mailParams
            }
            #endregion
        }
        #endregion

        $M = "Sent {0} mail{1}" -f $I, $(if ($I -ne 1) {'s'})
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