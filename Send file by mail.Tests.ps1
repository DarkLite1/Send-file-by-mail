#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML, Toolbox.General
#Requires -Version 7

BeforeAll {
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testInputFile = @{
        Tasks = @(
            @{
                Mail   = @{
                    From       = 'mike@contoso.con'
                    Header     = 'mail header'
                    Subject    = 'the subject'
                    To         = @('bob@contoso.com')
                    Cc         = @('jack@constoso.com')
                    Bcc        = @('drake@constoso.com')
                    Priority   = 'Normal'
                    Body       = 'Hello'
                    Attachment = (New-Item 'TestDrive:/a.txt' -ItemType File).FullName
                }
                Option = @{
                    ErrorWhen = @{
                        AttachmentNotFound = $true
                    }
                }
            }
        )
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = 'admin@contoso.com'
    }

    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To[0] -eq $ScriptAdmin[0]) -and
            ($To[1] -eq $ScriptAdmin[1]) -and
            ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            BeforeEach {
                $testNewInputFile = Copy-ObjectHC $testInputFile
            }
            It 'Tasks is missing' {
                $testNewInputFile.Tasks = $null
                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile* Property 'Tasks' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It '<_> is missing' -ForEach @(
                'Mail', 'Option'
            ) {
                $testNewInputFile.Tasks[0].$_ = $null
                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile* Property 'Tasks.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'Mail' {
                It '<_> is missing' -ForEach @(
                    'Header', 'To', 'Body', 'Priority', 'Subject'
                ) {
                    $testNewInputFile.Tasks[0].Mail.$_ = $null
                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile* Property 'Tasks.Mail.$_' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'Priority is incorrect' {
                    $testNewInputFile.Tasks[0].Mail.Priority = 'wrong'
                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile* Property 'Tasks.Mail.Priority' is not 'High', 'Low' or 'Normal'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
}
Describe 'Option.ErrorWhen.AttachmentNotFound is true' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Option.ErrorWhen.AttachmentNotFound = $true
    }
    It 'send an e-mail when all attachments are found' {
        $testNewInputFile.Tasks[0].Attachment = @(
                (New-Item 'TestDrive:/b.txt' -ItemType File).FullName
        )

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Times 1 -Exactly -ParameterFilter {
            ($To -eq $testNewInputFile.Tasks[0].Mail.To) -and
            ($Priority -eq $testNewInputFile.Tasks[0].Mail.Priority) -and
            ($Subject -eq $testNewInputFile.Tasks[0].Mail.Subject) -and
            ($Header -eq $testNewInputFile.Tasks[0].Mail.Header) -and
            ($BodyHTML -eq $testNewInputFile.Tasks[0].Mail.Body) -and
            ($Attachments -eq $testNewInputFile.Tasks[0].Attachment[0])
        }
    }
    It 'send no e-mail when an attachment is not found' {
        $testNewInputFile.Tasks[0].Attachment = @(
                (New-Item 'TestDrive:/c.txt' -ItemType File).FullName,
            'z:\notFound'
        )

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Not -ParameterFilter {
            ($To -eq $testNewInputFile.Tasks[0].Mail.To)
        }
    }
}
Describe 'Option.ErrorWhen.AttachmentNotFound is false' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].Option.ErrorWhen.AttachmentNotFound = $false
    }
    It 'send an e-mail when all attachments are found' {
        $testNewInputFile.Tasks[0].Attachment = @(
                (New-Item 'TestDrive:/b.txt' -ItemType File).FullName
        )

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Times 1 -Exactly -ParameterFilter {
            ($To -eq $testNewInputFile.Tasks[0].Mail.To) -and
            ($Priority -eq $testNewInputFile.Tasks[0].Mail.Priority) -and
            ($Subject -eq $testNewInputFile.Tasks[0].Mail.Subject) -and
            ($Header -eq $testNewInputFile.Tasks[0].Mail.Header) -and
            ($BodyHTML -eq $testNewInputFile.Tasks[0].Mail.Body) -and
            ($Attachments -eq $testNewInputFile.Tasks[0].Attachment[0])
        }
    }
    It 'send an e-mail when an attachment is not found' {
        $testNewInputFile.Tasks[0].Attachment = @(
                (New-Item 'TestDrive:/c.txt' -ItemType File).FullName,
            'z:\notFound'
        )

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Send-MailHC -Times 1 -Exactly -ParameterFilter {
            ($To -eq $testNewInputFile.Tasks[0].Mail.To) -and
            ($Priority -eq $testNewInputFile.Tasks[0].Mail.Priority) -and
            ($Subject -eq $testNewInputFile.Tasks[0].Mail.Subject) -and
            ($Header -eq $testNewInputFile.Tasks[0].Mail.Header) -and
            ($BodyHTML -eq $testNewInputFile.Tasks[0].Mail.Body) -and
            ($Attachments -eq $testNewInputFile.Tasks[0].Attachment[0])
        }
    }
}