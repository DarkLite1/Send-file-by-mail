#Requires -Modules Pester
#Requires -Modules Toolbox.EventLog, Toolbox.HTML, Toolbox.General
#Requires -Version 5.1

BeforeAll {
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }

    $testInputFile = @{
        Tasks = @(
            @{
                Mail         = @{
                    Header   = 'mail header'
                    Subject  = 'the subject'
                    To       = @('bob@contoso.com')
                    Cc       = @()
                    Bcc      = @()
                    Priority = 'Normal'
                    Body     = 'Hello'
                }
                ComputerName = @('PC1')
                File         = @('c:\test.txt')
            }
        )
    }

    Function ConvertTo-UncPathHC {
        Param (
            [Parameter(Mandatory)]
            [String]$ComputerName,
            [Parameter(Mandatory)]
            [String]$LocalPath
        )
    }

    Mock ConvertTo-UncPathHC
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
                $testNewInputFile | ConvertTo-Json -Depth 3 | 
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
                'Mail', 'ComputerName', 'File'
            ) {
                $testNewInputFile.Tasks[0].$_ = $null
                $testNewInputFile | ConvertTo-Json -Depth 3 | 
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
                    $testNewInputFile | ConvertTo-Json -Depth 3 | 
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
                    $testNewInputFile | ConvertTo-Json -Depth 3 | 
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
            It 'File is incorrect' {
                $testNewInputFile.Tasks[0].File[0] = 'wrong'
                $testNewInputFile | ConvertTo-Json -Depth 3 | 
                Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*: File 'wrong' is not supported, only local file paths are supported.*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
    }
}
Describe 'when the tests are successful' {
    Context 'and the file is found' {
        BeforeAll {
            $testFile = (New-Item 'TestDrive:/file.txt' -ItemType file).FullName

            Mock ConvertTo-UncPathHC {
                $testFile
            }
        
            $testNewInputFile = Copy-ObjectHC $testInputFile

            $testNewInputFile.Tasks[0].Computer = 'PC1'
            $testNewInputFile.Tasks[0].File = 'c:\file.txt'

            $testNewInputFile | ConvertTo-Json -Depth 3 | 
            Out-File @testOutParams
                
            .$testScript @testParams
        }
        It 'the file name is converted to a UNC path' {
            Should -Invoke ConvertTo-UncPathHC -Times 1 -Exactly -Scope 'Context'
        }
        It 'an e-mail is sent to the user with the file in attachment' {
            Should -Invoke Send-MailHC -Times 1 -Exactly -Scope 'Context' -ParameterFilter {
            ($To -eq $testInputFile.Tasks[0].Mail.To) -and
            ($Priority -eq $testInputFile.Tasks[0].Mail.Priority) -and
            ($Subject -eq $testInputFile.Tasks[0].Mail.Subject) -and
            ($Header -eq $testInputFile.Tasks[0].Mail.Header) -and
            ($BodyHTML -eq $testInputFile.Tasks[0].Mail.Body) -and
            ($Attachments -eq $testFile)
            }
        }
    }
    Context 'and the file is not found' {
        BeforeAll {
            $testFile = 'TestDrive:/fileNotExisting.txt'

            Mock ConvertTo-UncPathHC {
                $testFile
            }
            Mock Test-Path {
                $false
            } -ParameterFilter {
                $LiteralPath -eq $testFile
            }
        
            $testNewInputFile = Copy-ObjectHC $testInputFile

            $testNewInputFile.Tasks[0].Computer = 'PC1'
            $testNewInputFile.Tasks[0].File = 'c:\fileNotExisting.txt'

            $testNewInputFile | ConvertTo-Json -Depth 3 | 
            Out-File @testOutParams
                
            .$testScript @testParams
        }
        It 'the file name is converted to a UNC path' {
            Should -Invoke ConvertTo-UncPathHC -Times 1 -Exactly -Scope 'Context'
        }
        It 'an e-mail is sent to the user without the file in attachment' {
            Should -Invoke Send-MailHC -Times 1 -Exactly -Scope 'Context' -ParameterFilter {
            ($To -eq $testInputFile.Tasks[0].Mail.To) -and
            ($Priority -eq 'High') -and
            ($Subject -eq $testInputFile.Tasks[0].Mail.Subject) -and
            ($Header -eq $testInputFile.Tasks[0].Mail.Header) -and
            ($BodyHTML -like "*$($testInputFile.Tasks[0].Mail.Body)*$testFile*") -and
            (-not $Attachments)
            }
        }
    }
}