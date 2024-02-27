#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $importExcel = Get-Command Import-Excel

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        Path                   = New-Item 'TestDrive:/overview.xlsx' -ItemType File
        ScriptName             = 'Test (Brecht)'
        LogFolder              = New-Item 'TestDrive:/log' -ItemType Directory
        RequestTicketURL       = 'https://some-portal-url'
        ExcludedSamAccountName = 'ignoreMe'
        ScriptAdmin            = 'admin@contoso.com'
    }

    Mock Send-MailHC
    Mock Write-EventLog
    Mock Import-Excel
    Mock Get-ADObjectDetailHC
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('Path') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'a terminating error is thrown when' {
    Context 'the path to the Excel file' {
        It 'is not found' {
            $testNewParams = $testParams.Clone()
            $testNewParams.Path = 'notExisting'
            { .$testScript @testNewParams } |
            Should -Throw -ExpectedMessage "File path 'notExisting' not found"
        }
        It 'does not have the extension .xlsx' {
            $testNewParams = $testParams.Clone()
            $testNewParams.Path = New-Item "TestDrive:/test.txt" -ItemType File
            { .$testScript @testNewParams } |
            Should -Throw -ExpectedMessage "File path '*test.txt' does not have extension '.xlsx'"
        }
    }
    It 'an error is found after executing the job to retrieve AD details' {
        Mock Import-Excel {
            @(
                [PSCustomObject]@{
                    MatrixFileName = 'MI6 007 agents'
                    SamAccountName = 'oops'
                }
            )
        } -ParameterFilter { $WorksheetName -eq 'AdObjectNames' }
        Mock Import-Excel {
            @(
                [PSCustomObject]@{
                    MatrixFileName    = 'MI6 007 agents'
                    MatrixResponsible = 'm@contoso.com'
                    MatrixFolderPath  = '\\contoso.com\gbr\MI6\agents'
                    MatrixFilePath    = '\\contoso.com\input\007agents.xlsx'
                }
            )
        } -ParameterFilter { $WorksheetName -eq 'FormData' }
        Mock Get-ADObjectDetailHC {
            Write-Error "Something went wrong"
        }

        { .$testScript @testParams -EA SilentlyContinue } |
        Should -Throw -ExpectedMessage "Error after executing the job that retrieves AD object details, no emails are sent: Something went wrong"
    }
}
Describe 'when there is no terminating error' {
    BeforeAll {
        Mock Import-Excel {
            @(
                [PSCustomObject]@{
                    MatrixFileName = 'MI6 007 agents'
                    SamAccountName = 'craig'
                }
                [PSCustomObject]@{
                    MatrixFileName = 'MI6 007 agents'
                    SamAccountName = 'drNo'
                }
                [PSCustomObject]@{
                    MatrixFileName = 'MI6 007 agents'
                    SamAccountName = 'group1'
                }
                [PSCustomObject]@{
                    MatrixFileName = 'MI6 007 agents'
                    SamAccountName = 'group3'
                }
                [PSCustomObject]@{
                    MatrixFileName = 'Star Trek captains'
                    SamAccountName = 'kirk'
                }
                [PSCustomObject]@{
                    MatrixFileName = 'Star Trek captains'
                    SamAccountName = 'picard'
                }
                [PSCustomObject]@{
                    MatrixFileName = 'Star Trek captains'
                    SamAccountName = 'group2'
                }
                [PSCustomObject]@{
                    MatrixFileName = 'Star Trek captains'
                    SamAccountName = 'group3'
                }
                [PSCustomObject]@{
                    MatrixFileName = 'Team losers'
                    SamAccountName = 'group4'
                }
            )
        } -ParameterFilter { $WorksheetName -eq 'AdObjectNames' }
        Mock Import-Excel {
            @(
                [PSCustomObject]@{
                    MatrixFileName    = 'MI6 007 agents'
                    MatrixResponsible = 'm@contoso.com'
                    MatrixFolderPath  = '\\contoso.com\gbr\MI6\agents'
                    MatrixFilePath    = '\\contoso.com\input\007agents.xlsx'
                }
                [PSCustomObject]@{
                    MatrixFileName    = 'Star Trek captains'
                    MatrixResponsible = 'admiral@contoso.com'
                    MatrixFolderPath  = '\\contoso.com\usa\star-trek'
                    MatrixFilePath    = '\\contoso.com\input\star-trek.xlsx'
                }
                [PSCustomObject]@{
                    MatrixFileName    = 'Team losers'
                    MatrixResponsible = $null
                    MatrixFolderPath  = '\\contoso.com\usa\star-trek'
                    MatrixFilePath    = '\\contoso.com\input\star-trek.xlsx'
                }
            )
        } -ParameterFilter { $WorksheetName -eq 'FormData' }
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                samAccountName = 'craig'
                adObject       = @{
                    ObjectClass = 'user'
                    Name        = 'Craig Daniel'
                }
                adGroupMember  = $null
            }
            [PSCustomObject]@{
                samAccountName = 'kirk'
                adObject       = @{
                    ObjectClass = 'user'
                    Name        = 'James T. Kirk'
                }
                adGroupMember  = $null
            }
            [PSCustomObject]@{
                samAccountName = 'picard'
                adObject       = @{
                    ObjectClass = 'user'
                    Name        = 'Jean Luc Picard'
                }
                adGroupMember  = $null
            }
            [PSCustomObject]@{
                samAccountName = 'group1'
                adObject       = @{
                    ObjectClass = 'group'
                    Name        = 'Group1'
                    ManagedBy   = 'CN=ManagerGroup1,DC=contoso,DC=net'
                }
                adGroupMember  = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Sean Connery'
                        SamAccountName = 'connery'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Timothy Dalton'
                        SamAccountName = 'dalton'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Craig Daniel'
                        SamAccountName = 'craig'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Ignored account'
                        SamAccountName = 'ignoreMe'
                    }
                )
            }
            [PSCustomObject]@{
                samAccountName = 'group2'
                adObject       = @{
                    ObjectClass = 'group'
                    Name        = 'Group2'
                    ManagedBy   = 'CN=ManagerGroup1,DC=contoso,DC=net'
                }
                adGroupMember  = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Chuck Norris'
                        SamAccountName = 'cnorris'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Ignored account'
                        SamAccountName = 'ignoreMe'
                    }
                )
            }
            [PSCustomObject]@{
                samAccountName = 'group3'
                adObject       = @{
                    ObjectClass = 'group'
                    Name        = 'Group3'
                    ManagedBy   = 'CN=ManagerGroup1,DC=contoso,DC=net'
                }
                adGroupMember  = @{
                    ObjectClass    = 'user'
                    Name           = 'Ignored account'
                    SamAccountName = 'ignoreMe'
                }
            }
            [PSCustomObject]@{
                samAccountName = 'group4'
                adObject       = @{
                    ObjectClass = 'group'
                    Name        = 'group4'
                    ManagedBy   = 'CN=ManagerGroup1,DC=contoso,DC=net'
                }
                adGroupMember  = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'khan'
                        SamAccountName = 'khan'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Ignored account'
                        SamAccountName = 'ignoreMe'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'SamAccountName' }
        Mock Get-ADObjectDetailHC {
            [PSCustomObject]@{
                DistinguishedName = 'CN=ManagerGroup1,DC=contoso,DC=net'
                adObject          = @{
                    ObjectClass = 'group'
                    Name        = 'ManagerGroup1'
                }
                adGroupMember     = @(
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Tha Boss'
                        SamAccountName = 'boss'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'The Director'
                        SamAccountName = 'director'
                    }
                    @{
                        ObjectClass    = 'user'
                        Name           = 'Excluded user'
                        SamAccountName = 'ignoreMe'
                    }
                )
            }
        } -ParameterFilter { $Type -eq 'DistinguishedName' }

        .$testScript @testParams
    }
    Context 'the matrix data exported for Cherwell is imported from worksheet' {
        It '<_>' -ForEach @('AdObjectNames', 'FormData') {
            Should -Invoke Import-Excel -Times 1 -Exactly -Scope Describe -ParameterFilter { $WorksheetName -eq $_ }
        }
    }
    Context 'a matrix without MatrixResponsible is' {
        It 'collected in the script variable $matrixWithoutResponsible' {
            $matrixWithoutResponsible | Should -HaveCount 1
            $matrixWithoutResponsible.MatrixFileName | Should -Be 'Team losers'
        }
        It 'not checked for AD details and group members' {
            Should -Not -Invoke Get-ADObjectDetailHC -Scope Describe -ParameterFilter { $ADObjectName -contains 'group4' }
        }
        It 'not exported to an Excel file in the log folder' {
            $testGetParams = @{
                Path    = $testParams.LogFolder
                Recurse = $true
                Filter  = '*Team losers.xlsx'
                File    = $true
            }
            Get-ChildItem @testGetParams | Should -BeNullOrEmpty
        }
    }
    Context 'SamAccountNames used in the Cherwell file' {
        It 'are checked for ADObjectDetails for each unique SamAccountName' {
            foreach ($name in
                @(
                    'craig', 'drNo',
                    'group1', 'group2', 'group3',
                    'kirk', 'picard'
                )
            ) {
                Should -Invoke Get-ADObjectDetailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                    (($ADObjectName | Where-Object { $_ -eq $name }).count -eq 1)
                }
            }
        }
        It 'accounts in ExcludedSamAccountName are ignored as group member' {
            $ADObjectDetails.adGroupMember.SamAccountName |
            Should -Not -BeNullOrEmpty

            $ADObjectDetails.adGroupMember.SamAccountName |
            Should -Not -Contain 'ignoreMe'
        }
    }
    Context 'groups that have a manager' {
        It 'are checked for ADObjectDetails for each unique DistinguishedName' {
            Should -Invoke Get-ADObjectDetailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($ADObjectName -eq 'CN=ManagerGroup1,DC=contoso,DC=net')
            }
        }
        It 'accounts in ExcludedSamAccountName are ignored as group member' {
            $groupManagersAdDetails.adGroupMember.SamAccountName |
            Should -Not -BeNullOrEmpty

            $groupManagersAdDetails.adGroupMember.SamAccountName |
            Should -Not -Contain 'ignoreMe'
        }
    }
    Context 'an Excel file is created' {
        BeforeAll {
            $testGetParams = @{
                Path    = $testParams.LogFolder
                Recurse = $true
                Filter  = '*.xlsx'
                File    = $true
            }
            $testLogFiles = Get-ChildItem @testGetParams
        }
        It 'for each matrix with a MatrixResponsible in the log folder' {
            $testLogFiles | Should -HaveCount 2
            $testLogFiles[0].Name | Should -BeLike '*MI6 007 agents.xlsx'
            $testLogFiles[1].Name | Should -BeLike '*Star Trek captains.xlsx'
        }
        Context "the worksheet 'AccessList' contains" {
            BeforeAll {
                $testImportParams = @{
                    Path          = $testLogFiles[0].FullName
                    WorksheetName = 'AccessList'
                }
                $testExcelFileMatrix1 = & $importExcel @testImportParams

                $testImportParams.Path = $testLogFiles[1].FullName
                $testExcelFileMatrix2 = & $importExcel @testImportParams
            }
            It 'SamAccountName' {
                $testExcelFileMatrix1 | Should -HaveCount 5
                $testExcelFileMatrix1[0].SamAccountName | Should -Be 'craig'
                $testExcelFileMatrix1[1].SamAccountName | Should -Be 'group1'
                $testExcelFileMatrix1[2].SamAccountName | Should -Be 'group1'
                $testExcelFileMatrix1[3].SamAccountName | Should -Be 'group1'
                $testExcelFileMatrix1[4].SamAccountName | Should -Be 'group3'

                $testExcelFileMatrix2 | Should -HaveCount 4
                $testExcelFileMatrix2[0].SamAccountName | Should -Be 'kirk'
                $testExcelFileMatrix2[1].SamAccountName | Should -Be 'picard'
                $testExcelFileMatrix2[2].SamAccountName | Should -Be 'group2'
                $testExcelFileMatrix2[3].SamAccountName | Should -Be 'group3'
            }
            It 'Name' {
                $testExcelFileMatrix1 | Should -HaveCount 5
                $testExcelFileMatrix1[0].Name | Should -Be 'Craig Daniel'
                $testExcelFileMatrix1[1].Name | Should -Be 'Group1'
                $testExcelFileMatrix1[2].Name | Should -Be 'Group1'
                $testExcelFileMatrix1[3].Name | Should -Be 'Group1'
                $testExcelFileMatrix1[4].Name | Should -Be 'Group3'

                $testExcelFileMatrix2 | Should -HaveCount 4
                $testExcelFileMatrix2[0].Name | Should -Be 'James T. Kirk'
                $testExcelFileMatrix2[1].Name | Should -Be 'Jean Luc Picard'
                $testExcelFileMatrix2[2].Name | Should -Be 'Group2'
                $testExcelFileMatrix2[3].Name | Should -Be 'Group3'
            }
            It 'Type' {
                $testExcelFileMatrix1 | Should -HaveCount 5
                $testExcelFileMatrix1[0].Type | Should -Be 'user'
                $testExcelFileMatrix1[1].Type | Should -Be 'group'
                $testExcelFileMatrix1[2].Type | Should -Be 'group'
                $testExcelFileMatrix1[3].Type | Should -Be 'group'
                $testExcelFileMatrix1[4].Type | Should -Be 'group'

                $testExcelFileMatrix2 | Should -HaveCount 4
                $testExcelFileMatrix2[0].Type | Should -Be 'user'
                $testExcelFileMatrix2[1].Type | Should -Be 'user'
                $testExcelFileMatrix2[2].Type | Should -Be 'group'
                $testExcelFileMatrix2[3].Type | Should -Be 'group'
            }
            It 'MemberName' {
                $testExcelFileMatrix1 | Should -HaveCount 5
                $testExcelFileMatrix1[0].MemberName |
                Should -BeNullOrEmpty
                $testExcelFileMatrix1[1].MemberName |
                Should -Be 'Sean Connery'
                $testExcelFileMatrix1[2].MemberName |
                Should -Be 'Timothy Dalton'
                $testExcelFileMatrix1[3].MemberName |
                Should -Be 'Craig Daniel'
                $testExcelFileMatrix1[4].MemberName |
                Should -BeNullOrEmpty

                $testExcelFileMatrix2 | Should -HaveCount 4
                $testExcelFileMatrix2[0].MemberName | Should -BeNullOrEmpty
                $testExcelFileMatrix2[1].MemberName | Should -BeNullOrEmpty
                $testExcelFileMatrix2[2].MemberName | Should -Be 'Chuck Norris'
                $testExcelFileMatrix2[3].MemberName | Should -BeNullOrEmpty
            }
            It 'MemberSamAccountName' {
                $testExcelFileMatrix1 | Should -HaveCount 5
                $testExcelFileMatrix1[0].MemberSamAccountName |
                Should -BeNullOrEmpty
                $testExcelFileMatrix1[1].MemberSamAccountName |
                Should -Be 'connery'
                $testExcelFileMatrix1[2].MemberSamAccountName |
                Should -Be 'dalton'
                $testExcelFileMatrix1[3].MemberSamAccountName |
                Should -Be 'craig'
                $testExcelFileMatrix1[4].MemberSamAccountName |
                Should -BeNullOrEmpty

                $testExcelFileMatrix2 | Should -HaveCount 4
                $testExcelFileMatrix2[0].MemberSamAccountName |
                Should -BeNullOrEmpty
                $testExcelFileMatrix2[1].MemberSamAccountName |
                Should -BeNullOrEmpty
                $testExcelFileMatrix2[2].MemberSamAccountName |
                Should -Be 'cnorris'
                $testExcelFileMatrix2[3].MemberSamAccountName |
                Should -BeNullOrEmpty
            }
        }
        Context "the worksheet 'GroupManagers' contains" {
            BeforeAll {
                $testImportParams = @{
                    Path          = $testLogFiles[0].FullName
                    WorksheetName = 'GroupManagers'
                }
                $testExcelFileMatrix1 = & $importExcel @testImportParams

                $testImportParams.Path = $testLogFiles[1].FullName
                $testExcelFileMatrix2 = & $importExcel @testImportParams
            }
            It 'GroupName' {
                <#
                'MI6 007 agents'
                GroupName ManagerName   ManagerType ManagerMemberName
                --------- -----------   ----------- -----------------
                Group1    ManagerGroup1 group       Tha Boss
                Group1    ManagerGroup1 group       The Director
                Group3    ManagerGroup1 group       Tha Boss
                Group3    ManagerGroup1 group       The Director

                'Star Trek captains'
                GroupName ManagerName   ManagerType ManagerMemberName
                --------- -----------   ----------- -----------------
                Group2    ManagerGroup1 group       Tha Boss
                Group2    ManagerGroup1 group       The Director
                Group3    ManagerGroup1 group       Tha Boss
                Group3    ManagerGroup1 group       The Director
                #>

                $testExcelFileMatrix1 | Should -HaveCount 4
                $testExcelFileMatrix1[0].GroupName | Should -Be 'Group1'
                $testExcelFileMatrix1[1].GroupName | Should -Be 'Group1'
                $testExcelFileMatrix1[2].GroupName | Should -Be 'Group3'
                $testExcelFileMatrix1[3].GroupName | Should -Be 'Group3'

                $testExcelFileMatrix2 | Should -HaveCount 4
                $testExcelFileMatrix2[0].GroupName | Should -Be 'Group2'
                $testExcelFileMatrix2[1].GroupName | Should -Be 'Group2'
                $testExcelFileMatrix2[2].GroupName | Should -Be 'Group3'
                $testExcelFileMatrix2[3].GroupName | Should -Be 'Group3'
            }
            It 'ManagerName' {
                $testExcelFileMatrix1[0].ManagerName |
                Should -Be 'ManagerGroup1'
                $testExcelFileMatrix1[1].ManagerName |
                Should -Be 'ManagerGroup1'
                $testExcelFileMatrix1[2].ManagerName |
                Should -Be 'ManagerGroup1'
                $testExcelFileMatrix1[3].ManagerName |
                Should -Be 'ManagerGroup1'

                $testExcelFileMatrix2[0].ManagerName |
                Should -Be 'ManagerGroup1'
                $testExcelFileMatrix2[1].ManagerName |
                Should -Be 'ManagerGroup1'
                $testExcelFileMatrix2[2].ManagerName |
                Should -Be 'ManagerGroup1'
                $testExcelFileMatrix2[3].ManagerName |
                Should -Be 'ManagerGroup1'
            }
            It 'ManagerType' {
                $testExcelFileMatrix1[0].ManagerType | Should -Be 'group'
                $testExcelFileMatrix1[1].ManagerType | Should -Be 'group'
                $testExcelFileMatrix1[2].ManagerType | Should -Be 'group'
                $testExcelFileMatrix1[3].ManagerType | Should -Be 'group'

                $testExcelFileMatrix2[0].ManagerType | Should -Be 'group'
                $testExcelFileMatrix2[1].ManagerType | Should -Be 'group'
                $testExcelFileMatrix2[2].ManagerType | Should -Be 'group'
                $testExcelFileMatrix2[3].ManagerType | Should -Be 'group'
            }
            It 'ManagerMemberName' {
                $testExcelFileMatrix1[0].ManagerMemberName |
                Should -Be 'Tha Boss'
                $testExcelFileMatrix1[1].ManagerMemberName |
                Should -Be 'The Director'
                $testExcelFileMatrix1[2].ManagerMemberName |
                Should -Be 'Tha Boss'
                $testExcelFileMatrix1[3].ManagerMemberName |
                Should -Be 'The Director'

                $testExcelFileMatrix2[0].ManagerMemberName |
                Should -Be 'Tha Boss'
                $testExcelFileMatrix2[1].ManagerMemberName |
                Should -Be 'The Director'
                $testExcelFileMatrix2[2].ManagerMemberName |
                Should -Be 'Tha Boss'
                $testExcelFileMatrix2[3].ManagerMemberName |
                Should -Be 'The Director'
            }
        }
    }
    Context 'an e-mail is sent for each matrix to the user' {
        It 'defined in the FormData worksheet under MatrixResponsible' {
            Should -Invoke Send-MailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($To -eq 'm@contoso.com') -and
                ($Subject -eq 'MI6 007 agents, 3 users, 2 groups') -and
                ($Attachments -like '*MI6 007 agents.xlsx') -and
                ($Message -like '*
                *<a href="https://some-portal-url" target="_blank"><b>IT Self-service Portal</b></a>*
                *<a href="\\contoso.com\input\007agents.xlsx">MI6 007 agents.xlsx</a>*
                *Folder*<a href="\\contoso.com\gbr\MI6\agents"></a>*
                *Unique users*3*
                *Unique groups*2*
                *Check the attachment for details*')
            }
            Should -Invoke Send-MailHC -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($To -eq 'admiral@contoso.com') -and
                ($Subject -eq 'Star Trek captains, 3 users, 2 groups') -and
                ($Attachments -like '*Star Trek captains.xlsx') -and
                ($Message -like '*
                *<a href="https://some-portal-url" target="_blank"><b>IT Self-service Portal</b></a>*
                *<a href="\\contoso.com\input\star-trek.xlsx">Star Trek captains.xlsx</a>*
                *Folder*<a href="\\contoso.com\usa\star-trek">*
                *Unique users*3*
                *Unique groups*2*
                *Check the attachment for details*')
            }
        }
    }
}