#Requires -Version 5.1
#Requires -Modules ImportExcel, Toolbox.PermissionMatrix

<#
.SYNOPSIS
    Report about AD objects used in the matrix.

.DESCRIPTION
    The script reads a single Excel input file that contains the 
    SamAccountNames of all AD objects used in the matrix files. This is the
    file created by the 'Permission Matrix' script for displaying data in the 
    Cherwell forms.

    An e-mail is sent for each individual matrix to the e-mail address defined
    in the field 'MatrixResponsible' in the Cherwell file. The e-mail has an
    Excel file in attachment containing two worksheets. One with an overview
    of the AD Objects used in the matrix (SamAccountName, Name, ...) and the 
    group members in case groups are used. The other worksheet contains an
    overview of all the groups used in the matrix, with their manager and the
    members of the group managers (in case it concerns a group).
    
.PARAMETER Path
    Path to the Excel file containing the matrix AD object names. This is the file that has been exported previously by the 'Permission matrix' script 
    with the parameter '-Cherwell'.

.PARAMETER ExcludedSamAccountName
    SamAccountNames that are part of this list will be removed from group 
    memberships and will be disregarded by the entire script.

    This allows for the use of placeholder accounts within AD groups that 
    need to be ignored and are not important to end users.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [string]$Path,
    [String]$ScriptName = 'Permission matrix audit report (BNL)',
    [String]$RequestTicketURL = 'https://1itsm.grouphc.net/CherwellPortal',
    [String[]]$ExcludedSamAccountName = 'srvbatch',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Permission matrix\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start
        
        $Error.Clear()

        Write-EventLog @EventVerboseParams -Message "PSBoundParameters:`n $(
            $PSBoundParameters.GetEnumerator() | 
            ForEach-Object {"`n- $($_.Key): $($_.Value)" })"

        #region Test valid Path
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            throw "File path '$Path' not found"
        }

        if ($Path -notMatch '.xlsx$') {
            throw "File path '$Path' does not have extension '.xlsx'"
        }
        #endregion

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
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
}

Process {
    Try {
        #region Import worksheet AdObjectNames
        $importParams = @{
            WorksheetName = 'AdObjectNames'
            Path          = $Path
            ErrorAction   = 'Stop'
        }
        [array]$worksheetADObjectNames = Import-Excel @importParams
            
        $M = "Imported $($worksheetADObjectNames.Count) rows from the worksheet '$($importParams.WorksheetName)'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion

        #region Import worksheet FormData
        $importParams.WorksheetName = 'FormData'
        [array]$formData = Import-Excel @importParams
            
        $M = "Imported $($formData.Count) rows from the worksheet '$($importParams.WorksheetName)'"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion

        #region Get matrix with and without MatrixResponsible
        $matrixWithResponsible, $matrixWithoutResponsible = $formData.where( 
            { $_.MatrixResponsible }, 'Split')
        #endregion

        #region Get AD object details for objects in the cherwell file
        $samAccountNames = $worksheetADObjectNames | 
        Where-Object { 
            $matrixWithResponsible.MatrixFileName -contains $_.MatrixFileName 
        } | 
        Select-Object -Property @{
            Name       = 'name'; 
            Expression = { "$($_.SamAccountName)".Trim() } 
        } -Unique |
        Select-Object -ExpandProperty name

        $M = "Retrieve AD object details for $($samAccountNames.Count) objects used in the Cherwell file"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
        $getADObjectParams = @{
            SamAccountName   = $samAccountNames
            ADObjectProperty = 'ManagedBy'
        }
        $ADObjectDetails = Get-ADObjectDetailHC @getADObjectParams
        #endregion

        #region Get AD object details for group managers
        if (
            $groupManagers = $ADObjectDetails.ADObject.ManagedBy | 
            Sort-Object -Unique
        ) {
            $M = "Retrieve AD object details for $($groupManagers.Count) group managers"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
                
            $getADObjectParams = @{
                DistinguishedName = $groupManagers
            }
            $groupManagersAdDetails = Get-ADObjectDetailHC @getADObjectParams
        }
        #endregion

        #region Remove group members that are in the ExcludedSamAccountName
        if ($ExcludedSamAccountName) {
            foreach ($adObject in $ADObjectDetails) {
                $adObject.adGroupMember = $adObject.adGroupMember |
                Where-Object { 
                    $ExcludedSamAccountName -notContains $_.SamAccountName 
                }
            }
            foreach ($adObject in $groupManagersAdDetails) {
                $adObject.adGroupMember = $adObject.adGroupMember |
                Where-Object { 
                    $ExcludedSamAccountName -notContains $_.SamAccountName 
                }
            }
        }
        #endregion

        #region On error exit the script
        if ($error.Exception.Message) {
            throw "Error after executing the job that retrieves AD object details, no emails are sent: $($error.Exception.Message -join ', ')"
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
}

End {
    Try {
        #region HTML style
        $htmlStyle = @"
           <style>
               a {
                   color: black;
                   text-decoration: underline;
               }
               a:hover {
                   color: blue;
               }
           
               #matrixTable {
                   border: 1px solid Black;
                   /* padding-bottom: 60px; */
                   /* border-spacing: 0.5em; */
                   border-collapse: separate;
                   border-spacing: 0px 0.6em;
                   /* padding: 10px; */
                   /* width: 600px; */
               }
           
               #matrixTitle {
                   border: none;
                   background-color: lightgrey;
                   text-align: center;
                   padding: 6px;
               }
           
               #matrixHeader {
                   font-weight: normal;
                   letter-spacing: 5pt;
                   font-style: italic;
               }
           
               #matrixFileInfo {
                   font-weight: normal;
                   font-size: 12px;
                   font-style: italic;
                   text-align: center;
               }
           
               <! –– 
               table tbody tr td a {
                   display: block;
                   width: 100%;
                   height: 100%;
               }
               ––> 
           </style>
"@
        #endregion
   
        foreach ($matrix in $matrixWithResponsible) {
            $M = "Matrix '$($matrix.MatrixFileName)'"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
               
            $matrixSamAccountNames = $worksheetADObjectNames | 
            Where-Object { $matrix.MatrixFileName -eq $_.MatrixFileName } |
            Select-Object -ExpandProperty SamAccountName

            #region Create Excel worksheet 'AccessList'
            $AccessListToExport = foreach ($S in $matrixSamAccountNames) {
                $adData = $ADObjectDetails | 
                Where-Object { $S -EQ $_.samAccountName }
                   
                if (-not $adData.adObject) {
                    Write-Warning "SamAccountName '$S' not found in AD"
                }
                elseif (-not $adData.adGroupMember) {
                    $adData | Select-Object -Property SamAccountName, 
                    @{Name = 'Name'; Expression = { $_.adObject.Name } },
                    @{Name = 'Type'; Expression = { $_.adObject.ObjectClass } },
                    MemberName, MemberSamAccountName
                }
                else {
                    $adData.adGroupMember | Select-Object -Property @{
                        Name       = 'SamAccountName'; 
                        Expression = { $S } 
                    },
                    @{Name = 'Name'; Expression = { $adData.adObject.Name } },
                    @{Name = 'Type'; Expression = { $adData.adObject.ObjectClass } },
                    @{Name = 'MemberName'; Expression = { $_.Name } },
                    @{Name = 'MemberSamAccountName'; Expression = { $_.SamAccountName } }
                }
            }
            #endregion

            #region Create Excel worksheet 'GroupManagers'
            $GroupManagersToExport = foreach ($S in $matrixSamAccountNames) {
                $adData = (
                    $ADObjectDetails | Where-Object { 
                        ($S -EQ $_.samAccountName) -and
                        ($_.adObject.ObjectClass -eq 'group')
                    }
                )
                if ($adData) {
                    $groupManager = $groupManagersAdDetails | Where-Object {
                        $_.DistinguishedName -eq $adData.adObject.ManagedBy
                    }

                    if (-not $groupManager) {
                        [PSCustomObject]@{
                            GroupName         = $adData.adObject.Name
                            ManagerName       = $null
                            ManagerType       = $null
                            ManagerMemberName = $null
                        }
                    }
                    elseif (-not $groupManager.adGroupMember) {
                        [PSCustomObject]@{
                            GroupName         = $adData.adObject.Name
                            ManagerName       = $groupManager.adObject.Name
                            ManagerType       = $groupManager.adObject.ObjectClass
                            ManagerMemberName = $null
                        }
                    }
                    else {
                        foreach ($user in $groupManager.adGroupMember) {
                            [PSCustomObject]@{
                                GroupName         = $adData.adObject.Name
                                ManagerName       = $groupManager.adObject.Name
                                ManagerType       = $groupManager.adObject.ObjectClass
                                ManagerMemberName = $user.Name
                            }
                        }
                    }
                }
            }
            #endregion
   
            if ($AccessListToExport) {
                #region Export to Excel worksheet 'AccessList'
                $excelParams = @{
                    Path               = "$logFile- $($matrix.MatrixFileName).xlsx"
                    AutoSize           = $true
                    WorksheetName      = 'AccessList'
                    TableName          = 'AccessList'
                    FreezeTopRow       = $true
                    NoNumberConversion = '*'
                }
                   
                $M = "Export $($AccessListToExport.Count) AD objects to Excel file '$($excelParams.Path)' worksheet '$($excelParams.WorksheetName)'"
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
   
                $AccessListToExport | Export-Excel @excelParams
                #endregion

                #region Export to Excel worksheet 'GroupManagers'
                if ($GroupManagersToExport) {
                    $excelParams.WorksheetName = $excelParams.TableName = 'GroupManagers'
                    
                    $M = "Export $($GroupManagersToExport.Count) AD objects to Excel file '$($excelParams.Path)' worksheet '$($excelParams.WorksheetName)'"
                    Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
                    
                    $GroupManagersToExport | Export-Excel @excelParams
                }
                #endregion
                   
                #region Send mail to user
                $uniqueUserCount = (
                    @(($AccessListToExport.Where( { $_.Type -eq 'user' })).samAccountName) +
                    @(($AccessListToExport.Where( { $_.MemberSamAccountName })).MemberSamAccountName) | 
                    Sort-Object -Unique | Where-Object { $_ }
                ).Count
   
                $uniqueGroupCount = ($AccessListToExport.Where( { $_.Type -eq 'group' }) | Select-Object Name -Unique).Count
   
                $M = "Send mail with $uniqueUserCount unique user accounts and $uniqueGroupCount unique groups"
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
   
                $mailParams = @{
                    To          = $matrix.MatrixResponsible
                    Bcc         = $ScriptAdmin
                    Subject     = "$($matrix.MatrixFileName), $uniqueUserCount users, $uniqueGroupCount groups"
                    Attachments = $excelParams.Path
                    Message     =
                    "$htmlStyle
                       <p>Dear matrix responsible</p>
                       <p>Managing folder access is not always easy. People are joining and leaving the company, moving departments, changing jobs, ... . To facilitate this task we created the 'Permission matrix' script, an automated way to set permissions on files and folders that are shared with colleagues. This allows you to easily manage folder access by filling in an Excel worksheet containing the folder names, the user groups and the corresponding read or write permissions. </p>

                       <p>From experience we know that from time to time a short review of these permissions might be required. Please have a look at the details below and the file in attachment to see if they are still valid.</p>
                       
                       <p>In case something needs to be updated or changed, feel free to report this to us by submitting the form <b>`"Request folder/role access`"</b> on the <a href=`"$RequestTicketURL`" target=`"_blank`"><b>IT Self-service Portal</b></a>.</p>
   
                       <table id=`"matrixTable`">
                           <tr>
                               <th id=`"matrixTitle`" colspan=`"2`"><a href=""$($matrix.MatrixFilePath)"">$($matrix.MatrixFileName).xlsx</a></th>
                           </tr>
                           <tr>
                               <th>Category</th>
                               <td>$($matrix.MatrixCategoryName)</td>
                           </tr>
                           <tr>
                               <th>Sub category</th>
                               <td>$($matrix.MatrixSubCategoryName)</td>
                           </tr>
                           <tr>
                               <th>Folder</th>
                               <td><a href=""$($matrix.MatrixFolderPath)"">$($matrix.MatrixFolderDisplayName)</a></td>
                           </tr>
                           <tr>
                               <th>Responsible</th>
                               <td>$($matrix.MatrixResponsible -join ', ')</td>
                           </tr>
                       </table>
   
                       <p>Folder access summary:</p>
                       <table id=`"matrixTable`">
                           <tr>
                               <th>Unique users</th>
                               <td>$uniqueUserCount</td>
                           </tr>
                           <tr>
                               <th>Unique groups</th>
                               <td>$uniqueGroupCount</td>
                           </tr>
                       </table>
                           
                       <p><i>* Check the attachment for details</i></p>"
                    LogFolder   = $LogParams.LogFolder
                    Header      = $ScriptName
                    Save        = $LogFile + ' - Mail.html'
                }
                Send-MailHC @mailParams
                #endregion
            }
        }
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}