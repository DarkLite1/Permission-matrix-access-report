#Requires -Version 5.1
#Requires -Modules ImportExcel, SqlServer

<#
.SYNOPSIS
    Get all AD objects used in a matrix and retrieve the AD object details

.DESCRIPTION
    The script reads a single Excel input file that contains the 
    SamAccountNames of all  AD objects used in the matrix files. This is the
    file created by the Permission Matrix script for displaying date in the 
    Cherwell forms.

    An e-mail is sent to users defined in the field  'MatrixResponsible' with
    in attachment an overview of all users, groups and group members that have
    access to the folders.
    
.PARAMETER Path
    Path to the Excel file containing the matrix information that is used
    for the export to the Cherwell forms.

.PARAMETER MaxThreads
    Quantity of jobs allowed to run at the same time when  querying the 
    active director for details.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [string]$Path,
    [ValidateRange(1, 7)]
    [Int]$MaxThreads = 3,
    [String]$ScriptName = 'Permission matrix access (BNL)',
    [String]$LogFolder = "\\$env:COMPUTERNAME\Log",
    [String]$ScriptAdmin = 'Brecht.Gijbels@heidelbergcement.com'
)

Begin {
    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        Write-EventLog @EventVerboseParams -Message "PSBoundParameters:`n $(
            $PSBoundParameters.GetEnumerator() | 
            ForEach-Object {"`n- $($_.Key): $($_.Value)" })"

        $Error.Clear()

        #region Test valid Path
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            throw "File path '$Path' not found"
        }

        if ($Path -notMatch '.xlsx$') {
            throw "File path '$Path' does not have extension '.xlsx'"
        }
        #endregion

        #region Logging
        $LogParams = @{
            LogFolder    = New-FolderHC -Path $LogFolder -ChildPath "Permission matrix\$ScriptName"
            Name         = $ScriptName
            Date         = 'ScriptStartTime'
            NoFormatting = $true
        }
        $LogFile = New-LogFileNameHC @LogParams
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
        [array]$adObjectNames = Import-Excel @importParams
            
        $M = "Imported $($adObjectNames.Count) rows from the worksheet '$($importParams.WorksheetName)'"
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

        #region Get unique samAccountNames that need to be checked
        $uniqueSamAccountNamesToCheck = $adObjectNames | 
        Where-Object { 
            $matrixWithResponsible.MatrixFileName -contains $_.MatrixFileName 
        } | 
        Select-Object -Property @{
            Name       = 'name'; 
            Expression = { "$($_.SamAccountName)".Trim() } 
        } -Unique |
        Select-Object -ExpandProperty name
        #endregion
    
        #region Get AD object details and AD group members
        $jobs = $adQueryResults = @()

        $startJobParams = @{
            Init        = { Import-Module Toolbox.ActiveDirectory }
            ScriptBlock = {
                Param (
                    $SamAccountName
                )
                Try {
                    $adObject = Get-ADObject -Filter 'SamAccountName -eq $SamAccountName'

                    if ($adObject.ObjectClass -eq 'group') {
                        $adGroupMember = Get-ADGroupMember -Identity $adObject -Recursive
                    }

                    [PSCustomObject]@{
                        samAccountName = $SamAccountName
                        adObject       = $adObject
                        adGroupMember  = $adGroupMember
                    }
                }
                Catch {
                    $errorMessage = $_; $global:error.RemoveAt(0)
                    throw "Failed retrieving details for SamAccountName '$SamAccountName': $errorMessage"
                }
            }
        }

        $M = "Retrieve AD details for $($uniqueSamAccountNamesToCheck.Count) unique SamAccountNames"
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
        ForEach ($samAccountName in $uniqueSamAccountNamesToCheck) {
            Write-Verbose "Get AD details for SamAccountName '$samAccountName'"
            $jobs += Start-Job @startJobParams -ArgumentList $samAccountName
            Wait-MaxRunningJobsHC -Name $jobs -MaxThreads $MaxThreads
        }

        if ($jobs) {
            $adQueryResults = $jobs | Wait-Job | Receive-Job
            Write-Verbose 'Jobs done'
        }

        $M = 'Retrieved AD details'
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion

        #region On error exit the script
        if ($error) {
            throw "Error after executing the job that retrieves AD object details, no emails are sent: $($error.Exception.Message -join ', ')"
        }
        #endregion
            
        #region Create Excel file for each matrix and send mail
        foreach ($matrix in $matrixWithResponsible) {
            $M = "Matrix '$($matrix.MatrixFileName)'"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
            $matrixSamAccountNames = ($adObjectNames | Where-Object {
                    $matrix.MatrixFileName -eq $_.MatrixFileName 
                }).SamAccountName

            $adObjectsToExport = foreach ($s in $matrixSamAccountNames) {
                $adData = $adQueryResults | Where-Object {
                    $s -EQ $_.samAccountName }
                
                if (-not $adData.adObject) {
                    Write-Warning "SamAccountName '$s' not found in AD"
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
                        Expression = { $s } 
                    },
                    @{Name = 'Name'; Expression = { $adData.adObject.Name } },
                    @{Name = 'Type'; Expression = { $adData.adObject.ObjectClass } },
                    @{Name = 'MemberName'; Expression = { $_.Name } },
                    @{Name = 'MemberSamAccountName'; Expression = { $_.SamAccountName } }
                }
            }

            $M = "Created $($adObjectsToExport.Count) AD objects"
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            if ($adObjectsToExport) {
                #region Create Excel file
                $excelParams = @{
                    Path               = "$logFile- $($matrix.MatrixFileName).xlsx"
                    AutoSize           = $true
                    WorksheetName      = 'adObjects'
                    TableName          = 'adObjects'
                    FreezeTopRow       = $true
                    NoNumberConversion = '*'
                }
                
                $M = "Export $($adObjectsToExport.Count) AD objects to Excel file '$($excelParams.Path)'"
                Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

                $adObjectsToExport | Export-Excel @excelParams
                #endregion
                
                #region Send mail to user
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

                $uniqueUserCount = (
                    @(($adObjectsToExport.Where( { $_.Type -eq 'user' })).samAccountName) +
                    @(($adObjectsToExport.Where( { $_.MemberSamAccountName })).MemberSamAccountName) | 
                    Sort-Object -Unique | Where-Object { $_ }
                ).Count

                $uniqueGroupCount = ($adObjectsToExport.Where( { $_.Type -eq 'group' }) | Select-Object Name -Unique).Count

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
                    <p>From experience we know that from time to time a short review of these permissions might be required. Please have a look at the details below and the file in attachment to see if they are still valid. If something needs to be changed, feel free to let us know. We will assist you in bringing your matrix back up-to-date if needed.</p>

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
        #endregion
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