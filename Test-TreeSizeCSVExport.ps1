<#
.SYNOPSIS
Parses a TreeSize export CSV and validates the integrity of the scan and the CSV
output. Reports any identified issues to the console and, optionally, to a text file.

.DESCRIPTION
This script assumes that the user has already performed a scan using TreeSize and
followed methodology to export that scan to CSV. This script then reads the scanned
metadata and looks for access issues and other concerns. Finally, it reports the total
number of files and folders scanned and the total size (disk space) of the scan.

.PARAMETER CSVPath
Specifies the path to the TreeSize export CSV file. This parameter is required.

.PARAMETER ReportPath
Optionally, specifies the path to a report file to be created to log the identified
access issues/concerns. This parameter is optional; if not specified, the output will
be written to the console.

.PARAMETER AlsoWriteOutputToConsole
Optionally, specifies that the output should be written to the console in addition to
the report file. This parameter is optional; if not specified, the output will be
written only to the report file (unless no report file is specified, in which case the
output will be written to the console).

.PARAMETER DoNotCleanupMemory
Optionally, specifies that the script should not attempt to clean up memory. This makes
sense to include when the PowerShell workspace will clean up after conclusion of the
script; however, it is not recommended to use this parameter when running the script
with dot-sourcing (e.g., . .\Test-TreeSizeCSVExport.ps1) or other means where the
workspace will not be cleaned up after conclusion of the script.

.EXAMPLE
PS C:\> Test-TreeSizeCSVExport.ps1 -CSVPath 'C:\Users\username\Downloads\File_Server_Analysis\SERVER01_Public.csv' -ReportPath 'C:\Users\flesniak\Downloads\File_Server_Analysis\SERVER01_Public_Report.txt'

This example loads the TreeSize export CSV from the specified path, analyzes it, and
writes findings to the report path specified.

.OUTPUTS
None

.NOTES
TreeSize does not output a properly-formatted CSV, which prevents direct import into
PowerShell (e.g., using Import-Csv).
#>


[cmdletbinding()]

param (
    [Parameter(Mandatory = $true)][string]$CSVPath,
    [Parameter(Mandatory = $false)][string]$ReportPath,
    [Parameter(Mandatory = $false)][switch]$AlsoWriteOutputToConsole,
    [Parameter(Mandatory = $false)][switch]$DoNotCleanupMemory
)

#region License ####################################################################
# Copyright (c) 2023 Frank Lesniak
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this
# software and associated documentation files (the "Software"), to deal in the Software
# without restriction, including without limitation the rights to use, copy, modify,
# merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
# permit persons to whom the Software is furnished to do so, subject to the following
# conditions:
#
# The above copyright notice and this permission notice shall be included in all copies
# or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
# INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A
# PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT
# HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
# CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
# OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#endregion License ####################################################################

$strThisScriptVersionNumber = [version]'1.0.20230612.0'

$datetimeStartOfScript = Get-Date

function Split-StringOnLiteralString {
    # Split-StringOnLiteralString is designed to split a string the way the way that I
    # expected it to be done - using a literal string (as opposed to regex). It's also
    # designed to be backward-compatible with all versions of PowerShell and has been
    # tested successfully on PowerShell v1. My motivation for creating this function
    # was 1) I wanted a split function that behaved more like VBScript's Split
    # function, 2) I do not want to be messing around with RegEx, and 3) I needed code
    # that was backward-compatible with all versions of PowerShell.
    #
    # This function takes two positional arguments
    # The first argument is a string, and the string to be split
    # The second argument is a string or char, and it is that which is to split the string in the first parameter
    #
    # Note: This function always returns an array, even when there is zero or one element in it.
    #
    # Example:
    # $result = Split-StringOnLiteralString 'foo' ' '
    # # $result.GetType().FullName is System.Object[]
    # # $result.Count is 1
    #
    # Example 2:
    # $result = Split-StringOnLiteralString 'What do you think of this function?' ' '
    # # $result.Count is 7

    #region License
    ###############################################################################################
    # Copyright 2020 Frank Lesniak

    # Permission is hereby granted, free of charge, to any person obtaining a copy of this software
    # and associated documentation files (the "Software"), to deal in the Software without
    # restriction, including without limitation the rights to use, copy, modify, merge, publish,
    # distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the
    # Software is furnished to do so, subject to the following conditions:

    # The above copyright notice and this permission notice shall be included in all copies or
    # substantial portions of the Software.

    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
    # BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
    # NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
    # DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    # OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    ###############################################################################################
    #endregion License

    #region DownloadLocationNotice
    # The most up-to-date version of this script can be found on the author's GitHub repository
    # at https://github.com/franklesniak/PowerShell_Resources
    #endregion DownloadLocationNotice

    $strThisFunctionVersionNumber = [version]'2.0.20200820.0'

    trap {
        Write-Error 'An error occurred using the Split-StringOnLiteralString function. This was most likely caused by the arguments supplied not being strings'
    }

    if ($args.Length -ne 2) {
        Write-Error 'Split-StringOnLiteralString was called without supplying two arguments. The first argument should be the string to be split, and the second should be the string or character on which to split the string.'
        $result = @()
    } else {
        $objToSplit = $args[0]
        $objSplitter = $args[1]
        if ($null -eq $objToSplit) {
            $result = @()
        } elseif ($null -eq $objSplitter) {
            # Splitter was $null; return string to be split within an array (of one element).
            $result = @($objToSplit)
        } else {
            if ($objToSplit.GetType().Name -ne 'String') {
                Write-Warning 'The first argument supplied to Split-StringOnLiteralString was not a string. It will be attempted to be converted to a string. To avoid this warning, cast arguments to a string before calling Split-StringOnLiteralString.'
                $strToSplit = [string]$objToSplit
            } else {
                $strToSplit = $objToSplit
            }

            if (($objSplitter.GetType().Name -ne 'String') -and ($objSplitter.GetType().Name -ne 'Char')) {
                Write-Warning 'The second argument supplied to Split-StringOnLiteralString was not a string. It will be attempted to be converted to a string. To avoid this warning, cast arguments to a string before calling Split-StringOnLiteralString.'
                $strSplitter = [string]$objSplitter
            } elseif ($objSplitter.GetType().Name -eq 'Char') {
                $strSplitter = [string]$objSplitter
            } else {
                $strSplitter = $objSplitter
            }

            $strSplitterInRegEx = [regex]::Escape($strSplitter)

            # With the leading comma, force encapsulation into an array so that an array is
            # returned even when there is one element:
            $result = @([regex]::Split($strToSplit, $strSplitterInRegEx))
        }
    }

    # The following code forces the function to return an array, always, even when there are
    # zero or one elements in the array
    $intElementCount = 1
    if ($null -ne $result) {
        if ($result.GetType().FullName.Contains('[]')) {
            if (($result.Count -ge 2) -or ($result.Count -eq 0)) {
                $intElementCount = $result.Count
            }
        }
    }
    $strLowercaseFunctionName = $MyInvocation.InvocationName.ToLower()
    $boolArrayEncapsulation = $MyInvocation.Line.ToLower().Contains('@(' + $strLowercaseFunctionName + ')') -or $MyInvocation.Line.ToLower().Contains('@(' + $strLowercaseFunctionName + ' ')
    if ($boolArrayEncapsulation) {
        $result
    } elseif ($intElementCount -eq 0) {
        , @()
    } elseif ($intElementCount -eq 1) {
        , (, ($args[0]))
    } else {
        $result
    }
}

if ((Test-Path $CSVPath) -eq $false) {
    Write-Warning 'The a TreeSize CSV export does not exist at the specified path. Analysis cannot continue.'
    return
}

Write-Verbose 'Loading the malformed CSV file into memory...'
$arrContent = @(Get-Content -Path $CSVPath | Select-Object -Skip 4)

$strHeader = $arrContent[0]
$arrHeader = Split-StringOnLiteralString $strHeader ','

$hashtableHeaderStatus = @{
    'Full Path' = $false
    'Path' = $false
    'Size' = $false
    'Allocated' = $false
    'Last Modified' = $false
    'Last Accessed' = $false
    'Owner' = $false
    'Permissions' = $false
    'Inherited Permissions' = $false
    'Own Permissions' = $false
    'Type' = $false
    'Creation Date' = $false
}

$listUnmatchedHeaders = New-Object System.Collections.Generic.List[String]
$listRequiredHeadersNotFound = New-Object System.Collections.Generic.List[String]
$intColumnIndexOfPermissions = -1
$intColumnIndexOfAllocated = -1

foreach ($strHeaderElement in $arrHeader) {
    if ($hashtableHeaderStatus.ContainsKey($strHeaderElement)) {
        $hashtableHeaderStatus[$strHeaderElement] = $true
        if ($strHeaderElement -eq 'Permissions') {
            $intColumnIndexOfPermissions = $arrHeader.IndexOf($strHeaderElement)
        } elseif ($strHeaderElement -eq 'Allocated') {
            $intColumnIndexOfAllocated = $arrHeader.IndexOf($strHeaderElement)
        }
    } else {
        $listUnmatchedHeaders.Add($strHeaderElement)
    }
}

# Check $hashtableHeaderStatus for any headers with status $false and report them:
if ($hashtableHeaderStatus.Item('Full Path') -and $hashtableHeaderStatus.Item('Path')) {
    $listRequiredHeadersNotFound.Add('Full Path/Path')
}
foreach ($strHeaderElement in $hashtableHeaderStatus.Keys) {
    if ($strHeaderElement -ne 'Full Path' -and $strHeaderElement -ne 'Path') {
        if ($hashtableHeaderStatus.Item($strHeaderElement) -eq $false) {
            $listRequiredHeadersNotFound.Add($strHeaderElement)
        }
    }
}
if ($listRequiredHeadersNotFound.Count -gt 0) {
    Write-Warning ('The following required headers were not found in the CSV file: ' + $listRequiredHeadersNotFound)
}

if ($listUnmatchedHeaders.Count -gt 0) {
    Write-Warning ('The following headers were not matched to a known header: ' + $listUnmatchedHeaders)
}

if ($DoNotCleanupMemory.IsPresent -eq $false) {
    Write-Verbose 'Cleaning up memory...'
    [System.GC]::Collect() # Force garbage collection to free up memory
}

$datetimeEndOfScript = Get-Date
$timespanStartToFinish = $datetimeEndOfScript - $datetimeStartOfScript
if ($timespanStartToFinish.TotalHours -ge 1) {
    Write-Verbose ('Script completed in ' + [string]::Format('{0:0.00}', ($timespanStartToFinish.TotalHours)) + ' hours (' + [string]::Format('{0:0.00}', ($timespanStartToFinish.TotalMinutes)) + ' minutes; ' + [string]::Format('{0:0.00}', ($timespanStartToFinish.TotalSeconds)) + ' seconds).')
} elseif ($timespanStartToFinish.TotalMinutes -ge 1) {
    Write-Verbose ('Script completed in ' + [string]::Format('{0:0.00}', ($timespanStartToFinish.TotalMinutes)) + ' minutes (' + [string]::Format('{0:0.00}', ($timespanStartToFinish.TotalSeconds)) + ' seconds).')
} else {
    Write-Verbose ('Script completed in ' + [string]::Format('{0:0.00}', ($timespanStartToFinish.TotalSeconds)) + ' seconds.')
}
