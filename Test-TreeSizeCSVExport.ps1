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
PS C:\> Test-TreeSizeCSVExport.ps1 -CSVPath 'C:\Users\username\Downloads\File_Server_Analysis\SERVER01_Public.csv' -ReportPath 'C:\Users\username\Downloads\File_Server_Analysis\SERVER01_Public_Report.txt'

This example loads the TreeSize export CSV from the specified path, analyzes it, and
writes findings to the specified report path.

.OUTPUTS
None

.NOTES
TreeSize does not output a properly-formatted CSV, which prevents direct import into
PowerShell (e.g., using Import-Csv).
#>


[cmdletbinding()]

param (
    [Parameter(Mandatory = $true)][string]$CSVPath,
    [Parameter(Mandatory = $false)][string]$PathToOutputCSV = '',
    [Parameter(Mandatory = $false)][string]$ReportPath,
    [Parameter(Mandatory = $false)][switch]$AlsoWriteOutputToConsole,
    [Parameter(Mandatory = $false)][switch]$DoNotCleanupMemory,
    [Parameter(Mandatory = $false)][double]$SizeToleranceThreshold = 0.01
)

#region License ####################################################################
# Copyright (c) 2024 Frank Lesniak
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

$strThisScriptVersionNumber = [version]'1.2.20240823.0'

$datetimeStartOfScript = Get-Date

$__TREEINCLUDETYPE = $false
$__TREEINCLUDESIZE = $true
$__TREEINCLUDEALLOCATED = $false
$__TREEINCLUDELASTMODIFIED = $false
$__TREEINCLUDELASTACCESSED = $false
$__TREEINCLUDECREATIONDATE = $false
$__TREEINCLUDEOWNER = $false
$__TREEINCLUDEPERMISSIONS = $true
$__TREEINCLUDEINHERITEDPERMISSIONS = $false
$__TREEINCLUDEOWNPERMISSIONS = $false
$__TREEINCLUDEATTRIBUTES = $false

$boolDebug = $false

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

function New-PSObjectTreeDirectoryElement {
    $refPSObjectTreeElement = $args[0]

    $PSObjectTreeDirectoryElement = New-Object -TypeName PSObject

    $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'FullPath' -Value ''
    $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'Name' -Value ''
    $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'ParentPath' -Value ''
    $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'ReferenceToParentElement' -Value [ref]$null
    $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'ChildDirectories' -Value @{}
    $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'ChildFiles' -Value @{}
    if ($script:__TREEINCLUDETYPE -eq $true) {
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'Type' -Value 'Folder'
    }
    if ($script:__TREEINCLUDESIZE -eq $true) {
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'SizeInBytesAsReportedByTreeSize' -Value [uint64]0
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'RolledUpSizeInBytes' -Value [uint64]0
    }
    if ($script:__TREEINCLUDEALLOCATED -eq $true) {
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'DiskAllocationInBytesAsReportedByTreeSize' -Value [uint64]0
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'RolledUpDiskAllocationInBytes' -Value [uint64]0
    }
    if ($script:__TREEINCLUDELASTMODIFIED -eq $true) {
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'LastModified' -Value [datetime]::MinValue
    }
    if ($script:__TREEINCLUDELASTACCESSED -eq $true) {
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'LastAccessed' -Value [datetime]::MinValue
    }
    if ($script:__TREEINCLUDECREATIONDATE -eq $true) {
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'CreationDate' -Value [datetime]::MinValue
    }
    if ($script:__TREEINCLUDEOWNER -eq $true) {
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'Owner' -Value ''
    }
    if ($script:__TREEINCLUDEPERMISSIONS -eq $true) {
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'Permissions' -Value ''
    }
    if ($script:__TREEINCLUDEINHERITEDPERMISSIONS -eq $true) {
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'InheritedPermissions' -Value ''
    }
    if ($script:__TREEINCLUDEOWNPERMISSIONS -eq $true) {
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'OwnPermissions' -Value ''
    }
    if ($script:__TREEINCLUDEATTRIBUTES -eq $true) {
        $PSObjectTreeDirectoryElement | Add-Member -MemberType NoteProperty -Name 'Attributes' -Value ''
    }

    $refPSObjectTreeElement.Value = $PSObjectTreeDirectoryElement

    return $true
}

function New-PSObjectTreeFileElement {
    $refPSObjectTreeElement = $args[0]

    $PSObjectTreeFileElement = New-Object -TypeName PSObject

    $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'FullPath' -Value ''
    $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'Name' -Value ''
    $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'ParentPath' -Value ''
    $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'ReferenceToParentElement' -Value [ref]$null
    if ($script:__TREEINCLUDETYPE -eq $true) {
        $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'Type' -Value 'File'
    }
    if ($script:__TREEINCLUDESIZE -eq $true) {
        $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'SizeInBytesAsReportedByTreeSize' -Value [uint64]0
    }
    if ($script:__TREEINCLUDEALLOCATED -eq $true) {
        $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'DiskAllocationInBytesAsReportedByTreeSize' -Value [uint64]0
    }
    if ($script:__TREEINCLUDELASTMODIFIED -eq $true) {
        $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'LastModified' -Value [datetime]::MinValue
    }
    if ($script:__TREEINCLUDELASTACCESSED -eq $true) {
        $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'LastAccessed' -Value [datetime]::MinValue
    }
    if ($script:__TREEINCLUDECREATIONDATE -eq $true) {
        $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'CreationDate' -Value [datetime]::MinValue
    }
    if ($script:__TREEINCLUDEOWNER -eq $true) {
        $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'Owner' -Value ''
    }
    if ($script:__TREEINCLUDEPERMISSIONS -eq $true) {
        $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'Permissions' -Value ''
    }
    if ($script:__TREEINCLUDEINHERITEDPERMISSIONS -eq $true) {
        $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'InheritedPermissions' -Value ''
    }
    if ($script:__TREEINCLUDEOWNPERMISSIONS -eq $true) {
        $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'OwnPermissions' -Value ''
    }
    if ($script:__TREEINCLUDEATTRIBUTES -eq $true) {
        $PSObjectTreeFileElement | Add-Member -MemberType NoteProperty -Name 'Attributes' -Value ''
    }

    $refPSObjectTreeElement.Value = $PSObjectTreeFileElement

    return $true
}

function Add-RolledUpSizeOfTreeRecursively {
    $refUInt64RolledUpSizeInBytes = $args[0]
    $refUInt64RolledUpDiskAllocationInBytes = $args[1]
    $refPSObjectTreeDirectoryElement = $args[2]
    $boolRollUpSize = $args[3]
    $boolRollUpDiskAllocation = $args[4]
    $boolWarnWhenSizeDifferenceExceedsThreshold = $args[5]
    $boolReturnErrorWhenSizeDifferenceExceedsThreshold = $args[6]
    $boolWarnWhenDiskAllocationDifferenceExceedsThreshold = $args[7]
    $boolReturnErrorWhenDiskAllocationDifferenceExceedsThreshold = $args[8]
    $doubleThresholdInDecimal = $args[9] # 0.01 for 1%

    if ($__TREEINCLUDESIZE -eq $false) {
        $boolRollUpSize = $false
        $boolWarnWhenSizeDifferenceExceedsThreshold = $false
        $boolReturnErrorWhenSizeDifferenceExceedsThreshold = $false
    }

    if ($__TREEINCLUDEALLOCATED -eq $false) {
        $boolRollUpDiskAllocation = $false
        $boolWarnWhenDiskAllocationDifferenceExceedsThreshold = $false
        $boolReturnErrorWhenDiskAllocationDifferenceExceedsThreshold = $false
    }

    $uint64RunningTotalSizeInBytes = [uint64]0
    $uint64RunningTotalDiskAllocationInBytes = [uint64]0

    if ($null -eq ($refPSObjectTreeDirectoryElement.Value)) {
        return $false
    }

    # Process all the child folders recursively
    @((($refPSObjectTreeDirectoryElement.Value).ChildDirectories).PSObject.Properties | Where-Object { $_.Name -eq 'Keys' } | ForEach-Object { $_.Value }) | ForEach-Object {
        # $_ is the child directory name
        $refPSObjectTreeChildDirectoryElement = [ref]$null
        if ([string]::IsNullOrEmpty($_) -eq $false) {
            $refPSObjectTreeChildDirectoryElement = (($refPSObjectTreeDirectoryElement.Value).ChildDirectories).Item($_)
            $uint64ThisChildFolderSizeInBytes = [uint64]0
            $uint64ThisChildFolderDiskAllocationInBytes = [uint64]0
            $boolSuccess = Add-RolledUpSizeOfTreeRecursively ([ref]$uint64ThisChildFolderSizeInBytes) ([ref]$uint64ThisChildFolderDiskAllocationInBytes) $refPSObjectTreeChildDirectoryElement $boolRollUpSize $boolRollUpDiskAllocation $boolWarnWhenSizeDifferenceExceedsThreshold $boolReturnErrorWhenSizeDifferenceExceedsThreshold $boolWarnWhenDiskAllocationDifferenceExceedsThreshold $boolReturnErrorWhenDiskAllocationDifferenceExceedsThreshold $doubleThresholdInDecimal
            if ($boolSuccess -eq $false) {
                return $false
            } else {
                if ($boolRollUpSize -eq $true) {
                    $uint64RunningTotalSizeInBytes += $uint64ThisChildFolderSizeInBytes
                }
                if ($boolRollUpDiskAllocation -eq $true) {
                    $uint64RunningTotalDiskAllocationInBytes += $uint64ThisChildFolderDiskAllocationInBytes
                }
            }
        }
    }

    # Process all the child files
    @((($refPSObjectTreeDirectoryElement.Value).ChildFiles).PSObject.Properties | Where-Object { $_.Name -eq 'Keys' } | ForEach-Object { $_.Value }) | ForEach-Object {
        # $_ is the child file name
        $refPSObjectTreeChildFileElement = [ref]$null
        if ([string]::IsNullOrEmpty($_) -eq $false) {
            $refPSObjectTreeChildFileElement = (($refPSObjectTreeDirectoryElement.Value).ChildFiles).Item($_)

            if ($boolRollUpSize -eq $true) {
                $uint64RunningTotalSizeInBytes += ($refPSObjectTreeChildFileElement.Value).SizeInBytesAsReportedByTreeSize
            }
            if ($boolRollUpDiskAllocation -eq $true) {
                $uint64RunningTotalDiskAllocationInBytes += ($refPSObjectTreeChildFileElement.Value).DiskAllocationInBytesAsReportedByTreeSize
            }
        }
    }

    if ($boolRollUpSize -eq $true) {
        ($refPSObjectTreeDirectoryElement.Value).RolledUpSizeInBytes = $uint64RunningTotalSizeInBytes
        $refUInt64RolledUpSizeInBytes.Value = $uint64RunningTotalSizeInBytes

        if ($boolWarnWhenSizeDifferenceExceedsThreshold -eq $true) {
            $uint64SizeInBytesAsReportedByTreeSizeLowWatermark = [uint64](($refPSObjectTreeDirectoryElement.Value).SizeInBytesAsReportedByTreeSize * (1 - $doubleThresholdInDecimal))
            $uint64SizeInBytesAsReportedByTreeSizeHighWatermark = [uint64](($refPSObjectTreeDirectoryElement.Value).SizeInBytesAsReportedByTreeSize * (1 + $doubleThresholdInDecimal))
            if ($uint64RunningTotalSizeInBytes -lt $uint64SizeInBytesAsReportedByTreeSizeLowWatermark) {
                Write-Warning ('The rolled up size of the folder "' + ($refPSObjectTreeDirectoryElement.Value).FullPath + '" is ' + $uint64RunningTotalSizeInBytes + ' bytes, which is ' + (($refPSObjectTreeDirectoryElement.Value).SizeInBytesAsReportedByTreeSize - $uint64RunningTotalSizeInBytes) + ' bytes LESS THAN the size reported by TreeSize (' + ($refPSObjectTreeDirectoryElement.Value).SizeInBytesAsReportedByTreeSize + ' bytes). This may be caused by temporary/ephemeral files or other files that TreeSize does not have access to read in the child-most folder where this problem occurs.')
                if ($boolReturnErrorWhenSizeDifferenceExceedsThreshold -eq $true) {
                    return $false
                }
            } elseif ($uint64RunningTotalSizeInBytes -gt $uint64SizeInBytesAsReportedByTreeSizeHighWatermark) {
                Write-Information ('The rolled up size of the folder "' + ($refPSObjectTreeDirectoryElement.Value).FullPath + '" is ' + $uint64RunningTotalSizeInBytes + ' bytes, which is ' + ($uint64RunningTotalSizeInBytes - ($refPSObjectTreeDirectoryElement.Value).SizeInBytesAsReportedByTreeSize) + ' bytes MORE THAN the size reported by TreeSize (' + ($refPSObjectTreeDirectoryElement.Value).SizeInBytesAsReportedByTreeSize + ' bytes). This is unusual and may be caused by a bug in TreeSize.')
                if ($boolReturnErrorWhenSizeDifferenceExceedsThreshold -eq $true) {
                    return $false
                }
            }
        }
    }
    if ($boolRollUpDiskAllocation -eq $true) {
        ($refPSObjectTreeDirectoryElement.Value).RolledUpDiskAllocationInBytes = $uint64RunningTotalDiskAllocationInBytes
        $refUInt64RolledUpDiskAllocationInBytes.Value = $uint64RunningTotalDiskAllocationInBytes

        if ($boolWarnWhenDiskAllocationDifferenceExceedsThreshold -eq $true) {
            $uint64DiskAllocationInBytesAsReportedByTreeSizeLowWatermark = [uint64](($refPSObjectTreeDirectoryElement.Value).DiskAllocationInBytesAsReportedByTreeSize * (1 - $doubleThresholdInDecimal))
            $uint64DiskAllocationInBytesAsReportedByTreeSizeHighWatermark = [uint64](($refPSObjectTreeDirectoryElement.Value).DiskAllocationInBytesAsReportedByTreeSize * (1 + $doubleThresholdInDecimal))
            if ($uint64RunningTotalDiskAllocationInBytes -lt $uint64DiskAllocationInBytesAsReportedByTreeSizeLowWatermark) {
                Write-Warning ('The rolled up disk allocation of the folder "' + ($refPSObjectTreeDirectoryElement.Value).FullPath + '" is ' + $uint64RunningTotalDiskAllocationInBytes + ' bytes, which is ' + (($refPSObjectTreeDirectoryElement.Value).DiskAllocationInBytesAsReportedByTreeSize - $uint64RunningTotalDiskAllocationInBytes) + ' bytes LESS THAN the disk allocation reported by TreeSize (' + ($refPSObjectTreeDirectoryElement.Value).DiskAllocationInBytesAsReportedByTreeSize + ' bytes). This may be caused by temporary/ephemeral files or other files that TreeSize does not have access to read in the child-most folder where this problem occurs.')
                if ($boolReturnErrorWhenDiskAllocationDifferenceExceedsThreshold -eq $true) {
                    return $false
                }
            } elseif ($uint64RunningTotalDiskAllocationInBytes -gt $uint64DiskAllocationInBytesAsReportedByTreeSizeHighWatermark) {
                Write-Information ('The rolled up disk allocation of the folder "' + ($refPSObjectTreeDirectoryElement.Value).FullPath + '" is ' + $uint64RunningTotalDiskAllocationInBytes + ' bytes, which is ' + ($uint64RunningTotalDiskAllocationInBytes - ($refPSObjectTreeDirectoryElement.Value).DiskAllocationInBytesAsReportedByTreeSize) + ' bytes MORE THAN the disk allocation reported by TreeSize (' + ($refPSObjectTreeDirectoryElement.Value).DiskAllocationInBytesAsReportedByTreeSize + ' bytes). Most likely this is because TreeSize does not have access to all the files or subfolders in the folder.')
                if ($boolReturnErrorWhenDiskAllocationDifferenceExceedsThreshold -eq $true) {
                    return $false
                }
            }
        }
    }

    return $true
}

function Convert-TreeSizeSizeToUInt64Bytes {
    $refUInt64SizeInBytes = $args[0]
    $strTreeSizeSize = $args[1]

    $strWorkingTreeSizeSize = $strTreeSizeSize.Trim()
    $arrWorkingTreeSizeSize = Split-StringOnLiteralString $strWorkingTreeSizeSize ' '

    if ($arrWorkingTreeSizeSize[1] -eq 'Bytes') {
        $uint64SizeInBytes = [uint64]($arrWorkingTreeSizeSize[0])
    } elseif ($arrWorkingTreeSizeSize[1] -eq 'KB') {
        $uint64SizeInBytes = [uint64]([double]($arrWorkingTreeSizeSize[0]) * 1024)
    } elseif ($arrWorkingTreeSizeSize[1] -eq 'MB') {
        $uint64SizeInBytes = [uint64]([double]($arrWorkingTreeSizeSize[0]) * 1024 * 1024)
    } elseif ($arrWorkingTreeSizeSize[1] -eq 'GB') {
        $uint64SizeInBytes = [uint64]([double]($arrWorkingTreeSizeSize[0]) * 1024 * 1024 * 1024)
    } elseif ($arrWorkingTreeSizeSize[1] -eq 'TB') {
        $uint64SizeInBytes = [uint64]([double]($arrWorkingTreeSizeSize[0]) * 1024 * 1024 * 1024 * 1024)
    } elseif ($arrWorkingTreeSizeSize[1] -eq 'PB') {
        $uint64SizeInBytes = [uint64]([double]($arrWorkingTreeSizeSize[0]) * 1024 * 1024 * 1024 * 1024 * 1024)
    } else {
        # Write-Warning ('The TreeSize size "' + $strTreeSizeSize + '" could not be converted to a UInt64 value.')
        return $false
    }

    $refUInt64SizeInBytes.Value = $uint64SizeInBytes
    return $true
}

function Get-PSVersion {
    <#
    .SYNOPSIS
    Returns the version of PowerShell that is running

    .DESCRIPTION
    Returns the version of PowerShell that is running, including on the original
    release of Windows PowerShell (version 1.0)

    .EXAMPLE
    Get-PSVersion

    This example returns the version of PowerShell that is running. On versions of
    PowerShell greater than or equal to version 2.0, this function returns the
    equivalent of $PSVersionTable.PSVersion

    .OUTPUTS
    A [version] object representing the version of PowerShell that is running

    .NOTES
    PowerShell 1.0 does not have a $PSVersionTable variable, so this function returns
    [version]('1.0') on PowerShell 1.0
    #>

    [CmdletBinding()]
    [OutputType([version])]

    param ()

    #region License ################################################################
    # Copyright (c) 2023 Frank Lesniak
    #
    # Permission is hereby granted, free of charge, to any person obtaining a copy of
    # this software and associated documentation files (the "Software"), to deal in the
    # Software without restriction, including without limitation the rights to use,
    # copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the
    # Software, and to permit persons to whom the Software is furnished to do so,
    # subject to the following conditions:
    #
    # The above copyright notice and this permission notice shall be included in all
    # copies or substantial portions of the Software.
    #
    # THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    # IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
    # FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
    # COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN
    # AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
    # WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
    #endregion License ################################################################

    $versionThisFunction = [version]('1.0.20230613.0')

    if (Test-Path variable:\PSVersionTable) {
        $PSVersionTable.PSVersion
    } else {
        [version]('1.0')
    }
}

function Test-PermissionValidity {
    # From Convert-SummarizedNTFSPermissionsToCleanedUpAndRemappedNTFSPermissions.ps1

    $strExistingPermissionLevel = $args[0]
    $boolFunctionReturn = $false

    $strExistingPermissionLevel = $strExistingPermissionLevel
    if ($strExistingPermissionLevel -ceq 'full') {
        return $true
    } elseif ($strExistingPermissionLevel -eq '+r+w+x') {
        return $true
    } elseif ($strExistingPermissionLevel -eq '-r-w-x') {
        return $true
    } elseif ($strExistingPermissionLevel -eq '+r+x') {
        return $true
    } elseif ($strExistingPermissionLevel -eq '-r-x') {
        return $true
    } elseif ($strExistingPermissionLevel -eq '+x') {
        # bare execute permissions are very unusual on Windows, but technically valid
        return $true
    } elseif ($strExistingPermissionLevel -eq '-x') {
        # bare execute permissions are very unusual on Windows, but technically valid
        return $true
    } elseif ($strExistingPermissionLevel -eq '+r') {
        return $true
    } elseif ($strExistingPermissionLevel -eq '-r') {
        return $true
    }

    return $boolFunctionReturn
}

function Test-TreeSizePermissionsRecord {
    $refStringWarningMessage = $args[0]
    $strPermissions = $args[1]

    $strWorkingWarningMessage = ''

    $arrPermissions = Split-StringOnLiteralString $strPermissions ' | '

    $intPermissionEntryCount = $arrPermissions.Count
    for ($intCounter = 0; $intCounter -lt $intPermissionEntryCount; $intCounter++) {
        $arrSinglePermissionEntry = Split-StringOnLiteralString ($arrPermissions[$intCounter]) ': '
        if ($arrSinglePermissionEntry.Count -ne 2) {
            if ([string]::IsNullOrEmpty($strWorkingWarningMessage) -eq $false) {
                if ($arrPermissions[$intCounter] -eq 'The system cannot find the file specified') {
                    $strWorkingWarningMessage += '; the permission entry "' + $arrPermissions[$intCounter] + '" may be caused by a temporary/ephmeral file that TreeSize was not able to read'
                } else {
                    $strWorkingWarningMessage += '; the permission entry "' + $arrPermissions[$intCounter] + '" is not in the format "Account: Permission"'
                }
            } else {
                if ($arrPermissions[$intCounter] -eq 'The system cannot find the file specified') {
                    $strWorkingWarningMessage = 'The permission entry "' + $arrPermissions[$intCounter] + '" may be caused by a temporary/ephmeral file that TreeSize was not able to read'
                } else {
                    $strWorkingWarningMessage = 'The permission entry "' + $arrPermissions[$intCounter] + '" is not in the format "Account: Permission"'
                }
            }
        } else {
            # Correctly formatted permision entry in format "Account: Permission"
            $strAccount = $arrSinglePermissionEntry[0]
            if ($strAccount -eq 'The trust relationship between the primary domain and the trusted domain failed') {
                if ([string]::IsNullOrEmpty($strWorkingWarningMessage) -eq $false) {
                    $strWorkingWarningMessage += '; the permission entry "' + $arrPermissions[$intCounter] + '" shows a problem translating a SID to a principal name due to a failed client-domain trust relationship (or a failed domain-domain trust relationship)'
                } else {
                    $strWorkingWarningMessage = 'The permission entry "' + $arrPermissions[$intCounter] + '" shows a problem translating a SID to a principal name due to a failed client-domain trust relationship (or a failed domain-domain trust relationship)'
                }
            }

            $strPermission = $arrSinglePermissionEntry[1]
            $boolPermissionValid = Test-PermissionValidity $strPermission
            if ($boolPermissionValid -eq $false) {
                if ([string]::IsNullOrEmpty($strWorkingWarningMessage) -eq $false) {
                    $strWorkingWarningMessage += '; the permission entry "' + $arrPermissions[$intCounter] + '" is not known/valid'
                } else {
                    $strWorkingWarningMessage = 'The permission entry "' + $arrPermissions[$intCounter] + '" is not known/valid'
                }
            }
        }
    }

    if ([string]::IsNullOrEmpty($strWorkingWarningMessage) -eq $false) {
        $refStringWarningMessage.Value = $strWorkingWarningMessage
        return $false
    } else {
        return $true
    }
}

function Test-TreeSizeAttributesRecord {
    $refStringWarningMessage = $args[0]
    $strAttributes = $args[1]

    $strWorkingWarningMessage = ''

    if ([string]::IsNullOrEmpty($strAttributes) -eq $false) {
        # No attributes is perfectly OK
        return $true
    }

    for ($intPosition = 0; $intPosition -le $strAttributes.Length; $intPosition++) {
        $strCharacter = $strAttributes[$intPosition]
        if ($strCharacter -eq 'E') {
            $strWorkingWarningMessage = 'The item is encrypted'
        }
    }

    if ([string]::IsNullOrEmpty($strWorkingWarningMessage) -eq $false) {
        $refStringWarningMessage.Value = $strWorkingWarningMessage
        return $false
    } else {
        return $true
    }
}

$versionPS = Get-PSVersion

#region Quit if PowerShell Version is Unsupported ##################################
if ($versionPS -lt [version]'2.0') {
    Write-Warning 'This script requires PowerShell v2.0 or higher. Please upgrade to PowerShell v2.0 or higher and try again.'
    return # Quit script
}
#endregion Quit if PowerShell Version is Unsupported ##################################

if ((Test-Path $CSVPath) -eq $false) {
    Write-Warning 'The a TreeSize CSV export does not exist at the specified path. Analysis cannot continue.'
    return
}

$strFolderContainingCSV = Split-Path -Path $CSVPath -Parent
$strCSVFileName = Split-Path -Path $CSVPath -Leaf
$arrCSVFileNameParts = Split-StringOnLiteralString $strCSVFileName '.csv'
$strFileNamePrefix = $arrCSVFileNameParts[0]

if ([string]::IsNullOrEmpty($PathToOutputFile) -eq $true) {
    $strFixedCSVInputFilePath = Join-Path $strFolderContainingCSV ($strFileNamePrefix + '_Fixed.csv')
} else {
    $strFixedCSVInputFilePath = $PathToOutputCSV
}

Write-Verbose 'Loading the TreeSize CSV file into memory. This may take a while...'
$arrContent = @(Get-Content -Path $CSVPath | Select-Object -Skip 4)

Write-Verbose 'Writing the corrected CSV file to disk. This may take a while...'
$arrContent | Set-Content -Path $strFixedCSVInputFilePath
$arrContent = $null

if ($DoNotCleanupMemory.IsPresent -eq $false) {
    Write-Verbose 'Cleaning up memory. This may take a while...'
    [System.GC]::Collect() # Force garbage collection to free up memory
}

Write-Verbose 'Importing the CSV into memory. This may take a while...'
$arrCSV = @(Import-Csv -Path $strFixedCSVInputFilePath)

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
    'Attributes' = $false
}

$strNameOfColumnForFullPathOrPath = ''

$listUnmatchedHeaders = New-Object System.Collections.Generic.List[String]
$listRequiredHeadersNotFound = New-Object System.Collections.Generic.List[String]

if ($arrCSV.Count -ge 1) {
    (($arrCSV[0]).PSObject).Properties | ForEach-Object {
        $strHeaderElement = $_.Name
        if ($hashtableHeaderStatus.ContainsKey($strHeaderElement)) {
            $hashtableHeaderStatus.Item($strHeaderElement) = $true
        } else {
            $listUnmatchedHeaders.Add($strHeaderElement)
        }
        if ($strHeaderElement -eq 'Path') {
            if ([string]::IsNullOrEmpty($strNameOfColumnForFullPathOrPath) -eq $false) {
                Write-Warning ('Duplicate path column found. The header "' + $strHeaderElement + '" will be used for the full path. The conflicting column header is "' + $strNameOfColumnForFullPathOrPath + '".')
            }
            $strNameOfColumnForFullPathOrPath = $strHeaderElement
        } elseif ($strHeaderElement -eq 'Full Path') {
            if ([string]::IsNullOrEmpty($strNameOfColumnForFullPathOrPath) -eq $false) {
                Write-Warning ('Duplicate path column found. The header "' + $strNameOfColumnForFullPathOrPath + '" will be used for the full path. The conflicting column header is "' + $strHeaderElement + '".')
            } else {
                $strNameOfColumnForFullPathOrPath = $strHeaderElement
            }
        }
    }
}

$strPathSeparator = ''
$intIndex = 0
$intTotalRows = $arrCSV.Count
while ($strPathSeparator -eq '' -and $intIndex -lt $intTotalRows) {
    $strPath = ((($arrCSV[$intIndex]).PSObject).Properties) |
        Where-Object { $_.Name -eq $strNameOfColumnForFullPathOrPath } |
        Select-Object -ExpandProperty Value
    if ([string]::IsNullOrEmpty($strPath) -eq $false) {
        $strParentPath = Split-Path -Path $strPath -Parent
        if ([string]::IsNullOrEmpty($strParentPath) -eq $false) {
            if ($strParentPath.Length -gt 3) {
                # Not the root of a drive
                $strPathSeparator = $strPath.Substring($strParentPath.Length, 1)
            }
        }
    }
    $intIndex++
}
if ($strPathSeparator -eq '') {
    # Default to backslash
    $strPathSeparator = '\'
}
Write-Verbose ('The path separator is "' + $strPathSeparator + '".')

# Check $hashtableHeaderStatus for any headers with status $false and report them:
if (($hashtableHeaderStatus.Item('Full Path') -eq $false) -and ($hashtableHeaderStatus.Item('Path') -eq $false)) {
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

$hashtableRootElements = @{}
$hashtablePathsToFolderElements = @{}
$hashtableParentPathsToUnattachedChildDirectories = @{}
$hashtableParentPathsToUnattachedChildFiles = @{}

$intEmptyPathCounter = 0
$EMPTYSTRING = ''
$WILDCARDFILENAME = '*.*'
$MULTIPLEFILENAME = '[multiple]'
$FOLDERTYPE = 'Folder'
$VERSIONFIVE = [version]'5.0'
$VERSIONSIX = [version]'6.0'

$STRINGWITHSPACE = ' '
$SPACEOFSPACE = ' of '
$SPACELEFTPARENTHESIS = ' ('
$PERCENTRIGHTPARENTHESIS = '%)'
$sbCurrentOperation = New-Object -TypeName 'System.Text.StringBuilder'

#region Collect Stats/Objects Needed for Writing Progress ##########################
$intProgressReportingFrequency = 400
$intTotalItems = $intTotalRows
$strProgressActivity = 'Testing TreeSize CSV file for errors'
$strProgressStatus = 'Loading folder and file data into an in-memory tree'
$strProgressCurrentOperationPrefix = 'Processing item'
$boolReportOnUseCurrentOperation = $false
$timedateStartOfLoop = Get-Date
# Create a queue for storing lagging timestamps for ETA calculation
$queueLaggingTimestamps = New-Object System.Collections.Queue
$queueLaggingTimestamps.Enqueue($timedateStartOfLoop)
#endregion Collect Stats/Objects Needed for Writing Progress ##########################

Write-Verbose ($strProgressStatus + '...')

for ($intCounter = 0; $intCounter -lt $intTotalItems; $intCounter++) {
    #region Report Progress ########################################################
    $intCurrentItemNumber = $intCounter + 1 # Forward direction for loop
    if (($intCurrentItemNumber -ge 2000) -and ($intCurrentItemNumber % $intProgressReportingFrequency -eq 0)) {
        # Create a progress bar after the first 2000 items have been processed
        $timeDateLagging = $queueLaggingTimestamps.Dequeue()
        $datetimeNow = Get-Date
        $timespanTimeDelta = $datetimeNow - $timeDateLagging
        $intNumberOfItemsProcessedInTimespan = $intProgressReportingFrequency * ($queueLaggingTimestamps.Count + 1)
        $doublePercentageComplete = ($intCurrentItemNumber - 1) / $intTotalItems
        $intItemsRemaining = $intTotalItems - $intCurrentItemNumber + 1
        if ($boolReportOnUseCurrentOperation -eq $true) {
            [void]($sbCurrentOperation.Clear())
            [void]($sbCurrentOperation.Append($strProgressCurrentOperationPrefix))
            [void]($sbCurrentOperation.Append($STRINGWITHSPACE))
            [void]($sbCurrentOperation.Append($intCurrentItemNumber))
            [void]($sbCurrentOperation.Append($SPACEOFSPACE))
            [void]($sbCurrentOperation.Append($intTotalItems))
            [void]($sbCurrentOperation.Append($SPACELEFTPARENTHESIS))
            [void]($sbCurrentOperation.Append([string]::Format('{0:0.00}', ($doublePercentageComplete * 100))))
            [void]($sbCurrentOperation.Append($PERCENTRIGHTPARENTHESIS))
            [void]($strProgressCurrentOperationPrefix = $sbCurrentOperation.ToString())
            Write-Progress -Activity $strProgressActivity -Status $strProgressStatus -PercentComplete ($doublePercentageComplete * 100) -CurrentOperation ($sbCurrentOperation.ToString()) -SecondsRemaining (($timespanTimeDelta.TotalSeconds / $intNumberOfItemsProcessedInTimespan) * $intItemsRemaining)
        } else {
            Write-Progress -Activity $strProgressActivity -Status $strProgressStatus -PercentComplete ($doublePercentageComplete * 100) -SecondsRemaining (($timespanTimeDelta.TotalSeconds / $intNumberOfItemsProcessedInTimespan) * $intItemsRemaining)
        }
    }
    #endregion Report Progress ########################################################

    #region Extract Full Path and Type #############################################
    $boolMinimumElementsExtracted = $false

    try {
        $strFullPathOrPath = @((($arrCSV[$intCounter]).PSObject).Properties) |
            Where-Object { $_.Name -eq $strNameOfColumnForFullPathOrPath } |
            ForEach-Object { $_.Value }
    } catch {
        $strFullPathOrPath = $EMPTYSTRING
    }

    if (-not [string]::IsNullOrEmpty($strFullPathOrPath)) {
        # Path is not null or empty

        # Check to see if path is at least three characters long
        $boolWildcardFile = $false
        if ($strFullPathOrPath.Length -ge 3) {
            # Path is at least three characters long
            # Check to see if path ends in *.*
            # Note: in the following line,
            # $strFullPathOrPath.Substring($strFullPathOrPath.Length - 3, 3)
            # was refactored to
            # $strFullPathOrPath[-3..-1] -join ''
            if ($strFullPathOrPath[-3..-1] -join '' -eq $WILDCARDFILENAME) {
                # Path ends in *.*
                # This is a rollup of multiple file types that should be ignored
                # It seems older versions of TreeSize export this way no matter what
                $boolWildcardFile = $true
            }
        }

        if (-not $boolWildcardFile) {
            try {
                $strType = ($arrCSV[$intCounter]).Type
            } catch {
                $strType = $EMPTYSTRING
            }

            if (-not [string]::IsNullOrEmpty($strType)) {
                if ($strType -ne $MULTIPLEFILENAME) {
                    $boolMinimumElementsExtracted = $true

                    # ELSE:
                    # This is a rollup of multiple file types that should be ignored
                    # It seems older versions of TreeSize export this way no matter what
                }
            }
        }
    }
    #endregion Extract Full Path and Type #############################################

    if (-not $boolMinimumElementsExtracted) {
        if ([string]::IsNullOrEmpty($strFullPathOrPath)) {
            $intEmptyPathCounter++
        }
        if ($intEmptyPathCounter -eq 1) {
            $strMessage = 'Skipping the following row, which contains an empty path: ' + $arrCSV[$intCounter]
            if ($versionPS -ge $VERSIONFIVE) {
                Write-Information $strMessage
            } else {
                Write-Host $strMessage
            }
        } elseif ($strType -ne $MULTIPLEFILENAME -and $strFullPathOrPath.Length -ge 3) {
            $boolWildcardFile = $false
            # Path is at least three characters long
            # Check to see if path ends in *.*
            # Note: in the following line,
            # $strFullPathOrPath.Substring($strFullPathOrPath.Length - 3, 3)
            # was refactored to
            # $strFullPathOrPath[-3..-1] -join ''
            if ($strFullPathOrPath[-3..-1] -join '' -eq $WILDCARDFILENAME) {
                # Path ends in *.*
                # This is a rollup of multiple file types that should be ignored
                # It seems older versions of TreeSize export this way no matter what
                $boolWildcardFile = $true
            }

            if (-not $boolWildcardFile) {
                Write-Warning ('Failed to extract the minimum required elements from the following row (processing will continue, but this may indicate a problem with the TreeSize CSV file): ' + $arrCSV[$intCounter])
            }
        }
    } else {
        #region If The Type is a Folder, Ensure it Ends in Backslash ###############
        if ($strType -eq $FOLDERTYPE) {
            # Append a backslash if it's not already there
            if ($strFullPathOrPath.Substring($strFullPathOrPath.Length - 1, 1) -ne $strPathSeparator) {
                $strFullPathOrPath = $strFullPathOrPath + $strPathSeparator
            }
        }
        #endregion If The Type is a Folder, Ensure it Ends in Backslash ###############

        if ($boolDebug -eq $true) {
            Write-Debug ('Processing the following path: "' + $strFullPathOrPath + '"...')
        }

        #region Create the PSObjectTreeElement #####################################
        $refPSObjectTreeElement = [ref]$null
        if ($strType -eq $FOLDERTYPE) {
            if ($boolDebug -eq $true) {
                Write-Debug ('Creating a PSObjectTreeElement for a folder...')
            }
            $boolSuccess = New-PSObjectTreeDirectoryElement $refPSObjectTreeElement
        } else {
            # Assume file
            if ($boolDebug -eq $true) {
                Write-Debug ('Creating a PSObjectTreeElement for a file (type = "' + $strType + '")...')
            }
            $boolSuccess = New-PSObjectTreeFileElement $refPSObjectTreeElement
            if ($boolSuccess -eq $true) {
                if ($__TREEINCLUDETYPE -eq $true) {
                    ($refPSObjectTreeElement.Value).Type = $strType
                }
            }
        }
        #endregion Create the PSObjectTreeElement #####################################

        if ($boolSuccess -eq $false) {
            Write-Warning ('Failed to create a PSObjectTreeElement for the following row: ' + $arrCSV[$intCounter])
        } else {
            #region Store the Full Path ############################################
            ($refPSObjectTreeElement.Value).FullPath = $strFullPathOrPath
            if ($strType -eq $FOLDERTYPE) {
                # This is a folder
                if ($hashtablePathsToFolderElements.ContainsKey($strFullPathOrPath) -eq $true) {
                    # This path has already been added to the tree
                    Write-Warning ('The following path has already been added to the tree: ' + $strFullPathOrPath)
                    $hashtablePathsToFolderElements.Item($strFullPathOrPath) = $refPSObjectTreeElement
                } else {
                    # This path has not yet been added to the tree
                    if ($boolDebug -eq $true) {
                        Write-Debug ('Adding the element to the hashtable of paths to folder elements...')
                    }
                    $hashtablePathsToFolderElements.Add($strFullPathOrPath, $refPSObjectTreeElement)
                }
            }
            #endregion Store the Full Path ############################################

            #region Get and Store the Parent Path ##################################
            $strParentPath = Split-Path -Path $strFullPathOrPath -Parent
            if ([string]::IsNullOrEmpty($strParentPath) -eq $false) {
                # Append a backslash if it's not already there
                if ($strParentPath.Substring($strParentPath.Length - 1, 1) -ne $strPathSeparator) {
                    $strParentPath = $strParentPath + $strPathSeparator
                }
            }
            if ($boolDebug -eq $true) {
                Write-Debug ('Parent path = "' + $strParentPath + '"')
            }
            ($refPSObjectTreeElement.Value).ParentPath = $strParentPath
            #endregion Get and Store the Parent Path ##################################

            #region Get and Store This Object's Name ###############################
            $strName = Split-Path -Path $strFullPathOrPath -Leaf
            ($refPSObjectTreeElement.Value).Name = $strName
            #endregion Get and Store This Object's Name ###############################

            #region Attach to Existing Parent Element, or Store as Unattached ######
            if ([string]::IsNullOrEmpty($strParentPath) -eq $false) {
                # This path has an identifiable parent folder
                if ($hashtablePathsToFolderElements.ContainsKey($strParentPath) -eq $false) {
                    # This path's parent folder has not yet been added to the tree
                    $refToParentElement = [ref]$null
                    if ($strType -eq $FOLDERTYPE) {
                        # This is a folder
                        if ($hashtableParentPathsToUnattachedChildDirectories.ContainsKey($strParentPath) -eq $false) {
                            # New unattached parent folder
                            if ($versionPS -ge $VERSIONSIX) {
                                $hashtableParentPathsToUnattachedChildDirectories.Add($strParentPath, (New-Object System.Collections.Generic.List[ref]))
                            } else {
                                $hashtableParentPathsToUnattachedChildDirectories.Add($strParentPath, (New-Object System.Collections.ArrayList))
                            }
                        }
                        # Add this element to the existing list
                        if ($versionPS -ge $VERSIONSIX) {
                            ($hashtableParentPathsToUnattachedChildDirectories.Item($strParentPath)).Add($refPSObjectTreeElement)
                        } else {
                            [void](($hashtableParentPathsToUnattachedChildDirectories.Item($strParentPath)).Add($refPSObjectTreeElement))
                        }
                    } else {
                        # Assume file
                        if ($hashtableParentPathsToUnattachedChildFiles.ContainsKey($strParentPath) -eq $false) {
                            # New unattached parent folder
                            if ($versionPS -ge $VERSIONSIX) {
                                $hashtableParentPathsToUnattachedChildFiles.Add($strParentPath, (New-Object System.Collections.Generic.List[ref]))
                            } else {
                                $hashtableParentPathsToUnattachedChildFiles.Add($strParentPath, (New-Object System.Collections.ArrayList))
                            }
                        }
                        # Add this element to the existing list
                        if ($versionPS -ge $VERSIONSIX) {
                            ($hashtableParentPathsToUnattachedChildFiles.Item($strParentPath)).Add($refPSObjectTreeElement)
                        } else {
                            [void](($hashtableParentPathsToUnattachedChildFiles.Item($strParentPath)).Add($refPSObjectTreeElement))
                        }
                    }
                } else {
                    # This path's parent folder has already been added to the tree
                    $refToParentElement = $hashtablePathsToFolderElements.Item($strParentPath)
                    if ($strType -eq $FOLDERTYPE) {
                        # This is a folder
                        if ((($refToParentElement.Value).ChildDirectories).ContainsKey($strName) -eq $true) {
                            # This folder already exists in the parent folder
                            # This is a duplicate folder
                            # This is not allowed
                            Write-Warning ('Duplicate folder found: ' + $strFullPathOrPath)
                            (($refToParentElement.Value).ChildDirectories).Item($strName) = $refPSObjectTreeElement
                        } else {
                            # This folder does not already exist in the parent folder
                            # Add this folder to the parent folder
                            (($refToParentElement.Value).ChildDirectories).Add($strName, $refPSObjectTreeElement)
                        }
                    } else {
                        # Assume this is a file
                        if ($null -eq ($refToParentElement.Value).ChildFiles) {
                            Write-Warning ('The path "' + $strFullPathOrPath + '" has already been added to the tree. but the parent folder''s ChildFiles element is null. This should not be possible!')
                        } else {
                            if ((($refToParentElement.Value).ChildFiles).ContainsKey($strName) -eq $true) {
                                # This file already exists in the parent folder
                                # This is a duplicate file
                                # This is not allowed
                                Write-Warning ('Duplicate file found: ' + $strFullPathOrPath)
                                (($refToParentElement.Value).ChildFiles).Item($strName) = $refPSObjectTreeElement
                            } else {
                                # This file does not already exist in the parent folder
                                # Add this file to the parent folder
                                if ($boolDebug -eq $true) {
                                    Write-Debug('Adding file "' + $strName + '" to folder "' + $strParentPath + '"')
                                    Write-Debug('ChildFiles before adding: ' + (($refToParentElement.Value).ChildFiles).PSObject.Properties | Where-Object { $_.Name -eq 'Keys' } | ForEach-Object { $_.Value })
                                }
                                (($refToParentElement.Value).ChildFiles).Add($strName, $refPSObjectTreeElement)
                                if ($boolDebug -eq $true) {
                                    Write-Debug('ChildFiles after adding: ' + (($refToParentElement.Value).ChildFiles).PSObject.Properties | Where-Object { $_.Name -eq 'Keys' } | ForEach-Object { $_.Value })
                                }
                            }
                        }
                    }
                }
            } else {
                # This path has no identifiable parent folder
                $refToParentElement = [ref]$null
                # This path is the root of the tree
                if ($hashtableRootElements.ContainsKey($strFullPathOrPath) -eq $true) {
                    # This root path has already been added to the tree
                    Write-Warning ('The path "' + $strFullPathOrPath + '" has already been added as a root element. This should not be possible!')
                    $hashtableRootElements.Item($strFullPathOrPath) = $refPSObjectTreeElement
                } else {
                    # This root path has not yet been added to the tree
                    $hashtableRootElements.Add($strFullPathOrPath, $refPSObjectTreeElement)
                }
            }
            ($refPSObjectTreeElement.Value).ReferenceToParentElement = $refToParentElement
            #endregion Attach to Existing Parent Element, or Store as Unattached ######

            #region Attach to Already-Existing Children ############################
            # Start with folders/directories
            if ($hashtableParentPathsToUnattachedChildDirectories.ContainsKey($strFullPathOrPath) -eq $true) {
                # This path has unattached child directories
                $listChildElements = $hashtableParentPathsToUnattachedChildDirectories.Item($strFullPathOrPath)
                foreach ($refChildElement in $listChildElements) {
                    ($refChildElement.Value).ReferenceToParentElement = $refPSObjectTreeElement
                    if ((($refToParentElement.Value).ChildDirectories).ContainsKey(($refChildElement.Value).Name) -eq $true) {
                        # This folder already exists in the parent folder
                        # This is a duplicate folder
                        # This is not allowed
                        Write-Warning ('Duplicate folder found: ' + ($refChildElement.Value).FullPath)
                        (($refToParentElement.Value).ChildDirectories).Item(($refChildElement.Value).Name) = $refChildElement
                    } else {
                        # This folder does not already exist in the parent folder
                        # Add this folder to the parent folder
                        (($refToParentElement.Value).ChildDirectories).Add(($refChildElement.Value).Name, $refChildElement)
                    }
                }
                $hashtableParentPathsToUnattachedChildDirectories.Remove($strFullPathOrPath)
            }
            # Next, check files
            if ($hashtableParentPathsToUnattachedChildFiles.ContainsKey($strFullPathOrPath) -eq $true) {
                # This path has unattached child files
                $listChildElements = $hashtableParentPathsToUnattachedChildFiles.Item($strFullPathOrPath)
                foreach ($refChildElement in $listChildElements) {
                    ($refChildElement.Value).ReferenceToParentElement = $refPSObjectTreeElement
                    if ((($refToParentElement.Value).ChildFiles).ContainsKey(($refChildElement.Value).Name) -eq $true) {
                        # This file already exists in the parent folder
                        # This is a duplicate file
                        # This is not allowed
                        Write-Warning ('Duplicate file found: ' + ($refChildElement.Value).FullPath)
                        (($refToParentElement.Value).ChildFiles).Item(($refChildElement.Value).Name) = $refChildElement
                    } else {
                        # This file does not already exist in the parent folder
                        # Add this file to the parent folder
                        (($refToParentElement.Value).ChildFiles).Add(($refChildElement.Value).Name, $refChildElement)
                    }
                }
                $hashtableParentPathsToUnattachedChildFiles.Remove($strFullPathOrPath)
            }
            #endregion Attach to Already-Existing Children ############################

            if ($__TREEINCLUDESIZE -eq $true) {
                $strSize = $EMPTYSTRING
                $strSize = ($arrCSV[$intCounter]).Size
                if ([string]::IsNullOrEmpty($strSize) -eq $true) {
                    $strSize = $EMPTYSTRING
                    Write-Warning ('Unable to convert size "" to string for path "' + $strFullPathOrPath + '"')
                } else {
                    $int64Size = [uint64]0
                    $boolSuccess = Convert-TreeSizeSizeToUInt64Bytes ([ref]$int64Size) $strSize
                    if ($boolSuccess -ne $true) {
                        Write-Warning ('Unable to convert size "' + $strSize + '" to UInt64 bytes for path "' + $strFullPathOrPath + '"')
                    } else {
                        ($refPSObjectTreeElement.Value).SizeInBytesAsReportedByTreeSize = $int64Size
                    }
                }
            }

            if ($__TREEINCLUDEALLOCATED -eq $true) {
                $strAllocated = $EMPTYSTRING
                $strAllocated = ($arrCSV[$intCounter]).Allocated
                if ([string]::IsNullOrEmpty($strAllocated) -eq $true) {
                    $strAllocated = $EMPTYSTRING
                    Write-Warning ('Unable to convert disk allocation "" to string for path "' + $strFullPathOrPath + '"')
                } else {
                    $int64DiskAllocation = [uint64]0
                    $boolSuccess = Convert-TreeSizeSizeToUInt64Bytes ([ref]$int64DiskAllocation) $strAllocated
                    if ($boolSuccess -ne $true) {
                        Write-Warning ('Unable to convert disk allocation "' + $strAllocated + '" to UInt64 bytes for path "' + $strFullPathOrPath + '"')
                    } else {
                        ($refPSObjectTreeElement.Value).DiskAllocationInBytesAsReportedByTreeSize = $int64DiskAllocation
                    }
                }
            }

            if ($__TREEINCLUDELASTMODIFIED -eq $true) {
                $strLastModified = $EMPTYSTRING
                $strLastModified = ($arrCSV[$intCounter]).'Last Modified'
                if ([string]::IsNullOrEmpty($strLastModified) -eq $true) {
                    $strLastModified = $EMPTYSTRING
                    Write-Warning ('Unable to convert last modified date "" to string for path "' + $strFullPathOrPath + '"')
                } else {
                    # TODO: Create a function to do this conversion safely
                    $datetimeLastModified = [datetime]$strLastModified
                    ($refPSObjectTreeElement.Value).LastModified = $datetimeLastModified
                }
            }

            if ($__TREEINCLUDELASTACCESSED -eq $true) {
                $strLastAccessed = $EMPTYSTRING
                $strLastAccessed = ($arrCSV[$intCounter]).'Last Accessed'
                if ([string]::IsNullOrEmpty($strLastAccessed) -eq $true) {
                    $strLastAccessed = $EMPTYSTRING
                    Write-Warning ('Unable to convert last accessed date "" to string for path "' + $strFullPathOrPath + '"')
                } else {
                    # TODO: Create a function to do this conversion safely
                    $datetimeLastAccessed = [datetime]$strLastAccessed
                    ($refPSObjectTreeElement.Value).LastAccessed = $datetimeLastAccessed
                }
            }

            if ($__TREEINCLUDECREATIONDATE -eq $true) {
                $strCreationDate = $EMPTYSTRING
                $strCreationDate = ($arrCSV[$intCounter]).'Creation Date'
                if ([string]::IsNullOrEmpty($strCreationDate) -eq $true) {
                    $strCreationDate = $EMPTYSTRING
                    Write-Warning ('Unable to convert creation date "" to string for path "' + $strFullPathOrPath + '"')
                } else {
                    # TODO: Create a function to do this conversion safely
                    $datetimeCreationDate = [datetime]$strCreationDate
                    ($refPSObjectTreeElement.Value).CreationDate = $datetimeCreationDate
                }
            }

            if ($__TREEINCLUDEOWNER -eq $true) {
                $strOwner = $EMPTYSTRING
                $strOwner = ($arrCSV[$intCounter]).Owner
                if ([string]::IsNullOrEmpty($strOwner) -eq $true) {
                    $strOwner = $EMPTYSTRING
                    Write-Warning ('Unable to convert owner "" to string for path "' + $strFullPathOrPath + '"')
                } else {
                    if ($strOwner -eq 'The trust relationship between the primary domain and the trusted domain failed') {
                        Write-Warning ('Encountered a problem translating a SID to a principal name due to a failed client-domain trust relationship (or a failed domain-domain trust relationship) for path "' + $strFullPathOrPath + '"')
                    }
                    ($refPSObjectTreeElement.Value).Owner = $strOwner
                }
            }

            if ($__TREEINCLUDEPERMISSIONS -eq $true) {
                $strPermissions = $EMPTYSTRING
                $strPermissions = ($arrCSV[$intCounter]).Permissions
                if ([string]::IsNullOrEmpty($strPermissions) -eq $true) {
                    $strPermissions = $EMPTYSTRING
                    Write-Warning ('Unable to convert permissions "" to string for path "' + $strFullPathOrPath + '"')
                } else {
                    # Validate permissions
                    $strWarningMessage = $EMPTYSTRING
                    $boolSuccess = Test-TreeSizePermissionsRecord ([ref]$strWarningMessage) $strPermissions
                    if ($boolSuccess -eq $false) {
                        Write-Warning ('Encountered invalid permissions while processing path "' + $strFullPathOrPath + '": ' + $strWarningMessage)
                    } else {
                        ($refPSObjectTreeElement.Value).Permissions = $strPermissions
                    }
                }
            }

            if ($__TREEINCLUDEINHERITEDPERMISSIONS -eq $true) {
                $strInheritedPermissions = $EMPTYSTRING
                $strInheritedPermissions = ($arrCSV[$intCounter]).'Inherited Permissions'
                if ([string]::IsNullOrEmpty($strInheritedPermissions) -eq $true) {
                    $strInheritedPermissions = $EMPTYSTRING
                    Write-Warning ('Unable to convert inherited permissions "" to string for path "' + $strFullPathOrPath + '"')
                } else {
                    # Validate permissions
                    $strWarningMessage = $EMPTYSTRING
                    $boolSuccess = Test-TreeSizePermissionsRecord ([ref]$strWarningMessage) $strInheritedPermissions
                    if ($boolSuccess -eq $false) {
                        Write-Warning ('Encountered invalid inherited permissions while processing path "' + $strFullPathOrPath + '": ' + $strWarningMessage)
                    } else {
                        ($refPSObjectTreeElement.Value).InheritedPermissions = $strInheritedPermissions
                    }
                }
            }

            if ($__TREEINCLUDEOWNPERMISSIONS -eq $true) {
                $strOwnPermissions = $EMPTYSTRING
                $strOwnPermissions = ($arrCSV[$intCounter]).'Own Permissions'
                if ([string]::IsNullOrEmpty($strOwnPermissions) -eq $true) {
                    $strOwnPermissions = $EMPTYSTRING
                    Write-Warning ('Unable to convert own permissions "" to string for path "' + $strFullPathOrPath + '"')
                } else {
                    # Validate permissions
                    $strWarningMessage = $EMPTYSTRING
                    $boolSuccess = Test-TreeSizePermissionsRecord ([ref]$strWarningMessage) $strOwnPermissions
                    if ($boolSuccess -eq $false) {
                        Write-Warning ('Encountered invalid own permissions while processing path "' + $strFullPathOrPath + '": ' + $strWarningMessage)
                    } else {
                        ($refPSObjectTreeElement.Value).OwnPermissions = $strOwnPermissions
                    }
                }
            }

            if ($__TREEINCLUDEATTRIBUTES -eq $true) {
                $strAttributes = $EMPTYSTRING
                $strAttributes = ($arrCSV[$intCounter]).Attributes
                # Validate attributes
                $strWarningMessage = $EMPTYSTRING
                $boolSuccess = Test-TreeSizeAttributesRecord ([ref]$strWarningMessage) $strAttributes
                if ($boolSuccess -eq $false) {
                    Write-Warning ('Encountered problematic attribute(s) while processing path "' + $strFullPathOrPath + '": ' + $strWarningMessage)
                } else {
                    ($refPSObjectTreeElement.Value).Attributes = $strAttributes
                }
            }
        }
    }
    if ($intCounter % $intProgressReportingFrequency -eq 0) {
        # Add lagging timestamp to queue
        $queueLaggingTimestamps.Enqueue((Get-Date))
    }
}

# Set universal properties for size and disk allocation tallying
$boolRollUpSize = $true
$boolRollUpDiskAllocation = $false
$boolWarnWhenSizeDifferenceExceedsThreshold = $true
$boolReturnErrorWhenSizeDifferenceExceedsThreshold = $false
$boolWarnWhenDiskAllocationDifferenceExceedsThreshold = $true
$boolReturnErrorWhenDiskAllocationDifferenceExceedsThreshold = $false
$doubleThresholdInDecimal = $SizeToleranceThreshold

$strVerboseMessage = 'Rolling up file/folder sizes within root path '

# Process the root elements' size allocations if present
@($hashtableRootElements.Keys) | ForEach-Object {
    $strRootPath = $_
    $refPSObjectTreeRootElement = $hashtableRootElements.Item($strRootPath)
    Write-Verbose ($strVerboseMessage + '"' + $strRootPath + '"...')

    $uint64ThisFolderSizeInBytes = [uint64]0
    $uint64ThisFolderDiskAllocationInBytes = [uint64]0

    $boolSuccess = Add-RolledUpSizeOfTreeRecursively ([ref]$uint64ThisFolderSizeInBytes) ([ref]$uint64ThisFolderDiskAllocationInBytes) $refPSObjectTreeRootElement $boolRollUpSize $boolRollUpDiskAllocation $boolWarnWhenSizeDifferenceExceedsThreshold $boolReturnErrorWhenSizeDifferenceExceedsThreshold $boolWarnWhenDiskAllocationDifferenceExceedsThreshold $boolReturnErrorWhenDiskAllocationDifferenceExceedsThreshold $doubleThresholdInDecimal
    if ($boolSuccess -eq $false) {
        Write-Warning ('Encountered an error while rolling up size and disk allocation for root path "' + $strRootPath + '"')
    }
}

$strVerboseMessage = 'Rolling up top-level folder (not attached to parent folder) file/folder sizes '
@($hashtableParentPathsToUnattachedChildDirectories.Keys) | ForEach-Object {
    @($hashtableParentPathsToUnattachedChildDirectories.Item($_)) | ForEach-Object {
        $refPSObjectTreeRootElement = $_
        $strRootPath = ($refPSObjectTreeRootElement.Value).FullPath
        Write-Verbose ($strVerboseMessage + '"' + $strRootPath + '"...')

        $uint64ThisFolderSizeInBytes = [uint64]0
        $uint64ThisFolderDiskAllocationInBytes = [uint64]0

        $boolSuccess = Add-RolledUpSizeOfTreeRecursively ([ref]$uint64ThisFolderSizeInBytes) ([ref]$uint64ThisFolderDiskAllocationInBytes) $refPSObjectTreeRootElement $boolRollUpSize $boolRollUpDiskAllocation $boolWarnWhenSizeDifferenceExceedsThreshold $boolReturnErrorWhenSizeDifferenceExceedsThreshold $boolWarnWhenDiskAllocationDifferenceExceedsThreshold $boolReturnErrorWhenDiskAllocationDifferenceExceedsThreshold $doubleThresholdInDecimal
        if ($boolSuccess -eq $false) {
            Write-Warning ('Encountered an error while rolling up size and disk allocation for root path "' + $strRootPath + '"')
        }
    }
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
