# Copyright (c) STÜBER SYSTEMS GmbH. All rights reserved.
# Licensed under the MIT License.

Import-LocalizedData -BindingVariable stringTable

<#
    .Synopsis
    Lists all terms of a given Untis .gpn file

    .Description
    This cmdlet reads all defined terms from the Untis .gpn file and lists them. If no term is defined nothing is listed.

    .Parameter File
    The file name of an Untis .gpn file.

    .Example
    Get-UntisTerms -File example.gpn
#>
function Get-UntisTerms {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $File
    )
    process
    {
        try
        {
            $UntisContent = Get-Content -Path $File -Raw

            $RegExPattern = '(?<prefix>0P\s*),"?(?<sname>[^,"]*)"?,"?(?<lname>[^,"]*)"?,(?<fdate>\d*),(?<tdate>\d*),"?(?<pname>[^,"]*)"?,(?<status>A?)'
            $RegExMatches = [regex]::matches($UntisContent, $RegExPattern)

            $table = @()

            foreach($match in $RegExMatches)
            {
                $term = New-Object PSCustomObject
                $term | Add-Member -type NoteProperty -name ShortName -Value $match.Groups["sname"].value
                $term | Add-Member -type NoteProperty -name LongName -Value $match.Groups["lname"].value
                $term | Add-Member -type NoteProperty -name FromDate -Value (FormatDate $match.Groups["fdate"].value)
                $term | Add-Member -type NoteProperty -name ToDate -Value (FormatDate $match.Groups["tdate"].value)
                $term | Add-Member -type NoteProperty -name Active -Value (FormatStatus $match.Groups["status"].value)
                $table += $term
            }

            return $table
        }
        catch
        {
            $ErrorMessage = $_.Exception.Message
            Write-Host $ErrorMessage -ForegroundColor Red
        }
    }
}

<#
    .Synopsis
    Activates a choosen term in a given Untis .gpn file

    .Description
    This cmdlet changes the active term within a given Untis .gpn file. This cmdlet will modify the gpn file. Please be carefull while
    testing and always backup first.

    .Parameter File
    The file name of an Untis .gpn file.

    .Parameter ShortName
    Name of the term.

    .Example
    Set-UntisTermAsActive -File example.gpn -ShortName Periode1
    .Example
    Set-UntisTermAsActive -File example.gpn -ShortName 'Periode 2'
#>
function Set-UntisTermAsActive {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $File,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $ShortName
    )
    process
    {
        try
        {
            $UntisContent = Get-Content -Path $File -Raw

            $RegExPattern = '(?<prefix>0P\s*),"?(?<sname>[^,"]*)"?,"?(?<lname>[^,"]*)"?,(?<fdate>\d*),(?<tdate>\d*),"?(?<pname>[^,"]*)"?,A'
            $UntisContent = $UntisContent -Replace $RegExPattern, '${prefix},"${sname}","${lname}",${fdate},${tdate},"${pname}",'

            $RegExPattern = '(?<prefix>0P\s*),"?(?<sname>' + $ShortName + ')"?,"?(?<lname>[^,"]*)"?,(?<fdate>\d*),(?<tdate>\d*),"?(?<pname>[^,"]*)"?,'
            $UntisContent = $UntisContent -Replace $RegExPattern, '${prefix},"${sname}","${lname}",${fdate},${tdate},"${pname}",A'

            Set-Content -Path $File -Value $UntisContent
        }
        catch
        {
            $ErrorMessage = $_.Exception.Message
            Write-Host $ErrorMessage -ForegroundColor Red
        }
    }
}

<#
    .Synopsis
    Exports data from an Untis file (.gpn or .untis) or from an Untis database (Untis MultiUser)

    .Description
    This cmdlet exports data from an Untis file (.gpn or .untis) or from an Untis database (Untis MultiUser).
	Supported are XML export and GPU export. You can export single GPU files (e.g. GPU001.TXT) or all GPU files at once.

    .Parameter File
    The file name of an Untis file (.gpn or .untis). This parameter is mandatory for file access

    .Parameter SchoolNo
    The school no for accessing the Untis database. This parameter is mandatory for database access

    .Parameter SchoolYear
    The school year for accessing the Untis database. This parameter is mandatory for database access

    .Parameter Version
    The schedule version for accessing the Untis database. This parameter is optional. By default its version 1.

    .Parameter User
    The user name for accessing the Untis database. This parameter is mandatory for database access

    .Parameter Password
    The password as securestring for accessing the Untis database. This parameter is mandatory for database access.

    .Parameter Date
    A date within a defined Untis term. This paremeter is optional and only needed if the Untis file (.gpn or .untis) or the
	Untis database defines more than one term.

    .Parameter BackupFile
    The file name for the Untis backup file (.gpn or .untis). This paremeter is optional. By default a generated file name
	is used.

    .Parameter OutputFolder
    The folder in which all export files are created. This paremeter is optional. By default the current folder is used.

    .Parameter OutputType
    The type of export data. Possible values are GPU, GPU001, ..., GPU021 and XML. Since this parameter is definend as array
    you can freely combine the values.

    .Example
    Start-UntisExport -File example.gpn -OutputFolder c:\output -OutputType GPU
    .Example
    Start-UntisExport -File example.gpn -OutputFolder c:\output -OutputType GPU001
    .Example
    Start-UntisExport -File example.gpn -OutputFolder c:\output -OutputType XML
    .Example
    Start-UntisExport -File example.gpn -OutputFolder c:\output -OutputType GPU001,GPU002
    .Example
    Start-UntisExport -File example.gpn -Date 2009-07-11 -OutputFolder c:\output -OutputType GPU,XML
    .Example
    Start-UntisExport -SchoolNo 40042 -SchoolYear 2020-2021 -User Administrator -Password (ConvertTo-SecureString qwertz -AsPlainText -Force) -OutputFolder c:\output -OutputType GPU
    .Example
    Start-UntisExport -SchoolNo 40042 -SchoolYear 2020-2021 -User Administrator -Password (ConvertTo-SecureString qwertz -AsPlainText -Force) -Date 2009-07-11 -OutputFolder c:\output -OutputType GPU001,GPU002
#>
function Start-UntisExport {
    param(
        [Parameter(Mandatory=$true, ParameterSetName='File')]
        [ValidateNotNullOrEmpty()]
        [string]
        $File,
        [Parameter(Mandatory=$true, ParameterSetName='Database')]
        [ValidateNotNullOrEmpty()]
        [string]
        $SchoolNo,
        [Parameter(Mandatory=$true, ParameterSetName='Database')]
        [ValidateNotNullOrEmpty()]
        [string]
        $SchoolYear,
        [Parameter(ParameterSetName='Database')]
        [ValidateNotNullOrEmpty()]
        [int]
        $Version = 1,
        [Parameter(Mandatory=$true, ParameterSetName='Database')]
        [ValidateNotNullOrEmpty()]
        [string]
        $User,
        [Parameter(Mandatory=$true, ParameterSetName='Database')]
        [ValidateNotNullOrEmpty()]
        [securestring]
        $Password,
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [DateTime]
        $Date = 0,
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $BackupFile,
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $OutputFolder,
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [OutputType[]]
        $OutputType
    )
    process
    {
        try
        {
            $Stopwatch =  [system.diagnostics.stopwatch]::StartNew()

            Write-Host ([string]::Format($stringTable.StartExport, [Environment]::NewLine, [Environment]::NewLine))

            if (-not $BackupFile)
            {
                $BackupFile = Join-Path -Path $OutputFolder -ChildPath "untis.backup"
            }

            if ($File)
            {
                Copy-Item $File -Destination $BackupFile
            }
            else
            {
                RunUntisBackup -SchoolNo $SchoolNo -SchoolYear $SchoolYear -Version $Version -BackupFile $BackupFile -User $User -Password $Password
            }

            if ($OutputType.Contains([OutputType]::GPU))
            {
                RunUntisGpuExport -File $BackupFile -Date $Date -OutputFolder $OutputFolder
            }
            else
            {
                $gpuTable = @(
                    @([OutputType]::GPU001, "001"),
                    @([OutputType]::GPU002, "002"),
                    @([OutputType]::GPU003, "003"),
                    @([OutputType]::GPU004, "004"),
                    @([OutputType]::GPU005, "005"),
                    @([OutputType]::GPU006, "006"),
                    @([OutputType]::GPU007, "007"),
                    @([OutputType]::GPU008, "008"),
                    @([OutputType]::GPU009, "009"),
                    @([OutputType]::GPU010, "010"),
                    @([OutputType]::GPU011, "011"),
                    @([OutputType]::GPU012, "012"),
                    @([OutputType]::GPU013, "013"),
                    @([OutputType]::GPU014, "014"),
                    @([OutputType]::GPU015, "015"),
                    @([OutputType]::GPU016, "016"),
                    @([OutputType]::GPU017, "017"),
                    @([OutputType]::GPU018, "018"),
                    @([OutputType]::GPU019, "019"),
                    @([OutputType]::GPU020, "020"),
                    @([OutputType]::GPU021, "021"),
                    @([OutputType]::GPU022, "022"),
                    @([OutputType]::GPU023, "023")
                )
                foreach($gpu in $gpuTable)
                {
                    if ($OutputType.Contains($gpu[0]))
                    {
                        RunUntisSingleGpuExport -File $BackupFile -Date $Date -OutputFolder $OutputFolder -OutputType $gpu[1]
                    }
                }
            }
            if ($OutputType.Contains([OutputType]::XML))
            {
                $XmlFile = Join-Path -Path $OutputFolder -ChildPath "untis.xml"
                RunUntisXmlExport -File $BackupFile -Date $Date -XmlFile $XmlFile
            }

            $Stopwatch.Stop()

            Write-Host ([string]::Format($stringTable.TimeElapsed, [Environment]::NewLine, $Stopwatch.Elapsed))
        }
        catch
        {
            $ErrorMessage = $_.Exception.Message
            Write-Host $ErrorMessage -ForegroundColor Red
        }
    }
}

<#
    .Synopsis
    Creates an Untis backup file (.gpn or .untis) from an Untis database (Untis MultiUser)

    .Description
    This cmdlet exports all data from an Untis database (Untis MultiUser) to a newly created Untis file (.gpn or .untis).

    .Parameter SchoolNo
    The school no for accessing the Untis database. This parameter is mandatory.

    .Parameter SchoolYear
    The school year for accessing the Untis database. This parameter is mandatory.

    .Parameter Version
    The schedule version for accessing the Untis database. This parameter is optional. By default its version 1.

    .Parameter User
    The user name for accessing the Untis database. This parameter is mandatory.

    .Parameter Password
    The password as securestring for accessing the Untis database. This parameter is mandatory.

    .Parameter BackupFile
    The file name for the Untis backup file (.gpn or .untis). This paremeter is mandatory.

    .Example
    Start-UntisBackup -SchoolNo 40042 -SchoolYear 2020-2021 -User Administrator -Password (ConvertTo-SecureString qwertz -AsPlainText -Force) -BackupFile c:\output\backup.gpn
    .Example
    Start-UntisBackup -SchoolNo 40042 -SchoolYear 2021-2022 -User Administrator -Password (ConvertTo-SecureString qwertz -AsPlainText -Force) -BackupFile c:\output\backup.untis
#>
function Start-UntisBackup{
    param(
        [Parameter(Mandatory=$true, ParameterSetName='Database')]
        [ValidateNotNullOrEmpty()]
        [string]
        $SchoolNo,
        [Parameter(Mandatory=$true, ParameterSetName='Database')]
        [ValidateNotNullOrEmpty()]
        [string]
        $SchoolYear,
        [Parameter(ParameterSetName='Database')]
        [ValidateNotNullOrEmpty()]
        [int]
        $Version = 1,
        [Parameter(Mandatory=$true, ParameterSetName='Database')]
        [ValidateNotNullOrEmpty()]
        [string]
        $User,
        [Parameter(Mandatory=$true, ParameterSetName='Database')]
        [ValidateNotNullOrEmpty()]
        [securestring]
        $Password,
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $BackupFile
    )
    process
    {
        try
        {
            $Stopwatch =  [system.diagnostics.stopwatch]::StartNew()

            Write-Host ([string]::Format($stringTable.StartBackup, [Environment]::NewLine, [Environment]::NewLine))

            RunUntisBackup -SchoolNo $SchoolNo -SchoolYear $SchoolYear -Version $Version -User $User -Password $Password -BackupFile $BackupFile

            $Stopwatch.Stop()

            Write-Host ([string]::Format($stringTable.TimeElapsed, [Environment]::NewLine, $Stopwatch.Elapsed))
        }
        catch
        {
            $ErrorMessage = $_.Exception.Message
            Write-Host $ErrorMessage -ForegroundColor Red
        }
    }
}

function RunUntisGpuExport{
    param(
        [string]
        $File,
        [DateTime]
        $Date,
        [string]
        $OutputFolder
    )
    process
    {
        $ConsolePath = GetUntisConsolePath

        if (($ConsolePath) -and (Test-Path -Path $ConsolePath -PathType Leaf))
        {
            $UntisGpuFilePattern = Join-Path -Path $OutputFolder -ChildPath "GPU???.txt"

            Remove-Item $UntisGpuFilePattern

            if ($Date -gt 0)
            {
                Start-Process -FilePath $ConsolePath -ArgumentList "$($File)","/exp*=""$($OutputFolder)""","/date=$($Date.ToString("yyyyMMdd"))" -Wait
            }
            else
            {
                Start-Process -FilePath $ConsolePath -ArgumentList "$($File)","/exp*=""$($OutputFolder)""" -Wait
            }

            $UntisGpuFileCount = (Get-ChildItem -Path (Join-Path -Path $OutputFolder -ChildPath "*") -Include "GPU???.txt" | Measure-Object).Count

            if ($UntisGpuFileCount -gt 0)
            {
                Write-Host ([string]::Format($stringTable.GpuFilesExported, $UntisGpuFileCount, $OutputFolder))
            }
            else
            {
                $ErrorMessage = ([string]::Format($stringTable.ErrorGpuFilesNotCreated, $OutputFolder))
                throw $ErrorMessage
            }
        }
        else
        {
            $ErrorMessage = ([string]::Format($stringTable.ErrorUntisConsoleNotFound, $ConsolePath))
            throw $ErrorMessage
        }
    }
}

function RunUntisSingleGpuExport{
    param(
        [string]
        $File,
        [DateTime]
        $Date,
        [string]
        $OutputFolder,
        [string]
        $OutputType
    )
    process
    {
        $ConsolePath = GetUntisConsolePath

        if (($ConsolePath) -and (Test-Path -Path $ConsolePath -PathType Leaf))
        {
            $UntisGpuFile = Join-Path -Path $OutputFolder -ChildPath "GPU$($OutputType).TXT"

            if (Test-Path -Path $UntisGpuFile -PathType Leaf)
            {
                Remove-Item $UntisGpuFile
            }

            if ($Date -gt 0)
            {
                Start-Process -FilePath $ConsolePath -ArgumentList "$($File)","/exp$($OutputType)=""$($UntisGpuFile)""","/date=$($Date.ToString("yyyyMMdd"))" -Wait
            }
            else
            {
                Start-Process -FilePath $ConsolePath -ArgumentList "$($File)","/exp$($OutputType)=""$($UntisGpuFile)""" -Wait
            }

            if (Test-Path -Path $UntisGpuFile -PathType Leaf)
            {
                Write-Host ([string]::Format($stringTable.GpuFileExported, $UntisGpuFile))
            }
            else
            {
                $ErrorMessage = ([string]::Format($stringTable.ErrorGpuFileNotCreated, $UntisGpuFile))
                throw $ErrorMessage
            }
        }
        else
        {
            $ErrorMessage = ([string]::Format($stringTable.ErrorUntisConsoleNotFound, $ConsolePath))
            throw $ErrorMessage
        }
    }
}

function RunUntisXmlExport{
    param(
        [string]
        $File,
        [DateTime]
        $Date,
        [string]
        $XmlFile
    )
    process
    {
        $ConsolePath = GetUntisConsolePath

        if (($ConsolePath) -and (Test-Path -Path $ConsolePath -PathType Leaf))
        {
            if (Test-Path -Path $XmlFile -PathType Leaf)
            {
                Remove-Item $XmlFile
            }

            if ($Date -gt 0)
            {
                Start-Process -FilePath $ConsolePath -ArgumentList "$($File)","/xml=""$($XmlFile)""","/date=$($Date.ToString("yyyyMMdd"))" -Wait
            }
            else
            {
                Start-Process -FilePath $ConsolePath -ArgumentList "$($File)","/xml=""$($XmlFile)""" -Wait
            }

            if (Test-Path -Path $XmlFile -PathType Leaf)
            {
                Write-Host ([string]::Format($stringTable.XmlFileExported, $XmlFile))
            }
            else
            {
                $ErrorMessage = ([string]::Format($stringTable.ErrorXmlFileNotCreated, $XmlFile))
                throw $ErrorMessage
            }
        }
        else
        {
            $ErrorMessage = ([string]::Format($stringTable.ErrorUntisConsoleNotFound, $ConsolePath))
            throw $ErrorMessage
        }
    }
}

function RunUntisBackup{
    param(
        [string]
        $SchoolNo,
        [string]
        $SchoolYear,
        [int]
        $Version,
        [string]
        $User,
        [securestring]
        $Password,
        [string]
        $BackupFile
    )
    process
    {
        $ConsolePath = GetUntisConsolePath

        if (($ConsolePath) -and (Test-Path -Path $ConsolePath -PathType Leaf))
        {
            if (Test-Path -Path $BackupFile -PathType Leaf)
            {
                Remove-Item $BackupFile
            }

            if (IsPS7OrHigher)
            {
                $Pswd = ConvertFrom-SecureString -SecureString $Password -AsPlainText
            }
            else {
                $Pswd = ConvertFrom-SecureStringToPlainText -SecureString $Password
            }

            Start-Process -FilePath $ConsolePath -ArgumentList "DB~$($SchoolNo)~$($SchoolYear)~$($Version)","/backup=$($BackupFile)","/user=$($User)","/pw=$($Pswd)" -Wait

            if (Test-Path -Path $BackupFile -PathType Leaf)
            {
                Write-Host ([string]::Format($stringTable.BackupFileCreated, $BackupFile))
            }
            else
            {
                $ErrorMessage = ([string]::Format($stringTable.ErrorBackupFileNotCreated, $BackupFile))
                throw $ErrorMessage
            }
        }
        else
        {
            $ErrorMessage = ([string]::Format($stringTable.ErrorUntisConsoleNotFound, $ConsolePath))
            throw $ErrorMessage
        }
    }
}

function GetUntisConsolePath {
    process
    {
        $Versions = @("2023","2022","2021","2020","2019","2018","2017")

        for ($i=0; $i -lt $Versions.Count; $i++)
        {
            $RegKey64 = "HKLM:\SOFTWARE\WOW6432Node\Gruber&Petters\Untis $($Versions[$i])"
            $RegKey32 = "HKLM:\SOFTWARE\Gruber&Petters\Untis $($Versions[$i])"

            if ([Environment]::Is64BitProcess)
            {
                if (Test-Path -Path $RegKey64)
                {
                    $RegKey = Get-ItemProperty -Path $RegKey64 -Name Install_Dir
                    return Join-Path -Path $RegKey.Install_Dir -ChildPath "Untis.exe"
                }
            }
            else
            {
                if (Test-Path -Path $RegKey32)
                {
                    $RegKey = Get-ItemProperty -Path $RegKey32 -Name Install_Dir
                    return Join-Path -Path $RegKey.Install_Dir -ChildPath "Untis.exe"
                }
            }
        }
        return "Untis.exe"
    }
}

function FormatDate {
    param(
        [string]
        $DateNumber
    )
    process
    {
        return [datetime]::parseexact($DateNumber, 'yyyyMMdd', $null).ToString('dd-MM-yyyy')
    }
}

function FormatStatus {
    param(
        [string]
        $Status
    )
    process
    {
        return $Status.Contains("A")
    }
}

# List of supported export formats
Enum OutputType {
    GPU
    GPU001
    GPU002
    GPU003
    GPU004
    GPU005
    GPU006
    GPU007
    GPU008
    GPU009
    GPU010
    GPU011
    GPU012
    GPU013
    GPU014
    GPU015
    GPU016
    GPU017
    GPU018
    GPU019
    GPU020
    GPU021
    GPU022
    GPU023
    XML
}

# Fallback for missing ConvertFrom-SecureString -AsPlainText -Force in PowerShell 5.1
function ConvertFrom-SecureStringToPlainText{
    param(
        [SecureString]
        $SecureString
    )
    process
    {
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        return [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    }
}

# Helper for detecting PowerShell 7+ Version
function IsPS7OrHigher{
    process
    {
        return ($PSVersionTable.PSVersion.Major -ge 7)
    }
}

Export-ModuleMember -Function Get-UntisTerms
Export-ModuleMember -Function Set-UntisTermAsActive
Export-ModuleMember -Function Start-UntisBackup
Export-ModuleMember -Function Start-UntisExport
