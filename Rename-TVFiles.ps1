<#	
	.NOTES
	===========================================================================
	 Created with: 	Visual Studio Code 1.47.2
	 Created on:   	7/22/2020 11:09 AM
	 Created by:   	Jason Witters
	 Organization: 	Witters Inc.
	 Filename:     	Rename-TVFiles.ps1
	===========================================================================
	.DESCRIPTION
		Renames TV show files using information retrieved from TheTVDB API.
#>

param
(
    # Parent directory path for destination folders
    [Parameter(
        Mandatory = $false,
        Position = 0,
        ValueFromPipeline = $false)]
    [Alias("Source","Path")]
    [string]$SourcePath = "\\192.168.0.64\storage\Film\",

    # Parent directory path for individual download folders
    [Parameter(
        Mandatory = $false,
        Position = 1,
        ValueFromPipeline = $false)]
    [Alias("Downloads")]
    [string]$DownloadsDirectory = "\\192.168.0.64\storage\Film\_New",

    # Path to API key file, JSON format
    [Parameter(
        Mandatory = $false,
        Position = 2,
        ValueFromPipeline = $false)]
    [ValidatePattern("^.*\.json$")]
    [Alias("API","Key")]
    [string]$APIKey = "Z:\GitHub\TVDBKey.json"
)

# ScriptVersion = "1.0.10.1"

##################################
# Script Variables
##################################

$LoginURL = "https://api.thetvdb.com/login"
$SeriesSearchURL = "https://api.thetvdb.com/search/series?name="
$EpisodeSearchURL = "https://api.thetvdb.com/series/"
$EpisodeSearchString = "/episodes?page=1"
$global:WindowsFileNameRegex = '^(?:(?:[a-z]:|\\\\[a-z0-9_.$●-]+\\[a-z0-9_.$●-]+)\\|\\?[^\\\/:*?"<>|\r\n]+\\?)(?:[^\\\/:*?"<>|\r\n]+\\)*[^\\\/:*?"<>|\r\n]*$'
$global:StandardSeasonEpisodeFormatRegex = '(S|s)(\d{1,4})[ ]{0,1}(E|e|x|-)(\d{1,3})'
$global:SeasonRegex = '^(S|s)$'
$global:EpisodeRegex = '^(E|e|x|-)$'
$global:SeasonDigitRegex = '^\d{1,4}$'
$global:EpisodeDigitRegex = '^\d{1,3}$'
$TVRegex = "(?i)^0|TV$"
$AnimeRegex = "(?i)^1|Anime$"

##################################
# Script Functions
##################################

function Get-TVorAnimeDirectory {
    [CmdletBinding()]
    param (
        # Parent directory path
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $false)]
        [string]$SourcePath
    )
    
    begin
    {
        $TVorAnime = Read-Host "TV (0) or Anime (1)?"
    }
    
    process
    {
        if ($TVorAnime -match $TVRegex)
        {
            $FinalPath = Join-Path -Path $SourcePath -ChildPath "TV"
        }
        elseif ($TVorAnime -match $AnimeRegex)
        {
            $FinalPath = Join-Path -Path $SourcePath -ChildPath "Anime"
        }
        else
        {
            Write-Warning "You are not very smart!"
            Write-Warning "That was an invalid answer!"
            Write-Warning "As a reward, you have to start over now!"
            exit
        }
    }
    
    end
    {
        return $FinalPath
    }
}

function Get-APIToken {
    [CmdletBinding()]
    param (
        # Path to API key file, JSON format
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $false)]
        [string]$APIKey,
        # TheTVDB API login URL
        [Parameter(
            Mandatory = $true,
            Position = 1,
            ValueFromPipeline = $false)]
        [string]$LoginURL
    )
    
    begin
    {
        $ConvertedBody = Get-Content $APIKey -Raw
    }
    
    process
    {
        try
        {
            $token = Invoke-RestMethod -Method Post -Uri $LoginURL -Body $ConvertedBody -ContentType 'application/json' -ErrorAction Stop
            Write-Host "Successfully retrieved new API token"
        }
        catch
        {
            $Error[0]
            exit
        }
    }
    
    end
    {
        return $token.token
    }
}

function Get-TargetDirectory {
    [CmdletBinding()]
    param (
        # Parent directory path for individual download folders
        [Parameter(
            Mandatory = $false,
            Position = 0,
            ValueFromPipeline = $false)]
        [string]$DownloadsDirectory
    )
    
    begin
    {
        $SubFolders = Get-ChildItem -Path $DownloadsDirectory
    }
    
    process
    {
        if ($SubFolders.count -gt 1)
        {
            [int]$i = "0"

            Write-Host "`n"
            Write-Host "Folder Count: $($SubFolders.count)"
            Write-Host "`n"

            foreach ($SubFolder in $SubFolders)
            {
                Write-Host "$i - `"$($SubFolder.name)`""
                $i++
            }

            Write-Host "`n"
            $FolderNumber = Read-Host "Select folder"
            $TargetFolder = $SubFolders[$FolderNumber]
            Write-Host "Folder selected: `"$($TargetFolder.Name)`""
        }
        elseif ($SubFolders.count -eq 1)
        {
            $FolderNumber = "0"
        }
        else
        {
            Write-Warning "No folders in `"$DownloadsDirectory`" detected"
            exit
        }
    }
    
    end
    {
        $TargetFolder = $SubFolders[$FolderNumber]
        return $TargetFolder
    }
}

function Get-SeriesData {
    [CmdletBinding()]
    param (
        # Series Search String
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $false)]
        [string]$SeriesSearchString,

        # Series Search String
        [Parameter(
            Mandatory = $true,
            Position = 1,
            ValueFromPipeline = $false)]
        [string]$SeriesSearchURL,

        # API token string
        [Parameter(
            Mandatory = $true,
            Position = 2,
            ValueFromPipeline = $false)]
        [string]$APIToken
    )
    
    begin
    {
        $Headers = @{
            "ContentType" = "application/json"
            "Authorization" = "Bearer $APIToken"
        }

        $SeriesSearchURLTotal = $SeriesSearchURL + $SeriesSearchString
    }
    
    process
    {
        try
        {
            $SeriesData = Invoke-RestMethod -Method Get -Uri $SeriesSearchURLTotal -Headers $Headers -ErrorAction Stop

            if ($SeriesData.data.Count -gt 1)
            {
                [int]$i = "0"
                Write-Host "TVDB search returned $($SeriesSearchData.data.Count) results:"
                Write-Host "`n"
                foreach ($result in $SeriesData.data)
                {
                    Write-Host "$i - `"$($result.seriesName)`" ($($result.id))"
                    $i++
                }
                Write-Host "`n"
                $Number = Read-Host "Select correct series"
            }
            elseif ($SeriesData.data.Count -eq 1)
            {
                $Number = "0"
            }
            $SeriesName = $SeriesData.data[$Number].seriesName
            Write-Host "Received series data for: `"$SeriesName`""
        }
        catch
        {
            Write-Host $Error[0]
            exit
        }
    }
    
    end
    {
        return $SeriesData.data[$Number]
    }
}

function Get-EpisodeData {
    [CmdletBinding()]
    param (
        # Episode Search String
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $false)]
        [string]$EpisodeSearchString,

        # Episode Search URL
        [Parameter(
            Mandatory = $true,
            Position = 1,
            ValueFromPipeline = $false)]
        [string]$EpisodeSearchURL,

        # Series ID
        [Parameter(
            Mandatory = $true,
            Position = 2,
            ValueFromPipeline = $false)]
        [string]$SeriesID,
        
        # API token string
            [Parameter(
            Mandatory = $true,
            Position = 3,
            ValueFromPipeline = $false)]
        [string]$APIToken
    )
        
    begin
    {
        $Headers = @{
            "ContentType" = "application/json"
            "Authorization" = "Bearer $APIToken"
        }

        $EpisodeSearchURLTotal = $EpisodeSearchURL + $SeriesID + $EpisodeSearchString
    }
    
    process
    {
        try
        {
            $EpisodeData = Invoke-RestMethod -Method Get -Uri $EpisodeSearchURLTotal -Headers $Headers -ErrorAction Stop

            # For Series with more than 100 episodes, combine results from multiple pages
            if ($EpisodeData.links.last -gt 1)
            {
                for ([int]$i=2;$i -le $EpisodeData.links.last;$i++)
                {
                    $EpisodeSearchString = "/episodes?page=$i"
                    $EpisodeSearchURLTotal = $EpisodeSearchURL + $SeriesID + $EpisodeSearchString
                    try
                    {
                        $NewEpisodeData = Invoke-RestMethod -Method Get -Uri $EpisodeSearchURLTotal -Headers $Headers -ErrorAction Stop
                    }
                    catch
                    {
                        Write-Host $Error[0]
                        exit
                    }

                    $EpisodeData.data += $NewEpisodeData.data
                }
            }
        }
        catch
        {
            $Error[0]
            exit
        }
    }
    
    end
    {
        Write-Host "Received episode data for: `"$($SeriesSearchData.seriesName)`""
        Write-Host "Episode count: $($EpisodeData.data.Count)"
        Write-Host "`n"
        return $EpisodeData.data
    }
}

function Remove-BadFileTypes {
    [CmdletBinding()]
    param (
        # Target directory path
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true)]
        [string]$DirectoryPath
    )
    
    begin
    {
        $Files = Get-ChildItem -LiteralPath $DirectoryPath -Recurse -Force
    }
    
    process
    {
        foreach ($File in $Files)
        {
            # Delete .TXT .EXE .NFO files
            if ($File.name -match "(?i)\.(exe|nfo|txt)$")
            {
                Write-Host "Removing file: $($File.name)"
                Remove-Item -LiteralPath $File.fullname -Force
            }
        }
    }
    
    end {
        return
    }
}

function Remove-SubFolders {
    [CmdletBinding()]
    param (
        # Target directory path
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true)]
        [string]$DirectoryPath,

        # Destination directory path
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true)]
        [string]$DestinationFolderPath
    )
    
    begin
    {
        $Files = Get-ChildItem -LiteralPath $DirectoryPath -Recurse -Force
        Write-Host "Checking for subfolders at path: $DirectoryPath"
    }
    
    process
    {
        foreach ($File in $Files)
        {
            # Subfolders: move files to parent folder if folder named Subs\Subtitles\Season*
            if ($File.PSIsContainer -eq $true)
            {
                Write-Host "Subfolder found: $($File.name)"
                if (($File.Name -match "(?i)^(Subs|Subtitles)$") -or ($File.Name -match "(?i)^Season"))
                {
                    Write-Host "Acceptable subfolder detected: `"$($File.name)`""
                    $Subs = Get-ChildItem $File.fullname
                    foreach ($Subfile in $Subs)
                    {
                        Write-Host "Moving: $($Subfile.FullName)"
                        Write-Host "New path: $($TargetFolder.FullName)"
                        Move-Item -LiteralPath $subfile.fullname -Destination $TargetFolder.FullName -Force
                    }
                    # Delete Subfolder
                    Write-Host "Removing folder: $($File.FullName)"
                    Remove-Item -LiteralPath $File.FullName -Recurse -Force
                }
                # "Extras" subfolder
                elseif ($File.Name -match "(?i)^Extras$")
                {
                    $ExtrasFolder = Get-ChildItem -LiteralPath $File.FullName -Force
                    if ((($ExtrasFolder | Measure-Object).Count) -gt 0)
                    {
                        $ExtrasFolder
                        "`n"
                        Write-Host "0 - Move folder to destination folder"
                        Write-Host "1 - Move child items to current parent folder"
                        Write-Host "2 - Delete folder"
                        "`n"
                        $MoveExtras = Read-Host "Action to perform on `"Extras`" folder"

                        if ($MoveExtras -match "[0]")
                        {
                            Move-Item -LiteralPath $File.FullName -Destination $DestinationFolderPath -Force
                            Write-Host "Moving `"Extras`" folder to destination folder"
                        }
                        elseif ($MoveExtras -match "[1]")
                        {
                            Write-Host "Moving contents of `"Extras`" folder to parent directory"
                            foreach ($ExtrasItem in $ExtrasFolder)
                            {
                                Move-Item -LiteralPath $ExtrasItem.FullName -Destination $DirectoryPath -Force
                            }
                            # Delete empty folder after moving child items
                            if (((Get-ChildItem $File.FullName | Measure-Object).Count) -eq 0)
                            {
                                Write-Host "`"Extras`" folder is empty, deleting folder"
                                Remove-Item -LiteralPath $File.FullName -Recurse -Force
                            }
                        }
                        elseif ($MoveExtras -match "[2]")
                        {
                            Write-Host "Removing folder: $($File.FullName)"
                            Remove-Item -LiteralPath $File.FullName -Recurse -Force
                        }
                        else
                        {
                            Write-Host "You are not very smart!" -ForegroundColor Red
                            Write-Host "That was an invalid answer!" -ForegroundColor Red
                            Write-Host "As a reward, you have to start over now!" -ForegroundColor Red
                            exit
                        }
                    }
                }
                else
                {
                    Write-Host "Removing folder: $($File.FullName)"
                    Remove-Item -LiteralPath $File.FullName -Recurse -Force
                }
            }
        }
    }
    
    end {
        return
    }
}

function New-DestinationDirectory {
    [CmdletBinding()]
    param (
        # Series name from database
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $false)]
        [string]$SeriesSearchDataName,

        # Series name from directory path
        [Parameter(
            Mandatory = $true,
            Position = 1,
            ValueFromPipeline = $false)]
        [string]$TargetFolderName,

        # Selected media folder
        [Parameter(
            Mandatory = $true,
            Position = 2,
            ValueFromPipeline = $false)]
        [string]$TVDirectory
    )
    
    begin
    {
        # Determine destination folder/series name
        if ($SeriesSearchDataName -ne $TargetFolderName)
        {
            Write-Host "TVDB Series name does not match source folder name"
            Write-Host "Select correct series name:"
            Write-Host "0 - $SeriesSearchDataName"
            Write-Host "1 - $TargetFolderName"
            Write-Host "2 - Other"
            Write-Host "`n"
            $SeriesNameSelection = Read-Host "Selection"
            if ($SeriesNameSelection -eq "0")
            {
                $NewSeriesName = $SeriesSearchDataName
            }
            elseif ($SeriesNameSelection -eq "1")
            {
                $NewSeriesName = $TargetFolderName
            }
            elseif ($SeriesNameSelection -eq "2")
            {
                $NewSeriesName = Read-Host "Enter new series name"
            }
            else
            {
                Write-Host "Invalid response! Please type in correct series name" -ForegroundColor Yellow
                $NewSeriesName = Read-Host "Last chance moron"
            }
        }
        else
        {
            $NewSeriesName = $SeriesSearchDataName
        }
    }
    
    process
    {
        # Match series name to regex for windows file names
        if ($NewSeriesName -notmatch $WindowsFileNameRegex)
        {
            Write-Host "Series name does not conform to Windows file name rules" -ForegroundColor Yellow
            $TryAgainSeriesName = Read-Host "Enter in a different series name"
            if ($TryAgainSeriesName -match $WindowsFileNameRegex)
            {
                $DestinationFolderPath = Join-Path -Path $TVDirectory -ChildPath $TryAgainSeriesName
            }
            else
            {
                Write-Host "NewsSeries name STILL does not conform to Windows file name rules" -ForegroundColor Red
                Write-Host "Clearly you need time to think things over" -ForegroundColor Red
                Write-Host "Exiting..." -ForegroundColor Red
                exit
            }
        }
        else
        {
            $DestinationFolderPath = Join-Path -Path $TVDirectory -ChildPath $NewSeriesName
        }
    }
    
    end
    {
        # Create destination folder if doesn't exist
        if (!(Test-Path $DestinationFolderPath))
        {
            Write-Host "Destination folder `"$DestinationFolderPath`" does not exist"
            Write-Host "Creating destination folder `"$DestinationFolderPath`""
            try
            {
                New-Item -Path $TVDirectory -Name $NewSeriesName -ItemType Directory -ErrorAction Stop | Out-Null
                Write-Host "Successfully created destination folder: `"$DestinationFolderPath`""
            }
            catch
            {
                Write-Host "Failed to create destination folder: `"$DestinationFolderPath`"" -ForegroundColor Red
                exit
            }
        }
        else
        {
            Write-Host "Destination folder path detected: `"$DestinationFolderPath`""
        }
    
    # Return destination folder name and path
    $DestinationFolder = @{
        "Name" = $NewSeriesName
        "Path" = $DestinationFolderPath
    }
    return $DestinationFolder
    }
}

function Get-SeasonEpisodeNumbersFromString {
    [CmdletBinding()]
    param (
        # Source String
        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $false)]
        [string]$SourceString
    )
    
    begin
    {
        # NewName variable used as working file name
        $NewName = $SourceString
    }
    
    process
    {
        # Season and Episode number are in standard format
        if ($NewName -match $StandardSeasonEpisodeFormatRegex)
        {
            Write-Output "File name `"$NewName`" matches standard formatting"
        }
        # Season and Episode not in standard format
        else
        {
            Write-Warning "File name does NOT contain standard season and episode format"
            # Replace uncommon variations
            $NewName = $NewName.replace(' Chapter ','E')
            $NewName = $NewName.replace('_season_','S')
            $NewName = $NewName.replace(' Season ','S')
            $NewName = $NewName.Replace('_ep_','E')
            $NewName = $NewName.Replace(' Episode ','E')
            $NewName = $NewName.Replace('.Series.','S')
            $NewName = $NewName.Replace(' Series ','S')
            $NewName = $NewName.Replace('Series','S')

            # if file name matches ' 2x01 ' or ' 2e01 ' or [2x01] or [02x01] or [2e01], replace with S2E01
            if ($NewName -match ('(\[| )\d{1,4}(E|e|x|-)\d{1,3}(\]| )'))
            {
                $NewName = $NewName.Replace('[','S').replace(']','')
            }
            # if file name starts with 4 digits for season and episode number
            elseif ($NewName -match ('^\d{4}'))
            {
                $SeasonNumber = $NewName[0] + $NewName[1]
                $EpisodeNumber = $NewName[2] + $NewName[3]
            }
            #if file name starts with 3 digits for season and episode number
            elseif ($NewName -match ('^\d{3}'))
            {
                $SeasonNumber = "0" + $NewName[0]
                $EpisodeNumber = $NewName[1] + $NewName[2]
            }
        }
        # Parse season/episode numbers from file name if not already done
        if (!($SeasonNumber -or $EpisodeNumber))
        {
            Write-Verbose "Parsing Season\Episode numbers using standard format regex"
            $NewNameSplit = $NewName -split $StandardSeasonEpisodeFormatRegex

            # Parse out season/episode number
            for ($j=0;$j -lt ($NewNameSplit.count - 1); $j++)
            {
                if ($NewNameSplit[$j] -match $SeasonRegex)
                {
                    if ($NewNameSplit[($j + 1)] -match $SeasonDigitRegex)
                    {
                        $SeasonNumber = $NewNameSplit[$j + 1]
                    }
                }
                elseif ($NewNameSplit[$j] -match $EpisodeRegex)
                {
                    if ($NewNameSplit[($j + 1)] -match $EpisodeDigitRegex)
                    {
                        $EpisodeNumber = $NewNameSplit[$j + 1]
                    }
                }
            }
        }

        # if unable to parse season/episode number, prompt in console session
        if (!($SeasonNumber -or $EpisodeNumber))
        {
            Write-Warning "Episode or Season number not detected from file name"
            $Confirm = Read-Host "Input values [y/n]?"
            if ($Confirm -match "[yY]")
            {
                $SeasonNumber = Read-Host "Season Number"
                $EpisodeNumber = Read-Host "Episode Number"
            }
            else { exit }
        }

        ### Trim leading zeroes from season/episode numbers to match TheTVDB data

        # Double digit number with leading zero
        if ($SeasonNumber -match "^0[0-9]$")
        {
            $SeasonTrim = $SeasonNumber[1]
        }
        else
        {
            $SeasonTrim = $SeasonNumber
        }

        # Episode Number has one digit
        if ($EpisodeNumber -match "^\d{1}$")
        {
            $EpisodeTrim = $EpisodeNumber
        }
        # Two digits
        elseif ($EpisodeNumber -match "^\d{2}$")
        {
            # Double digit number with leading zero
            if ($EpisodeNumber -match "^0[0-9]$")
            {
                $EpisodeTrim = $EpisodeNumber[1]
            }
            # Double digit number with non-zero leading number
            else
            {
                $EpisodeTrim = $EpisodeNumber
            }
        }
        # Three digits
        elseif ($EpisodeNumber -match "^\d{3}$")
        {
            # Double leading zeroes
            if ($EpisodeNumber -match "^00[1-9]$")
            {
                $EpisodeTrim = $EpisodeNumber[2]
            }
            # Single leading zero
            elseif ($EpisodeNumber -match "^0[1-9][0-9]$")
            {
                $EpisodeTrim = $EpisodeNumber[1] + $EpisodeNumber[2]
            }
            else
            {
                $EpisodeTrim = $EpisodeNumber
            }
        }
    }
    
    end
    {
        Write-Host "Detected filename season number: $SeasonNumber"
        Write-Host "Detected filename episode number: $EpisodeNumber"
        Write-Host "Trimmed season number: $SeasonTrim"
        Write-Host "Trimmed episode number: $EpisodeTrim"
        
        # Return destination folder name and path
        $Numbers = @{
            "Season" = $SeasonNumber
            "Episode" = $EpisodeNumber
            "SeasonTrim" = $SeasonTrim
            "EpisodeTrim" = $EpisodeTrim
        }
        return $Numbers
    }
}

##################################
# Main
##################################

# Select correct media folder according to new media type
$TVDirectory = Get-TVorAnimeDirectory -SourcePath $SourcePath

# Determine target folder for processing
$TargetFolder = Get-TargetDirectory -DownloadsDirectory $DownloadsDirectory

# Get API token from TheTVDB
$APIToken = Get-APIToken -APIKey $APIKey -LoginURL $LoginURL

# Get series data from TheTVDB
$SeriesSearchData = Get-SeriesData -SeriesSearchString $TargetFolder.Name -SeriesSearchURL $SeriesSearchURL -APIToken $APIToken

# Get episode data from TheTVDB
$EpisodeData = Get-EpisodeData -EpisodeSearchString $EpisodeSearchString -EpisodeSearchURL $EpisodeSearchURL -SeriesID $SeriesSearchData.id -APIToken $APIToken

# Verify/create destination folder path
$DestinationFolder = New-DestinationDirectory -SeriesSearchDataName $SeriesSearchData.seriesName -TargetFolderName $TargetFolder.Name -TVDirectory $TVDirectory

# Eliminate subfolders in target folder
Remove-SubFolders -DirectoryPath $TargetFolder.FullName -DestinationFolderPath $DestinationFolder.Path

# Remove unnecessary files in target folder
Remove-BadFileTypes -DirectoryPath $TargetFolder.FullName

##################################
# Rename files
##################################

$Files = Get-ChildItem -LiteralPath $TargetFolder.FullName -Recurse

for ($i=0;$i -lt $Files.Count;$i++)
{
    $CurrentFile = $Files[$i]
    
    "`n"
    Write-Output "Parsing file: $($CurrentFile.name)"

    $NumbersFromFile = Get-SeasonEpisodeNumbersFromString -SourceString $CurrentFile.Name

    # Match episode information to TVDB data
    if ($EpisodeMatch) { Remove-Variable EpisodeMatch }
    $EpisodeMatch = $EpisodeData | Where-Object { ($_.airedepisodenumber -like $NumbersFromFile.EpisodeTrim) -and ($_.airedseason -like $NumbersFromFile.SeasonTrim) }
    if ($EpisodeMatch)
    {
        Write-Output "Episode data match found"
        Write-Output "Episode name: $($EpisodeMatch.episodeName)"
        Write-Output "DB Season number: $($EpisodeMatch.airedSeason)"
        Write-Output "DB Episode number: $($EpisodeMatch.airedEpisodeNumber)"
        $NewEpisodeName = $EpisodeMatch.episodeName.replace(":"," -")
        $NewEpisodeName = $NewEpisodeName.replace("/",", ")
        $NewEpisodeName = $NewEpisodeName.replace(" / ",", ")
        $NewEpisodeName = $NewEpisodeName.replace("`"","'")
        $NewEpisodeName = $NewEpisodeName.replace("\?", "")
        $NewEpisodeName = $NewEpisodeName.replace("?", "")
        Write-Output "Updated episode name: `"$NewEpisodeName`""

        if ($NewEpisodeName -notmatch $WindowsFileNameRegex)
        {
            Write-Warning "Episode name does NOT conform to Windows file name rules"
            $NewEpisodeName = Read-Host "Enter custom episode name"
        }
    
        $NewFileName = $DestinationFolder.Name + " - " + "S" + $NumbersFromFile.Season + "E" + $NumbersFromFile.Episode + " - " + $NewEpisodeName + $CurrentFile.extension
        $NewFilePath = Join-Path $DestinationFolder.Path -ChildPath $NewFileName
        Write-Output "New file name: `"$NewFileName`""
        Write-Output "Moving `"$($CurrentFile.Name)`" to $NewFilePath"
        Move-Item -LiteralPath $CurrentFile.FullName -Destination $NewFilePath -Force
    }
    else
    {
        Write-Warning "Could not find matching episode data from TVDB"
        Write-Warning "Season query: $($NumbersFromFile.SeasonTrim)"
        Write-Warning "Episode query: $($NumbersFromFile.EpisodeTrim)"
        $Confirm = Read-Host "Continue [y/n]?"
        if ($Confirm -match "[yY]")
        {
            Write-Output "Proceeding to next file"
        }
        else { exit }
    }
}