﻿<#	
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
    # Source directory path
    [Parameter(
        Mandatory = $false,
        Position = 0,
        ValueFromPipeline = $true)]
    [Alias("Source","Path")]
    [string]$SourcePath = "\\192.168.0.64\storage\Film\_New\",

    # Path to API key file, JSON format
    [Parameter(
        Mandatory = $false,
        Position = 1,
        ValueFromPipeline = $false)]
    [ValidatePattern("^.*\.json$")]
    [Alias("API","Key")]
    [string]$APIKey = "Z:\GitHub\TVDBKey.json",
    
    # Sets VerbosePreference variable to 'Continue', displaying Verbose output in console
    [Parameter(
        Mandatory = $false,
        Position = 1)]
	[switch]$VerboseOutput
)

#ScriptVersion = "1.0.0.0"

##################################
#Script Variables
##################################

if ($VerboseOutput)
{
    $VerbosePreference = "Continue"
}
else
{
    $VerbosePreference = "SilentlyContinue"
}

$ConvertedBody = Get-Content $APIKey -Raw
$LoginURL = "https://api.thetvdb.com/login"
$SeriesSearchURL = "https://api.thetvdb.com/search/series?name="
$Server = "192.168.0.64"
$WindowsFileNameRegex = '^(?:(?:[a-z]:|\\\\[a-z0-9_.$●-]+\\[a-z0-9_.$●-]+)\\|\\?[^\\\/:*?"<>|\r\n]+\\?)(?:[^\\\/:*?"<>|\r\n]+\\)*[^\\\/:*?"<>|\r\n]*$'
$StandardSeasonEpisodeFormatRegex = '(S|s)(\d{1,4})[ ]{0,1}(E|e|x|-)(\d{1,3})'
$SeasonRegex = '^(S|s)$'
$EpisodeRegex = '^(E|e|x|-)$'
$SeasonDigitRegex = '^\d{1,4}$'
$EpisodeDigitRegex = '^\d{1,3}$'
$SubFolders = Get-ChildItem -Path $SourcePath
$TVorAnime = Read-Host "TV (0) or Anime (1)?"
if (($TVorAnime -like "TV") -or ($TVorAnime -like "0"))
{
    $TVDirectory = "\\$server\storage\Film\TV"
}
elseif (($TVorAnime -like "Anime") -or ($TVorAnime -like "1"))
{
    $TVDirectory = "\\$server\storage\Film\Anime"
}
else
{
    Write-Warning "You are not very smart!"
    Write-Warning "That was an invalid answer!"
    Write-Warning "As a reward, you have to start over now!"
    exit
}

#Validate correct source folder
if ($SubFolders.count -gt 1)
{
    [int]$i = "0"
    "`n"
    Write-Output "Folder Count: $($SubFolders.count)"
    "`n"
    foreach ($SubFolder in $SubFolders)
    {
        Write-Output "$i - `"$($SubFolder.name)`""
        $i++
    }
    "`n"
    $FolderNumber = Read-Host "Select folder"
    $TargetFolder = $SubFolders[$FolderNumber]
    Write-Output "Processing folder: `"$($TargetFolder.Name)`""
}
elseif ($SubFolders.count -eq 1)
{
    $FolderNumber = "0"
    $Confirm = Read-Host "Use folder: `"$($TargetFolder.Name)`" [y/n]?"
    if ($Confirm -match "[yY]")
    {
        Write-Output "Processing folder: `"$($TargetFolder.Name)`""
    }
    else
    {
        exit
    }
}
else
{
    Write-Warning "No folders in `"$SourcePath`" detected"
    exit
}

$FolderName = $SubFolders[$FolderNumber].Name
$SeriesSearchString = $FolderName
$SeriesSearchURLTotal = $SeriesSearchURL + $SeriesSearchString

##################################
#Get API token
##################################

try
{
    $token = Invoke-RestMethod -Method Post -Uri $LoginURL -Body $ConvertedBody -ContentType 'application/json' -ErrorAction Stop
    Write-Output "Successfully retrieved new API token"
}
catch
{
    Write-Warning "Error retrieving API token from `'api.thetvdb.com`'"
    $Error[0]
    exit
}
finally
{
    $Headers = @{
        "ContentType" = "application/json"
        "Authorization" = "Bearer $($token.token)"
    }
}

##################################
#Get Series data
##################################

try
{
    $SeriesSearchData = Invoke-RestMethod -Method Get -Uri $SeriesSearchURLTotal -Headers $Headers -ErrorAction Stop
}
catch
{
    Write-Warning "Error searching for TV series: $SeriesSearchString"
    $Error[0]
    exit
}

if ($SeriesSearchData.data.Count -gt 1)
{
    [int]$i = "0"
    Write-Output "TVDB search returned $($SeriesSearchData.data.Count) results:"
    "`n"
    foreach ($result in $SeriesSearchData.data)
    {
        Write-Output "$i - `"$($result.seriesName)`" ($($result.id))"
        $i++
    }
    "`n"
    $Number = Read-Host "Select correct series"
}
elseif ($SeriesSearchData.data.Count -eq 1)
{
    $Number = "0"
}

Write-Output "Received series data for: `"$($SeriesSearchData.data[$Number].seriesName)`""

##################################
#Get Episode data
##################################

$EpisodeSearchURL = "https://api.thetvdb.com/series/"
$SeriesID = $SeriesSearchData.data[$Number].id
$EpisodeSearchString = "/episodes?page=1"
$EpisodeSearchURLTotal = $EpisodeSearchURL + $SeriesID + $EpisodeSearchString

try
{
    $EpisodeData = Invoke-RestMethod -Method Get -Uri $EpisodeSearchURLTotal -Headers $Headers -ErrorAction Stop
}
catch
{
    Write-Warning "Error searching for episode data"
    $Error[0]
    exit
}

#For Series with more than 100 episodes, combine results from multiple pages
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
            Write-Warning "Error searching for additional episode data pages"
            $Error[0]
            exit
        }

        $EpisodeData.data += $NewEpisodeData.data
    }
}

Write-Output "Received episode data for: `"$($SeriesSearchData.data[$Number].seriesName)`""
Write-Output "Episode count: $($EpisodeData.data.Count)"
"`n"

##################################
#Verify/create destination folder path
##################################

#Determine destination folder/series name
if ($SeriesSearchData.data[$Number].seriesName -ne $FolderName)
{
    Write-Output "TVDB Series name does not match source folder name"
    Write-Output "Select correct series name:"
    Write-Output "0 - $($SeriesSearchData.data[$Number].seriesName)"
    Write-Output "1 - $FolderName"
    Write-Output "2 - Other"
    "`n"
    $SeriesNameSelection = Read-Host "Selection"
    if ($SeriesNameSelection -eq "0")
    {
        $NewSeriesName = $SeriesSearchData.data[$Number].seriesName
    }
    elseif ($SeriesNameSelection -eq "1")
    {
        $NewSeriesName = $FolderName
    }
    elseif ($SeriesNameSelection = "2")
    {
        $NewSeriesName = Read-Host "Enter new series name"
    }
    else
    {
        Write-Warning "Invalid response! Please type in correct series name"
        $NewSeriesName = Read-Host "Last chance moron"
    }
}
else
{
    $NewSeriesName = $SeriesSearchData.data[$Number].seriesName
}

#Match series name to regex for windows file names
if ($NewSeriesName -notmatch $WindowsFileNameRegex)
{
    Write-Warning "Series name does not conform to Windows file name rules"
    $TryAgainSeriesName = Read-Host "Enter in a different series name"
    if ($TryAgainSeriesName -match $WindowsFileNameRegex)
    {
        $DestinationFolderPath = Join-Path -Path $TVDirectory -ChildPath $TryAgainSeriesName
    }
    else
    {
        Write-Warning "NewsSeries name STILL does not conform to Windows file name rules"
        Write-Warning "Clearly you need time to think things over"
        Write-Warning "Exiting..."
        exit
    }
}
else
{
    $DestinationFolderPath = Join-Path -Path $TVDirectory -ChildPath $NewSeriesName
}

#Create destination folder if doesn't exist
if (!(Test-Path $DestinationFolderPath))
{
    Write-Output "Destination folder `"$DestinationFolderPath`" does not exist"
    Write-Output "Creating destination folder `"$DestinationFolderPath`""
    try
    {
        New-Item -Path $TVDirectory -Name $NewSeriesName -ItemType Directory -ErrorAction Stop | Out-Null
        Write-Output "Successfully created destination folder: `"$DestinationFolderPath`""
    }
    catch
    {
        Write-Warning "Failed to create destination folder: `"$DestinationFolderPath`""
        exit
    }
}
else
{
    Write-Output "Destination folder path detected: `"$DestinationFolderPath`""
}

##################################
#Eliminate subfolders
##################################

$Files = Get-ChildItem $TargetFolder.FullName

Write-Output "Checking for subfolders"

foreach ($File in $Files)
{
    #Subfolders: move files to parent folder if folder named Subs\Subtitles\Season*
    if ($File.PSIsContainer -eq $true)
    {
        Write-Output "Subfolder found: $($File.name)"
        if (($File.Name -match "\bSubs\b|\bSubtitles\b") -or ($File.Name -match "\bSeason"))
        {
            Write-Output "Acceptable subfolder detected: `"$($File.name)`""
            $Subs = Get-ChildItem $File.fullname
            foreach ($Subfile in $Subs)
            {
                Write-Output "Moving: $($Subfile.FullName)"
                Write-Output "New path: $($TargetFolder.FullName)"
                Move-Item -LiteralPath $subfile.fullname -Destination $TargetFolder.FullName -Force
            }
            #Delete Subfolder
            Write-Output "Removing folder: $($File.FullName)"
            Remove-Item -LiteralPath $File.FullName -Recurse -Force
        }
        #"Extras" subfolder
        elseif ($File.Name -match "Extras")
        {
            $ExtrasFolder = Get-ChildItem -LiteralPath $File.FullName -Force
            if ((($ExtrasFolder | Measure-Object).Count) -gt 0)
            {
                $ExtrasFolder
                "`n"
                Write-Output "0 - Move folder to destination folder"
                Write-Output "1 - Move child items to current parent folder"
                Write-Output "2 - Delete folder"
                "`n"
                $MoveExtras = Read-Host "Action to perform on `"Extras`" folder"

                if ($MoveExtras -match "[0]")
                {
                    Move-Item -LiteralPath $File.FullName -Destination $DestinationFolderPath -Force
                    Write-Output "Moving `"Extras`" folder to destination folder"
                }
                elseif ($MoveExtras -match "[1]")
                {
                    Write-Output "Moving contents of `"Extras`" folder to parent directory"
                    foreach ($ExtrasItem in $ExtrasFolder)
                    {
                        Move-Item -LiteralPath $ExtrasItem.FullName -Destination $TargetFolder.FullName -Force
                    }
                    #Delete empty folder after moving child items
                    if (((Get-ChildItem $File.FullName | Measure-Object).Count) -eq 0)
                    {
                        Write-Output "`"Extras`" folder is empty, deleting folder"
                        Remove-Item -LiteralPath $File.FullName -Recurse -Force
                    }
                }
                elseif ($MoveExtras -match "[2]")
                {
                    Write-Output "Removing folder: $($File.FullName)"
                    Remove-Item -LiteralPath $File.FullName -Recurse -Force
                }
                else
                {
                    Write-Warning "You are not very smart!"
                    Write-Warning "That was an invalid answer!"
                    Write-Warning "As a reward, you have to start over now!"
                    exit
                }
            }
        }
        else
        {
            Write-Output "Removing folder: $($File.FullName)"
            Remove-Item -LiteralPath $File.FullName -Recurse -Force
        }
    }
}

##################################
#Remove unnecessary files
##################################

$Files = Get-ChildItem -LiteralPath $TargetFolder.FullName

foreach ($File in $Files)
{
    #Delete .TXT .EXE .NFO files
    if (($File.name -like "*.nfo") -or
        ($File.name -like "*.exe") -or
        ($File.name -like "*.txt"))
    {
        Write-Output "Removing file: $($File.name)"
        Remove-Item -LiteralPath $File.fullname -Force
    }
}

##################################
#Rename files
##################################

$Files = Get-ChildItem -LiteralPath $TargetFolder.FullName -Recurse

for ($i=0;$i -lt $Files.Count;$i++)
{
    if ($SeasonNumber) { Remove-Variable SeasonNumber }
    if ($EpisodeNumber) { Remove-Variable EpisodeNumber }
    
    #$Percent = [math]::Round($i/$Files.Count*100,0)
    #Write-Progress -Activity "Renaming file $i of $($Files.count)" -PercentComplete $Percent
    $CurrentFile = $Files[$i]
    
    "`n"
    Write-Output "Parsing file: $($CurrentFile.name)"

    #NewName variable used as working file name
    $NewName = $CurrentFile.Name

    #Season and Episode number are in standard format
    if ($NewName -match $StandardSeasonEpisodeFormatRegex)
    {
        Write-Output "File name `"$NewName`" matches standard formatting"
    }
    #Season and Episode not in standard format
    else
    {
        Write-Warning "File name does NOT contain standard season and episode format"
        #Replace uncommon variations
        $NewName = $NewName.replace(' Chapter ','E')
        $NewName = $NewName.replace('_season_','S')
        $NewName = $NewName.replace(' Season ','S')
        $NewName = $NewName.Replace('_ep_','E')
        $NewName = $NewName.Replace(' Episode ','E')
        $NewName = $NewName.Replace('.Series.','S')
        $NewName = $NewName.Replace(' Series ','S')
        $NewName = $NewName.Replace('Series','S')

        #if file name matches ' 2x01 ' or ' 2e01 ' or [2x01] or [02x01] or [2e01], replace with S2E01
        if ($NewName -match ('(\[| )\d{1,4}(E|e|x|-)\d{1,3}(\]| )'))
        {
            $NewName = $NewName.Replace('[','S').replace(']','')
        }
        #if file name starts with 4 digits for season and episode number
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
    #Parse season/episode numbers from file name if not already done
    if (!($SeasonNumber -or $EpisodeNumber))
    {
        Write-Verbose "Parsing Season\Episode numbers using standard format regex"
        $NewNameSplit = $NewName -split $StandardSeasonEpisodeFormatRegex

        #Parse out season/episode number
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

    #if unable to parse season/episode number, prompt in console session
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

    Write-Output "Detected filename season number: $SeasonNumber"
    Write-Output "Detected filename episode number: $EpisodeNumber"

    ### Trim leading zeroes from season/episode numbers to match TheTVDB data

    #Double digit number with leading zero
    if ($SeasonNumber -match "^0[0-9]$")
    {
        $SeasonTrim = $SeasonNumber.TrimStart('0')
    }
    else
    {
        $SeasonTrim = $SeasonNumber
    }

    #Episode Number has one digit
    if ($EpisodeNumber -match "^\d{1}$")
    {
        $EpisodeTrim = $EpisodeNumber
    }
    #Two digits
    elseif ($EpisodeNumber -match "^\d{2}$")
    {
        #Double digit number with leading zero
        if ($EpisodeNumber -match "^0[0-9]$")
        {
            $EpisodeTrim = $EpisodeNumber.TrimStart('0')
        }
        #Double digit number with non-zero leading number
        else
        {
            $EpisodeTrim = $EpisodeNumber
        }
    }
    #Three digits
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

    # Match episode information to TVDB data
    if ($EpisodeMatch) { Remove-Variable EpisodeMatch }
    $EpisodeMatch = $episodedata.data | Where-Object { ($_.airedepisodenumber -like $episodeTrim) -and ($_.airedseason -like $seasontrim) }
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
    
        $NewFileName = $NewSeriesName + " - " + "S" + $SeasonNumber + "E" + $EpisodeNumber + " - " + $NewEpisodeName + $CurrentFile.extension
        $NewFilePath = Join-Path $DestinationFolderPath -ChildPath $NewFileName
        Write-Output "New file name: `"$NewFileName`""
        Write-Output "Moving `"$($CurrentFile.Name)`" to $NewFilePath"
        Move-Item -LiteralPath $CurrentFile.FullName -Destination $NewFilePath -Force
    }
    else
    {
        Write-Warning "Could not find matching episode data from TVDB"
        Write-Warning "Season query: $SeasonTrim"
        Write-Warning "Episode query: $EpisodeTrim"
        $Confirm = Read-Host "Continue [y/n]?"
        if ($Confirm -match "[yY]")
        {
            Write-Output "Proceeding to next file"
        }
        else { exit }
    }
}