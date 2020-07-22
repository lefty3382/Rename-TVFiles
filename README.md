# Rename-TVFiles
A PowerShell script which renames TV media files using information from TheTVDB API.
## Detailed Description
This PowerShell script takes media files and renames them according to information queried from TheTVDB online database via API query.
## Prerequisities
* You must obtain your own API key from TheTVDB website, they are free for personal use. Sign up here [TheTVDB](https://thetvdb.com/).  Your personal API key information must be included in a JSON file referenced in the script parameter 'APIKey'.
* Files must include season number and episode number in file name (ex. S01E01) in order to properly identify new name.
## Script Parameters
`-SourcePath` (string) - (Example: '__\\\server\storage\Film__')

Folder path string which contains all pertinent files.  The script assumes you have these subfolders:
* \_New - The parent folder which contains all new files to be processed.  Each distinct TV series must be in its own sub-folder.  The name of each sub-folder must match the name of the TV series according to TheTVDB.  This is how the script identifies the correct series to pull episode information for.
* TV - The parent folder which all files will be moved to after processing.  Each distinct TV series will be placed in its own sub-folder using the name of the TV series from TheTVDB database.
* Anime - The parent folder which all designated Anime series will be moved to after processing.  Each distinct TV series will be placed in its own sub-folder using the name of the TV series from TheTVDB database.

`-APIKey` (string)

File path to a .JSON file containing API key information in order to access TheTVDB online database.

`-VerboseOutput` (switch)

Sets VerbosePreference variable to 'Continue', displaying Verbose output in console
## Behavior
The script is interactive in nature, prompting the user to make a series of choices depending on information gathered during execution.  The first choice will be to select the parent destination folder based on media type (TV|Anime).  The script will enumerate all subfolders from source path and prompt user to select which subfolder to process.  The script will then connect to TheTVDB using the API key information provided in the JSON file and attempt to obtain an access token.  If successful, it will then search for the correct TV series from the database based on the name of the subfolder selected.  If multiple values are returned, the script will prompt the user to select the correct one.  The script will then retrieve series and episode information for the chosen TV series.  Extraneous files (.exe | .nfo | .txt) files will be removed.  Each file will be renamed in this format:
* Series Name - S01E01 - Episode Name

Successfully renamed files will be placed in a new subfolder under the media type selected (ex. $SourcePath\TV\SeriesName\EpisodeFileName.ext).
## Example Output
```
PS Z:\GitHub\Rename-TVFiles> .\Rename-TVFiles.ps1 -SourcePath \\192.168.0.64\storage\Film\ -APIKey Z:\GitHub\APIKeySample.json

TV (0) or Anime (1)?: 0


Folder Count: 7    


0 - "Grand Designs"
1 - "It's Garry Shandling's Show"
2 - "Late Night with Seth Meyers"
3 - "Looney Tunes"
4 - "Love & Hip Hop Atlanta"     
5 - "Love & Hip Hop Hollywood"   
6 - "The Daily Show"


Select folder: 6
Processing folder: "The Daily Show"
Successfully retrieved new API token
TVDB search returned 6 results:


0 - "The Daily Show: Global Edition" (83846)
1 - "The Daily Show" (71256)
2 - "The Daily Show Nederlandse Editie" (228101)
3 - "Everyday Engineering Understanding the Marvels of Daily Life" (340305)
4 - "The Daly Show" (275661)
5 - "The Internatinoal" (345864)


Select correct series: 1
Received series data for: "The Daily Show"
Received episode data for: "The Daily Show"
Episode count: 3886


Destination folder path detected: "\\192.168.0.64\storage\Film\TV\The Daily Show"
Checking for subfolders


Parsing file: s25e127.mkv
File name "s25e127.mkv" matches standard formatting
Detected filename season number: 25
Detected filename episode number: 127
Episode data match found
Episode name: Michele Harper & Patton Oswalt
DB Season number: 25
DB Episode number: 127
Updated episode name: "Michele Harper & Patton Oswalt"
New file name: "The Daily Show - S25E127 - Michele Harper & Patton Oswalt.mkv"
Moving "s25e127.mkv" to \\192.168.0.64\storage\Film\TV\The Daily Show\The Daily Show - S25E127 - Michele Harper & Patton Oswalt.mkv
```
