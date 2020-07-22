# Rename-TVFiles
A PowerShell script which renames TV media files using information from TheTVDB API.
## Detailed Description
This PowerShell script takes media files and renames them according to information queried from TheTVDB online database via API query.
## Prerequisities
You must obtain your own API key from TheTVDB website, they are free for personal use. Sign up here [TheTVDB](https://thetvdb.com/).  Your personal API key information must be included in a JSON file referenced in the script parameter 'APIKey'.
## Script Parameters
`-SourcePath`

File path to the source directory where your individual download folders are.  Download folder names must match TV show names according to TheTVDB.
`-APIKey`

File path to a .JSON file containing API key information in order to access TheTVDB online database.
`-VerboseOutput`

Sets VerbosePreference variable to 'Continue', displaying Verbose output in console
## Example Output
```
PS Z:\GitHub\Rename-TVFiles> .\Rename-TVFiles.ps1 -SourcePath \\192.168.0.64\storage\Film\_New\ -APIKey Z:\GitHub\APIKeySample.json

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
