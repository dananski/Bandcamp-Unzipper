# Bandcamp albums download in zip format with a load of junk in the names and decompressing many zips in a row is annoying to do manually in Windows.
# This determines a cleaner structure more like Windows Media Player's default, with some allowance for customisation along the way, then unzips to that location and renames the track files.
#   "Artist - Album.zip" goes to "Artist\Album\" folder within $myMusicFolder
#   "Artist - Album - Track.*" in the zip goes to simply "Track.*"

# Usage: run with PowerShell using command:
# .\BandcampUnzipper.ps1

#Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope CurrentUser # Run this command once if powershell is restricted (as by default on Windows).

using namespace System.IO
using namespace System.IO.Compression

$myBandcampDownloadsFolder = 'C:\Downloads\Bandcamp\'
$myMusicFolder = 'F:\Music\'

Function SplitAndParse-ZipFileName {
    param (
        [string]$filepath
    )
    $unsplit = [Path]::GetFileNameWithoutExtension($filepath)
    $split = $unsplit.Split(([string[]] (" - ")), 0)
    $doUserEdit = "n"[0]
    if ($split.Length -eq 1) {
        $artist = $album = $split[0]
    }
    elseif ($split.Length -eq 2) {
        $artist = $split[0].Trim()
        $album = $split[1].Trim()
        Write-Host "Artist: $($artist)"
        Write-Host "Album: $($album)"
        $doUserEdit = Read-Host-SingleKey -Prompt "Do you wish to edit before continuing? (y/n; default n)" -allowedChars "yn" -default "n"[0]
    }
    else {
        Write-Host "Could not parse zip name $($unsplit)"
        $doUserEdit = "y"[0]
        $artist = "Unknown Artist"
        $album = "Unknown Album"
    }
    if ($doUserEdit -eq "y"[0]) {
        $artistFromUser = Read-Host -Prompt "Artist [$($artist)]"
        if ($artistFromUser -ne "") {
            $artist = $artistFromUser
        }
        $albumFromUser = Read-Host -Prompt "Album [$($album)]"
        if ($albumFromUser -ne "") {
            $album = $albumFromUser
        }
    }
    return @{ FullPath = $filepath; OrigName = $unsplit; Artist = $artist; Album = $album; NewRelativePath = Join-Path -Path $artist -ChildPath $album }
}

Function Read-Host-SingleKey {
    param (
        [string] $prompt,
        [string] $allowedChars,
        [char] $default
    )
    $keyPressed = "`0"[0] # nul char because unlikely to need to type it
    do {
        Write-Host $prompt
        $keyPressed = $Host.UI.RawUI.ReadKey().Character
        Write-Host
        if ([byte]$keyPressed -eq 13) { # on Enter
            $keyPressed = $default
        }
    }
    while (-not $allowedChars.Contains($keyPressed))
    return $keyPressed
}

Function Unzip-Album {
    param (
        [Hashtable] $nameInfo,
        [string] $outputDestination
    )
    $newLocation = Join-Path -Path $outputDestination -ChildPath $nameInfo.NewRelativePath
    if (Test-Path -Path $newLocation) {
        throw "$($newLocation) already exists"
    }
    Write-Host "Unzipping to $($newLocation)"
    Expand-Archive -Path $nameInfo.FullPath -DestinationPath $newLocation
    return $newLocation
}

Function Rename-Tracks {
    param (
        [string] $folderPath,
        [string] $expectedTrackPrefix
    )
    Write-Host "Renaming tracks containing $($expectedTrackPrefix)"
    $oldAndNewFilenames = Get-ChildItem $folderPath | Where-Object { $_.Name.StartsWith($expectedTrackPrefix) } | ForEach-Object {
        @{
            Old = $_.Name;
            New = $_.Name.Substring($expectedTrackPrefix.Length).Trim(" -");
        }
    }
    Write-Host "New names:"
    $oldAndNewFilenames | ForEach-Object {
        if ($_.Old -eq $_.New) {
            Write-Host "`tUnchanged: $($_.Old)"
        } else {
            Write-Host "`t$($_.Old) --> $($_.New)"
        }
    }
    $continue = Read-Host-SingleKey -prompt 'Continue? (y/n; default y)' -allowedChars "yn" -default "y"[0]
    if ($continue -eq "y"[0]) {
        foreach ($names in $oldAndNewFileNames) {
            Write-Host "Renaming $($names.Old) to $($names.New)"
            Rename-Item -Path (Join-Path -Path $folderPath -ChildPath $names.Old) -NewName $names.New
        }
    }
}

Function Process-NewBandcampDownloads {
    param (
        [string] $downloadsPath,
        [string] $musicFolder
    )
    ForEach ($zipFile in (Get-ChildItem $downloadsPath -Filter '*.zip')) {
        try {
            Write-Host "Processing $($zipFile)"
            $nameInfo = SplitAndParse-ZipFileName $zipFile
            $newLocation = Unzip-Album -nameInfo $nameInfo -outputDestination $musicFolder
            Rename-Tracks -folderPath $newLocation -expectedTrackPrefix $nameInfo.OrigName
        }
        catch {
            Write-Error "Exception occurred processing $($zipFile)"
            Write-Error $_.ToString()
            Write-Host 'Continuing to next zip'
        }
        Write-Host
    }
    Write-Host 'Done.'
}

Process-NewBandcampDownloads -downloadsPath $myBandcampDownloadsFolder -musicFolder $myMusicFolder