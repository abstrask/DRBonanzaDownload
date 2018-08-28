Param (
    [Parameter(Mandatory=$False)]
    [string]$SearchTerm = 'Matador - '
)



Function Get-DRBonanzaVideoWebUrl {

    Param (
        [Parameter(Mandatory=$True)]
        [string]$SearchTerm
    )

    # Init
    Write-Verbose "Searching DR Bonanza for ""$SearchTerm""" -Verbose
    Add-Type -AssemblyName System.Web
    $SearchUrl = "https://www.dr.dk/bonanza/sog?q=""$([System.Web.HttpUtility]::UrlEncode($SearchTerm))"""
    
    # Setup object, navigate to URL
    $ie = New-Object -ComObject "InternetExplorer.Application"
    $ie.Navigate($SearchUrl)
    
    # Wait for page to load - and a bit more to be safe
    while($ie.busy) { start-sleep 1 }
    Start-Sleep 3
    
    # Get properties from document
    $iedoc = $ie.Document
    $SearchResult = @($iedoc.getElementsByTagName("div") | Where-Object {$_.className -like "spot item*"})
    $SearchResult.children | Select-Object -ExpandProperty href
    Write-Verbose "$($SearchResult.Count) results returned" -Verbose

    # Cleanup IE object
    $ie.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) | Out-Null
    Remove-Variable ie

}



Function Get-DRBonanzaVideoProperties {

    Param (
        [Parameter(Mandatory=$True)]
        [string[]]$WebUrl
    )

    #Init
    $Prog = 0
    $Count = $WebUrl.Count
    
    ForEach ($Url in $WebUrl) {

        # Bonanza obfuscates content using client-side JS. Using an Internet Explorer object as workaround.
        Write-Verbose "Getting properties of $Url" -Verbose
        Write-Progress -Activity "Getting video properties" -Status "$Url ($($Prog+1) of $($Count))" -PercentComplete (100 * $Prog / $Count) -Id 1

        # Setup object, navigate to URL
        Try {$ie = New-Object -ComObject "InternetExplorer.Application"}
        Catch {
            Write-Warning "Failed to create Internet Explorer object for $Url"
            Continue
        }
        $ie.Navigate($Url)
        #$ie.visible = $True #debug

        # Wait for page to load - and a bit more to be safe
        while($ie.busy) { start-sleep 1 }
        Start-Sleep 3

        # Get properties from document
        $iedoc = $ie.Document 
        $VideoJsonString = $iedoc.getElementsByTagName("script") | Where-Object {$_.outerHTML -like '*AssetVideo_URL*'} | Select-Object -ExpandProperty text
        $VideoInfo = $iedoc.getElementsByTagName("div") | Where-Object {$_.className -eq "col-sm-8 asset-body"} | Select-Object -ExpandProperty  OuterText
        
        # Cleanup IE object
        $ie.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) | Out-Null
        Remove-Variable ie

        # Parse properties
        $VideoJson = $VideoJsonString.Replace('document.currentAsset = ','') | ConvertFrom-Json
        Add-Type -AssemblyName System.Web
        $Title = [System.Web.HttpUtility]::HtmlDecode($VideoJson.AssetTitle).Trim()
        $ImageUrl = "https:$($VideoJson.AssetImageUrl)"
        $VideoUrl = "https://vod-bonanza.dr.dk/bonanza/mp4:bonanza/bonanza$($VideoJson.AssetVideo_URL)/playlist.m3u8"
        $WebUrl = "https:$($VideoJson.Url)"

        # Parse detailed info
        $Description = $VideoInfo[0]
        $ProgramInfo = $VideoInfo[1]
        $Genre = $VideoInfo[2]
        $Duration = $VideoInfo[3]
        $AirDate = $VideoInfo[4]
        $Cast = $VideoInfo[5]

        # Strip invalid file name chars from title
        $InvalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
        $RegEx = "[{0}]" -f [RegEx]::Escape($InvalidChars)
        $FileName = $Title -replace $RegEx -replace "\.$" # also remove trailing dot

        # Return properties
        [pscustomobject]@{
            BonanzaId = $VideoJson.AssetId
            Title = $Title
            ImageUrl = $ImageUrl
            VideoUrl = $VideoUrl
            WebUrl = $WebUrl
            FileName = $FileName
            Description = $Description
            ProgramInfo = $ProgramInfo
            Genre = $Genre
            Duration = $Duration
            AirDate = $AirDate
            Cast = $Cast
        }

        $Prog++

    }

    Write-Progress -Activity "Getting video properties" -Completed

}




Function New-DRBonanzaVideoTextFile {

    Param (

        [Parameter(
            Mandatory=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [string]$FileName,

        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true
        )]
        $BonanzaId,
        
        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true
        )]
        $Title,

        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true
        )]
        $ImageUrl,

        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true
        )]
        $VideoUrl,

        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true
        )]
        $WebUrl,

        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true
        )]
        $Description,

        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true
        )]
        $ProgramInfo,

        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true
        )]
        $Genre,

        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true
        )]
        $Duration,

        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true
        )]
        $AirDate,

        [Parameter(
            Mandatory=$false,
            ValueFromPipelineByPropertyName=$true
        )]
        $Cast

    )

$String = @"
Titel:
$Title

Beskrivelse:
$Description

Programinfo:
$ProgramInfo

Genre:
$Genre

Tid:
$Duration

Sendt:
$AirDate

Medvirkende:
$Cast

"@

    $String | Out-File -FilePath "$FileName.txt"

}



# Define variables
$YoutubeDL = "C:\Users\Rasmus\Dropbox\Scripts\DRBonanzaDownload\youtube-dl.exe"
$FFMpeg = "C:\Users\Rasmus\Dropbox\Scripts\DRBonanzaDownload\ffmpeg.exe"

# Get webpage URLs by search
$Url = Get-DRBonanzaVideoWebUrl -SearchTerm $SearchTerm

# Get video properties using webpage URLs
$Video = Get-DRBonanzaVideoProperties -WebUrl $Url

# Process each video
$Prog = 0
$Count = $Video.Count

$Video | ForEach {

    Write-Progress -Activity "Downloading video" -Status "$($_.FileName) ($($Prog+1) of $($Count))" -PercentComplete (100 * $Prog / $Count) -Id 1

    # Generate text file containing various program information
    $_ | New-DRBonanzaVideoTextFile

    # Download video
    & $YoutubeDL --no-warnings --output """$($_.FileName)$('.%(ext)s')""" --add-metadata --console-title $_.VideoUrl --ffmpeg-location ""$FFMpeg""

    $Prog++

}
Write-Progress -Activity "Downloading video" -Completed