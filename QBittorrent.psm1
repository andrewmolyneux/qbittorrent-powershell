
# Possibly not strictly necessary, but it's nice to have an actual type for
# our QbtSession object.
Add-Type -ReferencedAssemblies ("Microsoft.Powershell.Commands.Utility") -TypeDefinition @"
using System;
using Microsoft.PowerShell.Commands;
public struct QbtSession
{
    public QbtSession(Uri UriIn)
    {
        Uri = UriIn;
        Session = new WebRequestSession();
    }
    public Uri Uri;
    public WebRequestSession Session;
}
"@

Add-Type -TypeDefinition @"
    public enum QbtSort
    {
        Hash,     Name,        Size,      Progress,      Dlspeed,  Upspeed, Priority,
        NumSeeds, NumComplete, NumLeechs, NumIncomplete, Ratio,    Eta,     State,
        SeqDl,    FLPiecePrio, Category,  SuperSeeding,  ForceStart
    }
"@

Add-Type -TypeDefinition @"
    public enum QbtFilter
    {
        All, Downloading, Completed, Paused, Active, Inactive
    }
"@

####################################################################################################
# Utility functions
####################################################################################################

function Join-Uri(
    [Parameter(Mandatory=$true)][Uri] $Uri,
    [Parameter(Mandatory=$true)][String] $Path) {
    New-Object System.Uri ($Uri, $Path)
}

function Get-UTF8JsonContentFromResponse(
    [Parameter(Mandatory=$true)][Microsoft.PowerShell.Commands.HtmlWebResponseObject] $Content) {
    # qBittorrent returns UTF-8 encoded results, but doesn't specify
    # charset=UTF-8 in the Content-Type response header. Technically,
	# because the Content-Type is application/json the default should
	# be UTF-8 anyway. Apparently Invoke-WebRequest doesn't know this.
    $Buffer = New-Object byte[] $Response.RawContentLength
    $Response.RawContentStream.Read($Buffer, 0, $Response.RawContentLength) | Out-Null
    $Decoded = [System.Text.Encoding]::UTF8.GetString($Buffer)
    $Decoded | ConvertFrom-Json
}

function Invoke-FormPost(
    [Parameter(Mandatory=$true)][String] $Uri,
    [Parameter(Mandatory=$true)][Hashtable] $Fields,
    [Microsoft.PowerShell.Commands.WebRequestSession] $WebSession) {
    $Form = New-Object System.Net.Http.MultipartFormDataContent
    foreach ($Field in $Fields.GetEnumerator()) {
      $Form.Add((ToHttpContent $Field.Value), $Field.Name)  
    }
    $Body = $Form.ReadAsStringAsync().GetAwaiter().GetResult()
    $Form.Headers.ContentType.CharSet="UTF-8"
    $ContentType = $Form.Headers.ContentType.ToString()
    Invoke-RestMethod -Method Post -Uri $Uri -ContentType $ContentType -Body $Body -WebSession $WebSession
}

function ConvertTo-TorrentObjects(
    [Parameter(Mandatory=$true)][PSObject] $Array) {
    foreach ($Object in $Array) {
        $NewObject = New-Object PSObject
        $NewObject.PSObject.TypeNames.Insert(0, 'Qbt.Torrent')
        $Object.PSObject.Properties | ForEach-Object {
            $Value = $_.Value
            if ($_.Name -in ("added_on","completion_on","last_activity","seen_complete")) {
                $Value = ConvertFrom-Timestamp $Value
            }
            $NewObject | Add-Member -NotePropertyName (ConvertTo-PowerShellName $_.Name) -NotePropertyValue $Value
        }
        Write-Output $NewObject
    }
}

function ConvertTo-TorrentProperties(
    [Parameter(Mandatory=$true)][PSObject] $Array) {
    foreach ($Object in $Array) {
        $NewObject = New-Object PSObject
        $NewObject.PSObject.TypeNames.Insert(0, 'Qbt.TorrentProperties')
        $Object.PSObject.Properties | ForEach-Object {
            $Value = $_.Value
            if ($_.Name -in ("addition_date","completion_date","creation_date","last_seen")) {
                $Value = ConvertFrom-Timestamp $Value
            }
            $NewObject | Add-Member -NotePropertyName (ConvertTo-PowerShellName $_.Name) -NotePropertyValue $Value
        }
        Write-Output $NewObject
    }
}

function ConvertFrom-Timestamp(
    [Parameter(Mandatory=$true)][Object] $Timestamp) {
    if ($Timestamp -eq -1 -or $Timestamp -eq [Uint32]::MaxValue) {
        [DateTime]::MinValue
    } else {
        (Get-Date "1970-01-01T00:00:00").AddSeconds($Timestamp)
    }
}

function ConvertTo-PowerShellName(
    [Parameter(Mandatory=$true)] [String] $QbtName) {
    # Split on underscores.
    $Words = $QbtName -split "_"
    # If the first character of each word is a letter, make it capital.
    $CapitalisedWords = $Words | ForEach-Object { $_.Substring(0,1).ToUpperInvariant() + $_.Substring(1) }
    # Join it all back together without the underscores.
    $CapitalisedWords -join ""
}

function ConvertTo-QbittorrentName(
    [Parameter(Mandatory=$true)] [String] $PowerShellName) {
    # Remember the first character.
    $FirstCharacter = $PowerShellName[0]
    # Split on capital letters.
    $Chunks = $PowerShellName.Substring(1) -csplit "([A-Z])"
    # Join it all back together with the capital letters prefixed with underscores.
    $Result = $FirstCharacter
    foreach ($Chunk in $Chunks) {
        if ($Chunk.Length -eq 1) {
            $Result += "_$Chunk"
        } else {
            $Result += $Chunk
        }
    }
    $Result.ToLowerInvariant()
}

####################################################################################################
# Exported functions
####################################################################################################

<#
    .Synopsis
    Logs into the qBittorrent WebUI API.

    .Description
    Logs into the qBittorrent WebUI API, specifying the URI of the qBittorrent WebUI and the username and password.

    .Parameter Uri
    The URI of the qBittorrent WebUI (default is http://localhost:8080).

    .Parameter Username
    The username to log into the qBittorrent WebUI (default is "admin").

    .Parameter Password
    The password to log into the qBittorrent WebUI (default is "adminadmin").

    .Outputs
    QbtSession
        A login session object that should be passed via the -Session parameter to other cmdlets.

    .Example
    # Log into the qBittorrent WebUI API on the local machine using the default port and username (you will be prompted for the password).
    $qbt = Open-QbtSession

    .Example
    # Log into the qBittorrent WebUI API with a specific URI, username and password.
    $qbt = Open-QbtSession -Uri http://qbtserver:8000 -Username admin -Password adminadmin
#>
function Open-QbtSession {
    [CmdletBinding()]
    param([Uri] $Uri = "http://localhost:8080",
          [String] $Username = "admin",
          [String] $Password = "adminadmin")

    $S = [QbtSession]::new($Uri)
    $Response = Invoke-WebRequest (Join-Uri $Uri login) -WebSession $S.Session -Method Post -Body @{username=$Username;password=$Password} -Headers @{Referer=$Uri}
    if ($Response.Content -ne "Ok.") {
        throw "Login failed: $($Response.Content)"
    }
    $S
}

<#
    .Synopsis
    Logs out of the qBittorrent WebUI API.

    .Description
    Logs out of the specified qBittorrent WebUI API login session.

    .Parameter Session
    A qBittorrent login session object returned by Open-QbtSession.

    .Outputs
    None.

    .Example
    # Log out of the qBittorrent WebUI session $qbt.
    Close-QbtSession $qbt
#>
function Close-QbtSession(
    [Parameter(Mandatory=$true)][QbtSession] $Session) {
    $Response = Invoke-RestMethod (Join-Uri $Session.Uri logout) -WebSession $Session.Session -Method Post | Out-Null
}

<#
    .Synopsis
    Gets the API version of the qBittorrent application.

    .Description
    Gets the API version of the qBittorrent application. This module has been tested with API version 14 (qBittorrent v3.3.13).

    .Parameter Session
    A qBittorrent login session object returned by Open-QbtSession.

    .Outputs
    Int32
        The qBittorrent API version.

    .Example
    # Get the API version supported by the qBittorrent session $qbt.
    $version = Get-QbtApiVersion $qbt
#>
function Get-QbtApiVersion(
    [Parameter(Mandatory=$true)][QbtSession] $Session) {
    Invoke-RestMethod (Join-Uri $Session.Uri version/api) -WebSession $Session.Session | ConvertFrom-Json
}

<#
    .Synopsis
    Gets the version of the qBittorrent application.

    .Description
    Gets the version of the qBittorrent application. This module has been tested with qBittorrent v3.3.13.

    .Parameter Session
    A qBittorrent login session object returned by Open-QbtSession.

    .Outputs
    String
        The qBittorrent version (e.g. "v3.3.13").

    .Example
    # Get the version of qBittorrent corresponding to the session $qbt.
    $version = Get-QbtVersion $qbt
#>
function Get-QbtVersion(
    [Parameter(Mandatory=$true)][QbtSession] $Session) {
    Invoke-RestMethod (Join-Uri $Session.Uri version/qbittorrent) -WebSession $Session.Session
}

<#
    .Synopsis
    Gets the minimum qBittorrent API version supported.

    .Description
    Gets the minimum API version supported. Any application designed to work with an API version greater than or equal to the minimum API version is guaranteed to work.

    .Parameter Session
    A qBittorrent login session object returned by Open-QbtSession.

    .Outputs
    Int32
        The minimum API version supported.

    .Example
    # Get the minimum API version supported by the session $qbt.
    $version = Get-QbtMinApiVersion $qbt
#>
function Get-QbtMinApiVersion(
    [Parameter(Mandatory=$true)][QbtSession] $Session) {
    Invoke-RestMethod (Join-Uri $Session.Uri version/api_min) -WebSession $Session.Session | ConvertFrom-Json
}

<#
    .Synopsis
    Shuts down the qBittorrent application.

    .Description
    Shuts down the qBittorrent application. This happens immediately and without warning the user. Note that it is not necessary to use Close-QbtSession as shutting down the application will invalidate the session.

    .Parameter Session
    A qBittorrent login session object returned by Open-QbtSession.

    .Outputs
    None.

    .Example
    # Shut down the qBittorrent application corresponding to the session $qbt.
    Stop-QbtApp $qbt
#>
function Stop-QbtApp(
    [Parameter(Mandatory=$true)][QbtSession] $Session) {
    Invoke-RestMethod (Join-Uri $Session.Uri command/shutdown) -WebSession $Session.Session | Out-Null
}

<#
    .Synopsis
    Gets torrents.

    .Description
    Gets torrents matching certain criteria, and optionally sorted.

    .Parameter Session
    A qBittorrent login session object returned by Open-QbtSession.

    .Parameter Filter
    Filter torrent list, e.g. to include only completed torrents.

    .Parameter NoCategory
    Include only torrents without an assigned category.

    .Parameter Category
    Include only torrents with the specified category.

    .Parameter Sort
    Sort torrents by the given key.

    .Parameter ReverseSort
    Enable reverse sorting (e.g. Z-A instead of A-Z, or 9-0 instead of 0-9).

    .Parameter Limit
    Limit the number of torrents returned. Combine with -Offset to allow "paging" when many torrents are present.

    .Parameter Offset
    Set offset. Negative offsets are offsets from the end of the list of torrents.

    .Outputs
    Qbt.Torrent
        This function returns objects that represent torrents.

    .Example
    # Get all torrents.
    $torrents = Get-QbtTorrent $qbt

    .Example
    # Get the first 5 torrents without an assigned category.
    $torrents = Get-QbtTorrent $qbt -NoCategory -Limit 5

    # Example
    # Get the last 5 torrents in the "Linux" category, sorted in ascending name order.
    $torrents = Get-QbtTorrent $qbt -Category Linux -Sort Name -Offset -5

    .Example
    # Get all completed torrents sorted in ascending name order.
    $torrents - Get-QbtTorrent $qbt -Filter Completed -Sort Name
#>
function Get-QbtTorrent(
    [Parameter(Mandatory=$true,Position=1,ParameterSetName="AnyCategory")]
    [Parameter(Mandatory=$true,Position=1,ParameterSetName="NoCategory")]
    [Parameter(Mandatory=$true,Position=1,ParameterSetName="WithCategory")]
    [QbtSession] $Session,
    [QbtFilter] $Filter,
    [Parameter(Mandatory=$true,ParameterSetName="NoCategory")]
    [Switch] $NoCategory,
    [Parameter(Mandatory=$true,ParameterSetName="WithCategory")]
    [String] $Category,
    [QbtSort] $Sort,
    [Switch] $ReverseSort,
    [Int32] $Limit,
    [Int32] $Offset) {
    $Params = @{}
    if ($Filter)      { $Params.Add("filter",   (ConvertTo-QbittorrentName $Filter.ToString())) }
    if ($Sort)        { $Params.Add("sort",     (ConvertTo-QbittorrentName $Sort.ToString())) }
    if ($Limit)       { $Params.Add("limit",    $Limit) }
    if ($Offset)      { $Params.Add("offset",   $Offset) }
    if ($NoCategory)  { $Params.Add("category", "") }
    if ($Category)    { $Params.Add("category", $Category) }
    if ($ReverseSort) { $Params.Add("reverse",  "true") }
    $Response = Invoke-WebRequest (Join-Uri $Session.Uri query/torrents) -WebSession $Session.Session -Body $Params
    ConvertTo-TorrentObjects (Get-UTF8JsonContentFromResponse $Response)
}

<#
    .Synopsis
    Gets generic properties for a torrent.

    .Description
    Gets generic properties for a torrent.

    .Parameter Session
    A qBittorrent login session object returned by Open-QbtSession.

    .Parameter Hash
    The requested torrent's hash (available from Get-QbtTorrent).

    .Outputs
    Qbt.TorrentProperties
        This function returns an object representing a single torrent's generic properties.
#>
function Get-QbtTorrentProperty(
    [Parameter(Mandatory=$true)][QbtSession] $Session,
    [Parameter(Mandatory=$true)][String] $Hash) {
    $Uri = Join-Uri $Session.Uri "query/propertiesGeneral/$Hash"
    $Response = Invoke-WebRequest $Uri -WebSession $Session.Session
    ConvertTo-TorrentProperties (Get-UTF8JsonContentFromResponse $Response)
}

function Get-QbtTorrentWebSeed(
    [Parameter(Mandatory=$true)][QbtSession] $Session,
    [Parameter(Mandatory=$true)][String] $Hash) {
    $Uri = Join-Uri $Session.Uri "query/propertiesWebSeeds/$Hash"
    $Response = Invoke-WebRequest $Uri -WebSession $Session.Session
    Get-UTF8JsonContentFromResponse $Response
}

function Get-QbtPreference(
    [Parameter(Mandatory=$true)][QbtSession] $Session) {
    $Response = Invoke-WebRequest (Join-Uri $Session.Uri query/preferences) -WebSession $Session.Session
    Get-UTF8JsonContentFromResponse $Response
}

function Add-QbtTorrent(
    [Parameter(Mandatory=$true)][QbtSession] $Session,
    [Parameter(Mandatory=$true)][Uri[]] $Uri,
    [String] $SavePath,
    [String] $Cookie,
    [String] $Category,
    [switch] $SkipChecking,
    [switch] $Paused) {
    $Urls = $Uri -join "`n"
    $Fields = @{urls=$Urls}
    if ($SavePath)     { $Fields.Add("save_path",     $SavePath) }
    if ($Cookie)       { $Fields.Add("cookie",        $Cookie)   }
    if ($Category)     { $Fields.Add("category"    ,  $Category) }
    if ($SkipChecking) { $Fields.Add("skip_checking", "true")    }
    if ($Paused)       { $Fields.Add("paused",        "true")    }
    Invoke-FormPost -Uri (Join-Uri $Session.Uri command/download) -Fields $Fields -WebSession $Session.Session | Out-Null
}

function Remove-QbtTorrent(
    [Parameter(Position=0,ParameterSetName="hashes")][string[]] $Hash,
    [Parameter(Position=0,ValueFromPipeline=$true,ParameterSetName="torrents")][PSObject] $Torrent,
    [Parameter(Mandatory=$true)][QbtSession] $Session,
    [switch] $WithData) {
    # $Torrent can be either a single torrent, or an array of torrents
    if ($Torrent) {
        $Hash = ($Torrent | Select -ExpandProperty hash)
    }
    # $Hash can be either a single hash, or an array of hashes
    if ($Hash) {
        $Hashes = $Hash -join "|"
    }
    $CommandPath = "command/delete"
    if ($WithData) {
        $CommandPath = "command/deletePerm"
    }
    Invoke-RestMethod (Join-Uri $Session.Uri $CommandPath) -WebSession $Session.Session -Method Post -Body @{hashes=$Hashes} | Out-Null
}

function Set-QbtPreference(
    [Parameter(Mandatory=$true)][QbtSession] $Session,
    [Hashtable] $Settings) {
    $Json = $Settings | ConvertTo-Json -Compress
    Invoke-RestMethod (Join-Uri $Session.Uri command/setPreferences) -WebSession $Session.Session -Method Post -Body @{json=$Json} | Out-Null
}

# TODO: Export-ModuleMember -Function Get-QbtTorrentTracker
# TODO: Export-ModuleMember -Function Get-QbtTorrentContent
# TODO: Export-ModuleMember -Function Get-QbtTorrentPieceState
# TODO: Export-ModuleMember -Function Get-QbtTorrentPieceHash
# TODO: Export-ModuleMember -Function Get-QbtTransferInfo
# TODO: Export-ModuleMember -Function Get-QbtPartialData
