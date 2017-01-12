Add-Type -AssemblyName System.Web.Extensions

function json
{
    <#
    .synopsis
    Decode JSON
    .example
    PS> echo "{'x':123,'y':22}" | json | % y
    22
    #>

    param(
        [Parameter(ValueFromPipeline = $true)][string]$i)

    begin
    {
        $jsser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    }
    process
    {
        $jsser.MaxJsonLength = $i.length + 100 # Make limit big enough
        $jsser.RecursionLimit = 100
        $jsser.DeserializeObject($i)
    }
}

function unjson
{
    <#
    .synopsis
    Encode JSON
    .parameter pretty
    Print prety JSON
    .example
    PS> echo "{'x':123,'y':22}" | json | unjson
    {"x":123,"y":22}
    #>

    param(
        [switch]$pretty,
        [Parameter(ValueFromPipeline = $true)]$obj)

    begin
    {
        $jsser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    }
    process
    {
        $r = $jsser.Serialize($obj)
        if ($pretty)
        {
            $r | jq .
        }
        else
        {
            $r
        }
    }
}

function whereis
{
    <#
    .synopsis
    Get location of file in PATH
    .parameter name
    Name of executable, omit if you want to get all names in PATH
    .example
    PS> whereis ping | % FullName
    C:\Windows\System32\PING.EXE
    #>

    param(
        [string]$name)

    $env:path -split ';' | % {
        if ($_ -and (test-path $_)) {
            ls -path $_ -filter "$($name).*"
        }
    }
}

set-alias which whereis

# download page
function download
{
    <#
    .synopsis
    Download single file
    .parameter url
    URL of file to download
    .parameter path
    Download destination. Default is current path
    .example
    download "www.foo.org/1.jpg"
    .example
    download "www.foo.org/1.jpg" subdir
    #>

    param(
        [Parameter(ValueFromPipeline = $true)]
        [string]$url,
        [string]$path)

    begin {
        if (!$path) {
            $path = (gl).Path
        }
        else {
            $path = gi $path
        }
    }
    process {
        if (test-path -pathType Container $path) {
            $fname = wget -method head $url | % { [uri]::UnescapeDataString($_.BaseResponse.ResponseUri.Segments[-1]) }
            $dst = join-path $path $fname
        }
        else {
            $dst = $path
        }

        echo "Downloading $url to $dst"

        $client = New-Object System.Net.WebClient
        $client.DownloadFile($url, $dst)        
    }
}

function wreq
{
    <#
    .synopsis
    Read web page
    .parameter url
    URL of page
    .parameter encoding
    Encoding
    .example
    wreq www.yandex.ru utf-8
    #>

    param(
        [Parameter(ValueFromPipeline = $true)]
        [string]$url,
        [string]$encoding = 'utf-8')

    if (!($url -match '^http(s?)://')) {
        $url = 'http://' + $url
    }

    $wr = new-object system.net.webclient
    $wr.encoding = [system.text.encoding]::getencoding($encoding)
    $wr.downloadstring([uri]$url)
}

$htmlagilitypack = whereis HtmlAgilityPack.dll

if ($htmlagilitypack) {
    add-type -path $htmlagilitypack.FullName
}

function html
{
    <#
    .synopsis
    Parse HTML
    #>

    param(
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [string]$contents)

    $doc = new-object HtmlAgilityPack.HtmlDocument
    $doc.LoadHtml($contents)
    $doc
}

function xpath
{
    <#
    .synopsis
    Select nodes from HTML with XPath
    .parameter document
    HTML document
    .parameter node
    HTML node
    .parameter path
    XPath expression
    #>

    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$path,
        [Parameter(ValueFromPipeline = $true, Mandatory = $true, ParameterSetName = "string")]
        [string]$contents,
        [Parameter(ValueFromPipeline = $true, Mandatory = $true, ParameterSetName = "document")]
        [HtmlAgilityPack.HtmlDocument]$document,
        [Parameter(ValueFromPipeline = $true, Mandatory = $true, ParameterSetName = "node")]
        [HtmlAgilityPack.HtmlNode]$node)

    process {
        $n = $null

        switch ($pscmdlet.parametersetname) {
            "string" {
                $doc = new-object HtmlAgilityPack.HtmlDocument
                $doc.LoadHtml($contents)
                $n = $doc.DocumentNode
            }
            "document" {
                $n = $document.DocumentNode
            }
            "node" {
                $n = $node
            }
        }

        $n.SelectNodes($path)
    }
}

function time([scriptblock]$s)
{
    <#
    .synopsis
    Measure time of script block
    .example
    time { foo; bar; baz; }
    #>
    $tm = get-date
    & $s
    (get-date) - $tm
}

function environment
{
    <#
    .synopsis
    Get/set environment as hash
    .parameter target
    Permanent store name: user, machine or process
    .parameter vars
    Environment to restore
    .parameter script
    Script to run
    #>

    [CmdLetBinding(DefaultParameterSetName = "get-set")]
    param(
        [Parameter(ParameterSetName = "get-set")]
        [ValidateSet("User", "Machine", "Process")]
        [string]$target,
        [Parameter(ValueFromPipeline = $true, ParameterSetName = "get-set")]
        [hashtable]$vars,
        [Parameter(Mandatory = $true, ParameterSetName = "script", Position = 0)]
        [scriptblock]$script)

    process
    {
        write-verbose "Parameter set: $($pscmdlet.parametersetname)"
        switch ($pscmdlet.parametersetname)
        {
            "get-set" {
                write-verbose "In get-set"
                if (!$vars)
                {
                    if ($target)
                    {
                        write-verbose "Getting from $($target)"
                        [Environment]::GetEnvironmentVariables([EnvironmentVariableTarget]::$target)
                    }
                    else
                    {
                        write-verbose "Get current"
                        ls env: | hash Key Value                    
                    }
                }
                else
                {
                    $vars | enumerate | % {
                        if ($target)
                        {
                            write-verbose "Setting in $($target): $($_.Key) = $($_.Value)"
                            [Environment]::SetEnvironmentVariable($_.Key, $_.Value, [EnvironmentVariableTarget]::$target)
                        }
                        else
                        {
                            write-verbose "Setting: $($_.Key) = $($_.Value)"
                            sc "env:\$($_.Key)" $_.Value
                        }
                    }
                }
                break
            }
            "script" {
                write-verbose "Saving environment variables"
                $e = ls env: | hash Key Value
                try
                {
                    write-verbose "Running script"
                    & $script                    
                }
                finally
                {
                    write-verbose "Restoring environment variables"
                    rm env:\*
                    $e | enumerate | % { sc "env:\$($_.Key)" $_.Value }
                }
            }
        }
    }
}

function invoke-cmd([string]$f)
{
    <#
    .synopsis
    Invoke .bat file
    .description
    Invokes .bat files and saves environment changes
    #>
    cmd /c "$f && set" | % {
        if ($_ -match "^(.*?)=(.*)$") {
            sc "env:\$($matches[1])" $matches[2]
        }
    }
}

function unzip
{
    <#
    .synopsis
    Extract zip archive
    .parameter archive
    Archive file
    .parameter Output
    Output directory
    .example
    PS> unzip 1.zip out
    #>
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$archive,

        [string]$Output = ".",

        [Parameter(ValueFromRemainingArguments = $true)]
        $files)
    7z x $archive "-o$($Output)" -y $files | ? { $_ -match "^Extracting\s+(.*)$" } | % { $matches[1] }
}

function zip([string]$archive)
{
    <#
    .synopsis
    Make zip archive
    .description
    Passes all parameters to 7z
    #>
    7z a $archive $args -y
}

function lszip([string]$archive, [string]$Mask)
{
    <#
    .synopsis
    List zip archive contents
    .description
    List files in archive like 'ls'
    .parameter archive
    Archive file
    .parameter Mask
    Mask to filter files like in 'ls'
    .example
    PS> lszip 1.zip *.txt
    #>
    $result = 7z l $archive $Mask
    $start = $result | ? { $_ -match "^(((-+)\s+)+)(-+)" } | select -first 1 | % {
        $matches[1].length
        $line = $matches[4]
    }

    $result = $result | ? { $_.length -ge $start } | % { $_.substring($start) }
    $result = $result | select -skip ($result.IndexOf($line) + 1)
    $result = $result | select -first ($result.IndexOf($line))

    if ($Mask)
    {
        $result | ? { $_ -like "$($Mask)" }
    }
    else
    {
        $result
    }
}

function codepage
{
    <#
    .synopsis
    Get or set codepage
    .parameter cp
    Codepage to set
    .parameter script
    Script to run with codepage. If set then codepage will be restored back after running script
    .example
    PS> codepage
    866
    PS> codepage 65001
    Active code page: 65001
    PS> codepage 65001 { cabal list }
    #>

    param(
        [Parameter(Position = 0)]
        [int]$cp,        
        [Parameter(Position = 1)]
        [scriptblock]$script)

    if (!$cp)
    {
        if ($(chcp) -match "\d+$") { $matches[0] }
    }
    elseif (!$script)
    {
        chcp $cp
    }
    else
    {
        $old = codepage
        try
        {
            codepage $cp | out-null
            & $script        
        }
        finally
        {
            codepage $old | out-null
        }
    }
}

function encoding
{
    <#
    .synopsis
    Get or set output encoding
    .parameter encoding
    Encoding to set
    .parameter script
    Script to run with encoding. If set then encoding will be restored back after running script
    #>

    param(
        [Parameter(Position = 0)]
        [string]$encoding,
        [Parameter(Position = 1)]
        [scriptblock]$script)

    if (!$encoding)
    {
        Get-Variable -Name OutputEncoding -Scope Global -ValueOnly
    }
    elseif (!$script)
    {
        Set-Variable -Name OutputEncoding -Value ([System.Text.Encoding]::GetEncoding($encoding)) -Scope Global
    }
    else
    {
        $old = encoding
        try
        {
            encoding $encoding
            & $script
        }
        finally
        {
            Set-Variable -Name OutputEncoding -Value $old -Scope Global
        }
    }
}

function verbose
{
    <#
    .synopsis Run script with verbose enabled
    .parameter script
    Script to run
    #>

    param(
        [scriptblock]$script)

    [System.Management.Automation.ActionPreference]$old = Get-Variable -Name VerbosePreference -Scope Global -ValueOnly
    try
    {
        Set-Variable -Name VerbosePreference -Value ([System.Management.Automation.ActionPreference]::Continue) -Scope Global
        & $script
    }
    finally
    {
        Set-Variable  -Name VerbosePreference -Value $old -Scope Global    
    }
}

function path
{
    <#
    .synopsis
    Get path by name
    .parameter name
    Name of path (MyDocuments,MyPictures etc.)
    .example
    PS> path MyDocuments
    d:\users\voidex\Documents
    #>

    param(
        [ValidateSet("User", "Desktop", "Programs", "MyDocuments", "Personal", "Favorites", "Startup", "Recent", "SendTo", "StartMenu", "MyMusic", "MyVideos", "DesktopDirectory", "MyComputer", "NetworkShortcuts", "Fonts", "Templates", "CommonStartMenu", "CommonPrograms", "CommonStartup", "CommonDesktopDirectory", "ApplicationData", "PrinterShortcuts", "LocalApplicationData", "InternetCache", "Cookies", "History", "CommonApplicationData", "Windows", "System", "ProgramFiles", "MyPictures", "UserProfile", "SystemX86", "ProgramFilesX86", "CommonProgramFiles", "CommonProgramFilesX86", "CommonTemplates", "CommonDocuments", "CommonAdminTools", "AdminTools", "CommonMusic", "CommonPictures", "CommonVideos", "Resources", "LocalizedResources", "CommonOemLinks", "CDBurning")]
        [string]$name)

    [Environment]::GetFolderPath($name)
}

function import-profile
{
    <#
    .synopsis
    Import profile scripts
    #>
    @(
        $Profile.AllUsersAllHosts,
        $Profile.AllUsersCurrentHost,
        $Profile.CurrentUserAllHosts,
        $Profile.CurrentUserCurrentHost        
    ) | % {
        if (test-path $_) {
            write-verbose "Running Import-Module $_ -Global -Force"
            import-module $_ -Global -Force
        }
    }
}

function rhistory
{
    <#
    .synopsis
    Search in history
    .parameter h
    Regex to match
    .example
    PS> rhistory path

        Id CommandLine                                                                                                                                                                                                                                                                                                               
        -- -----------                                                                                                                                                                                                                                                                                                               
        67 path MyDocuments                                                                                                                                                                                                                                                                                                          
        68 path Documents                                                                                                                                                                                                                                                                                                            
        69 path MyPictures                                                                                                                                                                                                                                                                                                           
        70 path Documents | clip                                                                                                                                                                                                                                                                                                     
        71 path MyDocuments | clip                                                                                                                                                                                                                                                                                                   
        72 rhistory path                                                                                                                                                                                                                                                                                                             

    #>
    param([Parameter(Mandatory=$true)][string]$h)
    history | ? { $_.CommandLine -match $h }
}

Add-Type -Assembly PresentationCore

function clear-clipboard
{
    <#
    .synopsis
    Clear clipboard
    #>
    [Windows.Clipboard]::Clear()
}

new-alias clear-clip clear-clipboard -scope global -erroraction silentlycontinue

function set-clipboard
{
    <#
    .synopsis
    Set clipboard text
    .example
    PS> set-clipboard 123
    PS> get-clipboard
    123
    #>
    param(
        [Parameter(ValueFromPipeline = $true)]
        [string]
        $inputText)

    [Windows.Clipboard]::SetText($inputText)
}

function out-clipboard
{
    <#
    .synopsis
    Output to clipboard
    .parameter Separator
    Separate lines with Separator, default is newline
    .parameter Echo
    Echo clipboarded data
    .example
    PS> 1, 2, 3, 4 | out-clipboard
    PS> get-clipboard
    1
    2
    3
    4
    PS> 1, 2, 3, 4 | out-clipboard -Separator ','
    PS> get-clipboard
    1,2,3,4
    PS> 1, 2, 3, 4 | out-clipboard -Separator ',' -Echo
    1,2,3,4
    #>
    param(
        [string]
        $Separator = "`r`n",
        [Parameter(ValueFromPipeline = $true)]
        [string]
        $InputObject,
        [switch]
        $Echo)

    begin
    {
        [string[]]$clipboard = @()
        clear-clipboard
    }
    process
    {
        $clipboard += $InputObject
    }
    end
    {
        $result = $clipboard -join $Separator
        [Windows.Clipboard]::SetText($result)
        if ($Echo)
        {
            $result
        }
    }
}

new-alias out-clip out-clipboard -scope global -erroraction silentlycontinue

function get-clipboard
{
    <#
    .synopsis
    Get clipboard contents
    .parameter Separator
    Separator to split clipboard with
    .example
    PS> 1, 2, 3, 4 | out-clipboard
    PS> get-clipboard | measure -Sum | select -exp Sum
    10
    PS> 1, 2, 3, 4 | out-clipboard -Separator ','
    PS> get-clipboard -Separator ','
    1
    2
    3
    4
    #>

    param(
        [string]
        $Separator = "`r`n")
    [Windows.Clipboard]::GetText() -split $Separator
}

new-alias get-clip get-clipboard -scope global -erroraction silentlycontinue

# taglib
$TagLib = $PSScriptRoot + "\taglib-sharp.dll"

[System.Reflection.Assembly]::LoadFile($TagLib) | out-null

# load file for tags
function tags
{
    <#
    .synopsis
    Get/set tags for audio file
    .description
    Get (all or some of) or set tags for audio file.
    .parameter path
    File to get/set tags
    .parameter get
    List of tags to get
    .parameter set
    Hashtable to set
    .example
    PS> tags foo.mp3 -get title, author
    @{Title='foo';'Author'='bar'}
    PS> tags foo.mp3 -set @{title='new title'}
    .example
    PS> ls *.mp3 | % { if ($_.Name -match '\d+\.\s(.*)\.mp3') { tags $_ -set @{Title=$matches[1]} } }
    #>

    [CmdLetBinding(DefaultParameterSetName = "get")]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [string]$path,
        [Parameter(ParameterSetName = "get")]
        [string[]]$get,
        [Parameter(ParameterSetName = "set")]
        [hashtable]$set)

    process {
        $t = [TagLib.File]::Create((gi $path))
        switch ($pscmdlet.parametersetname)
        {
            "get" {
                if (!$get) {
                    $get = $t.Tag | gm -membertype property | % name
                }
                new-object psobject -property ($get | % { @{Name=$_;Value=$t.Tag.$_} } | hash)
            }
            "set" {
                $set | enumerate | % {
                    $t.Tag.($_.Key) = $_.Value
                }
                $t.Save()
            }
        }
    }
}

function get-handle
{
    <#
    .synopsis
    Get opened handles with SysInternals handle util
    .description
    Returns opened handles
    .parameter Process
    Handle owning process (partial name)
    .parameter Name
    Handle partial name
    .parameter File
    Returns File handles
    .parameter Section
    Returns Section handles
    .example
    PS> get-handle -n txt -File | % Handle
    C:\ProgramData\Microsoft\Windows Defender\Network Inspection System\Support\NisLog.txt
    C:\Program Files (x86)\Steam\logs\bootstrap_log.txt
    C:\Program Files (x86)\Steam\logs\content_log.txt
    C:\Program Files (x86)\Steam\logs\remote_connections.txt
    C:\Program Files (x86)\Steam\logs\connection_log.txt
    C:\Program Files (x86)\Steam\logs\cloud_log.txt
    C:\Program Files (x86)\Steam\logs\parental_log.txt
    C:\Program Files (x86)\Steam\logs\appinfo_log.txt
    C:\Program Files (x86)\Steam\logs\stats_log.txt
    #>

    param(
        [string]
        $Process,
        [string]
        $Name,
        [switch]
        $File,
        [switch]
        $Section)

    if (!$File -and !$Section)
    {
        $File = $true
        $Section = $true
    }

    function psexe([string]$process_name)
    {
        if ($process_name -match '^(.*)\.exe$') { ps $matches[1] -erroraction silentlycontinue } else { $null }
    }
    function validate_type([ValidateSet('File', 'Section')][string]$type_name)
    {
        ($File -and ($type_name -eq 'File')) -or ($Section -and ($type_name -eq 'Section'))
    }

    $as = @()

    if ($Process)
    {
        $as = $as + @("-p", $Process)
    }
    if ($Name)
    {
        $as = $as + @($Name)
    }

    if ($Name)
    {
        &handle $as |
            ? { $_ -match "^(?<Process>[^\s]+)\s+pid:\s(?<PID>\d+)\s+type:\s(?<Type>\w+)\s+(?<ID>[\w\d]+):\s(?<Handle>.*)$" } |
            % { $matches } |
            ? { validate_type($_.Type) } |
            % {
                New-Object PSObject -Property @{
                    'Process'=psexe $_.Process;
                    'PID'=[int]$_.PID;
                    'Type'=$_.Type;
                    'ID'=$_.ID;
                    'Handle'=$_.Handle;
                }
            }
    }
    else
    {
        (&handle $as |
            % {
                if ($_ -match "^\-+$")
                {
                    $r = $result
                    $result = New-Object PSObject -Property @{
                        'Process'=$null;
                        'PID'=$null;
                        'User'=$null;
                        'Handles'=@()
                    }
                    $r
                }
                if ($_ -match "^(?<Process>[^\s]+)\spid:\s(?<PID>\d+)\s(?<User>.*)$")
                {
                    $matches | % {
                        $result.Process = psexe $_.Process
                        $result.PID = $_.PID
                        $result.User = $_.User
                    }
                }
                if ($_ -match "^\s*(?<ID>[\w\d]+):\s(?<Type>\w+)\s+(\((?<Read>[R\-])(?<Write>[W\-])(?<Delete>[D\-])\)\s+)?(?<Handle>.*)$")
                {
                    $matches | ? { validate_type($_.Type) } | % {
                        $h = New-Object PSObject -Property @{
                            'ID'=$_.ID;
                            'Type'=$_.Type;
                            'Access' = New-Object PSObject -Property @{
                                'Read'=$_.Read -eq 'R';
                                'Write'=$_.Write -eq 'W';
                                'Delete'=$_.Delete -eq 'D';
                            };
                            'Handle'=$_.Handle;
                        }
                        $result.Handles = $result.Handles + @($h)
                    }
                }
            }), $result
    }
}

function dictionary
{
    <#
    .synopsis
    Get/set value from dictionary
    .description
    Get/set value to key-value dictionary
    .parameter key
    Key to get/set value
    .parameter value
    If specified, sets this value
    .parameter delete
    Delete specified key
    .parameter pretty
    Print pretty JSON
    .example
    PS> gc dict.json | dictionary "foo"
    123
    PS> gc dict.json | dictionary "foo" 135 | out-file dict.json
    #>

    param(
        [Parameter(Mandatory=$true)][string]
        $key,
        $value,
        [switch]
        $delete,
        [switch]
        $pretty,
        [Parameter(ValueFromPipeline = $true)]
        $inputText)

    if (!$inputText)
    {
        $inputText = "{}"
    }

    if ($delete)
    {
        json $inputText | % { $_.Remove($key); $_ } | unjson -pretty:$pretty
    }
    elseif ($value)
    {
        json $inputText | % { $_[$key] = $value; $_ } | unjson -pretty:$pretty
    }
    else
    {
        json $inputText | % { $_[$key] }
    }
}

new-alias dict dictionary -scope global -erroraction silentlycontinue

function gdict
{
    <#
    .synopsis
    Get dictionary contents
    .parameter path
    File of dictionary
    .example
    PS> gdict "1.json"
    foo = 123
    bar = 22
    #>

    param(
        [Parameter(Mandatory = $true)][string]
        $path)

    -join (gc (ls $path)) | json | % { $d = $_; $d.Keys | % { $_ + " = " + $d[$_] } }
}

function select-match
{
    <#
    .synopsis
    Finds strings that matches regex, like `select-string` but also colorize output
    .parameter str
    String to find regex in
    .parameter pattern
    Regex pattern
    .parameter onlymatch
    Return (output) only matched value, implies nocolor
    .parameter caseinsensitive
    Case insensitive match
    .parameter group
    Match specific group, can be group index or name if there are named groups
    .parameter nocolor
    Don't color output and return it as string
    .example
    PS> 'foo 123 321' | select-match '\d+' -o
    123
    321
    PS> 'foo', 'some x=1', 'bar y=3' | select-match '[a-z]+=(?<val>\d+)'
    some x=1
    bar y=3
    PS> 'foo', 'some x=1', 'bar y=3' | select-match '[a-z]+=(?<val>\d+)' -o
    x=1
    y=3
    PS> 'foo', 'some x=1', 'bar y=3' | select-match '[a-z]+=(?<val>\d+)' -o -group val
    1
    3
    #>

    param(
        [Parameter(ValueFromPipeline = $true)]
        [string]$str,
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$pattern,
        [switch]$onlymatch,
        [switch]$caseinsensitive,
        $group = 0,
        [switch]$nocolor)

    process
    {
        $r = $str | select-string $pattern -allmatches -casesensitive:(!$caseinsensitive)
        if ($r) {
            if ($onlymatch) {
                $r.matches | % { $_.groups[$group] } | % {
                    $str.substring($_.index, $_.length)
                }
            }
            else {
                if ($nocolor) {
                    $str
                }
                else {
                    $index = 0
                    $r.matches | % { $_.groups[$group] } | % {
                        write-host $str.substring($index, $_.index - $index) -nonewline
                        write-host $str.substring($_.index, $_.length) -f red -nonewline
                        $index = $_.index + $_.length
                    }
                    write-host $str.substring($index)
                }
            }
        }
    }
}

function colorize
{
    <#
    .synopsis
    Colorize output. Pass regex for some color and it will colorize matched parts
    It accepts dictionary in format @{Color=Regex} or also regex
    for each color as separate flag
    .example
    colorize 'some string 123' @{Red="\\d+"}
    'some string 123' | colorize @{Red="\\d+"}
    'some string 123' | colorize -Red \\d+
    #>

    param(
        [Parameter(ValueFromPipeline = $true)]
        [string]$str,
        [hashtable]$colors=@{},
        [string]$Black,
        [string]$DarkBlue,
        [string]$DarkGreen,
        [string]$DarkCyan,
        [string]$DarkRed,
        [string]$DarkMagenta,
        [string]$DarkYellow,
        [string]$Gray,
        [string]$DarkGray,
        [string]$Blue,
        [string]$Green,
        [string]$Cyan,
        [string]$Red,
        [string]$Magenta,
        [string]$Yellow,
        [string]$White,
        [switch]$caseinsensitive)

    begin {
        if ($Black) { $colors.Black = $Black }
        if ($DarkBlue) { $colors.DarkBlue = $DarkBlue }
        if ($DarkGreen) { $colors.DarkGreen = $DarkGreen }
        if ($DarkCyan) { $colors.DarkCyan = $DarkCyan }
        if ($DarkRed) { $colors.DarkRed = $DarkRed }
        if ($DarkMagenta) { $colors.DarkMagenta = $DarkMagenta }
        if ($DarkYellow) { $colors.DarkYellow = $DarkYellow }
        if ($Gray) { $colors.Gray = $Gray }
        if ($DarkGray) { $colors.DarkGray = $DarkGray }
        if ($Blue) { $colors.Blue = $Blue }
        if ($Green) { $colors.Green = $Green }
        if ($Cyan) { $colors.Cyan = $Cyan }
        if ($Red) { $colors.Red = $Red }
        if ($Magenta) { $colors.Magenta = $Magenta }
        if ($Yellow) { $colors.Yellow = $Yellow }
        if ($White) { $colors.White = $White }
    }
    process {
        $ms = $null
        $colors | enumerate | % {
            $color = $_.Name
            $r = $str | select-string $_.Value -allmatches -casesensitive:(!$caseinsensitive)
            $ms += $r.matches | % { @{color=$color;match=$_} }
        }
        $index = 0
        $ms | sort -property { $_.match.index } | % {
            write-host $str.substring($index, $_.match.index - $index) -nonewline
            write-host $str.substring($_.match.index, $_.match.length) -f $_.color -nonewline
            $index = $_.match.index + $_.match.length
        }
        write-host $str.substring($index)
    }
}

function wait
{
    <#
    .synopsis
    Wait for process
    .parameter Process
    Process object
    .parameter Timeout
    Timeout to wait
    .parameter Kill
    To kill after timeout
    #>

    param(
        [Parameter(ValueFromPipeline = $true)]
        [System.Diagnostics.Process]$Process,
        [int]$Timeout = 0,
        [switch]$Kill)

    process
    {
        if ($Process)
        {
            $good = if (!$Timeout) { $Process.WaitForExit() } else { $Process.WaitForExit($Timeout) }
            if (!$good)
            {
                if ($Kill)
                {
                    $Process.Kill()
                }
            }
        }
    }
}

[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$src = @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PInvoke
{
    public static class NativeMethods
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("User32.dll")]
        static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("User32.dll")]
        static extern int ReleaseDC(IntPtr hwnd, IntPtr dc);

        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        public static void GetScreenSize(ref RECT lpRect)
        {
            IntPtr hdc = GetDC(IntPtr.Zero);
            int VERTRES = 117;
            int HORZRES = 118;
            int actualPixelsX = GetDeviceCaps(hdc, HORZRES);
            int actualPixelsY = GetDeviceCaps(hdc, VERTRES);
            ReleaseDC(IntPtr.Zero, hdc);
            lpRect.Left = 0;
            lpRect.Top = 0;
            lpRect.Right = actualPixelsX;
            lpRect.Bottom = actualPixelsY;
        }

        public static void ScaleToPixels(ref RECT lpRect)
        {
            int w = SystemInformation.VirtualScreen.Size.Width;
            int h = SystemInformation.VirtualScreen.Size.Height;
            RECT screen = new RECT();
            GetScreenSize(ref screen);
            lpRect.Left = lpRect.Left * screen.Right / w;
            lpRect.Top = lpRect.Top * screen.Bottom / h;
            lpRect.Right = lpRect.Right * screen.Right / w;
            lpRect.Bottom = lpRect.Bottom * screen.Bottom / h;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;        // x position of upper-left corner
        public int Top;         // y position of upper-left corner
        public int Right;       // x position of lower-right corner
        public int Bottom;      // y position of lower-right corner
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }
}
'@

Add-Type -TypeDefinition $src -ReferencedAssemblies 'System.Drawing', 'System.Windows.Forms'
Add-Type -AssemblyName System.Drawing

function save-image
{
    <#
    .synopsis
    Save image to file
    .parameter Out
    Output directory or file name, default is current directory
    #>

    param(
        [string]$Out,
        [Parameter(ValueFromPipeline = $true)]
        [Drawing.Bitmap]$Image)

    process
    {
        $res = $null
        $dir = $null
        $filter = '(?<N>\d+)\.png'
        $name = $null

        if ($Image.Tag) {
            if ($Image.Tag -match 'screenshot:(?<Name>.*)') {
                $name = $matches.Name
                $dir = join-path (path MyPictures) Screenshots
                $filter = "$name \((?<N>\d+)\)\.png"
            }
            elseif ($Image.Tag -eq 'screenshot') {
                $dir = join-path (path MyPictures) Screenshots
            }
        }

        if (!$Out)
        {
            if (!$dir) {
                $dir = gl
            }
            $num = ls $dir -Filter '*.png' | select-match $filter -Group N | measure -Maximum | ? { $_.Maximum -ne $null } | % { $_.Maximum + 1 }
            if (!$num) { $num = 0 }
            if ($name) {
                $res = join-path $dir "$name ($num).png"
            }
            else {
                $res = join-path $dir "$num.png"
            }
        }
        else
        {
            if (test-path $Out -PathType Container)
            {
                $dir = resolve-path $Out
                $num = ls $dir -Filter '*.png' | select-match $filter -Group N | measure -Maximum | ? { $_.Maximum -ne $null } | % { $_.Maximum + 1 }
                if (!$num) { $num = 0 }
                if ($name) {
                    $res = join-path $dir "$name ($num).png"
                }
                else {
                    $res = join-path $dir "$num.png"
                }
            }
            elseif (!(split-path $Out))
            {
                $dir = gl
                $res = join-path $dir $Out
            }
            elseif (test-path (split-path $Out) -PathType Container)
            {
                $dir = resolve-path (split-path $Out)
                $name = split-path $Out -Leaf
                $res = join-path $dir $name
            }
            else
            {
                throw "Invalid out"
            }
        }

        $Image.Save($res)
        $Image.Dispose()

        gi $res
    }
}

function screenshot
{
    <#
    .synopsis
    Take screenshot
    .parameter Process
    Process to take screenshot of, $null for full screen
    .parameter NoTag
    Don't tag image with 'screenshot' or 'screenshot:<process name>'
    #>

    param(
        [System.Diagnostics.Process]$Process,
        [switch]$Tag = $false)

    $bounds = $null
    if ($Process)
    {
        $rect = New-Object PInvoke.RECT
        if ([PInvoke.NativeMethods]::GetWindowRect($Process.MainWindowHandle, [ref]$rect))
        {
            [PInvoke.NativeMethods]::ScaleToPixels([ref]$rect)
            $bounds = [Drawing.Rectangle]::new($rect.Left, $rect.Top, $rect.Right - $rect.Left, $rect.Bottom - $rect.Top)
        }
    }
    if (!$bounds)
    {
        $rect = New-Object PInvoke.RECT
        [PInvoke.NativeMethods]::GetScreenSize([ref]$rect)
        $bounds = [Drawing.Rectangle]::new($rect.Left, $rect.Top, $rect.Right - $rect.Left, $rect.Bottom - $rect.Top)
        # $bounds = [Windows.Forms.SystemInformation]::VirtualScreen
    }

    $screen = New-Object Drawing.Bitmap $bounds.Width, $bounds.Height
    $graphics = [Drawing.Graphics]::FromImage($screen)
    $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.Size, [System.Drawing.CopyPixelOperation]::SourceCopy)
    $graphics.Dispose()
    if (!$NoTag) {
        if ($Process) {
            $screen.Tag = "screenshot:$($Process.Name)"    
        }
        else {
            $screen.Tag = "screenshot"
        }        
    }
    return $screen
}

function printscreen
{
    <#
    .synopsis
    Take screenshot via PrtScr button
    .parameter Active
    Screen active window (Alt+PrtScr)
    #>

    param(
        [switch]$Active)

    $key = "{PrtSc}"
    if ($Active) { $key = "%{PrtSc}" }

    [Windows.Forms.SendKeys]::SendWait($key)
    [Windows.Forms.Clipboard]::GetImage()
}

function test-ctrlc
{
    <#
    .synopsis
    Catch Ctrl-C and throw error
    #>

    if ($host.ui.rawui.keyavailable -and (3 -eq [int]$host.ui.rawui.readkey("AllowCtrlC,IncludeKeyUp,NoEcho").character)) {
        throw (new-object ExecutionEngineException "Ctrl+C pressed")
    }
}

function timer
{
    <#
    .synopsis
    Run some code with interval
    .parameter interval
    Interval in msecs (default 1 sec)
    .parameter count
    Times to invoke (infinite by default)
    .parameter script
    Script to run
    #>

    param(
        [scriptblock]$script,
        [int]$interval=100,
        [int]$count=0)

    try {
        & {
            $n = 0
            while (($n -lt $count) -or ($count -eq 0))
            {
                test-ctrlc
                sleep -Milliseconds $interval
                test-ctrlc
                $n = $n + 1
                $n
            }
        } | % $script
    }
    catch [ExecutionEngineException] {
    }
}

function record
{
    <#
    .synopsis
    Record screenshots to make animation
    .parameter process
    Process to record
    .parameter interval
    Screenshots interval
    .parameter out
    Output directory
    .parameter foreground
    Record only when process is foreground
    #>

    param(
        [System.Diagnostics.Process]$process,
        [int]$interval = 100,
        [string]$out,
        [switch]$foreground)

    set-variable -name images -value @() -scope global
    timer -interval $interval -script {
        if (!$foreground -or ($process.MainWindowHandle -eq [PInvoke.NativeMethods]::GetForegroundWindow())) {
        screenshot -Process $process | % {
            $image = save-image -Out $out -Image $_
            write-host $image
            $images = get-variable -name images -scope global -valueonly
            set-variable -name images -value ($images + $image) -scope global
        }
            
        }
    }
    $images = get-variable -name images -scope global -valueonly
    remove-variable -name images -scope global
    $images
}

function animate
{
    <#
    .synopsis
    Make gif animation from images
    .parameter out
    Output file
    .parameter delay
    Gif delay
    .parameter file
    Files to convert
    .parameter delete
    Delete files after convert
    #>

    param(
        [Parameter(Mandatory=$true)]
        [string]$out,
        [int]$delay = 20,
        [Parameter(ValueFromPipeline=$true)]
        [System.IO.FileInfo]$file,
        [switch]$delete)

    begin {
        if (!(whereis convert)) {
            throw [System.IO.FileNotFoundException] "ImageMagick's convert not found"
        }
        $files = @()
    }
    process {
        $files = $files + $file
    }
    end {
        ($files | sort -Property LastWriteTime | % { """$($_.FullName)""" }) -join ' ' | convert -delay $delay '@-' -layers optimize $out
        if ($delete) {
            $files | rm
        }
        gi $out
    }
}

function now
{
    <#
    .synopsis
    Returns current date or/and time
    .parameter Date
    Return date only
    .parameter Time
    Return time only
    .parameter File
    Result can be used as name of file, i.e. no semicolons
    .parameter Format
    Specify format
    .example
    PS> now
    2015-01-25 07:22:27
    PS> now -Date
    2015-01-25
    PS> now -File
    2015-01-25 07.25.03
    PS> now -Format HH
    07
    #>

    param(
        [switch]$Date,
        [switch]$Time,
        [switch]$File,
        [string]$Format)

    $fmt = "yyyy-MM-dd HH:mm:ss"
    if ($File) { $fmt = "yyyy-MM-dd HH.mm.ss" }
    if ($Date) { $fmt = "yyyy-MM-dd" }
    if ($Time) {
        if ($File) { $fmt = "HH.mm.ss" }
        else { $fmt = "HH:mm:ss" }
    }
    if ($Format) { $fmt = $Format }
    get-date -f $fmt
}

function enumerate
{
    <#
    .synopsis
    Convert container into sequence of elements
    .parameter collection
    Collection object
    .example
    PS> gc test.json | json | enumerate | % Key | select -first 2
    foo
    bar
    #>

    param(
        [Parameter(ValueFromPipeline = $true)]
        [object]$collection)

    process {
        $collection.GetEnumerator()
    }
}

function numerate
{
    <#
    .synopsis
    Numerate files
    .parameter Format
    Template for name, where {0} stands for number, default is {0}
    .parameter Digits
    Number of digits
    .parameter Start
    Starting number
    #>

    param(
        [string]$Format = "{0}",
        [int]$Digits = 0,
        [int]$Start = 1,
        [Parameter(ValueFromPipeline = $true)]
        [string]$Path)

    begin
    {
        $fmt = $Format -f "{0:d$($Digits)}{1}"    
    }
    process
    {
        ren $Path ($fmt -f $Start, (gi $Path).Extension)
        $Start = $Start + 1
    }
}

function hash
{
    <#
    .synopsis
    Make hashtable from input objects
    .parameter key
    Key expression
    .parameter value
    Value expression
    .example
    PS> ls | hash Name FullName
    ...
    #>

    param(
        [object]$key = "Name",
        [object]$value = "Value",
        [Parameter(ValueFromPipeline = $true, ValueFromRemainingArguments = $true)]
        [object]$input)

    begin
    {
        $h = @{}
        if ($key -is [scriptblock]) { $key_expr = $key } else { $key_expr = { $_.$key } }
        if ($value -is [scriptblock]) { $value_expr = $value } else { $value_expr = { $_.$value } }
    }
    process
    {
        $input | % { $h += @{(& $key_expr)=(& $value_expr)} }
    }
    end
    {
        $h
    }
}

function template
{
    <#
    .synopsis Simple template language
    .parameter vars
    Replace variables
    .parameter eval
    Eval expressions with $(...) syntax
    .parameter dest
    Target directory (for templating files)
    .example
    PS> 'Hello $x, $y...' | template @{x=12;y=(21,22,23)}
    Hello 12, 21
    Hello 12, 22
    Hello 12, 23
    PS> gc '$name.txt'
    x is $x
    PS> ls '$name.txt' | template @{name='foo';x=10}; gc 'foo.txt'
    x is 10
    #>

    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [hashtable]$vars = (get-variable | hash Name Value),
        [switch]$eval,
        [Parameter(ValueFromPipeline = $true, ValueFromRemainingArguments = $true, ParameterSetName = "string")]
        [string]$input,
        [Parameter(ParameterSetName = "file")]
        [string]$dest = '.',
        [Parameter(ValueFromPipeline = $true, ValueFromRemainingArguments = $true, ParameterSetName = "file")]
        [system.io.filesysteminfo]$path)

    begin
    {
        $dest = resolve-path $dest
    }
    process
    {
        switch ($pscmdlet.parametersetname)
        {
            "string" {
                $input | % {
                    if ($eval -and ($_ -match "\`$\(([^\)]+)\)\.\.\."))
                    {
                        $line = $_
                        $s = [scriptblock]::Create($matches[1])
                        & $s | % { $line -replace "\`$\(([^\)]+)\)\.\.\.", "$($_)" } | template -vars $vars -eval:$eval
                    }
                    elseif ($eval -and ($_ -match "\`$\(([^\)]+)\)"))
                    {
                        $s = [scriptblock]::Create($matches[1])
                        $_ -replace "\`$\(([^\)]+)\)", (& $s) | template -vars $vars -eval:$eval
                    }
                    elseif ($_ -match "\`$(\w+)\.\.\.")
                    {
                        $line = $_
                        $vars.($matches[1]) | % { $line -replace "\`$(\w+)\.\.\.", "$($_)" } | template -vars $vars -eval:$eval
                    }
                    elseif ($_ -match "\`$(\w+)")
                    {
                        $_ -replace "\`$(\w+)", ($vars.($matches[1])) | template -vars $vars -eval:$eval
                    }
                    else
                    {
                        $_
                    }
                }
                break
            }
            "file" {
                $destpath = $path.Name | template -vars $vars -eval:$eval
                $destfull = join-path $dest $destpath
                if ($path.PSIsContainer)
                {
                    if (test-path $destfull -pathtype container) { $newpath = resolve-path $destfull }
                    else { $newpath = resolve-path (new-item $destfull -itemtype directory) }
                    ls $path.FullName | template -vars $vars -eval:$eval -dest $newpath
                }
                else
                {
                    gc $path.FullName | template -vars $vars -eval:$eval | out-file $destfull -encoding utf8
                }
            }
        }
    }
}

function verbs
{
    <#
    .synopsis Get verbs for file
    .parameter file
    File to get verbs for
    #>

    param(
        [string]$file)

    (new-object System.Diagnostics.ProcessStartInfo $file).Verbs
}

function watch-path
{
    <#
    .synopsis
    Watch for file/dir created, changed or deleted in path. Call with no actions to drop watcher (with name or path only)
    .parameter path
    Path to watch in
    .parameter filter
    File filter
    .parameter created
    Action on creation
    .parameter changed
    Action on changed
    .parameter deleted
    Action on deleted
    .parameter name
    Name of watcher to remove, default is guid generated and returned
    .parameter recurse
    Watch in subdirectories
    #>

    param(
        [string]$path = '.',
        [string]$filter = '*',
        [scriptblock]$created,
        [scriptblock]$changed,
        [scriptblock]$deleted,
        [scriptblock]$renamed,
        [string]$name,
        [switch]$recurse)

    if (!$created -and !$changed -and !$deleted -and !$renamed) {
        if ($name) {
            "created", "changed", "deleted", "renamed" | % {
                get-eventsubscriber "$($name)-$($_)" -erroraction ignore | % { unregister-event $_.sourceidentifier }
                get-event "$($name)-$($_)" -erroraction ignore | % { remove-event $_.sourceidentifier }
            }
        }
        elseif ($path) {
            get-eventsubscriber | ? {
                ([IO.FileSystemWatcher]$_.SourceObject).Path -eq (resolve-path $path)
            } | % { unregister-event $_.sourceidentifier }
        }
        else {
            throw 'Specify name or path to unwatch'
        }
    }
    else {
        trap { throw "Error watching $($path)" }
        $fsw = new-object IO.FileSystemWatcher (resolve-path $path), $filter -Property @{ IncludeSubdirectories = $recurse; NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite' }
        if (!$name) { $name = (new-guid).Guid }
        register-objectevent $fsw created -sourceidentifier "$($name)-created" -Action $created | out-null
        register-objectevent $fsw changed -sourceidentifier "$($name)-changed" -Action $changed | out-null
        register-objectevent $fsw deleted -sourceidentifier "$($name)-deleted" -Action $deleted | out-null
        register-objectevent $fsw renamed -sourceidentifier "$($name)-renamed" -Action $renamed | out-null
        $name
    }
}

function beep
{
    <#
    .synopsis
    Beep in console
    .parameter freq
    Frequency
    .parameter duration
    Duration in msecs
    #>

    param(
        [int]$freq = 500,
        [int]$duration = 250)

    [console]::Beep($freq, $duration)
}

function escape
{
    <#
    .synopsis
    Escape quotes
    #>

    param(
        [string]$str)

    $str -creplace '\\"', '\\"' -creplace '"', '\"'
}

function update-file
{
    <#
    .synopsis
    Touch file: create empty or update timestamp
    #>

    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$file)

    if (test-path $file) {
        (gi $file).LastWriteTime = get-date
    }
    else {
        echo $null | out-file $file -encoding ascii
    }
}

new-alias touch update-file -scope global -erroraction silentlycontinue

function expand
{
    <#
    .synopsis
    Expand string
    #>

    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$s,
        [hashtable]$vars)

    # Use some name, that can't present in script as expandable variable
    set-variable ':script:' $s -scope local

    if ($vars) {
        $vars | enumerate | % { set-variable $_.Name $_.Value -scope local }
    }

    $ExecutionContext.InvokeCommand.ExpandString((get-variable ':script:' -scope local -valueonly))
}

function closure
{
    <#
    .synopsis
    Catch variables in closure via string expansion
    #>

    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [scriptblock]$s,
        [hashtable]$vars)

    [scriptblock]::create((expand ([string]$s) $vars))
}
