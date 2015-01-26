PSUtils
=======

My PowerShell utils<br>
Place it in `%MyDocuments%\WindowsPowerShell\Modules` and `Import-Module PSUtils`

* `json`, `unjson` - encode/decode JSON to/from PSObject
* `clear-clibboard`, `set-clipboard`, `out-clipboard`, `get-clipboard` with aliases `clear-clip`, `set-clip`, `out-clip`, `get-clip` - clipboard utils, `out-clip` appends data to clipboard when called in `foreach`
* `get-handle` - get handles with SysInternals handle util, returns `PSObject`s
* `select-group` - match & select group in one function, `| select-group 'foo(.*)'` is easier to write than `| ? { $_ -match 'foo(.*)' | % { $matches[1] }`
* `tags` - get tags of audio file with TagLib library
* `whereis` - locate executable with `where` util
* `invoke-cmd` - invoke .bat file and update environment variables changes
* `screenshot`, `save-screenshot` - make (save) screenshot of display or window
* `timer` - run script with interval
* `get-variables` - get variables as `hashtable`
* `template` - replace variables in text, `'$x = $y...' | template @{x=10;y=(1,2,3)}` ⇒ `10 = 1`, `10 = 2`, `10 = 3`

* `download` - download page
* `time` - measure script block running time
* `unzip`, `zip`, `lszip` - very simple zip utils in top of 7z
* `codepage` - get/set codepage or run block with specified codepage
* `encoding` - get/set encoding or run block with specified encoding, it's useful to run some cmd tool with $OutputEncoding set
* `path` - just `GetFolderPath`
* `rhistory` - search in history
* `dictionary` (alias `dict`) and `gdict` - working with simple dictionary in top of JSON
