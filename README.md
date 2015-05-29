PSUtils
=======

My PowerShell utils<br>
Place it in `%MyDocuments%\WindowsPowerShell\Modules` and `Import-Module PSUtils`

* `json`, `unjson` — encode/decode JSON to/from PSObject
* `clear-clibboard`, `set-clipboard`, `out-clipboard`, `get-clipboard` with aliases `clear-clip`, `set-clip`, `out-clip`, `get-clip` — clipboard utils, `out-clip` appends data to clipboard when called in `foreach`
* `get-handle` — get handles with SysInternals handle util, returns `PSObject`s
* `select-group` — match & select group in one function, `| select-group 'foo(.*)'` is easier to write than `| ? { $_ -match 'foo(.*)' | % { $matches[1] }`
* `tags` — get tags of audio file with TagLib library
* `whereis` — locate executable with `where` util
* `invoke-cmd` — invoke .bat file and update environment variables changes
* `screenshot`, `printscreen` — take screenshot with `Drawing.Graphics` or simulating `PrtScr` button
* `save-image` — save image to file, use it with function above: `screenshot | save-image pics`
* `timer` — run script with interval
* `hash` — create `hashtable` from input, `ls | hash Name FullName`
* `template` — replace variables in text/files, on string input replaces variables `'$x = $y...' | template @{x=10;y=(1,2,3)}` ⇒ `10 = 1`, `10 = 2`, `10 = 3`, on file/path input — replace variables both in file/path names and in its contents and copies substituted files to destination directory
* `enumerate` — convert container to sequence of elements
* `numerate` — numerate files, for example `ls *.png | sort -Property CreationTime | numerate -Digits 2` will rename to `01.png`, `02.png`, ... in order of `CreationTime`. `numerate -Digits 2 -Format 'foo {0}'` to specify template for new name, result will be `foo 01.png` etc.
* `watch-path` — watch for file creation/changing/deletion in path

* `download` — download page
* `time` — measure script block running time
* `unzip`, `zip`, `lszip` — very simple zip utils in top of 7z
* `codepage` — get/set codepage or run block with specified codepage
* `encoding` — get/set encoding or run block with specified encoding, it's useful to run some cmd tool with $OutputEncoding set
* `verbose` — set $VerbosePreference to `Continue` within script
* `verbs` — get verbs for file
* `path` — just `GetFolderPath`
* `import-profile` — reload profile scripts
* `rhistory` — search in history
* `environment` — get env as hashtable, or set env from hashtable
* `dictionary` (alias `dict`) and `gdict` — working with simple dictionary in top of JSON
