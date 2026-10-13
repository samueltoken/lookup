[CmdletBinding()]
param([string]$FixtureRoot)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\build\legacy-cleanup.ps1')
if (-not $FixtureRoot) { $FixtureRoot = Join-Path $PSScriptRoot ('..\output\verification\installer-fixtures-' + [Guid]::NewGuid().ToString('N')) }
$FixtureRoot = [IO.Path]::GetFullPath($FixtureRoot)
if (Test-Path -LiteralPath $FixtureRoot) { throw 'Fixture root must be a new directory.' }
[void](New-Item -ItemType Directory -Path $FixtureRoot)
$script:Assertions = 0
function Assert([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw ('Assertion failed: ' + $Message) }
    $script:Assertions++
}
function Put([string]$Path, [string]$Text) {
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($Path))
    [IO.File]::WriteAllText($Path, $Text, (New-Object Text.UTF8Encoding($false)))
}
function Fixture([string]$Name) {
    $base = Join-Path $FixtureRoot $Name
    $envData = @{LocalAppData=(Join-Path $base 'Local'); RoamingAppData=(Join-Path $base 'Roaming'); Documents=(Join-Path $base 'Documents');
        Desktop=(Join-Path $base 'Desktop'); Programs=(Join-Path $base 'Programs'); CommonDesktop=(Join-Path $base 'CommonDesktop');
        CommonPrograms=(Join-Path $base 'CommonPrograms'); ProgramFiles=(Join-Path $base 'ProgramFiles'); ProgramFilesX86=(Join-Path $base 'ProgramFilesX86')}
    return @{Environment=$envData; Root=(Join-Path $envData.LocalAppData 'Programs\lookup'); Base=$base}
}
function Marker([string]$Root, [object[]]$Files = @()) {
    if (-not (Test-Path -LiteralPath (Join-Path $Root 'lookup.exe'))) { Put (Join-Path $Root 'lookup.exe') 'old' }
    foreach ($core in @('lookup.exe','resources\app.asar')) {
        if ((Test-Path -LiteralPath (Join-Path $Root $core)) -and @($Files | Where-Object { $_.path -eq $core.Replace('\','/') }).Count -eq 0) {
            $Files += @{path=$core.Replace('\','/'); sha256=(Get-FileHash -LiteralPath (Join-Path $Root $core)).Hash}
        }
    }
    Put (Join-Path $Root 'resources\lookup-install.json') (@{schemaVersion=1;appId='com.lookup.pdfviewer';productName='lookup';executable='lookup.exe';packageVersion='1.3.6'; files=$Files} | ConvertTo-Json -Depth 8)
}
function Asar([string]$Root, [string]$Name = 'lookup', [string]$Main = 'main.js') {
    $package = [Text.Encoding]::UTF8.GetBytes((@{name=$Name; main=$Main} | ConvertTo-Json -Compress))
    $json = [Text.Encoding]::UTF8.GetBytes((@{files=@{'package.json'=@{size=$package.Length; offset='0'}}} | ConvertTo-Json -Depth 8 -Compress))
    $padded = [int]([Math]::Ceiling(($json.Length + 4) / 4.0) * 4)
    $memory = New-Object IO.MemoryStream
    $writer = New-Object IO.BinaryWriter($memory)
    $writer.Write([uint32]4); $writer.Write([uint32]($padded + 4)); $writer.Write([uint32]$padded); $writer.Write([uint32]$json.Length)
    $writer.Write($json)
    for ($i = $json.Length + 4; $i -lt $padded; $i++) { $writer.Write([byte]0) }
    $writer.Write($package)
    $path = Join-Path $Root 'resources\app.asar'
    [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($path))
    [IO.File]::WriteAllBytes($path, $memory.ToArray())
    $writer.Dispose(); $memory.Dispose()
}
function Record([string]$Key, [hashtable]$Values, [string]$Command = '', [string]$Hive = 'CurrentUser', [string]$View = 'Registry64') {
    return [pscustomobject]@{Key=$Key; Values=$Values; Command=$Command; Hive=$Hive; View=$View}
}
function Plan($Fixture, [object[]]$Registry = @(), [object[]]$Shortcuts = @(), [string[]]$Running = @(), [switch]$MachineOnly) {
    return New-LookupCleanupPlan -InstallDir $Fixture.Root -EnvironmentData $Fixture.Environment -RegistrySnapshot $Registry -ShortcutSnapshot $Shortcuts -RunningPaths $Running -MachineOnly:$MachineOnly
}

# Ownership from the real ASAR pickle layout; malformed/wrong packages are kept.
$f = Fixture 'asar'
Asar $f.Root
Put (Join-Path $f.Root 'lookup.exe') 'old executable'
Assert (Read-LookupAsarIdentity (Join-Path $f.Root 'resources\app.asar')) 'valid ASAR identity'
$p = Plan $f
Assert ($p.fileCount -eq 2) 'owned executable and ASAR are eligible'
$result = Invoke-LookupCleanupPlan $p -FileSystemOnly
Assert $result.success 'owned distribution removed'
Assert (-not (Test-Path -LiteralPath $f.Root)) 'empty installation root removed'
$f = Fixture 'wrong-asar'
Asar $f.Root 'other-app'
Put (Join-Path $f.Root 'lookup.exe') 'other application'
Assert ((Plan $f).fileCount -eq 0) 'wrong package name cannot prove ownership'
Put (Join-Path $f.Root 'resources\app.asar') 'truncated'
Assert (-not (Read-LookupAsarIdentity (Join-Path $f.Root 'resources\app.asar'))) 'truncated ASAR rejected'
Put (Join-Path $f.Root 'resources\lookup-install.json') '{"appId":"com.lookup.pdfviewer"}'
Assert ((Plan $f).fileCount -eq 0) 'incomplete marker cannot prove ownership'

# Marker manifests cannot authorize personal documents/path traversal; unknown files/settings survive.
$f = Fixture 'manifest'
Put (Join-Path $f.Root 'lookup.exe') 'old executable'
Put (Join-Path $f.Root 'saved.pdf') 'personal document'
Put (Join-Path $f.Root 'settings.json') 'personal settings'
Put (Join-Path $f.Base 'outside.dll') 'outside'
$saved = (Get-FileHash -LiteralPath (Join-Path $f.Root 'saved.pdf')).Hash
$settings = (Get-FileHash -LiteralPath (Join-Path $f.Root 'settings.json')).Hash
Marker $f.Root @(@{path='saved.pdf';sha256=$saved}, @{path='settings.json';sha256=$settings}, @{path='../../outside.dll';sha256=('0' * 64)})
Put (Join-Path $f.Environment.RoamingAppData 'lookup\settings.json') 'existing preferences'
[void][IO.Directory]::CreateDirectory((Join-Path $f.Root 'my-empty-folder'))
$p = Plan $f
Assert ($p.fileCount -eq 2) 'manifest allows only known executable/marker'
$result = Invoke-LookupCleanupPlan $p -FileSystemOnly
Assert $result.success 'selective cleanup succeeds'
foreach ($path in @((Join-Path $f.Root 'saved.pdf'),(Join-Path $f.Root 'settings.json'),(Join-Path $f.Base 'outside.dll'),
    (Join-Path $f.Environment.RoamingAppData 'lookup\settings.json'),(Join-Path $f.Root 'my-empty-folder'))) { Assert (Test-Path -LiteralPath $path) 'personal/unknown items remain' }

# A changed known file is preserved even if its name is on the fallback allowlist.
$f = Fixture 'manifest-hash'
Put (Join-Path $f.Root 'lookup.exe') 'modified executable'
Asar $f.Root
Marker $f.Root @(@{path='lookup.exe';sha256=('0' * 64)})
Assert (-not (@((Plan $f).files.Path) -contains (Join-Path $f.Root 'lookup.exe'))) 'manifest hash mismatch is preserved'

# Ownership evidence and all other files survive changes after planning or a lock.
$f = Fixture 'changed-after-plan'
Marker $f.Root
Put (Join-Path $f.Root 'lookup.exe') 'old'
Put (Join-Path $f.Root 'ffmpeg.dll') 'old dll'
$p = Plan $f
Put (Join-Path $f.Root 'ffmpeg.dll') 'changed dll'
$result = Invoke-LookupCleanupPlan $p -FileSystemOnly
Assert (-not $result.success -and $result.removedFiles -eq 0) 'changed leaf aborts before any mutation'
Assert (Test-Path -LiteralPath (Join-Path $f.Root 'resources\lookup-install.json')) 'proof retained after changed leaf'
$f = Fixture 'locked'
Marker $f.Root
Put (Join-Path $f.Root 'lookup.exe') 'old'
$p = Plan $f
$lock = [IO.File]::Open((Join-Path $f.Root 'lookup.exe'), [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
try { $result = Invoke-LookupCleanupPlan $p -FileSystemOnly } finally { $lock.Dispose() }
Assert (-not $result.success -and $result.removedFiles -eq 0) 'locked executable aborts cleanup'
Assert (Test-Path -LiteralPath (Join-Path $f.Root 'resources\lookup-install.json')) 'proof retained after lock'
$p = Plan $f -Running @((Join-Path $f.Root 'resources\accdb\java\bin\java.exe'))
Assert ($p.errors.Count -gt 0 -and $p.fileCount -eq 0) 'child runtime process blocks candidate'

# Reparse directories and every ancestor are rejected; targets never traversed.
$f = Fixture 'junction'
$target = Join-Path $f.Base 'outside'
Put (Join-Path $target 'lookup\lookup.exe') 'protected target'
Marker (Join-Path $target 'lookup')
[void][IO.Directory]::CreateDirectory((Split-Path -LiteralPath $f.Root))
[void](New-Item -ItemType Junction -Path $f.Root -Target (Join-Path $target 'lookup'))
Assert (-not (Test-LookupNoReparsePoint $f.Root)) 'candidate junction rejected'
Assert ((Plan $f).fileCount -eq 0) 'junction cleanup plan has no files'
Assert (Test-Path -LiteralPath (Join-Path $target 'lookup\lookup.exe')) 'junction target unchanged'
$f = Fixture 'ancestor-junction'
$target = Join-Path $f.Base 'outside'
Put (Join-Path $target 'Programs\lookup\lookup.exe') 'protected ancestor target'
Marker (Join-Path $target 'Programs\lookup')
[void][IO.Directory]::CreateDirectory($f.Base)
[void](New-Item -ItemType Junction -Path $f.Environment.LocalAppData -Target $target)
Assert ((Plan $f).fileCount -eq 0) 'ancestor junction rejected'
$f = Fixture 'child-junction'
Marker $f.Root
Put (Join-Path $f.Root 'lookup.exe') 'old'
$target = Join-Path $f.Base 'outside'
Put (Join-Path $target 'icon.ico') 'protected child target'
[void](New-Item -ItemType Junction -Path (Join-Path $f.Root 'resources\linked') -Target $target)
$p = Plan $f
Assert (-not (@($p.files.Path) -contains (Join-Path $target 'icon.ico'))) 'child link target never included'
Assert (@($p.preserved) -contains (Join-Path $f.Root 'resources\linked')) 'child junction preserved'

# Only known, dead/owned shortcut targets are eligible; user/external targets remain.
$f = Fixture 'shortcuts'
$known = Join-Path $f.Root 'lookup.exe'
$good = Join-Path $f.Environment.Programs 'lookup\lookup.lnk'
$other = Join-Path $f.Environment.Desktop 'lookup.lnk'
Put $good 'fake known-target shortcut'
Put $other 'fake unrelated-target shortcut'
$p = Plan $f -Shortcuts @([pscustomobject]@{Path=$good;Target=$known;Machine=$false},[pscustomobject]@{Path=$other;Target=(Join-Path $f.Base 'another\lookup.exe');Machine=$false})
Assert ($p.fileCount -eq 1 -and $p.files[0].Path -eq $good) 'exact dead shortcut selected'
Put $known 'unowned app named lookup'
Assert ((Plan $f -Shortcuts @([pscustomobject]@{Path=$good;Target=$known;Machine=$false})).fileCount -eq 0) 'live unowned target shortcut preserved'

# Stale exact defaults and OpenWith names only; never touch UserChoice or a portable copy.
$f = Fixture 'registry'
$records = @(
    (Record 'Software\Classes\.pdf' @{''='lookup PDF';Other='keep'}),
    (Record 'Software\Classes\.pdf\OpenWithProgids' @{'lookup.PDF'='';'lookup PDF'='';'MSEdgePDF'=''}),
    (Record 'Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\UserChoice' @{ProgId='lookup.PDF'}),
    (Record 'Software\Classes\.xlsx' @{''='Excel.Sheet.12'}),
    (Record 'Software\Classes\.pdf' @{''='lookup.PDF'} '' 'LocalMachine')
)
$p = Plan $f -Registry $records
Assert ($p.registryCount -eq 4) 'exact stale defaults and two OpenWith values only'
Assert (@($p.registry | Where-Object { $_.Key -match 'UserChoice' }).Count -eq 0) 'UserChoice never planned'
Assert $p.needsAdmin 'machine stale default requests elevation'
$portable = Join-Path $f.Base 'portable\lookup.exe'
Put $portable 'portable copy'
$records = @((Record 'Software\Classes\lookup PDF' @{} ('"' + $portable + '" "%1"')),
    (Record 'Software\Classes\.pdf' @{''='lookup PDF'}),
    (Record 'Software\Classes\.pdf\OpenWithProgids' @{'lookup PDF'=''}))
Assert ((Plan $f -Registry $records).registryCount -eq 0) 'live portable class/default/OpenWith preserved'
$records[0].Hive = 'LocalMachine'
Assert ((Plan $f -Registry $records).registryCount -eq 0) 'portable class inherited from HKLM preserves HKCU default/OpenWith'
$keys = @(Get-LookupRegistryPaths $false)
Assert ($keys.Count -eq 69) 'all exact registry snapshot paths enumerated separately'
Assert ($keys -contains ('Software\' + $script:LookupCleanupGuid)) 'install GUID key in snapshot path list'
Assert ($keys -contains ('Software\Microsoft\Windows\CurrentVersion\Uninstall\' + $script:LookupCleanupGuid)) 'uninstall GUID key in snapshot path list'
Assert ($keys -contains 'Software\Classes\Applications\lookup.exe') 'Applications key in snapshot path list'
Assert (@($keys | Where-Object { $_ -match 'UserChoice' }).Count -eq 0) 'UserChoice never even read'

# An exact historic uninstall GUID can prove the old installation even without InstallLocation.
$f = Fixture 'guid'
Put (Join-Path $f.Root 'lookup.exe') 'old'
Put (Join-Path $f.Root 'Uninstall lookup.exe') 'old uninstaller'
$record = Record ('Software\Microsoft\Windows\CurrentVersion\Uninstall\' + $script:LookupCleanupGuid) @{DisplayName='lookup 1.3.4';UninstallString=('"' + (Join-Path $f.Root 'Uninstall lookup.exe') + '" /currentuser')}
$p = Plan $f -Registry @($record)
Assert ($p.fileCount -eq 2 -and $p.owned[0].Identity -eq 'guid') 'precise GUID uninstall command proves ownership'

# Elevated phase can include only fixed machine candidates/common shortcuts/HKLM.
$f = Fixture 'machine-only'
Put (Join-Path $f.Root 'lookup.exe') 'user app'
Marker $f.Root
$machineRoot = Join-Path $f.Environment.ProgramFiles 'lookup'
Put (Join-Path $machineRoot 'lookup.exe') 'machine app'
Marker $machineRoot
$records = @((Record 'Software\Classes\.pdf' @{''='lookup PDF'}), (Record 'Software\Classes\.pdf' @{''='lookup.PDF'} '' 'LocalMachine'))
$p = Plan $f -Registry $records -MachineOnly
Assert (@($p.files | Where-Object { -not $_.Machine }).Count -eq 0) 'elevated phase excludes user files'
Assert (@($p.registry | Where-Object { $_.Hive -ne 'LocalMachine' }).Count -eq 0) 'elevated phase excludes HKCU'
Assert ($p.fileCount -eq 2 -and $p.registryCount -eq 1) 'machine-only scope is fixed'

# Shared 32/64 views may repeat one operation. An absent value is idempotent;
# an existing changed value is retained. The registry base below is entirely fake.
$script:FakeKey = [pscustomobject]@{Values=@{''='lookup PDF';Other='keep'}}
$script:FakeKey | Add-Member ScriptMethod GetValueNames { return [string[]]@($this.Values.Keys) }
$script:FakeKey | Add-Member ScriptMethod GetValue { param($Name,$Default,$Options) return $this.Values[$Name] }
$script:FakeKey | Add-Member ScriptMethod DeleteValue { param($Name,$Missing) $this.Values.Remove($Name) }
$script:FakeKey | Add-Member ScriptMethod Dispose { }
$script:FakeBase = [pscustomobject]@{}
$script:FakeBase | Add-Member ScriptMethod OpenSubKey { param($Key,$Writable) return $script:FakeKey }
$script:FakeBase | Add-Member ScriptMethod Dispose { }
function Get-LookupRegistryBase { return $script:FakeBase }
$operation = [pscustomobject]@{Hive='CurrentUser';View='Registry64';Key='Software\Classes\.pdf';Kind='Value';Name='';Value='lookup PDF'}
Remove-LookupRegistryOperation $operation
Remove-LookupRegistryOperation $operation
Assert (-not $script:FakeKey.Values.ContainsKey('') -and $script:FakeKey.Values.Other -eq 'keep') 'repeated shared-view deletion is selective/idempotent'
$script:FakeKey.Values[''] = 'MSEdgePDF'
$failed = $false
try { Remove-LookupRegistryOperation $operation } catch { $failed = $true }
Assert ($failed -and $script:FakeKey.Values[''] -eq 'MSEdgePDF') 'changed registry default remains'

# No live registry deletion: invoke the actual Plan entrypoint against snapshots by
# testing the same typed return path through a mocked adapter, then parse JSON.
function Get-LookupEnvironment { return (Fixture 'entrypoint').Environment }
function Get-LookupRegistrySnapshot { return @() }
function Get-LookupShortcutSnapshot { return @() }
function Get-LookupRunningPaths { return @() }
$InstallDir = $null; $Phase = 'Plan'; $MachineOnly = $false; $IncludeMachine = $false; $AllowElevation = $false
$outcome = Invoke-LookupCleanupMain
Assert ($outcome.Code -eq 0) 'Plan entrypoint typed zero exit status'
$parsed = $outcome.Report | ConvertTo-Json -Depth 12 -Compress | ConvertFrom-Json
Assert ($parsed.schema -eq 1 -and $parsed.fileCount -eq 0) 'Plan entrypoint is one valid JSON document'
$Phase = 'Apply'
$outcome = Invoke-LookupCleanupMain
Assert ($outcome.Code -eq 0 -and $outcome.Report.success) 'Apply empty fixture exits successfully'
function Get-LookupRegistrySnapshot { return @((Record 'Software\Classes\.pdf' @{''='lookup.PDF'} '' 'LocalMachine')) }
function Start-Process {
    [CmdletBinding()]param($FilePath,$ArgumentList,$Verb,$WindowStyle,[switch]$Wait,[switch]$PassThru)
    Assert ($Verb -eq 'RunAs' -and $WindowStyle -eq 'Hidden') 'elevation uses explicit hidden RunAs'
    Assert ($ArgumentList -match '-IncludeMachine -MachineOnly' -and $ArgumentList -notmatch '-InstallDir') 'elevation cannot inherit user directory'
    throw (New-Object OperationCanceledException('Mock administrator cancellation'))
}
$AllowElevation = $true
$outcome = Invoke-LookupCleanupMain
Assert ($outcome.Code -ne 0 -and -not $outcome.Report.success -and $outcome.Report.errors.Count -gt 0) 'administrator cancellation fails visibly'

[pscustomobject]@{success=$true; assertions=$script:Assertions; fixtureRoot=$FixtureRoot; liveRegistryMutations=0} | ConvertTo-Json -Compress
