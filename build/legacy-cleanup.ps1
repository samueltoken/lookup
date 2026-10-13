# Windows PowerShell 5.1. This helper is extracted and run before new app files.
[CmdletBinding()]
param(
    [string]$InstallDir,
    [ValidateSet('Plan', 'Apply')][string]$Phase = 'Plan',
    [switch]$IncludeMachine,
    [switch]$MachineOnly,
    [switch]$AllowElevation
)

$script:LookupCleanupScriptPath = $PSCommandPath
$script:LookupCleanupGuid = '5c2d5693-c260-51f6-830f-eac5d14767ef'
$script:LookupCleanupExtensions = @('pdf','hwp','hwpx','doc','docx','xls','xlsx','xlsm','accdb','ppt','pptx','html','htm')

function Get-LookupFullPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    try {
        $full = [IO.Path]::GetFullPath($Path)
        if ($full -eq [IO.Path]::GetPathRoot($full)) { return $full }
        return $full.TrimEnd('\')
    } catch { return $null }
}

function Test-LookupSamePath([string]$First, [string]$Second) {
    $a = Get-LookupFullPath $First
    $b = Get-LookupFullPath $Second
    return $a -and $b -and [string]::Equals($a, $b, [StringComparison]::OrdinalIgnoreCase)
}

function Test-LookupPathInside([string]$Path, [string]$Root) {
    $p = Get-LookupFullPath $Path
    $r = Get-LookupFullPath $Root
    return $p -and $r -and ($p.StartsWith($r + '\', [StringComparison]::OrdinalIgnoreCase))
}

function Test-LookupNoReparsePoint([string]$Path) {
    # Check the candidate itself and every existing ancestor. Never resolve links.
    $p = Get-LookupFullPath $Path
    if (-not $p) { return $false }
    while ($p) {
        try {
            if (Test-Path -LiteralPath $p) {
                $item = Get-Item -LiteralPath $p -Force -ErrorAction Stop
                if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return $false }
            }
        } catch { return $false }
        $parent = [IO.Path]::GetDirectoryName($p)
        if ($parent -eq $p) { break }
        $p = $parent
    }
    return $true
}

function Read-LookupAsarIdentity([string]$Path) {
    if (-not (Test-LookupNoReparsePoint $Path)) { return $false }
    $stream = $null
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        $reader = New-Object IO.BinaryReader($stream)
        if ($stream.Length -lt 16 -or $reader.ReadUInt32() -ne 4) { return $false }
        $headerSize = $reader.ReadUInt32()
        if ($headerSize -lt 8 -or $headerSize -gt 16777216 -or 8L + $headerSize -gt $stream.Length) { return $false }
        $payloadSize = $reader.ReadUInt32()
        $jsonSize = $reader.ReadUInt32()
        if ($payloadSize -gt $headerSize - 4 -or $jsonSize -gt $headerSize - 8) { return $false }
        $header = [Text.Encoding]::UTF8.GetString($reader.ReadBytes([int]$jsonSize)) | ConvertFrom-Json -ErrorAction Stop
        $entry = $header.files.'package.json'
        if (-not $entry -or $entry.link -or $entry.unpacked) { return $false }
        $size = [long]$entry.size
        $offset = [long]$entry.offset
        if ($size -lt 2 -or $size -gt 1048576 -or $offset -lt 0 -or 8L + $headerSize + $offset + $size -gt $stream.Length) { return $false }
        $stream.Position = 8L + $headerSize + $offset
        $package = [Text.Encoding]::UTF8.GetString($reader.ReadBytes([int]$size)) | ConvertFrom-Json -ErrorAction Stop
        return $package.name -ceq 'lookup' -and $package.main -ceq 'main.js'
    } catch { return $false } finally { if ($stream) { $stream.Dispose() } }
}

function Read-LookupInstallMarker([string]$Root) {
    foreach ($relative in @('resources\lookup-install.json', 'lookup-install.json')) {
        $path = Join-Path $Root $relative
        if (-not (Test-Path -LiteralPath $path -PathType Leaf) -or -not (Test-LookupNoReparsePoint $path)) { continue }
        try {
            if ((Get-Item -LiteralPath $path).Length -gt 16777216) { continue }
            $marker = [IO.File]::ReadAllText($path) | ConvertFrom-Json -ErrorAction Stop
            if ($marker.appId -cne 'com.lookup.pdfviewer' -or $marker.schemaVersion -ne 1 -or
                $marker.productName -cne 'lookup' -or $marker.executable -cne 'lookup.exe' -or
                -not $marker.packageVersion -or -not $marker.files) { continue }
            # A JSON string alone cannot authorize deletion. Require an actual
            # packaged core file hash so a forged/obsolete marker is not proof.
            foreach ($entry in @($marker.files)) {
                $relative = ([string]$entry.path).Replace('/', '\')
                if (@('lookup.exe','resources\app.asar') -notcontains $relative -or $entry.sha256 -notmatch '^[a-fA-F0-9]{64}$') { continue }
                $core = Join-Path $Root $relative
                if ((Test-Path -LiteralPath $core -PathType Leaf) -and (Test-LookupNoReparsePoint $core) -and
                    (Get-FileHash -LiteralPath $core -Algorithm SHA256 -ErrorAction Stop).Hash -ieq $entry.sha256) { return $marker }
            }
        } catch { }
    }
    return $null
}

function Get-LookupEnvironment {
    return @{
        LocalAppData = [Environment]::GetFolderPath('LocalApplicationData')
        RoamingAppData = [Environment]::GetFolderPath('ApplicationData')
        Desktop = [Environment]::GetFolderPath('DesktopDirectory')
        Programs = [Environment]::GetFolderPath('Programs')
        CommonDesktop = [Environment]::GetFolderPath('CommonDesktopDirectory')
        CommonPrograms = [Environment]::GetFolderPath('CommonPrograms')
        ProgramFiles = ${env:ProgramW6432}
        ProgramFilesX86 = ${env:ProgramFiles(x86)}
        Documents = [Environment]::GetFolderPath('MyDocuments')
    }
}

function Get-LookupCandidates([string]$InstallDir, [hashtable]$EnvironmentData, [bool]$MachineOnly) {
    $result = New-Object 'System.Collections.Generic.List[object]'
    $paths = @()
    if (-not $MachineOnly) {
        if ($InstallDir) { $paths += @{Path=$InstallDir; Machine=$false} }
        if ($EnvironmentData.LocalAppData) { $paths += @{Path=(Join-Path $EnvironmentData.LocalAppData 'Programs\lookup'); Machine=$false} }
    }
    foreach ($programRoot in @($EnvironmentData.ProgramFiles, $EnvironmentData.ProgramFilesX86)) {
        if ($programRoot) { $paths += @{Path=(Join-Path $programRoot 'lookup'); Machine=$true} }
    }
    foreach ($entry in $paths) {
        $path = Get-LookupFullPath $entry.Path
        if (-not $path -or [IO.Path]::GetFileName($path) -ine 'lookup') { continue }
        # App settings and documents can never be an installation cleanup root.
        if ((Test-LookupSamePath $path $EnvironmentData.Documents) -or
            (Test-LookupPathInside $path $EnvironmentData.Documents) -or
            (Test-LookupSamePath $path $EnvironmentData.RoamingAppData) -or
            (Test-LookupPathInside $path $EnvironmentData.RoamingAppData)) { continue }
        $machine = [bool]$entry.Machine
        foreach ($machineRoot in @($EnvironmentData.ProgramFiles, $EnvironmentData.ProgramFilesX86)) {
            if ($machineRoot -and (Test-LookupPathInside $path $machineRoot)) { $machine = $true }
        }
        if ($result | Where-Object { Test-LookupSamePath $_.Path $path }) { continue }
        $result.Add([pscustomobject]@{Path=$path; Machine=$machine})
    }
    return $result.ToArray()
}

function Get-LookupClassNames {
    foreach ($ext in $script:LookupCleanupExtensions) {
        'lookup.' + $ext.ToUpperInvariant()
        'lookup ' + $ext.ToUpperInvariant()
    }
}

function Get-LookupRegistryPaths([bool]$MachineOnly) {
    $keys = @(('Software\' + $script:LookupCleanupGuid),
        ('Software\Microsoft\Windows\CurrentVersion\Uninstall\' + $script:LookupCleanupGuid),
        'Software\Classes\Applications\lookup.exe',
        'Software\Microsoft\Windows\CurrentVersion\App Paths\lookup.exe')
    foreach ($class in Get-LookupClassNames) { $keys += 'Software\Classes\' + $class }
    foreach ($ext in $script:LookupCleanupExtensions) {
        $keys += 'Software\Classes\.' + $ext
        $keys += 'Software\Classes\.' + $ext + '\OpenWithProgids'
        if (-not $MachineOnly) { $keys += 'Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.' + $ext + '\OpenWithProgids' }
    }
    return $keys
}

function Get-LookupRegistrySnapshot([bool]$MachineOnly) {
    $result = New-Object 'System.Collections.Generic.List[object]'
    $hives = @('LocalMachine')
    if (-not $MachineOnly) { $hives = @('CurrentUser', 'LocalMachine') }
    $keys = @(Get-LookupRegistryPaths $MachineOnly)
    foreach ($hive in $hives) {
        foreach ($view in @('Registry64', 'Registry32')) {
            $base = $null
            try {
                $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::$hive, [Microsoft.Win32.RegistryView]::$view)
                foreach ($keyPath in $keys) {
                    $key = $base.OpenSubKey($keyPath, $false)
                    if (-not $key) { continue }
                    try {
                        $values = @{}
                        foreach ($name in $key.GetValueNames()) { $values[$name] = $key.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) }
                        $commandKey = $key.OpenSubKey('shell\open\command')
                        $command = $null
                        if ($commandKey) { try { $command = $commandKey.GetValue('') } finally { $commandKey.Dispose() } }
                        $result.Add([pscustomobject]@{Hive=$hive; View=$view; Key=$keyPath; Values=$values; Command=$command})
                    } finally { $key.Dispose() }
                }
            } finally { if ($base) { $base.Dispose() } }
        }
    }
    return $result.ToArray()
}

function Get-LookupCommandTarget([string]$Command) {
    if ([string]::IsNullOrWhiteSpace($Command)) { return $null }
    # No expansion or execution. NSIS's quoted executable is the only accepted form.
    if ($Command -match '^\s*"([^"\r\n]+)"(?:\s|$)') { return Get-LookupFullPath $Matches[1] }
    if ($Command -notmatch '["\r\n]' -and $Command -match '(?i)\.exe$') { return Get-LookupFullPath $Command }
    return $null
}

function Test-LookupKnownTarget([string]$Target, [object[]]$Candidates, [bool]$Uninstaller = $false) {
    foreach ($candidate in $Candidates) {
        if (Test-LookupSamePath $Target (Join-Path $candidate.Path 'lookup.exe')) { return $true }
        if ($Uninstaller -and (Test-LookupSamePath $Target (Join-Path $candidate.Path 'Uninstall lookup.exe'))) { return $true }
    }
    return $false
}

function Test-LookupEligibleTarget([string]$Target, [object[]]$Candidates, [object[]]$Owned, [bool]$Uninstaller = $false) {
    if (-not (Test-LookupKnownTarget $Target $Candidates $Uninstaller) -or -not (Test-LookupNoReparsePoint $Target)) { return $false }
    if (-not (Test-Path -LiteralPath $Target)) { return $true }
    return @($Owned | Where-Object { Test-LookupPathInside $Target $_.Path }).Count -gt 0
}

function Test-LookupClassCleanupEligible([string]$ClassName, [object[]]$Snapshot, [object[]]$Candidates, [object[]]$Owned) {
    # Classes can be inherited across HKCU/HKLM and shared registry views. Preserve
    # an exact name if any definition still belongs to an external/unknown app.
    $records = @($Snapshot | Where-Object { $_.Key -ieq ('Software\Classes\' + $ClassName) })
    foreach ($record in $records) {
        if (-not (Test-LookupEligibleTarget (Get-LookupCommandTarget $record.Command) $Candidates $Owned)) { return $false }
    }
    return $true
}

function Test-LookupDistributionFile([string]$Relative) {
    $name = $Relative.Replace('/', '\').ToLowerInvariant()
    $rootFiles = @('lookup.exe','uninstall lookup.exe','uninstallericon.ico','lookup-install.json','chrome_100_percent.pak','chrome_200_percent.pak',
        'd3dcompiler_47.dll','dxcompiler.dll','dxil.dll','ffmpeg.dll','icudtl.dat','libegl.dll','libglesv2.dll',
        'license.electron.txt','licenses.chromium.html','resources.pak','snapshot_blob.bin','v8_context_snapshot.bin',
        'vk_swiftshader.dll','vk_swiftshader_icd.json','vulkan-1.dll')
    if ($rootFiles -contains $name) { return $true }
    if ($name -match '^locales\\[a-z0-9_-]+\.pak$') { return $true }
    if (@('resources\app.asar','resources\app-update.yml','resources\elevate.exe','resources\icon.ico','resources\icon.png',
        'resources\lookup-install.json','resources\libreoffice\.keep',
        'resources\security\filepathcheckermoduleexample.dll','resources\security\lookupfilepathchecker.dll') -contains $name) { return $true }
    if ($name -match '^resources\\accdb\\lib\\(ucanaccess|jackcess|hsqldb|commons-logging|commons-lang3)-[0-9.]+\.jar$') { return $true }
    if ($name -match '^resources\\accdb\\classes\\lookupaccdbdump(\$[a-z0-9_]+)?\.class$') { return $true }
    # Bundled Java runtime only. No document/file format extensions are allowed.
    if ($name -match '^resources\\accdb\\java\\(bin|conf|legal|lib)\\' -and
        ($name -match '\.(dll|exe|jar|dat|lib|cfg|bfc|jfc|properties|src|policy|certs|md)$' -or
        $name -match '\\(modules|classlist|tzmappings|cacerts|readme|license|notice|assembly_exception)$')) { return $true }
    if (@('resources\accdb\java\release','resources\accdb\java\notice') -contains $name) { return $true }
    # Native modules unpacked by Electron; arbitrary files beside them are kept.
    if ($name -match '^resources\\app\.asar\.unpacked\\node_modules\\' -and $name -match '\.(node|dll|exe|wasm)$') { return $true }
    return $false
}

function Get-LookupTree([string]$Root) {
    $items = New-Object 'System.Collections.Generic.List[object]'
    $stack = New-Object 'System.Collections.Generic.Stack[string]'
    $stack.Push($Root)
    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        if (-not (Test-LookupNoReparsePoint $dir)) { continue }
        foreach ($item in Get-ChildItem -LiteralPath $dir -Force -ErrorAction Stop) {
            $linked = ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0
            $items.Add([pscustomobject]@{Item=$item; Linked=$linked})
            if ($item.PSIsContainer -and -not $linked) { $stack.Push($item.FullName) }
            if ($items.Count -gt 100000) { throw 'Installation file limit exceeded.' }
        }
    }
    return $items.ToArray()
}

function Get-LookupShortcutSnapshot([hashtable]$EnvironmentData, [bool]$MachineOnly) {
    $roots = @(@{Desktop=$EnvironmentData.CommonDesktop; Programs=$EnvironmentData.CommonPrograms; Machine=$true})
    if (-not $MachineOnly) { $roots += @{Desktop=$EnvironmentData.Desktop; Programs=$EnvironmentData.Programs; Machine=$false} }
    $result = New-Object 'System.Collections.Generic.List[object]'
    $shell = New-Object -ComObject WScript.Shell
    try {
        foreach ($root in $roots) {
            $paths = @()
            if ($root.Desktop) { $paths += Join-Path $root.Desktop 'lookup.lnk' }
            if ($root.Programs) {
                $paths += Join-Path $root.Programs 'lookup.lnk'
                $paths += Join-Path $root.Programs 'lookup\lookup.lnk'
                $paths += Join-Path $root.Programs 'lookup\lookup Uninstall.lnk'
            }
            foreach ($path in $paths) {
                if ((Test-Path -LiteralPath $path -PathType Leaf) -and (Test-LookupNoReparsePoint $path)) {
                    $link = $shell.CreateShortcut($path)
                    $result.Add([pscustomobject]@{Path=$path; Target=$link.TargetPath; Machine=$root.Machine})
                }
            }
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell) }
    return $result.ToArray()
}

function Get-LookupRunningPaths {
    $paths = @()
    # Include child runtimes such as java.exe/soffice.exe, without terminating them.
    foreach ($process in Get-CimInstance Win32_Process -ErrorAction Stop) {
        if ($process.ExecutablePath) { $paths += $process.ExecutablePath }
    }
    return $paths
}

function New-LookupCleanupPlan {
    [CmdletBinding()]
    param([string]$InstallDir, [hashtable]$EnvironmentData, [object[]]$RegistrySnapshot,
        [object[]]$ShortcutSnapshot, [string[]]$RunningPaths = @(), [switch]$MachineOnly)
    # Snapshots are inputs so fixture tests never touch live registry or shortcuts.
    $candidates = @(Get-LookupCandidates $InstallDir $EnvironmentData ([bool]$MachineOnly))
    $files = New-Object 'System.Collections.Generic.List[object]'
    $directories = New-Object 'System.Collections.Generic.List[object]'
    $registry = New-Object 'System.Collections.Generic.List[object]'
    $preserved = New-Object 'System.Collections.Generic.List[string]'
    $errors = New-Object 'System.Collections.Generic.List[string]'
    $owned = New-Object 'System.Collections.Generic.List[object]'
    foreach ($candidate in $candidates) {
        if (-not (Test-Path -LiteralPath $candidate.Path -PathType Container)) { continue }
        if (-not (Test-LookupNoReparsePoint $candidate.Path)) { $preserved.Add($candidate.Path); continue }
        if (@($RunningPaths | Where-Object { (Test-LookupPathInside $_ $candidate.Path) -or (Test-LookupSamePath $_ $candidate.Path) }).Count -gt 0) {
            $errors.Add('lookup is running: ' + $candidate.Path); continue
        }
        $marker = Read-LookupInstallMarker $candidate.Path
        $identity = $null
        if ($marker) { $identity = 'marker' }
        elseif (Read-LookupAsarIdentity (Join-Path $candidate.Path 'resources\app.asar')) { $identity = 'asar' }
        else {
            foreach ($record in $RegistrySnapshot) {
                if ($record.Key -ieq ('Software\Microsoft\Windows\CurrentVersion\Uninstall\' + $script:LookupCleanupGuid) -and
                    ([string]$record.Values.DisplayName -match '^lookup(?: [0-9.]+)?$') -and
                    (Test-LookupSamePath (Get-LookupCommandTarget $record.Values.UninstallString) (Join-Path $candidate.Path 'Uninstall lookup.exe'))) { $identity = 'guid'; break }
            }
        }
        if (-not $identity) { $preserved.Add($candidate.Path); continue }
        $owned.Add([pscustomobject]@{Path=$candidate.Path; Machine=$candidate.Machine; Identity=$identity})
        $manifest = @{}
        if ($marker -and $marker.files) {
            foreach ($entry in $marker.files) {
                $relative = [string]$entry.path
                if (-not $relative -or [IO.Path]::IsPathRooted($relative) -or $relative -match '(^|[\\/])\.\.([\\/]|$)' -or $entry.sha256 -notmatch '^[a-fA-F0-9]{64}$') { continue }
                $resolved = Get-LookupFullPath (Join-Path $candidate.Path $relative)
                if ((Test-LookupPathInside $resolved $candidate.Path) -and (Test-LookupDistributionFile $relative)) { $manifest[$resolved] = ([string]$entry.sha256).ToUpperInvariant() }
            }
        }
        try {
            foreach ($entry in Get-LookupTree $candidate.Path) {
                $item = $entry.Item
                if ($entry.Linked) { $preserved.Add($item.FullName); continue }
                if ($item.PSIsContainer) { continue }
                $relative = $item.FullName.Substring($candidate.Path.Length + 1)
                if (-not (Test-LookupDistributionFile $relative)) { $preserved.Add($item.FullName); continue }
                $hash = (Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
                if ($manifest.ContainsKey($item.FullName) -and $hash -ine $manifest[$item.FullName]) { $preserved.Add($item.FullName); continue }
                $files.Add([pscustomobject]@{Path=$item.FullName; Root=$candidate.Path; Machine=$candidate.Machine; Hash=$hash; Length=$item.Length})
                $parent = [IO.Path]::GetDirectoryName($item.FullName)
                while ($parent -and (Test-LookupPathInside $parent $candidate.Path)) {
                    $directories.Add([pscustomobject]@{Path=$parent; Root=$candidate.Path; Machine=$candidate.Machine})
                    $parent = [IO.Path]::GetDirectoryName($parent)
                }
            }
            $directories.Add([pscustomobject]@{Path=$candidate.Path; Root=$candidate.Path; Machine=$candidate.Machine})
        } catch { $errors.Add($_.Exception.Message) }
    }
    foreach ($link in $ShortcutSnapshot) {
        if ($MachineOnly -and -not $link.Machine) { continue }
        if ((Test-LookupEligibleTarget $link.Target $candidates $owned.ToArray() $true) -and (Test-LookupNoReparsePoint $link.Path)) {
            $files.Add([pscustomobject]@{Path=$link.Path; Root=$null; Machine=[bool]$link.Machine; Target=$link.Target; Hash=(Get-FileHash -LiteralPath $link.Path -Algorithm SHA256).Hash; Length=(Get-Item -LiteralPath $link.Path).Length})
            $parent = Split-Path -LiteralPath $link.Path
            if ([IO.Path]::GetFileName($parent) -ieq 'lookup') { $directories.Add([pscustomobject]@{Path=$parent; Root=$parent; Machine=[bool]$link.Machine}) }
        }
    }
    $classes = @(Get-LookupClassNames)
    foreach ($record in $RegistrySnapshot) {
        if ($MachineOnly -and $record.Hive -ne 'LocalMachine') { continue }
        $machine = $record.Hive -eq 'LocalMachine'
        $key = [string]$record.Key
        $removeKey = $false
        if ($key -ieq ('Software\' + $script:LookupCleanupGuid) -or $key -ieq ('Software\Microsoft\Windows\CurrentVersion\Uninstall\' + $script:LookupCleanupGuid)) {
            $removeKey = Test-LookupEligibleTarget (Get-LookupCommandTarget $record.Values.UninstallString) $candidates $owned.ToArray() $true
            if (-not $removeKey -and $record.Values.InstallLocation) {
                $removeKey = @($candidates | Where-Object { (Test-LookupSamePath $_.Path $record.Values.InstallLocation) -and
                    (-not (Test-Path -LiteralPath $_.Path) -or @($owned | Where-Object { Test-LookupSamePath $_.Path $record.Values.InstallLocation }).Count -gt 0) }).Count -gt 0
            }
        } elseif ($key -ieq 'Software\Microsoft\Windows\CurrentVersion\App Paths\lookup.exe') {
            $removeKey = Test-LookupEligibleTarget ([string]$record.Values['']) $candidates $owned.ToArray()
        } elseif ($key -ieq 'Software\Classes\Applications\lookup.exe' -or $classes -contains ($key -replace '^Software\\Classes\\','')) {
            $removeKey = Test-LookupEligibleTarget (Get-LookupCommandTarget $record.Command) $candidates $owned.ToArray()
        }
        if ($removeKey) { $registry.Add([pscustomobject]@{Hive=$record.Hive; View=$record.View; Key=$key; Kind='Key'; Name=$null; Machine=$machine; Expected=$record}); continue }
        foreach ($ext in $script:LookupCleanupExtensions) {
            if ($key -ieq ('Software\Classes\.' + $ext) -and $classes -contains [string]$record.Values['']) {
                $className = [string]$record.Values['']
                if (Test-LookupClassCleanupEligible $className $RegistrySnapshot $candidates $owned.ToArray()) {
                    $registry.Add([pscustomobject]@{Hive=$record.Hive; View=$record.View; Key=$key; Kind='Value'; Name=''; Machine=$machine; Value=$record.Values['']})
                }
            }
            if ($key -ieq ('Software\Classes\.' + $ext + '\OpenWithProgids') -or $key -ieq ('Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.' + $ext + '\OpenWithProgids')) {
                foreach ($name in $record.Values.Keys) {
                    if ($name -ieq ('lookup.' + $ext) -or $name -ieq ('lookup ' + $ext)) {
                        if (Test-LookupClassCleanupEligible $name $RegistrySnapshot $candidates $owned.ToArray()) {
                            $registry.Add([pscustomobject]@{Hive=$record.Hive; View=$record.View; Key=$key; Kind='Value'; Name=$name; Machine=$machine; Value=$record.Values[$name]})
                        }
                    }
                }
            }
        }
    }
    $machineCount = @($files.ToArray() + $registry.ToArray() | Where-Object { $_.Machine }).Count
    return [pscustomobject]@{schema=1; files=$files.ToArray(); directories=$directories.ToArray(); registry=$registry.ToArray(); owned=$owned.ToArray(); preserved=$preserved.ToArray(); errors=$errors.ToArray(); fileCount=$files.Count; registryCount=$registry.Count; needsAdmin=($machineCount -gt 0); machineCount=$machineCount}
}

function Get-LookupRegistryBase([string]$Hive, [string]$View) {
    return [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]$Hive, [Microsoft.Win32.RegistryView]$View)
}

function Remove-LookupRegistryOperation($Operation) {
    $base = Get-LookupRegistryBase $Operation.Hive $Operation.View
    try {
        if ($Operation.Kind -eq 'Value') {
            $key = $base.OpenSubKey($Operation.Key, $true)
            if (-not $key) { return }
            try {
                if ($key.GetValueNames() -notcontains $Operation.Name) { return }
                $actual = $key.GetValue($Operation.Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                if (($actual | ConvertTo-Json -Compress) -cne ($Operation.Value | ConvertTo-Json -Compress)) { throw 'Registry value changed during cleanup.' }
                $key.DeleteValue($Operation.Name, $false)
            } finally { $key.Dispose() }
        } else {
            $key = $base.OpenSubKey($Operation.Key, $false)
            if (-not $key) { return }
            try {
                foreach ($name in $Operation.Expected.Values.Keys) {
                    $actual = $key.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                    if (($actual | ConvertTo-Json -Compress) -cne ($Operation.Expected.Values[$name] | ConvertTo-Json -Compress)) { throw 'Registry key changed during cleanup.' }
                }
                if (@($key.GetValueNames()).Count -ne $Operation.Expected.Values.Count) { throw 'Registry key changed during cleanup.' }
                if ($Operation.Expected.Command) {
                    $command = $key.OpenSubKey('shell\open\command')
                    if (-not $command) { throw 'Registry command changed during cleanup.' }
                    try { if ($command.GetValue('') -cne $Operation.Expected.Command) { throw 'Registry command changed during cleanup.' } } finally { $command.Dispose() }
                }
            } finally { $key.Dispose() }
            $base.DeleteSubKeyTree($Operation.Key, $false)
        }
    } finally { $base.Dispose() }
}

function Invoke-LookupCleanupPlan {
    [CmdletBinding()]
    param($Plan, [switch]$IncludeMachine, [switch]$FileSystemOnly)
    $errors = New-Object 'System.Collections.Generic.List[string]'
    foreach ($errorText in $Plan.errors) { $errors.Add($errorText) }
    $removedFiles = 0
    $removedRegistry = 0
    $selected = @($Plan.files | Where-Object { -not $_.Machine -or $IncludeMachine })
    # Verify every selected leaf before making any changes. Changed files, junction
    # swaps and an executable still in use abort while ownership evidence survives.
    if ($errors.Count -eq 0) {
        foreach ($file in $selected) {
            try {
                if (-not (Test-Path -LiteralPath $file.Path -PathType Leaf)) { continue }
                if (-not (Test-LookupNoReparsePoint $file.Path)) { throw 'Reparse point detected during cleanup.' }
                if ($file.Root -and -not (Test-LookupPathInside $file.Path $file.Root)) { throw 'File escaped installation root.' }
                $item = Get-Item -LiteralPath $file.Path -Force -ErrorAction Stop
                if ($item.Length -ne $file.Length -or (Get-FileHash -LiteralPath $file.Path -Algorithm SHA256).Hash -ine $file.Hash) { throw 'File changed during cleanup.' }
                if ([IO.Path]::GetFileName($file.Path) -iin @('lookup.exe','Uninstall lookup.exe')) {
                    $probe = [IO.File]::Open($file.Path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
                    $probe.Dispose()
                }
            } catch { $errors.Add($file.Path + ': ' + $_.Exception.Message); break }
        }
    }
    # A plan containing an in-use/scan failure is never partially applied.
    if ($errors.Count -eq 0) {
        foreach ($file in @($selected | Sort-Object @{Expression={ if ($_.Path -match '(app\.asar|lookup-install\.json|Uninstall lookup\.exe)$') { 1 } else { 0 } }})) {
            try {
                if (-not (Test-Path -LiteralPath $file.Path -PathType Leaf)) { continue }
                if (-not (Test-LookupNoReparsePoint $file.Path)) { throw 'Reparse point detected during cleanup.' }
                if ($file.Root -and -not (Test-LookupPathInside $file.Path $file.Root)) { throw 'File escaped installation root.' }
                $item = Get-Item -LiteralPath $file.Path -Force -ErrorAction Stop
                if ($item.Length -ne $file.Length -or (Get-FileHash -LiteralPath $file.Path -Algorithm SHA256).Hash -ine $file.Hash) { throw 'File changed during cleanup.' }
                Remove-Item -LiteralPath $file.Path -Force -ErrorAction Stop
                $removedFiles++
            } catch { $errors.Add($file.Path + ': ' + $_.Exception.Message); break }
        }
        foreach ($directory in @($Plan.directories | Group-Object Path | ForEach-Object { $_.Group[0] } | Sort-Object @{Expression={$_.Path.Length};Descending=$true})) {
            if ($errors.Count -gt 0) { break }
            if ($directory.Machine -and -not $IncludeMachine) { continue }
            try {
                if (-not (Test-Path -LiteralPath $directory.Path -PathType Container)) { continue }
                if (-not (Test-LookupNoReparsePoint $directory.Path)) { continue }
                if (-not (Test-LookupSamePath $directory.Path $directory.Root) -and -not (Test-LookupPathInside $directory.Path $directory.Root)) { throw 'Directory escaped installation root.' }
                if (@(Get-ChildItem -LiteralPath $directory.Path -Force -ErrorAction Stop).Count -eq 0) { Remove-Item -LiteralPath $directory.Path -Force -ErrorAction Stop }
            } catch { $errors.Add($directory.Path + ': ' + $_.Exception.Message) }
        }
        if ($errors.Count -eq 0 -and -not $FileSystemOnly) {
            foreach ($operation in $Plan.registry) {
                if ($operation.Machine -and -not $IncludeMachine) { continue }
                try { Remove-LookupRegistryOperation $operation; $removedRegistry++ } catch { $errors.Add($operation.Key + ': ' + $_.Exception.Message) }
            }
        }
    }
    return [pscustomobject]@{success=($errors.Count -eq 0); removedFiles=$removedFiles; removedRegistry=$removedRegistry; preservedCount=@($Plan.preserved).Count; errors=$errors.ToArray(); needsAdmin=$Plan.needsAdmin}
}

function Invoke-LookupCleanupMain {
    $environmentData = Get-LookupEnvironment
    $plan = New-LookupCleanupPlan -InstallDir $InstallDir -EnvironmentData $environmentData -RegistrySnapshot @(Get-LookupRegistrySnapshot ([bool]$MachineOnly)) -ShortcutSnapshot @(Get-LookupShortcutSnapshot $environmentData ([bool]$MachineOnly)) -RunningPaths @(Get-LookupRunningPaths) -MachineOnly:$MachineOnly
    if ($Phase -eq 'Plan') { return @{Code=0; Report=$plan} }
    $result = Invoke-LookupCleanupPlan -Plan $plan -IncludeMachine:$IncludeMachine
    if ($result.success -and $plan.needsAdmin -and -not $IncludeMachine) {
        if (-not $AllowElevation) { $result.success = $false; $result.errors += 'Machine leftovers require administrator permission.' }
        else {
            try {
                if ($script:LookupCleanupScriptPath -match '["\r\n]') { throw 'Invalid helper path.' }
                $arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "' + $script:LookupCleanupScriptPath + '" -Phase Apply -IncludeMachine -MachineOnly'
                $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
                $process = Start-Process -FilePath $powershell -ArgumentList $arguments -Verb RunAs -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
                if ($process.ExitCode -ne 0) { throw 'Administrator cleanup failed (exit ' + $process.ExitCode + ').' }
            } catch { $result.success = $false; $result.errors += $_.Exception.Message }
        }
    }
    if ($result.success) { return @{Code=0; Report=$result} }
    return @{Code=2; Report=$result}
}

if ($MyInvocation.InvocationName -ne '.') {
    try { $outcome = Invoke-LookupCleanupMain; $outcome.Report | ConvertTo-Json -Depth 12 -Compress; exit $outcome.Code }
    catch { [pscustomobject]@{success=$false; errors=@($_.Exception.Message)} | ConvertTo-Json -Compress; exit 2 }
}
