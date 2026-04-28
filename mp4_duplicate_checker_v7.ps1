
param(
  [Parameter(ValueFromRemainingArguments=$true)]
  [string[]]$InitialFolders
)

$ErrorActionPreference = 'Stop'
$ToolVersion = '7.0'
$SampleFps = 1
$HashSize = 8
$BitsPerFrame = $HashSize * $HashSize
$LikelyThreshold = 0.93
$TopLimit = 50
$script:ReportDirForErrors = ''

function Pause-End {
  Write-Host ''
  Write-Host 'Press Enter to close this window.'
  Read-Host | Out-Null
}

function Percent([int64]$Current, [int64]$Total) {
  if ($Total -le 0) { return 0 }
  return [int][Math]::Round(($Current * 100.0) / $Total, 0)
}

function Pair-Key($AId, $BId) {
  $a = [int]$AId
  $b = [int]$BId
  if ($a -le $b) { return ('{0}|{1}' -f $a, $b) }
  return ('{0}|{1}' -f $b, $a)
}

function Clean-PathText([string]$p) {
  if ([string]::IsNullOrWhiteSpace($p)) { return '' }
  $x = $p.Trim()
  if ($x.StartsWith('"') -and $x.EndsWith('"') -and $x.Length -ge 2) { $x = $x.Substring(1, $x.Length - 2) }
  if ($x.StartsWith("'") -and $x.EndsWith("'") -and $x.Length -ge 2) { $x = $x.Substring(1, $x.Length - 2) }
  return $x.Trim()
}

function Get-FolderFullPath([string]$p) {
  $x = Clean-PathText $p
  if (-not (Test-Path -LiteralPath $x -PathType Container)) { return $null }
  return (Get-Item -LiteralPath $x).FullName.TrimEnd('\','/')
}

function Add-Folder($folders, $seen, [string]$p) {
  $full = Get-FolderFullPath $p
  if ($null -eq $full) {
    Write-Host ('Invalid folder path: {0}' -f $p)
    return $false
  }

  $key = $full.ToLowerInvariant()
  if ($seen.ContainsKey($key)) {
    Write-Host ('Already added: {0}' -f $full)
    return $false
  }

  [void]$folders.Add($full)
  $seen[$key] = $true
  Write-Host ('Added folder: {0}' -f $full)
  return $true
}

function Ask-Folder([string]$label) {
  while ($true) {
    Write-Host ''
    Write-Host ('Paste the path for {0}, then press Enter.' -f $label)
    Write-Host 'Tip: in File Explorer, open the folder, click the address bar, copy the path, then paste it here.'
    $p = Read-Host ('{0} folder path' -f $label)
    $full = Get-FolderFullPath $p
    if ($null -ne $full) { return $full }
    Write-Host ('Invalid folder path: {0}' -f $p)
  }
}

function Show-Folders($folders) {
  Write-Host ''
  Write-Host ('Current folder count: {0}' -f $folders.Count)
  $n = 0
  foreach ($f in $folders) {
    $n++
    Write-Host ('  {0}. {1}' -f $n, $f)
  }
  Write-Host ''
}

function Resolve-Ffmpeg {
  $cmd = Get-Command ffmpeg -ErrorAction SilentlyContinue
  if ($null -ne $cmd) { return $cmd.Source }
  throw 'ffmpeg.exe was not found. Install FFmpeg and confirm ffmpeg -version works.'
}

function Collect-Mp4($folders, [bool]$recursive) {
  $rows = @()
  $seen = @{}
  $id = 0
  $folderIndex = 0

  foreach ($folder in $folders) {
    $folderIndex++

    if ($recursive) {
      $files = Get-ChildItem -LiteralPath $folder -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ieq '.mp4' } | Sort-Object FullName
    } else {
      $files = Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ieq '.mp4' } | Sort-Object FullName
    }

    foreach ($file in $files) {
      $key = $file.FullName.ToLowerInvariant()
      if ($seen.ContainsKey($key)) { continue }

      $seen[$key] = $true
      $id++

      $rel = $file.FullName
      if ($file.FullName.Length -gt $folder.Length) { $rel = $file.FullName.Substring($folder.Length).TrimStart('\','/') }

      $rows += [pscustomobject]@{
        FileId = $id
        SourceFolderIndex = $folderIndex
        SourceFolder = $folder
        RelativePath = $rel
        FileName = $file.Name
        FullPath = $file.FullName
        SizeBytes = [int64]$file.Length
      }
    }
  }

  return @($rows)
}

function Run-Exact($files, [string]$csvPath, [string]$hashErrorPath) {
  Write-Host ''
  Write-Host 'Stage 1: exact duplicate check'

  $sizeBuckets = @{}
  foreach ($r in $files) {
    $sizeKey = [string]$r.SizeBytes
    if (-not $sizeBuckets.ContainsKey($sizeKey)) {
      $sizeBuckets[$sizeKey] = @()
    }
    $sizeBuckets[$sizeKey] = @($sizeBuckets[$sizeKey]) + @($r)
  }

  $candidates = @()
  foreach ($k in $sizeBuckets.Keys) {
    $bucket = @($sizeBuckets[$k])
    if ($bucket.Count -gt 1) {
      $candidates += $bucket
    }
  }

  Write-Host ('Files requiring SHA-256 hashing: {0}' -f $candidates.Count)

  $hashed = @()
  $errors = @()
  $i = 0

  foreach ($r in $candidates) {
    $i++
    Write-Progress -Activity 'Calculating SHA-256' -Status ('{0} / {1}: {2}' -f $i, $candidates.Count, $r.FileName) -PercentComplete (Percent $i $candidates.Count)
    Write-Host ('[hash] {0}/{1}: {2}' -f $i, $candidates.Count, $r.FileName)

    try {
      $h = Get-FileHash -LiteralPath $r.FullPath -Algorithm SHA256 -ErrorAction Stop
      $hashValue = [string]$h.Hash
      $exactKey = ('{0}|{1}' -f ([string]$r.SizeBytes), $hashValue)

      $hashed += [pscustomobject]@{
        FileId = [int]$r.FileId
        SourceFolderIndex = [int]$r.SourceFolderIndex
        SourceFolder = [string]$r.SourceFolder
        RelativePath = [string]$r.RelativePath
        FileName = [string]$r.FileName
        FullPath = [string]$r.FullPath
        SizeBytes = [int64]$r.SizeBytes
        SHA256 = $hashValue
        ExactKey = $exactKey
      }
    } catch {
      $errors += [pscustomobject]@{
        FileId = $r.FileId
        FileName = $r.FileName
        FullPath = $r.FullPath
        Error = $_.Exception.Message
      }
    }
  }

  Write-Progress -Activity 'Calculating SHA-256' -Completed

  if ($errors.Count -gt 0) {
    $errors | Export-Csv -LiteralPath $hashErrorPath -NoTypeInformation -Encoding UTF8
  }

  $hashBuckets = @{}
  foreach ($h in $hashed) {
    $key = [string]$h.ExactKey
    if (-not $hashBuckets.ContainsKey($key)) {
      $hashBuckets[$key] = @()
    }
    $hashBuckets[$key] = @($hashBuckets[$key]) + @($h)
  }

  $exactRows = @()
  $pairRows = @()
  $pairSet = @{}
  $groupId = 0

  foreach ($key in $hashBuckets.Keys) {
    $members = @($hashBuckets[$key])
    if ($members.Count -le 1) { continue }

    $groupId++

    foreach ($m in $members) {
      $exactRows += [pscustomobject]@{
        DuplicateGroupId = $groupId
        FileId = $m.FileId
        SourceFolderIndex = $m.SourceFolderIndex
        SourceFolder = $m.SourceFolder
        RelativePath = $m.RelativePath
        FileName = $m.FileName
        FullPath = $m.FullPath
        SizeBytes = $m.SizeBytes
        SHA256 = $m.SHA256
      }
    }

    for ($a = 0; $a -lt $members.Count; $a++) {
      for ($b = $a + 1; $b -lt $members.Count; $b++) {
        $am = $members[$a]
        $bm = $members[$b]
        $pairKey = Pair-Key $am.FileId $bm.FileId
        $pairSet[[string]$pairKey] = $true

        $pairRows += [pscustomobject]@{
          DuplicateGroupId = $groupId
          A_FileId = $am.FileId
          A_FileName = $am.FileName
          A_FullPath = $am.FullPath
          B_FileId = $bm.FileId
          B_FileName = $bm.FileName
          B_FullPath = $bm.FullPath
          SizeBytes = $am.SizeBytes
          SHA256 = $am.SHA256
        }
      }
    }
  }

  if ($exactRows.Count -gt 0) {
    $exactRows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
  } else {
    [pscustomobject]@{ Result='No exact duplicates found'; Rule='Same SizeBytes and SHA-256 hash' } | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
  }

  return [pscustomobject]@{
    GroupCount = [int]$groupId
    RowCount = [int]$exactRows.Count
    PairSet = $pairSet
    PairRows = @($pairRows)
    ErrorCount = [int]$errors.Count
  }
}

function Hamming([string]$a, [string]$b) {
  $len = $a.Length
  if ($b.Length -lt $len) { $len = $b.Length }

  $d = 0
  for ($i=0; $i -lt $len; $i++) {
    if ($a[$i] -ne $b[$i]) { $d++ }
  }

  return $d
}

function FrameHashes([byte[]]$bytes, [int]$frameSize) {
  $list = @()
  $count = [Math]::Floor($bytes.Length / $frameSize)

  for ($f=0; $f -lt $count; $f++) {
    $off = $f * $frameSize
    $sum = 0

    for ($i=0; $i -lt $frameSize; $i++) {
      $sum += [int]$bytes[$off + $i]
    }

    $avg = $sum / $frameSize
    $chars = New-Object char[] $frameSize

    for ($i=0; $i -lt $frameSize; $i++) {
      if ([int]$bytes[$off + $i] -ge $avg) { $chars[$i] = '1' } else { $chars[$i] = '0' }
    }

    $list += (-join $chars)
  }

  return @($list)
}

function VisualHash($ffmpeg, $record, [int]$index, [int]$total) {
  Write-Progress -Activity 'Building visual fingerprints' -Status ('{0} / {1}: {2}' -f $index, $total, $record.FileName) -PercentComplete (Percent $index $total)
  Write-Host ('[visual hash] {0}/{1}: {2}' -f $index, $total, $record.FileName)

  $raw = Join-Path $env:TEMP ('mp4_vhash_' + [Guid]::NewGuid().ToString('N') + '.raw')

  try {
    $vf = 'fps={0},scale={1}:{1}:flags=bicubic,format=gray' -f $SampleFps, $HashSize
    & $ffmpeg -hide_banner -loglevel error -y -i $record.FullPath -vf $vf -an -sn -dn -f rawvideo $raw

    if (-not (Test-Path -LiteralPath $raw)) { throw 'ffmpeg did not produce a raw fingerprint file.' }

    $bytes = [System.IO.File]::ReadAllBytes($raw)
    if ($bytes.Length -lt $BitsPerFrame) { throw 'No usable frames were extracted.' }

    $hashes = FrameHashes $bytes $BitsPerFrame

    return [pscustomobject]@{
      FileId = [int]$record.FileId
      SourceFolderIndex = [int]$record.SourceFolderIndex
      FileName = [string]$record.FileName
      FullPath = [string]$record.FullPath
      SizeBytes = [int64]$record.SizeBytes
      FrameCount = [int]$hashes.Count
      Fingerprints = @($hashes)
      Error = ''
    }
  } catch {
    return [pscustomobject]@{
      FileId = [int]$record.FileId
      SourceFolderIndex = [int]$record.SourceFolderIndex
      FileName = [string]$record.FileName
      FullPath = [string]$record.FullPath
      SizeBytes = [int64]$record.SizeBytes
      FrameCount = 0
      Fingerprints = @()
      Error = $_.Exception.Message
    }
  } finally {
    Remove-Item -LiteralPath $raw -Force -ErrorAction SilentlyContinue
  }
}

function Compare-VHash($a, $b) {
  $minFrames = [int]$a.FrameCount
  if ([int]$b.FrameCount -lt $minFrames) { $minFrames = [int]$b.FrameCount }

  $maxFrames = [int]$a.FrameCount
  if ([int]$b.FrameCount -gt $maxFrames) { $maxFrames = [int]$b.FrameCount }

  if ($minFrames -le 0) { return $null }

  $totalDistance = 0
  for ($i=0; $i -lt $minFrames; $i++) {
    $totalDistance += Hamming ([string]$a.Fingerprints[$i]) ([string]$b.Fingerprints[$i])
  }

  $avgDist = $totalDistance / $minFrames
  $visual = 1 - ($avgDist / $BitsPerFrame)
  $ratio = $minFrames / $maxFrames
  $weighted = $visual * $ratio

  return [pscustomobject]@{
    A_FileId = [int]$a.FileId
    A_FileName = [string]$a.FileName
    A_FullPath = [string]$a.FullPath
    A_SourceFolderIndex = [int]$a.SourceFolderIndex
    A_FrameSamples = [int]$a.FrameCount
    A_SizeBytes = [int64]$a.SizeBytes
    B_FileId = [int]$b.FileId
    B_FileName = [string]$b.FileName
    B_FullPath = [string]$b.FullPath
    B_SourceFolderIndex = [int]$b.SourceFolderIndex
    B_FrameSamples = [int]$b.FrameCount
    B_SizeBytes = [int64]$b.SizeBytes
    ComparedFrames = [int]$minFrames
    AverageHammingDistancePerFrame = [Math]::Round($avgDist,4)
    VisualSimilarity = [Math]::Round($visual,6)
    FrameCountRatio = [Math]::Round($ratio,6)
    WeightedSimilarity = [Math]::Round($weighted,6)
    LikelySameVisualVideo = ($weighted -ge $LikelyThreshold)
  }
}

function Run-Visual($files, $exactPairs, [string]$csvPath, [string]$errPath) {
  Write-Host ''
  Write-Host 'Stage 2: visual fingerprint check'

  $ffmpeg = Resolve-Ffmpeg
  Write-Host ('ffmpeg: {0}' -f $ffmpeg)

  $fps = @()
  $n = 0

  foreach ($r in $files) {
    $n++
    $fps += (VisualHash $ffmpeg $r $n $files.Count)
  }

  Write-Progress -Activity 'Building visual fingerprints' -Completed

  $errs = @($fps | Where-Object { $_.Error -ne '' })
  if ($errs.Count -gt 0) {
    $errs | Select-Object FileId,FileName,FullPath,Error | Export-Csv -LiteralPath $errPath -NoTypeInformation -Encoding UTF8
  }

  $valid = @($fps | Where-Object { $_.Error -eq '' -and $_.FrameCount -gt 0 })

  $rawPairs = 0
  if ($valid.Count -ge 2) { $rawPairs = [int64](($valid.Count * ($valid.Count - 1)) / 2) }

  Write-Host ('Visual eligible files: {0}' -f $valid.Count)
  Write-Host ('Raw visual pair count: {0}' -f $rawPairs)
  Write-Host 'Comparing visual fingerprints...'

  $rows = @()
  $pairIndex = 0
  $skipExact = 0

  for ($i=0; $i -lt $valid.Count; $i++) {
    for ($j=$i+1; $j -lt $valid.Count; $j++) {
      $pairIndex++

      if (($pairIndex % 100) -eq 0 -or $pairIndex -eq 1 -or $pairIndex -eq $rawPairs) {
        Write-Progress -Activity 'Comparing visual fingerprints' -Status ('{0} / {1} pairs' -f $pairIndex, $rawPairs) -PercentComplete (Percent $pairIndex $rawPairs)
      }

      $a = $valid[$i]
      $b = $valid[$j]
      $key = Pair-Key $a.FileId $b.FileId

      if ($exactPairs.ContainsKey([string]$key)) {
        $skipExact++
        continue
      }

      $cmp = Compare-VHash $a $b
      if ($null -ne $cmp) { $rows += $cmp }
    }
  }

  Write-Progress -Activity 'Comparing visual fingerprints' -Completed

  $sorted = @($rows | Sort-Object WeightedSimilarity -Descending)

  if ($sorted.Count -gt 0) {
    $sorted | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
  } else {
    [pscustomobject]@{ Result='No visual comparison rows generated' } | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
  }

  $likely = @($sorted | Where-Object { $_.LikelySameVisualVideo -eq $true })
  $top = @($sorted | Select-Object -First $TopLimit)

  return [pscustomobject]@{
    Rows = @($sorted)
    Likely = @($likely)
    Top = @($top)
    ErrorCount = [int]$errs.Count
    RawPairCount = [int64]$rawPairs
    SkippedExactPairs = [int]$skipExact
    Ffmpeg = $ffmpeg
  }
}

try {
  Write-Host ''
  Write-Host ('MP4 Duplicate Checker v{0}' -f $ToolVersion)
  Write-Host ''
  Write-Host 'Checks exact duplicates first, then optional visual duplicate candidates.'
  Write-Host 'This tool does not rename, move, delete, or edit any video files.'
  Write-Host ''
  Write-Host 'Input note: paste folder paths during the interactive prompts.'
  Write-Host 'Alternative: before the tool starts, you may drag one or more folders onto the BAT file icon.'
  Write-Host ''

  $folders = New-Object System.Collections.ArrayList
  $folderSeen = @{}

  foreach ($p in $InitialFolders) {
    if (-not [string]::IsNullOrWhiteSpace($p)) { [void](Add-Folder $folders $folderSeen $p) }
  }

  if ($folders.Count -eq 0) {
    [void](Add-Folder $folders $folderSeen (Ask-Folder 'Folder 1'))
  }

  while ($true) {
    Show-Folders $folders
    Write-Host 'Choose: A = add another folder, S = start checking, Q = quit'
    $choice = (Read-Host 'Your choice').Trim().ToUpperInvariant()

    if ($choice -eq 'A') {
      [void](Add-Folder $folders $folderSeen (Ask-Folder ('Folder {0}' -f ($folders.Count + 1))))
    } elseif ($choice -eq 'S') {
      break
    } elseif ($choice -eq 'Q') {
      Write-Host 'Cancelled.'
      Pause-End
      exit 0
    } else {
      Write-Host 'Unknown choice.'
    }
  }

  $recurseText = (Read-Host 'Scan subfolders too? [y/N]').Trim().ToUpperInvariant()
  $recursive = ($recurseText -eq 'Y' -or $recurseText -eq 'YES')

  $desktop = [Environment]::GetFolderPath('Desktop')
  $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $reportDir = Join-Path $desktop ('mp4_duplicate_checker_report_{0}' -f $stamp)
  $script:ReportDirForErrors = $reportDir
  [void](New-Item -ItemType Directory -Path $reportDir -Force)

  $summaryPath = Join-Path $reportDir ('duplicate_summary_{0}.txt' -f $stamp)
  $exactPath = Join-Path $reportDir ('exact_duplicate_report_{0}.csv' -f $stamp)
  $visualPath = Join-Path $reportDir ('visual_similarity_report_{0}.csv' -f $stamp)
  $visualErrPath = Join-Path $reportDir ('visual_processing_errors_{0}.csv' -f $stamp)
  $hashErrPath = Join-Path $reportDir ('hash_processing_errors_{0}.csv' -f $stamp)

  Write-Host ''
  Write-Host 'Scanning MP4 files...'
  $files = @(Collect-Mp4 $folders $recursive)

  Write-Host ('Total MP4 files found: {0}' -f $files.Count)
  if ($files.Count -lt 2) { throw 'At least two MP4 files are required.' }

  $maxPairs = [int64](($files.Count * ($files.Count - 1)) / 2)

  Write-Host ('Maximum visual pair count: {0}' -f $maxPairs)
  Write-Host ('Report folder: {0}' -f $reportDir)

  $exact = Run-Exact $files $exactPath $hashErrPath

  Write-Host ('Exact duplicate groups: {0}' -f $exact.GroupCount)
  Write-Host ('Exact duplicate pairs: {0}' -f $exact.PairRows.Count)

  $runVisualText = (Read-Host 'Run visual fingerprint check with FFmpeg? [Y/n]').Trim().ToUpperInvariant()
  $runVisual = -not ($runVisualText -eq 'N' -or $runVisualText -eq 'NO')

  $visual = $null
  $visualError = ''

  if ($runVisual) {
    try {
      $visual = Run-Visual $files $exact.PairSet $visualPath $visualErrPath
      Write-Host ('Likely visual duplicate pairs: {0}' -f $visual.Likely.Count)
    } catch {
      $visualError = $_.Exception.Message
      [pscustomobject]@{ Result='Visual check failed'; Error=$visualError } | Export-Csv -LiteralPath $visualPath -NoTypeInformation -Encoding UTF8
      Write-Host ('Visual check failed: {0}' -f $visualError)
    }
  } else {
    [pscustomobject]@{ Result='Visual check skipped by user' } | Export-Csv -LiteralPath $visualPath -NoTypeInformation -Encoding UTF8
  }

  $visualLikelyCount = 0
  if ($runVisual -and $null -ne $visual) { $visualLikelyCount = $visual.Likely.Count }

  $totalPossibleDupes = $exact.PairRows.Count + $visualLikelyCount

  $lines = New-Object System.Collections.Generic.List[string]
  [void]$lines.Add(('MP4 Duplicate Checker v{0} Summary' -f $ToolVersion))
  [void]$lines.Add(('Generated: {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
  [void]$lines.Add('')
  [void]$lines.Add('==============================')
  [void]$lines.Add('RESULT OVERVIEW')
  [void]$lines.Add('==============================')
  [void]$lines.Add(('Possible duplicate pairs total: {0}' -f $totalPossibleDupes))
  [void]$lines.Add(('Exact duplicate pairs by hash: {0}' -f $exact.PairRows.Count))
  [void]$lines.Add(('Visual duplicate candidate pairs by FFmpeg frame fingerprint: {0}' -f $visualLikelyCount))
  [void]$lines.Add('')
  [void]$lines.Add('Exact duplicate pairs:')

  if ($exact.PairRows.Count -gt 0) {
    $pairNumber = 0
    foreach ($p in $exact.PairRows) {
      $pairNumber++
      [void]$lines.Add(('[EXACT {0}]' -f $pairNumber))
      [void]$lines.Add(('A: {0}' -f $p.A_FullPath))
      [void]$lines.Add(('B: {0}' -f $p.B_FullPath))
      [void]$lines.Add('')
    }
  } else {
    [void]$lines.Add('None.')
    [void]$lines.Add('')
  }

  [void]$lines.Add('Visual duplicate candidate pairs:')

  if ($visualLikelyCount -gt 0) {
    $pairNumber = 0
    foreach ($p in $visual.Likely) {
      $pairNumber++
      [void]$lines.Add(('[VISUAL {0}] Score: {1}' -f $pairNumber, $p.WeightedSimilarity))
      [void]$lines.Add(('A: {0}' -f $p.A_FullPath))
      [void]$lines.Add(('B: {0}' -f $p.B_FullPath))
      [void]$lines.Add('')
    }
  } else {
    [void]$lines.Add('None.')
    [void]$lines.Add('')
  }

  [void]$lines.Add('==============================')
  [void]$lines.Add('DETAILS')
  [void]$lines.Add('==============================')
  [void]$lines.Add('')
  [void]$lines.Add('Folders:')

  $idx = 0
  foreach ($f in $folders) {
    $idx++
    [void]$lines.Add(('  {0}. {1}' -f $idx, $f))
  }

  [void]$lines.Add('')
  [void]$lines.Add(('Scan subfolders: {0}' -f $recursive))
  [void]$lines.Add(('Total MP4 files: {0}' -f $files.Count))
  [void]$lines.Add(('Maximum visual pair count: {0}' -f $maxPairs))
  [void]$lines.Add('')
  [void]$lines.Add('Exact duplicate rule: same file size and same SHA-256 hash.')
  [void]$lines.Add(('Exact duplicate groups: {0}' -f $exact.GroupCount))
  [void]$lines.Add(('Exact duplicate file rows: {0}' -f $exact.RowCount))
  [void]$lines.Add(('Exact duplicate report: {0}' -f $exactPath))

  if ($exact.ErrorCount -gt 0) {
    [void]$lines.Add(('Hash processing errors: {0}' -f $hashErrPath))
  }

  [void]$lines.Add('')

  if ($runVisual -and $null -ne $visual) {
    [void]$lines.Add('Visual method: FFmpeg samples 1 frame per second, converts frames to 8x8 grayscale hashes, then compares pairs.')
    [void]$lines.Add(('Likely threshold: WeightedSimilarity >= {0}' -f $LikelyThreshold))
    [void]$lines.Add(('Raw visual pair count: {0}' -f $visual.RawPairCount))
    [void]$lines.Add(('Exact pairs skipped in visual report: {0}' -f $visual.SkippedExactPairs))
    [void]$lines.Add(('Likely visual duplicate pairs: {0}' -f $visual.Likely.Count))
    [void]$lines.Add(('Visual similarity report: {0}' -f $visualPath))

    if ($visual.ErrorCount -gt 0) {
      [void]$lines.Add(('Visual processing errors: {0}' -f $visualErrPath))
    }

    [void]$lines.Add('')
    [void]$lines.Add(('Top {0} closest visual pairs:' -f $TopLimit))

    foreach ($r in $visual.Top) {
      [void]$lines.Add('')
      [void]$lines.Add(('Score: {0}, VisualSimilarity: {1}, FrameCountRatio: {2}' -f $r.WeightedSimilarity, $r.VisualSimilarity, $r.FrameCountRatio))
      [void]$lines.Add(('A: {0}' -f $r.A_FullPath))
      [void]$lines.Add(('B: {0}' -f $r.B_FullPath))
    }
  } elseif ($runVisual -and $visualError -ne '') {
    [void]$lines.Add(('Visual check failed: {0}' -f $visualError))
    [void]$lines.Add(('Visual report: {0}' -f $visualPath))
  } else {
    [void]$lines.Add('Visual check skipped by user.')
  }

  [void]$lines.Add('')
  [void]$lines.Add('Interpretation:')
  [void]$lines.Add('Exact duplicate pairs are byte-level matches.')
  [void]$lines.Add('Visual duplicate candidates require manual review before deletion or archival decisions.')
  [void]$lines.Add('This tool does not modify source files.')

  $lines | Set-Content -LiteralPath $summaryPath -Encoding UTF8

  Write-Host ''
  Write-Host 'Done.'
  Write-Host ('Possible duplicate pairs total: {0}' -f $totalPossibleDupes)
  Write-Host ('Exact duplicate pairs by hash: {0}' -f $exact.PairRows.Count)
  Write-Host ('Visual duplicate candidate pairs by FFmpeg frame fingerprint: {0}' -f $visualLikelyCount)
  Write-Host ''
  Write-Host ('Summary: {0}' -f $summaryPath)
  Write-Host ('Exact duplicate report: {0}' -f $exactPath)
  Write-Host ('Visual similarity report: {0}' -f $visualPath)
  Write-Host ('Report folder: {0}' -f $reportDir)

  Pause-End
} catch {
  Write-Host ''
  Write-Host ('ERROR: {0}' -f $_.Exception.Message)

  try {
    if (-not [string]::IsNullOrWhiteSpace($script:ReportDirForErrors)) {
      $errStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
      $errPath = Join-Path $script:ReportDirForErrors ('error_debug_{0}.txt' -f $errStamp)
      $errLines = @()
      $errLines += ('MP4 Duplicate Checker v{0} Error Debug' -f $ToolVersion)
      $errLines += ('Generated: {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
      $errLines += ''
      $errLines += ('Error: {0}' -f $_.Exception.Message)
      $errLines += ''
      $errLines += 'Script stack trace:'
      $errLines += $_.ScriptStackTrace
      $errLines | Set-Content -LiteralPath $errPath -Encoding UTF8
      Write-Host ('Error debug file: {0}' -f $errPath)
    }
  } catch {
  }

  Pause-End
}
