param(
  [string]$VaultPath,
  [string]$Notebook,
  [string]$Section,
  [switch]$Force,
  [switch]$List
)

$ErrorActionPreference = "Stop"

if (-not $VaultPath) {
  if (Test-Path -LiteralPath (Join-Path (Get-Location).Path ".obsidian")) {
    $VaultPath = (Get-Location).Path
  } elseif (Test-Path -LiteralPath (Join-Path $PSScriptRoot "..\..\.obsidian")) {
    $VaultPath = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
  } elseif (Test-Path -LiteralPath (Join-Path $PSScriptRoot "..\..\..\.obsidian")) {
    $VaultPath = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
  } else {
    $VaultPath = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
  }
}

$vault = $VaultPath
$output = Join-Path $vault "OneNote"
$assetsDir = Join-Path $output "Assets"
$manifestPath = Join-Path $PSScriptRoot ".onenote-sync-manifest.json"

New-Item -ItemType Directory -Path $output -Force | Out-Null
New-Item -ItemType Directory -Path $assetsDir -Force | Out-Null

try {
  $oneNote = New-Object -ComObject OneNote.Application
} catch {
  throw "Die Microsoft OneNote-Desktopanwendung konnte nicht angesprochen werden. Bitte stelle sicher, dass OneNote geöffnet oder installiert ist."
}

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$hierarchyXml = ""
try {
  $oneNote.GetHierarchy("", 4, [ref]$hierarchyXml)
} catch {
  throw "Fehler beim Abrufen der OneNote-Hierarchie: $($_.Exception.Message)"
}

if ([string]::IsNullOrWhiteSpace($hierarchyXml)) {
  throw "OneNote hat eine leere Hierarchie zurückgegeben."
}

[xml]$hierarchyDoc = $hierarchyXml
$ns = New-Object System.Xml.XmlNamespaceManager($hierarchyDoc.NameTable)
$ns.AddNamespace("one", $hierarchyDoc.DocumentElement.NamespaceURI)

function Get-RelativeFolderPath($pageNode) {
  $segments = [System.Collections.Generic.List[string]]::new()
  $curr = $pageNode.ParentNode
  while ($curr -and $curr.LocalName -ne "Notebooks" -and $curr.LocalName -ne "#document") {
    if ($curr.name) {
      $safeSegment = ($curr.name -replace '[<>:"/\\|?*]', "_").Trim()
      if ($safeSegment) {
        $segments.Add($safeSegment)
      }
    }
    $curr = $curr.ParentNode
  }
  $segments.Reverse()
  return ($segments -join "\")
}

function Get-SafeFileName($title) {
  $safe = ($title -replace '[<>:"/\\|?*]', "_").Trim()
  if ([string]::IsNullOrWhiteSpace($safe)) {
    $safe = "Seite ohne Titel"
  }
  return $safe
}

function Clean-HtmlText([string]$html) {
  if ([string]::IsNullOrWhiteSpace($html)) { return "" }
  $s = [System.Net.WebUtility]::HtmlDecode($html)
  $s = $s -replace "(?i)<br\s*/?>", "`n"
  $s = $s -replace "(?i)<span\s+style=['""][^'""]*font-weight\s*:\s*bold[^'""]*['""]>(.*?)</span>", '**$1**'
  $s = $s -replace "(?i)<span\s+style=['""][^'""]*font-style\s*:\s*italic[^'""]*['""]>(.*?)</span>", '*$1*'
  $s = $s -replace "(?i)<b>(.*?)</b>", '**$1**'
  $s = $s -replace "(?i)<i>(.*?)</i>", '*$1*'
  $s = $s -replace "(?s)<[^>]+>", ""
  return $s.Trim()
}

function Convert-PageXmlToMarkdown {
  param(
    [xml]$PageDoc,
    [System.Xml.XmlNamespaceManager]$NsManager,
    [string]$TargetAssetsDir,
    [string]$BaseFileName,
    [string]$NotebookName,
    [string]$SectionName,
    [string]$TargetFolder
  )

  $title = $PageDoc.DocumentElement.name
  $lines = [System.Collections.Generic.List[string]]::new()
  $lines.Add("# $title")
  $lines.Add("")
  $lines.Add("> [!info] Importierter OneNote-Inhalt")
  $lines.Add("> Notizbuch: $NotebookName | Abschnitt: $SectionName")
  $lines.Add("")

  $imgIndex = 0

  foreach ($outline in $PageDoc.SelectNodes("//one:Outline", $NsManager)) {
    $tables = $outline.SelectNodes(".//one:Table", $NsManager)
    if ($tables -and $tables.Count -gt 0) {
      foreach ($table in $tables) {
        $isFirst = $true
        foreach ($row in $table.SelectNodes("one:Row", $NsManager)) {
          $cells = foreach ($cell in $row.SelectNodes("one:Cell", $NsManager)) {
            $texts = foreach ($t in $cell.SelectNodes(".//one:T", $NsManager)) {
              Clean-HtmlText $t.InnerText
            }
            (($texts | Where-Object { $_ }) -join " ") -replace "\|", "&#124;"
          }
          $lines.Add("| " + ($cells -join " | ") + " |")
          if ($isFirst) {
            $seps = foreach ($c in $cells) { "---" }
            $lines.Add("| " + ($seps -join " | ") + " |")
            $isFirst = $false
          }
        }
        $lines.Add("")
      }
    }

    foreach ($oe in $outline.SelectNodes(".//one:OE", $NsManager)) {
      if ($oe.SelectSingleNode("ancestor::one:Cell", $NsManager)) {
        continue
      }

      $depth = 0
      $curr = $oe.ParentNode
      while ($curr -and $curr.LocalName -ne "Outline") {
        if ($curr.LocalName -eq "OEChildren") { $depth++ }
        $curr = $curr.ParentNode
      }
      $indent = "  " * [Math]::Max(0, $depth - 1)

      $listNode = $oe.SelectSingleNode("one:List", $NsManager)
      $prefix = ""
      if ($listNode) {
        if ($listNode.SelectSingleNode("one:Number", $NsManager)) {
          $prefix = "1. "
        } else {
          $prefix = "- "
        }
      }

      $tNode = $oe.SelectSingleNode("one:T", $NsManager)
      if ($tNode) {
        $text = Clean-HtmlText $tNode.InnerText
        if ($text -and $text -ne $title) {
          if ($prefix) {
            $lines.Add("$indent$prefix$text")
          } else {
            $lines.Add("$indent$text")
          }
        }
      }

      $imgNode = $oe.SelectSingleNode("one:Image", $NsManager)
      if ($imgNode) {
        $imgIndex++
        $dataNode = $imgNode.SelectSingleNode("one:Data", $NsManager)
        if ($dataNode -and -not [string]::IsNullOrWhiteSpace($dataNode.InnerText)) {
          $ext = if ($imgNode.format) { $imgNode.format.ToLower() } else { "png" }
          if ($ext -eq "emf") { $ext = "png" }
          $imgFileName = "${BaseFileName}_image_$($imgIndex.ToString('0000')).$ext"
          $imgPath = Join-Path $TargetAssetsDir $imgFileName
          try {
            $bytes = [Convert]::FromBase64String($dataNode.InnerText.Trim())
            [System.IO.File]::WriteAllBytes($imgPath, $bytes)
            
            $relPath = if ($TargetFolder -and (Test-Path -LiteralPath $TargetFolder)) {
              $fromUri = [System.Uri]((Resolve-Path -LiteralPath $TargetFolder).Path + "\")
              $toUri = [System.Uri]$imgPath
              [System.Uri]::UnescapeDataString($fromUri.MakeRelativeUri($toUri).ToString())
            } else {
              "../../Assets/$imgFileName"
            }
            $lines.Add("![image]($relPath)")
          } catch {}
        }
      }
    }
    $lines.Add("")
  }

  return ($lines -join "`r`n").Trim() + "`r`n"
}

$manifest = [ordered]@{}
if (Test-Path -LiteralPath $manifestPath) {
  try {
    $rawJson = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    if ($rawJson -and $rawJson.pages) {
      foreach ($prop in $rawJson.pages.PSObject.Properties) {
        $manifest[$prop.Name] = [ordered]@{
          lastModifiedTime = $prop.Value.lastModifiedTime
          path = $prop.Value.path
          title = $prop.Value.title
        }
      }
    }
  } catch {
    $manifest = [ordered]@{}
  }
}

$allPageNodes = $hierarchyDoc.SelectNodes("//one:Page", $ns)
$filteredPages = [System.Collections.Generic.List[System.Xml.XmlNode]]::new()

foreach ($page in $allPageNodes) {
  if ($Notebook) {
    $nb = $page.SelectSingleNode("ancestor::one:Notebook", $ns)
    if (-not $nb -or $nb.name -ne $Notebook) { continue }
  }
  if ($Section) {
    $sec = $page.SelectSingleNode("ancestor::one:Section", $ns)
    if (-not $sec -or $sec.name -ne $Section) { continue }
  }
  $filteredPages.Add($page)
}

if ($List) {
  $listItems = foreach ($page in $filteredPages) {
    $rel = Get-RelativeFolderPath $page
    $safe = Get-SafeFileName $page.name
    $filePath = Join-Path $rel "$safe.md"
    $fullPath = Join-Path $output $filePath
    $status = if (-not (Test-Path -LiteralPath $fullPath)) { "Neu" } else { "Vorhanden" }
    [PSCustomObject]@{
      Titel = $page.name
      Abschnitt = $rel
      Geaendert = $page.lastModifiedTime
      Status = $status
    }
  }
  $listItems | Format-Table -AutoSize
  exit 0
}

$toSync = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($page in $filteredPages) {
  $relFolder = Get-RelativeFolderPath $page
  $safeTitle = Get-SafeFileName $page.name
  $relFilePath = Join-Path $relFolder "$safeTitle.md"
  $absoluteFilePath = Join-Path $output $relFilePath
  $pageMod = [datetime]$page.lastModifiedTime

  $needsUpdate = $false
  if ($Force) {
    $needsUpdate = $true
  } elseif (-not (Test-Path -LiteralPath $absoluteFilePath)) {
    $needsUpdate = $true
  } elseif ($manifest.Contains($page.ID)) {
    $recordedMod = [datetime]$manifest[$page.ID].lastModifiedTime
    if ($pageMod -gt $recordedMod.AddSeconds(2)) {
      $needsUpdate = $true
    }
  } else {
    $fileInfo = Get-Item -LiteralPath $absoluteFilePath
    if ($pageMod -gt $fileInfo.LastWriteTime.ToUniversalTime().AddSeconds(5)) {
      $needsUpdate = $true
    } else {
      $manifest[$page.ID] = [ordered]@{
        lastModifiedTime = $page.lastModifiedTime
        path = $relFilePath
        title = $page.name
      }
    }
  }

  if ($needsUpdate) {
    $toSync.Add([PSCustomObject]@{
      Page = $page
      RelPath = $relFilePath
      AbsolutePath = $absoluteFilePath
    })
  }
}

if ($toSync.Count -eq 0) {
  $manifestData = [ordered]@{
    version = 1
    lastSync = (Get-Date).ToUniversalTime().ToString("o")
    pages = $manifest
  }
  $manifestJson = $manifestData | ConvertTo-Json -Depth 5
  [System.IO.File]::WriteAllText($manifestPath, $manifestJson, [System.Text.UTF8Encoding]::new($false))

  $sw.Stop()
  $sec = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
  Write-Host "OneNote ist aktuell. Keine Aenderungen gefunden ($($filteredPages.Count) Seiten in $($sec)s geprueft)."
  exit 0
}

Write-Host "OneNote: $($toSync.Count) von $($filteredPages.Count) Seite(n) haben Neuerungen. Synchronisiere..."

$successCount = 0
foreach ($item in $toSync) {
  $page = $item.Page
  Write-Host "  -> Importiere: $($page.name)"
  try {
    $pxml = ""
    $oneNote.GetPageContent($page.ID, [ref]$pxml, 7)
    [xml]$pdoc = $pxml
    $pns = New-Object System.Xml.XmlNamespaceManager($pdoc.NameTable)
    $pns.AddNamespace("one", $pdoc.DocumentElement.NamespaceURI)

    $nbNode = $page.SelectSingleNode("ancestor::one:Notebook", $ns)
    $secNode = $page.SelectSingleNode("ancestor::one:Section", $ns)
    $nbName = if ($nbNode) { $nbNode.name } else { "" }
    $secName = if ($secNode) { $secNode.name } else { "" }

    $targetFolder = Split-Path $item.AbsolutePath -Parent
    if (-not (Test-Path -LiteralPath $targetFolder)) {
      New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
    }

    $mdContent = Convert-PageXmlToMarkdown -PageDoc $pdoc -NsManager $pns `
      -TargetAssetsDir $assetsDir -BaseFileName (Get-SafeFileName $page.name) `
      -NotebookName $nbName -SectionName $secName -TargetFolder $targetFolder

    [System.IO.File]::WriteAllText($item.AbsolutePath, $mdContent, [System.Text.UTF8Encoding]::new($false))

    $manifest[$page.ID] = [ordered]@{
      lastModifiedTime = $page.lastModifiedTime
      path = $item.RelPath
      title = $page.name
    }
    $successCount++
  } catch {
    Write-Warning "Fehler beim Importieren von $($page.name): $($_.Exception.Message)"
  }
}

$manifestData = [ordered]@{
  version = 1
  lastSync = (Get-Date).ToUniversalTime().ToString("o")
  pages = $manifest
}
$manifestJson = $manifestData | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($manifestPath, $manifestJson, [System.Text.UTF8Encoding]::new($false))

$sw.Stop()
$sec = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
Write-Host "Fertig: $successCount Seite(n) erfolgreich synchronisiert (Dauer: $($sec)s)."
exit 0
