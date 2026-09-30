param(
  [string]$VaultPath,
  [string]$Notebook,
  [string]$Section,
  [switch]$Force,
  [switch]$List,
  [switch]$CleanDuplicates,
  [switch]$WithCanvas,
  [switch]$SkipCanvas,
  [switch]$SkipInkImages
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
  throw "Failed to connect to Microsoft OneNote desktop application. Please ensure OneNote is installed and running."
}

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$hierarchyXml = ""
try {
  $oneNote.GetHierarchy("", 4, [ref]$hierarchyXml)
} catch {
  throw "Error retrieving OneNote hierarchy: $($_.Exception.Message)"
}

if ([string]::IsNullOrWhiteSpace($hierarchyXml)) {
  throw "OneNote returned an empty hierarchy."
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

function Get-SubPageFolderPrefix {
  param([string]$SectionKey, [int]$Depth)
  if (-not $script:sectionCursors.ContainsKey($SectionKey)) {
    $script:sectionCursors[$SectionKey] = @("", "", "")
  }
  $cursors = $script:sectionCursors[$SectionKey]
  $parts = @()
  for ($i = 0; $i -lt $Depth; $i++) {
    if ($i -lt $cursors.Count -and $cursors[$i]) { $parts += $cursors[$i] }
  }
  return $parts
}

function Set-SubPageCursor {
  param([string]$SectionKey, [int]$Depth, [string]$Name)
  if (-not $script:sectionCursors.ContainsKey($SectionKey)) {
    $script:sectionCursors[$SectionKey] = @("", "", "")
  }
  $cursors = $script:sectionCursors[$SectionKey]
  if ($Depth -lt $cursors.Count) {
    $cursors[$Depth] = $Name
    for ($i = $Depth + 1; $i -lt $cursors.Count; $i++) { $cursors[$i] = "" }
  }
}

$script:sectionCursors = @{}

function Get-PageRelativePath {
  param($pageNode, [string]$safeTitle)
  $secNode = $pageNode.SelectSingleNode("ancestor::one:Section", $ns)
  $sectionKey = if ($secNode) { $secNode.ID } else { "" }

  $depth = 0
  if ($pageNode.pageLevel) {
    $parsed = 0
    if ([int]::TryParse($pageNode.pageLevel, [ref]$parsed) -and $parsed -gt 1) {
      $depth = [Math]::Min($parsed - 1, 2)
    }
  }

  $baseRelFolder = Get-RelativeFolderPath $pageNode
  $folderParts = [System.Collections.Generic.List[string]]::new()
  if ($baseRelFolder) { $folderParts.Add($baseRelFolder) }

  foreach ($prefixPart in (Get-SubPageFolderPrefix -SectionKey $sectionKey -Depth $depth)) {
    $folderParts.Add($prefixPart)
  }

  Set-SubPageCursor -SectionKey $sectionKey -Depth $depth -Name $safeTitle

  $finalFolder = ($folderParts -join "\")
  $relFile = if ($finalFolder) { Join-Path $finalFolder "$safeTitle.md" } else { "$safeTitle.md" }
  return [PSCustomObject]@{
    RelFolder = $finalFolder
    RelFilePath = $relFile
  }
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

function New-HexId {
  $bytes = New-Object byte[] 8
  [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
  return [System.BitConverter]::ToString($bytes).Replace("-", "").ToLowerInvariant()
}

function Get-InkRecognizedText {
  param($Scope, [System.Xml.XmlNamespaceManager]$NsManager)
  $sb = [System.Text.StringBuilder]::new()

  $candidates = [System.Collections.Generic.List[object]]::new()
  if ($Scope.LocalName -eq "InkWord") { $candidates.Add($Scope) }
  foreach ($w in $Scope.SelectNodes(".//one:InkWord", $NsManager)) { $candidates.Add($w) }

  foreach ($word in $candidates) {
    $recognized = $word.GetAttribute("recognizedText")
    if (-not [string]::IsNullOrWhiteSpace($recognized)) {
      [void]$sb.Append($recognized)
      continue
    }
    if ($word.SelectSingleNode("one:Space", $NsManager)) {
      [void]$sb.Append(" ")
    } elseif ($word.SelectSingleNode("one:EndOfLine", $NsManager)) {
      [void]$sb.Append("`n")
    }
  }
  return $sb.ToString().Trim()
}

function Add-InkBlobsFrom {
  param(
    $Scope,
    [System.Xml.XmlNamespaceManager]$NsManager,
    [System.Collections.Generic.List[byte[]]]$Target
  )

  if ($Scope.LocalName -in @("InkDrawing", "InkWord")) {
    $selfData = $Scope.SelectSingleNode("one:Data", $NsManager)
    if ($selfData -and -not [string]::IsNullOrWhiteSpace($selfData.InnerText)) {
      try { $Target.Add([Convert]::FromBase64String($selfData.InnerText.Trim())) } catch { }
    }
  }

  foreach ($node in $Scope.SelectNodes(".//one:InkDrawing | .//one:InkWord", $NsManager)) {
    $dataNode = $node.SelectSingleNode("one:Data", $NsManager)
    if (-not $dataNode -or [string]::IsNullOrWhiteSpace($dataNode.InnerText)) { continue }
    try { $Target.Add([Convert]::FromBase64String($dataNode.InnerText.Trim())) } catch { continue }
  }
}

function New-InkStrokesPng {
  param([System.Collections.Generic.List[byte[]]]$InkBlobs)

  if (-not $InkBlobs -or $InkBlobs.Count -eq 0) { return $null }
  Add-Type -AssemblyName WindowsBase, PresentationCore, System.Drawing -ErrorAction Stop

  $ink = New-Object System.Windows.Ink.StrokeCollection
  foreach ($bytes in $InkBlobs) {
    $ms = New-Object System.IO.MemoryStream -ArgumentList @(,$bytes)
    try {
      $part = New-Object System.Windows.Ink.StrokeCollection -ArgumentList @($ms)
      if ($part.Count -gt 0) { $ink.Add($part) }
    } catch {
    } finally { $ms.Dispose() }
  }

  if ($ink.Count -eq 0) { return $null }
  $bounds = $ink.GetBounds()
  if ($bounds.Width -le 0 -or $bounds.Height -le 0) { return $null }

  $margin = 8.0
  $scale = 2.0
  $w = [int][Math]::Max(1, [Math]::Ceiling(($bounds.Width + $margin * 2) * $scale))
  $h = [int][Math]::Max(1, [Math]::Ceiling(($bounds.Height + $margin * 2) * $scale))
  if ($w -gt 20000 -or $h -gt 20000) { return $null }

  $visual = New-Object System.Windows.Media.DrawingVisual
  $dc = $visual.RenderOpen()
  try {
    $dc.DrawRectangle([System.Windows.Media.Brushes]::Transparent, $null, (New-Object System.Windows.Rect(0, 0, $w, $h)))
    $dc.PushTransform((New-Object System.Windows.Media.ScaleTransform($scale, $scale)))
    $dc.PushTransform((New-Object System.Windows.Media.TranslateTransform((-$bounds.Left + $margin), (-$bounds.Top + $margin))))
    $ink.Draw($dc)
    $dc.Pop(); $dc.Pop()
  } finally { $dc.Close() }

  $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap($w, $h, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
  $rtb.Render($visual)

  $encoder = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
  $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
  $outStream = New-Object System.IO.MemoryStream
  try {
    $encoder.Save($outStream)
    return , $outStream.ToArray()
  } finally { $outStream.Dispose() }
}

function Convert-PageXmlToMarkdown {
  param(
    [xml]$PageDoc,
    [System.Xml.XmlNamespaceManager]$NsManager,
    [string]$TargetAssetsDir,
    [string]$BaseFileName,
    [string]$NotebookName,
    [string]$SectionName,
    [string]$TargetFolder,
    [switch]$CreateCanvas,
    [switch]$SkipInkImages
  )

  $cleanImgBase = ($BaseFileName -replace '\s+', '_')
  $title = $PageDoc.DocumentElement.name
  $lines = [System.Collections.Generic.List[string]]::new()
  $lines.Add("# $title")
  $lines.Add("")

  $canvasNodes = [System.Collections.Generic.List[PSCustomObject]]::new()
  $canvasEdges = [System.Collections.Generic.List[PSCustomObject]]::new()

  # Alle visuellen Inhaltselemente sammeln
  $childElements = @()
  foreach ($child in $PageDoc.DocumentElement.ChildNodes) {
    if ($child.LocalName -in @("Outline", "Image", "InsertedFile", "InkDrawing")) {
      $pos = $child.SelectSingleNode("one:Position", $NsManager)
      $size = $child.SelectSingleNode("one:Size", $NsManager)
      $x = if ($pos -and $pos.x) { [double]$pos.x } else { 0.0 }
      $y = if ($pos -and $pos.y) { [double]$pos.y } else { 0.0 }
      $w = if ($size -and $size.width) { [double]$size.width } else { 400.0 }
      $h = if ($size -and $size.height) { [double]$size.height } else { 250.0 }
      $childElements += [PSCustomObject]@{
        Node = $child
        Tag = $child.LocalName
        x = $x
        y = $y
        w = $w
        h = $h
      }
    }
  }

  $sortedElements = $childElements | Sort-Object y, x

  # Performante Gruppierung von InkDrawings via Bounding-Box / Distanz-Cluster:
  # Zeichnungen werden nur dann gebündelt, wenn kein Image/Outline dazwischen liegt.
  $groupedElements = [System.Collections.Generic.List[PSCustomObject]]::new()
  $currInkGroup = $null

  foreach ($elemItem in $sortedElements) {
    if ($elemItem.Tag -eq "InkDrawing") {
      if ($null -eq $currInkGroup) {
        $currInkGroup = [PSCustomObject]@{
          Tag = "InkGroup"
          Items = [System.Collections.Generic.List[PSCustomObject]]::new()
          minX = $elemItem.x
          minY = $elemItem.y
          maxX = $elemItem.x + $elemItem.w
          maxY = $elemItem.y + $elemItem.h
        }
        $currInkGroup.Items.Add($elemItem)
      } else {
        # Check if stroke is spatially close to current group (threshold: 100pt vertical / 250pt horizontal)
        $closeY = ($elemItem.y - $currInkGroup.maxY) -le 100.0
        $closeX = ($elemItem.x - $currInkGroup.maxX) -le 250.0 -and ($currInkGroup.minX - ($elemItem.x + $elemItem.w)) -le 250.0

        if ($closeY -and $closeX) {
          $currInkGroup.Items.Add($elemItem)
          $currInkGroup.minX = [Math]::Min($currInkGroup.minX, $elemItem.x)
          $currInkGroup.minY = [Math]::Min($currInkGroup.minY, $elemItem.y)
          $currInkGroup.maxX = [Math]::Max($currInkGroup.maxX, $elemItem.x + $elemItem.w)
          $currInkGroup.maxY = [Math]::Max($currInkGroup.maxY, $elemItem.y + $elemItem.h)
        } else {
          $groupedElements.Add($currInkGroup)
          $currInkGroup = [PSCustomObject]@{
            Tag = "InkGroup"
            Items = [System.Collections.Generic.List[PSCustomObject]]::new()
            minX = $elemItem.x
            minY = $elemItem.y
            maxX = $elemItem.x + $elemItem.w
            maxY = $elemItem.y + $elemItem.h
          }
          $currInkGroup.Items.Add($elemItem)
        }
      }
    } else {
      # Once an intervening element (Image, Outline, etc.) arrives, finalize the current ink group.
      # This strictly prevents ink strokes from merging across images or text blocks.
      if ($null -ne $currInkGroup) {
        $groupedElements.Add($currInkGroup)
        $currInkGroup = $null
      }
      $groupedElements.Add($elemItem)
    }
  }
  if ($null -ne $currInkGroup) {
    $groupedElements.Add($currInkGroup)
    $currInkGroup = $null
  }

  $imgIndex = 0
  $hasSpatial = ($childElements.Count -gt 1)

  $canvasFileName = "$BaseFileName.canvas"
  $canvasFilePath = if ($TargetFolder) { Join-Path $TargetFolder $canvasFileName } else { "" }

  $lines.Add("> [!info] Imported OneNote Content")
  $metaLine = "> Notebook: $NotebookName | Section: $SectionName"
  if ($CreateCanvas -and $hasSpatial) {
    $metaLine += " | Canvas View: [[$canvasFileName]]"
  }
  $lines.Add($metaLine)
  $lines.Add("")

  foreach ($item in $groupedElements) {
    $tag = $item.Tag

    if ($tag -eq "Image") {
      $elem = $item.Node
      $imgIndex++
      $dataNode = $elem.SelectSingleNode("one:Data", $NsManager)
      if ($dataNode -and -not [string]::IsNullOrWhiteSpace($dataNode.InnerText)) {
        $ext = if ($elem.format) { $elem.format.ToLower() } else { "png" }
        if ($ext -eq "emf") { $ext = "png" }
        $imgFileName = "${cleanImgBase}_image_$($imgIndex.ToString('0000')).$ext"
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
          $lines.Add("")

          $vaultAssetRel = "OneNote/Assets/$imgFileName"
          $canvasNodes.Add([ordered]@{
            id = (New-HexId)
            type = "file"
            file = $vaultAssetRel
            x = [int][Math]::Round($item.x)
            y = [int][Math]::Round($item.y)
            width = [int][Math]::Max(100, [Math]::Round($item.w))
            height = [int][Math]::Max(100, [Math]::Round($item.h))
          })
        } catch {}
      }
    } elseif ($tag -eq "InkGroup") {
      if (-not $SkipInkImages) {
        $inkBlobs = [System.Collections.Generic.List[byte[]]]::new()
        foreach ($subItem in $item.Items) {
          Add-InkBlobsFrom -Scope $subItem.Node -NsManager $NsManager -Target $inkBlobs
        }

        if ($inkBlobs.Count -gt 0) {
          $pngBytes = New-InkStrokesPng -InkBlobs $inkBlobs
          if ($pngBytes -and $pngBytes.Length -gt 0) {
            $imgIndex++
            $imgFileName = "${cleanImgBase}_drawing_$($imgIndex.ToString('0000')).png"
            $imgPath = Join-Path $TargetAssetsDir $imgFileName
            try {
              [System.IO.File]::WriteAllBytes($imgPath, $pngBytes)

              $relPath = if ($TargetFolder -and (Test-Path -LiteralPath $TargetFolder)) {
                $fromUri = [System.Uri]((Resolve-Path -LiteralPath $TargetFolder).Path + "\")
                $toUri = [System.Uri]$imgPath
                [System.Uri]::UnescapeDataString($fromUri.MakeRelativeUri($toUri).ToString())
              } else {
                "../../Assets/$imgFileName"
              }
              $lines.Add("![drawing]($relPath)")
              $lines.Add("")

              $vaultAssetRel = "OneNote/Assets/$imgFileName"
              $groupW = [Math]::Max(50.0, ($item.maxX - $item.minX))
              $groupH = [Math]::Max(50.0, ($item.maxY - $item.minY))
              $canvasNodes.Add([ordered]@{
                id = (New-HexId)
                type = "file"
                file = $vaultAssetRel
                x = [int][Math]::Round($item.minX)
                y = [int][Math]::Round($item.minY)
                width = [int][Math]::Max(100, [Math]::Round($groupW))
                height = [int][Math]::Max(100, [Math]::Round($groupH))
              })
            } catch {}
          }
        }
      }
    } elseif ($tag -eq "Outline") {
      $elem = $item.Node
      $outlineLines = [System.Collections.Generic.List[string]]::new()
      $tables = $elem.SelectNodes(".//one:Table", $NsManager)
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
            $outlineLines.Add("| " + ($cells -join " | ") + " |")
            if ($isFirst) {
              $seps = foreach ($c in $cells) { "---" }
              $outlineLines.Add("| " + ($seps -join " | ") + " |")
              $isFirst = $false
            }
          }
          $outlineLines.Add("")
        }
      }

      foreach ($oe in $elem.SelectNodes(".//one:OE", $NsManager)) {
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
              $outlineLines.Add("$indent$prefix$text")
            } else {
              $outlineLines.Add("$indent$text")
            }
          }
        }

        $inkText = Get-InkRecognizedText -Scope $oe -NsManager $NsManager
        if ($inkText) {
          if ($prefix) {
            $outlineLines.Add("$indent$prefix$inkText")
          } else {
            $outlineLines.Add("$indent$inkText")
          }
        }

        $imgNode = $oe.SelectSingleNode("one:Image", $NsManager)
        if ($imgNode) {
          $imgIndex++
          $dataNode = $imgNode.SelectSingleNode("one:Data", $NsManager)
          if ($dataNode -and -not [string]::IsNullOrWhiteSpace($dataNode.InnerText)) {
            $ext = if ($imgNode.format) { $imgNode.format.ToLower() } else { "png" }
            if ($ext -eq "emf") { $ext = "png" }
            $imgFileName = "${cleanImgBase}_image_$($imgIndex.ToString('0000')).$ext"
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
              $outlineLines.Add("![image]($relPath)")

              $vaultAssetRel = "OneNote/Assets/$imgFileName"
              $canvasNodes.Add([ordered]@{
                id = (New-HexId)
                type = "file"
                file = $vaultAssetRel
                x = [int][Math]::Round($item.x)
                y = [int][Math]::Round($item.y)
                width = [int][Math]::Max(100, [Math]::Round($item.w))
                height = [int][Math]::Max(100, [Math]::Round($item.h))
              })
            } catch {}
          }
        }
      }

      $outlineText = ($outlineLines -join "`r`n").Trim()
      if ($outlineText) {
        $lines.Add($outlineText)
        $lines.Add("")

        $textWithoutImages = ($outlineText -replace '!\[.*?\]\(.*?\)', '').Trim()
        if ($textWithoutImages) {
          $canvasNodes.Add([ordered]@{
            id = (New-HexId)
            type = "text"
            x = [int][Math]::Round($item.x)
            y = [int][Math]::Round($item.y)
            width = [int][Math]::Max(150, [Math]::Round($item.w))
            height = [int][Math]::Max(50, [Math]::Round($item.h))
            text = $outlineText
          })
        }
      }
    }
  }

  if ($CreateCanvas -and $hasSpatial -and $canvasFilePath -and $canvasNodes.Count -gt 1) {
    $canvasObj = [ordered]@{
      nodes = $canvasNodes
      edges = $canvasEdges
    }
    $canvasJson = $canvasObj | ConvertTo-Json -Depth 5
    [System.IO.File]::WriteAllText($canvasFilePath, $canvasJson, [System.Text.UTF8Encoding]::new($false))
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
    $safe = Get-SafeFileName $page.name
    $pathInfo = Get-PageRelativePath -pageNode $page -safeTitle $safe
    $fullPath = Join-Path $output $pathInfo.RelFilePath
    $status = if (-not (Test-Path -LiteralPath $fullPath)) { "New" } else { "Existing" }
    [PSCustomObject]@{
      Title = $page.name
      Section = $pathInfo.RelFolder
      Modified = $page.lastModifiedTime
      Status = $status
    }
  }
  $listItems | Format-Table -AutoSize
  exit 0
}

if ($CleanDuplicates) {
  Write-Host "Checking for legacy exporter duplicates (* (1).md, etc.)..."
  $duplicateFiles = Get-ChildItem -LiteralPath $output -Recurse -Filter '* (*).md' -File
  $removedDups = 0
  foreach ($f in $duplicateFiles) {
    $baseName = $f.Name -replace '\s*\(\d+\)\.md$', '.md'
    $basePath = Join-Path $f.DirectoryName $baseName
    if (Test-Path -LiteralPath $basePath) {
      Remove-Item -LiteralPath $f.FullName -Force
      $removedDups++
    }
  }
  Write-Host "Cleanup completed: $removedDups duplicate file(s) removed."
}

# Manifest auf existierende OneNote-Seiten bereinigen
$allPageIdSet = @{}
foreach ($p in $allPageNodes) {
  $allPageIdSet[$p.ID] = $true
}
$toPrune = @()
foreach ($k in $manifest.Keys) {
  if (-not $allPageIdSet.ContainsKey($k)) {
    $toPrune += $k
  }
}
foreach ($k in $toPrune) {
  $manifest.Remove($k)
}

$script:sectionCursors = @{}
$toSync = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($page in $filteredPages) {
  $safeTitle = Get-SafeFileName $page.name
  $pathInfo = Get-PageRelativePath -pageNode $page -safeTitle $safeTitle
  $relFolder = $pathInfo.RelFolder
  $relFilePath = $pathInfo.RelFilePath
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
  Write-Host "OneNote is up to date. No changes found ($($filteredPages.Count) pages checked in $($sec)s)."
  exit 0
}

Write-Host "OneNote: $($toSync.Count) of $($filteredPages.Count) page(s) have updates. Synchronizing..."

$successCount = 0
foreach ($item in $toSync) {
  $page = $item.Page
  Write-Host "  -> Importing: $($page.name)"
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

    $enableCanvas = (-not $SkipCanvas)
    $mdContent = Convert-PageXmlToMarkdown -PageDoc $pdoc -NsManager $pns `
      -TargetAssetsDir $assetsDir -BaseFileName (Get-SafeFileName $page.name) `
      -NotebookName $nbName -SectionName $secName -TargetFolder $targetFolder `
      -CreateCanvas:$enableCanvas -SkipInkImages:$SkipInkImages

    [System.IO.File]::WriteAllText($item.AbsolutePath, $mdContent, [System.Text.UTF8Encoding]::new($false))

    # If the page was previously at a different path (renamed/moved), remove old file
    if ($manifest.Contains($page.ID)) {
      $oldRel = $manifest[$page.ID].path
      if ($oldRel -and $oldRel -ne $item.RelPath) {
        $oldFull = Join-Path $output $oldRel
        if (Test-Path -LiteralPath $oldFull) {
          Remove-Item -LiteralPath $oldFull -Force
        }
        $oldCanvas = [System.IO.Path]::ChangeExtension($oldFull, ".canvas")
        if (Test-Path -LiteralPath $oldCanvas) {
          Remove-Item -LiteralPath $oldCanvas -Force
        }
      }
    }

    $manifest[$page.ID] = [ordered]@{
      lastModifiedTime = $page.lastModifiedTime
      path = $item.RelPath
      title = $page.name
    }
    $successCount++
  } catch {
    Write-Warning "Error importing $($page.name): $($_.Exception.Message)"
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
Write-Host "Done: $successCount page(s) successfully synchronized (Duration: $($sec)s)."
exit 0
