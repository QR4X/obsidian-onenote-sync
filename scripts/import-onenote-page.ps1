param(
  [Parameter(Mandatory)]
  [string]$PageTitle,
  [string]$Notebook,
  [string]$Section,
  [Parameter(Mandatory)]
  [string]$OutputDirectory
)

$ErrorActionPreference = "Stop"

function Convert-OneNoteTextToMarkdown {
  param([string]$Text)

  $markdown = [System.Net.WebUtility]::HtmlDecode($Text)
  $markdown = $markdown -replace '(?i)<br\s*/?>', "`n"
  $markdown = $markdown -replace "(?i)<span\s+style=['""][^'""]*font-weight\s*:\s*bold[^'""]*['""]>(.*?)</span>", '**$1**'
  $markdown = $markdown -replace '(?s)<[^>]+>', ''
  return $markdown.Trim()
}

$oneNote = New-Object -ComObject OneNote.Application
$hierarchyXml = ""
$oneNote.GetHierarchy("", 4, [ref]$hierarchyXml)

if ([string]::IsNullOrWhiteSpace($hierarchyXml)) {
  throw "OneNote returned no page hierarchy. Ensure the desktop OneNote application is running."
}

[xml]$hierarchy = $hierarchyXml
$namespace = New-Object System.Xml.XmlNamespaceManager($hierarchy.NameTable)
$namespace.AddNamespace("one", $hierarchy.DocumentElement.NamespaceURI)

$pages = $hierarchy.SelectNodes("//one:Page", $namespace) | Where-Object {
  $_.name -eq $PageTitle
}

if ($Notebook) {
  $pages = $pages | Where-Object {
    $_.SelectSingleNode("ancestor::one:Notebook", $namespace).name -eq $Notebook
  }
}

if ($Section) {
  $pages = $pages | Where-Object {
    $_.SelectSingleNode("ancestor::one:Section", $namespace).name -eq $Section
  }
}

$page = $pages | Sort-Object {
  if ($_.lastModifiedTime) {
    [datetime]$_.lastModifiedTime
  } else {
    [datetime]::MinValue
  }
} -Descending | Select-Object -First 1

if (-not $page) {
  throw "No matching OneNote page was found."
}

$pageXml = ""
$oneNote.GetPageContent($page.ID, [ref]$pageXml, 7)
[xml]$pageDocument = $pageXml
$pageNamespace = New-Object System.Xml.XmlNamespaceManager($pageDocument.NameTable)
$pageNamespace.AddNamespace("one", $pageDocument.DocumentElement.NamespaceURI)

$parts = @(
  foreach ($node in $pageDocument.SelectNodes("//one:T", $pageNamespace)) {
    $text = Convert-OneNoteTextToMarkdown $node.InnerText
    if ($text) {
      $text
    }
  }
)

if ($parts.Count -gt 0 -and $parts[0] -eq $page.name) {
  $parts = @($parts | Select-Object -Skip 1)
}

if ($parts.Count -eq 0) {
  throw "The OneNote page contains no text that can be imported."
}

New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$safeName = ($page.name -replace '[<>:"/\\|?*]', "_").Trim()
$destination = Join-Path $OutputDirectory "$safeName.md"
$content = "# $($page.name)`r`n`r`n> [!info] Untrusted imported content`r`n> Source: OneNote page imported via the local desktop application.`r`n`r`n" + ($parts -join "`r`n`r`n") + "`r`n"
[System.IO.File]::WriteAllText($destination, $content, [System.Text.UTF8Encoding]::new($false))

Write-Host "Imported OneNote page: $destination"
