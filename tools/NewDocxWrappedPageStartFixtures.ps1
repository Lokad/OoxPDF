# Public wrapped-note numeric before/after spacing controls; payloads reproduce exactly.
param([string]$OutputDirectory)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$cases=Join-Path $repo tests/Lokad.OoxPdf.Tests/Cases
if(!$OutputDirectory){$OutputDirectory=$cases}
New-Item -ItemType Directory -Force -Path $OutputDirectory|Out-Null
. (Join-Path $PSScriptRoot ZipPackage.ps1)
$word='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
foreach($spec in @(
 @{Seed='note-wrapped-calibri-mixed-run-early';Name='note-wrapped-calibri-before-four-early';Before=80;After=0},
 @{Seed='note-wrapped-calibri-mixed-met-later';Name='note-wrapped-calibri-both-eight-later';Before=160;After=160}
)){
 $entries=@{};$archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $cases ($spec.Seed+'.docx')))
 try {
  foreach($entry in $archive.Entries){
   $reader=[IO.StreamReader]::new($entry.Open())
   try{$entries[$entry.FullName]=$reader.ReadToEnd()}finally{$reader.Dispose()}
  }
 }finally{$archive.Dispose()}
 [xml]$body=$entries['word/document.xml']
 $ns=[Xml.XmlNamespaceManager]::new($body.NameTable);$ns.AddNamespace('w',$word)
 foreach($p in $body.SelectNodes('/w:document/w:body/w:p[w:r/w:footnoteReference]',$ns)){
  [void]$p.SelectSingleNode('w:pPr/w:spacing',$ns).SetAttribute('before',$word,[string]$spec.Before)
  [void]$p.SelectSingleNode('w:pPr/w:spacing',$ns).SetAttribute('after',$word,[string]$spec.After)
 }
 $entries['word/document.xml']=$body.OuterXml
 New-ZipPackage -Path (Join-Path $OutputDirectory ($spec.Name+'.docx')) -Entries $entries
}
